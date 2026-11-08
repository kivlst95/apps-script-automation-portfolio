function deleteInvoices() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetById(0);
  const existingInvoices = sheet.getRange("A2:G").getValues().filter(f => f[0] != "" && String(f[0]).includes("CC-") && f[6] == "DELETE")

  const ids = existingInvoices.map(m => m[1])
  Log.info(ids)
  bulkVoidInvoices(ids)


}


function bulkVoidInvoices(xeroInvoiceIds) {
  if (!xeroInvoiceIds || xeroInvoiceIds.length === 0) return;

  for (let i = 0; i < xeroInvoiceIds.length; i += 100) {
    const chunk = xeroInvoiceIds.slice(i, i + 100);

    const payload = {
      Invoices: chunk.map(id => ({
        InvoiceID: id,
        Status: "AUTHORISED"
      // Status: "VOIDED"
    }))
  };

  try {
    const response = xeroRequest("Invoices", "POST", payload);
    Log.info(`Successfully voided ${chunk.length} invoices.`);
  } catch (e) {
    Log.error(`Error voiding batch: ${e.message}`);
  }

  // Respect Xero's rate limits
  Utilities.sleep(1000);
}
}


/**
 * CONFIGURATION - Fill these in
 */


/**
 * RUN THIS: Syncs both for a specific date range
 */
function syncClinicDay() {
  const dateToSync = "2026-01-01"; // You can automate this to "yesterday"
  const to = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  Log.info("Starting Invoice Sync...");
  syncInvoices(dateToSync, to);

  Log.info("Starting Payment Sync...");
  syncPayments(dateToSync);
}

/**
 * 1. SYNC INVOICES
 */
function syncInvoices(date, to) {
  const ccInvoices = fetchFromClinicCards(`invoices?from=${date}&to=${to}`);

  ccInvoices.forEach(ccInv => {
    // Check if invoice already exists to avoid duplicates
    const existing = findXeroInvoiceByReference("CC-" + ccInv.id);
    if (existing) {
      Log.warn(`Invoice CC-${ccInv.id} already exists in Xero. Skipping.`);
      return;
    }

    const contactId = getOrCreateXeroContact(ccInv.patient_name, ccInv.patient_email);
    const payload = {
      Type: "ACCREC",
      Contact: { ContactID: contactId },
      Date: ccInv.date,
      DueDate: ccInv.due_date || ccInv.date,
      Reference: "CC-" + ccInv.id, // THE BRIDGE
      LineItems: [{
        Description: "Service: " + ccInv.description,
        Quantity: 1,
        UnitAmount: ccInv.total,
        AccountCode: "200"
      }],
      Status: "AUTHORISED"
    };

    xeroRequest("https://api.xero.com/api.xro/2.0/Invoices", "POST", payload);
  });
}

/**
 * 2. SYNC PAYMENTS
 */
function syncPayments(date) {
  const ccPayments = fetchFromClinicCards(`payments?date=${date}`);

  ccPayments.forEach(ccPay => {
    // Find the Xero Invoice this payment belongs to using the Reference
    const xeroInvoice = findXeroInvoiceByReference("CC-" + ccPay.invoice_id);

    if (xeroInvoice) {
      const payload = {
        Invoice: { InvoiceID: xeroInvoice.InvoiceID },
        Account: { Code: getRequiredProperty_('XERO_BANK_ACCOUNT_CODE') },
        Date: ccPay.payment_date,
        Amount: ccPay.amount,
        Status: "AUTHORISED"
      };

      xeroRequest("https://api.xero.com/api.xro/2.0/Payments", "POST", payload);
      Log.info(`Applied ${ccPay.amount} to Xero Invoice ${xeroInvoice.InvoiceNumber}`);
    } else {
      Log.error(`Could not find Xero invoice for CC ID: ${ccPay.invoice_id}`);
    }
  });
}

/**
 * SEARCH HELPER: Finds Xero Invoice by ClinicCards ID
 */
function findXeroInvoiceByReference(ref) {
  const url = `https://api.xero.com/api.xro/2.0/Invoices?where=Reference=="${ref}"`;
  const response = xeroRequest(url, "GET");
  return (response.Invoices && response.Invoices.length > 0) ? response.Invoices[0] : null;
}

/**
 * CORE REQUEST HANDLERS
 */
function xeroRequest(url, method, payload = null) {
  const service = getXeroService_();
  const options = {
    method: method,
    headers: {
      "Authorization": "Bearer " + service.getAccessToken(),
      "Xero-tenant-id": getXeroTenantId_(),
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);

  const res = UrlFetchApp.fetch(url, options);
  return JSON.parse(res.getContentText());
}

function fetchFromClinicCards(endpoint) {
  const url = `https://cliniccards.com/api/${endpoint}`;
  const res = UrlFetchApp.fetch(url, { headers: { "Token": CC_API_KEY } });
  return JSON.parse(res.getContentText());
}

function getXeroInvoiceMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetById(0);

  const date = new Date();
  date.setDate(date.getDate() - 60);
  const dateStr = "2024-10-01";

  // const map = {};
  let page = 1;
  let keepFetching = true;
  let logData = []; // Buffer to write to sheet in one go (faster)

  Log.info("Starting pagination for Xero Invoices...");

  while (keepFetching) {
    const endpoint = `Invoices?Statuses=AUTHORISED,PAID&page=${page}`;
    // const endpoint = `Invoices?page=${page}`;

    const res = xeroRequest(endpoint, "GET");

    if (res.Invoices && res.Invoices.length > 0) {
      res.Invoices.forEach(inv => {
        if (inv.Reference) {
          logData.push([inv.Reference, inv.InvoiceID, inv.Contact.Name, inv.Total, inv.Status, inv.DateString]);
        }
      });

      Log.info(`Page ${page}: Found ${res.Invoices.length} invoices.`);

      if (res.Invoices.length < 100) {
        keepFetching = false;
      } else {
        page++;
      }
    } else {
      keepFetching = false;
    }
  }

  // Write all logs to the spreadsheet at once
  if (logData.length > 0) {
    sheet.getRange("A2:F").clearContent()
    sheet.getRange(sheet.getLastRow() + 1, 1, logData.length, logData[0].length).setValues(logData);
  }

  // return map;
}

function syncClinicCardsToXero(fromDate = new Date(2024, 9, 1), toDate = new Date(2024, 9, 10)) {


  // A. Fetch ClinicCards Invoices (Simulated Endpoint)
  // Replace with your actual DNS-fixed URL

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetById(0);
  const existingInvoices = sheet.getRange("A2:E").getValues().filter(f => f[0] != "")
    ?.reduce((accum, [refer, id, cus, amount, status]) => {


      if (!accum.byContact[cus]) accum.byContact[cus] = [];

      accum.byContact[cus].push(

        {
          id: id,
          amount: amount,
          status: status
        })


      accum.byReference[refer] = id


      return accum

    }, { byReference: {}, byContact: {} })


  const allCCInvoices = fetchCCInvoicesPaginated("2024-10-01", "2024-10-31").filter(f => f.amount);
  const allCCPayments = fetchCCPaymentsPaginated("2024-10-01", "2024-10-31").filter(f => f.amount);
  const paymentsWithRef = allCCPayments.filter(f => f?.invoice_id)


  const invoicesToCreate = allCCInvoices.filter(inv => !existingInvoices.byReference["CC-" + inv.id]);


  if (invoicesToCreate.length > 0) {
    for (let i = 0; i < invoicesToCreate.length; i += 100) {
      const chunk = invoicesToCreate.slice(i, i + 100);
      const created = bulkCreateXeroInvoices(chunk);
      // Update map so payments can find these new invoices
      created.Invoices.forEach(inv => {
        existingInvoices.byReference[inv.Reference] = inv.InvoiceID

        if (!existingInvoices.byContact[inv.Contact.Name]) existingInvoices.byContact[inv.Contact.Name] = [];

        existingInvoices.byContact[inv.Contact.Name].push(

          {
            id: inv.InvoiceID,
            amount: inv.Total,
            status: inv.Status
          })


      });

      Utilities.sleep(1000); // Rate limit protection
    }
  }


  processPayments(allCCPayments, existingInvoices);

}


function processPayments(allCCPayments, existingInvoices) {
  const paymentBatch = [];
  const prepaymentBatch = []

  allCCPayments.forEach(pay => {
    // Search by Contact Name and Amount (since no invoice_id is present)
    const match = findInvoiceByPatientAndAmount(pay.patient_id, pay.amount, existingInvoices.byContact);

    if (match) {
      paymentBatch.push({
        Invoice: { InvoiceID: match },
        Account: { AccountID: BANK_ACCOUNT_ID },
        Date: pay.date.split(' ')[0], // Fix: Strip time
        Amount: Number(pay.amount),   // Fix: Ensure Number
        Reference: "CC-PAY-" + (pay.payment_id || pay.id),
        Status: "AUTHORISED"
      });
    } else {
      prepaymentBatch.push({
        "Type": "RECEIVE-PREPAYMENT",
        Contact: { Name: pay.patient_id || "Unknown Patient" },
        BankAccount: { AccountID: BANK_ACCOUNT_ID },
        Date: pay.date.split(' ')[0], // Fix: Strip time
        // Amount: Number(pay.amount),   // Fix: Ensure Number
        LineItems: [{
          Description: `Prepayment - Ref: ${pay.payment_id || pay.id}`,
          Quantity: 1,
          UnitAmount: Number(pay.amount),
          AccountCode: "200",
          // Prepayments often require a TaxType (e.g., "OUTPUT" or "NONE")
          TaxType: "NONE"
        }],
        Reference: "CC-PAY-" + (pay.payment_id || pay.id),
      });

      // createXeroOverpayment(pay);
      // Utilities.sleep(1000)
    }
  });

  // Execute matched payments in bulk
  if (paymentBatch.length > 0) {
    for (let i = 0; i < paymentBatch.length; i += 100) {
      xeroRequest("Payments", "POST", { Payments: paymentBatch.slice(i, i + 100) });
      Utilities.sleep(1000);
    }
  }

  // if (prepaymentBatch.length > 0) {
  //   for (let i = 0; i < prepaymentBatch.length; i += 100) {
  //     xeroRequest("BankTransactions", "POST", { BankTransactions: prepaymentBatch.slice(i, i + 100) });
  //     Utilities.sleep(1000);
  //   }
  // }
}


function createXeroOverpayment(pay) {
  // If your 'pay' object contains a breakdown of the deposit (uncommon but possible), 
  // you could map them here. Otherwise, we use a structured single line item.

  const payload = {
    // BankTransactions: [{
    Type: "RECEIVE-PREPAYMENT",
    Contact: {
      Name: pay.patient_id || "Unknown Patient"
    },
    Date: pay.date.split(' ')[0],
    // Reference is vital for matching back to ClinicCards later
    Reference: "PREPAY-" + (pay.payment_id || "000"),

    LineItems: [{
      Description: pay.description || `Prepayment for services - Ref: ${pay.payment_id}`,
      Quantity: 1,
      UnitAmount: Number(pay.amount),
      AccountCode: "200", // Your Sales or Unearned Revenue code
      // You can also add Tracking Categories here if your clinic uses them
      // Tracking: [{ Name: "Region", Option: "North" }]
    }],

    BankAccount: {
      AccountID: BANK_ACCOUNT_ID
    }
    // }]
  };

  try {
    const response = xeroRequest("BankTransactions", "POST", payload);
    Log.info(`Overpayment created for ${pay.patient_id}: ${pay.amount}`);
    return response;
  } catch (e) {
    Log.error(`Failed to create overpayment for ${pay.patient_id}: ${e.message}`);
    return null;
  }
}

// B. Get existing Xero data to avoid per-item API calls
// const existingInvoices = getXeroInvoiceMap();

// const toCreate = [];
// const paymentsToProcess = [];

// invoices.forEach(ccInv => {
//   const ref = "CC-" + ccInv.id;

//   // Check if invoice needs creating
//   if (existingInvoices?.get(ref)) {
//     toCreate.push(ccInv);
//   }

//   // Check if there's a payment to apply
//   if (ccInv.payment_received > 0) {
//     paymentsToProcess.push(ccInv);
//   }
// });

// // C. Bulk Create Invoices (Up to 100)
// if (toCreate.length > 0) {
//   const bulkPayload = {
//     Invoices: toCreate.map(inv => ({
//       Type: "ACCREC",
//       Contact: { Name: inv.patient_name },
//       Date: inv.date,
//       DueDate: inv.due_date || inv.date,
//       Reference: "CC-" + inv.id,
//       Status: "AUTHORISED",
//       LineItems: [{
//         Description: "Clinic Services",
//         Quantity: 1,
//         UnitAmount: inv.total_amount,
//         AccountCode: "200"
//       }]
//     }))
//   };
//   const created = xeroRequest("Invoices", "POST", bulkPayload);
//   // Add newly created IDs to our map for the payment step
//   created.Invoices.forEach(inv => existingInvoices[inv.Reference] = inv.InvoiceID);
// }

// // D. Process Payments (Normal or Overpayment)
// paymentsToProcess.forEach(pay => {
//   const ref = "CC-" + pay.id;
//   const xeroId = existingInvoices[ref];

//   if (xeroId) {
//     // Regular Payment
//     xeroRequest("Payments", "POST", {
//       Invoice: { InvoiceID: xeroId },
//       Account: { AccountID: BANK_ACCOUNT_ID },
//       Date: pay.payment_date,
//       Amount: pay.payment_received,
//       Status: "AUTHORISED"
//     });
//   } else {
//     // PREPAYMENT LOGIC: Create Overpayment
//     xeroRequest("BankTransactions", "POST", {
//       BankTransactions: [{
//         Type: "RECEIVE-OVERPAYMENT",
//         Contact: { Name: pay.patient_name },
//         Date: pay.payment_date,
//         Reference: "PREPAY-" + pay.id,
//         BankAccount: { AccountID: BANK_ACCOUNT_ID },
//         LineItems: [{
//           Description: "Prepayment for services",
//           Quantity: 1,
//           UnitAmount: pay.payment_received,
//           AccountCode: "200"
//         }]
//       }]
//     });
//   }
// });

// console.log("Sync Complete.");
// }


/**
 * Searches the local cache of Xero invoices for a match.
 * @param {string} patientName - Name from ClinicCards payment
 * @param {number} amount - Payment amount
 * @param {Object} fullXeroMap - The enhanced map containing invoice details
 * @return {string|null} - The Xero InvoiceID if found, else null.
 */
function findInvoiceByPatientAndAmount(patientName, amount, fullXeroMap) {
  // 1. Normalize name for better matching (lowercase, trim spaces)
  const searchName = patientName;

  // 2. Look into our 'byContact' index
  const patientInvoices = fullXeroMap[searchName];

  if (!patientInvoices || patientInvoices.length === 0) {
    return null; // No invoices found at all for this patient
  }

  // 3. Look for an exact amount match among that patient's invoices
  // We prefer 'AUTHORISED' (unpaid) invoices
  const match = patientInvoices.find(inv =>
    Math.abs(inv.amount - amount) < 0.01 && inv.status === 'AUTHORISED'
  );

  if (match) {
    Log.info(`Match found! Linked ${patientName} payment to Invoice ${match.id}`);
    return match.id;
  }

  return null;
}


/**
 * Applies a payment to a specific invoice in Xero.
 * @param {string} xeroInvoiceId - The UUID of the invoice.
 * @param {number} amount - The amount to pay.
 * @param {string} date - The date of the payment (YYYY-MM-DD).
 */
function applyXeroPayment(xeroInvoiceId, amount, date, id) {
  const payload = {
    Payments: [{
      Invoice: {
        InvoiceID: xeroInvoiceId
      },
      Account: {
        AccountID: BANK_ACCOUNT_ID // Your Xero Bank Account UUID
      },
      Date: date.split(' ')[0],
      Amount: Number(amount),
      Reference: "CC-PAY-" + id,
    }]
  };

  try {
    const response = xeroRequest("Payments", "PUT", payload);
    Log.info(`Payment of ${amount} applied to Xero Invoice ${xeroInvoiceId}`);
    return response;
  } catch (e) {
    // If the invoice is already paid or locked, Xero will throw an error
    Log.error(`Failed to apply payment to Invoice ${xeroInvoiceId}: ${e.message}`);
    return null;
  }
}

function fetchCCInvoicesPaginated(startDate, endDate) {
  const CC_API_KEY = PropertiesService.getScriptProperties().getProperty("CLINICCARDS_API_KEY");
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, "yyyy-MM-dd");

  const final = new Date(endDate);
  let allInvoices = [];
  let current = new Date(startDate);

  while (current <= final) {
    // the window is always calculated from current, as a separate object
    let to = new Date(current);
    to.setDate(to.getDate() + 7);
    if (to > final) to = new Date(final);

    const dateStr = fmt(current);
    const dateStrTo = fmt(to);
    Log.info(`Fetching ClinicCards Invoices: ${dateStr} — ${dateStrTo}`);

    try {
      const url = `https://cliniccards.com/api/invoices?from=${dateStr}&to=${dateStrTo}`;
      const response = UrlFetchApp.fetch(url, {
        headers: { "Token": CC_API_KEY },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        if (Array.isArray(data.data)) {
          allInvoices = allInvoices.concat(data.data);
        }
      } else {
        Log.warn(`Could not fetch ${dateStr}—${dateStrTo}: ${response.getContentText()}`);
      }
    } catch (e) {
      Log.error(`Error on ${dateStr}: ${e.message}`);
    }

    // the next window starts one day after the end of the previous one
    current = new Date(to);
    current.setDate(current.getDate() + 1);
  }

  const targetMonth = Utilities.formatDate(new Date(endDate), tz, "MMM-yy");
  Log.info(`Total Invoices fetched: ${allInvoices.length}`);
  return allInvoices.filter(f =>
    f.purpose !== "PREINVOICE"
  );
}


function fetchCCPaymentsPaginated(startDate, endDate) {
  const CC_API_KEY = PropertiesService.getScriptProperties().getProperty("CLINICCARDS_API_KEY")
  let allInvoices = [];
  let current = new Date(startDate);
  let toIterartion = new Date(startDate);
  toIterartion.setDate(toIterartion.getDate() + 7);

  const final = new Date(endDate);

  while (current <= final) {
    // Format date as YYYY-MM-DD
    const dateStr = current.toISOString().split('T')[0];
    const dateSteTo = toIterartion.toISOString().split('T')[0];


    Log.info(`Fetching ClinicCards Payments for: ${dateStr}`);

    try {
      const url = `https://cliniccards.com/api/payments?from=${dateStr}&to=${dateSteTo}`;
      const response = UrlFetchApp.fetch(url, {
        headers: { "Token": CC_API_KEY },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        if (Array.isArray(data.data)) {
          allInvoices = allInvoices.concat(data.data);
        }
      } else {
        Log.warn(`Could not fetch data for ${dateStr}: ${response.getContentText()}`);
      }
    } catch (e) {
      Log.error(`Error on ${dateStr}: ${e.message}`);
    }

    // Move to the next day
    current.setDate(current.getDate() + 8);
    toIterartion.setDate(current.getDate() + 7);

  }

  Log.info(`Total Payments fetched from ClinicCards: ${allInvoices.length}`);
  return allInvoices;
}

function bulkCreateXeroInvoices(chunk) {
  const payload = {
    Invoices: chunk.map(inv => {

      // Map Line Items: If CC provides an array of services, use it. 
      // Otherwise, create a single line item from the total.
      const lineItems = inv.invoice_items.map(item => ({
        Description: item?.plan_item_name || "Clinic Service",
        Quantity: item?.quantity || 1,
        UnitAmount: item?.price || 0,
        AccountCode: "200", // Sales/Revenue
        TaxType: "NONE",
        // DiscountRate: item.discountrate
      }))

      return {
        Type: "ACCREC",
        Contact: { Name: inv.patient_id || "Unknown Patient" },
        Date: inv.date_created,
        DueDate: inv?.due_date || inv.date_created,
        Reference: "CC-" + inv.id,
        Status: "AUTHORISED",
        CurrencyCode: "UAH",
        LineItems: lineItems // Injected line items array here
      };
    })
  };

  try {
    const response = xeroRequest("Invoices", "POST", payload);
    Log.info(`Successfully created ${response.Invoices.length} invoices.`);
    return response;
  } catch (e) {
    Log.error("Bulk Creation Error: " + e.message);
    throw e;
  }
}

/******************************************************************
 * INVOICE UPLOAD: one Xero invoice per patient + day,
 * each clinic service record as a separate line. Replaces the whole content of invoice_upload.gs.
 *
 * Invoice reference:  CC-2026-10-01-1234567   (date + patient)
 * Line description:   CC-12345678 | service name
 *
 * Sheet CC_Invoices_Log: which service records are already uploaded (duplicate protection).
 * If a new service record appears in the clinic system for a day already uploaded,
 * it goes as a separate invoice with suffix -2, -3 and ends up in Issues.
 ******************************************************************/

const INV_LOG_SHEET = 'CC_Invoices_Log'

function uploadInvoicesByDate() {

  const ss = SpreadsheetApp.getActive()
  const sheet = ss.getSheetById(RECON.SETTINGS_SHEET_ID)
  const tz = ss.getSpreadsheetTimeZone()

  const datesRange = sheet.getRange("B2:B3")
  const [fromDate, toDate] = datesRange.getValues().flat()

  if (!fromDate || !toDate || toDate < fromDate) {
    SpreadsheetApp.getUi().alert("Missing Dates", "Please enter from and to in correct way.", SpreadsheetApp.getUi().ButtonSet.OK);
    return
  }

  const fromS = Utilities.formatDate(fromDate, tz, "yyyy-MM-dd")
  const toS = Utilities.formatDate(toDate, tz, "yyyy-MM-dd")

  // What already exists in Xero (updates the invoces_list sheet)
  getXeroInvoiceMap()
  Utilities.sleep(1000)
  const existingRefs = new Set(
    ss.getSheetById(0).getRange("A2:A").getValues().flat().filter(String).map(String)
  )

  // Which service records are already uploaded with the new scheme
  const log = invLogSheet_()
  const loggedActs = new Set(log.getRange("A2:A").getValues().flat().filter(String).map(String))

  const allCCInvoices = (fetchCCInvoicesPaginated(fromS, toS) || [])
    .filter(f => f.amount && f.date_created)
    .filter(f => { const d = String(f.date_created).slice(0, 10); return d >= fromS && d <= toS })

  // Record already exists: either in the log (new scheme) or as an old invoice CC-<id>
  const actsToCreate = allCCInvoices.filter(inv =>
    !loggedActs.has(String(inv.id)) && !existingRefs.has(`CC-${inv.id}`))

  if (!actsToCreate.length) {
    datesRange.clearContent()
    SpreadsheetApp.getUi().alert("No new service records for the period.")
    return
  }

  // Group by: day + patient
  const groups = {}
  actsToCreate.forEach(inv => {
    const d = String(inv.date_created).slice(0, 10)
    const patient = String(inv.patient_id || "Unknown Patient")
    const key = d + "|" + patient
    if (!groups[key]) groups[key] = { d, patient, acts: [] }
    groups[key].acts.push(inv)
  })

  const issues = []
  const toCreate = Object.values(groups).map(g => {
    let ref = `CC-${g.d}-${g.patient}`
    if (existingRefs.has(ref)) {
      let n = 2
      while (existingRefs.has(`${ref}-${n}`)) n++
      issues.push([g.d, g.patient, '', '', g.acts.map(a => a.id).join(', '),
        `Invoice ${ref} for this day already existed; uploaded as ${ref}-${n}`])
      ref = `${ref}-${n}`
    }
    existingRefs.add(ref)
    return { ...g, ref }
  })

  const logRows = []
  for (let i = 0; i < toCreate.length; i += 50) {
    const chunk = toCreate.slice(i, i + 50)
    const res = bulkCreateXeroDayInvoices(chunk)
    const created = (res && res.Invoices) || []

    chunk.forEach((g, idx) => {
      const x = created[idx]
      const errs = x && (x.ValidationErrors || []).map(e => e.Message).join(' | ')
      if (!x || errs || x.StatusAttributeString === 'ERROR') {
        issues.push([g.d, g.patient, '', '', g.acts.map(a => a.id).join(', '), 'Invoice not created: ' + (errs || 'no response')])
        return
      }
      g.acts.forEach(a => logRows.push([String(a.id), g.ref, g.d, g.patient, Number(a.amount), x.InvoiceID, new Date()]))
    })
    Utilities.sleep(2000)
  }

  if (logRows.length) log.getRange(log.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows)
  if (issues.length) writeInvIssues_(issues)

  datesRange.clearContent()
  SpreadsheetApp.getUi().alert(`Service records: ${actsToCreate.length}\nInvoices (patient/day): ${toCreate.length}\nIssues: ${issues.length}`)
}


function bulkCreateXeroDayInvoices(groups) {
  const payload = {
    Invoices: groups.map(g => ({
      Type: "ACCREC",
      Contact: { Name: g.patient },
      Date: g.d,
      DueDate: g.d,
      Reference: g.ref,
      Status: "AUTHORISED",
      CurrencyCode: "UAH",
      LineItems: g.acts.flatMap(act =>
        (act.invoice_items && act.invoice_items.length ? act.invoice_items : [{ plan_item_name: "Clinic Service", quantity: 1, price: act.amount }])
          .map(item => ({
            Description: `CC-${act.id} | ${item?.plan_item_name || "Clinic Service"}`,
            Quantity: item?.quantity || 1,
            UnitAmount: item?.price || 0,
            AccountCode: "200",
            TaxType: "NONE",
          }))
      ),
    }))
  }

  const response = xeroRequest("Invoices?summarizeErrors=false", "POST", payload)
  Log.info(`Created/processed ${(response.Invoices || []).length} day-invoices.`)
  return response
}


function invLogSheet_() {
  const ss = SpreadsheetApp.getActive()
  let sh = ss.getSheetByName(INV_LOG_SHEET)
  if (!sh) {
    sh = ss.insertSheet(INV_LOG_SHEET)
    sh.getRange(1, 1, 1, 7).setValues([['act_id', 'xero_reference', 'date', 'patient', 'amount', 'xero_invoice_id', 'posted_at']]).setFontWeight('bold')
    sh.setFrozenRows(1)
  }
  return sh
}

function writeInvIssues_(rows) {
  const ss = SpreadsheetApp.getActive()
  let sh = ss.getSheetByName('Issues')
  if (!sh) {
    sh = ss.insertSheet('Issues')
    sh.getRange(1, 1, 1, 7).setValues([['date', 'patient', 'method', 'amount', 'ids', 'problem', 'logged_at']]).setFontWeight('bold')
    sh.setFrozenRows(1)
  }
  const data = rows.map(r => [...r, new Date()])
  sh.getRange(sh.getLastRow() + 1, 1, data.length, data[0].length).setValues(data)
}

/******************************************************************
 * PAYMENT UPLOAD: groups the statement by patient + day, separately per sheet
 * (Cash / Credit_card). Replaces the whole content of payment_upload.gs.
 ******************************************************************/
function uploadPaymentsByDate() {

  const ss = SpreadsheetApp.getActive()
  const sheet = ss.getSheetById(RECON.SETTINGS_SHEET_ID)
  const tz = ss.getSpreadsheetTimeZone()

  const datesRange = sheet.getRange("E2:E3")
  const [fromDate, toDate] = datesRange.getValues().flat()

  if (!fromDate || !toDate || toDate < fromDate) {
    SpreadsheetApp.getUi().alert("Missing Dates", "Please enter from and to in correct way.", SpreadsheetApp.getUi().ButtonSet.OK);
    return
  }

  const fromS = Utilities.formatDate(fromDate, tz, "yyyy-MM-dd")
  const toS = Utilities.formatDate(toDate, tz, "yyyy-MM-dd")

  // fetch may include days after To, so they are cut off
  const allCCPayments = (fetchCCPaymentsPaginated(fromS, toS) || [])
    .filter(p => Number(p.amount) && p.date)
    .filter(p => { const d = String(p.date).slice(0, 10); return d >= fromS && d <= toS })

  if (!allCCPayments.length) return

  const tabSettings = sheet.getRange("G2:H").getValues().filter(f => f[0] != "")

  tabSettings.forEach(([method, sheetId]) => {

    const targetSheet = ss.getSheetById(sheetId)

    // group by: day + patient
    const groups = {}
    allCCPayments
      .filter(p => p.type == method)
      .forEach(p => {
        const d = String(p.date).slice(0, 10)
        const patient = String(p.patient_id || "Unknown Patient")
        const key = d + "|" + patient
        if (!groups[key]) groups[key] = { d, patient, sum: 0, ids: [] }
        groups[key].sum += Number(p.amount)
        groups[key].ids.push(p.payment_id)
      })

    const targetData = Object.values(groups)
      .map(g => ({ ...g, sum: Math.round(g.sum * 100) / 100 }))
      .filter(g => Math.abs(g.sum) >= 0.005)           // payment and change netted to 0, no row
      .sort((a, b) => a.d.localeCompare(b.d) || a.patient.localeCompare(b.patient))
      .map(g => [g.d, g.sum, g.patient, `Payment id: ${g.ids.join(", ")}`, ""])

    targetSheet.getRange("A2:E").clearContent()
    if (targetData.length) {
      targetSheet.getRange(2, 1, targetData.length, targetData[0].length).setValues(targetData)
    }
  })

  datesRange.clearContent()
}

/******************************************************************
 * CLINIC SYSTEM -> XERO: PAYMENTS WITHOUT MANUAL RECONCILIATION
 *
 * What it does:
 * 1. Pulls payments from the clinic system for the period (Settings E2:E3).
 * 2. Groups them: patient + day + method = one net amount
 *    (payment 39,600 and change -100 -> 39,500).
 * 3. Allocates the net amount to the patient unpaid invoices:
 *    first those referenced by the payment itself (invoice_id),
 *    then in order from the oldest (only invoices dated on or before the payment).
 * 4. The remainder -> Overpayment (patient deposit, kept as a credit on the contact).
 * 5. Allocates open Overpayments to new patient invoices (allocate).
 * 6. Everything that failed -> Issues sheet. Processed payment_id -> sheet
 *    CC_Payments_Log, so a rerun for the same period does not duplicate.
 *
 * Uses: xeroRequest(), fetchCCPaymentsPaginated().
 ******************************************************************/

const RECON = {
  // Xero bank account AccountIDs for each clinic payment method.
  // Take them from logBankAccounts() -> the ID: line.
  ACCOUNTS: {
    CASH: getProperty_('XERO_CASH_ACCOUNT_ID', ''),
    CREDIT_CARD: getProperty_('XERO_CARD_ACCOUNT_ID', ''),
  },
  // true  -> payments are created already reconciled, the Cash/Card statement is NOT imported into Xero.
  // false -> payments wait for a statement line (then import the grouped statement).
  MARK_RECONCILED: true,
  SETTINGS_SHEET_ID: Number(getProperty_('SETTINGS_SHEET_GID', '0')),
  LOG_SHEET: 'CC_Payments_Log',
  ISSUES_SHEET: 'Issues',
};

/* ===================== MAIN FUNCTION ===================== */

function postPaymentsToXero() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const tz = ss.getSpreadsheetTimeZone();
  const settings = ss.getSheetById(RECON.SETTINGS_SHEET_ID);
  const range = settings.getRange('E2:E3');
  const [from, to] = range.getValues().flat();

  if (!from || !to || to < from) {
    ui.alert('Missing Dates', 'Set From / To in Payment settings.', ui.ButtonSet.OK);
    return;
  }
  const fromS = Utilities.formatDate(from, tz, 'yyyy-MM-dd');
  const toS = Utilities.formatDate(to, tz, 'yyyy-MM-dd');

  // 1. Clinic payments: strictly cut by dates (fetch may include extra days)
  const raw = (fetchCCPaymentsPaginated(fromS, toS) || [])
    .filter(p => Number(p.amount) && p.date)
    .filter(p => { const d = p.date.slice(0, 10); return d >= fromS && d <= toS; });

  // 2. Skip already processed payment_id
  const log = getSheet_(RECON.LOG_SHEET,
    ['payment_id', 'date', 'patient', 'method', 'amount', 'group_net', 'result', 'details', 'posted_at']);
  const done = new Set(log.getRange('A2:A').getValues().flat().filter(String).map(String));
  const fresh = raw.filter(p => !done.has(String(p.payment_id)));
  // dedup within one upload
  const seen = new Set();
  const payments = fresh.filter(p => { const k = String(p.payment_id); if (seen.has(k)) return false; seen.add(k); return true; });

  if (!payments.length) {
    ui.alert('No new payments for the period.');
    return;
  }

  // 3. Group by: day + patient + method
  const groups = {};
  payments.forEach(p => {
    const date = p.date.slice(0, 10);
    const patient = String(p.patient_id || 'Unknown Patient');
    const key = [date, patient, p.type].join('|');
    if (!groups[key]) groups[key] = { key, date, patient, type: p.type, net: 0, items: [], ccRefs: [] };
    const g = groups[key];
    g.net = r2_(g.net + Number(p.amount));
    g.items.push(p);
    if (p.invoice_id) g.ccRefs.push('CC-' + p.invoice_id);
  });

  // 4. Open Xero invoices per patient
  const openByContact = getOpenInvoicesByContact_();
  const arCode = getArAccountCode_();

  const payBatch = [], payMeta = [];
  const opBatch = [], opMeta = [];
  const issues = [];
  const logRows = [];
  const now = new Date();

  Object.values(groups)
    .sort((a, b) => a.date.localeCompare(b.date))
    .forEach(g => {
      const ids = g.items.map(i => i.payment_id).join(', ');
      const ref = 'CC-PAY-' + g.items[0].payment_id;
      const account = RECON.ACCOUNTS[g.type];

      const logGroup = (result, details) => g.items.forEach(i =>
        logRows.push([String(i.payment_id), g.date, g.patient, g.type, Number(i.amount), g.net, result, details, now]));

      if (!account || account.startsWith('PASTE_')) {
        issues.push([g.date, g.patient, g.type, g.net, ids, 'No AccountID for payment method ' + g.type + ' in RECON.ACCOUNTS']);
        return; // not logged, so it is picked up after the fix
      }
      if (Math.abs(g.net) < 0.005) {
        logGroup('NETTED_TO_ZERO', 'Payment and refund on the same day netted to 0');
        return;
      }
      if (g.net < 0) {
        issues.push([g.date, g.patient, g.type, g.net, ids, 'Refund without a payment on the same day: post manually (refund overpayment / spend money)']);
        return;
      }

      // Allocate to invoices
      let left = g.net;
      const applied = [];
      const candidates = (openByContact[g.patient] || [])
        .filter(inv => inv.due > 0.005 && inv.date <= g.date)
        .sort((a, b) =>
          (g.ccRefs.includes(b.ref) - g.ccRefs.includes(a.ref)) || a.date.localeCompare(b.date));

      for (const inv of candidates) {
        if (left < 0.005) break;
        const amt = r2_(Math.min(left, inv.due));
        payBatch.push({
          Invoice: { InvoiceID: inv.id },
          Account: { AccountID: account },
          Date: g.date,
          Amount: amt,
          Reference: ref,
          IsReconciled: RECON.MARK_RECONCILED,
        });
        payMeta.push({ g, inv, amt });
        inv.due = r2_(inv.due - amt);
        left = r2_(left - amt);
        applied.push(`${inv.ref}: ${amt}`);
      }

      // Remainder -> patient deposit
      if (left > 0.005) {
        opBatch.push({
          Type: 'RECEIVE-OVERPAYMENT',
          Contact: { Name: g.patient },
          Date: g.date,
          Reference: ref,
          BankAccount: { AccountID: account },
          IsReconciled: RECON.MARK_RECONCILED,
          LineAmountTypes: 'NoTax',
          LineItems: [{
            Description: `Patient deposit, CC payment id: ${ids}`,
            Quantity: 1,
            UnitAmount: left,
            AccountCode: arCode,
          }],
        });
        opMeta.push({ g, amt: left });
        applied.push(`DEPOSIT: ${left}`);
      }

      logGroup('POSTED', applied.join('; '));
    });

  // 5. Send to Xero (with per-element error checks)
  const failedKeys = new Set();
  sendInChunks_('Payments?summarizeErrors=false', 'Payments', payBatch, payMeta, issues, failedKeys);
  sendInChunks_('BankTransactions?summarizeErrors=false', 'BankTransactions', opBatch, opMeta, issues, failedKeys);

  // Only groups without errors are written to the log
  const okRows = logRows.filter(r => {
    const key = [r[1], r[2], r[3]].join('|');
    return !failedKeys.has(key);
  });
  if (okRows.length) log.getRange(log.getLastRow() + 1, 1, okRows.length, okRows[0].length).setValues(okRows);

  // 6. Allocate deposits to new invoices
  const allocated = allocateOpenOverpayments(true);

  writeIssues_(issues);
  range.clearContent();

  ui.alert('Done',
    `Clinic payments: ${payments.length}\n` +
    `Groups (patient/day/method): ${Object.keys(groups).length}\n` +
    `Payments in Xero: ${payBatch.length}\n` +
    `Deposits (overpayments): ${opBatch.length}\n` +
    `Deposits allocated to invoices: ${allocated}\n` +
    `Issues: ${issues.length}`, ui.ButtonSet.OK);
}

/* ===================== DEPOSIT ALLOCATION ===================== */

/**
 * Takes all open Overpayments and uses them to close unpaid invoices of the same
 * patient, starting from the oldest. Can be run separately from the menu.
 */
function allocateOpenOverpayments(silent) {
  const openByContact = getOpenInvoicesByContact_();
  const overpayments = getOpenOverpayments_();
  let count = 0;
  const issues = [];

  overpayments.forEach(op => {
    let credit = op.remaining;
    const invs = (openByContact[op.contact] || [])
      .filter(i => i.due > 0.005)
      .sort((a, b) => a.date.localeCompare(b.date));

    const allocations = [];
    for (const inv of invs) {
      if (credit < 0.005) break;
      const amt = r2_(Math.min(credit, inv.due));
      allocations.push({
        Invoice: { InvoiceID: inv.id },
        Amount: amt,
        Date: inv.date > op.date ? inv.date : op.date,
      });
      inv.due = r2_(inv.due - amt);
      credit = r2_(credit - amt);
    }
    if (!allocations.length) return;

    const res = xeroRequest(`Overpayments/${op.id}/Allocations`, 'PUT', { Allocations: allocations });
    const errs = collectErrors_(res);
    if (errs) {
      issues.push([op.date, op.contact, 'ALLOCATION', op.remaining, op.id, errs]);
    } else {
      count += allocations.length;
    }
    Utilities.sleep(1100); // Xero limit 60 requests/min
  });

  writeIssues_(issues);
  if (!silent) SpreadsheetApp.getUi().alert(`Deposits allocated to invoices: ${count}. Errors: ${issues.length}`);
  return count;
}

/* ===================== XERO HELPERS ===================== */

function getOpenInvoicesByContact_() {
  const map = {};
  let page = 1;
  while (true) {
    const res = xeroRequest(`Invoices?Statuses=AUTHORISED&page=${page}`, 'GET');
    const list = (res && res.Invoices) || [];
    list.filter(i => i.Type === 'ACCREC').forEach(i => {
      const c = String(i.Contact.Name);
      (map[c] = map[c] || []).push({
        id: i.InvoiceID,
        ref: i.Reference || '',
        date: (i.DateString || '').slice(0, 10),
        due: Number(i.AmountDue) || 0,
      });
    });
    if (list.length < 100) break;
    page++;
    Utilities.sleep(1100);
  }
  return map;
}

function getOpenOverpayments_() {
  const out = [];
  let page = 1;
  while (true) {
    const res = xeroRequest(`Overpayments?page=${page}`, 'GET');
    const list = (res && res.Overpayments) || [];
    list
      .filter(o => o.Type === 'RECEIVE-OVERPAYMENT' && o.Status === 'AUTHORISED' && Number(o.RemainingCredit) > 0.005)
      .forEach(o => out.push({
        id: o.OverpaymentID,
        contact: String(o.Contact.Name),
        date: (o.DateString || '').slice(0, 10),
        remaining: Number(o.RemainingCredit),
      }));
    if (list.length < 100) break;
    page++;
    Utilities.sleep(1100);
  }
  return out.sort((a, b) => a.date.localeCompare(b.date));
}

/** Accounts Receivable account code, taken from Xero automatically. */
function getArAccountCode_() {
  const res = xeroRequest('Accounts', 'GET');
  const ar = (res.Accounts || []).find(a => a.SystemAccount === 'DEBTORS');
  if (!ar) throw new Error('Accounts Receivable account not found in Xero');
  return ar.Code;
}

function sendInChunks_(endpoint, key, batch, meta, issues, failedKeys) {
  for (let i = 0; i < batch.length; i += 50) {
    const chunk = batch.slice(i, i + 50);
    const chunkMeta = meta.slice(i, i + 50);
    const res = xeroRequest(endpoint, 'PUT', { [key]: chunk });
    const items = (res && res[key]) || [];

    if (!items.length) {
      // the whole request failed
      const msg = collectErrors_(res) || JSON.stringify(res).slice(0, 300);
      chunkMeta.forEach(m => { failedKeys.add(m.g.key); issues.push(issueRow_(m, msg)); });
    } else {
      items.forEach((it, idx) => {
        const errs = collectErrors_(it);
        if (errs || it.StatusAttributeString === 'ERROR') {
          const m = chunkMeta[idx];
          failedKeys.add(m.g.key);
          issues.push(issueRow_(m, errs || 'ERROR'));
        }
      });
    }
    Utilities.sleep(1100);
  }
}

function collectErrors_(obj) {
  if (!obj) return 'Empty response';
  const errs = [];
  (obj.ValidationErrors || []).forEach(e => errs.push(e.Message));
  (obj.Elements || []).forEach(el => (el.ValidationErrors || []).forEach(e => errs.push(e.Message)));
  if (obj.Message && obj.ErrorNumber) errs.push(obj.Message);
  return errs.length ? errs.join(' | ') : '';
}

function issueRow_(m, msg) {
  return [m.g.date, m.g.patient, m.g.type, m.amt,
    m.g.items.map(i => i.payment_id).join(', ') + (m.inv ? ` → ${m.inv.ref}` : ' → DEPOSIT'), msg];
}

/* ===================== SHEET HELPERS ===================== */

function getSheet_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function writeIssues_(rows) {
  if (!rows.length) return;
  const sh = getSheet_(RECON.ISSUES_SHEET, ['date', 'patient', 'method', 'amount', 'payment ids', 'problem', 'logged_at']);
  const now = new Date();
  const data = rows.map(r => [...r, now]);
  sh.getRange(sh.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}

function r2_(n) { return Math.round(Number(n) * 100) / 100; }

/* ===================== MENU (replaces onOpen) ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 XERO Import')
    .addItem('Add Invoices', 'uploadInvoicesByDate')
    .addSeparator()
    .addItem('Import Payments (to the Cash / CC sheets)', 'uploadPaymentsByDate')
    .addItem('Post Payments to Xero (auto-allocation)', 'postPaymentsToXero')
    .addItem('Allocate deposits to invoices', 'allocateOpenOverpayments')
    .addToUi();
}

/******************************************************************
 * RESET CC INVOICES: removes payments and voids old invoices
 * per service record (Reference = CC-<record id>) for the period from Settings B2:B3.
 *
 * Touches ONLY ACCREC invoices with a Reference like CC-12345678.
 * New patient+day invoices (CC-2026-09-01-...), Historical, Other revenue are NOT touched.
 *
 * The script can be rerun: each time it takes from Xero
 * what is not voided yet. With many invoices it stops after ~5 min
 * and asks to run again (Apps Script limit 6 min, Xero 60 requests/min).
 *
 * Dates B2:B3 are NOT cleared after the run, so Add Invoices can run right away.
 ******************************************************************/

function resetClinicInvoicesForPeriod() {
  const started = Date.now()
  const TIME_LIMIT = 5 * 60 * 1000

  const ss = SpreadsheetApp.getActive()
  const ui = SpreadsheetApp.getUi()
  const tz = ss.getSpreadsheetTimeZone()
  const [fromDate, toDate] = ss.getSheetById(RECON.SETTINGS_SHEET_ID).getRange("B2:B3").getValues().flat()

  if (!fromDate || !toDate || toDate < fromDate) {
    ui.alert("Missing Dates", "Set From / To in Invoices settings (B2:B3).", ui.ButtonSet.OK)
    return
  }

  const fromS = Utilities.formatDate(fromDate, tz, "yyyy-MM-dd")
  const toS = Utilities.formatDate(toDate, tz, "yyyy-MM-dd")

  const invoices = getOldCCInvoices_(fromS, toS)
  if (!invoices.length) {
    ui.alert(`No old CC-<id> invoices for ${fromS} - ${toS}.`)
    return
  }

  const paymentsCount = invoices.reduce((s, i) => s + (i.Payments || []).length, 0)
  const answer = ui.alert("Confirm",
    `Period: ${fromS} - ${toS}\n` +
    `Invoices per service record: ${invoices.length}\n` +
    `Payments on them that will be deleted: ${paymentsCount}\n\n` +
    `Invoices will be voided (VOIDED). Continue?`, ui.ButtonSet.YES_NO)
  if (answer !== ui.Button.YES) return

  const issues = []
  let deletedPayments = 0
  const readyToVoid = []
  let stoppedEarly = false

  // 1. Delete payments
  for (const inv of invoices) {
    if (Date.now() - started > TIME_LIMIT) { stoppedEarly = true; break }

    let ok = true
    for (const p of (inv.Payments || [])) {
      const res = xeroRequest(`Payments/${p.PaymentID}`, "POST", { Status: "DELETED" })
      const err = resetErr_(res)
      if (err) {
        ok = false
        issues.push([inv.DateString.slice(0, 10), inv.Contact.Name, '', p.Amount, inv.Reference, 'Payment not deleted: ' + err])
      } else {
        deletedPayments++
      }
      Utilities.sleep(1100)
    }

    // Allocated deposits / credit notes cannot be removed via the API, only manually
    const other = (inv.Overpayments || []).length + (inv.Prepayments || []).length + (inv.CreditNotes || []).length
    if (other) {
      ok = false
      issues.push([inv.DateString.slice(0, 10), inv.Contact.Name, '', '', inv.Reference,
        'The invoice has an overpayment / prepayment / credit note: remove the allocation manually in Xero'])
    }

    if (ok) readyToVoid.push(inv.InvoiceID)
  }

  // 2. Void in batches
  let voided = 0
  for (let i = 0; i < readyToVoid.length; i += 50) {
    const chunk = readyToVoid.slice(i, i + 50)
    const res = xeroRequest("Invoices?summarizeErrors=false", "POST",
      { Invoices: chunk.map(id => ({ InvoiceID: id, Status: "VOIDED" })) })
    ;(res.Invoices || []).forEach((x, idx) => {
      const err = resetErr_(x)
      if (err || x.StatusAttributeString === 'ERROR') {
        issues.push(['', '', '', '', chunk[idx], 'Not voided: ' + (err || 'ERROR')])
      } else {
        voided++
      }
    })
    Utilities.sleep(1100)
  }

  if (issues.length) writeInvIssues_(issues)   // from invoice_upload.gs

  // update the invoces_list sheet so Add Invoices does not treat them as existing
  getXeroInvoiceMap()

  ui.alert(stoppedEarly ? "Partially done: run again" : "Done",
    `Payments deleted: ${deletedPayments}\n` +
    `Invoices voided: ${voided}\n` +
    `Problems (Issues sheet): ${issues.length}\n` +
    (stoppedEarly ? `\nTime limit reached: run Reset again, it will continue.` :
      `\nNext: Add Invoices -> Import Payments -> import the clearing statement.`),
    ui.ButtonSet.OK)
}


/** ACCREC, AUTHORISED/PAID, date within the period, Reference = CC-<digits only> */
function getOldCCInvoices_(fromS, toS) {
  const [fy, fm, fd] = fromS.split('-').map(Number)
  const [ty, tm, td] = toS.split('-').map(Number)
  const where = `Type=="ACCREC" AND Date>=DateTime(${fy},${fm},${fd}) AND Date<=DateTime(${ty},${tm},${td})`

  const out = []
  let page = 1
  while (true) {
    const res = xeroRequest(`Invoices?where=${where}&Statuses=AUTHORISED,PAID&page=${page}`, "GET")
    const list = (res && res.Invoices) || []
    list.filter(i => /^CC-\d+$/.test(i.Reference || '')).forEach(i => out.push(i))
    if (list.length < 100) break
    page++
    Utilities.sleep(1100)
  }
  return out
}

function resetErr_(obj) {
  if (!obj) return 'empty response'
  const e = []
  ;(obj.ValidationErrors || []).forEach(v => e.push(v.Message))
  ;(obj.Elements || []).forEach(el => (el.ValidationErrors || []).forEach(v => e.push(v.Message)))
  if (obj.ErrorNumber && obj.Message) e.push(obj.Message)
  return e.join(' | ')
}
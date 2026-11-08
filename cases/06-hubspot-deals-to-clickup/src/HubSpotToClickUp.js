// function onOpen() {

//   regularHandbookUpdate()

// }


function regularHandbookUpdate() {

  let hubspotOptionsObject = []
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Hubspot enumeration");


  exportClickUpListCustomFieldsToSheet()
  hubspotOptionsObject = getHubspotOrganizationFields(hubspotOptionsObject)
  hubspotOptionsObject = getHubspotContactsFields(hubspotOptionsObject)
  hubspotOptionsObject = getHubspotDealsFields(hubspotOptionsObject)
  getClickUpUsers()

  sheet.getRange("A2:C").clearContent()
  sheet.getRange(2, 1, hubspotOptionsObject.length, hubspotOptionsObject[0].length).setValues(hubspotOptionsObject)

}


function exportClickUpListCustomFieldsToSheet() {
  const CLICKUP_TOKEN = PropertiesService.getScriptProperties()
    .getProperty('CLICKUP_API_TOKEN');
  const LIST_ID = getRequiredProperty_('CLICKUP_DEALS_LIST_ID');
  const SHEET_NAME = 'Handbook';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const url = `https://api.clickup.com/api/v2/list/${LIST_ID}/field`;

  const options = {
    method: 'get',
    headers: {
      Authorization: CLICKUP_TOKEN,
      'Content-Type': 'application/json'
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (!data.fields || !data.fields.length) {
    Log.info('No custom fields found');
    return;
  }

  const rows = data.fields.map(field => [
    field.id,
    field.name,
    field.type,
    field.required,
    JSON.stringify(field.type_config || {}),
    LIST_ID
  ]);

  sheet.getRange("A7:F").clearContent()
  sheet.getRange(7, 1, rows.length, rows[0].length).setValues(rows);

  // sheet.autoResizeColumns(1, 6);

  // Logger.log(`✅ Exported ${rows.length} custom fields`);
}


function getHubspotOrganizationFields(hubspotOptionsObject) {
  // Replace with your HubSpot user access token

  const url = '/crm/v3/properties/companies';
  const response = callHubSpot("GET", url, "")

  Log.info(response)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Handbook");

  const allFields = []

  response.results?.forEach(obj => {


    const label = obj.label
    const name = obj.name
    const type = obj.type


    allFields.push([label, name, type])

    if (type == "enumeration") {

      obj.options?.forEach(({ label, value }) => hubspotOptionsObject.push([name, value, label]))

    }

  })

  sheet.getRange("L2:N").clearContent()
  sheet.getRange(2, 12, allFields.length, allFields[0].length).setValues(allFields)

  return hubspotOptionsObject

}


function getHubspotContactsFields(hubspotOptionsObject) {
  // Replace with your HubSpot user access token

  const url = '/crm/v3/properties/contacts';
  const response = callHubSpot("GET", url, "")

  Log.info(response)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Handbook");

  const allFields = []

  response.results?.forEach(obj => {

    const label = obj.label
    const name = obj.name
    const type = obj.type

    allFields.push([label, name, type])

    if (type == "enumeration") {

      obj.options?.forEach(({ label, value }) => hubspotOptionsObject.push([name, value, label]))

    }


  })


  sheet.getRange("P2:R").clearContent()
  sheet.getRange(2, 16, allFields.length, allFields[0].length).setValues(allFields)

  return hubspotOptionsObject

}


function getHubspotDealsFields(hubspotOptionsObject) {
  // Replace with your HubSpot user access token

  const url = '/crm/v3/properties/deals';
  const response = callHubSpot("GET", url, "")

  Log.info(response)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Handbook");

  const allFields = []

  response.results?.forEach(obj => {

    const label = obj.label
    const name = obj.name
    const type = obj.type

    allFields.push([label, name, type])

    if (type == "enumeration") {

      obj.options?.forEach(({ label, value }) => hubspotOptionsObject.push([name, value, label]))

    }


  })


  sheet.getRange("T2:V").clearContent()
  sheet.getRange(2, 20, allFields.length, allFields[0].length).setValues(allFields)

  return hubspotOptionsObject

}


function getClickUpUsers() {
  const CLICKUP_TOKEN = PropertiesService.getScriptProperties()
    .getProperty('CLICKUP_API_TOKEN');
  const url = 'https://api.clickup.com/api/v2/team';


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Handbook");

  const allMembers = []

  const options = {
    'method': 'GET',
    'headers': {
      'Authorization': CLICKUP_TOKEN,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      Log.error("Error: " + data.err);
      return;
    }

    // Iterate through Workspaces (Teams)
    data.teams.forEach(team => {
      Log.info("--- Workspace: " + team.name + " (ID: " + team.id + ") ---");

      // Iterate through members in that Workspace
      team.members.forEach(member => {
        const user = member.user;
        allMembers.push([user.id, user.email, user.username])
      });
    });

    sheet.getRange("H2:J").clearContent()
    sheet.getRange(2, 8, allMembers.length, allMembers[0].length).setValues(allMembers)


  } catch (e) {
    Log.error("Execution failed: " + e.toString());
  }
}


/*
 * This sample demonstrates how to configure the library for the HubSpot API.
 * Instructions on how to generate OAuth credentuals is available here:
 * https://developers.hubspot.com/docs/api/oauth-quickstart-guide
 */

/** Sheet IDs (gid) inside the handbook spreadsheet. Script Property HANDBOOK_SHEET_GIDS (JSON). */
const HANDBOOK_SHEET_GIDS = getJsonProperty_('HANDBOOK_SHEET_GIDS', {});
/** ClickUp custom field IDs of the deals list. Script Property CLICKUP_FIELD_IDS (JSON). */
const CLICKUP_FIELD_IDS = getJsonProperty_('CLICKUP_FIELD_IDS', {});

var HUBSPOT_CLIENT_ID = PropertiesService.getScriptProperties().getProperty("HUBSPOT_CLIENT_ID");
var HUBSPOT_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty("HUBSPOT_CLIENT_SECRET");

/**
 * Authorizes and makes a request to the HubSpot API.
 */
function syncClosedWonDealsToClickUp() {
  var service = getService_();
  if (service.hasAccess()) {
    // Make a request to retrieve the list of CRM owners.
    var url = 'https://api.hubapi.com/crm/v3/owners/';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      },
    });
    var result = JSON.parse(response.getContentText());
    Log.info(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    var ui = SpreadsheetApp.getUi();
    ui.prompt(
      'Action Required: Connect to HubSpot',
      'Please copy the link below, paste it into your browser, and log in:\n\n' + authorizationUrl,
      ui.ButtonSet.OK
    );
  }
}

function callHubSpot(method, endpoint, payload) {
  var service = getService_();

  // 1. Authorization Wrap: Check if the user is connected
  if (!service.hasAccess()) {
    var authorizationUrl = service.getAuthorizationUrl();
    var ui = SpreadsheetApp.getUi();
    ui.prompt(
      'Action Required: Connect to HubSpot',
      'Please copy the link below, paste it into your browser, and log in:\n\n' + authorizationUrl,
      ui.ButtonSet.OK
    );
    return null; // Stop the function here until they authorize
  }

  // 2. Prepare the Request
  var url = 'https://api.hubapi.com' + endpoint;
  var options = {
    'method': method,
    'headers': {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };

  if (payload && method !== 'GET') {
    options.payload = JSON.stringify(payload);
  }

  // 3. Execute the Request
  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();
  var content = response.getContentText();

  if (statusCode >= 200 && statusCode < 300) {
    return JSON.parse(content);
  } else {
    Log.error('❌ HubSpot Error ' + statusCode + ': ' + content);
    return null;
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService_().reset();
}

/**
 * Configures the service.
 */
function getService_() {
  return OAuth2.createService('HubSpot')
    // Set the endpoint URLs.
    .setAuthorizationBaseUrl('https://app.hubspot.com/oauth/authorize')
    .setTokenUrl('https://api.hubapi.com/oauth/v1/token')

    // Set the client ID and secret.
    .setClientId(HUBSPOT_CLIENT_ID)
    .setClientSecret(HUBSPOT_CLIENT_SECRET)

    // Set the name of the callback function that should be invoked to
    // complete the OAuth flow.
    .setCallbackFunction('authCallback')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getScriptProperties())

    // Set the scopes to request from the user. The full list of scopes is
    // available here:
    // https://developers.hubspot.com/docs/api/working-with-oauth#scopes
    .setScope('automation crm.objects.companies.write crm.objects.companies.read crm.objects.contacts.write crm.objects.contacts.read crm.objects.deals.write crm.objects.deals.read crm.objects.users.write crm.objects.users.read crm.schemas.companies.write crm.schemas.companies.read crm.schemas.contacts.write crm.schemas.contacts.read crm.schemas.deals.write crm.schemas.deals.read forms');
};

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService_();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redirect URI to register.
 */
function logRedirectUri() {
  Log.info(OAuth2.getRedirectUri());
}

function toClickUpMs(val) {
  if (!val) return null;

  // Convert string to number if possible
  let n = isNaN(val) ? new Date(val).getTime() : parseFloat(val);

  // If it's 10 digits (seconds), scale to milliseconds
  return n < 3e10 ? n * 1000 : n;
}


function clickupValue(tasksID, type, value, fieldID) {

  const url = `https://api.clickup.com/api/v2/task/${tasksID}/field/${fieldID}`

  const payload = {
    ...(type != "date") && { "value": value },
    ...(type == "date") && {
      "value_options": {
        "time": false
      },
      "value": value
    },

  }


  clickUpRequest(url, "POST", payload)


}


function clickUpRequest(url, method, payload = "") {

  const CLICKUP_API_TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN")


  const options = {
    method: method,
    headers: {
      'Authorization': CLICKUP_API_TOKEN,
      'content-type': 'application/json'
    },
    ...(payload && { payload: JSON.stringify(payload) })
  }


  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  return data


}


function syncHubSpotToClickUp() {

  getClosedWonDeals()
  Utilities.sleep(1000)
  checkForms()

}

function checkForms() {

  const ss = SpreadsheetApp.openById(getRequiredProperty_('HANDBOOK_SPREADSHEET_ID'))
  const settings = ss.getSheetById(Number(HANDBOOK_SHEET_GIDS.settings))
  const handbook = ss.getSheetById("0")
  const enumeration = ss.getSheetById(Number(HANDBOOK_SHEET_GIDS.enumeration))
  const submissionsList = ss.getSheetById(Number(HANDBOOK_SHEET_GIDS.submissions))


  const appliedSubmissions = submissionsList.getRange("A2:A").getValues().filter(f => f[0] != "").flat()


  const enumerationHubspot = enumeration.getRange("A2:C").getValues().reduce((acc, [key, value, label]) => {
    // Initialize the array if the key doesn't exist yet
    if (!acc[key]) acc[key] = [];

    // Push the object into the correct group
    acc[key].push({ label: label, value: value });

    return acc;
  }, {});

  const formID = settings.getRange("D1").getValue()

  const now = new Date();
  const yesterday = new Date();
  yesterday.setDate(now.getDate() - 2); // Subtract 1 day

  const startDate = yesterday.getTime();
  const endDate = now.getTime();


  let allFilteredSubmissions = [];
  let afterOffset = null;
  let keepFetching = true;

  while (keepFetching) {
    // Construct URL - append the 'after' offset only if it exists from a previous page
    let url = `/form-integrations/v1/submissions/forms/${formID}?limit=50`;
    if (afterOffset) {
      url += `&after=${afterOffset}`;
    }


    const data = callHubSpot("GET", url, "")

    if (!data) {
      Log.error("API Error: " + response.getContentText());
      break;
    }

    const results = data.results || [];

    // Filter this batch
    const batchFiltered = results.filter(sub => sub.submittedAt >= startDate && !appliedSubmissions.includes(sub.conversionId));
    allFilteredSubmissions = allFilteredSubmissions.concat(batchFiltered);

    if (data.paging && data.paging.next && results.length > 0) {
      const lastItemTimestamp = results[results.length - 1].submittedAt;

      if (lastItemTimestamp >= startDate) {
        afterOffset = data.paging.next.after;
      } else {
        keepFetching = false; // The end of this batch is older than 24h
      }
    } else {
      keepFetching = false; // No more pages
    }
  }


  if (allFilteredSubmissions.length) {

    const formSubmition = new Map(settings.getRange("A3:E").getValues().filter(f => f[0])?.map(m => [m[0], [m[1], m[2], m[3], m[4]]]))

    const clickupUsers = new Map(handbook.getRange("H2:J").getValues().filter(f => f[0])?.map(m => [m[2], m[0]]))

    const keys = [...formSubmition.keys()]

    const contactsFields = keys.filter(f => f.includes("cont"))?.map(m => String(m).split("¬")[1])
    const companiesFields = keys.filter(f => f.includes("comp"))?.map(m => String(m).split("¬")[1])
    companiesFields.push("clickup_id")


    allFilteredSubmissions.forEach(submission => {

      const values = submission?.values


      const email = values?.find(f => f.name == "email")?.value
      const domain = values?.find(f => f.name == "domain")?.value

      const { company, contact } = getClickUpIdFromEmail(email, domain)

      const companyValues = company
        ? callHubSpot("GET", `/crm/v3/objects/companies/${company}?properties=${companiesFields.join(",")}`, "")
        : {}

      const contactValues = contact
        ? callHubSpot("GET", `/crm/v3/objects/contacts/${contact}?properties=${contactsFields.join(",")}`, "")
        : {}


      const associatedClickup = companyValues?.properties?.clickup_id


      if (associatedClickup) {


        keys.forEach(key => {
          const [pref, propertyName] = key.split("¬");

          let value = ""

          switch (pref) {

            case "cont":

              value = findEnum_(enumerationHubspot, propertyName, contactValues.properties?.[propertyName])

              break;


            case "comp":

              value = findEnum_(enumerationHubspot, propertyName, companyValues.properties?.[propertyName])

              break;


          }


          const findSetting = formSubmition?.get(key)


          if (findSetting) {

            const clickupField = findSetting[1]
            const clickupType = findSetting[2]
            const options = findSetting[3]
              ? JSON?.parse(findSetting[3])?.options
              : ""
            let valueToAdd = ""

            switch (clickupType) {

              case "users":

                const userID = clickupUsers?.get(value)

                if (userID) {

                  valueToAdd = {
                    "add": [
                      userID
                    ]
                  }

                }

                break;

              case "drop_down":

                const findDd = options?.find(f => f.name == value)

                if (findDd) {

                  valueToAdd = findDd.id

                }

                break;


              case "labels":

                const optionsLb = options?.find(f => f.label == value)
                if (optionsLb) {

                  valueToAdd = [value]

                }

                break;

              case "date":

                const insertDate = toClickUpMs(value)

                if (insertDate) {

                  valueToAdd = insertDate


                }


                break;


              default:
                valueToAdd = value

            }

            Log.info(valueToAdd)
            clickupValue(associatedClickup, clickupType, valueToAdd, clickupField)


          }


        })


      }


      submissionsList.appendRow([
        submission.conversionId,
        submission.submittedAt,
        submission?.values ? JSON.stringify(submission.values) : "{}"])


    })


  }


}


function getClickUpIdFromEmail(email, domain) {
  // 1. Search for Contact by Email


  if (email) {

    const contactSearchUrl = "/crm/v3/objects/contacts/search";
    const contactPayload = {
      "filterGroups": [{
        "filters": [{ "propertyName": "email", "operator": "EQ", "value": email }]
      }],
      "properties": ["associatedcompanyid"],
      "limit": 1
    };

    const contactRes = callHubSpot("POST", contactSearchUrl, contactPayload);
    let contact = contactRes?.results?.[0];

    if (contact) {
      return {
        company: contact.properties.associatedcompanyid || null,
        contact: contact.id
      };
    }
  }
  // 2. Fallback: Search for Company by Domain
  Log.warn("Contact not found by email. Searching by domain: " + domain);
  const companySearchUrl = "/crm/v3/objects/companies/search";
  const companyPayload = {
    "filterGroups": [{
      "filters": [{ "propertyName": "domain", "operator": "EQ", "value": domain }]
    }],
    "limit": 1
  };

  const companyRes = callHubSpot("POST", companySearchUrl, companyPayload);
  let company = companyRes?.results?.[0];

  if (company) {
    return {
      company: company.id,
      contact: null // We found the business, but we don't know which person
    };
  }

  Log.info("No Contact or Company found for: " + email + " / " + domain);
  return null;
}


function getContactEmailsForCompany(companyId) {
  // 1. Get the Company and its associated contact IDs
  const companyData = getCompanyAssociationById(companyId);

  // if (!companyData || companyData.contactIds.length === 0) {
  //   return [];
  // }

  // 2. Prepare the Batch Request payload
  const url = "/crm/v3/objects/contacts/batch/read";
  const payload = {
    inputs: companyData?.associations?.contacts?.reults?.map(id => ({ id: id })),
    properties: ["email", "firstname", "lastname"]
  };

  try {
    const response = callHubSpot("POST", url, payload);

    // 3. Extract emails from the results
    return response.results.map(contact => contact.properties.email);
  } catch (err) {
    Log.error(`Error batch fetching contacts: ${err}`);
    return [];
  }
}

function getCompanyAssociationById(companyId) {
  if (!companyId) return null;

  // Define which properties you want to retrieve
  const props = ["name", "domain", "clickup_id", "city"];
  const url = `/crm/v3/objects/companies/${companyId}?properties=${props.join(',')}&associations=contacts`;

  try {
    const response = callHubSpot("GET", url);

    // The CRM v3 API returns properties inside a 'properties' object
    return response || null;
  } catch (err) {
    Log.error(`Error fetching company ${companyId}: ${err}`);
    return null;
  }
}


function getCompanyById(companyId) {
  if (!companyId) return null;

  // Define which properties you want to retrieve
  const props = ["name", "domain", "clickup_id", "city"];
  const url = `/crm/v3/objects/companies/${companyId}?properties=${props.join(',')}&associations=contacts`;

  try {
    const response = callHubSpot("GET", url);

    // The CRM v3 API returns properties inside a 'properties' object
    return response?.properties || null;
  } catch (err) {
    Log.error(`Error fetching company ${companyId}: ${err}`);
    return null;
  }
}


function getClosedWonDeals() {
  const url = '/crm/v3/objects/deals/search';


  const ss = SpreadsheetApp.openById(getRequiredProperty_('HANDBOOK_SPREADSHEET_ID'))
  const settings = ss.getSheetById(Number(HANDBOOK_SHEET_GIDS.settings))
  const enumeration = ss.getSheetById(Number(HANDBOOK_SHEET_GIDS.enumeration))

  const handbook = ss.getSheetById("0")

  const formSubmition = new Map(settings.getRange("G3:K").getValues().filter(f => f[0])?.map(m => [m[0], [m[1], m[2], m[3], m[4]]]))

  const clickupUsers = new Map(handbook.getRange("H2:J").getValues().filter(f => f[0])?.map(m => [m[2], m[0]]))

  const enumerationHubspot = enumeration.getRange("A2:C").getValues().reduce((acc, [key, value, label]) => {
    // Initialize the array if the key doesn't exist yet
    if (!acc[key]) acc[key] = [];

    // Push the object into the correct group
    acc[key].push({ label: label, value: value });

    return acc;
  }, {});

  const keys = [...formSubmition.keys()]

  const dealsFields = keys.filter(f => f.includes("deal"))?.map(m => String(m).split("¬")[1])
  const contactsFields = keys.filter(f => f.includes("cont"))?.map(m => String(m).split("¬")[1])
  const companiesFields = keys.filter(f => f.includes("comp"))?.map(m => String(m).split("¬")[1])

  companiesFields.push("clickup_id")


  // Define your date range (Year, Month [0-11], Day)
  const now = new Date();
  const yesterday = new Date();
  yesterday.setDate(now.getDate() - 1); // Subtract 1 day

  const startDate = yesterday.getTime();
  const endDate = now.getTime();

  // const startDate = new Date(2025, 0, 01).getTime();   // Jan 1, 2024
  // const endDate = new Date(2026, 1, 31).getTime();  // Jan 31, 2024


  const payload = {
    filterGroups: [
      {
        filters: [
          {
            propertyName: 'dealstage',
            operator: 'EQ',
            value: 'closedwon'
          },
          {
            propertyName: 'closedate',
            operator: 'GTE',
            value: startDate.toString()
          },
          {
            propertyName: 'closedate',
            operator: 'LTE',
            value: endDate.toString()
          }
        ]
      }
    ],
    properties: dealsFields,
    associations: ["companies"],
    limit: 100
  };

  const response = callHubSpot("POST", url, payload);
  if (response) {
    const deals = response.results;

    // const companyIds = [...new Set(deals
    //   .map(d => d.associations?.companies?.results[0]?.id)
    //   .filter(id => id != null))];

    // const companyMap = {};
    // if (companyIds.length > 0) {
    //   const batchUrl = '/crm/v3/objects/companies/batch/read';
    //   const batchPayload = { inputs: companyIds.map(id => ({ id })), properties: ["name", "clickup_id"] };

    //   const batchRes = callHubSpot("POST", batchUrl, batchPayload);

    //   batchRes?.results.forEach(co => {
    //     companyMap[co.id] = co.properties.clickup_id;
    //   });
    // }


    deals?.forEach(deal => {

      const companyID = callHubSpot("GET", `/crm/v3/objects/deals/${deal.id}?associations=companies`, "")

      const associatedCompany = companyID?.associations?.companies?.results[0]?.id

      const contactsID = callHubSpot("GET", `/crm/v3/objects/deals/${deal.id}?associations=contacts`, "")


      const associatedContact = contactsID?.associations?.contacts?.results[0]?.id


      const companyValues = callHubSpot("GET", `/crm/v3/objects/companies/${associatedCompany}?properties=${companiesFields.join(",")}`, "")
      const associatedClickup = companyValues?.properties?.clickup_id

      if (associatedClickup) {

        return

      }


      const contactValues = contactsFields.length
        ? callHubSpot("GET", `/crm/v3/objects/contacts/${associatedContact}?properties=${contactsFields.join(",")}`, "")
        : ""


      // Create new Clickup object 

      const payload = {

        "status": "onboarding",
        "custom_fields": [

          {
            id: CLICKUP_FIELD_IDS.company,
            value: associatedCompany
          }


        ]

      }


      keys.forEach(key => {
        const [pref, propertyName] = key.split("¬");

        let value = ""

        switch (pref) {

          case "deal":

            value = findEnum_(enumerationHubspot, propertyName, deal.properties?.[propertyName])

            break;


          case "cont":

            value = findEnum_(enumerationHubspot, propertyName, contactValues.properties?.[propertyName])

            break;


          case "comp":

            value = findEnum_(enumerationHubspot, propertyName, companyValues.properties?.[propertyName])

            break;


        }


        const findSetting = formSubmition?.get(key)

        if (findSetting) {

          const clickupField = findSetting[1]
          const clickupType = findSetting[2]
          const options = findSetting[3]
            ? JSON?.parse(findSetting[3])?.options
            : ""

          if (["name", "description", "assignees", "due_date", "start_date"].includes(findSetting[0])) {

            switch (findSetting[0]) {

              case "name":

                payload.name = value

                break;


              case "description":

                payload.description = value

                break;


              case "assignees":

                const userID = clickupUsers?.get(value)

                if (userID) {

                  payload.assignees = [userID]

                }


                break;


              case "due_date":

                const insertDate = toClickUpMs(value)

                if (insertDate) {

                  payload.due_date = insertDate
                  payload.due_date_time = false


                }

                break;

              case "start_date":

                const insertDateStart = toClickUpMs(value)

                if (insertDateStart) {

                  payload.start_date = insertDateStart
                  payload.start_date_time = false


                }

                break;


            }


          }

          else {

            switch (clickupType) {

              case "users":

                const userID = clickupUsers?.get(value)

                if (userID) {

                  payload.custom_fields.push(

                    {
                      "value": {
                        "add": [
                          userID
                        ]
                      },
                      "id": clickupField
                    }

                  )

                }

                break;

              case "drop_down":

                const findDd = options?.find(f => f.name == value)

                if (findDd) {

                  payload.custom_fields.push(

                    {
                      "id": clickupField,
                      "value": findDd.id
                    }

                  )
                }

                break;


              case "labels":

                const optionsLb = options?.find(f => f.label == value)
                if (optionsLb) {

                  payload.custom_fields.push(

                    {
                      "value": [
                        optionsLb.id
                      ],
                      "id": clickupField
                    }

                  )
                }

                break;

              case "date":

                const insertDate = toClickUpMs(value)

                if (insertDate) {
                  payload.custom_fields.push(

                    {
                      "value_options": {
                        "time": false
                      },
                      "id": clickupField,
                      "value": insertDate
                    }

                  )

                }


                break;


              default:


                if (value) {

                  payload.custom_fields.push(

                    {
                      "id": clickupField,
                      "value": value
                    }

                  )

                }

            }


          }

        }


      })


      const newClickup = clickUpRequest("https://api.clickup.com/api/v2/list/${getRequiredProperty_('CLICKUP_DEALS_LIST_ID')}/task", "POST", payload)
      const newClickupID = newClickup?.id

      if (newClickupID) {

        const url = `/crm/v3/objects/companies/${associatedCompany}`

        callHubSpot("PATCH", url,
          {
            "properties": {
              "clickup_id": newClickupID,
            }
          }

        )

      }


    })


  }
}


const findEnum_ = (map, property, value) => {
  // 1. Get the array for that specific property (e.g., 'company_type')
  const enumList = map?.[property];

  // 2. If no list exists, return the value as-is
  if (!Array.isArray(enumList)) return value;

  // 3. Find the matching object. 
  // Using == instead of === allows it to match "1" to 1 (string vs number)
  const match = enumList.find(f => f.value == value || f.value == Boolean(value));

  // 4. Return the label if found, otherwise the original value
  return match ? match.label : value;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Update Script')
    .addItem('Update fields lists', 'regularHandbookUpdate')
    .addSeparator()
    .addItem('Initiate Form Submition', 'checkForms')
    .addSeparator()
    .addItem('Initiate Closed Won', 'getClosedWonDeals')


    .addToUi();
}
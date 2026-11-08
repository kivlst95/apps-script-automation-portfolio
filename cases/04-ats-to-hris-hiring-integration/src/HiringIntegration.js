/**
 * Organisation-specific IDs, kept out of the code (Script Properties, JSON).
 * See config.example.json for the expected shape.
 */
const MANATAL_STAGE_IDS = getJsonProperty_('MANATAL_STAGE_IDS', {});
const PEOPLEFORCE_IDS = getJsonProperty_('PEOPLEFORCE_IDS', { leavePolicies: [] });

/** Base URL of the PeopleForce web app, e.g. https://acme.peopleforce.io */
function getPeopleForceBaseUrl_() {
  return `https://${getRequiredProperty_('PEOPLEFORCE_SUBDOMAIN')}.peopleforce.io`;
}

function createEmployeeInPeopleForceOffer(matchID, manual = false) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN")
  const findCandidateURL = `https://api.manatal.com/open/v3/matches/${matchID}/`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountHistory = ss.getSheetById("0")


  const orhanizationMap = new Map(accountHistory.getRange("J2:K").getValues().filter(f => f[0] != ""))

  const manatalOptions = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const responseMatch = UrlFetchApp.fetch(findCandidateURL, manatalOptions);
  const statusCode = responseMatch.getResponseCode();
  const responseBody = responseMatch.getContentText();

  Log.info("Status: " + statusCode);
  Log.info("Response: " + responseBody);

  if (statusCode === 200) {
    const data = JSON.parse(responseBody);
    // 👇 You can now access data.candidate or similar
    Log.info("Match ID: " + data.id);

    const getCandidate = data.candidate
    const organization = data.organization || ""


    const organizationName = orhanizationMap.get(organization) || ""

    if (getCandidate && organizationName != "NDA") return createPFEmployeeFromManatalOffer_(getCandidate, manatalOptions, matchID, manual)


  }
  else if (statusCode === 422) {

    Log.error("❌ Error fetching match: " + responseBody);
    throw Error(responseBody)


  }

  else {
    Log.error("❌ Error fetching match: " + responseBody);
  }
}

function messagePeopleForceHired(matchID) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN")
  const findCandidateURL = `https://api.manatal.com/open/v3/matches/${matchID}/`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountHistory = ss.getSheetById("0")

  const orhanizationMap = new Map(accountHistory.getRange("J2:K").getValues().filter(f => f[0] != ""))


  const manatalOptions = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const responseMatch = UrlFetchApp.fetch(findCandidateURL, manatalOptions);
  const statusCode = responseMatch.getResponseCode();
  const responseBody = responseMatch.getContentText();

  Log.info("Status: " + statusCode);
  Log.info("Response: " + responseBody);

  if (statusCode === 200) {
    const data = JSON.parse(responseBody);
    // 👇 You can now access data.candidate or similar
    Log.info("Match ID: " + data.id);

    const getCandidate = data.candidate
    const organization = data.organization || ""


    const organizationName = orhanizationMap.get(organization) || ""

    if (getCandidate && organizationName != "NDA") {

      Utilities.sleep(1000)

      const dataCandidate = getManatalCandidate_(getCandidate)


      const emailBody =
        `Hi!\n candidat ${getCandidate} - ${dataCandidate.full_name} is moved to "Hired" stage. Please be sure that all needed info will complete till tomorrow 00:00.
        \n https://app.manatal.com/candidates/${getCandidate}?tab=more`

      sendEmail_(emailBody, `${dataCandidate.full_name}  - New Hire detected`, "Hire Integration")

      return ""

    }

    return


  } else {
    Log.error("❌ Error fetching match: " + responseBody);
  }
}

function createEmployeeInPeopleForceHired(matchID) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN")
  const findCandidateURL = `https://api.manatal.com/open/v3/matches/${matchID}/`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountHistory = ss.getSheetById("0")


  const orhanizationMap = new Map(accountHistory.getRange("J2:K").getValues().filter(f => f[0] != ""))

  const manatalOptions = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const responseMatch = UrlFetchApp.fetch(findCandidateURL, manatalOptions);
  const statusCode = responseMatch.getResponseCode();
  const responseBody = responseMatch.getContentText();

  Log.info("Status: " + statusCode);
  Log.info("Response: " + responseBody);

  if (statusCode === 200) {
    const data = JSON.parse(responseBody);
    // 👇 You can now access data.candidate or similar
    Log.info("Match ID: " + data.id);

    const getCandidate = data.candidate
    const organization = data.organization || ""


    const organizationName = orhanizationMap.get(organization) || ""

    if (getCandidate && organizationName != "NDA") return createPFEmployeeFromManatalHired_(getCandidate, manatalOptions, matchID)


  } else {
    Log.error("❌ Error fetching match: " + responseBody);
  }
}

function createEmployeeInPeopleForce(matchID) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN")
  const findCandidateURL = `https://api.manatal.com/open/v3/matches/${matchID}/`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountHistory = ss.getSheetById("0")


  const orhanizationMap = new Map(accountHistory.getRange("J2:K").getValues().filter(f => f[0] != ""))

  const manatalOptions = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const responseMatch = UrlFetchApp.fetch(findCandidateURL, manatalOptions);
  const statusCode = responseMatch.getResponseCode();
  const responseBody = responseMatch.getContentText();

  Log.info("Status: " + statusCode);
  Log.info("Response: " + responseBody);

  if (statusCode === 200) {
    const data = JSON.parse(responseBody);
    // 👇 You can now access data.candidate or similar
    Log.info("Match ID: " + data.id);

    const getCandidate = data.candidate
    const organization = data.organization || ""


    const organizationName = orhanizationMap.get(organization) || ""

    if (getCandidate && organizationName != "NDA") return createPFEmployeeFromManatal_(getCandidate, manatalOptions, matchID)


  } else {
    Log.error("❌ Error fetching match: " + responseBody);
  }
}


function createPFEmployeeFromManatalOffer_(candidateID, manatalOptions, matchID, manual = false) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountHistory = ss.getSheetById("0")

  const data = getManatalCandidate_(candidateID)
  let { custom_fields, gender, picture, resume, source_type } = data


  if (!custom_fields?.legalnameukr && !custom_fields?.legalsurnameukr) {

    const emailBody =
      `Hi!\n We faced an error during PeopleForce New hire creation (Name_UKR and Surname_UKR are empty).\n Please check this candidate in Manatal and use Manual tool try one more time - ${getProperty_('MANUAL_HIRE_TOOL_URL', '')}.\n 
 Manatal - https://app.manatal.com/candidates/${candidateID}?tab=more`

    sendEmail_(emailBody, `Name and Surname weren't provided`, "Offer Integration")

    return ""


  }


  const payloadPeopleForce =
  {
    "first_name": custom_fields?.legalnameukr || "",
    "last_name": custom_fields?.legalsurnameukr || "",
    "middle_name": custom_fields?.legalmiddlenameukr || "",
    "gender": gender,
    // "hired_on": hiredDateToAdd || "",
    // "client_id": clientIDMatch ? Number(clientIdMap.get(clientIDMatch)) : "",
    "personal_email": custom_fields?.legalpersonalemail || "",
    // "client_start_date": clientDateToAdd || "",
    "eng_name": custom_fields?.legalnameengpassport || "",
    "eng_surname": custom_fields?.legalsurnameengpassport || "",
    // "telegram": custom_fields?.telegram || "",
    "manatal_id": `${candidateID}¬${matchID}` || "",
    // "invoicing_team": invoiceTeamMatch ? Number(invoiceTeam.get(invoiceTeamMatch)) : "",
    "mobile_number": custom_fields?.legalphonenumber || "",
    // "recruiter_feedback": custom_fields?.recruitersfeedback || "",
    // "candidate_source": source_type || "",
    // "candidate_sub_source": custom_fields?.referralnameifany || "",
    "offer_on_hire": spanExtract_(custom_fields.offerlink) || "",
    "legal_entity_on_hire": custom_fields.typeofemployment,
    "contract_details_on_hire": spanExtract_(custom_fields.detailsforcontract)

  }


  const pf = custom_fields.peopleforceid
  let newEmployeeID = ""

  if (pf) {
    updateEmployee_(pf, payloadPeopleForce)
    newEmployeeID = pf
  }

  else {

    newEmployeeID = createEmployee_(payloadPeopleForce)
    if (!newEmployeeID) return
    let updateFields = { "custom_fields": JSON.parse(JSON.stringify(custom_fields)) }
    updateFields["custom_fields"]["peopleforceid"] = String(newEmployeeID)
    const updateManatal = returnNewIdManatal_(candidateID, updateFields)


  }


  const emailBody =
    `Hi!\n New offer from Manatal to Peopleforce ${manual ? "initited manually" : ""} (${custom_fields?.legalnameukr} ${custom_fields?.legalsurnameukr}) \n 
 Manatal - https://app.manatal.com/candidates/${candidateID}?tab=more ||  PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${newEmployeeID}`

  sendEmail_(emailBody, `${custom_fields?.legalsurnameukr} ${custom_fields?.legalnameukr}`, "Offer Integration")

  return [new Date(), "offer", candidateID, matchID, newEmployeeID]


}


function createPFEmployeeFromManatalHired_(candidateID, manatalOptions, matchID) {


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountHistory = ss.getSheetById("0")


  const positionMap = new Map(accountHistory.getRange("A2:B").getValues().filter(f => f[0] != ""))
  const levelMap = new Map(accountHistory.getRange("D2:E").getValues().filter(f => f[0] != ""))
  const locationMap = new Map(accountHistory.getRange("G2:H").getValues().filter(f => f[0] != ""))
  const invoiceTeam = new Map(accountHistory.getRange("M2:N").getValues().filter(f => f[0] != ""))
  const invoiceTeamValues = Array.from(invoiceTeam.keys())

  const clientIdMap = new Map(accountHistory.getRange("P2:Q").getValues().filter(f => f[0] != ""))
  const clientIdValues = Array.from(clientIdMap.keys())


  const data = getManatalCandidate_(candidateID)

  let { custom_fields, gender, picture, resume, source_type } = data

  const hiredDateToAdd = Utilities.formatDate(new Date(custom_fields.startdate), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")
  const clientDateToAdd = Utilities.formatDate(new Date(custom_fields.startdateclient), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")


  const peopleForceID = custom_fields?.peopleforceid

  if (!peopleForceID || peopleForceID == "null") {


    return createPFEmployeeFromManatal_(candidateID, manatalOptions, matchID)

  }


  const invoiceTeamMatch = invoiceTeamValues.find(f => custom_fields?.teamforinvoicing?.includes(f))
  const clientIDMatch = clientIdValues.find(f => custom_fields?.clientid?.includes(f))

  const payloadPeopleForce =
  {
    "hired_on": hiredDateToAdd || "",
    "client_id_new": clientIDMatch ? Number(clientIdMap.get(clientIDMatch)) : "",
    "client_start_date": clientDateToAdd || "",
    "telegram": custom_fields?.telegram || "",
    "invoicing_team": invoiceTeamMatch ? Number(invoiceTeam.get(invoiceTeamMatch)) : "",
    "recruiter_feedback": custom_fields?.recruitersfeedback || "",
    "candidate_source": source_type || "",
    "candidate_sub_source": custom_fields?.referralnameifany || "",
    "local_address": spanExtract_(custom_fields?.postaddress),
    "rate_after_probation": custom_fields?.rateafterprobation ? Number(custom_fields?.rateafterprobation) : "",


  }


  updateEmployee_(peopleForceID, payloadPeopleForce) // Find employee ID


  //// Position
  const position = String(custom_fields?.position)?.trim() || ""
  const jobLevel = String(custom_fields?.level)?.trim() || ""
  const location = String(custom_fields?.location)?.trim() || ""

  const positionID = positionMap.get(position) || ""
  const jobLevelID = levelMap.get(jobLevel) || ""
  const locationID = locationMap.get(location) || ""

  applyPosition_(peopleForceID, hiredDateToAdd, positionID, jobLevelID, locationID)

  applyStatus_(peopleForceID, hiredDateToAdd)

  /// Salary 
  const rate = custom_fields?.hourlyrate
  if (rate) applySalary_(peopleForceID, hiredDateToAdd, rate)

  const leavePolicyObject =

    PEOPLEFORCE_IDS.leavePolicies.map(p => ({ type_id: p.type_id, policy_id: p.policy_id, effective_on: hiredDateToAdd }))

  applyLeavePolicies_(peopleForceID, leavePolicyObject)


  /// Document
  const docName = `${custom_fields?.fullnamepassport} ${custom_fields?.surnamepassport}` || `${custom_fields?.nameukr} ${custom_fields?.surnameukr}`

  if (resume) createDocument_(resume, PEOPLEFORCE_IDS.resumeDocumentTypeId, docName, peopleForceID)

  /// Avatar
  if (picture) updateAvatar_(picture, peopleForceID)

  const emailBody = `Hi!<br>
New hire from Manatal to Peopleforce (${custom_fields?.legalnameukr} ${custom_fields?.legalsurnameukr})<br>
${custom_fields?.rateafterprobation ? '<span style="color: red; font-weight: bold;">Probation period rate change</span><br>' : ''}
Manatal - https://app.manatal.com/candidates/${candidateID}?tab=more || PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${peopleForceID}`;

  sendEmail_(emailBody, `${custom_fields?.level || ""} ${custom_fields?.position} - ${custom_fields?.legalsurnameukr} ${custom_fields?.legalnameukr} - ${hiredDateToAdd}`, "Hiring Integration", true)

  return [new Date(), "hire", candidateID, matchID, peopleForceID]


}


function createPFEmployeeFromManatal_(candidateID, manatalOptions, matchID) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountHistory = ss.getSheetById("0")

  const positionMap = new Map(accountHistory.getRange("A2:B").getValues().filter(f => f[0] != ""))
  const levelMap = new Map(accountHistory.getRange("D2:E").getValues().filter(f => f[0] != ""))
  const locationMap = new Map(accountHistory.getRange("G2:H").getValues().filter(f => f[0] != ""))
  const invoiceTeam = new Map(accountHistory.getRange("M2:N").getValues().filter(f => f[0] != ""))
  const invoiceTeamValues = Array.from(invoiceTeam.keys())

  const clientIdMap = new Map(accountHistory.getRange("P2:Q").getValues().filter(f => f[0] != ""))
  const clientIdValues = Array.from(clientIdMap.keys())


  const data = getManatalCandidate_(candidateID)

  let { custom_fields, gender, picture, resume, source_type } = data

  const hiredDateToAdd = Utilities.formatDate(new Date(custom_fields.startdate), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")
  const clientDateToAdd = Utilities.formatDate(new Date(custom_fields.startdateclient), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  if (!custom_fields?.legalnameukr && !custom_fields?.legalsurnameukr) {

    const emailBody =
      `Hi!\n We faced an error during PeopleForce New hire creation (Name_UKR and Surname_UKR are empty).\n Please check this candidate in Manatal and use Manual tool try one more time - ${getProperty_('MANUAL_HIRE_TOOL_URL', '')}.\n 
 Manatal - https://app.manatal.com/candidates/${candidateID}?tab=more`

    sendEmail_(emailBody, `Name and Surname weren't provided`, "Hiring Integration")

    return ""


  }

  const invoiceTeamMatch = invoiceTeamValues.find(f => custom_fields?.teamforinvoicing?.includes(f))
  const clientIDMatch = clientIdValues.find(f => custom_fields?.clientid?.includes(f))

  const payloadPeopleForce =
  {
    "first_name": custom_fields?.legalnameukr || "",
    "last_name": custom_fields?.legalsurnameukr || "",
    "middle_name": custom_fields?.legalmiddlenameukr || "",
    "gender": gender,
    "hired_on": hiredDateToAdd || "",
    "client_id_new": clientIDMatch ? Number(clientIdMap.get(clientIDMatch)) : "",
    "email": custom_fields?.legalpersonalemail || "",
    "personal_email": custom_fields?.legalpersonalemail || "",
    "client_start_date": clientDateToAdd || "",
    "eng_name": custom_fields?.legalnameengpassport || "",
    "eng_surname": custom_fields?.legalsurnameengpassport || "",
    "telegram": custom_fields?.telegram || "",
    "manatal_id": `${candidateID}¬${matchID}` || "",
    "invoicing_team": invoiceTeamMatch ? Number(invoiceTeam.get(invoiceTeamMatch)) : "",
    "mobile_number": custom_fields?.phonenumber || "",
    "recruiter_feedback": custom_fields?.recruitersfeedback || "",
    "candidate_source": source_type || "",
    "candidate_sub_source": custom_fields?.referralnameifany || "",
    "offer_on_hire": spanExtract_(custom_fields.offerlink) || "",
    "legal_entity_on_hire": custom_fields?.typeofemployment || "",
    "contract_details_on_hire": spanExtract_(custom_fields.detailsforcontract) || "",
    "local_address": spanExtract_(custom_fields?.postaddress),
    "rate_after_probation": custom_fields?.rateafterprobation ? Number(custom_fields?.rateafterprobation) : "",


  }

  const newEmployeeID = createEmployee_(payloadPeopleForce)
  let updateFields = { "custom_fields": JSON.parse(JSON.stringify(custom_fields)) }
  updateFields["custom_fields"]["peopleforceid"] = String(newEmployeeID)
  const updateManatal = returnNewIdManatal_(candidateID, updateFields)


  //// Position
  const position = String(custom_fields?.position)?.trim() || ""
  const jobLevel = String(custom_fields?.level)?.trim() || ""
  const location = String(custom_fields?.location)?.trim() || ""

  const positionID = positionMap.get(position) || ""
  const jobLevelID = levelMap.get(jobLevel) || ""
  const locationID = locationMap.get(location) || ""

  applyPosition_(newEmployeeID, hiredDateToAdd, positionID, jobLevelID, locationID)

  applyStatus_(newEmployeeID, hiredDateToAdd)

  /// Salary 
  const rate = custom_fields?.hourlyrate
  if (rate) applySalary_(newEmployeeID, hiredDateToAdd, rate)

  const leavePolicyObject =

    PEOPLEFORCE_IDS.leavePolicies.map(p => ({ type_id: p.type_id, policy_id: p.policy_id, effective_on: hiredDateToAdd }))

  applyLeavePolicies_(newEmployeeID, leavePolicyObject)


  /// Document
  const docName = `${custom_fields?.fullnamepassport} ${custom_fields?.surnamepassport}` || `${custom_fields?.nameukr} ${custom_fields?.surnameukr}`

  if (resume) createDocument_(resume, PEOPLEFORCE_IDS.resumeDocumentTypeId, docName, newEmployeeID)

  /// Avatar
  if (picture) updateAvatar_(picture, newEmployeeID)

  //   const emailBody =
  //     `Hi!\n New hire from Manatal to Peopleforce (${custom_fields?.nameukr} ${custom_fields?.surnameukr}) \n 
  //  Manatal - https://app.manatal.com/candidates/${candidateID}?tab=more ||  PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${newEmployeeID}`

  const emailBody = `Hi!<br>
New hire from Manatal to Peopleforce (${custom_fields?.legalnameukr} ${custom_fields?.legalsurnameukr})<br>
${custom_fields?.rateafterprobation ? '<span style="color: red; font-weight: bold;">Probation period rate change</span><br>' : ''}
Manatal - https://app.manatal.com/candidates/${candidateID}?tab=more || PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${newEmployeeID}`;

  sendEmail_(emailBody, `${custom_fields?.level || ""} ${custom_fields?.position} - ${custom_fields?.surnameukr} ${custom_fields?.nameukr} - ${hiredDateToAdd}`, "Hiring Integration", true)

  return [new Date(), "hire", candidateID, matchID, newEmployeeID]


}


function createDocument_(doc, folderId, name, employeeId) {
  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/${employeeId}/documents`;

  const payload = {
    document_folder_id: Number(folderId),
    name: name,
    url: doc
  };

  const options = {
    method: 'post',
    headers: { 'X-API-KEY': PFKEY, 'Content-Type': 'application/json' },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(URL, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code >= 200 && code < 300) {
    try {
      Log.info(`Document created successfully for Employee ID: ${employeeId}.`);
      return JSON.parse(body);
    } catch (e) {
      Log.error(`Success, but failed to parse response body.`);
      return { success: true };
    }
  } else {
    Log.error(`ERROR creating document for Employee ID ${employeeId} (${code}): ${body}`);
    return null;
  }
}

function updateAvatar_(pic, employeeId) {
  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/${employeeId}/update_avatar`;

  try {
    // 1. Fetch image and get Blob
    const blob = UrlFetchApp.fetch(pic).getBlob();

    // 2. Configure options for PUT file upload (multipart/form-data)
    const options = {
      method: 'put',
      headers: { 'X-API-KEY': PFKEY },
      muteHttpExceptions: true,
      // CRITICAL: Payload must be an object containing the Blob, NOT a stringified JSON
      payload: { avatar: blob }
    };

    const response = UrlFetchApp.fetch(URL, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code >= 200 && code < 300) {
      Log.info(`Avatar updated successfully for ID: ${employeeId}.`);
      return JSON.parse(body);
    } else {
      Log.error(`ERROR updating avatar (${code}): ${body}`);
      return null;
    }
  } catch (e) {
    Log.error(`Script error during avatar update for ID ${employeeId}: ${e.toString()}`);
    return null;
  }
}


function createEmployee_(payload) {
  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = "https://app.peopleforce.io/api/public/v3/employees";

  const options = {
    method: 'post',
    // Remove 'Content-Type': 'application/json'—it's redundant for UrlFetchApp with JSON payload
    headers: {
      'X-API-KEY': PFKEY,
      'Content-Type': 'application/json; charset=utf-8'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Capture response for status check
  };

  const response = UrlFetchApp.fetch(URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 200 && responseCode < 300) {
    // Success: Parse and return the new employee ID
    try {
      return JSON.parse(responseText).data.id;
    } catch (e) {
      Log.error("Failed to parse successful response: " + responseText);
      return null;
    }
  } else {
    // Failure: Log error and return null
    writeLog('createEmployee_', `Employee creation failed (${responseCode}): ${responseText}`)
    Log.error(`Employee creation failed (${responseCode}): ${responseText}`);
    return null;
  }
}

function updateEmployee_(id, payload) {
  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/${id}`;

  const options = {
    method: 'put',
    // Remove 'Content-Type': 'application/json'—it's redundant for UrlFetchApp with JSON payload
    headers: {
      'X-API-KEY': PFKEY,
      'Content-Type': 'application/json; charset=utf-8'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Capture response for status check
  };

  const response = UrlFetchApp.fetch(URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 200 && responseCode < 300) {
    // Success: Parse and return the new employee ID
    try {
      return JSON.parse(responseText).data.id;
    } catch (e) {
      Log.error("Failed to parse successful response: " + responseText);
      return null;
    }
  } else {
    // Failure: Log error and return null
    writeLog('updateEmployee_', `Employee update failed (${responseCode}): ${responseText}`)
    Log.error(`Employee update failed (${responseCode}): ${responseText}`);
    return null;
  }
}

function returnNewIdManatal_(candidateId, custom_fields) {
  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");
  const URL = `https://api.manatal.com/open/v3/candidates/${candidateId}/`;

  const options = {
    method: 'patch',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(custom_fields)
  };

  const response = UrlFetchApp.fetch(URL, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code >= 200 && code < 300) {
    // Success: Return the parsed response
    try {
      return JSON.parse(body);
    } catch (e) {
      Log.error(`Success, but failed to parse response body.`);
      return { success: true };
    }
  } else {
    // Failure: Log error and return null
    Log.error(`ERROR updating Manatal Candidate ${candidateId} (${code}): ${body}`);
    return null;
  }
}

function applyPosition_(employee_id, effective_from, position_id, job_level_id, location_id) {

  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/${employee_id}/positions`;

  const payload =
  {
    "effective_from": effective_from,
    ...(position_id && { "position_id": Number(position_id) }),
    ...(job_level_id && { "job_level_id": Number(job_level_id) }),
    ...(location_id && { "location_id": Number(location_id) })
  }


  const options = {
    method: 'post',
    // Remove 'Content-Type': 'application/json'—it's redundant for UrlFetchApp with JSON payload
    headers: {
      'X-API-KEY': PFKEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Capture response for status check
  };

  const response = UrlFetchApp.fetch(URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 200 && responseCode < 300) {
    // Success: Parse and return the new employee ID
    try {
      return JSON.parse(responseText).data.id;
    } catch (e) {
      Log.error("Failed to create a position successful response: " + responseText);
      return null;
    }
  } else {
    // Failure: Log error and return null
    Log.error(`Position applying failed (${responseCode}): ${responseText}`);
    return null;
  }
}

function applyLeavePolicies_(employee_id, arrayOfPolicies) {

  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/${employee_id}/leave_types`;

  arrayOfPolicies.forEach(policy => {

    const payload =
    {
      "effective_on": policy?.effective_on || "",
      "leave_type_id": policy?.type_id || "",
      "leave_type_policy_id": policy?.policy_id || "",
    }


    const options = {
      method: 'post',
      // Remove 'Content-Type': 'application/json'—it's redundant for UrlFetchApp with JSON payload
      headers: {
        'X-API-KEY': PFKEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // Capture response for status check
    };

    const response = UrlFetchApp.fetch(URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      // Success: Parse and return the new employee ID
      try {
        return JSON.parse(responseText).data.id;
      } catch (e) {
        Log.error("Failed to apply a leave policy successful response: " + responseText);
        return null;
      }
    } else {
      // Failure: Log error and return null
      Log.error(`Leave policy applying failed (${responseCode}): ${responseText}`);
      return null;
    }


  })


}

function applyStatus_(employee_id, effective_from) {

  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/${employee_id}/employment_statuses`;

  const payload =
  {
    "effective_from": effective_from,
    "probation_policy_id": PEOPLEFORCE_IDS.probationPolicyId,
    "employment_type_id": PEOPLEFORCE_IDS.employmentTypeId,
    "working_pattern_id": PEOPLEFORCE_IDS.workingPatternId
  }


  const options = {
    method: 'post',
    // Remove 'Content-Type': 'application/json'—it's redundant for UrlFetchApp with JSON payload
    headers: {
      'X-API-KEY': PFKEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Capture response for status check
  };

  const response = UrlFetchApp.fetch(URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 200 && responseCode < 300) {
    // Success: Parse and return the new employee ID
    try {
      return JSON.parse(responseText).data.id;
    } catch (e) {
      Log.error("Failed to create a employment statuses successful response: " + responseText);
      return null;
    }
  } else {
    // Failure: Log error and return null
    Log.error(`Employment statuses applying failed (${responseCode}): ${responseText}`);
    return null;
  }
}


function applySalary_(employee_id, effective_from, ratePerHour) {

  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/${employee_id}/salaries`;

  const payload =
  {
    "effective_on": effective_from,
    "per": "hour",
    "currency_code": "EUR",
    "amount": parseFloat(ratePerHour),
    "pay_schedule_id": PEOPLEFORCE_IDS.payScheduleId
  }


  const options = {
    method: 'post',
    // Remove 'Content-Type': 'application/json'—it's redundant for UrlFetchApp with JSON payload
    headers: {
      'X-API-KEY': PFKEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Capture response for status check
  };

  const response = UrlFetchApp.fetch(URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 200 && responseCode < 300) {
    // Success: Parse and return the new employee ID
    try {
      return JSON.parse(responseText).data.id;
    } catch (e) {
      Log.error("Failed to create a salary successful response: " + responseText);
      return null;
    }
  } else {
    // Failure: Log error and return null
    Log.error(`Salary applying failed (${responseCode}): ${responseText}`);
    return null;
  }
}

function updateManatalMatch_(matchID, payload) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");
  const URL = `https://api.manatal.com/open/v3/matches/${matchID}/`;

  const options = {
    method: 'patch',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(URL, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code >= 200 && code < 300) {
    // Success: Return the parsed response
    try {
      return JSON.parse(body);
    } catch (e) {
      Log.error(`Success, but failed to parse response body.`);
      return { success: true };
    }
  } else {
    // Failure: Log error and return null
    Log.error(`ERROR updating Manatal Match ${matchID} (${code}): ${body}`);
    return null;
  }
}

function getManatalCandidate_(candidateID) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
  };

  const candidateURL = `https://api.manatal.com/open/v3/candidates/${candidateID}/`;
  const responseCandidate = UrlFetchApp.fetch(candidateURL, options);
  const statusCode = responseCandidate.getResponseCode();
  const responseBody = responseCandidate.getContentText();

  Log.info("Status: " + statusCode);
  Log.info("Response: " + responseBody);

  if (statusCode === 200) {
    const data = JSON.parse(responseBody);
    // 👇 You can now access data.candidate or similar
    Log.info("Candidate id: " + data.id);
    return data
  }

  else {
    Log.error("❌ Error fetching candidate: " + responseBody);
    return responseBody
  }


}


function getManatalMatch_(matchID) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
  };

  const candidateURL = `https://api.manatal.com/open/v3/matches/${candidateID}/`;
  const responseCandidate = UrlFetchApp.fetch(candidateURL, options);
  const statusCode = responseCandidate.getResponseCode();
  const responseBody = responseCandidate.getContentText();

  Log.info("Status: " + statusCode);
  Log.info("Response: " + responseBody);

  if (statusCode === 200) {
    const data = JSON.parse(responseBody);
    // 👇 You can now access data.candidate or similar
    Log.info("Candidate id: " + data.id);
    return data
  }

  else {
    Log.error("❌ Error fetching candidate: " + responseBody);
    return responseBody
  }


}

function identifyManatalMatchByCandidate_(candidateID) {

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
  };

  const candidateURL = `https://api.manatal.com/open/v3/candidates/${candidateID}/matches/`;
  const responseCandidate = UrlFetchApp.fetch(candidateURL, options);
  const statusCode = responseCandidate.getResponseCode();
  const responseBody = responseCandidate.getContentText();

  Log.info("Status: " + statusCode);
  Log.info("Response: " + responseBody);

  if (statusCode === 200) {
    const data = JSON.parse(responseBody);
    // 👇 You can now access data.candidate or similar
    Log.info("Candidate id: " + data.results?.[0]?.id);
    return data.results?.[0]?.id || ""
  }

  else {
    Log.error("❌ Error fetching candidate: " + responseBody);
    return responseBody
  }


}


/**
 * Main entry point for the Manatal Webhook with Deduplication.
 */
function doPost(e) {
  Log.info('doPost: request received', { hasBody: !!(e && e.postData), params: (e && e.parameter) ? Object.keys(e.parameter).filter(function (k) { return k !== 'token'; }) : [] });
  var lock = LockService.getScriptLock();
  var cache = CacheService.getScriptCache();

  try {
    // 1. Acquire lock (wait up to 5 seconds)
    // We keep this short so we don't hold up the webhook response line
    lock.waitLock(5000);

    // 2. Parse payload
    var payload = JSON.parse(e.postData.contents);

    // 3. Deduplication Check
    // Manatal usually provides a unique 'id' for every webhook event
    var eventId = `${payload.id}_${payload.job_pipeline_stage}_${payload.created_at}` || "no_id_found";

    if (cache.get(eventId) !== null) {
      Log.warn("Duplicate request detected for ID: " + eventId);
      return ContentService.createTextOutput(JSON.stringify({ "status": "ignored", "reason": "duplicate" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Logic Execution
    if (payload && payload.job_pipeline_stage) {
      handleMatchStatusChange(payload);
    }

    // 5. Mark as processed in Cache (valid for 20 minutes / 1200 seconds)
    cache.put(eventId, "processed", 1200);

    return ContentService.createTextOutput(JSON.stringify({ "status": "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Log.error("Execution Error: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    // 6. CRITICAL: Release the lock so the next request can enter
    lock.releaseLock();
  }
}

/**
 * Helper function to write to the 'Logs' sheet.
 */
function writeLog(functionName, message) {
  var ss = SpreadsheetApp.openById(getRequiredProperty_('HIRING_SPREADSHEET_ID'));
  var logSheet = ss.getSheetByName("Logs");

  if (!logSheet) {
    logSheet = ss.insertSheet("Logs");
    logSheet.appendRow(["Timestamp", "Function", "Message"]);
    logSheet.getRange("A1:C1").setFontWeight("bold");
  }

  logSheet.appendRow([new Date(), functionName, message]);
}


/**
 * Logic to run when a match status changes
 */
function handleMatchStatusChange(data) {
  var matchId = data.id;
  var newStatus = data.job_pipeline_stage;
  const ss = SpreadsheetApp.openById(getRequiredProperty_('HIRING_SPREADSHEET_ID'))
  const historySheet = ss
    .getSheetByName("Hired history")


  // Basic validation to ensure we have a status
  if (!newStatus) return;

  Log.info("Processing status: " + newStatus + " for Match ID: " + matchId);

  let record;

  switch (newStatus) {
    case "Offered":
      writeLog("payload_Offered", JSON.stringify(data))
      record = createEmployeeInPeopleForceOffer(matchId);
      break;

    case "Hired":
      writeLog("payload_Hired", JSON.stringify(data))
      record = messagePeopleForceHired(matchId);
      break;

    default:
      Log.warn("Status " + newStatus + " is not a trigger stage. Skipping.");
  }


  if (record && Array.isArray(record)) {
    try {
      // appendRow is safer than getRange(getLastRow + 1) for concurrent scripts
      if (newStatus = "Offered") historySheet.appendRow(record);
      Log.info("Successfully logged record for Match ID: " + matchId);
    } catch (e) {
      Log.error("Failed to write to sheet: " + e.message);
    }
  }
}


function checkOffered() {

  // getPositionHandbook()


  const today = new Date()
  const toURL = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")
  today.setDate(today.getDate() - 1)
  const fromURL = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  const ss = SpreadsheetApp.getActive()
  const historySheet = ss.getSheetByName("Hired history")
  const history = historySheet.getRange("D2:D").getValues().filter(f => f[0] != "").flat()


  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");

  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
  };

  let page = 1
  let nextPage = true

  const matchesOffered = []
  const offeredInfo = []

  while (nextPage) {
    const URL = `https://api.manatal.com/open/v3/matches/?stage__in=${MANATAL_STAGE_IDS.offer}&offer_at__gte=${fromURL}&offer_at__lte=${toURL}&page=${page}`;
    const response = UrlFetchApp.fetch(URL, options);
    const data = JSON.parse(response.getContentText())

    nextPage = data?.next

    if (nextPage) page++

    matchesOffered.push(...data?.results)
  }


  matchesOffered
    .filter(f => !history.includes(f.id))
    .forEach(match => {

      try {

        // 👉 Run your next action here (e.g., send email, update sheet, etc.)
        const recordToHire = createEmployeeInPeopleForceOffer(match.id);
        if (recordToHire?.length) offeredInfo.push(recordToHire)


      } catch (error) {
        Log.error("❌ Error: " + error);
      }


    })


  if (offeredInfo.length) {

    const historySheet = ss
      .getSheetByName("Hired history")

    historySheet.getRange(historySheet.getLastRow() + 1, 1, offeredInfo.length, offeredInfo[0].length).setValues(offeredInfo)


  }

}


function checkHiredNew() {

  // getPositionHandbook()
  const today = new Date()
  const toURL = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")
  today.setDate(today.getDate() - 1)
  const fromURL = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  const ss = SpreadsheetApp.getActive()
  const historySheet = ss.getSheetByName("Hired history")
  const history = historySheet.getRange("B2:D").getValues()
    .filter(f => f[0] != "" && f[0] != "offer")
    .map(m => m[2])
    .flat()


  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");

  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
  };

  let page = 1
  let nextPage = true

  const matchesOffered = []
  const hiredInfo = []

  while (nextPage) {
    const URL = `https://api.manatal.com/open/v3/matches/?stage__in=${MANATAL_STAGE_IDS.hired}&hired_at__gte=${fromURL}&hired_at__lte=${toURL}&page=${page}`;
    const response = UrlFetchApp.fetch(URL, options);
    const data = JSON.parse(response.getContentText())

    nextPage = data?.next

    if (nextPage) page++

    matchesOffered.push(...data?.results)
  }


  matchesOffered
    .filter(f => !history.includes(f.id))
    .forEach(match => {

      try {

        // 👉 Run your next action here (e.g., send email, update sheet, etc.)
        const recordToHire = createEmployeeInPeopleForceHired(match.id);
        if (recordToHire?.length) hiredInfo.push(recordToHire)


      } catch (error) {
        Log.error("❌ Error: " + error);
      }


    })


  if (hiredInfo.length) {

    const historySheet = ss
      .getSheetByName("Hired history")

    historySheet.getRange(historySheet.getLastRow() + 1, 1, hiredInfo.length, hiredInfo[0].length).setValues(hiredInfo)


  }


}


function checkHired() {

  // checkOffered()
  // Utilities.sleep(1000)
  checkHiredNew()
  getPositionHandbook()
  ScriptApp.newTrigger("checkProbation")
    .timeBased()
    .after(2 * 1000) // run after 5 seconds
    .create();


}


// function checkHired() {

//   getPositionHandbook()

//   const today = new Date()


//   const toURL = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

//   today.setDate(today.getDate() - 1)

//   const fromURL = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")


//   getPositionHandbook()


//   const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");

//   const options = {
//     method: 'get',
//     headers: {
//       'Authorization': 'Token ' + MANATAL_KEY,
//       'Content-Type': 'application/json'
//     },
//     muteHttpExceptions: true,
//   };

//   let page = 1
//   let nextPage = true

//   const matchesOffered = []
//   const hiredInfo = []


//   // {
//   //     "name": "Offered",
//   //       "rank": 7
//   // },

//   while (nextPage) {
//     const URL = `https://api.manatal.com/open/v3/matches/?stage__in=${MANATAL_STAGE_IDS.hired}&hired_at__gte=${fromURL}&hired_at__lte=${toURL}&page=${page}`;
//     const response = UrlFetchApp.fetch(URL, options);
//     const data = JSON.parse(response.getContentText())

//     nextPage = data?.next

//     if (nextPage) page++

//     matchesOffered.push(...data?.results)
//   }


//   matchesOffered.forEach(match => {

//     try {

//       // 👉 Run your next action here (e.g., send email, update sheet, etc.)
//       const recordToHire = createEmployeeInPeopleForce(match.id);
//       if (recordToHire?.length) hiredInfo.push(recordToHire)


//     } catch (error) {
//       Logger.log("❌ Error: " + error);
//     }


//   })


//   if (hiredInfo.length) {

//     const historySheet = SpreadsheetApp.getActiveSpreadsheet()
//       .getSheetByName("Hired history")

//     historySheet.getRange(historySheet.getLastRow() + 1, 1, hiredInfo.length, hiredInfo[0].length).setValues(hiredInfo)


//   }

//   ScriptApp.newTrigger("checkProbation")
//     .timeBased()
//     .after(2 * 1000) // run after 5 seconds
//     .create();


function checkProbation(e) {

  const from = new Date()
  from.setMonth(from.getMonth() - 3)
  from.setDate(from.getDate() - 1)
  const fromURL = Utilities.formatDate(from, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  // const to = new Date()
  // to.setMonth(to.getMonth() - 3)
  // const toURL = Utilities.formatDate(to, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees?status=active&page=1&hired_on[gte]=${fromURL}&hired_on[lte]=${fromURL}`;


  const options = {
    method: 'get',
    headers: { 'X-API-KEY': PFKEY },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(URL, options);
  const data = JSON.parse(response.getContentText())?.data

  if (data.length) {

    Log.info(`Find ${data.length} members that passed Probation Period`)

    data.forEach(({ id }, i) => {


      const urlEmployee = `https://app.peopleforce.io/api/public/v3/employees/${id}`
      const emplResponse = UrlFetchApp.fetch(urlEmployee, options);
      const emplData = JSON.parse(emplResponse.getContentText())?.data


      const name = emplData?.full_name
      const getManatalID = emplData?.fields?.["manatal_id"]?.value || ""
      if (getManatalID) {


        const [candidate, match] = getManatalID?.split("¬")

        if (match) {


          const updatePayload =
            { "stage": { "id": MANATAL_STAGE_IDS.probationPassed } }

          updateManatalMatch_(match, updatePayload)

          const emailBody =
            `Hi!\n${name} passed Probation period and moved to the Probation passed stage in Manatal (Match ID ${match}) \n 
 Manatal - https://app.manatal.com/candidates/${candidate}?tab=more ||  PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${id}`

          sendEmail_(emailBody, "[Notif] - Probation passed", "HR Integration")


        }


      }

    })


  }

  if (e && e.triggerUid) {
    ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid));
  }

  ScriptApp.newTrigger("checkTermination")
    .timeBased()
    .after(2 * 1000) // run after 5 seconds
    .create();

}


function checkTermination(e) {

  const from = new Date()
  from.setDate(from.getDate() - 1)
  const fromURL = Utilities.formatDate(from, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")


  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/terminated?terminated_on[gte]=${fromURL}&terminated_on[lte]=${fromURL}&page=1`;


  const options = {
    method: 'get',
    headers: { 'X-API-KEY': PFKEY },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(URL, options);
  const data = JSON.parse(response.getContentText())?.data
  // ?.filter(f => !isMoreThanThreeMonths_(f.hired_on, fromURL))

  if (data.length) {


    Log.info(`Find ${data.length} members that not passed Probation Period`)

    data.forEach((info, i) => {

      const terminationComment = info.termination_comment
      const terminationType = info.termination_type?.name
      const terminationReason = info.termination_reason?.name
      const terminationDate = info.termination_effective_on

      const urlEmployee = `https://app.peopleforce.io/api/public/v3/employees/${info.id}`
      const emplResponse = UrlFetchApp.fetch(urlEmployee, options);
      const emplData = JSON.parse(emplResponse.getContentText())?.data

      const getManatalID = emplData?.fields?.["manatal_id"]?.value || ""
      if (getManatalID) {

        const [candidate, match] = getManatalID.split("¬")

        if (match) {

          const updatePayload =
            { "stage": { "id": MANATAL_STAGE_IDS.terminated } }

          updateManatalMatch_(match, updatePayload)

          const emailBody =
            `Hi!\n${name} not passed Probation period and moved to the Probation passed stage in Manatal (Match ID ${match}) \n 
 Manatal - https://app.manatal.com/candidates/${candidate}?tab=more ||  PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${id}`

          sendEmail_(emailBody, "[Notif] - Probation passed", "HR Integration")


        }

        if (candidate) {
          const getCustomFields = getManatalCandidate_(candidate)?.custom_fields
          getCustomFields.terminationdate = `${terminationDate}`
          getCustomFields.terminationtype = terminationType
          getCustomFields.terminationreason = terminationReason
          getCustomFields.terminationcomments = terminationComment

          const payload = { "custom_fields": getCustomFields }
          returnNewIdManatal_(candidate, payload)

        }


      }


    })


  }

  if (e && e.triggerUid) {
    ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid));
  }

  ScriptApp.newTrigger("laptopCheck")
    .timeBased()
    .after(2 * 1000) // run after 5 seconds
    .create();


}

// function checkProbation(e) {

//   const from = new Date()
//   from.setMonth(from.getMonth() - 3)
//   from.setDate(from.getDate() - 1)
//   const fromURL = Utilities.formatDate(from, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

//   // const to = new Date()
//   // to.setMonth(to.getMonth() - 3)
//   // const toURL = Utilities.formatDate(to, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

//   const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
//   const URL = `https://app.peopleforce.io/api/public/v3/employees?status=active&page=1&hired_on[gte]=${fromURL}&hired_on[lte]=${fromURL}`;


//   const options = {
//     method: 'get',
//     headers: { 'X-API-KEY': PFKEY },
//     muteHttpExceptions: true,
//   };

//   const response = UrlFetchApp.fetch(URL, options);
//   const data = JSON.parse(response.getContentText())?.data

//   if (data.length) {

//     Logger.log(`Find ${data.length} members that passed Probation Period`)

//     data.forEach(({ id }, i) => {


//       const urlEmployee = `https://app.peopleforce.io/api/public/v3/employees/${id}`
//       const emplResponse = UrlFetchApp.fetch(urlEmployee, options);
//       const emplData = JSON.parse(emplResponse.getContentText())?.data


//       const name = emplData?.full_name
//       const getManatalID = emplData?.fields?.["manatal_id"]?.value || ""
//       if (getManatalID) {

//         const [candidate, match] = getManatalID?.split("¬")

//         if (match) {

//           const updatePayload =
//             { "stage": { "id": MANATAL_STAGE_IDS.customProbationPassed } }

//           updateManatalMatch_(match, updatePayload)

//           const emailBody =
//             `Hi!\n${name} passed Probation period and moved to the Probation passed stage in Manatal (Match ID ${match}) \n 
//  Manatal - https://app.manatal.com/candidates/${candidate}?tab=more ||  PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${id}`

//           sendEmail_(emailBody, "[Notif] - Probation passed", "HR Integration")

//         }


//       }


//     })


//   }

//   if (e && e.triggerUid) {
//     ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid));
//   }

//   ScriptApp.newTrigger("checkTermination")
//     .timeBased()
//     .after(2 * 1000) // run after 5 seconds
//     .create();

// }


function customCheckProbation(e) {

  if (e && e.triggerUid) {
    ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid));
  }

  const from = new Date()
  from.setMonth(from.getMonth() - 3)
  from.setDate(from.getDate() - 2)
  const fromURL = Utilities.formatDate(from, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  // const to = new Date()
  // to.setMonth(to.getMonth() - 3)
  // const toURL = Utilities.formatDate(to, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees?status=active&page=1&hired_on[gte]=${fromURL}&hired_on[lte]=${fromURL}`;


  const options = {
    method: 'get',
    headers: { 'X-API-KEY': PFKEY },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(URL, options);
  const data = JSON.parse(response.getContentText())?.data

  if (data.length) {

    Log.info(`Find ${data.length} members that passed Probation Period`)

    data.forEach(({ id }, i) => {


      const urlEmployee = `https://app.peopleforce.io/api/public/v3/employees/${id}`
      const emplResponse = UrlFetchApp.fetch(urlEmployee, options);
      const emplData = JSON.parse(emplResponse.getContentText())?.data


      const name = emplData?.full_name
      const getManatalID = emplData?.fields?.["manatal_id"]?.value || ""
      if (getManatalID) {

        const [candidate, match] = getManatalID?.split("¬")

        if (match) {

          const updatePayload =
            { "stage": { "id": MANATAL_STAGE_IDS.customProbationPassed } }

          updateManatalMatch_(match, updatePayload)

          const emailBody =
            `Hi!\n${name} passed Probation period and moved to the Probation passed stage in Manatal (Match ID ${match}) \n 
 Manatal - https://app.manatal.com/candidates/${candidate}?tab=more ||  PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${id}`

          sendEmail_(emailBody, "[Notif] - Probation passed", "HR Integration")

        }


      }


    })


  }

  if (e && e.triggerUid) {
    ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid));
  }

  ScriptApp.newTrigger("checkTermination")
    .timeBased()
    .after(2 * 1000) // run after 5 seconds
    .create();

}

function customCheckTermination(e) {

  const from = new Date()
  from.setDate(from.getDate() - 2)
  const fromURL = Utilities.formatDate(from, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")

  if (e && e.triggerUid) {
    ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid));
  }


  const PFKEY = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");
  const URL = `https://app.peopleforce.io/api/public/v3/employees/terminated?terminated_on[gte]=${fromURL}&terminated_on[lte]=${fromURL}&page=1`;


  const options = {
    method: 'get',
    headers: { 'X-API-KEY': PFKEY },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(URL, options);
  const data = JSON.parse(response.getContentText())?.data?.filter(f => !isMoreThanThreeMonths_(f.hired_on, fromURL))

  if (data.length) {


    Log.info(`Find ${data.length} members that not passed Probation Period`)

    data.forEach((info, i) => {

      const terminationComment = info.termination_comment
      const terminationType = info.termination_type?.name
      const terminationReason = info.termination_reason?.name
      const terminationDate = info.termination_effective_on

      const urlEmployee = `https://app.peopleforce.io/api/public/v3/employees/${info.id}`
      const emplResponse = UrlFetchApp.fetch(urlEmployee, options);
      const emplData = JSON.parse(emplResponse.getContentText())?.data

      const getManatalID = emplData?.fields?.["manatal_id"]?.value || ""
      if (getManatalID) {

        const [candidate, match] = getManatalID.split("¬")

        if (match) {

          const updatePayload =
            { "stage": { "id": MANATAL_STAGE_IDS.terminated } }

          updateManatalMatch_(match, updatePayload)

          const emailBody =
            `Hi!\n${info.full_name} not passed Probation period and moved to the Probation not passed stage in Manatal (Match ID ${match}) \n 
 Manatal - https://app.manatal.com/candidates/${candidate}?tab=more ||  PeopleForce - ${getPeopleForceBaseUrl_()}/employees/${info.id}`

          sendEmail_(emailBody, "[Notif] - Probation passed", "HR Integration")

        }

        if (candidate) {
          const getCustomFields = getManatalCandidate_(candidate)?.custom_fields
          getCustomFields.terminationdate = `${terminationDate}`
          getCustomFields.terminationtype = terminationType
          getCustomFields.terminationreason = terminationReason
          getCustomFields.terminationcomments = terminationComment

          const payload = { "custom_fields": getCustomFields }
          returnNewIdManatal_(candidate, payload)

        }


      }


    })


  }

}


function laptopCheck(e) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hiredHistory = ss.getSheetByName("Hired history")
  const historySheet = hiredHistory.getRange("A2:F").getValues().map((m, i) => [...m, i]).filter(f => f[0] != "" && f[5] == "" && f[4] != "" && f[1] == "hired")

  historySheet.forEach(record => {


    const candidateData = getManatalCandidate_(record[1])
    const laptopData = candidateData?.custom_fields?.laptopsn || ""

    if (laptopData) {

      updateEmployee_(record[3], { "equipment_on_first_day": laptopData })

      hiredHistory.getRange(record[record.length - 1] + 2, 6).setValue(new Date())

    }

  })


  if (e && e.triggerUid) {
    ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid));
  }

}

function manualNewHire() {

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const manualHireSheet = ss.getSheetById(Number(getRequiredProperty_('MANUAL_HIRE_SHEET_GID')))
  const historySheet = ss.getSheetById(Number(getRequiredProperty_('MANUAL_HIRE_HISTORY_SHEET_GID')))

  const getCandidateID = manualHireSheet.getRange("B3")
  const getCandidateType = manualHireSheet.getRange("B2")


  if (!getCandidateID.getValue()) {

    SpreadsheetApp.getActive().toast("⚠️ Manatal ID wasn't provided")
    return ""
  }

  if (!getCandidateType.getValue()) {

    SpreadsheetApp.getActive().toast("⚠️ Manatal Stage wasn't provided")
    return ""
  }


  const findMatch = identifyManatalMatchByCandidate_(getCandidateID.getValue())


  if (!findMatch) {

    SpreadsheetApp.getActive().toast("⚠️ Manatal Match wasn't find")
    return ""

  }


  switch (getCandidateType.getValue()) {

    case "Offer":

      const newOffer = createEmployeeInPeopleForceOffer(findMatch, true)
      if (newOffer.length) {
        historySheet.appendRow(newOffer)


      }

      break;

    case "Hire":
      const newHire = createEmployeeInPeopleForceHired(findMatch)
      if (newHire.length) {
        historySheet.appendRow(newHire)

      }

      break


  }

  createEmployeeInPeopleForceHired(findMatch);
  getCandidateID.clearContent()
  getCandidateType.clearContent()
}


function getPositionHandbook(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const positionsHandbook = ss.getSheetById("0")

  const key = PropertiesService.getScriptProperties().getProperty("PEOPLEFORCE_API_KEY");

  const options = {
    headers: {
      "X-API-KEY": key,
    },
  };

  const MANATAL_KEY = PropertiesService.getScriptProperties().getProperty("MANATAL_API_TOKEN");
  const optionsManatal = {
    method: 'get',
    headers: {
      'Authorization': 'Token ' + MANATAL_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true,
  };


  const urlLevels = "https://app.peopleforce.io/api/public/v3/job_levels"

  const jobLevels = JSON.parse(UrlFetchApp.fetch(urlLevels, options).getContentText()).data?.map(({ id, name }, i) => [name, id]);

  const allPositions = [];
  let nextPage = 1;
  let hasNext = true;

  while (hasNext) {
    const url = `https://app.peopleforce.io/api/public/v3/positions?page=${nextPage}`;
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    allPositions.push(...data.data.map(({ id, name }) => [name, id]));

    const totalPages = data.metadata?.pagination?.pages || nextPage;
    nextPage++;
    hasNext = nextPage <= totalPages; // Fixed logic
  }

  const allLocations = []
  let nextPageLoc = 1;
  let hasNextLoc = true;

  while (hasNextLoc) {
    const urlLoc = `https://app.peopleforce.io/api/public/v3/locations?page=${nextPageLoc}`;
    const responseLoc = UrlFetchApp.fetch(urlLoc, options);
    const dataLoc = JSON.parse(responseLoc.getContentText());

    allLocations.push(...dataLoc.data.map(({ id, name }) => [name, id]));

    const totalPagesLoc = dataLoc.metadata?.pagination?.pages || nextPageLoc;
    nextPageLoc++;
    hasNextLoc = nextPageLoc <= totalPagesLoc; // Fixed logic
  }

  const allOrganizations = []
  let nextPageOrg = 1;
  let hasNextOrg = true;

  while (hasNextOrg) {
    const urlORG = `https://api.manatal.com/open/v3/organizations?page=${nextPageOrg}`;
    const responseOrg = UrlFetchApp.fetch(urlORG, optionsManatal);
    const dataOrg = JSON.parse(responseOrg.getContentText());

    allOrganizations.push(...dataOrg?.results?.map(({ id, name }) => [id, name]));

    nextPageOrg++;
    hasNextOrg = optionsManatal.next // Fixed logic
  }


  const allStages = []
  let nextPageSt = 1;
  let hasNextSt = true;

  while (hasNextSt) {
    const urlORG = `https://api.manatal.com/open/v3/job-pipelines?page=${nextPageSt}`;
    const responseOrg = UrlFetchApp.fetch(urlORG, optionsManatal);
    const dataOrg = JSON.parse(responseOrg.getContentText());

    dataOrg?.results?.map((pipeline) => pipeline.job_pipeline_stages.forEach(({ name, id }) => allStages.push(([pipeline.name, id, name]))));

    nextPageSt++;
    hasNextSt = optionsManatal.next // Fixed logic
  }


  let invoicngFields = []
  let nextPageTeam = 1;
  let hasNextTeam = true;


  while (hasNextTeam) {
    const urlTeam = `https://app.peopleforce.io/api/public/v3/employee_fields?page=${nextPageTeam}`;
    const responseTeam = UrlFetchApp.fetch(urlTeam, options);
    const dataTeam = JSON.parse(responseTeam.getContentText());

    invoicngFields.push(...dataTeam.data);

    const totalPagesTeam = dataTeam.metadata?.pagination?.pages || nextPageTeam;
    nextPageTeam++;
    hasNextTeam = nextPageTeam <= totalPagesTeam; // Fixed logic
  }

  const invoicngTeam = invoicngFields.filter(f => f.name == "Invoicing team (dropdown)")?.[0]
  const teamOptions = invoicngTeam.options.map(({ value, id }) => [value, id])

  const clientID = invoicngFields.filter(f => f.name == "Client ID_new")?.[0]
  const clientOptions = clientID.options.map(({ value, id }) => [value, id])


  positionsHandbook.getRange("A2:U").clearContent()


  positionsHandbook.getRange(2, 13, teamOptions.length, teamOptions[0].length).setValues(teamOptions)
  positionsHandbook.getRange(2, 1, allPositions.length, allPositions[0].length).setValues(allPositions)
  positionsHandbook.getRange(2, 4, jobLevels.length, jobLevels[0].length).setValues(jobLevels)
  positionsHandbook.getRange(2, 7, allLocations.length, allLocations[0].length).setValues(allLocations)
  positionsHandbook.getRange(2, 10, allOrganizations.length, allOrganizations[0].length).setValues(allOrganizations)
  positionsHandbook.getRange(2, 16, clientOptions.length, clientOptions[0].length).setValues(clientOptions)
  positionsHandbook.getRange(2, 19, allStages.length, allStages[0].length).setValues(allStages)


}


// {
//   "effective_from": "2025-10-19",
//   "probation_policy_id": PEOPLEFORCE_IDS.probationPolicyId,
//   "employment_type_id": PEOPLEFORCE_IDS.employmentTypeId,
//   "working_pattern_id": PEOPLEFORCE_IDS.workingPatternId
// }



// function sendEmail_(body, subject, name, htmlBody = false) {

//   const emailTo = SpreadsheetApp.getActiveSpreadsheet().getRange("Settings!B1").getValue()


//   if (emailTo) {
//     GmailApp.sendEmail(emailTo, subject, htmlBody ? "" : body,
//       {
//         name: name,
//         && htmlBody ? ...    {
//       htmlBody: emailBody
//     } : ""
//       })

// }


// }

function sendEmail_(body, subject, name, htmlBody = false) {
  const emailTo = SpreadsheetApp.getActiveSpreadsheet()
    .getRange("Settings!B1")
    .getValue();

  if (!emailTo) return;

  const options = {
    name,
    ...(htmlBody && { htmlBody: body }),
  };

  GmailApp.sendEmail(
    emailTo,
    subject,
    htmlBody ? "" : body,
    options
  );
}


function isMoreThanThreeMonths_(date1, date2) {
  const d1 = new Date(date1);
  const d2 = new Date(date2);

  // Find the difference in months
  const monthsDiff =
    (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());

  // If the difference is more than 3 months, return true right away
  if (monthsDiff > 3) return true;

  // If exactly 3 months, check the days
  if (monthsDiff === 3 && d2.getDate() > d1.getDate()) return true;

  return false;
}


function spanExtract_(str) {
  if (!str) return ""

  return str
    .replace(/<[^>]+>/g, '')   // remove tags
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")   // replace &nbsp;
    .trim();
}
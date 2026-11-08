/**
 * ClickUp custom field IDs of the HR list, keyed by purpose.
 * Script Property CLICKUP_FIELD_IDS (JSON), see config.example.json.
 */
const CLICKUP_FIELD_IDS = getJsonProperty_('CLICKUP_FIELD_IDS', {});

const BALANCE_FIELD_BY_REQUEST_TYPE =

{

  "Vacation": CLICKUP_FIELD_IDS.vacationbalance,
  "Overtime": CLICKUP_FIELD_IDS.dayoffbalance,
  "Sick leave": CLICKUP_FIELD_IDS.sickleavebalance,

}


/**
 * Option IDs of the ClickUp dropdown field request_type.
 * Script Property CLICKUP_REQUEST_TYPE_OPTIONS (JSON), see config.example.json.
 */
const REQUEST_TYPE_OPTION_IDS = getJsonProperty_('CLICKUP_REQUEST_TYPE_OPTIONS', {});


const TOTAL_SICK_DAYS = 7
const TOTAL_VACATION_DAYS = 24
const TOTAL_OVERTIME_DAYS_OFF = 4
const HR_LIST_ID = getProperty_('CLICKUP_HR_LIST_ID', '')
const CLICKUP_WORKSPACE_ID = getProperty_('CLICKUP_WORKSPACE_ID', '')
const VACATION_DAYS_PER_MONTH = 2


function findTaskByName_(list, exasctName) {


  const listTasks = fetchTasksFromClickUp_(list)
  const foundedTask = listTasks?.find(f => f.name == exasctName && !getCustomFieldByNameType(f.custom_fields, CLICKUP_FIELD_IDS.request_type)) || ""


  return foundedTask


}


function fetchTaskByIdClickUp_(taskId) {

  const TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN")
  const url = `https://api.clickup.com/api/v2/task/${taskId}`;
  const options = {
    method: 'GET',
    headers: {
      'Authorization': TOKEN
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON?.parse(response?.getContentText());

  return data || "";
}

function fetchTasksFromClickUp_(listId) {

  const TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN")
  let page = 0;
  let allTasks = [];
  let hasMore = true;

  while (hasMore) {
    const url = `https://api.clickup.com/api/v2/list/${listId}/task?page=${page}&subtasks=true&include_closed=true`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': TOKEN
      }
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    allTasks = allTasks.concat(data.tasks);
    page += 1;

    // Check if the returned task count is less than 100 to exit the loop
    hasMore = !data.last_page;
  }
  return allTasks;
}

function dateParser(date) {

  if (!date) return ''

  return new Date(parseInt(date)) || null

}

function dropDownParser(dropDownFields, fullOption) {

  const keyValue = dropDownFields?.value

  const result = dropDownFields?.type_config?.options?.filter(f => f.orderindex == keyValue)?.[0]
  // ?.name


  return fullOption
    ? result || ''
    : result?.name

}

function labelParser(labelField, fullOption) {

  const keyValue = labelField?.value?.[0]
  const result = labelField?.type_config?.options?.filter(f => f.id == keyValue)?.[0]


  return fullOption
    ? result || ''
    : result?.label


}

function getCustomFieldByNameType(fields, fieldId, fullOption = false) {

  const selectedField = fields?.filter(({ id }) => id == fieldId)[0] || ''
  const type = selectedField?.type

  if (type == 'drop_down') {
    return dropDownParser(selectedField, fullOption)
  }
  else if (type == 'labels') {
    return labelParser(selectedField, fullOption)
  }

  else if (type == 'list_relationship') {
    return selectedField?.value?.[0]?.name || ''
  }

  else if (type == 'date') {
    return dateParser(selectedField?.value)
  }

  else if (type == 'users') {

    return fullOption
      ? selectedField?.value
      : selectedField?.value?.[0]?.username || ''

  }

  else {
    return selectedField?.value || ''
  }

}


function getTimeEntriesClickUp_(teamID, assignees, startDate, endDate) {

  const CLICKUP_API_TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN")

  if ((startDate - endDate) < 0) {

    const timeUrlClickup = `https://api.clickup.com/api/v2/team/${teamID}/time_entries?assignee=${assignees}&start_date=${startDate}&end_date=${endDate}&include_task_tags=true`

    const options = {
      method: "get",
      headers: {
        'Authorization': CLICKUP_API_TOKEN,
        'content-type': 'application/json',
        muteHttpExceptions: true,
      },
    }

    const response = UrlFetchApp.fetch(timeUrlClickup, options);
    const data = JSON?.parse(response.getContentText()) || "";

    return data


  }

  else {

    return []

  }


}

function getPrivateChatId(userID) {
  const CLICKUP_API_TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN");
  const scriptProps = PropertiesService.getScriptProperties();
  const cacheKey = `dm_chat_${userID}`;

  // ---------------------------
  // 1. Check cache
  // ---------------------------
  const cachedId = scriptProps.getProperty(cacheKey);
  if (cachedId) return cachedId;

  // ---------------------------
  // 2. GET all DM channels and search for existing
  // ---------------------------
  const listUrl = `https://api.clickup.com/api/v3/workspaces/${CLICKUP_WORKSPACE_ID}/chat/channels?type=DM`;

  const listOptions = {
    method: "get",
    headers: {
      "Authorization": CLICKUP_API_TOKEN
    },
    muteHttpExceptions: true
  };

  const listResponse = UrlFetchApp.fetch(listUrl, listOptions);
  const listData = JSON.parse(listResponse.getContentText());

  if (listData?.data) {
    for (const channel of listData.data) {
      if (channel.type === "DM" && channel.members?.some(m => m.id == userID)) {
        scriptProps.setProperty(cacheKey, channel.id); // save
        return channel.id;
      }
    }
  }

  // ---------------------------
  // 3. No existing DM → create new one
  // ---------------------------
  const createUrl = `https://api.clickup.com/api/v3/workspaces/${CLICKUP_WORKSPACE_ID}/chat/channels/direct_message`;

  const createOptions = {
    method: "post",
    headers: {
      "Authorization": CLICKUP_API_TOKEN,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      "user_ids": [userID]
    }),
    muteHttpExceptions: true
  };

  const createResponse = UrlFetchApp.fetch(createUrl, createOptions);
  const createData = JSON.parse(createResponse.getContentText());

  const newId = createData?.data?.id;

  if (newId) {
    scriptProps.setProperty(cacheKey, newId); // cache new DM
    return newId;
  }

  // ---------------------------
  // 4. Ultimate fallback
  // ---------------------------
  return "";
}

// function getPrivateChatId(userID) {

//   const CLICKUP_API_TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN")
//   const URL = `https://api.clickup.com/api/v3/workspaces/${CLICKUP_WORKSPACE_ID}/chat/channels/direct_message`

//   const options = {
//     method: "post",
//     headers: {
//       'Authorization': CLICKUP_API_TOKEN,
//       'content-type': 'application/json',
//       muteHttpExceptions: true,
//       payload: JSON.stringify({
//         "user_ids": [
//           userID
//         ]
//       }),

//     }

//   }

//   const response = UrlFetchApp.fetch(URL, options);
//   const data = JSON?.parse(response.getContentText()) || "";


//   return data?.data?.id || ""

// }

function sendClickUpMessage(userID, text) {
  const CLICKUP_API_TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN");
  const chatID = getPrivateChatId(userID);

  if (!chatID) {
    Log.error("Cannot get chat ID for user:", userID);
    return;
  }

  // Ensure text is string and max 40000 chars
  text = String(text).substring(0, 40000);

  const URL = `https://api.clickup.com/api/v3/workspaces/${CLICKUP_WORKSPACE_ID}/chat/channels/${chatID}/messages`;

  const payload = {
    "type": "message",
    "content_format": "text/plain",
    "assignee": userID,
    "content": text
  };

  const options = {
    method: "post",
    headers: {
      'Authorization': CLICKUP_API_TOKEN,
      'Content-Type': 'application/json',
      muteHttpExceptions: true,
    },
    payload: JSON.stringify(payload),
  };

  const response = UrlFetchApp.fetch(URL, options);
  const data = JSON.parse(response.getContentText());
  Log.info(data);
}


// async function clickUpRequest(url, method, payload = "", mute = true) {

//   const CLICKUP_API_TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN")


//   const options = {
//     method: method,
//     headers: {
//       'Authorization': CLICKUP_API_TOKEN,
//       'content-type': 'application/json',
//       ...(mute && {
//         muteHttpExceptions: true,
//       })
//     },
//     ...(payload && { payload: JSON.stringify(payload) }),

//   }


//   const response = await UrlFetchApp.fetch(url, options);
//   const data = JSON?.parse(response.getContentText()) || "";

//   return data


// }

function getUserIdsFromField_(fields, fieldId) {
  const field = (fields || []).find(({ id }) => id == fieldId);
  if (!field || !Array.isArray(field.value)) {
    Log.warn(`Field ${fieldId}: missing or empty`, JSON.stringify(field || null));
    return [];
  }

  const ids = [];
  field.value.forEach(entry => {
    const raw = (entry && typeof entry === 'object') ? entry.id : entry;
    const num = Number(raw);
    if (Number.isFinite(num)) {
      ids.push(num);
    } else {
      Log.warn(`Field ${fieldId}: non-user entry skipped (group?)`, JSON.stringify(entry));
    }
  });

  return ids;
}

function clickUpRequest(url, method, payload = "", mute = true) {
  const TOKEN = PropertiesService.getScriptProperties().getProperty("CLICKUP_API_TOKEN");

  const options = {
    method: method,
    headers: { 'Authorization': TOKEN },
    contentType: 'application/json',
    muteHttpExceptions: mute,
    ...(payload && { payload: JSON.stringify(payload) })
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code >= 300) {
    Log.error(`ClickUp ${method.toUpperCase()} ${url} → ${code}: ${text}`);
    throw new Error(`ClickUp ${code}: ${text}`);
  }

  return text ? JSON.parse(text) : "";
}


function getWorkingDaysWithHolidays_(startTimestamp, endTimestamp, holidayList) {
  // Convert to Date objects
  var startDate = new Date(startTimestamp);
  var endDate = new Date(endTimestamp);

  // Ensure startDate <= endDate
  if (startDate > endDate) {
    var temp = startDate;
    startDate = endDate;
    endDate = temp;
  }

  // Normalize times
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);

  // Convert holiday list to normalized dates
  var holidays = holidayList.map(function (h) {
    var d = new Date(h);
    d.setHours(0, 0, 0, 0);
    return d.getTime();
  });

  var workingDays = 0;
  var current = new Date(startDate);

  while (current <= endDate) {
    var day = current.getDay(); // 0=Sun, 6=Sat
    var currentTime = current.getTime();

    var isWeekend = (day === 0 || day === 6);
    var isHoliday = holidays.indexOf(currentTime) !== -1;

    if (!isWeekend && !isHoliday) {
      workingDays++;
    }

    current.setDate(current.getDate() + 1);
  }

  return workingDays;
}

function daysUntilEndOfYear_(timestamp) {
  const date = new Date(timestamp);

  // Last day of the year at midnight
  const endOfYear = new Date(date.getFullYear(), 11, 31); // December = 11
  endOfYear.setHours(23, 59, 59, 999);

  const diffMs = endOfYear - date;   // difference in milliseconds
  const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));

  return diffDays;
}


const getMonthStartEndUTC_ = (date) => {
  const d = new Date(date);
  return {
    start: Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1),
    end: Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0, 23, 59, 59, 999)
  };
};


function logToSheet(message, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  sheet.appendRow([
    new Date(),
    message,
    typeof data === "object" ? JSON.stringify(data) : data
  ]);
}

function doPost(e) {
  Log.info('doPost: request received', { hasBody: !!(e && e.postData), params: (e && e.parameter) ? Object.keys(e.parameter).filter(function (k) { return k !== 'token'; }) : [] });
  try {
    const data = JSON.parse(e.postData.contents);
    const event = data.event;
    const taskId = data.task_id;

    if (!event || !taskId) return ok();

    // ===============================
    // TASK CREATED
    // ===============================
    else if (event == "taskCreated") {

      const taskInfo = fetchTaskByIdClickUp_(taskId)
      const request_type = getCustomFieldByNameType(taskInfo?.custom_fields, CLICKUP_FIELD_IDS.request_type, true)

      if (request_type?.name) {

        const name = getCustomFieldByNameType(taskInfo?.custom_fields, CLICKUP_FIELD_IDS.created_user)
        const findParentTask = findTaskByName_(HR_LIST_ID, name)

        /// Create subtasks \\\

        // if (findParentTask?.id) {

        //   const manager = fetchTaskByIdClickUp_(findParentTask?.id)


        //   const payload =
        //   {
        //     "name": taskInfo?.name,
        //     "parent": findParentTask?.id,
        //     "custom_item_id": 1017,
        //     "due_date": taskInfo.due_date,
        //     "start_date": taskInfo.start_date,
        //     "status": "off | for review",
        //     "assignees": [
        //       ...getCustomFieldByNameType(manager?.custom_fields, CLICKUP_FIELD_IDS.manager, true)?.map(m => m.id)

        //     ],

        //     "custom_fields": [
        //       {
        //         "value": request_type?.id
        //       },
        //       {
        //         "value": getWorkingDaysWithHolidays_(Number(taskInfo.start_date), Number(taskInfo.due_date), []) /// day difference
        //       },

        //       {
        //         "value": getCustomFieldByNameType(taskInfo?.custom_fields, CLICKUP_FIELD_IDS.comment)

        //       }


        //       // {
        //       //   "id": "fdhfdhfc5c1767a-3456-42cb-acae-d8fbfffba65cdh",
        //       //   "value": {
        //       //     "add": [
        //       //       Number(name?.id)
        //       //     ]
        //       //   },

        //       // }
        //     ],


        //   }

        //   clickUpRequest(`https://api.clickup.com/api/v2/list/${HR_LIST_ID}/task`, "post", payload)


        //   /// Delete essential task \\\

        //   clickUpRequest(`https://api.clickup.com/api/v2/task/${taskId}`, "delete", {})

        // }

        if (findParentTask?.id) {
          const manager = fetchTaskByIdClickUp_(findParentTask.id);
          const assignees = getUserIdsFromField_(manager?.custom_fields, CLICKUP_FIELD_IDS.manager);

          if (assignees.length === 0) {
            Log.error(`No manager resolved on parent task ${findParentTask.id} — aborting, source task kept`);
            return ok();
          }

          const payload = {
            name: taskInfo?.name,
            parent: findParentTask.id,
            custom_item_id: 1017,
            due_date: taskInfo.due_date,
            start_date: taskInfo.start_date,
            status: "off | for review",
            assignees: assignees,
            custom_fields: [
              { id: CLICKUP_FIELD_IDS.request_type, value: request_type?.id },
              { id: CLICKUP_FIELD_IDS.off_days, value: getWorkingDaysWithHolidays_(Number(taskInfo.start_date), Number(taskInfo.due_date), []) },
              { id: CLICKUP_FIELD_IDS.comment, value: getCustomFieldByNameType(taskInfo?.custom_fields, CLICKUP_FIELD_IDS.comment) }
            ]
          };

          const created = clickUpRequest(
            `https://api.clickup.com/api/v2/list/${HR_LIST_ID}/task`, "post", payload
          );

          Log.info("created task", created?.id, "assignees:", JSON.stringify(created?.assignees));

          // Only destroy the source once the replacement demonstrably exists
          if (created?.id) {
            clickUpRequest(`https://api.clickup.com/api/v2/task/${taskId}`, "delete", {});
          }
        }


      }

      return ok();
    }

    // ===============================
    // STATUS UPDATED
    // ===============================
    else if (event === "taskStatusUpdated" && isAlreadyProcessed(data.history_items[0]?.id) == false) {
      if (data.history_items[0]?.after?.status === "off | approved") {
        approveProcessing(taskId);
      }

      else if (data.history_items[0]?.after?.status === "off| rejected") {
        rejectedProcessing(taskId);
      }


      else if (data.history_items[0]?.after?.status === "new hire") {
        sickLeaveNew(taskId);
      }


    }

    return ok();

  } catch (err) {
    Log.error(err);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// function logToSheet(message, data) {
//   sheet.appendRow([
//     new Date(),
//     message,
//     typeof data === "object" ? JSON.stringify(data) : data
//   ]);
// }


function isAlreadyProcessed(eventId) {
  const cache = CacheService.getScriptCache();

  const key = `clickup_event_${eventId}`;
  const exists = cache?.get(key);

  if (exists) return true;

  // Store for 10 minutes (600 sec) – you can change to 300 if you want
  cache.put(key, "1", 600);

  return false;
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}


function approveProcessing(taskID) {

  const getTask = fetchTaskByIdClickUp_(taskID)

  const getParent = fetchTaskByIdClickUp_(getTask.parent)


  const getDays = getCustomFieldByNameType(getTask?.custom_fields, CLICKUP_FIELD_IDS.off_days)
  const requestType = getCustomFieldByNameType(getTask?.custom_fields, CLICKUP_FIELD_IDS.request_type)
  const fromDate = Utilities.formatDate(new Date(Number(getTask?.start_date)), "UTC", "dd-MMM-yy")
  const toDate = Utilities.formatDate(new Date(Number(getTask?.due_date)), "UTC", "dd-MMM-yy")

  const requestedField = BALANCE_FIELD_BY_REQUEST_TYPE[requestType]

  clickUpRequest(`https://api.clickup.com/api/v2/task/${taskID}`, "put", { "status": "complete" })
  clickUpRequest(`https://api.clickup.com/api/v2/task/${taskID}/field/${requestedField}`, "post", { "value": true })


  const oldBalance = getCustomFieldByNameType(getParent?.custom_fields, requestedField)
  const getUserID = getCustomFieldByNameType(getParent?.custom_fields, CLICKUP_FIELD_IDS.created_user, true)?.[0]?.id
  const manager = getCustomFieldByNameType(getParent?.custom_fields, CLICKUP_FIELD_IDS.manager)
  const newValue = oldBalance - getDays

  clickUpRequest(`https://api.clickup.com/api/v2/task/${getTask.parent}/field/${requestedField}`, "post", { "value": newValue })


  /// Add message for Approve

  const textToSend = `Hi, your ${requestType} for ${fromDate} - ${toDate} (${getDays} days) has been approved. Please contact your manager ${manager} for details`


  sendClickUpMessage(getUserID, textToSend)

}

function rejectedProcessing(taskID) {

  const getTask = fetchTaskByIdClickUp_(taskID)
  const getParent = fetchTaskByIdClickUp_(getTask.parent)

  const manager = getCustomFieldByNameType(getParent?.custom_fields, CLICKUP_FIELD_IDS.manager)


  const getDays = getCustomFieldByNameType(getTask?.custom_fields, CLICKUP_FIELD_IDS.off_days)
  const requestType = getCustomFieldByNameType(getTask?.custom_fields, CLICKUP_FIELD_IDS.request_type)
  const fromDate = Utilities.formatDate(new Date(Number(getTask?.start_date)), "UTC", "dd-MMM-yy")
  const toDate = Utilities.formatDate(new Date(Number(getTask?.due_date)), "UTC", "dd-MMM-yy")

  const getUserID = getCustomFieldByNameType(getParent?.custom_fields, CLICKUP_FIELD_IDS.created_user, true)?.id


  /// Add message for Rejection

  const text = `Hi, your ${requestType} for ${fromDate} - ${toDate} (${getDays} days) has been rejected. Please contact your manager ${manager} for details`
  sendClickUpMessage(getUserID, text)


}


function calculateBalances() {

  const taskTypes = [1017, 1018]
  const allTasks = fetchTasksFromClickUp_(HR_LIST_ID)

  const allTasksObject = allTasks.reduce((accum, value) => {
    accum[value.id] = value

    return accum

  }, {})

  let tasks = allTasks.filter(({ parent }) => !parent)
  // let tasks_ids = tasks.map(({ id }) => { return id })


  /// Subtasks 
  let subtasks = allTasks.filter(({ parent }) => parent)


  // Dependancy processing
  let arrayWithDependencies = tasks.reduce((result, task) => {

    let subtasksOfTask = subtasks.filter((subtask) => subtask.parent == task.id && taskTypes.includes(subtask.custom_item_id) && !getCustomFieldByNameType(taskInfo?.custom_fields, CLICKUP_FIELD_IDS.inbalance) && subtask.status == "off | approved")

    if (subtasksOfTask.length) {

      result[task] = [...subtasksOfTask];

    }

    return result


  }, {})


  const keys = Object.keys(arrayWithDependencies)

  keys.forEach(key => {

    const employer = allTasksObject?.[key]

    let vacationBalance = getCustomFieldByNameType(employer?.custom_fields, CLICKUP_FIELD_IDS.vacationbalance)
    let sickleaveBalance = getCustomFieldByNameType(employer?.custom_fields, CLICKUP_FIELD_IDS.sickleavebalance)
    let dayoffBalance = getCustomFieldByNameType(employer?.custom_fields, CLICKUP_FIELD_IDS.dayoffbalance)

    arrayWithDependencies?.[key]?.forEach(subtask => {

      const subtaskDetails = allTasksObject?.[subtask]
      const taskType = subtaskDetails.custom_item_id
      const type = getCustomFieldByNameType(subtaskDetails?.custom_fields, CLICKUP_FIELD_IDS.request_type)

      const offdays = taskType == 1017
        ? -getCustomFieldByNameType(subtaskDetails?.custom_fields, CLICKUP_FIELD_IDS.off_days)
        : getCustomFieldByNameType(subtaskDetails?.custom_fields, CLICKUP_FIELD_IDS.off_days)


      switch (type) {
        case "Vacation":
          vacationBalance = vacationBalance + offdays > TOTAL_VACATION_DAYS
            ? TOTAL_VACATION_DAYS
            : vacationBalance + offdays
          break;

        case "Day-off":
          dayoffBalance += offdays
          break;

        case "Sick leave":
          sickleaveBalance += offdays

        default:
          ""
      }

      clickUpRequest(`https://api.clickup.com/api/v2/task/${subtask}/field/${CLICKUP_FIELD_IDS.inbalance}`, "post", { "value": true })


    })

    if (vacationBalance != getCustomFieldByNameType(employer?.custom_fields, CLICKUP_FIELD_IDS.vacationbalance)) {

      clickUpRequest(`https://api.clickup.com/api/v2/task/${key}/field/${CLICKUP_FIELD_IDS.vacationbalance}`, "post", { "value": vacationBalance })

    }


    if (dayoffBalance != getCustomFieldByNameType(employer?.custom_fields, CLICKUP_FIELD_IDS.dayoffbalance)) {

      clickUpRequest(`https://api.clickup.com/api/v2/task/${key}/field/${CLICKUP_FIELD_IDS.dayoffbalance}`, "post", { "value": dayoffBalance })

    }


    if (sickleaveBalance != getCustomFieldByNameType(employer?.custom_fields, CLICKUP_FIELD_IDS.sickleavebalance)) {

      clickUpRequest(`https://api.clickup.com/api/v2/task/${key}/field/${CLICKUP_FIELD_IDS.sickleavebalance}`, "post", { "value": sickleaveBalance })


    }


  })


}

function manualWrapper() {

  sickLeaveNew("869bx8q01")

}


function sickLeaveNew(taskId) {


  const taskInfo = fetchTaskByIdClickUp_(taskId)
  const you = getCustomFieldByNameType(taskInfo?.custom_fields, CLICKUP_FIELD_IDS.created_user)
  const joiningDate = getCustomFieldByNameType(taskInfo?.custom_fields, CLICKUP_FIELD_IDS.joining_date)


  if (taskInfo.custom_item_id == 1015 && !taskInfo.parent && joiningDate && you) {
    const tillEndYear = daysUntilEndOfYear_(joiningDate)
    const prorated = Math.min(tillEndYear / 365, 1)
    const calculateBalance = Math?.floor(prorated * TOTAL_SICK_DAYS)

    if (calculateBalance) {
      clickUpRequest(`https://api.clickup.com/api/v2/task/${taskId}/field/${CLICKUP_FIELD_IDS.sickleavebalance}`, "post", { "value": calculateBalance })

    }

  }


}

///// Sick Leave Accrual - every 1 january \\\\

function sickLeaveYear() {

  const today = new Date()

  if (today.getDate() == 1 && today.getMonth() == 0) {

    const allTasksToChange = fetchTasksFromClickUp_(HR_LIST_ID)?.filter(f =>
      !f.parent && !f.archived &&
      !["complete", "former employee"].includes(f.status?.status) &&
      f.custom_item_id == 1015 &&
      getCustomFieldByNameType(f?.custom_fields, CLICKUP_FIELD_IDS.created_user)
    )
    allTasksToChange.forEach(person => {

      clickUpRequest(`https://api.clickup.com/api/v2/task/${person.id}/field/${CLICKUP_FIELD_IDS.sickleavebalance}`, "post", { "value": TOTAL_SICK_DAYS })

      /// Transfer balance

      const vacationBalance =
        getCustomFieldByNameType(person?.custom_fields, CLICKUP_FIELD_IDS.vacationbalance) || 0

      clickUpRequest(`https://api.clickup.com/api/v2/task/${person.id}/field/${CLICKUP_FIELD_IDS.vacationbeginbalance}`, "post", { "value": vacationBalance })


    })


  }


}


///// Vacation Accrual - every 1 st day of month \\\\

function vacationDayoffAccrual() {

  const today = new Date()
  // const today = new Date(2026, 3, 1)
  const calcDate = new Date(today)
  calcDate.setDate(calcDate.getDate() - 1)


  const { start, end } = getMonthStartEndUTC_(calcDate)

  if (today.getDate() == 1) {

    const allTasksToChange = fetchTasksFromClickUp_(HR_LIST_ID)?.filter(f =>
      !f.parent && !f.archived && // Except subtasks
      !["complete", "former employee"].includes(f.status?.status) && // Except completed
      f.custom_item_id == 1015 && // Taks type Person
      getCustomFieldByNameType(f?.custom_fields, CLICKUP_FIELD_IDS.created_user) // You is presented
    )

    allTasksToChange.forEach(person => {
      const userID = getCustomFieldByNameType(person?.custom_fields, CLICKUP_FIELD_IDS.created_user, true)?.[0]

      const vacationBalance =
        getCustomFieldByNameType(person?.custom_fields, CLICKUP_FIELD_IDS.vacationbalance) || 0

      const dayoffBalance =
        getCustomFieldByNameType(person?.custom_fields, CLICKUP_FIELD_IDS.dayoffbalance) || 0

      const vacationBeginBalance =
        getCustomFieldByNameType(person?.custom_fields, CLICKUP_FIELD_IDS.vacationbeginbalance) || 0

      const probationEnd =
        getCustomFieldByNameType(person?.custom_fields, CLICKUP_FIELD_IDS.probationend)

      // ✅ FIRST ACCRUAL MONTH = probation ended in accrual month
      const vacationNewCheck =
        probationEnd &&
        vacationBalance === 0 &&
        new Date(probationEnd).getMonth() === calcDate.getMonth() &&
        new Date(probationEnd).getFullYear() === calcDate.getFullYear()

      const startToUse = vacationNewCheck
        ? Date.UTC(
          new Date(probationEnd).getUTCFullYear(),
          new Date(probationEnd).getUTCMonth(),
          new Date(probationEnd).getUTCDate() + 1
        )
        : (end - new Date(probationEnd).getTime() > 0 || !probationEnd)
          ? start
          : 0

      /* ---------------- VACATION ---------------- */

      if (startToUse != 0) {

        const vacEntries = getTimeEntriesClickUp_(CLICKUP_WORKSPACE_ID, userID.id, startToUse, end)

        const vacHours =
          vacEntries?.data.reduce((s, e) => s + Number(e.duration), 0) / 3600000 || 0

        const calculateVacation =
          vacHours >= 160
            ? VACATION_DAYS_PER_MONTH
            : Math.floor((vacHours / 160) * VACATION_DAYS_PER_MONTH)

        const canAccrual = TOTAL_VACATION_DAYS - vacationBeginBalance

        const addDaysOfVacation = Math.max(
          0,
          Math.min(calculateVacation, canAccrual)
        )

        if (addDaysOfVacation > 0) {
          clickUpRequest(
            `https://api.clickup.com/api/v2/list/${HR_LIST_ID}/task`,
            "post",
            {
              name: "Vacation Accrual",
              parent: person.id,
              custom_item_id: 1018,
              status: "complete",
              custom_fields: [
                { id: CLICKUP_FIELD_IDS.request_type, value: REQUEST_TYPE_OPTION_IDS.Vacation },
                { id: CLICKUP_FIELD_IDS.off_days, value: addDaysOfVacation },
                { id: CLICKUP_FIELD_IDS.inbalance, value: true }
              ]
            }
          )

          clickUpRequest(
            `https://api.clickup.com/api/v2/task/${person.id}/field/${CLICKUP_FIELD_IDS.vacationbalance}`,
            "post",
            { value: Number(vacationBalance) + Number(addDaysOfVacation) }
          )
        }

      }


      // /* ---------------- DAY OFF / OVERTIME ---------------- */

      const allEntries = getTimeEntriesClickUp_(CLICKUP_WORKSPACE_ID, userID.id, start, end)

      const totalHours =
        allEntries?.data.reduce((s, e) => s + Number(e.duration), 0) / 3600000 || 0

      const calculationDayOff =
        totalHours > 160 ? Math.floor((totalHours - 160) / 8) : 0

      const addDaysOfOffs = Math.max(
        0,
        Math.min(calculationDayOff, TOTAL_OVERTIME_DAYS_OFF - dayoffBalance)
      )

      if (addDaysOfOffs > 0) {
        clickUpRequest(
          `https://api.clickup.com/api/v2/list/${HR_LIST_ID}/task`,
          "post",
          {
            name: "Day-Off Accrual",
            parent: person.id,
            custom_item_id: 1018,
            status: "complete",
            custom_fields: [
              { id: CLICKUP_FIELD_IDS.request_type, value: REQUEST_TYPE_OPTION_IDS.Overtime },
              { id: CLICKUP_FIELD_IDS.off_days, value: addDaysOfOffs },
              { id: CLICKUP_FIELD_IDS.inbalance, value: true }
            ]
          }
        )

        clickUpRequest(
          `https://api.clickup.com/api/v2/task/${person.id}/field/${CLICKUP_FIELD_IDS.dayoffbalance}`,
          "post",
          { value: Number(dayoffBalance) + Number(addDaysOfOffs) }
        )
      }
    })
  }
}


function monthlyTrigger() {

  sickLeaveYear()
  Utilities.sleep(500)
  vacationDayoffAccrual()

}

// URL and API Key should be added in manually when attaching the code to the relevant Sheet.
let schedulerApiURL = "";
let schedulerApiKey = "";

let ui = SpreadsheetApp.getUi();


let SchedulerStatus = {
  ACTIVE: 0,
  PAUSED: 1,
  UNREACHABLE: 2
}


function setupSchedulerApiUi()
{
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Scheduler')
      .addItem('Pause', 'requestPause')
      .addItem('Unpause', 'requestUnpause')
      .addSeparator()
      .addItem('Force Update', 'requestForceUpdate')
      .addSeparator()
      .addItem('One-Off Email', 'showOneOffEmailDialog')
      .addSeparator()
      .addItem('Latest Logs', 'requestLatestLogs')
      .addItem('Check Status', 'displayStatus')
      .addToUi();
}


function requestPause()
{
  var request =
  {
    "method": "post",
    "headers": { "X-Api-Key": schedulerApiKey },
    "payload": { "Body": "" },
    'validateHttpsCertificates': false
  };

  try
  {
    var response = UrlFetchApp.fetch(schedulerApiURL + "/pause", request);
    ui.alert("Pause successful.")
  }
  catch (exception)
  {
    ui.alert("Failed to reach the server when trying to pause the scheduler.\n\n" + exception);
  }
}


function requestUnpause()
{
  var request =
  {
    "method": "post",
    "headers": { "X-Api-Key": schedulerApiKey },
    "payload": { "Body": "" },
    'validateHttpsCertificates': false
  };

  try
  {
    var response = UrlFetchApp.fetch(schedulerApiURL + "/unpause", request);
    ui.alert("Unpause successful.")
  }
  catch (exception)
  {
    ui.alert("Failed to reach the server when trying to unpause the scheduler.\n\n" + exception);
  }
}


function requestForceUpdate()
{
  var request =
  {
    "method": "post",
    "headers": { "X-Api-Key": schedulerApiKey },
    "payload": { "Body": "" },
    'validateHttpsCertificates': false
  };

  try
  {
    var response = UrlFetchApp.fetch(schedulerApiURL + "/forceupdate", request);
    ui.alert("Force update successful.")
  }
  catch (exception)
  {
    ui.alert("Failed to reach the server when trying to force update.\n\n" + exception);
  }
}


function checkStatus()
{
  var request =
  {
    "method": "post",
    "headers": { "X-Api-Key": schedulerApiKey },
    "payload": { "Body": "" },
    'validateHttpsCertificates': false
  };

  let status;

  try
  {
    let response = UrlFetchApp.fetch(schedulerApiURL + "/status", request);
    status = response == "1" ? SchedulerStatus.ACTIVE : SchedulerStatus.PAUSED;
  }
  catch (exception)
  {
    status = SchedulerStatus.UNREACHABLE;
  }

  return status;
}


function displayStatus()
{
  status = checkStatus();

  if (status == SchedulerStatus.ACTIVE)
  {
    ui.alert("The scheduler is currently running.");
  }
  else if (status == SchedulerStatus.PAUSED)
  {
    ui.alert("The scheduler is currently paused.");
  }
  else
  {
    ui.alert("The scheduler is currently unreachable. This most likely means the server is offline" +
             " or the IP Address/API Key has not been properly configured.");
  }
}

function showOneOffEmailDialog()
{
    var campaignResult = ui.prompt(
        "Getting ready to send a One-off email",
        "Specify the target Campaign:",
        ui.ButtonSet.OK_CANCEL);

    if (campaignResult.getSelectedButton() == ui.Button.OK)
    {
        var templateResult = ui.prompt(
            "Getting ready to send a One-off email",
            "Specify the Template ID from SendGrid:",
            ui.ButtonSet.OK_CANCEL);

        if (templateResult.getSelectedButton() == ui.Button.OK)
        {
            campaign = campaignResult.getResponseText();
            template = templateResult.getResponseText();

            requestOneOffEmail(campaign, template);
        }
    }
}

function requestOneOffEmail(campaign, template)
{
    let jsonBody = {
        "campaign": campaign,
        "template": template
    }

    var request =
    {
        "method": "post",
        "contentType" : "application/json",
        "headers": { "X-Api-Key": schedulerApiKey },
        "payload": JSON.stringify(jsonBody),
        'validateHttpsCertificates': false
    };

    try
    {
        var response = UrlFetchApp.fetch(schedulerApiURL + "/oneoff", request);
        ui.alert("Emails were sent out.\n\n" + response);
    }
    catch (exception)
    {
        ui.alert("Failed to send out one-off email.\n\n" + exception);
    }
}

function requestLatestLogs()
{
    var request =
    {
        "method": "post",
        "headers": { "X-Api-Key": schedulerApiKey },
        "payload": { "Body": "" },
        'validateHttpsCertificates': false
    };

    try
    {
        var response = UrlFetchApp.fetch(schedulerApiURL + "/fetchlogs", request);
        saveLogs(response);
        ui.alert("Logs saved to a file in your Google Drive (DripConfig-Logs.txt).")
    }
    catch (exception)
    {
        ui.alert("Failed to reach the server when trying to force update.\n\n" + exception);
    }
}

function saveLogs(text)
{
    let folder = DriveApp.getRootFolder();
    let file = folder.createFile("DripConfig-Logs.txt", text);
}

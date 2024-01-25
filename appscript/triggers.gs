// Error checking for the Templates sheet doesn't
// happen until after a user makes an edit.
function onOpen(e)
{
  checkContacts();
  checkTemplates();
  checkCampaigns();
  checkSchedule();
  checkSubscriptions();

  setupSchedulerApiUi();
  checkErrorsUI();
  addTriggerSearchMenu();
}

// Clicking add a new row to the Templates sheet does not seem to
// trigger onEdit() so error checking doesn't occur until after
// the user has entered something.
function onEdit(e)
{
  checkContacts();
  checkTemplates();
  checkCampaigns();
  checkSchedule();
  checkSubscriptions();
  updateMatch();
}
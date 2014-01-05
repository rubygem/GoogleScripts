var DATE_OF_BOUT = '<date_of_bout>';
var RETURN_GAME = '<return_game_required>';
var REQUIRED_TO_RETURN = '<required_to_return>';
var REQUIRED_WITHIN = '<required_within>';
var MIN_SKILLS = '<min_skills_year>';
var AWAY_TEAM = '<away_team>';
var HOME_TEAM = '<home_team>';
var totalSteps = 0;
var B_TEAM_MIN_SKILLS = '2010';
var A_TEAM_MIN_SKILLS = '2013';

function onOpen() {
  DocumentApp.getUi().createMenu('Update Contract')
      .addItem('Wizard', 'wizard')
      .addToUi();
}

function wizard(){
  totalSteps = 5;
  dates(1);
  isReturnBout(2);
  specifyTeam(3);
  specifyHomeTeam(4);
  specifyAwayTeam(5);
}

function dates(step){
  var text = "Date of bout (mm/dd/yyyy <- silly Americans, don't worry I will make sure it's right on the form):";
  var result = createOKDialog(step, text);
  var boutDate = DATE_OF_BOUT;
  
  if (clickedOK(result)) {
    boutDate = new Date(result.getResponseText());
  }
  
  updateDates(boutDate);
}

function specifyAwayTeam(step){
  var text = "Name of Away team:";
  var result = createOKDialog(step, text);
  var teamName = AWAY_TEAM;
  
  if (clickedOK(result)) {
    teamName = result.getResponseText();
  }
  
  update(teamName, AWAY_TEAM);
}

function specifyHomeTeam(step){
  var text = "Name of Home team:";
  var result = createOKDialog(step, text);
  var teamName = HOME_TEAM;
  
  if (clickedOK(result)) {
    teamName = result.getResponseText();
  }
  
  update(teamName, HOME_TEAM);
}

function isReturnBout(step) {
  var text = 'Is a return game required?';
  var result = createYesNoDialog(step, text);
  var returnGameRequired = 'No';
  var requiredToReturn = 'not required';
  var requiredWithin = '';

  if (clickedYes(result)) {
    returnGameRequired = 'Yes';
    requiredToReturn = 'required';
    requiredWithin = ' within 18 months of the bout date, on a date that is to yet to be agreed';
  }
  
  update(returnGameRequired, RETURN_GAME);
  update(requiredToReturn, REQUIRED_TO_RETURN);
  update(requiredWithin, REQUIRED_WITHIN);
}

function specifyTeam(step){
  var text = 'Is this an A team game (so playing to higher skills)?';
  var result = createYesNoDialog(step, text);
  var team = 'B';

  if (clickedYes(result)) {
    team = 'A';
  }

  var minSkillsYear = B_TEAM_MIN_SKILLS;
  if (team == "A"){
    minSkillsYear = A_TEAM_MIN_SKILLS;
  }

  update(minSkillsYear, MIN_SKILLS);
}

function createOKDialog(step, text){
  return DocumentApp.getUi().prompt(wizardTitle(step), text, DocumentApp.getUi().ButtonSet.OK_CANCEL);
}

function wizardTitle(step){
  return 'New contract wizard - Step ' + step + ' of ' + totalSteps;
}

function clickedOK(result){
  return result.getSelectedButton() == DocumentApp.getUi().Button.OK;
}

function clickedYes(result){
  return result == DocumentApp.getUi().Button.YES;
}

function createYesNoDialog(step, text){
  return DocumentApp.getUi().alert(wizardTitle(step), text, DocumentApp.getUi().ButtonSet.YES_NO);
}

function update(required, text){
   var bodyElement = DocumentApp.getActiveDocument().getBody();
   bodyElement.replaceText(text, required);
}

function updateDates(boutDate){
  var bodyElement = DocumentApp.getActiveDocument().getBody();
  bodyElement.replaceText(DATE_OF_BOUT, Utilities.formatDate(boutDate, "GMT", "dd/MM/yyyy"));
  
  bodyElement.replaceText('<4_weeks_prior>', weeksPrior(boutDate, 4));
  bodyElement.replaceText('<3_weeks_prior>', weeksPrior(boutDate, 3));
  bodyElement.replaceText('<2_weeks_prior>', weeksPrior(boutDate, 2));
  bodyElement.replaceText('<1_weeks_prior>', weeksPrior(boutDate, 1));
  bodyElement.replaceText('<4_weeks_after>', weeksAfter(boutDate, 4));
}

function weeksPrior(boutDate, weeks){
  var today = new Date(boutDate);

  var numberOfDays = 7*weeks;
  var date = new Date(today).setDate(today.getDate()-(numberOfDays));
  return Utilities.formatDate(new Date(date), "GMT", "dd/MM/yyyy");;
}

function weeksAfter(boutDate, weeks){
  var today = new Date(boutDate);

  var numberOfDays = 7*weeks;
  var date = new Date(today).setDate(today.getDate()+(numberOfDays));
  return Utilities.formatDate(new Date(date), "GMT", "dd/MM/yyyy");;
}

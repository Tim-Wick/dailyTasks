// Created by Tim Wick - tjwick77@gmail.com

/*
Known potential issues:
If there are no jobs due on a day, pushLive will enter an infinite loop as it searches endlessly for a job

*/

// Home of all scripts/functions for the Daily Tasks
// This script file is to set global variables relating to various spreadsheets to be used by other script files and create functions used by multiple scripts

// Main sheet
dailyTasksSheet = SpreadsheetApp.openById("1bkN0XQVSYENpmf6AZAy8l4ehwaFd0dVSL6kG13ctB68");
// Task list sheet
taskList = dailyTasksSheet.getSheetByName("Task List");
// Current Tasks
currentTasks = dailyTasksSheet.getSheetByName("Current Tasks");
//Resources
resources = dailyTasksSheet.getSheetByName("Resources");
// Archive
archive = dailyTasksSheet = SpreadsheetApp.openById("1DTY-niO7CcmQMgupFwSp9Q6OXvuW-Sh0-Lc64gbCmuU").getSheetByName("Archived Tasks");
// Admin/monitoring email to use for alerts, use a comma separated list within the quotes to add multiple emails
adminEmail = "tjwick77@gmail.com";


function doGet() { // Function to run on load of the file
   return HtmlService.createTemplateFromFile('index').evaluate().setTitle("Daily Tasks").setFaviconUrl("https://b.kisscc0.com/20180705/byq/kisscc0-wireless-router-cisco-systems-computer-icons-compu-router-5b3dbedad7b141.4725061115307732108835.png");  
   populateTable() // Runs populate table the initial time
}

function include(filename) { // Function to be able to include the CSS file
  return HtmlService.createHtmlOutputFromFile(filename) // Calls getContent on the passed file
      .getContent();
}

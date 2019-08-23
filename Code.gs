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


//This function pushes the next task that is run on this day to the spreadsheet when called
function pushLive(completedRow){ // Takes in the row that was completed to check if it was thelast row in the spreadsheet
  var curRunDay = new Date().getDay(); // Sets curRunDay as the day of the week from 0-6
  
  // Sets up current task list variables
  var curTaskLast = currentTasks.getLastRow(); // Sets curTaskLast as the last occupied row in the currentTasks sheet
  var curTaskArrTemp = currentTasks.getRange(2,2,curTaskLast-1,1).getValues(); // Makes a temp array of values. -1 to account for occupied first row
  var curTaskArr = []; // Sets curTaskArr as a blank array
  for(i in curTaskArrTemp){curTaskArr.push(curTaskArrTemp[i][0]);} // Populates curTaskArr with values from temp array. getValues returns an array of arrays that is difficult to work with, so pulling the values into an array makes it easier
  
  // Sets up all task list variables
  var allTaskLast = taskList.getLastRow(); // Sets allTaskLast as the last occupied row in the taskList sheet
  var allTaskArrTemp = taskList.getRange(2, 2, allTaskLast-1, 1).getValues(); // Makes a temp array of values. -1 to not get a blank cell due to occupied first row
  var allTaskArr = []; //Sets allTaskArr as a blank array
  for(i in allTaskArrTemp){allTaskArr.push(allTaskArrTemp[i][0]);} // Populates allTaskArr with values from temp array. getValues returns an array of arrays that is difficult to work with, so pulling the values into an array makes it easier
  var lastTaskPos = allTaskArr.length; // Grabs last task postion in array to later check if we need to loop around to the first task
  
  var curTaskPos = allTaskArr.indexOf(curTaskArr[curTaskLast-2]); // Gets the position in allTaskArr of the current task. -2 to and account for occupied first row and arrays starting at 0 where sheets start at 1
  var nextTaskPos = 0; // Initializes the nextTaskPos variable. 0 is arbitrary and can be any integer or float 
  curTaskPos == lastTaskPos ?  nextTaskPos = 0 : nextTaskPos = curTaskPos += 1; // Checks if the curTaskPos is equal to the lastTaskPos, if it is, it leaves the nextTaskPos as 0 to loop back around to the beginning of the task list, otherwise incremenets by 1

  while(true){ // Sets up a while loop so we will continue until there is a job due today
    if (taskList.getRange((nextTaskPos+2),4,1,1).getValue().indexOf(curRunDay) >= 0){ // Checks if job should be run today. Script pulls a string from the spreadsheet, GAS does not support .includes() so we check if indexOf is greater than or equal to 0. indexOf returns -1 if curRunDay is not in the values     
      currentTasks.getRange((curTaskLast+1),1,1,5).setValues(taskList.getRange((nextTaskPos+2),1,1,5).getValues()); // If the job is due today, paste the information into the currentTasks sheet
      break // Breaks out of the loop once one new task is entered
    }
    else{nextTaskPos == lastTaskPos ? nextTaskPos = 0 : nextTaskPos = curTaskPos += 1;} // Checks if the curTaskPos is equal to the lastTaskPos, if it is, it leaves the nextTaskPos as 0 to loop back around to the beginning of the task list, otherwise increments by 1 so the loop can check the next task
  }
  currentTasks.deleteRow(completedRow); // Deletes the row that was checked as passed from the completeTask function
}

// Function to record a task when completed on the archive sheet and delete the row

// Function takes in the row of the job to record

function completeTask(row) {
  var user = Session.getActiveUser(); // Gets active user to record who completed job
  var fullDate = new Date (); // Sets a date variable
  var date = Utilities.formatDate(fullDate, "PST", "MM/dd/yyyy"); // Pulls date from fullDate
  var time = Utilities.formatDate(fullDate, "PST", "HH:mm"); // Pulls hours/minutes from fullDate
  var task = currentTasks.getRange(row, 1, 1, 5).getDisplayValues(); // Gets values from the row that was checked
  var onTime = isOnTime(task[0][0], time); // Runs the isOnTime function passing the time string from the row and the time variable, gets back no/yes
  var archArr = [date, task[0][1], user, task[0][0], time, onTime]; // Creates array to add to archiving sheet
  archive.insertRows(2); // Inserts a row at the top of the archive sheet. This way it doesn't matter how large the sheet gets, the scripts will only care about the first row and it is also easier to see recent jobs instead of scrolling
  archive.getRange(2,1,1,6).setValues([archArr]); // Sets the values from the archArr in the newly created row
  pushLive(row); // Calls pushLive, passing the row that was completed
  }
}

// Function to check if a job is over an hour past due. If so, an email will be sent to the admin email from Code.gs

function pastDueCheck() {
  var curDate = new Date(); // Initializes a date
  var curTime = (curDate.getTime()/1000); // Gets current date/time as unix timestamp in seconds (divided by 1000 as unix is in milliseconds)
  
  // Sets up current task list variables
  var curTaskLast = currentTasks.getLastRow(); // Sets curTaskLast as the last occupied row in the currentTasks sheet
  var curTaskArrTemp = currentTasks.getRange(2,1,curTaskLast-1,1).getDisplayValues(); // Makes a temp array of values. -1 to account for occupied first row
  var curTaskArr = []; // Sets curTaskArr as a blank array
  for(i in curTaskArrTemp){curTaskArr.push(curTaskArrTemp[i][0]);} // Populates curTaskArr with values from temp array. getValues returns an array of arrays that is difficult to work with, so pulling the values into an array makes it easier

  var pastDue = 0 // Initializes past due jobs
  for(i in curTaskArr){ // Loops over curTasksArr
    var split = curTaskArr[i].split(" "); // splits the by the space, ex. "7:30 AM" becomes [7:30, AM]
    var splitTwo = split[0].split(":"); // splits the previous split by the colon, creating [7, 30]
    var jobDate = new Date(curDate.getFullYear(), curDate.getMonth(), curDate.getDate(), splitTwo[0], splitTwo[1]); // Sets up a JS date object using the current date and the time previously split out
    split[1] == "PM" ? jobDate.setHours(jobDate.getHours() + 12) : jobDate = jobDate // if the job time was PM, it adds 12 hours to the date object. This is a ternary operator, basically an if statements. The dueDate = dueDate was just to fill out the right side of the colon, otherwise it didn't work
    if(curDate.getHours() >= 12 && split[1] == "AM"){jobDate.setDate(jobDate.getDate() + 1)} // Checks if the current time is PM (>= 12 hours in 24 hour) and the time of the job is AM. If so, it increments the job date by 1 from the current date. This negates false positives if the script is checking jobs in the PM that are due the next day AM.
    var jobTime = (jobDate.getTime()/1000); // Gets the jobTime as a unix time stamp in seconds (divided by 1000 as unix is in milliseconds)
    (curTime - jobTime) > 3600 ? pastDue++ : pastDue = pastDue // Checks if the job is over one hour old (3600 seconds), if so, increments pastDue by one
  }
  
  if(pastDue > 0){sendEmail(pastDue)} // If pastDue is over 0, calls sendEmail with the number of past due jobs
  
}

// Function to send an email to the admin email. Takes in the number of jobs that are past due.
function sendEmail(number){
  var message = "There are currently " + number + " daily tasks over one hour past due."
  MailApp.sendEmail(adminEmail,"Past Due Daily Tasks",message);
}

// Function to check temp tasks and delete those past expiration. Needs to be run at midnight each night to properly compare times. The time to live needs to include the day added as time is calculated from midnight the day added.

function tempTaskCheck() {
  var curDate = new Date(); // Initializes date
  var tasksLast = taskList.getLastRow(); // Gets last row of tasks
  var tasks = taskList.getRange(2,5,(tasksLast-1),3).getValues(); // Gets the temp y/n, created, and time to live columns. -1 to account for first row of spreadsheet being populated
  
  for(var i = tasks.length - 1; i >= 0; i--){ //  Loops through tasks from the end to the beginning. This is so a row can be deleted without changing position of the rest of the rows compared to the array of tasks
    if(tasks[i][0] == "Y"){ // Checks if the task is temp
      // Checks if unix time stamp of right now (should be midnight or close to it if scheduled correctly) is greater than midnight the day the temp task was created, plus the number of days times 86,400, the number of seconds in a day.
      // Dividing by 1000 is used to get seconds as it is initially grabbed in milliseconds 
      // Since the created timestamp will be midnight the day it is created, the script adds a day to the time to live so the day created is technically day 1
      if((curDate.getTime()/1000) > ((tasks[i][1].getTime()/1000) + ((tasks[i][2]+1)*86400))){ 
        taskList.deleteRow(i+2); // Deletes the position in the array, plus to to account for first row of spreadsheet and arrays starting at 0
      }
    }
  } 
}


function getTableContents() { // Function to get information from Current Tasks spreadsheet
  var currentLast = currentTasks.getLastRow() // Gets the last row of the spreadsheet
  var tasks = currentTasks.getRange(2,1,currentLast-1,3).getDisplayValues(); // Gets values from the sheet
  
  for(i in tasks){ // Loops through the tasks
    var replaceArray = []; // Sets up replaceArray as a blank array
    var resourceArray = tasks[i][2].split(","); // Splits the string of values into an array of links so it can handle multiple links
    for(j in resourceArray){ // Loops over the array of links
       var resourceInfo = getResourceInfo(resourceArray[j]); // Calls getResouceLink for the resource number. Returns with a link from the resources tab
       Logger.log(resourceInfo);
       var returnLink = "<a href='" + resourceInfo[0][1] + "' target='_blank'><div class='tooltip'><i class='material-icons'>info</i><span class='tooltiptext'>" + resourceInfo[0][0] + "</span></div></a>"; // Sets up the HTML code to display an icon with the corresponding link
       replaceArray.push(returnLink); // Pushes the link to replaceArray
    }
    var replaceArray2 = replaceArray.join(" ") // Splits replaceArray into a string separated by spaces
    tasks[i][2] = replaceArray2 // Sets the i index of tasks to the string of html code
  }
  return JSON.stringify(tasks); // Returns a JSON string of the tasks
}


// Function to alert when a task is created. Also acts as a stop to make it more difficult to click multiple jobs at once.

function alertInfo(row) {
  var taskName = currentTasks.getRange(row, 2).getDisplayValue(); // Gets the task name from the row that was checked
  var taskString = "Task '" + taskName + "' marked as complete. \nPlease wait for table to reload before marking another task as complete." // Sets up a string to return for an alert
  return taskString // Returns the string
}


function getResourceInfo(resourceNumber) { // Function to take in a resource number and return the corresponding resource name and link
  resourceNumber++
  var returnInfo = resources.getRange(resourceNumber++,2,1,2).getDisplayValues(); // Gets the value from the corresponding row and columns 2&3 of the resources spreadsheet. +1 to account for occupied first row
  return returnInfo // Returns the link
}


// This function takes in a job due time "timeDue", a time in an AM/PM format, and a current time, "curTime", creates a date object from the AM/PM time, and returns if the job was done on time
function isOnTime(timeDue, curTime){
  var date = new Date(); // initializes a date variable
  var dueSplit = timeDue.split(" "); // splits the timeDue by the space, ex. "7:30:00 AM" becomes [7:30:00, AM]
  var dueSplitTwo = dueSplit[0].split(":") // splits the previous split by the colon, creating [7, 30, 00]
  var dueDate = new Date(date.getFullYear(), date.getMonth(), date.getDate(), dueSplitTwo[0], dueSplitTwo[1]) // Sets up a JS date object using the current date and the time previously split out
  dueSplit[1] == "PM" ? dueDate.setHours(dueDate.getHours() + 12) : dueDate = dueDate // if the job time was PM, it adds 12 hours to the date object. This is a ternary operator, basically an if statements. The dueDate = dueDate was just to fill out the right side of the colon, otherwise it didn't work
  var dueTime = Utilities.formatDate(dueDate, "PST", "HH:mm"); // Pulls out the hours/minutes from the time dueTime object
  var onTime = "" // initializes a variable to describe if job is on time or not
  curTime > dueTime ? onTime = "No" : onTime = "Yes"; // Checks if the current time is larger than the task due time. If so sets onTime to "No"
  return onTime;
}

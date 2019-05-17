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

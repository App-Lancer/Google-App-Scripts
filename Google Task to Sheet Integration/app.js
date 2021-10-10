let sheetClient = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1Xlo3Z_LO0GI5mgEC9NCDNVH1TbcK1sE7x3Xp-fr9kZ0/edit')

function createSheet(id, name){
  let idMetaData = sheetClient.createDeveloperMetadataFinder().withKey("taskListId").withValue(id).find()[0]
  let nameMetaData = sheetClient.createDeveloperMetadataFinder().withKey(name).find()[0]
  
  var sheetTitle = name
  
  if (nameMetaData === undefined){
    sheetClient.addDeveloperMetadata(sheetTitle, '0')
    nameMetaData = sheetClient.createDeveloperMetadataFinder().withKey(sheetTitle).find()[0]
  }

  if (idMetaData){
    let loc = idMetaData.getLocation()
    if (loc.getLocationType() == SpreadsheetApp.DeveloperMetadataLocationType.SHEET){
      let sheet = loc.getSheet()
      let sheetNameSplit = sheet.getName().split("#")
      sheetNameSplit.pop()
      let sheetName = sheetNameSplit.join("#")
      if (!(sheetName === name)){
        try{
          return sheet.setName(sheetTitle)
        }
        catch(e){
          let count = parseInt(nameMetaData.getValue())
          nameMetaData.setValue(`${count+1}`)
          return sheet.setName(sheetTitle + `#${count+1}`)
        }
      }
      else{
        return sheet
      }
    }
  }
  else{
    try{
      return sheetClient.insertSheet(sheetTitle).addDeveloperMetadata('taskListId', id)
    }
    catch(e){
      let count = parseInt(nameMetaData.getValue())
      nameMetaData.setValue(`${count+1}`)
      return sheetClient.insertSheet(sheetTitle + `#${count+1}`).addDeveloperMetadata('taskListId', id)
    }
  }

}

function firstImport() {
  var taskLists = Tasks.Tasklists.list();
  if (taskLists.items) {
    for (var i = 0; i < taskLists.items.length; i++) {
      var taskList = taskLists.items[i];
      let tasks = Tasks.Tasks.list(taskList.id, {
        showCompleted: true,
        showDeleted: true,
        showHidden: true,
      });

      Logger.log(taskList.title)
      Logger.log(tasks.items)

      var sheet = createSheet(taskList.id, taskList.title)

      if (tasks.items){
        for (let task of tasks.items){

          let finder = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
          .createDeveloperMetadataFinder()
          .onIntersectingLocations()
          .withKey('taskId')
          .withValue(task.id).find()[0]

          if(finder !== undefined){
            let loc = finder.getLocation()
            if(loc.getLocationType() === SpreadsheetApp.DeveloperMetadataLocationType.ROW){
              if(task.deleted){
                sheet.deleteRow(loc.getRow().getRowIndex())
                // loc.getRow().deleteCells(SpreadsheetApp.Dimension.ROWS)
              }
              else{
                let row = sheet.getRange(loc.getRow().getRowIndex(), 1, 1, 4)
                // Logger.log("%d %d", row.getNumRows(), row.getNumColumns())
                row.setValues([[task.title , task.notes, task.kind, task.completed]])
              }
            }
          }
          else{
            // Logger.log("%s %s %s %s", task.title , task.notes, task.kind, task.completed)
            sheet.appendRow([task.title , task.notes, task.kind, task.completed])
            let row = sheet.getRange(sheet.getLastRow() + ":" + sheet.getLastRow())
            row.addDeveloperMetadata('taskId', task.id)
          }
          

        }
      }

    }
  } else {
    Logger.log('No task lists found.');
  }
  
}
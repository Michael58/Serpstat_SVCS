function getTaskProgress(taskID) {
  var formData = {
    'jsonrpc': '2.0',
    'id': '12345',
    'method': 'tasks.getTaskResult',
    'params': {
      'taskId': taskID
    }
  };
  var payload = JSON.stringify(formData);
  var options = {
    'method' : 'post',
    'payload' : payload
  };
  var accessToken = getAccessToken();
  
  // make an API request for the Task and get the information back
  var response = UrlFetchApp.fetch('https://serpstat.com/rt/api/v2?token=' + accessToken, options);  
  var rawData = JSON.parse(response.getContentText());
  
  if ('result' in rawData) {
    if ('progress' in rawData['result']) {
      // if progress is in the response, assign it to the Task
      return rawData['result']['progress'];
    }
    // if no progress is in the response, Task is finished
    return 'finished';
  } else {
    return 0;
  }
}

function refresh() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
  // we start from row number 6, and check if any Task is unfinished
  var rowIndex = 6;
  
  while (true) {
    if (!sheet.getRange(rowIndex, 8).isBlank()) {
      if (sheet.getRange(rowIndex, 9).getValue() != 'finished') {        
        // if the row is not finished, get a progress update
        var taskProgress = getTaskProgress(sheet.getRange(rowIndex, 8).getValue());
        sheet.getRange(rowIndex, 9).setValue(taskProgress);
      }
    } else {
      break;
    }
    rowIndex++;
  }
}



function getAccessToken() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Run');
  var accessToken = sheet.getRange('D5').getValue();
  return accessToken;
}

function parseRecords(responseData) {
  rows = [];
  
  for (i = 0; i < responseData['result']['tops'].length; i++) {
    r = responseData['result']['tops'][i];      
    var keyword = r['keyword'];   
    
    if (!('keyword_data' in r) || r['keyword_data'].length <= 0) {
      continue; 
    }      
    
    for (j = 0; j < r['keyword_data']['top'].length; j++) {
      top = r['keyword_data']['top'][j];
      var row = {};
      row['keyword'] = keyword;
      row['Result Type'] = 'Top';
      row['url'] = top['url'];
      row['domain'] = top['domain'];
      row['title'] = top['title'];
      row['snippet'] = top['snippet'];
      row['position'] = top['position'];
      row['text'] = '';
      rows.push(row);
    }
    
    for (j = 0; j < r['keyword_data']['ads']['1'].length; j++) {
      ads = r['keyword_data']['ads']['1'][j];
      var row = {};
      row['keyword'] = keyword;
      row['Result Type'] = 'Ads Top';
      row['url'] = ads['url'];
      row['domain'] = ads['domain'];
      row['title'] = ads['title'];
      row['snippet'] = '';
      row['text'] = ads['text'];
      row['position'] = ads['position'];
      rows.push(row);
    }
    
    for (j = 0; j < r['keyword_data']['ads']['3'].length; j++) {
      ads = r['keyword_data']['ads']['3'][j];
      var row = {};
      row['keyword'] = keyword;
      row['Result Type'] = 'Ads Bottom';
      row['url'] = ads['url'];
      row['domain'] = ads['domain'];
      row['title'] = ads['title'];
      row['snippet'] = '';
      row['text'] = ads['text'];
      row['position'] = ads['position'];
      rows.push(row);
    }
    
  }
  
  return rows;  
}

function getTaskRecords(accessToken, taskID) {
  var formData = {
    'jsonrpc': '2.0',
    'id': '12345',
    'method': 'tasks.getTaskResult',
    'params': {
      'taskId': taskID
    }
  };
  var payload = JSON.stringify(formData);
  var options = {
    'method' : 'post',
    'payload' : payload,
  };
  
  var response = UrlFetchApp.fetch('https://serpstat.com/rt/api/v2?token=' + accessToken, options);  
  var responseData = JSON.parse(response.getContentText()); 
  var rawData = [];
  
  while (true) {
    var parsedData = parseRecords(responseData);
    
    for (index = 0; index < parsedData.length; ++index) {
      rawData.push(parsedData[index]);
    }
    
    // break the loop if no more record pages
    if (!('next_page' in responseData) && !('next_page' in responseData['result'])) {
      break;
    }
    
    var nextPageNumber = responseData['result']['next_page'];
    
    formData = {
      'jsonrpc': '2.0',
      'id': '12345',
      'method': 'tasks.getTaskResult',
      'params': {
        'taskId': taskID,
        'page': nextPageNumber
      }
    };
    payload = JSON.stringify(formData);
    options = {
      'method' : 'post',
      'payload' : payload
    };
    response = UrlFetchApp.fetch('https://serpstat.com/rt/api/v2?token=' + accessToken, options);
    responseData = JSON.parse(response.getContentText());    
  }
  
  return rawData;  
}

function displayRecords() {
  // remove records from raw records tab
  var rawRecordsSheet = SpreadsheetApp.getActive().getSheetByName('Raw Records');
  rawRecordsSheet.clear();
  
  // set the headers for raw records sheet
  var columns = [[
    "keyword",
    "Result Type",
    "url",
    "domain",
    "title",
    "snippet",
    "position",
    "text",
  ]];  
  rawRecordsSheet.getRange(1, 1, 1, 8).setValues(columns);
  
  // Get Task ID of the Task that we want to get records from
  var sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
  var taskID = parseInt(sheet.getRange('I3').getValue());

  var accessToken = getAccessToken();
  
  // Get raw records via Serpstat API
  var taskData = getTaskRecords(accessToken, taskID);
  
  // Display row data inside Raw Records tab
  var rawSheetData = [];
  
  taskData.forEach(function (item) {
    var sheetRow = [];
    
    sheetRow.push(item['keyword']);
    sheetRow.push(item['Result Type']);
    sheetRow.push(item['url']);
    sheetRow.push(item['domain']);
    sheetRow.push(item['title']);
    sheetRow.push(item['snippet']);
    sheetRow.push(item['position']);
    sheetRow.push(item['text']);    
    
    rawSheetData.push(sheetRow);
  });
  rawRecordsSheet.getRange(2, 1, rawSheetData.length, 8).setValues(rawSheetData);
}

function getParameters() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Run');
  var accessToken = sheet.getRange('D5').getValue();
  var country = sheet.getRange('D7').getValue();
  var language = sheet.getRange('D8').getValue();
  var region = sheet.getRange('D9').getValue();
  var device = sheet.getRange('D10').getValue();
  var taskId = getNextId(); 
  
  var rowIndex = 6;
  var keywordList = [];
  while (true) {
    if (sheet.getRange(rowIndex, 6).isBlank()) {
      break;
    }
    var keyword = sheet.getRange(rowIndex, 6).getValue();
    keywordList.push(keyword);
    rowIndex++;
  }
  
  var parameters = {
    'accessToken': accessToken,
    'country': country,
    'language': language,
    'region': region,
    'keywords': keywordList.join(','),
    'keywordSize': keywordList.length,
    'device': device,
    'taskId': taskId
  }; 
  
  return parameters;
}

function getNextId() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
  var rowIndex = 6;
  var taskId = 1;
  while (true) {
    if (sheet.getRange(rowIndex, 3).isBlank()) {
      return taskId;
    }
    taskId = parseInt(sheet.getRange(rowIndex, 3).getValue()) + 1;
    rowIndex++;
  }
}

function getNextTaskRow() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
  var rowIndex = 6;
  while (true) {
    if (sheet.getRange(rowIndex, 3).isBlank()) {
      return rowIndex;
    }
    rowIndex++;
  }
}

function postTaskCreation(parameters) {  
  var formData = {
    'jsonrpc': '2.0',
    'id': parameters['taskId'],
    'method': 'tasks.addTask',
    'params': {
      'keywords': parameters['keywords'],
      'typeId': parseInt(parameters['device']),
      'seId': 1,
      'countryId': parseInt(parameters['country']),
      'regionId': parseInt(parameters['region']),
      'langId': parseInt(parameters['language'])
    }
  };
  var payload = JSON.stringify(formData);
  var options = {
    'method' : 'post',
    'payload' : payload,
  };
  
  var response = UrlFetchApp.fetch('https://serpstat.com/rt/api/v2?token=' + parameters['accessToken'], options);  
  var responseData = JSON.parse(response.getContentText());  
  
  if (response.getResponseCode() != 200) {
    SpreadsheetApp.getUi().alert('Error creating task');
    throw 'Error creating task';
  }
  
  if ('error' in responseData) {
    SpreadsheetApp.getUi().alert(responseData['error']['message']);
    throw responseData['error']['message'];
  }
  
  SpreadsheetApp.getUi().alert('Task successfully created');
  return responseData;  
}

function appendTask(parameters, response) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Tasks');
  var dateString = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd HH:mm:ss');
  var rowIndex = getNextTaskRow();
  
  sheet.getRange('C' + rowIndex).setValue(parameters['taskId'].toString());
  sheet.getRange('D' + rowIndex).setValue(dateString);
  sheet.getRange('E' + rowIndex).setValue(parameters['country'].toString());
  sheet.getRange('F' + rowIndex).setValue(parameters['language'].toString());
  sheet.getRange('G' + rowIndex).setValue(parameters['region'].toString());
  sheet.getRange('H' + rowIndex).setValue(response['result']['task_id'].toString());
  sheet.getRange('I' + rowIndex).setValue('0.00%');
  sheet.getRange('J' + rowIndex).setValue(parameters['keywordSize'].toString());  
  sheet.getRange('L' + rowIndex).setValue(parameters['device'].toString());
  
  sheet.getRange('C' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('D' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('E' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('F' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('G' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('H' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('I' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('J' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('K' + rowIndex).setBackground('#fff2cc');
  sheet.getRange('L' + rowIndex).setBackground('#fff2cc');
}

function createTask() {
  var parameters = getParameters();
  var response = postTaskCreation(parameters);  
  appendTask(parameters, response);  
}

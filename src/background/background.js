const googleSheetsUrlPattern = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/;
const spreadsheetIdPattern = /\/spreadsheets\/d\/([^/]+)/;
const cellPattern = /^[A-Za-z]+[0-9]{1,2}$/;
let spreadsheetId = null;
let activeSheetName = null;
let currentTabId = null;
const columnPattern = /^[A-Za-z]{1,2}$/;
let numericVal=null;
chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  console.log(tabId)
  
  if (changeInfo.status === 'complete' && googleSheetsUrlPattern.test(tab.url)) {
    const spreadsheetIdMatch = tab.url.match(spreadsheetIdPattern);
    if (spreadsheetIdMatch) {
      console.log(tab.url)
      currentTabId = tabId;
      spreadsheetId = spreadsheetIdMatch[1];
      activeSheetName = await getActiveSheetName(tabId);
      await authenticateUser();
     numericVal=null;
      // calculateAverageAndAddToRange(spreadsheetId,"A2:A10","A11");
      createEmptyTable(spreadsheetId,activeSheetName,"B3",5,5);
      batchClearValues(spreadsheetId,activeSheetName+"!1:1");
    }
  } 
});

chrome.tabs.onActivated.addListener(async (activeInfo) => {
  const tabId = activeInfo.tabId;
  chrome.tabs.get(tabId, async (tab) => {
    console.log(tabId);
    if (googleSheetsUrlPattern.test(tab.url)) {
      const spreadsheetIdMatch = tab.url.match(spreadsheetIdPattern);
      if (spreadsheetIdMatch) {
        console.log(tab.url);
        currentTabId = tabId;
        spreadsheetId = spreadsheetIdMatch[1];
        activeSheetName = await getActiveSheetName(tabId);
        await authenticateUser();
      }
    }
  });
});

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
chrome.runtime.onMessage.addListener(async(message, sender, sendResponse) => {
  if (spreadsheetId && activeSheetName && currentTabId && sender.tab && sender.tab.id === currentTabId) {
    if (message.action === 'recognizedText') {
      const recognizedText = message.text.toLowerCase();

      const response_entities_array = await extractEntitiesFromText(recognizedText)
      console.log('Recognized Text:', recognizedText);
      console.log('model response:', response_entities_array);
      // const wordArray = recognizedText.split(' ');
      // if (wordArray[0].toLowerCase() === 'delete' && cellPattern.test(wordArray[1])) {
      //   clearParticularCell(spreadsheetId, activeSheetName + '!' + wordArray[1].toUpperCase());
      // } else if (wordArray[0].toLowerCase() === 'insert' && cellPattern.test(wordArray[1])) {
      //   let val = '';
      //   for (let i = 2; i < wordArray.length; i++) {
      //     if (i === 2) {
      //       val += wordArray[i];
      //     } else {
      //       val += ' ' + wordArray[i];
      //     }
      //   }

      response_entities_array.forEach((response_entities) => {
        if (response_entities.commands[0].toLowerCase() === 'delete' && cellPattern.test(response_entities.cells[0])) {
          updateCellValue(
            spreadsheetId,
            activeSheetName + '!' + response_entities.cells[0].toUpperCase(),
            '',
            'USER_ENTERED',
          );
        } else if (
          response_entities.commands[0].toLowerCase() === 'insert' &&
          cellPattern.test(response_entities.cells[0])
        ) {
          let val = response_entities.values[0];
          console.log(val);
          updateCellValue(
            spreadsheetId,
            activeSheetName + '!' + response_entities.cells[0].toUpperCase(),
            val,
            'USER_ENTERED',
          );
        } else if (response_entities.commands[0].toLowerCase() === 'delete') {
          const alphaVal = (s) => s.toLowerCase().charCodeAt(0) - 97 + 1;
          if (response_entities.column.length != 0) {
            deleteColumn(spreadsheetId, alphaVal(response_entities.column[0]));
          }
          if (response_entities.row.length !== 0) {
            deleteRow(spreadsheetId, response_entities.row[0]);
          }
        } else if (response_entities.commands[0].toLowerCase() === 'replace') {
          // let find = '';
          // let replace = '';
          // let i = 1;
          // let isReplace = false;

          // for (i; i < wordArray.length; i++) {
          //   if (wordArray[i] == 'with') {
          //     isReplace = true;
          //     continue;
          //   }

          //   if (isReplace) {
          //     if (replace === '') {
          //       replace += wordArray[i];
          //     } else {
          //       replace += ' ' + wordArray[i];
          //     }
          //   } else {
          //     if (find === '') {
          //       find += wordArray[i];
          //     } else {
          //       find += ' ' + wordArray[i];
          //     }
          //   }
          // }

          let find = response_entities.values[0];
          let replace = response_entities.values[1];
          console.log('Find : ' + find + ' Replace : ' + replace);
          find_replace(spreadsheetId, find, replace);

          // find_replace(spreadsheetId, 'king', 'prathamesh');
        } else if (
          response_entities.commands[0].toLowerCase() === 'bold' &&
          cellPattern.test(response_entities.cells[0])
        ) {
          bold_text(
            spreadsheetId,
            activeSheetName +
              '!' +
              response_entities.cells[0].toUpperCase() +
              ':' +
              response_entities.cells[0].toUpperCase(),
          );
        } else if (
          response_entities.commands[0].toLowerCase() === 'italic' &&
          cellPattern.test(response_entities.cells[0])
        ) {
          italic_text(
            spreadsheetId,
            activeSheetName +
              '!' +
              response_entities.cells[0].toUpperCase() +
              ':' +
              response_entities.cells[0].toUpperCase(),
          );
        } else if (response_entities.commands[0].toLowerCase() === 'merge') {
          merge_cells(
            spreadsheetId,
            activeSheetName +
              '!' +
              response_entities.cells[0].toUpperCase() +
              ':' +
              response_entities.cells[1].toUpperCase(),
          );
        } else if (response_entities.commands[0].toLowerCase() === 'chart') {
          insert_chart(
            spreadsheetId,
            activeSheetName +
              '!' +
              response_entities.cells[0].toUpperCase() +
              ':' +
              response_entities.cells[1].toUpperCase(),
            response_entities.chart[0].toUpperCase(),
          );
        }else if(response_entities.commands[0].toLowerCase() === 'sumall'){
          sumAndAddValuesToRange(
            spreadsheetId,
            activeSheetName +
              '!' +
              response_entities.cells[0].toUpperCase() +
              ':' +
              response_entities.cells[1].toUpperCase(),
              (response_entities.cells.length==3?response_entities.cells[2].toUpperCase():null)
          )
        }else if(response_entities.commands[0].toLowerCase() === 'multiplyall'){
          multiplyAndAddValuesToRange(spreadsheetId, activeSheetName +
              '!' +
              response_entities.cells[0].toUpperCase() +
              ':' +
              response_entities.cells[1].toUpperCase(), (response_entities.cells.length==3?response_entities.cells[2].toUpperCase():null));

        }else if(response_entities.commands[0].toLowerCase() === 'averageall'){
          calculateAverageAndAddToRange(spreadsheetId, activeSheetName +
              '!' +
              response_entities.cells[0].toUpperCase() +
              ':' +
              response_entities.cells[1].toUpperCase(),(response_entities.cells.length==3?response_entities.cells[2].toUpperCase():null));

        }

      });


    }
    
  }
});

const extractEntitiesFromText = async(text) => {
  try {
    const response = await fetch('http://127.0.0.1:5000/predict', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ text }),
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const dataArray = await response.json();
    console.log(dataArray);

    // Process each element in the array
    const extractedDataArray = dataArray.map((data) => {
      const values = data.values;
      const commands = data.type_labels;
      const cells = data.cells;
      const column = data.column;
      const row = data.row;
      const chart = data.chart;

      return { commands, cells, values, column, row, chart };
      

    });

    return extractedDataArray;

    // const data = await response.json();
    // console.log(data)
    // const values = data.values;
    // const commands = data.type_labels;
    // const cells = data.cells;
    // const column = data.column;
    // const row = data.row;
    // const chart = data.chart;

    // return {commands : commands,cells:cells,values:values, column:column, row: row, chart: chart}
  } catch (error) {
    console.error('Fetch error:', error);
  }
}
// mathmatical operations

async function sumAndAddValuesToRange(spreadsheetId, range,cell) {
  try {
    const token = await checkAuthentication(); // Assuming you have a function to check authentication

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`;

    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        Authorization:` Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to retrieve values from range: ${responseData.error.message}`);
    }

    const values = responseData.values;

    // Calculate the sum of numeric values
    const numericSum = values
      .flat() // Flatten the 2D array
      .filter(value => !isNaN(value)) // Filter out non-numeric values
      .reduce((sum, value) => sum + Number(value), 0);

    console.log(`Sum of numeric values in range ${range}: ${numericSum}`);
    numericVal=numericSum;

    if(cell){
    updateCellValue(spreadsheetId, activeSheetName + '!' + cell.toUpperCase(), numericSum, 'USER_ENTERED')
    }
    

    // await addValuesToRange(spreadsheetId, sumRange, sumValues);
  } catch (error) {
    console.error('Error summing and adding values to range:', error);
  }
}

// Function to multiply and add numeric values to a range in a Google Sheet
async function multiplyAndAddValuesToRange(spreadsheetId, range,cell) {
  try {
    const token = await checkAuthentication(); // Assuming you have a function to check authentication

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`;

    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to retrieve values from range: ${responseData.error.message}`);
    }

    const values = responseData.values;
    console.log(values);

    // Calculate the product of numeric values
   const numericProduct = values
      .flat() // Flatten the 2D array
      .map(value => isNaN(Number(value)) ? 1 : Number(value)) // Convert to numbers and handle non-numeric values
      .reduce((product, value) => product * value, 1);

    console.log(`Product of numeric values in range ${range}: ${numericProduct}`);

    // Now, you can add the product to a specific cell or range using the addValuesToRange function
    const productRange = 'Sheet1!D1'; // Change this to the cell where you want to put the product
    const productValues = [[numericProduct]];
      if(cell){
      
    updateCellValue(spreadsheetId, activeSheetName + '!' + cell.toUpperCase(), numericProduct, 'USER_ENTERED')
    }else{
      numericVal=numericProduct
    }



  } catch (error) {
    console.error('Error multiplying and adding values to range:', error);
  }
}


// Function to calculate average and add to a range in a Google Sheet
async function calculateAverageAndAddToRange(spreadsheetId, range,cell) {
  try {
    const token = await checkAuthentication(); // Assuming you have a function to check authentication

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`;

    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to retrieve values from range: ${responseData.error.message}`);
    }

    const values = responseData.values;

    // Calculate the average of numeric values
    const numericValues = values
      .flat() // Flatten the 2D array
      .filter(value => !isNaN(value)) // Filter out non-numeric values
      .map(Number);

    const sum = numericValues.reduce((acc, value) => acc + value, 0);
    const average = sum / numericValues.length;

    console.log(`Average of numeric values in range ${range}: ${average}`);

    // Now, you can add the average to a specific cell or range using the addValuesToRange function
    const averageRange = 'Sheet1!E1'; // Change this to the cell where you want to put the average
    const averageValues = [[average]];
       if(cell){
      
    updateCellValue(spreadsheetId, activeSheetName + '!' + cell.toUpperCase(), average, 'USER_ENTERED')
    }else{
      numericVal=average;
    }


  } catch (error) {
    console.error('Error calculating average and adding to range:', error);
  }
}



//table
async function createEmptyTable(spreadsheetId, sheetName, startCell, numRows, numCols) {
  try {
    const token = await checkAuthentication(); // Assuming you have a function to check authentication

    const apiUrl =` https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(sheetName)}:batchUpdate`;

    // Convert the startCell to row and column indices
    const startRowIndex = parseInt(startCell.match(/\d+/)[0]) - 1;
    const startColumnIndex = startCell.match(/[A-Z]+/)[0].split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 'A'.charCodeAt(0) + 1, 0) - 1;

    // Create the requests for batch update
    const requests = [
      {
        updateCells: {
          start: {
            sheetId: 0, // Assuming the sheet has ID 0
            rowIndex: startRowIndex,
            columnIndex: startColumnIndex,
          },
          rows: Array.from({ length: numRows }, () => ({
            values: Array.from({ length: numCols }, () => ({})), // Empty values
          })),
          fields: 'userEnteredValue', // Only update userEnteredValue field
        },
      },
    ];

    // Send the batch update request
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ requests }),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to create table: ${responseData.error.message}`);
    }

    console.log(`Table created with ${numRows} rows and ${numCols} columns starting from ${startCell}`);
  } catch (error) {
    console.error('Error creating table:', error);
  }
}











const getActiveSheetName = (tabId) => {
  return new Promise(async (resolve, reject) => {
    try {
      const result = await chrome.scripting.executeScript({
        target: { tabId: tabId },
        function: async () => {
          const sheetTabs = document.querySelectorAll('.docs-sheet-tab');
          let activeSheetName = '';

          sheetTabs.forEach((tab) => {
            if (tab.classList.contains('docs-sheet-active-tab')) {
              const sheetName = tab.querySelector('.docs-sheet-tab-name').textContent;
              activeSheetName = sheetName;
            }
          });
          return activeSheetName;
        },
      });
      resolve(result[0].result);
    } catch (error) {
      reject(error);
    }
  });
};

const checkAuthentication = () => {
  return new Promise((resolve) => {
    chrome.identity.getAuthToken({ interactive: false }, (token) => {
      resolve(token);
      console.log(token);
    });
  });
};

const authenticateUser = async () => {
  const isAuthenticated = await checkAuthentication();
  if (!isAuthenticated) {
    const token = await new Promise((resolve) => {
      chrome.identity.getAuthToken({ interactive: true }, (token) => {
        resolve(token);
      });
    });

    if (token) {
      console.log('Authenticated with token:', token);
    } else {
      console.error('Authentication error');
    }
  }
};

const unauthenticateUser = () => {
  chrome.identity.getAuthToken({ interactive: false }, (token) => {
    if (!chrome.runtime.lastError) {
      chrome.identity.removeCachedAuthToken({ token }, () => {
        console.log('User has been unauthenticated.');
      });
    }
  });
};

const updateCellValue = async (spreadsheetId, range, newValue, valueInputOption) => {
  try {
    // while(!newValue || !numericVal);
    await sleep(1000);
    console.log("1st parameter:",newValue);
    console.log("global paramet",numericVal);
    const token = await checkAuthentication();
    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      spreadsheetId,
    )}/values/${encodeURIComponent(range)}?valueInputOption=${encodeURIComponent(valueInputOption)}`;
    const requestBody = {
      values: [[(newValue?newValue:numericVal)]],
    };
    numericVal=null;
    const response = await fetch(apiUrl, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to update cell value: ${responseData.error.message}`);
    }

    console.log(`Updated cell value: ${newValue}`);
  } catch (error) {
    console.error('Error updating cell value:', error);
  }
};

const batchClearValues = async (spreadsheetId, ranges) => {
  try {
    const token = await checkAuthentication();
    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values:batchClear`;

    const requestBody = {
      ranges: ranges,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    if (response.ok) {
      console.log('Batch clear operation completed successfully.');
    } else {
      const responseData = await response.json();
      console.error(`Failed to perform batch clear operation: ${responseData.error.message}`);
    }
  } catch (error) {
    console.error('Error performing batch clear operation:', error);
  }
};





const deleteColumn = async (spreadsheetId, columnIndex) => {
  try {
    const token = await checkAuthentication();

    console.log('columnIndex :' + columnIndex);

    const sheetID = parseInt(spreadsheetId);
    console.log('spreadsheetId :' + spreadsheetId);
    console.log('sheetID :' + sheetID);
    const startIndex = parseInt(columnIndex) - 1;
    const endIndex = parseInt(columnIndex);
    console.log('endIndex :' + endIndex);

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;

    const requests = [
      {
        deleteDimension: {
          range: {
            // sheetId: 0,
            dimension: 'COLUMNS',
            startIndex: startIndex,
            endIndex: endIndex,
          },
        },
      },
    ];

    const requestBody = {
      requests: requests,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to delete column: ${responseData.error.message}`);
    }

    console.log(`Deleted column at index ${columnIndex}`);
  } catch (error) {
    console.error('Error deleting column:', error);
  }
};




const deleteRow = async (spreadsheetId, rowIndex) => {
  try {
    const token = await checkAuthentication();

    console.log('rowIndex :' + rowIndex);

    const sheetID = parseInt(spreadsheetId);
    console.log('spreadsheetId :' + spreadsheetId);
    console.log('sheetID :' + sheetID);
    const startIndex = parseInt(rowIndex) - 1;
    const endIndex = parseInt(rowIndex);
    console.log('endIndex :' + endIndex);

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;

    const requests = [
      {
        deleteDimension: {
          range: {
            // sheetId: 0,
            dimension: 'ROWS',
            startIndex: startIndex,
            endIndex: endIndex,
          },
        },
      },
    ];

    const requestBody = {
      requests: requests,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to delete row: ${responseData.error.message}`);
    }

    console.log(`Deleted row at index ${rowIndex}`);
  } catch (error) {
    console.error('Error deleting row:', error);
  }
};


const find_replace = async (spreadsheetId, find, replace) => {
  try {
    console.log('Find : ' + find + ' Replace : ' + replace);
    const token = await checkAuthentication();

    const sheetID = parseInt(spreadsheetId);
    console.log('spreadsheetId :' + spreadsheetId);
    console.log('sheetID :' + sheetID);

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;

    const requests = [
      {
        findReplace: {
          find: String(find),
          replacement: String(replace),

          matchCase: false,
          matchEntireCell: false,

          range: {},
        },
      },
    ];

    const requestBody = {
      requests: requests,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to replace values: ${responseData.error.message}`);
    }

    console.log(`Replaced the values`);
  } catch (error) {
    console.error('Error replacing values', error);
  }
};



function columnLetterToIndex(letter = null) {
  letter = letter.toUpperCase();
  return [...letter].reduce((c, e, i, a) => (c += (e.charCodeAt(0) - 64) * Math.pow(26, a.length - i - 1)), -1);
}

function columnIndexToLetter(index = null) {
  return (a = Math.floor(index / 26)) >= 0 ? columnIndexToLetter(a - 1) + String.fromCharCode(65 + (index % 26)) : '';
}

function convA1NotationToGridRange(sheetId, a1Notation) {
  const { col, row } = a1Notation
    .toUpperCase()
    .split('!')
    .map((f) => f.split(':'))
    .pop()
    .reduce(
      (o, g) => {
        var [r1, r2] = ['[A-Z]+', '[0-9]+'].map((h) => g.match(new RegExp(h)));
        o.col.push(r1 && columnLetterToIndex(r1[0]));
        o.row.push(r2 && Number(r2[0]));
        return o;
      },
      { col: [], row: [] },
    );
  col.sort((a, b) => (a > b ? 1 : -1));
  row.sort((a, b) => (a > b ? 1 : -1));
  const [start, end] = col.map((e, i) => ({ col: e, row: row[i] }));
  const gridRange = {
    startRowIndex: start?.row && start.row - 1,
    endRowIndex: end?.row ? end.row : start.row,
    startColumnIndex: start && start.col,
    endColumnIndex: end ? end.col + 1 : 1,
  };
  if (gridRange.startRowIndex === null) {
    gridRange.startRowIndex = 0;
    delete gridRange.endRowIndex;
  }
  if (gridRange.startColumnIndex === null) {
    gridRange.startColumnIndex = 0;
    delete gridRange.endColumnIndex;
  }
  return gridRange;
}

const bold_text = async (spreadsheetId, range) => {
  try {
    const token = await checkAuthentication();

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;

    var gridRange = convA1NotationToGridRange(spreadsheetId, range);

    const requests = [
      {
        repeatCell: {
          range: {
            startRowIndex: gridRange.startRowIndex,
            endRowIndex: gridRange.endRowIndex,
            startColumnIndex: gridRange.startColumnIndex,
            endColumnIndex: gridRange.endColumnIndex,
          },
          cell: {
            userEnteredFormat: {
              textFormat: {
                bold: true,
              },
            },
          },
          fields: 'userEnteredFormat(textFormat)',
        },
      },
    ];

    const requestBody = {
      requests: requests,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to make bold: ${responseData.error.message}`);
    }

    console.log(`Made cell bold`);
  } catch (error) {
    console.error('Error to make bold', error);
  }
};

const italic_text = async (spreadsheetId, range) => {
  try {
    const token = await checkAuthentication();

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;

    var gridRange = convA1NotationToGridRange(spreadsheetId, range);

    const requests = [
      {
        repeatCell: {
          range: {
            startRowIndex: gridRange.startRowIndex,
            endRowIndex: gridRange.endRowIndex,
            startColumnIndex: gridRange.startColumnIndex,
            endColumnIndex: gridRange.endColumnIndex,
          },
          cell: {
            userEnteredFormat: {
              textFormat: {
                italic: true,
              },
            },
          },
          fields: 'userEnteredFormat(textFormat)',
        },
      },
    ];

    const requestBody = {
      requests: requests,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to make italic: ${responseData.error.message}`);
    }

    console.log(`Made cell italic`);
  } catch (error) {
    console.error('Error to make italic', error);
  }
};

const merge_cells = async (spreadsheetId, range) => {
  try {
    const token = await checkAuthentication();

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;

    var gridRange = convA1NotationToGridRange(spreadsheetId, range);

    const requests = [
      {
        mergeCells: {
          range: {
            startRowIndex: gridRange.startRowIndex,
            endRowIndex: gridRange.endRowIndex,
            startColumnIndex: gridRange.startColumnIndex,
            endColumnIndex: gridRange.endColumnIndex,
          },
          mergeType: 'MERGE_ALL',
        },
      },
    ];

    const requestBody = {
      requests: requests,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to merge: ${responseData.error.message}`);
    }

    console.log(`merged successfully`);
  } catch (error) {
    console.error('Error to merge', error);
  }
};


const insert_chart = async (spreadsheetId, range, chart) => {
  try {
    const token = await checkAuthentication();

    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;

    const gridRange = convA1NotationToGridRange(spreadsheetId, range);

    const columnCount = gridRange.endColumnIndex - gridRange.startColumnIndex + 1;
    const seriesRequests = [];

    for (let i = 1; i < columnCount; i++) {
      const seriesRequest = {
        series: {
          sourceRange: {
            sources: [
              {
                startRowIndex: gridRange.startRowIndex,
                endRowIndex: gridRange.endRowIndex,
                startColumnIndex: gridRange.startColumnIndex + i,
                endColumnIndex: gridRange.startColumnIndex + i + 1,
              },
            ],
          },
        },
        targetAxis: 'LEFT_AXIS',
      };
      
      seriesRequests.push(seriesRequest);
    }

    const chartRequest = {
      addChart: {
        chart: {
          spec: {
            title: `${chart} Chart`,
            basicChart: {
              chartType: chart,
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                // x-axis
                {
                  position: 'BOTTOM_AXIS',
                  title: 'X-AXIS',
                },
                // y-axis
                {
                  position: 'LEFT_AXIS',
                  title: 'Y-AXIS',
                },
              ],
              domains: [
                {
                  domain: {
                    sourceRange: {
                      sources: [
                        {
                          startRowIndex: gridRange.startRowIndex,
                          endRowIndex: gridRange.endRowIndex,
                          startColumnIndex: gridRange.startColumnIndex,
                          endColumnIndex: gridRange.startColumnIndex + 1,
                        },
                      ],
                    },
                  },
                },
              ],
              series: seriesRequests,
              headerCount: 1,
            },
          },
          position: {
            newSheet: true,
          },
        },
      },
    };

    const requests = [chartRequest];

    const requestBody = {
      requests: requests,
    };

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const responseData = await response.json();

    if (!response.ok) {
      throw new Error(`Failed to merge: ${responseData.error.message}`);
    }

    console.log(`Chart added successfully`);
  } catch (error) {
    console.error('Error adding chart', error);
  }
};

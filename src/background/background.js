const googleSheetsUrlPattern = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/;
const spreadsheetIdPattern = /\/spreadsheets\/d\/([^/]+)/;
const cellPattern = /^[A-Za-z]+[0-9]{1,2}$/;
let spreadsheetId = null;
let activeSheetName = null;
let currentTabId = null;
const columnPattern = /^[A-Za-z]{1,2}$/;

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
      batchClearValues(spreadsheetId,activeSheetName+"!1:1")
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


chrome.runtime.onMessage.addListener(async(message, sender, sendResponse) => {
  if (spreadsheetId && activeSheetName && currentTabId && sender.tab && sender.tab.id === currentTabId) {
    if (message.action === 'recognizedText') {
      const recognizedText = message.text;

      const response_entities = await extractEntitiesFromText(recognizedText)
      console.log('Recognized Text:', recognizedText);
      console.log('model response:', response_entities);
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
      if (response_entities.commands[0].toLowerCase() === 'type b' && cellPattern.test(response_entities.cells[0])) {
        updateCellValue(spreadsheetId, activeSheetName + '!' + response_entities.cells[0].toUpperCase(),'', 'USER_ENTERED');
      } else if (response_entities.commands[0].toLowerCase() === 'type a' && cellPattern.test(response_entities.cells[0])) {
        let val = response_entities.values[0]
        console.log(val);
        updateCellValue(spreadsheetId, activeSheetName + '!' + response_entities.cells[0].toUpperCase(), val, 'USER_ENTERED');
      }
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

    const data = await response.json();
    console.log(data)
    const values = data.values;
    const commands = data.type_labels;
    const cells = data.cells;

    return {commands : commands,cells:cells,values:values}
  } catch (error) {
    console.error('Fetch error:', error);
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
    const token = await checkAuthentication();
    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      spreadsheetId,
    )}/values/${encodeURIComponent(range)}?valueInputOption=${encodeURIComponent(valueInputOption)}`;
    const requestBody = {
      values: [[newValue]],
    };
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
            sheetId: 0,
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
            sheetId: 0,
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
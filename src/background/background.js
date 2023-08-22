chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  const googleSheetsUrlPattern = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/;
  const spreadsheetIdPattern = /\/spreadsheets\/d\/([^/]+)/;
  var cellPattern = /^[A-Za-z]+[0-9]{1,2}$/;

  if (changeInfo.status === 'complete' && googleSheetsUrlPattern.test(tab.url)) {
    const spreadsheetIdMatch = tab.url.match(spreadsheetIdPattern);
    if (spreadsheetIdMatch) {
      const spreadsheetId = spreadsheetIdMatch[1];
      const activeSheetName = await getActiveSheetName(tabId);
      await authenticateUser();

      chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        if (message.action === 'recognizedText') {
          const recognizedText = message.text;
          console.log('Recognized Text:', recognizedText);
          const wordArray = recognizedText.split(' ');
          if (wordArray[0].toLowerCase() == 'delete' && cellPattern.test(wordArray[1])) {
            clearParticularCell(spreadsheetId, activeSheetName + '!' + wordArray[1].toUpperCase());
          } else if (wordArray[0].toLowerCase() == 'insert' && cellPattern.test(wordArray[1])) {
            let val="";
            for(let i=2;i<wordArray.length;i++){
              if(i==2){
                 val+=wordArray[i];
              }else{
                 val+=" "+wordArray[i];
              }
             
            }
            console.log(val);
            updateCellValue(
              spreadsheetId,
              activeSheetName + '!' + wordArray[1].toUpperCase(),
              val,
              'USER_ENTERED',
            );
          }
        }
      });
    }
  }
});

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

const fetchDataFromGoogleSheet = async (spreadsheetId, range) => {
  try {
    const token = await checkAuthentication();

    const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });
    if (!response.ok) {
      throw new Error('Failed to fetch data from Google Sheets.');
    }

    const data = await response.json();
    console.log(data);
    return data;
  } catch (error) {
    console.error('API request error:', error);
    throw error;
  }
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

const clearParticularCell = async (spreadsheetId, range) => {
  try {
    const token = await checkAuthentication();
    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      spreadsheetId,
    )}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
    const requestBody = {
      values: [['']],
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
  } catch (error) {
    console.error('Error updating cell value:', error);
  }
};

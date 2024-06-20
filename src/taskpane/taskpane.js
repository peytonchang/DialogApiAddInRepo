(function () {
    let currentDialog = null;
    let loggedIn = false; // Flag to check login status
  
    Office.onReady((info) => {
        // Ensures the add-in initializes correctly depending on the host
        if (info.host === Office.HostType.Excel) {
            // Add event listener to login button
            document.getElementById('loginButton').addEventListener('click', login);
        }
    });
    
    function login() {
        const passwordInput = document.getElementById('passwordInput').value;
        const universalPassword = "BlueSage"; // Set your universal password
    
        if (passwordInput === universalPassword) {
            // Hide login section and show main content upon successful login
            document.getElementById('loginSection').style.display = 'none';
            document.getElementById('mainContent').style.display = 'block';
            console.log('Login successful.');
        } else {
            alert("Incorrect password. Please try again.");
        }
    }
  
    function checkLoginState() {
        var storedLoggedIn = localStorage.getItem('loggedIn') === 'true';
        if (storedLoggedIn) {
            loggedIn = true;
            document.getElementById('loginSection').style.display = 'none';
            document.getElementById('mainContent').style.display = 'block';
        }
    }
  
    function openDialog() {
        if (!loggedIn) {
            alert("Please log in first.");
            return;
        }
  
        console.dir(Office.context.ui);  // Logs the Office context UI for debugging
        const dialogUrl = 'https://peytonchang.github.io/DialogApiAddInRepo/src/dialog.html';
        Office.context.ui.displayDialogAsync(dialogUrl, { height: 50, width: 50 }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Failed to open dialog: ' + result.error.message);
            } else {
                currentDialog = result.value;
                console.log('Dialog opened successfully.');
                currentDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessageFromDialog);
                currentDialog.addEventHandler(Office.EventType.DialogEventReceived, handleDialogEvent);
            }
        });
    }

  function processMessageFromDialog(arg) {
      const message = arg.message;
      console.log('Message from dialog: ', message);

      if (message === 'capture') {
          Excel.run(async (context) => {
              console.log('Processing capture request');
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const cell = sheet.getRange("A1");
              cell.load("values");
              await context.sync();
              const cellValue = cell.values[0][0];
              console.log('Captured value from A1: ', cellValue);
              // Send the captured value back to the dialog
              if (currentDialog) {
                  console.log('Sending value to dialog: ', cellValue);
                  console.dir(currentDialog);
                  currentDialog.messageChild(JSON.stringify({ value: cellValue }));
              } else {
                  console.error('No dialog instance found.');
              }
          }).catch((error) => {
              console.error('Error capturing value from A1: ', error);
          });
      } else if (message.startsWith('paste:')) {
          const valueToPaste = message.split(':')[1];
          Excel.run(async (context) => {
              console.log('Processing paste request with value: ', valueToPaste);
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const cell = sheet.getRange("A1");
              cell.values = [[valueToPaste]];
              await context.sync();
          }).catch((error) => {
              console.error('Error pasting value to A1: ', error);
          });
      }
  }

  function handleDialogEvent(event) {
      if (event.type === "dialogClosed") {
          currentDialog = null;
          console.log('Dialog closed.');
      }
  }
})();

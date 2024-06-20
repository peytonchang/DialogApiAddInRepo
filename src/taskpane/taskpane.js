Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('loginButton')?.addEventListener('click', login);
        checkLoginState();
    }
});

function checkLoginState() {
    const loggedIn = localStorage.getItem('loggedIn') === 'true';
    if (!loggedIn && window.location.pathname.endsWith('taskpane.html')) {
        window.location.href = 'login.html';
    } else if (loggedIn && window.location.pathname.endsWith('login.html')) {
        window.location.href = 'taskpane.html';
    }
}

function login() {
    const passwordInput = document.getElementById('passwordInput').value;
    const universalPassword = "YourSecurePassword"; // Set your universal password

    if (passwordInput === universalPassword) {
        localStorage.setItem('loggedIn', 'true');
        window.location.href = 'taskpane.html'; // Redirect to the main content page
    } else {
        document.getElementById('errorMessage').style.display = 'block';
    }
}

function openDialog() {
    const dialogUrl = 'https://peytonchang.github.io/DialogApiAddInRepo/src/dialog.html'; // Adjust as necessary
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
            if (currentDialog) {
                console.log('Sending value to dialog: ', cellValue);
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

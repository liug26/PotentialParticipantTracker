// GLOBAL VARIABLES -----------------------------------------------------------:


// On Open function to create custom menu
function onOpen() {
	createMenuWithSubMenu();
}

// Generate Custom Menu
function createMenuWithSubMenu() {
	SpreadsheetApp.getUi()
		.createMenu('🚀 Automation')
		.addItem('Process Emails 📧', 'processAllEmails')
		.addSeparator()
    .addItem('Process Forms 📝', 'processAllForms')
		.addSeparator()
    .addItem('Undo 😬', 'undo')
    .addSeparator()
		.addItem('Cleanup Labels 🧹', 'cleanupLabels')
		.addSeparator()
		.addItem('Cleanup Backup 🧹', 'cleanupBackups')
		.addToUi();
}

// Spreadsheet Functions  -----------------------------------------------------:


// Undo Main function
function undoMain() {
	// Password check
	let userAuthenticationResponse = userAuthentication();
	// If user selects a valid uuid continue, else exit script.
	if (userAuthenticationResponse === 'pass') {
		// Get uuid from user to undo
		let userJobSelectionResponse = userJobSelection();
		// If invalid selection, exit script.
		if (userJobSelectionResponse === 'fail') {
			return;
		}
		// Gather emails based on uuid
		let allThreads = gatherEmailsUndo(userJobSelectionResponse);
		// Undo gmail process.
		for (let i = 0; i < allThreads.length; i++) {
			undoProcessThread(allThreads[i], userJobSelectionResponse);
		}
		// Clear unused labels
		cleanupLabelsInsecure();
		// Restore spreadsheet from backup uuid
		restoreSpreadsheet(userJobSelectionResponse);
	} else {
		return;
	}
	SpreadsheetApp.getUi().alert(
		'🙂 Success!\n\nPLEASE CLOSE THIS SPREADSHEET AND OPEN THE POTENTIAL PARTICIPANT TRACKER SPREADSHEET - RESTORED FILE FROM THE HOME DIRECTORY IN GOOGLE DRIVE TO COMPLETE THE UNDO PROCESS'
	);
}

// Cleanup labels function
function cleanupLabels() {
	// Password check
	let userAuthenticationResponse = userAuthentication();
	// If user enters the correct password, continue, else exit script.
	if (userAuthenticationResponse === 'pass') {
		let allLabels = getLabelList();
		for (let i = 0; i < allLabels.length; i++) {
			let label = GmailApp.getUserLabelByName(allLabels[i]);
			if (label.getThreads()[0] == null && isUUID(label.getName())) {
				Logger.log(`${label.getName()} will be deleted`);
				label.deleteLabel();
			} else {
				Logger.log(`${label.getName()} contains active threads`);
			}
		}
	} else if (userAuthenticationResponse === 'fail') {
		Logger.log('ERROR');
		return;
	}
	SpreadsheetApp.getActive().toast('Script is Complete', '🙂 Success!');
}

// Process Forms
function processForm() {
	// Generate job-id
	const jobID = generateFormJobID();
	const timestamp = getDate();

	// Authenticate user
	let userAuthenticationResponse = userAuthentication();
	if (userAuthenticationResponse == 'pass') {
		// Make copy of current spreadsheet
		backupSpreadsheet(jobID, timestamp);
		// Get array of all form objects
		let allFormResponses = getFormResponses();
		// Process each object
		for (let i = 0; i < allFormResponses.length; i++) {
			processFormResponse(allFormResponses[i], jobID, timestamp);
		}
	} else return;
	SpreadsheetApp.getActive().toast('Script is Complete', '🙂 Success!');
}

// PROGRAM FUNCTIONS ------------------------------------------------------------:



// Get a list of all labels to delete unused labels
function getLabelList() {
	let labelList = [];
	let labelObject = GmailApp.getUserLabels();
	for (let i = 0; i < labelObject.length; i++) {
		labelList[i] = labelObject[i].getName();
	}
	return labelList;
}

// User input to select which job to reverse.
function userJobSelection() {
	let label = getLabelList();
	let ui = SpreadsheetApp.getUi();
	let uuidLabels = [];
	for (let i = 0; i < label.length; i++) {
		if (isUUID(label[i])) {
			uuidLabels.push(label[i]);
		}
	}
	let result = ui.prompt(
		'Select a label',
		`Please select the job to undo:\n${uuidLabels.join('\n')}`,
		ui.ButtonSet.OK_CANCEL
	);
	let button = result.getSelectedButton();
	let userResponse = result.getResponseText();
	// User clicks cancel
	if (button == ui.Button.CANCEL) {
		ui.alert('Not running script.');
		return 'fail';
		// User enters incorrect label.
	} else if (button == ui.Button.OK && uuidLabels.includes(userResponse)) {
		ui.alert(`Confirmation received.\nReversing operations on ${userResponse}`);
		return userResponse;
		// Edge case catcher
	} else if (button == ui.Button.OK && !uuidLabels.includes(userResponse)) {
		ui.alert('ERROR: Label not on list');
		return 'fail';
	}
}

// Gather emails based on user input label
function gatherEmailsUndo(uuid) {
	let label = GmailApp.getUserLabelByName(`${uuid}`);
	let threads = label.getThreads();
	let allThreads = [];
	for (let i = 0; i < threads.length; i++) {
		let thread = threads[i];
		let currentMessage = threads[i].getMessages()[0];
		let currentMessageBody = threads[i].getMessages()[0].getPlainBody();
		let currentMessageSubject = threads[i].getMessages()[0].getSubject();
		let currentMessageSender = threads[i].getMessages()[0].getFrom();
		allThreads[i] = [
			currentMessageBody,
			currentMessageSubject,
			currentMessageSender,
			currentMessage,
			thread,
		];
	}
	return allThreads;
}


// Process Form Responses
// Process Each Form Response
function processFormResponse(response, jobid, timestamp) {
	const name = response.getItemResponses()[0].getResponse();
	const phone = response.getItemResponses()[1].getResponse();
	const email = response.getItemResponses()[2].getResponse();
	const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
	const MASTER_LIST = SPREADSHEET.getSheetByName('Master List');
	const INITIAL_OUTREACH = SPREADSHEET.getSheetByName('Initial Outreach');
	const DQ_LIST = SPREADSHEET.getSheetByName("DQ'd");
	const MASTER_LIST_EMAILS = MASTER_LIST.getRange('C2:C1000')
		.getValues()
		.toLocaleString()
		.replace(/,/g, '\n');
	const DQ_LIST_EMAILS = DQ_LIST.getRange('C2:C1000')
		.getValues()
		.toLocaleString()
		.replace(/,/g, '\n');
	const MASTER_LIST_PHONES = MASTER_LIST.getRange('B2:B1000')
		.getValues()
		.toString()
		.replace(/,/g, '\n')
		.replace(/[^\d\n]/g, '');
	const DQ_LIST_PHONES = DQ_LIST.getRange('B2:B1000')
		.getValues()
		.toString()
		.replace(/,/g, '\n')
		.replace(/[^\d\n]/g, '');
	const phoneCheck = phone.replace(/\D/g, '');
	if (MASTER_LIST_EMAILS.includes(email) || DQ_LIST_EMAILS.includes(email)) {
		Logger.log(`Participant ${name}'s email is already on the Master List.`);
	} else if (
		MASTER_LIST_PHONES.includes(phoneCheck) ||
		DQ_LIST_PHONES.includes(phoneCheck)
	) {
		Logger.log(
			`Participant ${name}'s phone number is already on the Master List.`
		);
	} else {
		Logger.log(
			`${name} is a new potential participant. Adding to Master List and Initial Initial Outreach Tabs...`
		);
		// Add participant to Master List and Initial Outreach tabs
		MASTER_LIST.appendRow([`${name}`, `${phone}`, `${email}`]);
		INITIAL_OUTREACH.appendRow([
			`${name}`,
			`${phone}`,
			`${email}`,
			`${timestamp} (Computer)`,
		]);
		// Create draft email and mark with uuid label
		let myLabel = GmailApp.createLabel(`${jobid}`);
		try {
			let myDraft = GmailApp.createDraft(
				`${email}`,
				`Thank You For Contacting Us About Our Study...`,
				``,
				{
					name: `UCLA Translational Neuroimaging Lab`,
					htmlBody: `<p>Hello,</p><p>Thank you for contacting the Translational Neuroimaging Lab at UCLA. In order to find out if you're eligible for the research study, we will need to set up a time to speak on the phone. When we talk, I can tell you more about the study. If you decide you're interested in participating, I'll ask you questions to determine your eligibility. We will need about 10-15 minutes, and it would be good for you to be in a private location where you can speak freely.</p><p>Please reply to this message with your phone number and a few dates and times that you can speak freely on the phone for <strong>10-15 minutes</strong>. We have the most availability during business hours <strong>(9am-5pm Monday through Friday)</strong>.</p><p>Sincerely,<br>Translational Neuroimaging Research Team</p><p style='font-size: .75rem; color: grey;'>--</p><p style='font-size: .75rem; color: grey;'>Translational Neuroimaging Lab</p><p style='font-size: .75rem; color: grey;'>Department of Psychiatry & Biobehavioral Sciences</p><p style='font-size: .75rem; color: grey;'>University of California, Los Angeles</p><p><a href='https://www.translational-neuroimaging.com/' style='font-size: .75rem;'>https://www.translational-neuroimaging.com/</a></p><p><a href='tel:4245323802' style='font-size: .75rem;'>424-532-3802</a></p>`,
				}
			);
			myDraft.getMessage().getThread().addLabel(myLabel);
		} catch (e) {
			Logger.log(`Email address is invalid for ${name}`);
		}
	}
}

// Get All Form Responses
function getFormResponses() {
	const FORM = FormApp.openById(FORMID);
	const allResponses = FORM.getResponses();
	return allResponses;
}



// Restore spreadsheet function
function restoreSpreadsheet(userJobSelectionResponse) {
	const currentSPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
	const backupSPREADSHEETS = DriveApp.getFolderById(BACKUPFOLDERID).getFiles();
	let backupSPREADSHEET;
	while (backupSPREADSHEETS.hasNext()) {
		backupSPREADSHEET = backupSPREADSHEETS.next();
		if (backupSPREADSHEET.getName().includes(userJobSelectionResponse)) {
			Logger.log(
				`The Identified Backup Spreadsheet is ${backupSPREADSHEET.getName()}\nRestoring this version now...`
			);
			const DEST_FOLDER = DriveApp.getRootFolder();
			DriveApp.getFileById(backupSPREADSHEET.getId()).makeCopy(
				`${currentSPREADSHEET.getName()}`,
				DEST_FOLDER
			);
			DriveApp.getFileById(currentSPREADSHEET.getId()).setTrashed(true);
		}
	}
}

// Non user authenticated method for calling in other functions
function cleanupLabelsInsecure() {
	let allLabels = getLabelList();
	for (let i = 0; i < allLabels.length; i++) {
		let label = GmailApp.getUserLabelByName(allLabels[i]);
		if (label.getThreads()[0] == null && isUUID(label.getName())) {
			Logger.log(`${label.getName()} will be deleted`);
			label.deleteLabel();
		} else {
			Logger.log(`${label.getName()} contains active threads`);
		}
	}
}

// Remove labels from threads and delete drafts
function undoProcessThread(thread, uuid) {
	let messageSender = thread[2];
	let messageEmailAddress = messageSender.substring(
		messageSender.indexOf('<') + 1,
		messageSender.indexOf('>')
	);
	Logger.log(messageEmailAddress);
	let currentEmailThread = GmailApp.search(
		`label:${uuid} from:${messageEmailAddress}`
	)[0];
	Logger.log(currentEmailThread);
	currentEmailThread.moveToInbox();
	currentEmailThread.markUnread();
	const messageLength = currentEmailThread.getMessageCount();
	const messages = currentEmailThread.getMessages();
	if (messageLength > 1) {
		const lastMessage = messages[messageLength - 1];
		lastMessage.moveToTrash();
	}
	const threadLabels = currentEmailThread.getLabels();
	Logger.log(`The labels are ${threadLabels}`);
	for (let i = 0; i < threadLabels.length; i++) {
		let finalLabel = threadLabels[i].getName();
		Logger.log(`The final label is ${finalLabel}`);
		currentEmailThread.removeLabel(GmailApp.getUserLabelByName(finalLabel));
	}
}

function tempFunction() {
	let ui = SpreadsheetApp.getUi();
	ui.alert('This function is currently unavailable', ui.ButtonSet.OK);
}

// change these to correct IDs before adapting to actual use
const BACKUPFOLDERID = '1j7AEamTmraI7RMusLABMyWyc18bSaubS';
const PASSWORDHASH = '6d9b298d80f6a2dab59d4879c2e1a8a9';
const FORMID = '1TYpUodL9s4bS0nlNPY5-HH5pNVGt8uAyfbxjmhIma8c';

// 
let startTimestamp = new Date().getTime();
const MAX_RUNTIME = 5.5 * 60;

let masterList, initialOutreach;
let masterEmailsLength, dqEmailsLength;
let masterEmails, dqEmails;

/** Creates an MD5 hash from an input string.
 * ------------------------------------------
 *   MD5 function for GAS(GoogleAppsScript)
 *
 * You can get a MD5 hash value and even a 4digit short Hash value of a string.
 * ------------------------------------------------------------------------------
 * @param {(string|Bytes[])} input The value to hash.
 * @param {boolean} isShortMode Set true for 4 digit shortend hash, else returns usual MD5 hash.
 * @return {string} The hashed input
 * @customfunction
 */
function MD5(input, isShortMode) {
	var isShortMode = !!isShortMode; // Be sure to be bool
	var txtHash = '';
	var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);

	if (!isShortMode) {
		for (i = 0; i < rawHash.length; i++) {
			var hashVal = rawHash[i];
			if (hashVal < 0) {
				hashVal += 256;
			}
			if (hashVal.toString(16).length == 1) {
				txtHash += '0';
			}
			txtHash += hashVal.toString(16);
		}
	} else {
		for (j = 0; j < 16; j += 8) {
			hashVal =
				(rawHash[j] + rawHash[j + 1] + rawHash[j + 2] + rawHash[j + 3]) ^
				(rawHash[j + 4] + rawHash[j + 5] + rawHash[j + 6] + rawHash[j + 7]);

			if (hashVal < 0) {
				hashVal += 1024;
			}
			if (hashVal.toString(36).length == 1) {
				txtHash += '0';
			}
			txtHash += hashVal.toString(36);
		}
	}
	// change below to "txtHash.toUpperCase()" if needed
	return txtHash;
}

// User password authentication on button click
function userAuthentication()
{
  return true;
	let ui = SpreadsheetApp.getUi();
	let result = ui.prompt(
		'Are you sure you want to continue?',
		'Please enter the password:',
		ui.ButtonSet.OK_CANCEL
	);
	let button = result.getSelectedButton();
	let passwordResponse = MD5(result.getResponseText());
	// User clicks cancel
	if (button == ui.Button.CANCEL) {
		ui.alert('Not running script.');
		return false;
		// Incorrect password
	} else if (passwordResponse !== PASSWORDHASH) {
		ui.alert('Incorrect password.\nNot running script.');
		return false;
		// Correct password. Run script
	} else if (button == ui.Button.OK && passwordResponse == PASSWORDHASH) {
		return true;
		// Edge case catcher
	} else {
		ui.alert('ERROR\nUNCAUGHT EDGE CASE');
		return false;
	}
}

function loadEmailLists()
{
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
	masterList = spreadSheet.getSheetByName('Master List');
  initialOutreach = spreadSheet.getSheetByName('Initial Outreach');
	let dqList = spreadSheet.getSheetByName("DQ'd");
	masterEmails = masterList.getRange('C2:C').getValues();
	dqEmails = dqList.getRange('C2:C').getValues();
  for (masterEmailsLength = masterEmails.length - 1; masterEmailsLength >= 0 && masterEmails[masterEmailsLength][0] == '';  masterEmailsLength--);
  masterEmailsLength++;
  for (dqEmailsLength = dqEmails.length - 1; dqEmailsLength >= 0 && dqEmails[dqEmailsLength][0] == ''; dqEmailsLength--);
  dqEmailsLength++;
}

function isInMaster(emailAddress)
{
  for (let i = 0; i < masterEmailsLength; i++)
    if (emailAddress == masterEmails[i])
      return true;
  return false;
}

function isInDQ(emailAddress)
{
  for (let i = 0; i < dqEmailsLength; i++)
    if (emailAddress == dqEmails[i])
      return true;
  return false;
}

function addToMaster(emailAddress)
{
  masterEmails[masterEmailsLength++] = emailAddress;
}

// Backup spreadsheet before doing anything
function backupSpreadsheet(jobID)
{
	const fileName = `${getDate()}-(${jobID})`;
	const DEST_FOLDER = DriveApp.getFolderById(BACKUPFOLDERID);
	const spreadSheetID = SpreadsheetApp.getActiveSpreadsheet().getId();
	DriveApp.getFileById(spreadSheetID).makeCopy(fileName, DEST_FOLDER);
}

// Generarte a uuid for the job
function generateJobID() {
	return 'xxxx-xxxx'.replace(/[xy]/g, function (c) {
		var r = (Math.random() * 16) | 0,
			v = c == 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
	});
}

// Regex phone number search function
function findPhoneNumber(messageBody) {
	const REGEX = /\(?\d{3}\)?-?.? *\d{3}-?.? *-?\d{4}/g;
	const match = REGEX.exec(messageBody);
	if (match) {
		const phoneNumber = match[0];
		return phoneNumber;
	} else {
		return null;
	}
}

// Get current date
function getDate() {
	let date = new Date();
	let month = date.getMonth() + 1;
	let day = date.getDate();
	let year = date.getFullYear();
	let hour = date
		.getHours()
		.toLocaleString('en-US', { minimumIntegerDigits: 2 });
	let minute = date
		.getMinutes()
		.toLocaleString('en-US', { minimumIntegerDigits: 2, useGrouping: false });
	let fullDate = `${month}/${day}/${year}-${hour}:${minute}`;
	return fullDate;
}

// Get current date
function getTime() {
  let date = new Date();
	let month = date.getMonth() + 1;
	let day = date.getDate();
	let year = date.getFullYear();
	let hour = date
		.getHours()
		.toLocaleString('en-US', { minimumIntegerDigits: 2 });
	let minute = date
		.getMinutes()
		.toLocaleString('en-US', { minimumIntegerDigits: 2, useGrouping: false });
  let second = date
		.getSeconds()
		.toLocaleString('en-US', { minimumIntegerDigits: 2, useGrouping: false });
	return `${month}/${day}/${year}-${hour}:${minute}:${second}`;
}

function inTime()
{
  return new Date().getTime() - startTimestamp < MAX_RUNTIME * 1000;
}

function passedTime()
{
  return (new Date().getTime() - startTimestamp) / 1000;
}

// Check if string is in the format of xxxx-xxxx
function isJobID(jobID) {
	const regex = /[0-9a-fA-F]+-[0-9a-fA-F]+/g;
	const match = regex.exec(jobID);
	return match;
}

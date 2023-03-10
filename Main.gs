/** Potential Participant Tracker
 *            __..--''``---....___   _..._    __
 *        _.-'    .-/";  `        ``<._  ``.''_ `. 
 *    _.-' _..--.'_    \                    `( ) ) 
 *   (_..-'    (< _     ;_..__               ; `' 
 *              `-._,_)'      ``--...____..-' 
 * Created by Mike
 * Revised by Jason Liu on 2/21/2023
 * */

// On Open function to create custom menu
function onOpen() {
	createMenuWithSubMenu();
}

// Generate Custom Menu
function createMenuWithSubMenu() {
	SpreadsheetApp.getUi()
		.createMenu('๐ Automation')
		.addItem('Process Emails ๐ง', 'processAllEmails')
		.addSeparator()
    .addItem('Process Forms ๐', 'processForm')
		.addSeparator()
    .addItem('Restore Version ๐ฌ', 'restoreVersion')
    .addSeparator()
		.addItem('Cleanup Labels ๐งน', 'cleanupLabels')
		.addSeparator()
		.addItem('Cleanup Backups ๐งน', 'cleanupBackups')
    .addSeparator()
		.addItem('Cleanup Logs ๐งน', 'cleanupLogs')
    .addSeparator()
		.addItem('Reset Form Response Tracker ๐๏ธ', 'resetLastResponseIndex')
		.addToUi();
}

// so that we can go through the entire form responses
function resetLastResponseIndex()
{
  if (!userAuthentication())
    return;
  PropertiesService.getScriptProperties().setProperty('lastResponseIndex', 0);
  SpreadsheetApp.getActive().toast(`Reset successful`, '๐ Success!');
}

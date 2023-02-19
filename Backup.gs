SHEET_MIMETYPE = 'application/vnd.google-apps.spreadsheet';
FORM_MIMETYPE = 'application/vnd.google-apps.form';

function undo()
{

}

function getLatestBackups(num)
{
  let folder = DriveApp.getFolderById(BACKUPFOLDERID);
  let files = folder.getFiles();
  let allFiles = [];
  while (files.hasNext()) {
    file = files.next();
    if (!isValidBackupName(file.getName()))
    {
      continue;
    }
    let type = '';
    if (file.getMimeType() == SHEET_MIMETYPE)
      type = 'Sheet';
    else if (file.getMimeType() == FORM_MIMETYPE)
      type = 'Form';
    else
    {
      continue;
    }
    allFiles.push([file.getName(), file.getId(), type]);
  }

  allFiles.sort(sortByFirst);
  let latestFiles = [];
  for (let i = 0; i < allFiles.length; i++)
  {
    if (allFiles[i][2] == 'Sheet')
    {
      if (i + 1 < allFiles.length && allFiles[i + 1][2] == 'Form' && allFiles[i + 1][0] == allFiles[i][0])
      {
        latestFiles.push([allFiles[i][0], allFiles[i][1], allFiles[i + 1][0], allFiles[i + 1][1]]);
        i++;
      }
      else
      {
        latestFiles.push([allFiles[i][0], allFiles[i][1], null, null]);
      }
    }
  }
  
  for (let i = latestFiles.length; i < num; i++)
    latestFiles.push([null, null, null, null]);
  
  return latestFiles.slice(0, num);
}

function isValidBackupName(name)
{
  if (name.length < 12 || name[name.length - 12] != '-')
    return false;
  let jobID = name.substring(name.length - 10, name.length - 1);
  if (!isJobID(jobID))
    return false;
  if (isNaN(new Date(name.substring(0, name.length - 12)).getTime()))
    return false;
  return true;
}

function sortByFirst(a, b) {
  let aTime = new Date(a[0].substring(0, a[0].length - 12)).getTime();
  let bTime = new Date(b[0].substring(0, b[0].length - 12)).getTime();
  if (aTime === bTime) {
    if (a[2] === b[2])
      return 0;
    else if(a[2] === 'Sheet')
      return -1;
    else
      return 1;
  } else {
    return (aTime > bTime) ? -1 : 1;
  }
}


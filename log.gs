function rnp(msg)
{
  Logger.log(msg)
}

function logScriptStart()
{
  rnp(`Script execution starts at ${getTime()}`);
}

function log(str)
{
  rnp(`[${getTime()}]${str}`);
}

function logScriptEnd()
{
  rnp(`Script execution ends, execution time: ${passedTime()}`);
}

function warning(str)
{
  rnp(`Warning: [${getTime()}]${str}`)
}

function error(str)
{
  rnp(`ERROR: [${getTime()}]${str}`)
}

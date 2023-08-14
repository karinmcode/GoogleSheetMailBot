function cleanBody(t) {
  //var t = "Hi Nevena,\r\n\r\nI will see with Liam and Claudia if we can meet on Sunday and come back to\r\nyou. >>We are quiet busy during the week also.\r\n\r\nCheers and welcome to Geneva,\r\n\r\nKarin\r\n\ " ;
  var a = "";
  var tout = "";
  try{
  for (i=0;i<t.length;i++){
    
    if (t[i] == ">"){
      tout = a;
      Logger.log(tout)
      return tout

    }
    a = [a + t[i]];

  }
  }
  catch(e){};
  tout =a;
  return tout

  
}


// Basic website validation
function isValidWebsite(url) {
  // Check if it contains at least one dot and no spaces (you can expand this check as needed)
  return url.includes('.') && !url.includes(' ');
}

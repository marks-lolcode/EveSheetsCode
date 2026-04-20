 /**
 // Converts a datetime string to a datetime string in a targe timezone.
 //
 //param {"September 1, 2018 1:00 PM PST"} datetimeString Date, time and timezone.
 //param {"GMT"} timeZone Target timezone
 //param {"YYYY-MM-dd hh:mm a z"} Datetime format
 //@customfunction
 //
*/

function convertTimeZone(datetimeString,timeZone="",format="YYYY-MM-dd hh:mm") {
  var moment = new Date(datetimeString);
  return Utilities.formatDate(moment, timeZone, format);
}

 /**
 //Time Zone abbreviations: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
 //America/Chicago
 //UTC
 */


/** 

 * @customfunction 

 * Your function description here 

 * 

 * @param {"September 1, 2018 1:00 PM PST"} datetime Date, time and timezone.

 * @param {"CST"} fromTZ From timezone
 * 
 * @param {"EST"} toTZ Target timezone

 * @return {"September 1, 2018 2:00 PM"} 

 */
function convertTZ(datetime="11/18/2024 23:13:06", fromTZ="GMT", toTZ="America/Chicago") { 
  var moment = new Date(datetime);
  console.log(moment);
  var fromOffset = Utilities.formatDate(moment, fromTZ, "Z");
  console.log(fromOffset);
  var toOffset = Utilities.formatDate(moment, toTZ, "Z");
  console.log(toOffset);
  var offset = (toOffset - fromOffset) / 100;  //60
  console.log(offset);
  var offsetSeconds=(offset * 3600000);
  console.log(offsetSeconds);
  var convertedDate= new Date(moment.getTime()+offsetSeconds);
  console.log(moment.getTime());
  console.log(convertedDate);
  return convertedDate;
}

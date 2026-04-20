/**
 * Converts values like "-1,228,000,000 ISK" to numbers across a range.
 *
 * Usage in Sheets:
 *   =ISK_RANGE_TO_NUMBERS(L24:L29)                 // keeps negatives (default)
 *   =ISK_RANGE_TO_NUMBERS(L24:L29, TRUE)           // removes negative sign (makes all results >= 0)
 *
 * Parameters:
 * - rangeInput: a cell or range of cells
 * - removeNegative (optional): TRUE to convert negative values to positive. Default FALSE.
 *
 * Returns a 2D array so it "spills" properly in Sheets.
 * - Blank input -> blank output
 * - If a cell cannot be parsed -> blank output (you can change this behavior)
 */
/**
 * Converts ISK strings like "-1,228,000,000 ISK" into numbers (supports ranges too).
 *
 * Usage:
 *   =ISK_RANGE_TO_NUMBERS(L24:L29)
 *   =ISK_RANGE_TO_NUMBERS(L24:L29, TRUE)
 *
 * @param {any[][]|any} rangeInput A cell or range to convert.
 * @param {boolean} [removeNegative] TRUE to remove negative sign (absolute value).
 * @return {any[][]} A 2D array of numbers (or blanks).
 * @customfunction
 */
function ISK_RANGE_TO_NUMBERS(rangeInput, removeNegative) {
  try {
    // Default behavior: keep negative signs.
    var shouldRemoveNegative = _toBooleanWithDefault_(removeNegative, false);

    console.log("ISK_RANGE_TO_NUMBERS called.");
    console.log("shouldRemoveNegative:", shouldRemoveNegative);

    // Custom functions receive:
    // - a 2D array when passed a range
    // - a single value when passed a single cell
    var values = rangeInput;

    // If a single cell was passed, wrap it into a 2D array to keep output consistent.
    if (!Array.isArray(values)) {
      values = [[values]];
    }

    // If it's a 1D array for some reason, normalize it to 2D.
    if (Array.isArray(values) && values.length > 0 && !Array.isArray(values[0])) {
      values = values.map(function (v) { return [v]; });
    }

    console.log("Normalized input size:", values.length, "rows");

    // Convert each cell in the 2D array
    var output = values.map(function (row, rIdx) {
      return row.map(function (cell, cIdx) {
        var parsed = _iskParseToNumberOrBlank_(cell, shouldRemoveNegative);
        console.log("Parsed cell", rIdx, cIdx, "from", cell, "to", parsed);
        return parsed;
      });
    });

    return output;
  } catch (err) {
    console.error("ISK_RANGE_TO_NUMBERS error:", err);
    return [["ERROR: " + err.message]];
  }
}

/**
 * Helper: parse a single cell value into a number or blank.
 * Keeps negative sign unless shouldRemoveNegative is TRUE.
 */
function _iskParseToNumberOrBlank_(input, shouldRemoveNegative) {
  // Keep blanks blank.
  if (input === null || input === undefined || input === "") {
    return "";
  }

  // If it's already a number, apply the negative-handling rule and return.
  if (typeof input === "number") {
    return shouldRemoveNegative ? Math.abs(input) : input;
  }

  // Normalize to string.
  var text = String(input).trim();

  // Remove the ISK label, commas, and whitespace.
  text = text
    .replace(/isk/gi, "")       // remove ISK (any case)
    .replace(/,/g, "")          // remove thousands separators
    .replace(/\u00A0/g, " ")    // convert NBSP to normal space
    .replace(/\s+/g, "")        // remove any remaining whitespace
    .trim();

  // Pull the first number-like token (supports optional leading - and optional decimals).
  var match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) {
    // If you prefer an error marker instead of blank, replace "" with something like "PARSE_ERROR".
    return "";
  }

  var value = Number(match[0]);
  if (Number.isNaN(value)) {
    return "";
  }

  // If requested, remove the negative sign by taking absolute value.
  return shouldRemoveNegative ? Math.abs(value) : value;
}

/**
 * Converts common Sheets inputs to boolean with a default.
 * In Sheets, user might pass TRUE/FALSE, 1/0, "TRUE"/"FALSE", "yes"/"no".
 */
function _toBooleanWithDefault_(value, defaultValue) {
  if (value === null || value === undefined || value === "") {
    return defaultValue;
  }

  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  var text = String(value).trim().toLowerCase();
  if (text === "true" || text === "t" || text === "yes" || text === "y" || text === "1") {
    return true;
  }
  if (text === "false" || text === "f" || text === "no" || text === "n" || text === "0") {
    return false;
  }

  // If it's something unexpected, fall back to the default to avoid surprises.
  return defaultValue;
}





/**
 * Old Spreadsheet Field Code: 
 * =if(A21=FALSE,"",if(int(G21-now())<0,"PRODUCTION COMPLETE",(if(int(G21-now())>0,if(int(G21-now())>1,int(G21-now())&" Days ",int(G21-now())&" Day "),""))&if(hour(abs((G21-now())))>1,hour(abs((G21-now())))&" Hours ",int(G21-now())&" Hour ")&if(MINUTE(abs(G21-now()))>1,MINUTE(abs(G21-now()))&" Mins ",MINUTE(abs(G21-now()))&" Min ")))
 * 
 */


function testDCD() {
  var testD = dateCountDown("11/28/22 17:57",false);
  console.log (testD);
}

function dateCountDown(countDownDateIn,refreshToggle=true) {
console.log(countDownDateIn);
// Set the date we're counting down to
//var endDate = startTime+duration;
var countDownDate = new Date(countDownDateIn);
console.log(countDownDate);

// Get today's date and time
var now = new Date().getTime();
console.log(now);

// Find the distance between now and the count down date
var distance = countDownDate - now;
console.log(distance);

//Initialize the values and make sure they are zero.
var days = 0;
var hours = 0;
var minutes = 0;
var seconds = 0;

// Time calculations for days, hours, minutes and seconds
days = Math.floor(distance / (1000 * 60 * 60 * 24));
hours = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
seconds = Math.floor((distance % (1000 * 60)) / 1000);

var displayFormat = "";

//Check to see if no time is left and return complete.
if ((days <= 0) && (hours <= 0) && (minutes <= 0)) {
  displayFormat = "COMPLETE";
  console.log(displayFormat);``
  return displayFormat;
}

//Check each value to see if it should be included and append to return value if should.
if (days > 0){
  if (days <10) {
    displayFormat = displayFormat + "0" + days + "d ";
  } else {
    displayFormat = displayFormat + days + "d ";
  }
}
console.log ("Days: " + days);
console.log (displayFormat);

if ((days > 0) || (hours > 0)) {
  if (hours <10) {
    displayFormat = displayFormat + "0" + hours + "h ";
  } else {
    displayFormat = displayFormat + hours + "h ";
  }
}

console.log ("Hours: " + hours);
console.log (displayFormat);

if ((days > 0 || hours > 0) || (minutes > 0)) {
  if (minutes <10) {
    displayFormat = displayFormat + "0" + minutes + "m ";
  } else {
    displayFormat = displayFormat + minutes + "m ";
  }
}

console.log ("Minutes: " + minutes);
console.log (displayFormat);

return displayFormat;
}

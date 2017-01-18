var _LOCAL_PEOPLE_RESULT_ID = 'B09A7990-05EA-4AF9-81EF-EDFAB16C4E31';
var arrStaff = [];
var context = new SP.ClientContext.get_current();
var web = context.get_web();
var user = web.get_currentUser();

$(document).ready(function() {
  // If the window has been resized
  $(window).resize(function() {
  });

  context.load(user);
  context.executeQueryAsync(function() {
    console.log(user);
  }, function() {
    alert("Connection Failed. Refresh the page and try again. :(");
  });
});

function escapeURL(string) {
  return String(string).replace(/[&<>\-'\/]/g, function(s) {
    var entityMap = {
      "&": "%26",
      "<": "%3C",
      ">": "%3E",
      "-": "\\-",
      "'": '%60',
      "\\": "%5C",
      "[": "%5B",
      "]": "%5D"
    };

    return entityMap[s];
  });
}

/**
  Reloads the phonebook page with new data
**/
function refreshPhonebook() {
  $("article.phonebook").empty();

  if (arrStaff.length == 0)
    return;

  for (var i = 0; i < arrStaff.length; i++) {
    var entry = $("<div/>", {
      html: '<span>' + arrStaff[i].Name + '</span><span>' + arrStaff[i].CellPhone + '</span>'});

    $("article.phonebook").append(entry);
  }
}

function enableFields() {
  $("#itin-main button, #itin-main input").removeClass("disabled");
  $("#itin-splash button").removeClass("disabled");
}

function onQueryFailed(sender, args) {
  console.log('Request failed.' + args.get_message() +
    '\n' + args.get_stackTrace());
}

function onQueryFailedJSON(data, errorCode, errorMessage) {
  if (data.responseJSON != null)
    console.log(data.responseJSON.error.code + ': ' + data.responseJSON.error.message.value);
  else
    console.log(data);
}

// Get phone numbers of all staff
function getStaffPhones() {
  searchTerm = '(JobTitle:"a*" JobTitle:"b*" JobTitle:"c*" JobTitle:"d*" JobTitle:"e*"' +
    'JobTitle:"f*" JobTitle:"g*" JobTitle:"h*" JobTitle:"i*" JobTitle:"j*"' +
    'JobTitle:"k*" JobTitle:"l*" JobTitle:"m*" JobTitle:"n*" JobTitle:"o*"' +
    'JobTitle:"p*" JobTitle:"q*" JobTitle:"r*" JobTitle:"s*" JobTitle:"t*"' +
    'JobTitle:"u*" JobTitle:"v*" JobTitle:"w*" JobTitle:"x*" JobTitle:"y*"' +
    'JobTitle:"z*" JobTitle:"0*" JobTitle:"1*" JobTitle:"2*" JobTitle:"3*"' +
    'JobTitle:"4*" JobTitle:"5*" JobTitle:"6*" JobTitle:"7*" JobTitle:"8*"' +
    'JobTitle:"9*")';
  searchUrl = _spPageContextInfo.webAbsoluteUrl +
    "/_api/search/query?querytext='" + searchTerm +
    "'&sortlist='PreferredName:ascending'&" +
    "selectproperties='PreferredName,WorkPhone,CellPhone,HomePhone'&" +
    "sourceid='" + _LOCAL_PEOPLE_RESULT_ID + "'&" +
    "rowlimit='100'";

  //console.log(searchUrl);

  $.ajax({
    url: searchUrl,
    type: "GET",
    headers: {
      "Accept": "application/json; odata=verbose"
    },
    success: onStaffPhoneJSON,
    error: onQueryFailedJSON
  });
}

// Writes Staff phone numbers into arrStaff and reloads the phonebook
function onStaffPhoneJSON(data) {
  arrStaff = [];

  var results = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
  if (results.length > 0) {
    $.each(results, function(index, result) {
      var item = {
        'Name': result.Cells.results[2].Value,
        'WorkPhone': (result.Cells.results[3].Value != null ? result.Cells.results[3].Value : ''),
        'CellPhone': (result.Cells.results[4].Value != null ? result.Cells.results[4].Value : ''),
        'HomePhone': (result.Cells.results[5].Value != null ? result.Cells.results[5].Value : '')
      };

      arrStaff.push(item);
    });

    refreshPhonebook();

  // No results were found.
  } else {
    $("article.phonebook").html("Can't find any staff in the directory.");
  }

}

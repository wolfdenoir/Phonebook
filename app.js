var _LOCAL_PEOPLE_RESULT_ID = 'B09A7990-05EA-4AF9-81EF-EDFAB16C4E31';
var arrStaff = [];
var prop;
var context = new SP.ClientContext.get_current();
var web = context.get_web();
var user = web.get_currentUser();

$(document).ready(function() {
  // If the window has been resized
  $(window).resize(function() {});

  context.load(user);
  context.executeQueryAsync(function() {
    console.log(user);
    getStaffPhones();
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
    if (arrStaff[i].CellPhone != '') {
      var cellphones = $("<div/>", {
        html: '<p>' + arrStaff[i].Name + '</p><p>' + arrStaff[i].CellPhone + '</p>'
      });
      $("#cellphone article.phonebook").append(cellphones);
    }

    if (arrStaff[i].HomePhone != '') {
      var homephones = $("<div/>", {
        html: '<p>' + arrStaff[i].Name + '</p><p>' + arrStaff[i].HomePhone + '</p>'
      });
      $("#homephone article.phonebook").append(homephones);
    }

    if (arrStaff[i].WorkPhone != '') {
      var workphones = $("<div/>", {
        html: '<p>' + arrStaff[i].Name + '</p><p>' + arrStaff[i].WorkPhone + '</p>'
      });
      $("#workphone article.phonebook").append(workphones);
    }
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
  searchTerm = '(JobTitle:"Support" JobTitle:"Servicing")';
  searchUrl = _spPageContextInfo.webAbsoluteUrl +
    "/_api/search/query?querytext='" + searchTerm +
    "'&sortlist='PreferredName:ascending'&" +
    "selectproperties='PreferredName,AccountName,WorkPhone'&" +
    "sourceid='" + _LOCAL_PEOPLE_RESULT_ID + "'&" +
    "rowlimit='400'";

  console.log(searchUrl);

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
  var numResults = results.length;
  if (numResults > 0) {
    numReady = 0;

    // Prepare the Progressbar for update
    var progressBar = $("#progress-phonelist");
    if ($("div.progress").hasClass("hidden")) {
      $("div.progress").removeClass("hidden");
    }

    $.each(results, function(index, result) {
      var item = {
        'Name': result.Cells.results[2].Value,
        'WorkPhone': (result.Cells.results[4].Value != null ? result.Cells.results[4].Value : ''),
        'CellPhone': '',
        'HomePhone': ''
      };

      arrStaff.push(item);

      var searchUrl = _spPageContextInfo.webAbsoluteUrl +
        "/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)?@v='" +
        String(result.Cells.results[3].Value).replace(/'/g, "''") + "'";
      $.ajax({
        url: searchUrl,
        type: "GET",
        headers: {
          "Accept": "application/json; odata=verbose"
        },
        indexValue: index,
        total: numResults,
        progressbar: progressBar,
        success: function(data) {
          arrStaff[index].CellPhone = (data.d.UserProfileProperties.results[48].Value != null ? data.d.UserProfileProperties.results[48].Value : '');
          arrStaff[index].HomePhone = (data.d.UserProfileProperties.results[50].Value != null ? data.d.UserProfileProperties.results[50].Value : '');
          progressBar.html("Getting phone numbers for " + numReady++ + " out of " + numResults + " Users.");
          progressBar.attr("aria-valuenow", Math.ceil(numReady / numResults * 100) + "");
          progressBar.attr("style", "width: " + Math.ceil(numReady / numResults * 100) + "%");
          if (arrStaff.length == numReady) {
            if (!$("div.progress").hasClass("hidden"))
              $("div.progress").addClass("hidden");
            refreshPhonebook();
          }
        },
        error: onQueryFailedJSON
      });
    });

    //console.log("arrStaff: " + arrStaff.length);

    // No results were found.
  } else {
    $("article.phonebook").html("Can't find any staff in the directory.");
  }

}

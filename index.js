// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. see LICENSE in the project root for license information.
"use strict";

Office.initialize = function(reason) {
  $(document).ready(function(){
    loadEntities();
  });
}

function loadEntities() {
  // getSelectedRegExMatches is in preview, so need to test for it
  if (Office.context.mailbox.item.getSelectedRegExMatches !== undefined) {
    var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
    if (selectedMatches) {
      var url = 'https://centralusdtapp52.epicorsaas.com/SaaS5127/Apps/Erp/Home/#/view/OMGO3004/Erp.UI.PartEntry?channelid=0&layerVersion=0&baseAppVersion=0&company=161830&site=MfgSys&pageId=Details&pageChanged=true&KeyFields.PartNum=' + selectedMatches.OrderNumber;
      $("#selected-match").html('<a href="' + url + '" target="_blank">Click here to open Part Tracker</a>');
    } else {
      $("#selected-match").text("Selected matches was null");
    }
  } else {
    $("#selected-match").text("Method not supported on your client");
  }
  // Get all matches
  var allMatches = Office.context.mailbox.item.getRegExMatches();
  if (allMatches) {
    var url = 'https://centralusdtapp52.epicorsaas.com/SaaS5127/Apps/Erp/Home/#/view/OMGO3004/Erp.UI.PartEntry?channelid=0&layerVersion=0&baseAppVersion=0&company=161830&site=MfgSys&pageId=Details&pageChanged=true&KeyFields.PartNum=' + allMatches.OrderNumber;
    $("#all-matches").html('<a href="' + url + '" target="_blank">Click here to open Part Tracker</a>');
  } else {
    $("#all-matches").text("All matches was null");
  }
}

// function loadEntities() {
//   // getSelectedRegExMatches is in preview, so need to test for it
//   if (Office.context.mailbox.item.getSelectedRegExMatches !== undefined) {
//     var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
//     if (selectedMatches) {
//       // Note that the use of selectedMatches.PartNumber, where
//       // PartNumber corresponds to the RegExName attribute of the Rule element
//       // in the manifest
//       $("#selected-match").text(JSON.stringify(selectedMatches.PartNumber, null, 2));
//     } else {
//       $("#selected-match").text("Selected matches was null");
//     }
//   } else {
//     $("#selected-match").text("Method not supported on your client");
//   }

//   // Get all matches
//   var allMatches = Office.context.mailbox.item.getRegExMatches();
//   if (allMatches) {
//     // Note that the use of selectedMatches.PartNumber, where
//     // PartNumber corresponds to the RegExName attribute of the Rule element
//     // in the manifest
//     $("#all-matches").text(JSON.stringify(allMatches.PartNumber, null, 2));
//   } else {
//     $("#all-matches").text("All matches was null");
//   }
// }

function showError(message) {
  $("#error-msg").text(message);
  $("#error").show();
}
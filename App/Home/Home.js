/// <reference path="../App.js" />
/*global app*/

(function () {
	'use strict';

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			$('#get-data-from-selection').click(getDataFromSelection);
		});
	};

	// Reads data from current document selection and displays a notification
	function getDataFromSelection() {
		Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					app.showNotification('The selected text is:', '"' + result.value + '"');
				} else {
					app.showNotification('Error:', result.error.message);
				}
			}
		);
	}
	
})();

function createFunction()
{
  var inputReminder = document.getElementById("input");
  if(inputReminder.value !== "")
  {
  var table = document.getElementById("myTable");
  var row = table.insertRow(0);
  
  var cell1 = row.insertCell(0);
  var element1 = document.createElement("input");
  element1.type = "checkbox";
  cell1.appendChild(element1);
  
  var cell2 = row.insertCell(1);
  cell2.innerHTML = 
  '<textarea maxlength="5000" cols="40" rows="40" style="width: 300px; height: 70px; max-width:300px;"></textarea>'
  var element2 = cell2.getElementsByTagName('textarea')[0];
  element2.value = inputReminder.value;
  //var element2 = document.createElement("input");
  //element2.type = "textarea";
  //element2.style.height = 'auto'; element2.style.width = "300px";
  //element2.style.height = inputReminder.value.scrollHeight + 'px';
  //element2.wrap = "virtual";
  //element2.value = inputReminder.value; 
  cell2.appendChild(element2);
  }
}

function deleteFunction() 
{
   try 
   {
   var table = document.getElementById("myTable");
   var rowCount = table.rows.length;
 
   for(var i=0; i<rowCount; i++) 
   {
     var row = table.rows[i];
     var chkbox = row.cells[0].childNodes[0];
     if(null !== chkbox && true === chkbox.checked)
     {
         table.deleteRow(i);
         rowCount--;
         i--;
     }
   }
   }
   
   catch(e)
    {
       alert(e);
    }
}

//function deleteRow()
//{
//	var table = document.getElementById("myTable");	 
//}
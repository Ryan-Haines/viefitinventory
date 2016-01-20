//Ryan Haines January 2016
//This script uses two external libraries: JS-XLSX for parsing an excel file and jquery-csv for parsing a CSV file
$(document).ready(function() {
  
  //variables to track items sold, master inventory, barcodes extracted from inventory
  var soldItems = [];
  var inventory = [];
  var barcodes =[];
  
  //variables to track what column in inventory corresponds to listed trait
  var invUPCCol; //"Variant Barcode" column
  var invSzCol; //"Option2 Value" (sizes)
  var invQtyCol; //"Variant Inventory Qty" column
  var invColorCol; //"Option1 Value" (color) column
  var invBodyCol; //"Body (HTML)" column

  var invName = '';

  var debugLog = '';

  if(isAPIAvailable()) {
    $("#update").click(updateInventory);
    $("#inv").bind('change', handleInv)
    $("#txns").bind('change', handleTxn);
    $("#debug").click(writeDebug);
  }

  function isAPIAvailable() {
    // Check for the various File API support.
    if (window.File && window.FileReader && window.FileList && window.Blob) {
      // Great success! All the File APIs are supported.
      return true;
    } else {
      // source: File API availability - http://caniuse.com/#feat=fileapi
      // source: <output> availability - http://html5doctor.com/the-output-element/
      document.writeln('The HTML5 APIs used in this form are only available in the following browsers:<br />');
      // 6.0 File API & 13.0 <output>
      document.writeln(' - Google Chrome: 13.0 or later<br />');
      // 3.6 File API & 6.0 <output>
      document.writeln(' - Mozilla Firefox: 6.0 or later<br />');
      // 10.0 File API & 10.0 <output>
      document.writeln(' - Internet Explorer: Not supported (partial support expected in 10.0)<br />');
      // ? File API & 5.1 <output>
      document.writeln(' - Safari: Not supported<br />');
      // ? File API & 9.2 <output>
      document.writeln(' - Opera: Not supported');
      return false;
    }
  }

  function handleTxn(e) {
    var files = e.target.files;
    var file = files[0];
    console.log(file);
    var reader = new FileReader();
    var name = file.name;
    reader.onload = function(e) {
      var data = e.target.result;

      //workbook is our XLSX file converted to a javascript friendly object
      var workbook = XLSX.read(data, {type: 'binary'});
      var nullCell = false;
      var index = 2;

      var prodCol = "B";
      var colorCol = "C";
      var szCol = "D"
      var barcodeCol = "E";
      var qtyCol = "F";
      
      var sheetName = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[sheetName];
      
      while(!nullCell){
        var prodAddr = prodCol + index;
        var colorAddr = colorCol + index;
        var szAddr = szCol + index;
        var barcodeAddr = barcodeCol + index;
        var qtyAddr = qtyCol + index;

        var prodCell = worksheet[prodAddr];
        var colorCell = worksheet[colorAddr];
        var szCell = worksheet[szAddr];
        var barcodeCell = worksheet[barcodeAddr];
        var qtyCell = worksheet[qtyAddr];
      
        if(barcodeCell == null || barcodeCell == undefined)
          nullCell = true;
        else{
          var txnBarcode = barcodeCell.v;
          //trim all leading 0s
          while(txnBarcode.charAt(0) == '0'){
            txnBarcode = txnBarcode.substring(1, txnBarcode.length);
          }
          var soldItem = [prodCell.v, colorCell.v, szCell.v, txnBarcode, qtyCell.v];
          soldItems.push(soldItem);
          index++;
        }
      }
      console.log(soldItems);
    };
    reader.readAsBinaryString(file);
  }

  function handleInv(evt) {
    var invFiles = $('#inv').prop('files');;
    var invFile = invFiles[0];
    console.log(invFile);
    debugLog += invFile + '\n';

    // read the file metadata
    var invOutput = ''
    invOutput += '<span style="font-weight:bold;">' + escape(invFile.name) + '</span><br />\n';
    invOutput += ' - FileType: ' + (invFile.type || 'n/a') + '<br />\n';
    invOutput += ' - FileSize: ' + invFile.size + ' bytes<br />\n';
    invOutput += ' - LastModified: ' + (invFile.lastModifiedDate ? invFile.lastModifiedDate.toLocaleDateString() : 'n/a') + '<br />\n';

    invName = invFile.name;

    //print the file contents
    printInventoryTable(invFile);
    //handleTxnFile(txnFile);

    // post the results
    //$('#invHolder').append(output);
  }

  function printInventoryTable(file) {
    var reader = new FileReader();
    reader.readAsText(file);
    reader.onload = function(event){
      var csv = event.target.result;
      var data = $.csv.toArrays(csv);
      var html = '';
      /*
      //table formatted output for debugging
      for(var row in data) {
        html += '<tr>\r\n';
        for(var item in data[row]) {
          html += '<td>' + data[row][item] + '</td>\r\n';
        }
        html += '</tr>\r\n';
      }
      */

      //maintain formatting
      for(var row in data) {
        var thisRow = [];
        var thisCell;
        for(var item in data[row]) {
          thisCell = data[row][item];
          //trim leading ' from UPC column
          if(item == invUPCCol && thisCell.charAt(0) ==='\''){
            thisCell = thisCell.substring(1, thisCell.length);
          }
          //manually add quotations to body column, otherwise formatting is disturbed
          if(item == invBodyCol){
            thisCell = "\"" + thisCell + "\"";
          }
          thisRow.push(thisCell); //builds into array
          if(row == 0){
            if(thisCell === "Variant Barcode"){
              invUPCCol = item;
              console.log("UPC are in column " + invUPCCol);
            }
            else if(thisCell === "Option2 Value"){
              invSzCol = item;
              console.log("Sizes are in column " + invSzCol);
            }
            else if(thisCell === "Variant Inventory Qty"){
              invQtyCol = item;
              console.log("Quantity is in column " + invQtyCol);
            }
            else if(thisCell === "Option1 Value"){
              invColorCol = item;
              console.log("Color is in column " + invColorCol);
            }
            else if(thisCell === "Body (HTML)"){
              invBodyCol = item;
              console.log("Body is in column " + invBodyCol);
            }
          }
          thisCell = "";
        }
        inventory.push(thisRow);
        thisRow = []; //clear this row
      }

      console.log(inventory);
    };
    reader.onerror = function(){ alert('Unable to read ' + file.fileName); };
  }


  //for each entry in soldItems, searches for a corresponding item in the inventory and reduces its quantity by the given amount
  function updateInventory(){
    console.log("updating inventory...");
    debugLog += "updating inventory... "+ '\n';
    for(var i = 0; i < soldItems.length; i++){
      var indexOfMatchingBarcodes = [];
      var matchFound = false;
      var sellingItem = soldItems[i];
      var sellingBarcode = sellingItem[3];
      var sellingQTY = sellingItem[4];
      console.log("removing item " + sellingBarcode + " in the amount of "+ sellingQTY);
      debugLog += "removing item " + sellingBarcode + " in the amount of "+ sellingQTY + '\n';
      
      var invItem;
      for(var j = 0; j < inventory.length; j++){
        invItem = inventory[j]
        if(invItem[invUPCCol] === sellingBarcode){ //barcode match
          matchFound = true;
          indexOfMatchingBarcodes.push(j);
        }
      }
      //check if there was more than one match, if so, also match on size & color fields
      if(matchFound){
        var barcodeMatchItem;
        if(indexOfMatchingBarcodes.length == 1){
          barcodeMatchItem = inventory[indexOfMatchingBarcodes[0]];
          console.log("Found match! Item " + barcodeMatchItem + " has barcode " + sellingBarcode);
          debugLog += "Found match! Item " + barcodeMatchItem + " has barcode " + sellingBarcode + '\n';
          barcodeMatchItem[invQtyCol]-=sellingQTY; //reduce inventory by QTY of sold item 
          console.log("Item now has quantity: " + barcodeMatchItem[invQtyCol]);
          debugLog += "Item now has quantity: " + barcodeMatchItem[invQtyCol] + '\n';
        }
        else{
          //check that the barcodeMatchItem matches other fields of sellingItem (we already know barcode is a match)
          for(var k = 0; k < indexOfMatchingBarcodes.length; k++){
            console.log("Found " + indexOfMatchingBarcodes.length + " matching barcodes, narrowing search");
            debugLog += "Found " + indexOfMatchingBarcodes.length + " matching barcodes, narrowing search" + '\n';
            barcodeMatchItem = inventory[indexOfMatchingBarcodes[k]];
            if(barcodeMatchItem[invSzCol] === sellingItem[2] && barcodeMatchItem[invColorCol] === sellingItem[1]){
              console.log("Found match! Item " + barcodeMatchItem + " has barcode " + sellingBarcode + "and matches other fields");
              debugLog += "Found match! Item " + barcodeMatchItem + " has barcode " + sellingBarcode + "and matches other fields" + '\n';
              barcodeMatchItem[invQtyCol]-=sellingQTY; //reduce inventory by QTY of sold item 
              console.log("Item now has quantity: " + barcodeMatchItem[invQtyCol]);
              debugLog += "Item now has quantity: " + barcodeMatchItem[invQtyCol] + '\n';
            }
          }
        }
      }
      else{
        console.log("Couldn't find match for " + soldItems[i][3]);
        debugLog += "Couldn't find match for " + soldItems[i][3] + '\n';
      }
      matchFound = false;
      indexOfMatchingBarcodes = [];
    }
    console.log(inventory);
    console.log("inventory successfully updated");
    debugLog += "inventory successfully updated" + '\n';
    writeInventory();
  }

  //Writes our inventory to file
  function writeInventory(){
    var inventoryText = '';
    for(var i = 0; i < inventory.length-1; i++){
      inventoryText += inventory[i] + '\n';
    }
    inventoryText += inventory[inventory.length-1];
    var blob = new Blob([inventoryText], {type: "text/plain;charset=utf-8"});
    console.log(inventoryText);
    saveAs(blob, changeFilename(invName, "_SOLD_"));

  }

  function writeDebug(){
    var blob = new Blob([debugLog], {type: "text/plain;charset=utf-8"});
    saveAs(blob, changeFilename("debug_log.txt", '_'));
  }

  function changeFilename(str, custom){
    var name = ''
    var extension = ''
    var d = new Date();
    var currentDate = d.getMonth() + "-" + d.getDate() + "-" + d.getFullYear() + "T" + d.getHours() + ":" + d.getMinutes() + ":" + d.getSeconds();
    for(var i = 0; i<str.length; i++){
      if(str.charAt(i) == '.'){
        var name = str.substring(0, i) + custom +currentDate + str.substring(i, str.length);
        break;
      }
    }
    return name;
  }

  function download(filename, text) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);

    element.style.display = 'none';
    document.body.appendChild(element);

    element.click();

    document.body.removeChild(element);
  }

});
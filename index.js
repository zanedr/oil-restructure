// XLSX = require('xlsx')
const select = document.getElementById("myFile");
const mainTable = document.getElementById("main-table");

const wellNameAndNumberArray = []
const APIArray = []
const topPerfArray = []
const bottomPerfArray = []
const stageNoArray = []
const ISIPStartArray = []
const threeMinArray = []
const ISIPEndArray = []

select.addEventListener('change', handleFile, false);

function fixdata(data) {
  var o = "", l = 0, w = 10240;
  for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
  o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
  return o;
}

const fileNames = []
let counter = 0

var rABS = true;

function handleFile(e) {
  var files = e.target.files;
  var i,f;
  for (i = 0; i != files.length; ++i) {
    f = files[i];
    var reader = new FileReader();
    var name = f.name;
    fileNames.push(name)
    reader.onload = function(e) {
      var data = e.target.result;

      var workbook;
      if(rABS) {
        /* if binary string, read with type 'binary' */
        workbook = XLSX.read(data, {type: 'binary'});
      } else {
        /* if array buffer, convert to base64 */
        let arr = fixdata(data);
        workbook = XLSX.read(btoa(arr), {type: 'base64'});
      }
        let fracReport = workbook.SheetNames[3];
        let breakdown = workbook.SheetNames[6];
        let summary = workbook.SheetNames[7];

        // let wellNameAndNumber = 'C6';
      /* DO SOMETHING WITH workbook HERE */
        console.log(' ');
        console.log('FILE:', fileNames[counter]);
        let tableRow = document.createElement('tr')
        // file name
        let fileNameCell = document.createElement('th')
        let fileNameVal = document.createTextNode(`${fileNames[counter]}`)
        counter++

        fileNameCell.appendChild(fileNameVal)
        tableRow.appendChild(fileNameCell)

        // wellNameAndNumber
        let nameNumbCell = document.createElement('th')
        let nameNumVal = document.createTextNode(`${getValue(workbook, breakdown, 'C1', 'wellNameAndNumber')}`)

        nameNumbCell.appendChild(nameNumVal)
        tableRow.appendChild(nameNumbCell)

        //API
        let apiCell = document.createElement('th')
        let apiVal = document.createTextNode(`${getValue(workbook, breakdown, 'C2', 'API')}`)

        apiCell.appendChild(apiVal)
        tableRow.appendChild(apiCell)

        //topPerf
        let topPerfCell = document.createElement('th')
        let topPerfVal = document.createTextNode(`${getValue(workbook, fracReport, 'C6', 'topPerf')}`)

        topPerfCell.appendChild(topPerfVal)
        tableRow.appendChild(topPerfCell)

        // bottomPerf
        let bottomPerfCell = document.createElement('th')
        let bottomPerfVal = document.createTextNode(`${getValue(workbook, fracReport, 'E6', 'bottomPerf')}`)

        bottomPerfCell.appendChild(bottomPerfVal)
        tableRow.appendChild(bottomPerfCell)

        // stage number
        let stageNoCell = document.createElement('th')
        let stageNoVal = document.createTextNode(`${getValue(workbook, summary, 'B1', 'stage number')}`)

        stageNoCell.appendChild(stageNoVal)
        tableRow.appendChild(stageNoCell)

        //ISIP start
        let isipStartCell = document.createElement('th')
        let isipStartVal = document.createTextNode(`${getValue(workbook, summary, 'C20', 'ISIP start')}`)

        isipStartCell.appendChild(isipStartVal)
        tableRow.appendChild(isipStartCell)

        // 3 min
        let threeMinCell = document.createElement('th')
        let threeMinVal = document.createTextNode(`${getValue(workbook, summary, 'C21', '3 Minutes')}`)

        threeMinCell.appendChild(threeMinVal)
        tableRow.appendChild(threeMinCell)

        // ISIP end
        let isipEndCell = document.createElement('th')
        let isipEndVal = document.createTextNode(`${getValue(workbook, summary, 'C35', 'ISIP End')}`)

        isipEndCell.appendChild(isipEndVal)
        tableRow.appendChild(isipEndCell)


        mainTable.appendChild(tableRow)

    };
    reader.readAsBinaryString(f);
  }
}

function getValue(workbook, workFile, address_of_cell, title) {
  // console.log(title);
  const worksheet = workbook.Sheets[workFile];
  const desired_cell = worksheet[address_of_cell];
  const desired_value = (desired_cell ? desired_cell.v : undefined);
  // console.log('value:', desired_value);
  return desired_value
}

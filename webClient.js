class webClient {
    self;
    constructor() {
        self = this;

        //let baseUrl = "http://api.altbv.com/V5/webclient/"
        let baseUrl = "https://localhost:44334/webclient/";

        this.configuration = {
            templateId: 2,
            processUrl: baseUrl + "Process"
            /*downloadFileName: "NTVX.xlsx",
            processUrl: baseUrl + "Process",
            downloadUrl: baseUrl + "DownloadExcel"*/
        }
        // other variables
       // this.downloadServerFileName = "" // this is sent by server
        //let response = null;
    }

    initialize() {
        //let url = this.configuration.templateUrl + templateId
        //this.httpRequest(url, "GET", undefined,
        //    data => {
        //        //console.log(data);
        //        //console.log(url);
        //        this.template = data;
        //        this.initializeControl();
        //        this.loadExcel();
        //        this.showButtons();
        //    }
        //)
    }

    //initializeControl() {
    //    //Initialize Spreadsheet component
    //    this.spreadsheet = new ej.spreadsheet.Spreadsheet({
    //        openUrl: this.configuration.openUrl,
    //        saveUrl: this.configuration.saveUrl,
    //        beforeSave: (args) => {
    //            args.fileName = this.template.templateFile.replace(".xlsx", "");
    //        },
    //        created: () => {
    //            self.spreadsheet.hideRibbonTabs(["Home", "Formulas", "Insert", "Data", "View"]);
    //        },
    //        cellEdit: (args) => {
    //            // console.log(args);
    //            let unlockRanges = this.template.actions[0].unlockRanges
    //            if (!self.IsCellInRange(args.address, unlockRanges))
    //                args.cancel = true;
    //        },
    //        openComplete: (args) => {
    //            self.spreadsheet.sheets[0].showHeaders = false;
    //        },
    //        fileMenuBeforeOpen: (args) => {
    //            if (args.parentItem.text === "File") {
    //                self.spreadsheet.hideFileMenuItems(["New", "Open"]);
    //            } else if (args.parentItem.text === "Save As") {
    //                self.spreadsheet.hideFileMenuItems(["Comma-separated values"]);
    //            }
    //        },
    //    });

    //    //Render initialized Spreadsheet component
    //    this.spreadsheet.appendTo('#spreadsheet');
    //}

    //loadExcel() {
    //    let request = new XMLHttpRequest();
    //    request.responseType = "blob";
    //    request.onload = () => {
    //        let file = new File([request.response], this.template.templateFile);
    //        this.spreadsheet.open({ file: file });
    //    }

    //    request.open("GET", this.configuration.fileUrl + this.template.templateId);
    //    request.send();
    //}

    //showButtons() {

    //    let ribbonTabItems = [{
    //        header: { text: "NTVX" },
    //        content: []
    //    }];
    //    this.template.actions.forEach(action => {
    //        ribbonTabItems[0].content.push({ text: action.button, cssClass: "ntvx-btn", click: this.postToServer });
    //    })
    //    this.spreadsheet.addRibbonTabs(ribbonTabItems);
    //    this.spreadsheet.element.querySelector('.e-ribbon .e-tab-header .e-toolbar-items').children[7].click();
    //}


    //IsCellInRange(cellId, unlockranges) {

    //    let cell = ej.spreadsheet.getRangeIndexes(cellId);
    //    let currentSheetName = cellId.split("!")[0]
    //    let ranges = unlockranges.split(";");
    //    let counter;
    //    for (counter = 0; counter < ranges.length; counter++) {
    //        let rangeSplit = ranges[counter].split("~");

    //        if (currentSheetName == rangeSplit[0]) {
    //            let sheetRanges = rangeSplit[1].split(",");

    //            for (let rngCtr = 0; rngCtr < sheetRanges.length; rngCtr++) {
    //                let range = ej.spreadsheet.getRangeIndexes(sheetRanges[rngCtr]);

    //                let inRange =
    //                    currentSheetName == rangeSplit[0] &&
    //                    range[0] <= cell[0] &&
    //                    range[2] >= cell[0] &&
    //                    range[1] <= cell[1] &&
    //                    range[3] >= cell[1]; // condition to check whether the cell is in between particular range 
    //                if (inRange) {
    //                    return true;
    //                }
    //            }
    //        }
    //    }
    //    return false
    //}

    async postToServer(buttonName) {
        const form = document.querySelector("form");
        const data = Object.fromEntries(new FormData(form).entries());
        data.buttonName = buttonName
        // uncomment
                 //   console.log(webModel);
        const formData = JSON.stringify(data);
        self.httpRequest(self.configuration.processUrl, "POST", formData, self.processServerResponse)

    }

    processServerResponse(serverData) {
        console.log(serverData)
        //this.response = serverData
        let keys = Object.keys(serverData)
        console.log(keys)

        for (let i = 0; i < keys.length; i++) {
            let key = keys[i]
            if (serverData.hasOwnProperty(key)) {
                let elem = document.getElementById(key)
                if (elem != null)
                    elem.innerHTML = serverData[key]
            }
        }

       // self.downloadServerFileName = serverData["downloadServerFileName"]
    }

   /* downloadExcel() {
        if (self.downloadServerFileName == "") {
            alert('Run calculator before downloading')
            return
        }

        let loadingDiv = document.getElementById("loading");
        loadingDiv.classList.remove("hidden")

        const data = { "file": self.downloadServerFileName }
        const formData = JSON.stringify(data);

        const xhttp = new XMLHttpRequest();
        xhttp.open("POST", self.configuration.downloadUrl);
        xhttp.withCredentials = true;
        xhttp.setRequestHeader("Content-Type", "application/json");
        xhttp.responseType = 'blob';

        xhttp.onload = function () {
            //Convert the Byte Data to BLOB object.
            var blob = new Blob([xhttp.response], { type: "application/octetstream" });

            //Check the Browser type and download the File.
            var isIE = false || !!document.documentMode;
            if (isIE) {
                window.navigator.msSaveBlob(blob, fileName);
            } else {
                let url = window.URL || window.webkitURL;
                let link = url.createObjectURL(blob);
                let a = document.createElement("a");
                a.setAttribute("download", self.configuration.downloadFileName);
                a.setAttribute("href", link);
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                loadingDiv.classList.add("hidden")
            }
        };
        xhttp.send(formData);
    }*/


   

     // reloadData() {
       // if (this.response) {
           /* let keys = Object.keys(serverData)
            console.log(keys)

            for (let i = 0; i < keys.length; i++) {
                let key = keys[i]
                if (serverData.hasOwnProperty(key)) {
                    let elem = document.getElementById(key)
                    if (elem != null)
                        elem.innerHTML = serverData[key]
                }
            }
          }*/
         // saveData(serverData[key])
     
    httpRequest(url, method, formData, successFunction) {
        //showing loading icon
        let loadingDiv = document.getElementById("loading");
       /* loadingDiv.classList.remove("hidden")

        const xhttp = new XMLHttpRequest();
        xhttp.open(method, url);
        xhttp.withCredentials = true;
        xhttp.setRequestHeader("Content-Type", "application/json");
        xhttp.send(formData);

        xhttp.onreadystatechange = (e) => {
            if (xhttp.readyState === XMLHttpRequest.DONE) {

                if (xhttp.status === 200) {
                    const data = JSON.parse(xhttp.responseText);
                    successFunction.call(this, data);
                    loadingDiv.classList.add("hidden")
                }

                if (xhttp.status === 500) {
                    console.error(xhttp.responseText);
                    alert(xhttp.responseText);
                    loadingDiv.classList.add("hidden")
                }
            }*/
    }



    //pasteServerSheets(serverData) {

    //    let ranges = serverData.pasteSheetsAndRange.split(";");

    //    let counter, item;
    //    let sourceSheetName, sourceRange, destSheetName, destStart;
    //    let sourceStartRow, sourceStartCol, sourceEndRow, sourceEndCol, destRow, destCol;
    //    let rangeStartEnd, destAddr, sourceAddr;
    //    let serverSheet, serverCell, cellValue;
    //    for (counter = 0; counter < ranges.length; counter++) {
    //        let range = ranges[counter];
    //        if (range === "")
    //            continue;
    //        item = range.split("~");
    //        sourceSheetName = item[0];
    //        sourceRange = item[1];
    //        destSheetName = item[2];
    //        destStart = item[3];

    //        serverSheet = serverData.sheets.find(s => s.name === sourceSheetName);

    //        rangeStartEnd = sourceRange.split(":");
    //        sourceStartRow = parseInt(rangeStartEnd[0].replace(/[A-Z]/g, ""));
    //        sourceStartCol = parseInt(this.columnNameToNumber(rangeStartEnd[0].replace(/[0-9]/g, "")));
    //        if (rangeStartEnd.length > 1) {
    //            sourceEndRow = parseInt(rangeStartEnd[1].replace(/[A-Z]/g, ""));
    //            sourceEndCol = parseInt(this.columnNameToNumber(rangeStartEnd[1].replace(/[0-9]/g, "")));
    //        }
    //        else {
    //            sourceEndRow = sourceStartRow;
    //            sourceEndCol = sourceStartCol;
    //        }
    //        destRow = parseInt(destStart.replace(/[A-Z]/g, ""));
    //        destCol = parseInt(this.columnNameToNumber(destStart.replace(/[0-9]/g, "")));

    //        for (let col = sourceStartCol; col <= sourceEndCol; col++) {
    //            for (let row = sourceStartRow; row <= sourceEndRow; row++) {
    //                destAddr = this.columnNumberToName(destCol + col - sourceStartCol) + (destRow + row - sourceStartRow);
    //                sourceAddr = this.columnNumberToName(col) + row;

    //                serverCell = serverSheet.cells.find(c => c.key === sourceAddr);
    //                cellValue = serverCell !== undefined ? serverCell.value : "";
    //                this.spreadsheet.updateCell({ value: cellValue }, destSheetName + "!" + destAddr);
    //            }
    //        }
    //    }
    //}

    //columnNameToNumber(columnName) {
    //    let columnNumber = 0;
    //    for (let counter = 0; counter < columnName.length; counter++) {
    //        columnNumber = (columnName.charCodeAt(counter) - 64) + columnNumber * 26;
    //    }

    //    return columnNumber;
    //}

    //columnNumberToName(columnNumber) {
    //    let columnName = "";
    //    let reminder;
    //    while (columnNumber > 0) {
    //        reminder = (columnNumber - 1) % 26;
    //        columnName = String.fromCharCode(reminder + 65) + columnName;
    //        columnNumber = parseInt((columnNumber - reminder) / 26);
    //    }

    //    return columnName;
    //}
}

let client = new webClient();
client.initialize();

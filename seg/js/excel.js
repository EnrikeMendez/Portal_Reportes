function sheet2blob(sheet, sheetName) {
    sheetName = sheetName || 'sheet1';

    var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;

    var wopts = {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);

    var blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    });
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    return blob;
}
function sheet2blob(sheet, sheet2, sheetName, sheetName2) {
    sheetName = sheetName || 'sheet1';
    sheetName2 = sheetName2 || 'sheet2';

    var workbook = {
        SheetNames: [sheetName, sheetName2],
        Sheets: {}
    };

    workbook.Sheets[sheetName] = sheet;
    workbook.Sheets[sheetName2] = sheet2;

    var wopts = {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);

    var blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    });
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    return blob;
}
function sheet2blob(sheet, sheet2, sheet3, sheetName, sheetName2, sheetName3) {
    sheetName = sheetName || 'sheet1';
    sheetName2 = sheetName2 || 'sheet2';
    sheetName3 = sheetName3 || 'sheet3';

    var workbook = {
        SheetNames: [sheetName, sheetName2, sheetName3],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    workbook.Sheets[sheetName2] = sheet2;
    workbook.Sheets[sheetName3] = sheet3;

    var wopts = {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);

    var blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    });
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}

function openDownloadDialog(url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url);
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || '';
    var event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}

function export3sheetsTOexcel(idTbl1, idTbl2, idTbl3, sheetNamesByComma, fileName) {
    var sheetNames;
    var tbl1, tbl2, tbl3;
    var sheet1, sheet2, sheet3;
    var sheetName1, sheetName2, sheetName3;

    sheetName1 = "";
    sheetName2 = "";
    sheetName3 = "";

    if (idTbl1 != "") { tbl1 = document.querySelector("#" + idTbl1); }
    if (idTbl2 != "") { tbl2 = document.querySelector("#" + idTbl2); }
    if (idTbl3 != "") { tbl3 = document.querySelector("#" + idTbl3); }

    if (tbl1 != null) { sheet1 = XLSX.utils.table_to_sheet(tbl1); }
    if (tbl2 != null) { sheet2 = XLSX.utils.table_to_sheet(tbl2); }
    if (tbl3 != null) { sheet3 = XLSX.utils.table_to_sheet(tbl3); }
    
    if (sheetNamesByComma != "") {
        sheetNames = sheetNamesByComma.split(",");
        
        for (var i = 0; i < sheetNames.length; i++) {
            if (i == 0) { sheetName1 = sheetNames[i] == "" ? "" : sheetNames[i]; }
            if (i == 1) { sheetName2 = sheetNames[i] == "" ? "" : sheetNames[i]; }
            if (i == 2) { sheetName3 = sheetNames[i] == "" ? "" : sheetNames[i]; }
        }
    }
    
    if (sheet1 != null && sheet2 != null && sheet3 != null) {
        openDownloadDialog(sheet2blob(sheet1, sheet2, sheet3, sheetName1, sheetName2, sheetName3), fileName + '.xlsx');
    }
    else if (sheet1 != null && sheet2 != null) {
        openDownloadDialog(sheet2blob(sheet1, sheet2, sheetName1, sheetName2), fileName + '.xlsx');
    }
    else if (sheet1 != null) {
        openDownloadDialog(sheet2blob(sheet1, sheetName1), fileName + '.xlsx');
    }
    else {
        alert("No se puede generar un archivo xls vacio.");
    }
}
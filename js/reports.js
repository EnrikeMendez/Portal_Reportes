function GenerarExcel() {
	showLoading();

	const fecha = new Date();
	var NombreReporte = "Rpt_ConsultaGeneralReportes";
	var CmbEstatus = document.getElementById("select_activos");
	if (CmbEstatus.value != "") {
		NombreReporte = NombreReporte.replace("General", "");
		NombreReporte += CmbEstatus.options[CmbEstatus.selectedIndex].text;
	}
	NombreReporte += "_" + fecha.getFullYear().toString() + ("0" + (fecha.getMonth() + 1).toString()).slice(-2) + ("0" + fecha.getDate().toString()).slice(-2) + ("0" + fecha.getHours().toString()).slice(-2);
	exportTableToExcel("select_reporte", NombreReporte.replace(" ", ""));

	hideLoading();
}

function exportTableToExcel(tableID, filename = '') {
	var objById;
	var objHTML;
	var tableHTML;
	var tableSelect;
	var downloadLink;
	var dataType = 'application/vnd.ms-excel';
	var arrColumnsToDelete = document.getElementsByClassName("delC");
	var arrElementsToDelete = ['select_activos', 'hdnEtape_4', 'table-buscar', 'select_activos'];
	var arrHeadersToChange = ['thNo', 'thName', 'thArea', 'thPrio', 'thType', 'thCus', 'thFre', 'tdDS', 'tdDM', 'tdDSe', 'tdH', 'tdM', 'tdUC', 'tdFC', 'tdUM', 'tdFM', 'tdP1', 'tdP2', 'tdP3', 'tdP4', 'thCmd'];

	showLoading();

	tableSelect = document.getElementById(tableID);
	tableHTML = tableSelect.outerHTML;

	for (var i in arrHeadersToChange) {
		objById = document.getElementById(arrHeadersToChange[i]);

		if (objById != null) {
			objHTML = document.getElementById(arrHeadersToChange[i]).outerHTML;
			tableHTML = tableHTML.replace(objHTML, objHTML.replace('rowspan="2"', 'rowspan="1"'));
		}

		objById = null;
	}

	for (var i in arrColumnsToDelete) {
		objHTML = arrColumnsToDelete[i].outerHTML;
		tableHTML = tableHTML.replace(objHTML, "");
	}

	for (var i in arrElementsToDelete) {
		objById = document.getElementById(arrElementsToDelete[i]);

		if (objById != null) {
			objHTML = document.getElementById(arrElementsToDelete[i]).outerHTML;
			tableHTML = tableHTML.replace(objHTML, "");
		}

		objById = null;
	}

	var colObjHTML;
	var s = 0;

	tableHTML = "<meta http-equiv='Content-Type' content='text/html;' charset='utf-8' /> " + tableHTML;
	tableHTML = "<style>.delC{display:none;visibility:collapse;} td{border:thin solid;} .trHeader{background-color:#223F94;font-family:'Roboto',sans-serif;color:#FFFFFF;}<style> " + tableHTML;


	var newObjHTML;
	var arrId = ["thNo", "tdDS", "tdDM", "tdDSe", "tdUC", "tdFC", "tdUM", "tdFM"];
	var arrAnt = ["N°", "Días", "Días", "Días", "creación", "creación", "modificación", "modificación"];
	var arrNew = ["No.", "Dias", "Dias", "Dias", "creacion", "creacion", "modificacion", "modificacion"];

	for (var j in arrId) {
		objById = document.getElementById(arrId[j]);

		if (objById != null) {
			objHTML = document.getElementById(arrId[j]).outerHTML;
			newObjHTML = document.getElementById(arrId[j]).outerHTML.replace(arrAnt[j], arrNew[j]);
			newObjHTML = newObjHTML.replace(arrAnt[j], arrNew[j]);
			tableHTML = tableHTML.replace(objHTML, newObjHTML);
		}

		objById = null;
	}

	var arrTR = document.getElementsByTagName("tr");

	for (var k in arrTR) {
		try {
			if (typeof arrTR[k] != 'undefined' && arrTR[k] != null) {
				if (arrTR[k].style.display == 'none') {
					objHTML = arrTR[k].outerHTML;
					tableHTML = tableHTML.replace(objHTML, "");
				}
			}
		}
		catch { }
	}

	objHTML = null;
	newObjHTML = null;

	filename = filename ? filename + '.xls' : 'excel_data.xls';
	downloadLink = document.createElement("a");
	document.body.appendChild(downloadLink);


	if (navigator.msSaveOrOpenBlob) {
		var blob = new Blob(['ufeff', tableHTML], {
			type: dataType
		});
		navigator.msSaveOrOpenBlob(blob, filename);
	} else {
		downloadLink.href = 'data:' + dataType + ', ' + encodeURIComponent(tableHTML);
		downloadLink.download = filename;
		downloadLink.click();
	}

	document.body.removeChild(downloadLink);

	hideLoading();
}

function showLoading() {
	var dvloading = null;
	dvloading = document.getElementById("dvloading");

	if (dvloading != null) {
		dvloading.style.display = "";
		dvloading.style.visibility = "visible";
	}
}

function hideLoading() {
	var dvloading = null;
	dvloading = document.getElementById("dvloading");

	if (dvloading != null) {
		dvloading.style.display = "none";
		dvloading.style.visibility = "collapse";
	}
}
const formatter = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2, });
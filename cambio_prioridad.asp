<%@  language="VBScript" %>
<% option explicit %>
<!--#include file="include/include.asp"-->

<%
'admin of logis web site :
'modificacion de reportes
	Response.Expires = 0
	call check_session()
	dim SQL, arrayRS, SQL_02, arrayRS2, i, rst, arrayRS3, log_prioridad_dinamica, j
	set rst = Server.CreateObject("ADODB.Recordset")

Function NVL(str)
	if IsNull(str) then
		NVL = "" 
	else 
		NVL = str
	end if
End Function

	'if Request("reporte[]") <> "" then	
	'	SQL = "update rep_chron set priorite = " & Request("select_prioridad") & " where id_chron in (" & Request("reporte[]") & ")"
	'	rst.Open SQL, Connect(), 0, 1, 1
	'	Response.Write  "<br>Se cambio a prioridad " & Request("select_prioridad") & " los id_chron " & Request("reporte[]") & " de la tabla rep_cron."
	'end if
%>
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html;" charset="iso-8859-1" />

    <!--ORP: WS AJAX-->
    <!--call print_style()-->
    <link href="css/print_style.css" type="text/css" rel="stylesheet" />
    <!--ORP: WS AJAX-->


    <link href="include/logis.css" type="text/css" rel="stylesheet" />
    <link href="css/logis_style.css" type="text/css" rel="stylesheet" />
    <script language="JavaScript" src="./include/tigra_tables.js"></script>
    <script type="text/javascript" src="js/reports.min.js"></script>
    <!--ORP: WS AJAX-->
    <script src="js/jquery-1.3.2.min.js"></script>
    <script src="js/main.js"></script>
    <!--ORP: WS AJAX-->

    <script type="text/javascript">



        var MinutosRecargarPagina = 5;

        var totLabel = "";
        const fecha = new Date();
        var tot = (fecha.getFullYear() - 2000) + (fecha.getMonth() + 1) + fecha.getDate() + fecha.getHours() + fecha.getMinutes() + fecha.getSeconds();
        tot = tot / 60;
        totLabel = tot.toFixed(1).toString();
    </script>
    <title>Cambio de prioridad</title>
</head>
<body onload="refresca_pagina();">
    <div id="dvloading" style="display: block!important; visibility: visible!important;">
        <center>Procesando </center>
        <center>
            <img alt=". . ." id="imgPuntos" src="images/puntosSuspensivos.gif" /></center>
    </div>
    <div class="contenedorMenu">
        <div class="dvMenu">
            <ul id="menu">
                <div class="logo-logis">
                    <img src="images/logo-logis-s.png" style="height: 50px;" />
                </div>
                <li onclick="window.location.href='menu.asp';" class="link_cursor">Inicio
                </li>
                <li id="imgXls" alt="Exportar" title="Exportar consulta" onclick="GeneraExcel('CambioPrioridad','select_reporte')" class="link_cursor">Exportar consulta
                </li>
            </ul>
            <h2>CAMBIO DE PRIORIDAD
            </h2>
        </div>
    </div>

    <%
	
			'< ---CHG-DESA-27042022-01
			'invoca el seteo dinamico de prioridad
			call sub_procesos_prioridad_dinamica()
			' CHG-DESA-27042022-01-- >

			if not IsArray(arrayRS) then
				Response.Write "<script>MinutosRecargarPagina = totLabel;</script>"
			end if

    %>

    <form action="cambio_prioridad.asp" method="post">
        <center>
            <table width="98%" border="0" class="tbl-shadow">
                <tr>
                    <td colspan="2">
                        <input id="table-buscar" type="text" class="form-control rounded-txt" placeholder="Escriba algo para filtrar" style="width: 100%;" />
                    </td>
                </tr>
                <tr>
                    <td class="width-15p">
                        <label>Prioridad</label>
                        <select name="select_prioridad" id="select_prioridad" class="form-control rounded-cmb">
                            <option value="0" selected>0</option>
                            <option value="1">1</option>
                            <option value="2">2</option>
                            <option value="3">3</option>
                            <option value="4">4</option>
                            <option value="5">5</option>
                            <option value="6">6</option>
                            <option value="7">7</option>
                            <option value="8">8</option>
                            <option value="9">9</option>
                        </select>
                    </td>
                    <td>

                        <!--ORP: WS AJAX-->
                        <!--<button type="submit" class="rounded-btn">Guardar</button>-->
                        <button id="btn_guardar" class="rounded-btn" onclick="ftn_GettModificaCambioPrioridad()">Guardar</button>
                        <!--ORP: WS AJAX-->

                    </td>
                </tr>
            </table>
        </center>
        <br />
        <table width="100%" border="0" id="select_reporte" class="tblContent">
            <thead>
                <tr align="center">
                    <th>&nbsp;</th>
                    <th>ID rep</th>
                    <th>Nombre</th>
                    <th>Prioridad</th>
                    <th>Cliente</th>
                    <th>ID cron</th>
                    <th>Nombre detalle</th>
                    <th>Fecha</th>
                    <th>Email</th>
                </tr>
            </thead>
            <tbody id="print_data">
            </tbody>
        </table>
    </form>

    <div id="clockCounter" name="clockCounter" class="fixedBottomRightLabel">
        La p&aacute;gina se actualizar&aacute en
			<label id="lblTimer" name="lblTimer">60</label>
        <label id="lblMedidaTiempo" name="lblMedidaTiempo">segundo</label>(s).
    </div>

    <script type="text/javascript">
        var reloadTime = MinutosRecargarPagina * 60 * 1000;

        //<!--
        tigra_tables('select_reporte', 4, 0, '#ffffff', '#ffffcc', '#ffcc66', '#cccccc');
        // -->

        function refresca_pagina() {
            counter();
            setTimeout(function () {
                location.reload();
            }, reloadTime);
            hideLoading();
        }
        function counter() {
            var dNow = new Date();
            var redColorSeconds = 5;
            var t = Math.round(reloadTime / 1000);
            var lblTimer = document.getElementById("lblTimer");
            var lblText = document.getElementById("clockCounter");
            var lblMedidaTiempo = document.getElementById("lblMedidaTiempo");

            try { redColorSeconds = ((dNow.getDate() + dNow.getMonth()) * 0.7) + 1; }
            catch { redColorSeconds = redColorSeconds + 1; }

            lblTimer.innerHTML = "<b>" + MinutosRecargarPagina + "</b>";
            lblMedidaTiempo.innerHTML = "minuto";

            window.setInterval(function () {
                lblTimer.innerHTML = t - 1;
                t--;

                if (t <= 60) {
                    lblMedidaTiempo.innerHTML = "segundo";
                }
                if (t < redColorSeconds) {
                    lblTimer.style.color = "red";
                }
                if (t <= (redColorSeconds * 1.5)) {
                    lblText.style.display = "block";
                }
                else {
                    lblText.style.display = "none";
                }
            }, 1000);
        }
        function showLoading() {
            document.getElementById("dvloading").style.display = "";
            document.getElementById("dvloading").style.visibility = "visible";
        }
        function hideLoading() {
            document.getElementById("dvloading").style.display = "none";
            document.getElementById("dvloading").style.visibility = "collapse";
        }

        $TableFilter = function (id, value) {
            var rows = document.querySelectorAll(id + ' tbody tr');

            for (var i = 0; i < rows.length; i++) {
                var showRow = false;

                var row = rows[i];
                row.style.display = 'none';

                for (var x = 0; x < row.childElementCount; x++) {
                    if (row.children[x].textContent.toLowerCase().indexOf(value.toLowerCase().trim()) > -1) {
                        showRow = true;
                        break;
                    }
                }

                if (showRow) {
                    row.style.display = null;
                }
            }
        }

        document.querySelector("#table-buscar").onkeyup = function () {
            $TableFilter("#select_reporte", this.value);
        }


        //< !--ORP: WS AJAX-- >
        function ftn_GetConsultaCambioPrioridad() {
            const xhr = new XMLHttpRequest();
            //const url = "http://localhost:51687/Report_Service.svc/GetConsultaCambioPrioridad";
            const url = urlWebService + "GetConsultaCambioPrioridad";

            xhr.onreadystatechange = function () {
                if (xhr.readyState == XMLHttpRequest.DONE) {
                    ftn_consulta_cambio_prioridad(xhr.responseText);
                }
            }

            xhr.open("GET", url, true);
            xhr.send();
        }

        function ftn_GettModificaCambioPrioridad() {
            const xhr = new XMLHttpRequest();

            var id_crons = "";
            var prioridad = 0;

            prioridad = document.getElementById("select_prioridad").value;
            id_crons = document.getElementById("reporte[]").value;

            if (id_crons != "") {
                //const url = "http://localhost:51687/Report_Service.svc/GetModificaCambioPrioridad?id_crons = "+ id_crons +" ";
                const url = urlWebService + "GetModificaCambioPrioridad?id_crons=" + id_crons + "&prioridad= " + prioridad + "";

                xhr.onreadystatechange = function () {
                    if (xhr.readyState == XMLHttpRequest.DONE) {
                        ftn_modifica_cambio_prioridad(xhr.responseText);
                    }
                }
                xhr.open("GET", url, true);
                xhr.send();
            }
            else {
                alert("Seleccione un reporte primero.");
            }

            
        }



        function ftn_consulta_cambio_prioridad(wsResponseText) {

            var objResult = JSON.parse(wsResponseText);
            var info = objResult.GetConsultaCambioPrioridadResult;
            var arrayData = JSON.parse(info);

            var i = 0;
            var htmlTable = "";

            if (arrayData.length == 0) {
                htmlTable = htmlTable + "<tr class='center' >";
                htmlTable = htmlTable + "<td colspan='9' class='center'>";
                htmlTable = htmlTable + " No hay Reportes en ejecuci&oacute;n.";
                htmlTable = htmlTable + "</td>";
                htmlTable = htmlTable + "</tr>";

                $("#print_data").append(htmlTable);

            } else {

                for (i = 0; i < arrayData.length; i++) {

                    htmlTable = "";

                    htmlTable = htmlTable + "<tr>";

                    htmlTable = htmlTable + " <td align='center'><input type='checkbox' name='reporte[]' id='reporte[]' value='" + arrayData[i].ID_CHRON + "'></td> \n\n"


                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].ID_REP;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td> <font class='carmin'>";
                    htmlTable = htmlTable + arrayData[i].NOMBRE_REPORTE;
                    htmlTable = htmlTable + "</font> </td> \n\n";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].PRIORITE;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].CLIENTE;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].ID_CRON;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td>";
                    htmlTable = htmlTable + arrayData[i].NOMBRE_DETALLE;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td>";
                    htmlTable = htmlTable + arrayData[i].HORA_CREACION;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td>";
                    if (arrayData[i].DEST_MAIL != "") {
                        htmlTable = htmlTable + arrayData[i].DEST_MAIL;
                    }
                    htmlTable = htmlTable + "</td> \n\n";


                    htmlTable = htmlTable + "</tr> \n\n";
                }

                $("#print_data").append(htmlTable);
            }

        }
        $(document).ready(ftn_GetConsultaCambioPrioridad);

        function ftn_modifica_cambio_prioridad(wsResponseText) {

            var objResult = JSON.parse(wsResponseText);
            var info = objResult.GetConsultaCambioPrioridadResult;
            alert(info);
        }

            //< !--ORP: WS AJAX-- >



    </script>
</body>

<%
		'< ---CHG-DESA-27042022-01
		'----------------------------'Setear la prioridad de los procesos dinamicamente:----------------------------
			Sub sub_procesos_prioridad_dinamica()
				dim bandera,SQL2,arrayRS2
				    bandera = 0

    	    SQL2 = " select reporte.id_rep,  chron.priorite, rep_det.id_cron, chron.id_chron " & VbCrLf  
			SQL2 = SQL2 & " from rep_detalle_reporte rep_det  " & VbCrLf 
			SQL2 = SQL2 & " join REP_CHRON chron on chron.id_rapport = rep_det.id_cron " & VbCrLf  
            SQL2 = SQL2 & " join rep_reporte reporte on reporte.id_rep = rep_det.id_rep  " & VbCrLf 
            SQL2 = SQL2 & " where chron.active = 1  " & VbCrLf 
			SQL2 = SQL2 & " and chron.MINUTES is null and chron.HEURES is null and chron.JOURS is null and chron.MOIS is null and chron.JOUR_SEMAINE is null and chron.LAST_EXECUTION is null  " & VbCrLf 
            SQL2 = SQL2 & " order by chron.priorite, id_cron desc "
            arrayRS2 = GetArrayRS(SQL2)

				if not IsArray(arrayRS2) then
					'Response.Write "Prioridad dinamica en espera ..."
					Response.Write "<script>console.log('Prioridad dinamica en espera ...')</script>"
				else
					'Response.Write "Reportes modificados de prioridad dinamicamente: "
					Response.Write "<script>console.log('Reportes modificados de prioridad dinamicamente:')</script>"
				dim arrTmp
				arrTmp = CatalogoPrioridadDinamica()
				if IsArray(arrTmp) then
				For i=0 to UBound(arrayRS2,2)
				     for j=0 to UBound(arrTmp,2)
				     	if CStr(arrayRS2(0,i)) = CStr(arrTmp(0,j)) Then
				     		if CStr(arrayRS2(1,i)) <> CStr(arrTmp(1,j)) Then
                                  call sub_set_prioridad(CStr(arrayRS2(0,i)),CStr(arrTmp(1,j)),CStr(arrayRS2(2,i))) 
				     			  bandera = 1
                            else
                                  Response.Write "<script>console.log('OK: "& CStr(arrayRS2(0,i)) & " " & CStr(arrayRS2(2,i)) & " " & CStr(arrTmp(1,j)) &"')</script>"
                                  bandera = 1
					end if
				     	end if	
				     next  
				Next
				end if
				if bandera = 0 Then
				  	'Response.Write " Ninguno de la lista actual." 
					Response.Write "<script>console.log(' Ninguno de la lista actual.')</script>"
				end if
				end if				
			End Sub
			Sub	sub_set_prioridad(id_rep, prioridad, id_chron)
				SQL ="update rep_chron set priorite = "& prioridad &" where ID_CHRON IN (select ID_CHRON  from  rep_chron chron" & VbCrLf 
				SQL = SQL & "JOIN rep_detalle_reporte rep_detalle on chron.id_rapport = rep_detalle.id_cron   " & VbCrLf
				SQL = SQL & "JOIN rep_reporte rep on rep.ID_REP = rep_detalle.id_rep " & VbCrLf
				SQL = SQL & "where chron.active = 1" & VbCrLf
				SQL = SQL & "and chron.MINUTES is null and chron.HEURES is null and chron.JOURS is null and chron.MOIS is null and chron.JOUR_SEMAINE is null and chron.LAST_EXECUTION is null " & VbCrLf 
				SQL = SQL & "and rep_detalle.id_rep  = "& id_rep &" )" & VbCrLf
				rst.Open SQL, Connect(), 0, 1, 1
				'Response.Write  "("& id_rep & ") " & id_chron & " -" & "- > " & prioridad & "|" 
				Response.Write "<script>console.log('CHANGED: "& id_rep & " " & id_chron & " " & prioridad &"')</script>"
			End Sub
		'-----------------------------------------------------------------------------------------------------------
		' CHG-DESA-27042022-01-- >
%>
</html>

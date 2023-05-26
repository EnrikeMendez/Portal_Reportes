<%@  language="VBScript" %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
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

	if Request("reporte[]") <> "" then	
		SQL = "update rep_chron set priorite = " & Request("select_prioridad") & " where id_chron in (" & Request("reporte[]") & ")"
		rst.Open SQL, Connect(), 0, 1, 1
		Response.Write  "<br>Se cambio a prioridad " & Request("select_prioridad") & " los id_chron " & Request("reporte[]") & " de la tabla rep_cron."
	end if
%>
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html;" charset="iso-8859-1" />
    <% call print_style() %>
    <link href="include/logis.css" type="text/css" rel="stylesheet" />
    <link href="css/logis_style.css" type="text/css" rel="stylesheet" />
    <script language="JavaScript" src="./include/tigra_tables.js"></script>
	<script type="text/javascript" src="js/reports.min.js"></script>
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
				<li id="imgXls" alt="Exportar" title="Exportar consulta" onclick="GeneraExcel('CambioPrioridad','select_reporte')" class="link_cursor">
					Exportar consulta
				</li>
            </ul>
            <h2>CAMBIO DE PRIORIDAD
            </h2>
        </div>
    </div>

    <%


			SQL = "select reporte.id_rep, reporte.name as nombre_reporte, chron.priorite, cliente, rep_det.id_cron, rep_det.name as nombre_detalle, TO_char(date_created,'DD/MON/YYYY HH24:MI') as hora_creacion, rep_det.dest_mail, id_chron " & VbCrLf 
			SQL = SQL & " from rep_detalle_reporte rep_det " & VbCrLf 
			SQL = SQL & " join REP_CHRON chron on chron.id_rapport = rep_det.id_cron " & VbCrLf 
			SQL = SQL & "  join rep_reporte reporte on reporte.id_rep = rep_det.id_rep " & VbCrLf 
			SQL = SQL & "  where chron.active = 1 " & VbCrLf 
			SQL = SQL & " and chron.MINUTES is null and chron.HEURES is null and chron.JOURS is null and chron.MOIS is null and chron.JOUR_SEMAINE is null and chron.LAST_EXECUTION is null " & VbCrLf 
			SQL = SQL & "  order by chron.priorite, id_cron desc " 
			arrayRS = GetArrayRS(SQL)
			
		
			
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
                        <select name="select_prioridad" class="form-control rounded-cmb">
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
                        <button type="submit" class="rounded-btn">Guardar</button>
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
            <tbody>
                <%
					if not IsArray(arrayRS) then
						Response.Write "<tr class='center'>"
						Response.Write "	<td colspan='9' class='center'>"
						Response.Write "		No hay Reportes en ejecuci&oacute;n."
						Response.Write "	</td>"
						Response.Write "</tr>"
					else
						for i = 0 to UBound(arrayRS,2)
							Response.Write "<tr>"
							Response.Write "	<td align='center'><input type='checkbox' name='reporte[]' value='" & arrayRS(8,i) & "'></td>" & vbCrLf & vbTab
							Response.Write "	<td align='center'>"& arrayRS(0,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td> <font class='carmin'>" & arrayRS(1,i) & "</font></td>"
							Response.Write "	<td align='center'>"& arrayRS(2,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td align='center'>"& arrayRS(3,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td align='center'>"& arrayRS(4,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td>"& arrayRS(5,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td>"& arrayRS(6,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td>"& arrayRS(7,i) &"</td>" & vbCrLf & vbTab
							Response.Write "<tr>"
						next
					end if
                %>
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

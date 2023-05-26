<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'modificacion de reportes
	Response.Expires = 0
	call check_session()
	dim SQL, arrayRS, SQL_02, arrayRS2, i, i2, rst, bandera, log_prioridad_dinamica, j
	set rst = Server.CreateObject("ADODB.Recordset")
	'<<MRG 
	Dim nowDay, backDays
	backDays = 1
	nowDay = Now
	'MRG>>
Function NVL(str)
	if IsNull(str) then
		NVL = "" 
	else 
		NVL = str
	end if
End Function
	
	if Request("num_reporte") <> "" then 
		'Reprocesar insert
		SQL = "insert into rep_chron (id_chron, id_rapport, priorite, test, active) " & VbCrLf 
		SQL = SQL & " values (SEQ_CHRON.nextval, '" & SQLEscape(Request.Form("num_reporte")) & "', 1,0, 1) "
		rst.Open SQL, Connect(), 0, 1, 1
		Response.Redirect("monitoreo_reportes.asp?MENSAJE=Se reproceso el id_rapport " & SQLEscape(Request.Form("num_reporte")))
	end if
	if  REQUEST("mensaje") <> ""  then 
		Response.Write  REQUEST("mensaje")
	end if 
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
		<script language="JavaScript" src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
		
		<script type="text/javascript">
			var MinutosRecargarPagina = 5;

			var totLabel = "";
			const fecha = new Date();
            var tot = (fecha.getFullYear() - 2000) + (fecha.getMonth() + 1) + fecha.getDate() + fecha.getHours() + fecha.getMinutes() + fecha.getSeconds();
			tot = tot / 60;
			totLabel = tot.toFixed(1).toString();
        </script>
		
		<title>MONITOREO REPORTES</title>
	</head>
	<body onload="refresca_pagina();">
		<style>
			.hidden 
			{
				visibility: hidden;
		    }
		    .text-right
			{
				text-align: right;
		    }
			.tblContent > tbody > tr:nth-child(odd) > td, .tblContent > tbody > tr:nth-child(odd) > th
			{
				background: none;
			}
		    .error-tr 
			{
				background: #BB2D3B;
				color: white;
		    }
			.error-tr:hover, .success-tr:hover
			{
				color: black;
		    }

		    .success-tr
			{
				background: #157347;
				color: white;
		    }
			.tr-F2F2F2
			{
				background: #F2F2F2;
		    }
		</style>
		<div id="dvloading" style="display:block!important;visibility:visible!important;">
			<center>Procesando </center>
			<center><img alt=". . ." id="imgPuntos" src="images/puntosSuspensivos.gif" /></center>
		</div>
		<div class="contenedorMenu">
			<div class="dvMenu">
				<ul id="menu">
					<div class="logo-logis">
						<img src="images/logo-logis-s.png" style="height:50px;" />
					</div>
					<li onclick="window.location.href='menu.asp';" class="link_cursor">
						Inicio
					</li>
				</ul>
				<h2>
					MONITOREO REPORTES
				</h2>
			</div>
		</div>

		<%
		'<<MRG
		'El lunes se mostrarán desde el viernes por la tarde:
			if Weekday(nowDay,1) = 2 then
				backDays = 3
			end if
		'MRG>>
			SQL = "select  cron.id_rapport as id_cron, TO_char(cron.last_execution,'DD/MON/YYYY HH24:MI') as hora_creacion, rep_detalle.name,  " & VbCrLf 
			SQL = SQL & " rep.id_rep || ' - ' ||rep.name as tipo_reporte, " & VbCrLf 
			SQL = SQL & " (cron.MOIS || cron.JOUR_SEMAINE || cron.HEURES || cron.MINUTES || cron.JOURS) as programacion, " & VbCrLf 
			SQL = SQL & "  cron.priorite, cron.test, cron.in_progress, rep_detalle.id_cron, " & VbCrLf 
			SQL = SQL & "  (select error.log from rep_chron_error error, rep_detalle_reporte reporte where trunc(error.date_created) = trunc(sysdate) and error.id_reporte = reporte.id_cron and error.id_reporte = rep_detalle.id_cron and rownum = 1)as errores " & VbCrLf 
			SQL = SQL & "  , cron.id_chron, reprocesos.nombre_proceso ,reprocesos.status" & VbCrLf 
			SQL = SQL & "  from REP_CHRON cron " & VbCrLf 
			SQL = SQL & "  JOIN rep_detalle_reporte rep_detalle on cron.id_rapport = rep_detalle.id_cron   " & VbCrLf 
			SQL = SQL & "  JOIN rep_reporte rep on rep.ID_REP = rep_detalle.id_rep  " & VbCrLf 
			SQL = SQL & "  LEFT OUTER JOIN rep_reprocesos_reporte reprocesos on reprocesos.id_cron = cron.id_rapport  " & VbCrLf 
			SQL = SQL & "  where cron.active <> 0 " & VbCrLf 
			'<<MRG
			'SQL = SQL & "  and trunc(cron.last_execution) = trunc(sysdate) " & VbCrLf
			SQL	=	SQL	&	"AND cron.last_execution	between sysdate - " & backDays & " and sysdate "	&	VbCrLf
			'MRG>>
			SQL = SQL & "  order by  cron.in_progress desc, hora_creacion desc " 
			
			arrayRS = GetArrayRS(SQL)

			'< -- CHG-DESA-27042022-01
			'invoca el seteo dinamico de prioridad
			call sub_procesos_prioridad_dinamica()
			' CHG-DESA-27042022-01 -- >

			if not IsArray(arrayRS) then
				Response.Write "<script>MinutosRecargarPagina = totLabel;</script>"
			end if
		%>
		
		<form action="monitoreo_reportes.asp" method="post">
			<center>
				<table width="98%" border="0" class="tbl-shadow">
					<tr>
						<td colspan="2">
							<input id="table-buscar" type="text" class="form-control rounded-txt" placeholder="Escriba algo para filtrar" style="width: 100%;"/>
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
						<th>Hora de creacion</th>
						<th style="text-align:left;">Nombre</th>
						<th style="text-align:left;">Tipo de reporte</th>
						<th>Programacion</th>
						<th style="text-align:right;">Priorite</th>
						<th style="text-align:right;">In_progress</th>
						<th style="text-align:right;">Error</th>
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
							
							if arrayRS(7,i) <> "0" then
								Response.Write "<tr class='success-tr'>"
							elseif arrayRS(9,i) <> "" then
								Response.Write "<tr class='error-tr'>"
							else
								if Int(i / 2) * 2 = i then 
									Response.Write "<tr class='tr-F2F2F2'>"
								else 
									Response.Write "<tr>"
								end if 
							end if
					
							Response.Write "	<td align='center'><input type='checkbox' name='reporte[]' value='" & arrayRS(10,i) & "'></td>" & vbCrLf & vbTab
							Response.Write "	<td align='center'>"& arrayRS(0,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td>" & arrayRS(1,i) & "</td>"
							Response.Write "	<td aling='left'>"& arrayRS(2,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td aling='left'>"& arrayRS(3,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td align='center'>"& arrayRS(4,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td align='right'>"& arrayRS(5,i) &"</td>" & vbCrLf & vbTab
							Response.Write "	<td align='right'>"& arrayRS(7,i) &"</td>" & vbCrLf & vbTab
							if arrayRS(9,i)  <> "" then 
								Response.Write "	<td align='right'><button type='button' class='modal-click' data-idcron='"& arrayRS(8,i) &"' data-status='"& arrayRS(12,i) &"'>Ver error</button></td>" & vbCrLf & vbTab
							else	
								Response.Write "	<td align='right'></td>" & vbCrLf & vbTab
							end if 
							
							Response.Write "<tr>"
						next
					end if
				%>
				</tbody>
			</table>
		</form>
		<!-- The Modal -->
		<div id="myModal" class="modal">
			<!-- Modal content -->
			<div class="modal-content">
				<span class="close">&times;</span>
				
				<table width="100%" border="0" id="select_reporte" class="tblContent">
				<thead>
					<tr>
						<th>ID</th>
						<th style="text-align: left;">Errores</th>
						<th>Fecha</th>
					</tr>
				</thead>
				<tbody>
					<% 
						if IsArray(arrayRS) then
							SQL = "select * from rep_chron_error where id_reporte in( "
							bandera = 0
							for i = 0 to UBound(arrayRS,2)
								if(not IsArray(arrayRS(8,i))) then
									if bandera = 0 then 
										SQL = SQL & arrayRS(0,i) 
										bandera = 1
									end if 
									SQL = SQL & ", " & arrayRS(0,i) 
								end if
							next
							SQL = SQL & ") and trunc(date_created) = trunc(sysdate) " 
							
							arrayRS2 = GetArrayRS(SQL)
							
							if IsArray(arrayRS2) then
								for i2 = 0 to UBound(arrayRS2,2)
									Response.Write "<tr class='tr-moda tr-moda-"& arrayRS2(1,i2) &"'>"
									Response.Write "	<td align='center'>"& arrayRS2(1,i2) &"</td>" & vbCrLf & vbTab
									Response.Write "	<td align='left'>"& arrayRS2(4,i2) &"</td>" & vbCrLf & vbTab
									Response.Write "	<td align='center'>"& arrayRS2(6,i2) &"</td>" & vbCrLf & vbTab
									Response.Write "<tr>"
								next
							end if
						end if
						
					%>
				</tbody>
			</table>
			<form class="hidden text-right" id="form_reproceso" method="POST">					
				<input id="num_reporte" class="hidden" name="num_reporte"/>
				<button type="submit">Rerpocesar</button>
			</form>
			</div>
		</div>
		<div id="clockCounter" name="clockCounter" class="fixedBottomRightLabel">
			La p&aacute;gina se actualizar&aacute en
			<label id="lblTimer" name="lblTimer">60</label>
			<label id="lblMedidaTiempo" name="lblMedidaTiempo">segundo</label>(s).
		</div>
		<script type="text/javascript">
			var reloadTime = MinutosRecargarPagina * 60 * 1000;
			
			<!--
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
            var modal = document.getElementById("myModal");
			$(".modal-click").click(function ()
			{
				modal.style.display = "block";
                $(".tr-moda").hide("hidden");
                $(".tr-moda-" + $(this).data("idcron")).show("hidden");
                var idcron = $(this).data("idcron");
				var status = $(this).data("status");
                $("#form_reproceso").addClass("hidden");
				if (status != "")
				{
					$("#form_reproceso").removeClass("hidden");
                    $("#num_reporte").val(idcron);
				}
			});
			$(".close").click(function ()
			{
                modal.style.display = "none";
			});
            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
			}
        </script>
	</body>


	
	<% 
			'< -- CHG-DESA-27042022-01
			'----------------------------'Setear la prioridad de los procesos dinamicamente:----------------------------
			Sub sub_procesos_prioridad_dinamica()
				dim bandera,SQL2,arrayRS2
				    bandera = 0

			SQL2 = "select  cron.id_rapport," & VbCrLf 
			SQL2 = SQL2 & "  cron.priorite, rep_detalle.id_cron, cron.id_chron, cron.in_progress, TO_char(cron.last_execution,'DD/MON/YYYY HH24:MI') as hora_creacion" & VbCrLf 
			SQL2 = SQL2 & "  from REP_CHRON cron " & VbCrLf 
			SQL2 = SQL2 & "  JOIN rep_detalle_reporte rep_detalle on cron.id_rapport = rep_detalle.id_cron   " & VbCrLf 
			SQL2 = SQL2 & "  JOIN rep_reporte rep on rep.ID_REP = rep_detalle.id_rep  " & VbCrLf 
			SQL2 = SQL2 & "  LEFT OUTER JOIN rep_reprocesos_reporte reprocesos on reprocesos.id_cron = cron.id_rapport  " & VbCrLf 
			SQL2 = SQL2 & "  where cron.active <> 0 " & VbCrLf 
			SQL2 = SQL2 & "  and trunc(cron.last_execution) = trunc(sysdate) " & VbCrLf 
			SQL2 = SQL2 & "  order by  cron.in_progress desc, hora_creacion desc " 
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
							if CStr(arrayRS2(1,i)) <> CStr(arrTmp(1,j))  Then
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
				SQL ="update rep_chron set priorite = "& prioridad &" where ID_CHRON IN (select ID_CHRON from  rep_chron cron" & VbCrLf 
				SQL = SQL & "JOIN rep_detalle_reporte rep_detalle on cron.id_rapport = rep_detalle.id_cron   " & VbCrLf
				SQL = SQL & "JOIN rep_reporte rep on rep.ID_REP = rep_detalle.id_rep " & VbCrLf
				SQL = SQL & "where cron.active <> 0" & VbCrLf
				SQL = SQL & "and trunc(cron.last_execution) = trunc(sysdate) " & VbCrLf
				SQL = SQL & "and rep_detalle.id_rep  = "& id_rep &" )" & VbCrLf
				rst.Open SQL, Connect(), 0, 1, 1
				'Response.Write  "("& id_rep & ") " & id_chron & " -" & "- > " & prioridad & "|" 
				Response.Write "<script>console.log('CHANGED: "& id_rep &" "& id_chron &" "& prioridad &"')</script>"
			End Sub
		'-----------------------------------------------------------------------------------------------------------
		' CHG-DESA-27042022-01 -- >
%>	

</html>
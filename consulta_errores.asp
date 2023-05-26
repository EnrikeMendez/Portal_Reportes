<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'modificacion de reportes
Response.Expires = 0
call check_session()
dim SQL, arrayRS, SQL_02, arrayRS2, i, rst, arrayRS3, nowDay, backDays
set rst = Server.CreateObject("ADODB.Recordset")
backDays = 1
nowDay = Now

Function NVL(str)
	if IsNull(str) then
		NVL = "" 
	else 
		NVL = str
	end if
End Function

%>
<!DOCTYPE html>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html;" charset="iso-8859-1" />
		<% call print_style() %>
		<link type="text/css" href="css/logis_style.min.css" rel="stylesheet" />
		<link href="include/logis.css" type="text/css" rel="stylesheet" />
		<script type="text/javascript" src="js/reports.min.js"></script>
		<script language="JavaScript" src="./include/tigra_tables.js"></script>
		<style type="text/css">
			.dvHoraActual{
				display:		none;
				font-size:		small;
				font-weight:	bold;
				text-align:		right;
				visibility:		collapse;
				width:			100%;
			}
		</style>
		<script type="text/javascript">
			var MinutosRecargarPagina = 5;
			function r() { setTimeout(function () { location.reload(); }, 1000); }
            var totLabel = "";
            const fecha = new Date();
            var tot = (fecha.getFullYear() - 2000) + (fecha.getMonth() + 1) + fecha.getDate() + fecha.getHours() + fecha.getMinutes() + fecha.getSeconds();
            tot = tot / 60;
            totLabel = tot.toFixed(1).toString();
        </script>
		<title>Consulta de errores</title>
	</head>
	<body onload="refresca_pagina();">
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
					<li id="imgXls" alt="Exportar" title="Exportar consulta" onclick="GeneraExcel('ConsultaErrores','select_reporte')" class="link_cursor">
						Exportar consulta
					</li>
				</ul>
				<h2>
					CONSULTA DE ERRORES
				</h2>
			</div>
		</div>

		<%
'		SQL = "select reporte.name,  error.id_reporte as id_cron, error.id_chron_error, error.tipo_error, error.log, TO_char(error.date_created,'DD/MON/YYYY HH24:MI') as hora " & VbCrLf 
'		SQL = SQL & " from rep_chron_error error, " & VbCrLf 
'		SQL = SQL & " rep_detalle_reporte reporte " & VbCrLf 
'		SQL = SQL & "  where trunc(error.date_created) = trunc(sysdate) " & VbCrLf 
'			SQL = SQL & "  	where error.date_created >= ((sysdate-1) + (8/24))  " & VbCrLf 
'			SQL = SQL & "  where (TO_CHAR(error.date_created,'dd/MM/yyyy HH24:mm')) >= (TO_CHAR(trunc(SYSDATE)-(6/24),'dd/MM/yyyy HH24:mm')) " & VbCrLf 
'		SQL = SQL & "  and error.id_reporte = reporte.id_cron " & VbCrLf 

'		SQL = SQL & "  union " & VbCrLf 
'		SQL = SQL & "  	select '<i>Reporte generado bajo demanda</i>' name,  " & VbCrLf 
'		SQL = SQL & "  	select concat(replace(replace(substr(error.log,instr(error.log,'|',-1)+1   ),' Reporte : ',''),'.xls',''), '<br>(<i>Reporte generado bajo demanda</i>)') name,  " & VbCrLf 
		'SQL = SQL & "  	select concat(replace(INITCAP(replace(substr(error.log,instr(error.log,'|',-1)+1   ),' Reporte : ','')),'.xls',''), '<br>(<i>Reporte generado bajo demanda</i>)') name,  " & VbCrLf 
'		SQL = SQL & "  	  error.id_reporte as id_cron, error.id_chron_error, error.tipo_error, error.log, TO_char(error.date_created,'DD/MON/YYYY HH24:MI') as hora  " & VbCrLf 
'		SQL = SQL & "  	from rep_chron_error error  " & VbCrLf 
'		SQL = SQL & "  	where trunc(error.date_created) = trunc(sysdate)  " & VbCrLf 
'			SQL = SQL & "  	where error.date_created >= ((sysdate-1) + (8/24))  " & VbCrLf 
'			SQL = SQL & "  	where trunc(error.date_created) >= trunc((sysdate-1) + (6/24))  " & VbCrLf 
'		SQL = SQL & "  	    and error.id_reporte not in (select distinct reporte.id_cron from rep_detalle_reporte reporte)  " & VbCrLf 

		'SQL = SQL & " order by error.date_created desc " 
'		SQL = SQL & " order by hora desc, id_cron desc "

			'El lunes se mostrarán los errores que ocurrieron desde el viernes por la tarde:
			if Weekday(nowDay,1) = 2 then
				backDays = 3
			end if

		SQL	=	"SELECT	 reporte.name			AS	name "	&	VbCrLf
		SQL	=	SQL	&	" 		,error.id_reporte		AS	id_cron "	&	VbCrLf
		SQL	=	SQL	&	"		,error.id_chron_error	AS	id_chron_error "	&	VbCrLf
		SQL	=	SQL	&	"		,error.tipo_error		AS	tipo_error "	&	VbCrLf
		SQL	=	SQL	&	"		,error.log				AS	log "	&	VbCrLf
		SQL	=	SQL	&	"		,TO_CHAR(error.date_created,'DD/MON/YYYY HH24:MI')	AS	hora "	&	VbCrLf
		SQL	=	SQL	&	"FROM	 rep_chron_error error "	&	VbCrLf
		SQL	=	SQL	&	"	INNER	JOIN	rep_detalle_reporte reporte "	&	VbCrLf
		SQL	=	SQL	&	"		 ON	error.id_reporte	=	reporte.id_cron "	&	VbCrLf
		SQL	=	SQL	&	"WHERE	 error.date_created		>=	((sysdate - " & backDays & ") + (8/24)) "	&	VbCrLf
		SQL	=	SQL	&	"UNION "	&	VbCrLf
		SQL	=	SQL	&	"SELECT	 CONCAT(REPLACE(REPLACE(SUBSTR(error.log,INSTR(error.log,'|',-1)+1   ),' Reporte : ',''),'.xls',''), '<br>(<i>Reporte generado bajo demanda</i>)')	AS	name "	&	VbCrLf
		SQL	=	SQL	&	"		,error.id_reporte		AS	id_cron "	&	VbCrLf
		SQL	=	SQL	&	"		,error.id_chron_error	AS	id_chron_error "	&	VbCrLf
		SQL	=	SQL	&	"		,error.tipo_error		AS	tipo_error "	&	VbCrLf
		SQL	=	SQL	&	"		,error.log				AS	log "	&	VbCrLf
		SQL	=	SQL	&	"		,TO_CHAR(error.date_created,'DD/MON/YYYY HH24:MI')	AS	hora "	&	VbCrLf
		SQL	=	SQL	&	"FROM	 rep_chron_error error "	&	VbCrLf
		SQL	=	SQL	&	"WHERE	 error.date_created	>=	((sysdate - " & backDays & ") + (8/24)) "	&	VbCrLf
		SQL	=	SQL	&	"	AND	 error.id_reporte	NOT IN	(SELECT	DISTINCT "	&	VbCrLf
		SQL	=	SQL	&	"											reporte.id_cron "	&	VbCrLf
		SQL	=	SQL	&	"									 FROM	rep_detalle_reporte reporte) "	&	VbCrLf
		SQL	=	SQL	&	"ORDER	 BY	hora	DESC "	&	VbCrLf
		SQL	=	SQL	&	"		,id_cron	DESC "	&	VbCrLf

'		response.Write SQL
		arrayRS = GetArrayRS(SQL)
			if not IsArray(arrayRS) then
				Response.Write "<script>MinutosRecargarPagina = totLabel;</script>"
			end if
		%>
		<center>
			<table width="98%" border="0" class="tbl-shadow">
				<tr>
					<td>
						<input id="table-buscar" type="text" class="form-control rounded-txt" placeholder="Escriba algo para filtrar" style="width: 100%;"/>
					</td>
				</tr>
			</table>
		</center>
		<br />
		<table width="100%" border="0" id="select_reporte" class="tblContent">
			<thead>
				<tr align="center">
					<th>Nombre</th>
					<th class="width-8p">ID cron</th>
					<th class="width-8p">Error</th>
					<!--<th class="width-10p">Lista Correo</th>-->
					<th>Log</th>
					<th class="width-12p">Fecha</th>
				</tr>
			</thead>
			<tbody>
			<%
			if not IsArray(arrayRS) then
				Response.Write "<tr class='center'>"
				Response.Write "	<td colspan='5' class='center'>"
				Response.Write "		No hay Errores registrados."
				Response.Write "	</td>"
				Response.Write "</tr>"
			else
				for i = 0 to UBound(arrayRS,2)
					Response.Write "<tr>" & vbCrLf & vbTab 
					if instr(arrayRS(0,i),"@") = 0 then
						Response.Write "	<td>"& arrayRS(0,i) &"</td>" & vbCrLf & vbTab 
					else
						Response.Write "	<td><i>Reporte generado bajo demanda</i></td>" & vbCrLf & vbTab 
					end if
					Response.Write "	<td align='center'> <font class='carmin'>" & arrayRS(1,i) & "</font></td>"
					Response.Write "	<td align='center'>"& arrayRS(2,i) &"</td>" & vbCrLf & vbTab 
					'Response.Write "	<td>"& arrayRS(3,i) &"</td>" & vbCrLf & vbTab 
					Response.Write "	<td>"& arrayRS(4,i) &"</td>" & vbCrLf & vbTab 
					Response.Write "	<td align='center'>"& arrayRS(5,i) &"</td>" & vbCrLf & vbTab 
					Response.Write "</tr>"
				next
			end if
			%>
			</tbody>
		</table>
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
				showTime();
                setTimeout(function () {
                    location.reload();
				}, (reloadTime));
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
			function showTime() {
				var myDate, hours, minutes, seconds, dvHoraActual, dato;
				myDate = new Date();
				hours = myDate.getHours();
				minutes = myDate.getMinutes();
				seconds = myDate.getSeconds();
				if (hours < 10) hours = 0 + hours;
				if (minutes < 10) minutes = "0" + minutes;
				if (seconds < 10) seconds = "0" + seconds;
				
				dvHoraActual = document.getElementById("HoraActual");
				dato = (hours + ":" + minutes + ":" + seconds);
				dvHoraActual.innerHTML = dato;
				setTimeout("showTime()", 1000);
			}
            function hideLoading() {
                document.getElementById("dvloading").style.display = "none";
                document.getElementById("dvloading").style.visibility = "collapse";
            }

			$TableFilter = function (id, value) {
				var rows = document.querySelectorAll(id + ' tbody tr');

				if (MinutosRecargarPagina != null) { MinutosRecargarPagina = 5; }

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
		<div id="HoraActual" name="HoraActual" class="dvHoraActual"> </div>
	</body>
</html>
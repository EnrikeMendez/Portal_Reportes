<%@ Language=VBScript %>
<% option explicit
Response.Expires = -1
%><!--#include file="include/include.asp"-->
<%
dim qa
	qa = "_qa"
	'call check_session()
	
	dim array_Talones
	dim i, num_reporte
	dim iColLoop, iRowLoop
	dim tipo, msg, lblClass
	dim SQL, sqlFiltro, rst
	dim script_include, style_include
	dim filFecha_Ini, filFecha_Fin, filMail

	lblClass = "tdLabel center"
	msg = Request.QueryString("msg")
	tipo = Request.QueryString("tipo")
	
'	if Session("internal_login") <> 2 and Session("internal_login") <> 3 then
'		if UCase(tipo) = "LTL" then
'			Response.Redirect "ltl_consulta" & qa & ".asp?msg=" & Server.URLEncode("El usuario " & Session("array_client")(3,0) & " no tiene permitido el acceso.")
'		else
'			Response.Redirect "cd_consulta" & qa & ".asp?msg=" & Server.URLEncode("El usuario " & Session("array_client")(3,0) & " no tiene permitido el acceso.")
'		end if
'	end if

	if Request.Form("etapa") = "1" then
		filMail = Request.Form("txtMail")
		filFecha_Ini = Request.Form("txtFecha_Ini")
		filFecha_Fin = Request.Form("txtFecha_Fin")
		msg = "El reporte solicitado se está generando y en breve será enviado al correo indicado."
		
		if filFecha_Ini <> "" and filFecha_Fin <> "" then
			sqlFiltro = "	AND	W.DATE_CREATED	BETWEEN	TO_DATE('" & filFecha_Ini & "','DD/MM/YYYY')	AND	TO_DATE('" & filFecha_Fin & "','DD/MM/YYYY')	" & vbCrLf
		else
			sqlFiltro = "	AND	W.DATE_CREATED	BETWEEN	TO_DATE('01/' || TO_CHAR(ADD_MONTHS(SYSDATE,-1),'MM/YYYY'))	AND	LAST_DAY(ADD_MONTHS(SYSDATE,-1))	" & vbCrLf
		end if
		
		SQL = "SELECT SEQ_CHRON.nextval FROM DUAL"
		Session("SQL") = SQL
		array_Talones = GetArrayRS(SQL)
		
		if IsArray(array_Talones) then
			num_reporte = array_Talones(0,0)
			
			SQL = "" & vbCrLf
			SQL = SQL & "INSERT INTO	REP_DETALLE_REPORTE (ID_CRON, ID_REP, DEST_MAIL, MAIL_OK, MAIL_ERROR, NAME, CLIENTE, FRECUENCIA, FILE_NAME, CARPETA, PARAM_1, PARAM_2, " & vbCrLf
			SQL = SQL & "									DAYS_DELETED, LAST_CREATED, LAST_CONF_DATE_1, LAST_CONF_DATE_2, TEST,PARAM_3, CREATED_BY, DATE_CREATED)" & vbCrLf
			SQL = SQL & "	SELECT	'" & SQLEscape(num_reporte) & "',ID_REP,'web-reports@logis.com.mx;" & filMail & "', MAIL_OK, MAIL_ERROR, NAME, CLIENTE, FRECUENCIA,  FILE_NAME, CARPETA, PARAM_1, PARAM_2, " & vbCrLf
			SQL = SQL & "			DAYS_DELETED, LAST_CREATED, TO_CHAR(TO_DATE('" & filFecha_Ini & "', 'DD/MM/YYYY')) LAST_CONF_DATE_1, TO_DATE('" & filFecha_Fin & "', 'DD/MM/YYYY') LAST_CONF_DATE_2, TEST, PARAM_3, 'REP_ANOMALIA_WEB', SYSDATE " & vbCrLf
			SQL = SQL & "	FROM	REP_DETALLE_REPORTE " & vbCrLf
			SQL = SQL & "	WHERE	ID_CRON	=	'206974' " & vbCrLf
			Session("SQL") = SQL
			set rst = Server.CreateObject("ADODB.Recordset")
			rst.Open SQL, Connect(), 0, 1, 1
			
			SQL = "" & vbCrLf
			SQL = SQL & " INSERT INTO	REP_CHRON (ID_CHRON, ID_RAPPORT, PRIORITE, TEST, ACTIVE) " & vbCrLf
			SQL = SQL & " 	VALUES (SEQ_CHRON.nextval,'" & SQLEscape(num_reporte) & "',1,0, 1) " & vbCrLf
			set rst = Server.CreateObject("ADODB.Recordset")
			rst.Open SQL, Connect(), 0, 1, 1
		end if
	else
		sqlFiltro = "	AND	W.DATE_CREATED	BETWEEN	TO_DATE('01/' || TO_CHAR(ADD_MONTHS(SYSDATE,-1),'MM/YYYY'))	AND	LAST_DAY(ADD_MONTHS(SYSDATE,-1))	" & vbCrLf
		
		SQL = " SELECT TO_CHAR(TO_DATE('01/' || TO_CHAR(ADD_MONTHS(SYSDATE,-1),'MM/YYYY')),'DD/MM/YYYY') FECHA_INI, TO_CHAR(LAST_DAY(ADD_MONTHS(SYSDATE,-1)),'DD/MM/YYYY') FECHA_FIN FROM DUAL "
		Session("SQL") = SQL
		array_Talones = GetArrayRS(SQL)
		
		if IsArray(array_Talones) then
			filFecha_Ini = array_Talones(0,0)
			filFecha_Fin = array_Talones(1,0)
		end if
	end if

	SQL = "SELECT	 WEL_CLICLEF AS NUM_CLIENTE " & vbCrLf
	SQL = SQL & "	, CLINOM AS NOMBRE_CLIENTE " & vbCrLf
	SQL = SQL & "	, TO_CHAR(TRUNC(W.DATE_CREATED, 'MM'), 'YYYY/MM/DD') AS FECHA_CREACION " & vbCrLf
	SQL = SQL & "	, TO_CHAR(WELCONS_GENERAL, 'FM0000000') || '-' || GET_CLI_ENMASCARADO(WEL_CLICLEF) AS TALON " & vbCrLf
	SQL = SQL & "	, WELIMPORTE AS IMPORTE_TALON " & vbCrLf
	SQL = SQL & "	, NVL(WELIMPORTE_DIVCLEF, 'MXN') AS IMPORTE_DIVISA " & vbCrLf
	SQL = SQL & "	, WLC_IMPORTE AS IMPORTE_SEGURO " & vbCrLf
	SQL = SQL & "	, WLMBASECALCULO " & vbCrLf
	SQL = SQL & "	, WLMCUOTAUNIDAD " & vbCrLf
	SQL = SQL & "	, WLMCUOTAMIN " & vbCrLf
	SQL = SQL & "	, WLMDESCRIPCION " & vbCrLf
	SQL = SQL & "FROM	WEB_LTL W " & vbCrLf
	SQL = SQL & "	 , ECLIENT " & vbCrLf
	SQL = SQL & "	 , WEB_LTL_CONCEPTOS " & vbCrLf
	SQL = SQL & "	 , ECONCEPTOSHOJA " & vbCrLf
	SQL = SQL & "	 , WEB_LTL_METODOS " & vbCrLf
	SQL = SQL & "WHERE	1	=	1 " & vbCrLf
	SQL = SQL & vbCrLf & sqlFiltro & " " & vbCrLf
	SQL = SQL & "	AND	WELSTATUS = 1 " & vbCrLf
	SQL = SQL & "	AND	WLCSTATUS = 1 " & vbCrLf
	SQL = SQL & "	AND	CHONUMERO = 183 " & vbCrLf
	SQL = SQL & "	AND	NVL(WLMSTATUS, 1) = 1 " & vbCrLf
	SQL = SQL & "	AND	CLICLEF = WEL_CLICLEF " & vbCrLf
	SQL = SQL & "	AND	WLC_WELCLAVE = WELCLAVE " & vbCrLf
	SQL = SQL & "	AND	WLC_CHOCLAVE = CHOCLAVE " & vbCrLf
	SQL = SQL & "	AND	WLM_WELCLAVE (+)= WLC_WELCLAVE " & vbCrLf
	SQL = SQL & "	AND	WLM_CHOCLAVE (+)= WLC_CHOCLAVE " & vbCrLf
	SQL = SQL & "ORDER	BY WEL_CLICLEF, WELCONS_GENERAL" & vbCrLf

	if Request.Form("etapa") = "1" then
		Session("SQL") = SQL
'		array_Talones = GetArrayRS(SQL)
		'response.write Replace(SQL,vbCrLf,"<br>")
	end if
	
	script_include = "<!-- main calendar program -->" & vbCrLf & _
								"<script type=""text/javascript"" src=""jscalendar/calendar.js""></script>" & vbCrLf & _
								"<!-- language for the calendar -->" & vbCrLf & _
								"<script type=""text/javascript"" src=""jscalendar/lang/calendar-es.js""></script>" & vbCrLf & _
								"<!-- the following script defines the Calendar.setup helper function, which makes" & vbCrLf & _
								"      adding a calendar a matter of 1 or 2 lines of code. -->" & vbCrLf & _
								"<script type=""text/javascript"" src=""jscalendar/calendar-setup.js""></script>" & vbCrLf & _
								"<script src=""js/DynamicOptionList.js"" type=""text/javascript"" language=""javascript""></script>" & vbCrLf

	style_include = "<!-- calendar stylesheet -->" & vbCrLf & _
					"<link rel=""stylesheet"" type=""text/css"" media=""all"" href=""jscalendar/skins/aqua/theme.css"" title=""Aqua"" />" & vbCrLf
	style_include = style_include & "<link href='css/logis_style.css' media='all' type='text/css' rel='stylesheet' />" & vbCrLf
	style_include = style_include & "<style type='text/css'>" & vbCrLf
	style_include = style_include & "	.auto-style1 { width: 56%; }" & vbCrLf
	style_include = style_include & "	.trHeader>th { font-size: 15px !important; }" & vbCrLf
	style_include = style_include & "	.tdInstrucciones { font-size: 14px !important; }" & vbCrLf
	style_include = style_include & "	.tdLabel { font-size: 13px !important; }" & vbCrLf
	style_include = style_include & "</style>" & vbCrLf

'	if UCase(tipo) = "LTL" then
'		Response.Write print_headers_nocache("Talones con seguro", "ltl", script_include, style_include, "")
'	else
'		Response.Write print_headers_nocache("Talones con seguro", "cd", script_include, style_include, "")
'	end if
%>
<html>
<head>
	<%
		call print_style()
		'Response.Write script_include
		'Response.Write style_include
		
	%>
	<style type="text/css">
		.trHeader
				{
					background-color: #223F94;
					font-family: "Roboto",sans-serif;
					font-size: 14px;
					color:#FFFFFF;
					text-align: center;
				}
		.trHeader>th
		{
			background-color: #223F94;
		}
		.trHeader>td
		{
			font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
			font-style: normal;
			font-weight: bold;
			font-size: 14px;
		}
		td{
			font-size: 12px;
		}
		.dvFormCancel
		{
			display: block;
			font-size: 14px;
			padding-top: 100pt;
			width: 400pt;
		}
		.calendar
		{
			cursor: pointer;
			background-color: #fff;
		}
		.success-bg
		{
			font-size: 11pt;
			background-color: #CCFFCC;
			font-family: Arial, sans-serif;
			font-weight: bold;
			color: #48A838 !important;
		}
		.error-bg
		{
			font-size: 10pt;
			background-color: #FFCCCC;
			font-family: Arial, sans-serif;
			font-weight: bold;
			color: #DC3545 !important;
		}
		.wn
		{
			font-size: 10pt;
			font-weight: normal;
		}
	</style>
	
	<script type="text/javascript" src="include/jscalendar/calendar.js"></script>
	<!-- language for the calendar -->
	<script type="text/javascript" src="include/jscalendar/lang/calendar-es.js"></script>
	<!-- the following script defines the Calendar.setup helper function, which makes
	      adding a calendar a matter of 1 or 2 lines of code. -->
	<script type="text/javascript" src="include/jscalendar/calendar-setup.js"></script>
	<script src="js/DynamicOptionList.js" type=""text/javascript" language="javascript"></script>
	
	<link media="screen" href="dyncalendar.css" type="text/css" rel="stylesheet">
	<script src="include/browserSniffer.js" type="text/javascript" language="javascript"></script>
	<script src="include/dyncalendar.js" type="text/javascript" language="javascript"></script>
</head>
<body>
<center>
	<div class="dvFormCancel">
<%
	if Request.QueryString("msg") <> "" then
		if inStr(Request.QueryString("msg"),"Error") then
			lblClass = "error-bg"
		else
			lblClass = "success-bg"
		end if
		%>
			<center>
				<label id="lblMsg" name="lblMsg" class="<%=lblClass%>">
					<%=Request.QueryString("msg")%>
				</label>
			</center>
			<br />
		<%
	else
		if msg <> "" then
			if inStr(msg,"Error") then
				lblClass = "error-bg"
			else
				lblClass = "success-bg"
			end if
			%>
				<center>
					<label id="lblMsg" name="lblMsg" class="<%=lblClass%>">
						<%=msg%>
					</label>
				</center>
				<br />
			<%
		end if
	end if
%>
		<form id="seg_form" name="seg_form" action="<%=asp_self%>" autocomplete="off" method="post" class="form-cancel-mas">
			<center>
				<table border="0" cellspacing="0" width="100%" align="center">
					<thead>
						<tr class="trHeader bold">
							<th>
								Reporte de talones con seguro
							</th>
						</tr>
					</thead>
					<tbody>
						<tr> 
							<td> 
								<table cellpadding="5" cellspacing="0" width="100%" style="border: solid 1px blue;">
									<tr valign="top">
										<td class="tdInstrucciones" colspan="2">
											Seleccione los criterios de consulta:
										</td>
									</tr>
									<tr valign="top">
										<td class="tdLabel">
											Fecha Inicial<font color="red">*</font>:
										</td>
										<td class="tdField">
											<input type="text" size="12" class="light" id="txtFecha_Ini" name="txtFecha_Ini" readonly value="<%=day(now)%>/<%=month(now)%>/<%=year(now)%>">
											<img src="include/dynCalendar/dynCalendar.gif" id="txtFecha_Ini_trigger" title="Date selector" alt="Date selector"  valign="top"/>
											<script type="text/javascript">
												Calendar.setup({
													inputField: "txtFecha_Ini",     // id of the input field
													ifFormat: "%d/%m/%Y",      // format of the input field
													button: "txtFecha_Ini_trigger",  // trigger for the calendar (button ID)
													//align          :    "Tl",           // alignment (defaults to "Bl")
													singleClick: true
												});
											</script>&nbsp;&nbsp;&nbsp;
										</td>
									</tr>
									<tr valign="top">
										<td class="tdLabel">
											Fecha Final<font color="red">*</font>:
										</td>
										<td class="tdField">
											<input type="text" size="12" class="light" id="txtFecha_Fin" name="txtFecha_Fin" readonly value="<%=day(now)%>/<%=month(now)%>/<%=year(now)%>">
											<img src="include/dynCalendar/dynCalendar.gif" id="txtFecha_Fin_trigger" title="Date selector" alt="Date selector"  valign="top"/>
											<script type="text/javascript">
												Calendar.setup({
													inputField: "txtFecha_Fin",     // id of the input field
													ifFormat: "%d/%m/%Y",      // format of the input field
													button: "txtFecha_Fin_trigger",  // trigger for the calendar (button ID)
													//align          :    "Tl",           // alignment (defaults to "Bl")
													singleClick: true
												});
											</script>&nbsp;&nbsp;&nbsp;
										</td>
									</tr>
									<tr>
										<td class="tdLabel">
											Correo electr&oacute;nico<font color="red">*</font>:
										</td>
										<td>
											<input type="text" class="light" id="txtMail" name="txtMail" value="<%=filMail%>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</tbody>
				</table>
				<br />
				<input type="hidden" name="etapa" id="etapa" value="1" />
				<input type="button" value="Generar Reporte" class="buttonsBlue" onclick="javascript:validarForm();" name="btnEnviar" id="btnEnviar" />
				<button onclick="location.href='/reporte_anomalia.asp'"> Regresar </button>
			</center>
		</form>
	</div>
	<br/>
	<%
		if isArray(array_Talones) then
			if UBound(array_Talones,2) > 0 and UBound(array_Talones,2) < 100 then
				%>
					<div style="width:90%;">
						<table align="center" border="1" cellpadding="2" cellspacing="0" class="tblGeneralContent form-GeneralContent" id="table_data" name="table_data" style="display: inline-table;">
							<thead>
								<tr class="titulo_trading_bold" valign="center" align="center">
									<th colspan="1">N°</th>
									<th colspan="1">Cliente</th>
									<th colspan="1">Mes de creacion</th>
									<th colspan="1">Talon</th>
									<th colspan="1">Importe Mercancia</th>
									<th colspan="1">Divisa</th>
									<th colspan="1">Importe Seguro</th>
									<th colspan="1">Base calculo (Importe Mercancia)</th>
									<th colspan="1">Cuota por unidad</th>
									<th colspan="1">Cuota Min.</th>
									<th colspan="1">Descripcion</th>
								</tr>
							</thead>
							<tbody>
								<%
									iColLoop = 0
									iRowLoop = 0
									
									For iRowLoop = 0 to UBound(array_Talones,2)
										%>
											<tr>
												<%
													For iColLoop = 0 to 10
														lblClass = ""
														
														if iColLoop = 0 or iColLoop = 2 or iColLoop = 4 or iColLoop = 5 or iColLoop = 6 then
															lblClass = "center"
														end if
														
														if iColLoop = 4 or iColLoop = 6 then
															lblClass = "right"
														end if
														%>
															<td class="<%=lblClass%>">
																<%
																	if iColLoop = 4 or iColLoop = 6 then
																		if array_Talones(iColLoop,iRowLoop) <> "" then
																			Response.Write FormatCurrency(array_Talones(iColLoop,iRowLoop))
																		end if
																	else
																		Response.Write array_Talones(iColLoop,iRowLoop)
																	end if
																%>
															</td>
														<%
													Next
												%>
											</tr>
										<%
									Next
								%>
							</tbody>
						</table>
						<%
							if isArray(array_Talones) then
								Response.write "<div style='text-align:right;'><i>"
									Response.Write UBound(array_Talones,2)
								Response.write "</i> registro(s)</div>"
							else
								Response.write "<div style='text-align:right;'><i>0</i> registro(s)</div>"
							end if
						%>
					</div>
				<%
			end if
		end if
	%>
</center>
<!--
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js"></script>
-->
<script type="text/javascript">
	function validarForm()
	{
		var dFechaIni = "",
			dFechaFin = "",
			sEmail = "";
		
		dFechaIni = document.getElementById("txtFecha_Ini").value;
		dFechaFin = document.getElementById("txtFecha_Fin").value;
		sEmail = document.getElementById("txtMail").value;
		
        var noValido = /\s/;
		if (dFechaIni == "")
		{
			alert("Debe seleccionar un rango de fechas válido.");
			document.getElementById("txtFecha_Ini").value = "";
			document.getElementById("txtFecha_Ini").focus();
		}
		else if (dFechaIni == "")
		{
			alert("Debe seleccionar un rango de fechas válido.");
			document.getElementById("txtFecha_Fin").value = "";
			document.getElementById("txtFecha_Fin").focus();
		}
        else if (sEmail == "")
		{
            alert("Debe capturar un correo electronico.");
            document.getElementById("txtMail").value = "";
            document.getElementById("txtMail").focus();
        }
		else
		{
			document.getElementById("seg_form").submit();
		}
	}
	function IsNumeric(sText)
	{
		if (!/^([0-9])*$/.test(sText))
		{
			return false;
		}

		return true;
	}
	function limpiaTxt()
	{
        document.getElementById("txtMail").value = "";
        document.getElementById("txtFecha_Ini").value = "";
        document.getElementById("txtFecha_Fin").value = ""
	}
</script>
</body>
</html>
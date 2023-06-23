<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'modificacion de reportes
Response.Expires = 0
call check_session()
dim SQL, arrayRS, SQL_02, arrayRS2, i, rst, arrayRS3
set rst = Server.CreateObject("ADODB.Recordset")

Function NVL(str)
	if IsNull(str) then
		NVL = "" 
	else 
		NVL = str
	end if
End Function

	' <- CHG-DESA-30062021-01
			dim arr1(3), arr2(9)

			'14,173,24
			arr1(0) = "ALEJANDROLE"
			arr1(1) = "JAVIERD"
			arr1(2) = "ALMALFS"

			'174
			arr2(0) = "YAZMINCC"
			arr2(1) = "EVELINGB"
			arr2(2) = "ELIZABETHBM"
			arr2(3) = "LUISFR"
			arr2(4) = "DULCELO"
			arr2(5) = "ALOURDESC"
			arr2(6) = "ALEXANDRAMM"
			arr2(7) = "LGABRIELAM"
			arr2(8) = "MLOURDESB"
	' CHG-DEA-30062021-01 ->

Select Case Request.Form("Etape")
%>			
<%
case ""
		'SQL = " select repdet.ID_CRON, repdet.NAME, rep.ID_REP, rep.name, repdet.CLIENTE, InitCap(cli.clinom)  " & VbCrLf 
		'SQL = SQL & "  , repdet.CARPETA, repdet.FILE_NAME, repdet.LAST_CREATED, repdet.MAIL_OK, repdet.MAIL_ERROR " & VbCrLf 
		'SQL = SQL & "  , repdet.FRECUENCIA, tipo.DESCRIPCION, cron.HEURES, cron.MINUTES, cron.JOURS " & vbCrLf 
		'SQL = SQL & "  , cron.MOIS, cron.JOUR_SEMAINE, cron.PRIORITE, nvl(rep.subcarpeta,' '), rep.TEMP_MENSAJE_FECHA " & VbCrLf 
		'SQL = SQL & "  , rep.TEMP_MENSAJE, rep.NUM_OF_PARAM, decode(repdet.confirmacion, '1', 'checked') " & VbCrLf 
		'SQL = SQL & "  , repdet.days_deleted " & VbCrLf 
		'SQL = SQL & " from rep_detalle_reporte repdet  " & VbCrLf 
		'SQL = SQL & "  , rep_reporte rep  " & VbCrLf 
		'SQL = SQL & "  , rep_chron cron  " & VbCrLf 
		'SQL = SQL & "  , REP_TIPO_FRECUENCIA tipo  " & VbCrLf 
		'SQL = SQL & "  , eclient cli  " & VbCrLf 
		'SQL = SQL & " where rep.ID_REP = repdet.id_rep  " & VbCrLf 
		'SQL = SQL & "  and cron.ID_RAPPORT = repdet.id_cron   " & VbCrLf 
		'SQL = SQL & "  and tipo.ID_TIPO_FREC = repdet.FRECUENCIA  " & VbCrLf 
		'SQL = SQL & "  and cli.cliclef = repdet.cliente  " & VbCrLf 
		'SQL = SQL & "  and repdet.ID_CRON = " & SQLEscape(Request.Form("id_reporte"))
	' <- CHG-DESA-30062021-01
		'for i=0 to UBound(arr1)
		'	if Session("array_user")(0,0) = arr1(i) then
		'		SQL = SQL + " and rep.id_rep in (14,173,24) "
		'		exit for
		'	end if
		'next

		'for i=0 to UBound(arr2)
		'	if Session("array_user")(0,0) = arr2(i) then
		'		SQL = SQL + " and rep.id_rep = 174 "
		'		exit for
		'	end if
		'next
' CHG-DESA-30062021-01 ->
		'arrayRS = GetArrayRS(SQL)
		'Response.Write replace(SQL, vbcrlf, "<br />")
		'Response.End 
		'if not IsArray(arrayRS) then
		'	Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Reporte no existente.")
		'end if
		
		'SQL = " select nombre_proceso from REP_REPROCESOS_REPORTE where id_cron = " & arrayRS(0,0) & " and status = 1  " & VbCrLf 
		'arrayRS3 = GetArrayRS(SQL)
		%>
<!DOCTYPE html>
		<html>
			<head>
				<link href="include/logis.css" type="text/css" rel="stylesheet" />
				<link href="css/logis_style.min.css" type="text/css" rel="stylesheet" />
				<script src="js/jquery-1.3.2.min.js"></script>
				<script src="js/main.js"></script>
				<script language="JavaScript" src="./include/tigra_tables.js"></script>
				<script type="text/javascript" src="js/reports.min.js"></script>
				<script src="js/jquery-1.3.2.min.js"></script>
				<script src="js/main.js"></script>

				<script type="text/javascript">
                    var Type;
                    var Url;
                    var Data;
                    var ContentType;
                    var DataType;
					var ProcessData;
                    $(document).ready(

                        function () {
                            showLoading();
                            tmp_ws();
                        }
					);

                    function tmp_ws() {
						const xhr = new XMLHttpRequest('select_activos');
						//var idcron = document.getElementsByName('idreporte').value.trim();
                        var idcron = '<%=Request.QueryString("reporte")%>';
                        var usr = '<%=Session("array_user")(0,0)%>';
                        const url = urlWebService + "GetModificaReporte?usuario=" + usr + "&idCron=" + idcron;

                        var someHandler = "ok";

                        xhr.onreadystatechange = function () {
                            if (xhr.readyState == XMLHttpRequest.DONE) {
                                mostrarResultado(xhr.responseText);

                            }

                        }

                        xhr.open("GET", url, true);
                        xhr.send();
					}

                    function mostrarResultado(wsResponseText) {
						var objResult = JSON.parse(wsResponseText);
                        var info = objResult.GetModificaReporteResult;
						var arrayRS3 = JSON.parse(info);
						var i = 0;
						var htmlTable = "";
						var SQL = "";
						var bandera = 0;
						//$("#tbResult").empty();
						$("#lblnumero").html(arrayRS3[0].ID_CRON);
						$("#report_name").val(arrayRS3[0].NAME);
						$("#file_name").val(arrayRS3[0].FILE_NAME);
						$("#carpeta").val(arrayRS3[0].CARPETA);
						$("#subcarpeta").html(arrayRS3[0].SUBCARPETA);
						$("#diasServidor").val(arrayRS3[0].DAYS_DELETED);
                        $("#frecuencia").val(arrayRS3[0].FRECUENCIA);
						$("#hora").val(arrayRS3[0].HEURES);
						$("#minutos").val(arrayRS3[0].MINUTES);
						$("#dia").val(arrayRS3[0].JOURS);
						$("#mes").val(arrayRS3[0].MOIS);
						$("#dia_semana").val(arrayRS3[0].JOUR_SEMAINE);
						$("#idPrioridad").val(arrayRS3[0].PRIORITE);                        
						comboTipoReporte(arrayRS3[0].ID_REP);
						comboFrecuancia(arrayRS3[0].FRECUENCIA);

						for (var a = 1; a <= 9; a++) {
                            const opcion = document.createElement("option");

							opcion.value = a
                            opcion.text = a
                            if (arrayRS3[i].PRIORITE == a) {
                                opcion.selected = true
                            }
                            document.getElementById('prioridad').appendChild(opcion);
						}
                        prioridad


						//$("#tbResult").append(htmlTable);
                        hideLoading();
					}

                    function comboTipoReporte(idreporte) {
                        const xhr = new XMLHttpRequest('select_activos');
                        //var idcron = document.getElementsByName('idreporte').value.trim();
                        var usr = '<%=Session("array_user")(0,0)%>';
                        const url = urlWebService + "GetReporte?usuario=" + usr + "&idreporte=" + idreporte;

                        var someHandler = "ok";

                        xhr.onreadystatechange = function () {
                            if (xhr.readyState == XMLHttpRequest.DONE) {
                                mostrarTipoReporte(xhr.responseText);
                            }

                        }

                        xhr.open("GET", url, true);
                        xhr.send();
					}
                    function mostrarTipoReporte(wsResponseText) {
                        var objResult = JSON.parse(wsResponseText);
                        var info = objResult.GetReporteResult;
                        var arrayRS3 = JSON.parse(info);
                        var i = 0;
                        var htmlTable = "";
                        var SQL = "";
						var bandera = 0;						
						for (i = 0; i < arrayRS3.length; i++) {
							const opcion = document.createElement("option");
							opcion.value = arrayRS3[i].ID_REP
							opcion.text = arrayRS3[i].NAME
							if (arrayRS3[i].SELECCION == "selected") {
								opcion.selected = true
							}
							document.getElementById('tipo_reporte').appendChild(opcion);
						}
					}

                    function comboFrecuancia(idfrecuencia) {
                        const xhr = new XMLHttpRequest('select_activos');
                        //var idcron = document.getElementsByName('idreporte').value.trim();
                        var usr = '<%=Session("array_user")(0,0)%>';
                        const url = urlWebService + "GetReferencia?idfrecuencia=" + idfrecuencia;

                        var someHandler = "ok";

                        xhr.onreadystatechange = function () {
                            if (xhr.readyState == XMLHttpRequest.DONE) {
                                mostrarFrecuencia(xhr.responseText);
                            }

                        }

                        xhr.open("GET", url, true);
                        xhr.send();
                    }
                    function mostrarFrecuencia(wsResponseText) {
                        var objResult = JSON.parse(wsResponseText);
                        var info = objResult.GetReferenciaResult;
                        var arrayRS3 = JSON.parse(info);
                        var i = 0;
                        var htmlTable = "";
                        var SQL = "";
                        var bandera = 0;
                        for (i = 0; i < arrayRS3.length; i++) {
                            const opcion = document.createElement("option");
                            opcion.value = arrayRS3[i].ID_TIPO_FREC
                            opcion.text = arrayRS3[i].DESCRIPCION
                            if (arrayRS3[i].SELECCION == "selected") {
                                opcion.selected = true
                            }
                            document.getElementById('frecuencia').appendChild(opcion);
                        }
                    }
                </script>
				<style>
					td{font-size:12px !important;}
					.width-344{width:344px;}
				</style>
				<link href="css/logis_style.min.css" type="text/css" rel="stylesheet" />

				<link media="screen" href="include/dyncalendar.css" type="text/css" rel="stylesheet" />
				<script src="include/browserSniffer.js" type="text/javascript" language="javascript"></script>
				<script src="include/dyncalendar.js" type="text/javascript" language="javascript"></script>
				
				<title>Modificacion reporte</title>
			</head>
			<tbody>				
								<div class="contenedorMenu">
					<div class="dvMenu">
						<ul id="menu">
							<div class="logo-logis">
								<img src="images/logo-logis-s.png" alt="Logo de Logis" height="55" />
							</div>
							<li onclick="window.location.href='menu.asp';">
								Menu
							</li>
							<li onclick="window.location.href='modif_reporte.asp';">
								Administraci&oacute;n de reportes
							</li>
						</ul>
					</div>
				</div>
				<hr />
				<center>
					<form name="modif_rep" id="modif_rep" action="<%=asp_self()%>" method="post">
						<input type="hidden" name="idreporte" value=<%=Request.Form("reporte")%>  />
						<input type="hidden" id="idPrioridad" name="idPrioridad"/>
						<table border="1" width="700">
							<tr class="trHeader">
								<th colspan="2" class="font-size-13_5">Par&aacute;metros generales</th>
							</tr>
							<tr valign="top">
								<td width="500">
									<table class="width-100p">
										<tr class="tr-padding-top-bottom-10">
											<td class="tdLabel">
												N&uacute;mero del reporte:
											</td>
											<td class="tdField" id="lblnumero">
												
											</td>
										</tr>
										<tr class="tr-padding-top-bottom-10">
											<td class="tdLabel">
												Tipo del reporte:
											</td>
											<td class="tdField">
												<select id="tipo_reporte" name="tipo_reporte" class="light width-344 height-20">
													
												</select>
											</td>
										</tr>
										<tr class="tr-padding-top-bottom-10">
											<td class="tdLabel">
												Nombre del reporte:
											</td>
											<td class="tdField">
												<input type="text" id="report_name" name="report_name" class="light width-344 height-20 required" size="40" maxlength="100"/>
											</td>
										</tr>
										<tr class="tr-padding-top-bottom-10">
											<td class="tdLabel">
												Nombre del archivo:
											</td>
											<td class="tdField">
												<input type="text" id="file_name" name="file_name" class="light width-344 height-20 required" size="40" onblur="Remplace(this.form.file_name);"  maxlength="100" />
											</td>
										</tr>
										<tr>
											<td class="padding-left-20" colspan="2">
												<div class="font-size-10 sangria-izquierda">
													- Puedes usar signos especiales:<br />
													<label class="padding-left-50">
														%P -> el rango de fecha de los datos (Mar-01-2003_to_Mar-07-2003)
													</label><br />
													<label class="padding-left-50">
														%p -> la fecha de los datos
													</label><br />
													<label class="padding-left-50">
														%M, %D, %Y -> el nombre corto del mes, el n&uacute;mero del d&iacute;a y del a&ntilde;o
													</label>
												</div>
											</td>
										</tr>
									</table>
								</td>
								<td>
									Carpeta:<br /><input type="text" class="light width-200 height-20 required" id="carpeta" name="carpeta"  onblur="Remplace(this.form.carpeta);" maxlength="30"/>
									<br /><br />
									Subcarpeta: <br />
									<label class="disabled width-200 height-20 lblBorder padding-top-5" id="subcarpeta" id="subcarpeta" name="subcarpeta">
										
									</label>
									<br />
									D&iacute;as en el servidor:<br />
									<input type="text" class="light width-200 height-20 required" id="diasServidor" name="diasServidor" id="diasServidor"  onblur="Remplace(this.form.diasServidor);" maxlength="5"/>
									<br /><br />
									Frecuencia:<br />
									<select id="frecuencia" name="frecuencia" class="light width-200 height-20">										
									</select><br /><br />
									<input type="checkbox" name="con_conf" class="" />
									Con confirmaci&oacute;n.<br /><br />
								</td>
							</tr>
							<tr class="trHeader">
								<th colspan="2" class="font-size-13_5">Captura los par&aacute;metros</th>
							</tr>
							<tr class="trHeader">
								<th colspan="2" class="font-size-13_5">Programaci&oacute;n</th>
							</tr>
							<tr>
								<td colspan="2" class="padding-top-bottom-10">
									<center>
										<table border="0" align="center">
											<tr class="trHeader center">
												<td>hora</td>
												<td>minutos</td>
												<td><a href="javascript:void(0);" onmouseover="return overlib('Define cual dia del mes<br />se va a ejecutar el reporte.<br />Ej.: 1 para que cada primero de mes se ejecuta. Vacio para todos los dias.<br />Ver la ayuda abajo.');" onmouseout="return nd();">d&iacute;a</a></td>
												<td><a href="javascript:void(0);" onmouseover="return overlib('Ver la ayuda del dia.<br />2 solo para que se ejecuta en febrero.');" onmouseout="return nd();">Mes</a></td>
												<td><a>D&iacute;a de la semana</a></td>
												<td>Prioridad</td>
											</tr>
											<tr align="center">
												<td>
													<input type="text" id="hora" name="hora" class="light height-20" size="7" />
												</td>
												<td>
													<input type="text" id="minutos" name="minutos" class="light height-20" size="7" />
												</td>
												<td>
													<input type="text" id="dia" name="dia" size="7" class="light height-20" maxlength="20"/>
												</td>
												<td>
													<input type="text" id="mes" name="mes" size="7" class="light height-20" maxlength="20"/>
												</td>
												<td>
													<input type="text" id="dia_semana" name="dia_semana" size="7" class="light height-20 width-80p" maxlength="20" />
												</td>
												<td>
													<select id="prioridad" name="prioridad" class="light height-20">
														
													</select>
												</td>
											</tr>
										</table>
									</center>
								</td>
							</tr>
							<tr class="trHeader">
								<th colspan="2" class="font-size-13_5">Mensaje del correo autom&aacute;tico</th>
							</tr>
							<tr valign="top">
								<td colspan="2" class="padding-top-bottom-10">
									<table class="width-95p">
										<tr>
											<td class="tdLabel width-25p">
												Fecha l&iacute;mite de aparici&oacute;n <i>(mm/dd/yyyy)</i>: 
											</td>
											<td class="tdField">
												<input type="text" name="TEMP_MENSAJE_FECHA" size="10" class="light height-20" align="middle" />
												<script language="JavaScript" type="text/javascript">
                                                    //<!--
                                                    if (is_ie5up || is_nav6up || is_gecko) {
                                                        Calendar = new dynCalendar('Calendar', 'CalendarCallback');
                                                        Calendar.setOffset(240, 395);
                                                    }
													//-->
                                                </script>
												<script type="text/javascript">
                                                    //<!--
                                                    // Calendar callback. When a date is clicked on the calendar
                                                    // this function is called so you can do as you want with it
                                                    function CalendarCallback(date, month, year) {
                                                        date = month + '/' + date + '/' + year;
                                                        document.modif_rep.TEMP_MENSAJE_FECHA.value = date;
                                                    }
													// -->
                                                </script>
											</td>
										</tr>
										<tr>
											<td class="tdLabel">
												Texto:
											</td>
											<td class="tdField">
												<textarea name="TEMP_MENSAJE" class="light padding-left-10 width-100p"  rows="4" cols="80"></textarea>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td class="padding-left-20 padding-top-bottom-10">
									<input type="submit" class="buttonsBlue" value="Guardar" /><br />
									<input type="hidden" name="etape" value="2" />
									<!--<input type="hidden" name="num_param" value="" />
									<input type="hidden" name="num_reporte" value="" />-->
								</td>
								<td class="padding-left-20 padding-top-bottom-10">
									<center>
										<%
											'if IsArray(arrayRS3) then 
												'if arrayRS3(0,0) = 1  AND (Session("array_user")(0,0) = "HECTORRR" OR Session("array_user")(0,0) = "CESARRP")  then 												
												if (Session("array_user")(0,0) = "HECTORRR" OR Session("array_user")(0,0) = "CESARRP")  then 										%>
													<input type="date" class="light height-20"  align="middle" name="fecha-inicio"/>
													<input type="date" class="light height-20"  align="middle" name="fecha-fin"/>
													<textarea type="text" class="light height-20"  align="middle" name="email-reproceso"> </textarea>
													<input type=button class="buttonsBlue" value="Reprocesar" onclick="function_reprocesos()" /><br />
													<input type=hidden name="reprocesos"  id="reprocesos" />										<%
												end if
											'end if
										%>
									</center>
								</td>
							</tr>
		
							<script language="javascript">
                                function function_reprocesos() {
                                    document.getElementById("reprocesos").value = "reprocesar";
                                    document.getElementById("modif_rep").submit();
                                }
                                function Remplace(expr) {
                                    var new_name = expr.value;
                                    var Forbidden_char = "\\/:*?\"\'<>|;,.~ &";
                                    var i = 0;
                                    for (var i = 0; i <= new_name.length; i++) {
                                        for (var j = 0; j < Forbidden_char.length; j++) {
                                            if (new_name.charAt(i) == Forbidden_char.charAt(j)) {
                                                new_name = new_name.substring(0, i) + '_' + new_name.substring(i + 1);
                                            }
                                        }
                                    }
                                    expr.value = new_name;
                                }
                                function ValidateForm() {
                                    var msg = "";
                                    if (document.modif_rep.file_name.value == "") { msg = "- el archivo no tiene nombre.\n" };
                                    if (document.modif_rep.report_name.value == "") { msg += "- el reporte no tiene nombre.\n" };
                                    if (document.modif_rep.carpeta.value == "") { msg += "- no hay nombre de carpeta." };
                                    if ((document.modif_rep.TEMP_MENSAJE_FECHA.value != '') && (document.modif_rep.TEMP_MENSAJE.value == "")) { msg += "- Ingresar un mensaje" };

                                    if (msg == "") return true;

                                    alert("Verifica los datos : \n" + msg);
                                    return false;
                                }
                                function check_opcion(param, op) {
                                    var error = "";
                                    for (var i = 0; i < op.length; i++) {
                                        if ((op[i].value == 0) && (param[i].value == "")) {
                                            {
                                                if (error != "") { error = error + "," + (i + 1); }
                                                else { error = error + (i + 1); }

                                            }
                                        }
                                    }
                                    if (error == "") { return 0; }
                                    else { return error; }
                                }
                            </script>
						</table>
					</form>
				</center>
			</tbody>
		</html>
		<%

case "1"
	
	if Request.Form("reprocesos") = "reprocesar" then		
		dim num_reporte, email_reproceso, dia_inicio, fin_final, mes_inicio, mes_fin, ano_inicio, ano_final, last_conf_date_1, last_conf_date_2
		num_reporte = Request.Form("num_reporte")
		if Request.Form("fecha-inicio") <> "" then 
			dia_inicio = DAY(Request.Form("fecha-inicio"))
			fin_final =  DAY(Request.Form("fecha-fin"))
			mes_inicio = Month(Request.Form("fecha-inicio"))
			mes_fin = Month(Request.Form("fecha-fin"))
			ano_inicio = Year(Request.Form("fecha-inicio"))
			ano_final = Year(Request.Form("fecha-fin"))
			email_reproceso = Request.Form("email-reproceso")
			last_conf_date_1 =  mes_inicio & "/" & dia_inicio & "/" & ano_inicio
			last_conf_date_2 = mes_fin & "/" & fin_final & "/" & ano_final
			SQL = "select SEQ_CHRON.nextval from dual"
			arrayRS = GetArrayRS(SQL)
			num_reporte = arrayRS(0,0)
			
			'num_reporte
			SQL = "insert into REP_DETALLE_REPORTE " & vbCrLf 
			SQL = SQL &" (id_cron, id_rep, dest_mail, mail_ok, mail_error, name, cliente, frecuencia, " & vbCrLf 
			SQL = SQL &" file_name, carpeta, param_1, param_2, days_deleted, last_created, last_conf_date_1, last_conf_date_2, test, " & vbCrLf 
			SQL = SQL &" param_3, created_BY, date_modified) " & vbCrLf 
			'SQL = SQL &" select " & num_reporte & ", id_rep, 'web-reports@logis.com.mx; " & email_reproceso & "' dest_mail, mail_ok, mail_error, name, cliente, frecuencia, " & vbCrLf 
			SQL = SQL &" select " & num_reporte & ", id_rep, 'desarrollo_web@logis.com.mx', 6381, 6381, name, cliente, frecuencia, " & vbCrLf 
			SQL = SQL &" file_name, carpeta, param_1, param_2, days_deleted, last_created, to_char(to_date('" & last_conf_date_1 & "', 'mm/dd/yyyy')) last_conf_date_1, to_date('" & last_conf_date_2 & "', 'mm/dd/yyyy') last_conf_date_2, test, " & vbCrLf 
			SQL = SQL &" param_3, created_BY, date_modified from REP_DETALLE_REPORTE where ID_CRON = '" & SQLEscape(Request.Form("num_reporte")) & "' "   
			Response.Write SQL
			rst.Open SQL, Connect(), 0, 1, 1
			response.end
		END IF
		'Reprocesar insert
			SQL = "insert into rep_chron (id_chron, id_rapport, priorite, test, active) " & VbCrLf 
			SQL = SQL & " values (SEQ_CHRON.nextval, '" & SQLEscape(num_reporte) & "', 1,0, 1) "
			Response.Write SQL
			'rst.Open SQL, Connect(), 0, 1, 1
			response.end
	ELSE  
		'modificacion de los datos en la base...
		dim params
		'actualisacion datos de reporte :
		SQL = " update REP_DETALLE_REPORTE " 
		SQL = SQL &" set ID_REP = '" & SQLEscape(Request.Form("tipo_reporte")) & "' " & vbCrLf 
		SQL = SQL &" , NAME ='" & SQLEscape(Request.Form("report_name")) & "' " & vbCrLf 
		SQL = SQL &" , FILE_NAME = '" & SQLEscape(Request.Form("file_name")) & "' " & vbCrLf 
		SQL = SQL &" , CARPETA = '" & SQLEscape(Request.Form("carpeta")) & "' " & vbCrLf 
		SQL = SQL &" , FRECUENCIA = '" & SQLEscape(Request.Form("frecuencia")) & "' " & vbCrLf 
		SQL = SQL &" , CONFIRMACION = '" & SQLEscape(Request.Form("con_conf")) & "' " & vbCrLf 

		'< JEMV
			SQL = SQL &" , DAYS_DELETED = '" & SQLEscape(Request.Form("diasServidor")) & "' " & vbCrLf 
		' JEMV >
		
		SQL = SQL & ", param_1 = '" & trim(SQLEscape(Request.Form("param_1"))) & "'  " & vbCrLf 
		SQL = SQL & ", param_2 = '" & trim(SQLEscape(Request.Form("param_2"))) & "'  " & vbCrLf 
		SQL = SQL & ", param_3 = '" & trim(SQLEscape(Request.Form("param_3"))) & "'  " & vbCrLf 
		SQL = SQL & ", param_4 = '" & trim(SQLEscape(Request.Form("param_4"))) & "'  " & vbCrLf 
		
		SQL = SQL &", MODIFIED_BY = '" &  Session("array_user")(0,0) & "' " & vbCrLf 	
		SQL = SQL &", DATE_MODIFIED = sysdate " & vbCrLf 
		SQL = SQL &" where ID_CRON = '" & SQLEscape(Request.Form("num_reporte")) & "' "
		RESPONSE.Write SQL
		rst.Open SQL, Connect(), 0, 1, 1
		
		'actualizacion de la programacion (Cron)
		
		SQL = "update REP_CHRON " & VbCrLf 
		SQL = SQL & " set HEURES = '" & SQLEscape(Request.Form("hora")) & "' " & VbCrLf 
		SQL = SQL & " , MINUTES = '" & SQLEscape(Request.Form("minutos")) & "' " & VbCrLf 
		SQL = SQL & " , JOURS = '" & SQLEscape(Request.Form("dia")) & "' " & VbCrLf 
		SQL = SQL & " , MOIS = '" & SQLEscape(Request.Form("mes")) & "' " & VbCrLf 
		SQL = SQL & " , JOUR_SEMAINE = '" & SQLEscape(Request.Form("dia_semana")) & "' " & VbCrLf 
		SQL = SQL & " , PRIORITE = '" & SQLEscape(Request.Form("prioridad")) & "' " & VbCrLf 
		SQL = SQL & " , LAST_EXECUTION = null " & VbCrLf 
		SQL = SQL & " where ID_RAPPORT = '" & SQLEscape(Request.Form("num_reporte")) & "' "
		
		'Response.Write SQL
		rst.Open SQL, Connect(), 0, 1, 1
		
		'modificacion del mensaje del reporte.
		SQL = "update REP_REPORTE " & VbCrLf 
		SQL = SQL & " set TEMP_MENSAJE = '" & SQLEscape(Request.Form("TEMP_MENSAJE")) & "' " & VbCrLf 
		SQL = SQL & " , TEMP_MENSAJE_FECHA = to_date('" & SQLEscape(Request.Form("TEMP_MENSAJE_FECHA")) & "', 'mm/dd/yyyy') " & VbCrLf 
		SQL = SQL & " where ID_REP = '" & SQLEscape(Request.Form("tipo_reporte")) & "' "
		RESPONSE.Write SQL
		response.end
		rst.Open SQL, Connect(), 0, 1, 1
		
    end if
	Response.Redirect "menu.asp?msg=" & Server.URLEncode ("Reporte modificado.")
	
case "3"
	
	'desactivar o reactivar un reporte
	if Request.Form("accion") = "desactivar" then
		SQL = "update rep_chron set active = 0 where id_rapport = '"& Request.Form("id_reporte") &"'"
		rst.Open SQL, Connect(), 0, 1, 1
		Response.Redirect "menu.asp?msg=" & Server.URLEncode ("Reporte desactivado.")
	elseif Request.Form("accion") = "reactivar" then
	'activar de nuevo
		SQL = "update rep_chron set active = 1, last_execution = null where id_rapport = '"& Request.Form("id_reporte") &"'"
		rst.Open SQL, Connect(), 0, 1, 1
		Response.Redirect "menu.asp?msg=" & Server.URLEncode ("Reporte reactivado.")
	end if
end select

%>
	</body>
</html>
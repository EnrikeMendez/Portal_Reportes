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
	' CHG-DESA-30062021-01 ->

Select Case Request.Form("Etape")
	Case ""

%>
<!DOCTYPE html>
		<html>
		<head>
			<!-- <- CHG-DESA-30062021-01  -->
			<meta http-equiv="Content-Type" content="text/html;" charset="iso-8859-1" />
			<% call print_style() %>
			<!-- CHG-DESA-30062021-01 ->  -->
			<link href="include/logis.css" type="text/css" rel="stylesheet" />
			<link href="css/logis_style.min.css" type="text/css" rel="stylesheet" />
			<script language="JavaScript" src="./include/tigra_tables.js"></script>
			<script type="text/javascript" src="js/reports.min.js"></script>
			<script type="text/javascript">
				
				function FilterTableBy2Params() {
					var input = document.getElementById("table-buscar");
					var combo = document.getElementById("select_activos");
					var filterTxt = input.value.toUpperCase();
					var filterCmb = combo.value.toUpperCase();
					var table = document.getElementById("select_reporte");
					var tr = table.getElementsByTagName("tr");
					var filterType = 0;
					var newClass = "";
                    var bgColor = 0;

					try {
						showLoading();

						tr = table.getElementsByClassName("f");
						if (filterTxt != "" && filterCmb != "") {
							filterType = 1;
						}
						else {
							if (filterTxt != "") {
								filterType = 2;
							}
							if (filterCmb != "") {
								filterType = 3;
							}
						}

						for (var i = 0; i < tr.length; i++) {
							var rowContent = tr[i].innerHTML.toUpperCase();
                            if (bgColor % 2 == 0) { newClass = "tr-odd";}
                            else { newClass = "tr-even"; }

							switch (filterType) {
								case 1:
									if (rowContent.indexOf(filterTxt) != -1 && rowContent.indexOf(filterCmb) != -1) {
										tr[i].style.display = "";
                                        tr[i].setAttribute('class', "f " + newClass);
                                        bgColor++;
									}
									else {
										tr[i].style.display = "none";
										tr[i].setAttribute('class', "f delC");
									}
									break;
								case 2:
									if (rowContent.indexOf(filterTxt) == -1) {
										tr[i].style.display = "none";
										tr[i].setAttribute('class', "f delC");
									}
									else {
										tr[i].style.display = "";
                                        tr[i].setAttribute('class', "f " + newClass);
                                        bgColor++;
									}
									break;
								case 3:
									if (rowContent.indexOf(filterCmb) == -1) {
										tr[i].style.display = "none";
										tr[i].setAttribute('class', "f delC");
									}
									else {
										tr[i].style.display = "";
                                        tr[i].setAttribute('class', "f " + newClass);
										bgColor++;
									}
									break;
								default:
									tr[i].style.display = "";
                                    tr[i].setAttribute('class', "f " + newClass);
                                    bgColor++;
									break;
							}
						}
					}
					catch { }
					finally {
						hideLoading();
					}
				}
				function showActive() {
					document.getElementById("select_activos").value = "Desactivar";
					FilterTableBy2Params();
					hideLoading();
				}

				function showLoading() {
                    document.getElementById("dvloading").style.display = "";
                    document.getElementById("dvloading").style.visibility = "visible";
				}
				function hideLoading() {
                    document.getElementById("dvloading").style.display = "none";
                    document.getElementById("dvloading").style.visibility = "collapse";
                }
            </script>
			<!-- CHG-DESA-30062021-01 ->  -->
			<title>Administracion de reportes</title>
		</head>
		<body onload="showActive();">
			<!-- <-	CHG-DESA-30062021-01   --
			<div id="dvloading" style="display:block!important;visibility:visible!important;">
				<center>Procesando </center>
				<center><img alt=". . ." id="imgPuntos" src="images/puntosSuspensivos.gif" /></center>
			</div>
			<!-- CHG-DESA-30062021-01 ->  -->
		<%
		
		
		SQL = "select repdet.ID_CRON, trim(nvl(repdet.NAME,'')), rep.id_rep, trim(nvl(rep.name,'')), repdet.CLIENTE, InitCap(cli.clinom) " & VbCrLf 
		SQL = SQL & " , repdet.FRECUENCIA, tipo.DESCRIPCION, repdet.mail_ok, nvl(cron.active, 0), repdet.mail_error" & VbCrLf 
' <- CHG-DESA-30062021-01
			SQL = SQL & " , cli.clirfc" & VbCrLf 
			SQL = SQL & " , case when  rep.id_rep in (80,98,106,108,110,114,117,118,126,130,142,159,160, 169,171,174,179,175,176,183,186,199,201,221,228,236,240,242,244,248,249,260,263,287,288,290) then 'DISTRIBUCION' else 'COEX' end Area_Negocio" & VbCrLf 
			SQL = SQL & " , cron.priorite as Prioridad" & VbCrLf 
			SQL = SQL & " , tipo.DESCRIPCION as frecuencia_desc, repdet.days_deleted as dias_servidor" & VbCrLf 
			SQL = SQL & " , cron.jours AS DIA_MES, cron.jour_semaine AS DIA_SEMANA" & VbCrLf 
			SQL = SQL & " , cron.heures AS HORA" & VbCrLf 
			SQL = SQL & " , cron.minutes AS MINUTO" & VbCrLf 
			SQL = SQL & " ,repdet.CLIENTE || ' ' || repdet.NAME AS Num_Nom" & VbCrLf 
			SQL = SQL & " , repdet.param_1, repdet.param_2,repdet.param_3,repdet.param_4" & VbCrLf 
			SQL = SQL & " , repdet.created_by, repdet.date_created, repdet.modified_by, repdet.date_modified" & VbCrLf 
			SQL = SQL & " , rep.COMMAND AS COMMAND" & VbCrLf 
'CHG-DESA-30062021-01->
		SQL = SQL & " from rep_detalle_reporte repdet " & VbCrLf 
		SQL = SQL & " , rep_reporte rep " & VbCrLf 
		SQL = SQL & " , rep_chron cron " & VbCrLf 
		SQL = SQL & " , REP_TIPO_FRECUENCIA tipo " & VbCrLf 
		SQL = SQL & " , eclient cli " & VbCrLf 
		SQL = SQL & " where rep.ID_REP = repdet.id_rep " & VbCrLf 
		SQL = SQL & " and cron.ID_RAPPORT(+) = repdet.id_cron  " & VbCrLf 
		SQL = SQL & " and tipo.ID_TIPO_FREC = repdet.FRECUENCIA " & VbCrLf 
		SQL = SQL & " and cli.cliclef = repdet.cliente " & VbCrLf 
		'SQL = SQL & " and repdet.test = 0 " & VbCrLf 
		'SQL = SQL & " and cron.active = 1 " & VbCrLf 

' <- CHG-DESA-30062021-01
		for i=0 to UBound(arr1)
			if Session("array_user")(0,0) = arr1(i) then
				SQL = SQL + " and rep.id_rep in (14,173,24) "
				exit for
			end if
		next

		for i=0 to UBound(arr2)
			if Session("array_user")(0,0) = arr2(i) then
				SQL = SQL + " and rep.id_rep = 174 "
				exit for
			end if
		next
' CHG-DESA-30062021-01 ->


		SQL = SQL & " order by 2 "

		arrayRS = GetArrayRS(SQL)

'			response.write Replace(SQL,VbCrLf,"<br>")
'			response.End
		
		if not IsArray(arrayRS) then 
			Response.Write "No hay reportes registrados."
			Response.End 
		end if
		
		'affichage du popup pour la fonction filtre_col
		call print_popup()
		%>
		<!-- <- CHG-DESA-30062021-01  -->
		<div class="contenedorMenu">
			<div class="dvMenu">
				<ul id="menu">
					<div class="logo-logis">
						<img src="images/logo-logis-s.png" style="height:50px;" />
					</div>
					<li onclick="window.location.href='menu.asp';" class="link_cursor">
						Inicio
					</li>
					<li id="imgXls" alt="Exportar" title="Exportar consulta" onclick="GenerarExcel('ConsultaGeneralReportes','select_activos','select_reporte')" class="link_cursor">
						Exportar consulta
					</li>
				</ul>
				<h2>
					ADMINISTRACI&Oacute;N DE REPORTES
				</h2>
			</div>
		</div>
			<center>
				<%if Request("msg") <> "" then	Response.Write "<table width='100%'><tr><td align='center' colspan='2'><font class='lblMSG red error' size='2'>" & Request("msg") & "</font></td></tr></table><br/>"%>
				<table width="98%" border="0" class="tbl-shadow">
					<tr>
						<td>
							<input id="table-buscar" type="text" class="form-control rounded-txt" placeholder="Escriba algo para filtrar" style="width: 100%;"/>
						</td>
					</tr>
					<tr>
						<td class="width-15p">
							<label>
								Estatus:
							</label>
							<select name="cars" id="select_activos" class="form-control rounded-cmb">
								<option value="" selected>Todos</option>
								<option value="Desactivar">Activos</option>
								<option value="Reactivar">No activos</option>
							</select>
						</td>
					</tr>
				</table>
			</center>
			<form name="frmExcel" id="frmExcel" action="<%=asp_self()%>" method="post">
				<input type="hidden" name="etape" value="4" id="hdnEtape_4" />
				<img id="imgXlsWait" alt="Generando" title="Generando archivo . . ." src="../images/excel.gif" class="logo_excel_disabled" style="visibility:collapse;"/>
			</form>
			<br />
			<table width="100%" border="0" id="select_reporte" class="tblContent">
				<thead>
					<tr align="center" id="trHeader" class="trHeader">
						<td rowspan="2" id="thNo">No.</td>
						<td rowspan="2" id="thName">Nombre</td>
						<td rowspan="2" id="thArea" class="trHeader">Area Negocio</td>
						<td rowspan="2" id="thPrio">Prioridad</td>
						<td rowspan="2" id="thType">Tipo reporte</td>
						<td colspan="3" class="delC" id="thListC">Lista Correo</td>
						<td rowspan="2" id="thCus">Cliente</td>
						<td rowspan="2" id="thFre">Frecuencia</td>
						<td rowspan="2" class="delC">Accion</td>
						<td rowspan="2" id="tdDS">D&iacute;as en el Servidor</td>
						<td rowspan="2" id="tdDM" title="D&iacute;as del mes en que se genera el reporte">D&iacute;as del mes</td>
						<td rowspan="2" id="tdDSe" title="D&iacute;as de la semana en que se genera el reporte">D&iacute;as de la semana</td>
						<td rowspan="2" id="tdH" name="td3rows" title="Horas del d&iacute;a en que se ejecuta el proceso">Hora (s)</td>
						<td rowspan="2" id="tdM" title="Minutos dentro de cada hora programada en que se ejecuta el proceso">Minuto (s)</td>
						<td rowspan="2" id="tdUC">Usuario creaci&oacute;n</td>
						<td rowspan="2" id="tdFC">Fecha creaci&oacute;n</td>
						<td rowspan="2" id="tdUM">Usuario modificaci&oacute;n</td>
						<td rowspan="2" id="tdFM">Fecha modificaci&oacute;n</td>
						<td rowspan="2" id="tdP1">Param_1</td>
						<td rowspan="2" id="tdP2">Param_2</td>
						<td rowspan="2" id="tdP3">Param_3</td>
						<td rowspan="2" id="tdP4">Param_4</td>
						<td rowspan="2" id="thCmd">Command</td>
					</tr>
					<tr class="delC trHeader" align="center">
						<td class="trHeader delC">Ver</td>
						<td class="trHeader delC">Normal</td>
						<td class="trHeader delC">Error</td>
					</tr>
			<!-- CHG-DESA-30062021-01 ->  -->
		</thead>
		<%
		for i = 0 to UBound(arrayRS,2)
			Response.Write "<tr class='f'> "
			
			' No.
			Response.Write "<td>"& arrayRS(0,i) &"</td>" & vbCrLf & vbTab 
			
			' Nombre
			Response.Write "<td>"
			if arrayRS(9,i) = "0" then
				Response.Write "<font color='red'>" & arrayRS(1,i) & "</font></td>"
			else 
				Response.Write arrayRS(1,i) & "</td>"
			end if 
			
			' <- CHG-DESA-30062021-01
			' Area Negocio
			Response.Write "<td>" & arrayRS(12,i) & "</td>"
			
			' Prioridad
			Response.Write "<td>" & arrayRS(13,i) & "</td>"
			
			' Tipo Reporte
			Response.Write vbCrLf & vbTab

			' Lista Correo -> Ver
			Response.Write "<td align=""left"">" & JSescape(arrayRS(2,i)) & " - " & JSescape(arrayRS(3,i)) & "</td>" & vbCrLf & vbTab
			' Lista Correo -> Normal
			Response.Write "<td class=""delC"" align=""center""><a href=""javascript:ver_lista('"& arrayRS(8,i) &"','" & arrayRS(0,i) & "','" & arrayRS(4,i) & "');""; name='nVer'>Ver</a></td>" & vbCrLf & vbTab
			' Lista Correo -> Error
			Response.Write "<td class=""delC"" style=""font-size: 9.5px;""><a title=""Modificar la lista de correos"" href=""javascript:modif_list("& arrayRS(8,i) &", "& arrayRS(4,i) &");"">Mod._norm.</a></td>" & vbCrLf & vbTab
			
			' Cliente
			Response.Write "<td class=""delC"" style=""font-size: 9.5px;""><a title=""Modificar la lista de correos en caso de error"" href=""javascript:modif_list("& arrayRS(10,i) &", "& arrayRS(4,i) &");"">Mod._err</a></td>" & vbCrLf & vbTab
			
			' Frecuencia
			Response.Write "<td align=""left"">" & JSescape(arrayRS(4,i)) & " - " & arrayRS(5,i) & "</td>" & vbCrLf & vbTab
			
			' Accion
			Response.Write "<td align=""left"">" & JSescape(arrayRS(6,i)) & " - " & arrayRS(7,i) & "</td>" & vbCrLf & vbTab
			Response.Write "<td align=""center"" class=""delC""><a href=""javascript:modif_reporte("& arrayRS(0,i) &", 'mod');"">Modificar</a>_|_"
			'CHG-DESA-30062021-01 ->

			if arrayRS(9,i) = "1" then
				Response.Write "<a href=""javascript:modif_reporte("& arrayRS(0,i) &", 'desactivar');"">Desactivar</a></td>" & vbCrLf
			else
				Response.Write "<a href=""javascript:modif_reporte("& arrayRS(0,i) &", 'reactivar');"">Reactivar</a></td>" & vbCrLf
			end if

			'<- CHG-DESA-30062021-01
			' Dias en el servidor
			Response.Write "<td>" & arrayRS(15,i) & "</td>"
			
			' Dias del mes
			Response.Write "<td>" & arrayRS(16,i) & "</td>"
			
			' Dias de la semana
			Response.Write "<td>" 
			if InStr(1,arrayRS(17,i),"-") > 0 then
				Response.Write "'" & arrayRS(17,i)
			else
				Response.Write arrayRS(17,i)
			end if
			Response.Write "</td>"

			' Hora(s)
			Response.Write "<td>" & arrayRS(18,i) & "</td>"
			
			' Minuto(s)
			Response.Write "<td>" & arrayRS(19,i) & "</td>"
			
			' Usuario creacion
			Response.Write "<td>" & arrayRS(25,i) & "</td>"
			
			' Fecha creacion
			Response.Write "<td>" & arrayRS(26,i) & "</td>"
			
			' Usuario modificacion
			Response.Write "<td>" & arrayRS(27,i) & "</td>"
			
			' Fecha modificacion
			Response.Write "<td>" & arrayRS(28,i) & "</td>"
			
			' Param_1
			if replace(trim(nvl(arrayRS(21,i)))," ","") = "" then
				Response.Write "<td></td>"
			else
				Response.Write "<td>'" & arrayRS(21,i) & "</td>"
			end if
			
			' Param_2
			if replace(trim(nvl(arrayRS(22,i)))," ","") = "" then
				Response.Write "<td></td>"
			else
				Response.Write "<td>'" & arrayRS(22,i) & "</td>"
			end if
			
			' Param_3
			if replace(trim(nvl(arrayRS(23,i)))," ","") = "" then
				Response.Write "<td></td>"
			else
				Response.Write "<td>'" & arrayRS(23,i) & "</td>"
			end if

			' Param_4
			if replace(trim(nvl(arrayRS(24,i)))," ","") = "" then
				Response.Write "<td></td>"
			else
				Response.Write "<td>'" & arrayRS(24,i) & "</td>"
			end if

			' Command
			Response.Write "<td>" & trim(nvl(arrayRS(29,i))) & "</td>"
			'CHG-DESA-30062021-01 ->

			Response.Write "</tr>" & vbCrLf 
		next
		%>
		<script language="javascript">
            //<!--
			function modif_reporte(id_rep, accion) {
				document.modif_rep.id_reporte.value = id_rep;
				document.modif_rep.accion.value = accion;
				if ((accion == 'desactivar') || (accion == 'reactivar')) {
					if (confirm('� Esta seguro de ' + accion + ' el reporte no. ' + id_rep + ' ?')) {
						document.modif_rep.etape.value = 3;
						document.modif_rep.submit();
					};
				}
				else {
					document.modif_rep.etape.value = 1;
					document.modif_rep.submit();
				}
			}

			function modif_list(id_list, id_client) {
				document.modif_list.mail_list.value = id_list;
				document.modif_list.id_client.value = id_client;
				document.modif_list.submit();
			}

			//<- CHG-DESA-30062021-01
			function ver_lista(lista, Num, Cli) {
				localStorage.setItem('Cli', Cli);
				localStorage.setItem('sURI_list', "ver_lista.asp?liste=" + lista + "&Num=" + Num);
				localStorage.setItem('Id_Cron', Num);
                localStorage.setItem('pop', 1);
                window.showModalDialog("ver_lista.asp?liste=" + lista + "&Num=" + Num, "Lista_contactos", "toolbar=no, location=no, directories=no, status=yes, scrollbars=yes, resizable=no, copyhistory=no, width=765, height=444, left=333, top=111, center=yes");
			}
            //CHG-DESA-30062021-01 ->
		//-->
        </script>
		<form name="modif_rep" action="<%=asp_self()%>" method="post">
			<input type="hidden" name="etape" value="" />
			<input type="hidden" name="id_reporte" value="" />
			<input type="hidden" name="accion" value="" />
		</form>
		<form name="modif_list" action="mail_modif.asp" method="post">
			<input type="hidden" name="mail_list" value="" />
			<input type="hidden" name="id_client" value="" />
		</form>
		<script language="JavaScript">
			tigra_tables('select_reporte', 4, 0, '#ffffff', '#ffffcc', '#ffcc66', '#cccccc');
        </script>
		</table>
		<script>
			var select_activos = document.getElementById("select_activos");
			// <-	CHG-DESA-30062021-01
			document.querySelector("#table-buscar").onkeyup = function()
			{
				
			$TableFilter("#select_reporte", this.value);
			/*
				if(this.value.length > 3)
				{
					//$TableFilter("#select_reporte", this.value);
					FilterTableBy2Params();
				}
			*/
			}
				
			select_activos.addEventListener("change", function()
			{
				showLoading();
				//$TableFilter("#select_reporte", this.value);
				FilterTableBy2Params();

			});
			//CHG-DESA-30062021-01	->
		
			$TableFilter = function(id, value)
			{
				var rows = document.querySelectorAll(id + ' tbody tr');
				var estatus = document.getElementById("select_activos").value.toLowerCase().trim();
			
				for(var i = 0; i < rows.length; i++){
					var showRow = false;
				
					var row = rows[i];
					row.style.display = 'none';
				
					for(var x = 0; x < row.childElementCount; x++){
						if(row.children[x].textContent.toLowerCase().indexOf(value.toLowerCase().trim()) > -1 && row.children[10].textContent.toLowerCase().indexOf(estatus) > -1){
							showRow = true;
							break;
						}
					}
				
					if(showRow){
						row.style.display = null;
					}
				}
			}
		</script>
<%
case "1"
		SQL = " select repdet.ID_CRON, repdet.NAME, rep.ID_REP, rep.name, repdet.CLIENTE, InitCap(cli.clinom)  " & VbCrLf 
		SQL = SQL & "  , repdet.CARPETA, repdet.FILE_NAME, repdet.LAST_CREATED, repdet.MAIL_OK, repdet.MAIL_ERROR " & VbCrLf 
		SQL = SQL & "  , repdet.FRECUENCIA, tipo.DESCRIPCION, cron.HEURES, cron.MINUTES, cron.JOURS " & vbCrLf 
		SQL = SQL & "  , cron.MOIS, cron.JOUR_SEMAINE, cron.PRIORITE, nvl(rep.subcarpeta,' '), rep.TEMP_MENSAJE_FECHA " & VbCrLf 
		SQL = SQL & "  , rep.TEMP_MENSAJE, rep.NUM_OF_PARAM, decode(repdet.confirmacion, '1', 'checked') " & VbCrLf 
		SQL = SQL & "  , repdet.days_deleted " & VbCrLf 
		SQL = SQL & " from rep_detalle_reporte repdet  " & VbCrLf 
		SQL = SQL & "  , rep_reporte rep  " & VbCrLf 
		SQL = SQL & "  , rep_chron cron  " & VbCrLf 
		SQL = SQL & "  , REP_TIPO_FRECUENCIA tipo  " & VbCrLf 
		SQL = SQL & "  , eclient cli  " & VbCrLf 
		SQL = SQL & " where rep.ID_REP = repdet.id_rep  " & VbCrLf 
		SQL = SQL & "  and cron.ID_RAPPORT = repdet.id_cron   " & VbCrLf 
		SQL = SQL & "  and tipo.ID_TIPO_FREC = repdet.FRECUENCIA  " & VbCrLf 
		SQL = SQL & "  and cli.cliclef = repdet.cliente  " & VbCrLf 
		SQL = SQL & "  and repdet.ID_CRON = " & SQLEscape(Request.Form("id_reporte"))
	' <- CHG-DESA-30062021-01
		for i=0 to UBound(arr1)
			if Session("array_user")(0,0) = arr1(i) then
				SQL = SQL + " and rep.id_rep in (14,173,24) "
				exit for
			end if
		next

		for i=0 to UBound(arr2)
			if Session("array_user")(0,0) = arr2(i) then
				SQL = SQL + " and rep.id_rep = 174 "
				exit for
			end if
		next
' CHG-DESA-30062021-01 ->
		arrayRS = GetArrayRS(SQL)
'		Response.Write replace(SQL, vbcrlf, "<br />")
'		Response.End 
		if not IsArray(arrayRS) then
			Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Reporte no existente.")
		end if
		
		SQL = " select nombre_proceso from REP_REPROCESOS_REPORTE where id_cron = " & arrayRS(0,0) & " and status = 1  " & VbCrLf 
		arrayRS3 = GetArrayRS(SQL)
		%>
		<html>
			<head>
				<!-- <- CHG-DESA-30062021-01  -->
				<meta http-equiv="Content-Type" content="text/html;" charset="iso-8859-1" />
				<!-- CHG-DESA-30062021-01 ->  -->
				<script type="text/javascript" src="js/jquery-1.3.2.min.js"></script>
				<script type="text/javascript" src="js/jquery.validate.min.js"></script>
				<script type="text/javascript">
					$().ready(function () {
						$("#modif_rep").validate();
						jQuery.extend(jQuery.validator.messages, {
							required: "Campo obligatorio."
						});
					});
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
			<body>
				<%call print_style()
					'affichage du popup pour la fonction filtre_col
					call print_popup()%>
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
											<td class="tdField">
												<%=arrayRS(0,0)%>
											</td>
										</tr>
										<tr class="tr-padding-top-bottom-10">
											<td class="tdLabel">
												Tipo del reporte:
											</td>
											<td class="tdField">
												<select name="tipo_reporte" class="light width-344 height-20">
													<%
														SQL_02 = " select id_rep, name, decode(id_rep, "& arrayRS(2,0) &", 'selected') " & VbCrLf 
														SQL_02 = SQL_02 & "  from rep_reporte " & VbCrLf 
														' <- CHG-DESA-30062021-01
														SQL_02 = SQL_02 + " WHERE 1=1 "
														for i=0 to UBound(arr1)
															if Session("array_user")(0,0) = arr1(i) then
																SQL_02 = SQL_02 + " and id_rep in (14,173,24) "
																exit for
															end if
														next

														for i=0 to UBound(arr2)
															if Session("array_user")(0,0) = arr2(i) then
																SQL_02 = SQL_02 + " and id_rep = 174 "
																exit for
															end if
														next
														' CHG-DESA-30062021-01 ->
														SQL_02 = SQL_02 & "  order by 2"
										'				response.Write SQL_02
														arrayRS2 = GetArrayRS(SQL_02)

														for i=0 to UBound(arrayRS2,2)
															' <- CHG-DESA-30062021-01
															Response.Write vbTab & vbTab &"<option value=" & arrayRS2(0,i) & " " &  arrayRS2(2,i) & ">" & arrayRS2(1,i)  & "</option>"
															'CHG-DESA-30062021-01 ->
														next
													%>	
												</select>
											</td>
										</tr>
										<tr class="tr-padding-top-bottom-10">
											<td class="tdLabel">
												Nombre del reporte:
											</td>
											<td class="tdField">
												<input type="text" name="report_name" class="light width-344 height-20 required" size="40" maxlength="100" value="<%=HTMLEscape(arrayRS(1,0))%>" />
											</td>
										</tr>
										<tr class="tr-padding-top-bottom-10">
											<td class="tdLabel">
												Nombre del archivo:
											</td>
											<td class="tdField">
												<input type="text" name="file_name" class="light width-344 height-20 required" size="40" onblur="Remplace(this.form.file_name);"  maxlength="100"  value="<%=HTMLEscape(arrayRS(7,0))%>" />
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
									Carpeta:<br /><input type="text" class="light width-200 height-20 required" name="carpeta" value="<%=arrayRS(6,0)%>" onblur="Remplace(this.form.carpeta);" maxlength="30"/>
									<br /><br />
									Subcarpeta: <br />
									<label class="disabled width-200 height-20 lblBorder padding-top-5" name="subcarpeta">
										<%=arrayRS(19,0)%>
									</label>
									<br />
									D&iacute;as en el servidor:<br />
									<input type="text" class="light width-200 height-20 required" name="diasServidor" id="diasServidor" value="<%=arrayRS(24,0)%>" onblur="Remplace(this.form.diasServidor);" maxlength="5"/>
									<br /><br />
									Frecuencia:<br />
									<select name="frecuencia" class="light width-200 height-20">
										<%
											SQL_02 = " select tipo.ID_TIPO_FREC, tipo.DESCRIPCION, decode(tipo.ID_TIPO_FREC, "& arrayRS(11,0) &", 'selected')  " & VbCrLf 
											SQL_02 = SQL_02 & "   from REP_TIPO_FRECUENCIA tipo " & VbCrLf 
											SQL_02 = SQL_02 & "   order by 2 "
											arrayRS2 = GetArrayRS(SQL_02)
											
											for i=0 to UBound(arrayRS2,2)
												'<- CHG-DESA-30062021-01
												Response.Write vbTab & vbTab &"<option value=" & arrayRS2(0,i) & " " &  arrayRS2(2,i) & ">" & arrayRS2(1,i) & "</option>"
												'CHG-DESA-30062021-01 ->
											next
										%>	
									</select><br /><br />
									<input type="checkbox" name="con_conf" value="1 <%=arrayRS(23,0)%>" class="" />
									Con confirmaci&oacute;n.<br /><br />
								</td>
							</tr>
							<tr class="trHeader">
								<th colspan="2" class="font-size-13_5">Captura los par&aacute;metros</th>
							</tr>
							<%
								if arrayRS(22,0) <> "0" then
									SQL_02 = "select  "
									for i=1 to CInt(arrayRS(22,0))
										if i <> 1 then SQL_02 = SQL_02 & ","
										SQL_02 = SQL_02 & "name_param_" & i
										SQL_02 = SQL_02 & ", opcion_" & i
										SQL_02 = SQL_02 & ", param_" & i
									next
									SQL_02 = SQL_02 & " from rep_reporte rep " & vbCrLf
									SQL_02 = SQL_02 & ", rep_detalle_reporte repdet " & vbCrLf 
									SQL_02 = SQL_02 & " where rep.ID_REP = repdet.ID_REP and repdet.id_cron='"& arrayRS(0,0) &"' "  

									'	response.write SQL_02
									arrayRS2 = GetArrayRS(SQL_02)
									if IsArray(arrayRS2) then
			
										for i=0 to UBound(arrayRS2,1) step 3
											Response.Write "<tr>" & vbCrLf & vbTab
											Response.Write "<td colspan='2'>"

											Response.Write "<table class='width-100p'><tr>"
											
											'Numero y nombre del par�metro:
											Response.Write "<td>"
											if arrayRS2(i+1,0) = "1" then Response.Write "<i>"
												Response.Write (i/3) + 1  &".&nbsp;&nbsp;" & arrayRS2(i,0) & "&nbsp;:&nbsp;&nbsp;"
											if arrayRS2(i+1,0) = "1" then Response.Write "</i>"
											Response.Write "</td>"

											'Campo de texto:
											Response.Write "<td class='width-400'>"
											Response.Write "&nbsp;&nbsp;&nbsp;<input type='text' name='param_" & (i/3) + 1 & "' size='50' class='light height-20 width-350"
											'parametros obligatorios
											if arrayRS2(i+1,0) = "0" then Response.Write " required"
											'Response.Write "' />"
											'Response.Write "</td>"
				
											' Valor del par�metro:
											'Response.Write "<td>"
											Response.Write "' value='"& HTMLEscape(arrayRS2(i+2,0)) &"' /><input type='hidden' name='opcion' value="& arrayRS2(i+1,0) &" /> " & vbCrLf 
											Response.Write "</td>"
											
											Response.Write "</tr></table>"
											
											Response.Write "</tr>" & vbCrLf
										next
										Response.Write "<tr><td colspan='2'><div class='font-size-10_5'><i>En it&aacute;lico, los par&aacute;metros son opcionales.</i><br /><br /></div></td></tr>" & vbCrLf
									end if
								end if
							%>
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
												<td><a href="javascript:void(0);" onmouseover="return overlib('1-&gt;Lunes<br />2-&gt;Martes<br />...<br />7-&gt;Domingo.<br />1-5 para de lunes a viernes.');" onmouseout="return nd();">D&iacute;a de la semana</a></td>
												<td>Prioridad</td>
											</tr>
											<tr align="center">
												<td>
													<input type="text" name="hora" class="light height-20" value="<%=NVL(arrayRS(13,0))%>" size="7" />
												</td>
												<td>
													<input type="text" name="minutos" class="light height-20" value="<%=NVL(arrayRS(14,0))%>" size="7" />
												</td>
												<td>
													<input type="text" name="dia" size="7" class="light height-20" maxlength="20" value="<%=arrayRS(15,0)%>" />
												</td>
												<td>
													<input type="text" name="mes" size="7" class="light height-20" maxlength="20" value="<%=arrayRS(16,0)%>" />
												</td>
												<td>
													<input type="text" name="dia_semana" size="7" class="light height-20 width-80p" maxlength="20" value="<%=arrayRS(17,0)%>" />
												</td>
												<td>
													<select name="prioridad" class="light height-20">
														<%
															for i=1 to 9
																Response.Write vbCrLf & vbTab &"<option value=" & i 
																if i= CInt(arrayRS(18,0)) then Response.Write " selected "
																'<- CHG-DESA-30062021-01
																Response.Write ">" & i & "</option>"
																'CHG-DESA-30062021-01 ->
															next
														%>
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
												<input type="text" name="TEMP_MENSAJE_FECHA" size="10" class="light height-20" value="<%=arrayRS(20,0)%>" align="middle" />
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
												<textarea name="TEMP_MENSAJE" class="light padding-left-10 width-100p" value="<%=arrayRS(21,0)%>" rows="4" cols="80"></textarea>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td class="padding-left-20 padding-top-bottom-10">
									<input type="submit" class="buttonsBlue" value="Guardar" /><br />
									<input type="hidden" name="etape" value="2" />
									<input type="hidden" name="num_param" value="<%=arrayRS(22,0)%>" />
									<input type="hidden" name="num_reporte" value="<%=arrayRS(0,0)%>" />
								</td>
								<td class="padding-left-20 padding-top-bottom-10">
									<center>
										<%
											'if IsArray(arrayRS3) then 
												'if arrayRS3(0,0) = 1  AND (Session("array_user")(0,0) = "HECTORRR" OR Session("array_user")(0,0) = "CESARRP")  then 
												if (Session("array_user")(0,0) = "HECTORRR" OR Session("array_user")(0,0) = "CESARRP")  then 
										%>
													<input type="date" class="light height-20"  align="middle" name="fecha-inicio"/>
													<input type="date" class="light height-20"  align="middle" name="fecha-fin"/>
													<textarea type="text" class="light height-20"  align="middle" name="email-reproceso"> </textarea>
													<input type=button class="buttonsBlue" value="Reprocesar" onclick="function_reprocesos()" /><br />
													<input type=hidden name="reprocesos"  id="reprocesos" />
										<%
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
			</body>
		</html>
		<%

case "2"
	
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
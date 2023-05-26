<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
Response.Expires = 0
call check_session()

Dim Precio_P, Precio_R, state
Dim arrLineCollection, contents, col3
Dim i, x, newLine, arrLineData, lineData
Dim sFuncionDescarga, cmdDescargaStyle, cmdCargarStyle
Dim arrInfoArchivo, sNombreArchivo, first_path, FSO, contenido
Dim Agente, Pedim, Tipo, Firma, Fecha, Hora, RFC, Reg, Archivo, Op


Function NVL(str)
	if IsNull(str) then
		NVL = "" 
	else 
		NVL = str
	end if
End Function


if Request.Form("state") = "1" then
	state = ""
	cmdDescargaStyle = ""
	cmdCargarStyle = "visibility:collapse;display:none;"
	
	
	'------------------------------------------------'
	'VALIDAR QUE SE CAPTUREN LOS DOS DATOS NECESARIOS'
	'------------------------------------------------'
	if Request.Form("txtPrecio_P") = "" then
		Response.Redirect "formato_carga.asp?msg=" & Server.URLEncode("Error: El precio unitario de PEDIMENTOS es obligatorio.")
	end if
	if Request.Form("txtPrecio_R") = "" then
		Response.Redirect "formato_carga.asp?msg=" & Server.URLEncode("Error: El precio unitario de RECTIFICADOS es obligatorio.")
	end if
	
	'---------------'
	'LEER EL ARCHIVO'
	'---------------'
	if Cstr(request.Form("tipo")) = "txt" or Cstr(request.Form("tipo")) = "TXT" then
		contents = request.Form("container")
		arrLineCollection = Split(contents, vbCRLF)
		
		if UBound(arrLineCollection) = 0 then
			Response.Redirect "formato_carga.asp?msg=" & Server.URLEncode("Error: El archivo no tiene informacion.")
		end if
	else
		Response.Redirect "formato_carga.asp?msg=" & Server.URLEncode("Error: El tipo de archivo es incorrecto.")
	end if
	
	'--------------------------------'
	'OBTENER LOS DATOS DEL FORMULARIO'
	'--------------------------------'
	sNombreArchivo = Request.Form("sNombreArchivo")
	Precio_P = Request.Form("txtPrecio_P")
	Precio_R = Request.Form("txtPrecio_R")
	col3 = Request.Form("col3")
	
	
	'------------------------------------------------------------------'
	'GENERAR EL ARREGLO CON LAS LINEAS QUE SE ALMACENARÁN EN EL ARCHIVO'
	'------------------------------------------------------------------'
	contenido = ""
	x = UBound(arrLineCollection)
	ReDim arrInfoArchivo(x-12)
	for i = 13 to UBound(arrLineCollection)
		newLine = ""
		lineData = Trim(arrLineCollection(i))
		'arrLineData = Split(lineData, " ")
		
		'if UBound(arrLineData) > 0 then
		if Len(lineData) > 4 then
			Agente = Trim(Mid(lineData,1,4))
			Pedim = Trim(Mid(lineData,7,7))
			Tipo = Trim(Mid(lineData,17,5))
			Firma = Trim(Mid(lineData,23,8))
			Fecha = Trim(Mid(lineData,32,10))
			Hora = Trim(Mid(lineData,43,5))
			RFC = Trim(Mid(lineData,49,12))
			Reg = Trim(Mid(lineData,67,1))
			Archivo = Trim(Mid(lineData,71,8))
			Op = Trim(Mid(lineData,80,1))
			
			if Trim(Agente) = "" or Trim(Agente) = "T." then
				Exit For
			end if
			
			if Reg <> "B" then
				newLine = Reg & "|" & Pedim & "|" & col3 
				
				if Reg = "P" then
					newLine = newLine & "|" & Precio_P
				else
					newLine = newLine & "|" & Precio_R
				end if
				
				newLine = newLine & "|" & Firma & "|" & Fecha & "|" & Hora & "|"
				
				arrInfoArchivo(i-12) = newLine
				if contenido = "" then
					contenido = Trim(newLine) & vbCRLF
				else
					contenido = contenido & newLine & vbCRLF
				end if
			end if
		else
			Exit For
		end if
	next
	
	'----------------'
	'CREAR EL ARCHIVO'
	'----------------'
	contenido = Trim(contenido)
	sNombreArchivo = "S" & Trim(Mid(sNombreArchivo,2,Len(sNombreArchivo)))

	sFuncionDescarga = ""
	sFuncionDescarga = sFuncionDescarga & "<script type='text/javascript'>" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "	function download()" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "	{" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "		var	filename = '" & sNombreArchivo & "';" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "		var textInput = '" & Replace(contenido,vbCRLF,"\n") & "';" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "		var element = document.createElement('a');" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "		element.setAttribute('href','data:text/plain;charset=utf-8,' + encodeURIComponent(textInput));" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "		element.setAttribute('download', filename);" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "		document.body.appendChild(element);" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "		element.click();" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "	}" & vbCRLF
	sFuncionDescarga = sFuncionDescarga & "</script>" & vbCRLF
	
	Response.Write sFuncionDescarga
else
	state = "1"
	cmdCargarStyle = ""
	cmdDescargaStyle = "visibility:collapse;display:none;"
end if

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
			input[type=file]
			{
				cursor: pointer !important;
			}
			.lblSelecciona
			{
				font-family: "Roboto",sans-serif !important;
				font-size: 14px !important;
				font-weight: bold !important;
				text-align: right !important;
				width: 70% !important;
			}
		</style>
		<title>Crear formato de carga</title>
	</head>
	<body>
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
					CREAR FORMATO DE CARGA
				</h2>
			</div>
		</div>
		<center>
			<div style="width: 800px;">
				<form id="frmCargarArchivo" name="frmCargarArchivo" action="<%=asp_self%>" autocomplete="off" method="post" class="form-cancel-mas">
					<table width="98%" border="0" class="tbl-shadow">
						<tr>
							<td colspan="2" class="center">
								<%
									if Request.QueryString("msg") <> "" then
										%>
											<label class="lblMSG error red" style="padding:3px;">
												<%=Request.QueryString("msg")%>
											</label>
											<br><br>
										<%
									end if
								%>
							</td>
						</tr>
						<tr>
							<td class="lblSelecciona">
								Precio unitario de PEDIMENTOS POR SERVICIOS COMPLEMENTARIOS (P):
							</td>
							<td>
								<input id="txtPrecio_P" name="txtPrecio_P" type="number" class="form-control rounded-txt" title="En blanco se considera como cero." />
							</td>
						</tr>
						<tr>
							<td class="lblSelecciona">
								Precio unitario de RECTIFICADOS SERVICIOS COMPLEMENTARIOS (R):
							</td>
							<td>
								<input id="txtPrecio_R" name="txtPrecio_R" type="number" class="form-control rounded-txt" title="En blanco se considera como cero." />
							</td>
						</tr>
						<tr>
							<td class="lblSelecciona">
								Seleccionar archivo:
							</td>
							<td>
								<input id="txtArchivo" type="file" class="form-control rounded-txt" onchange="showFile(this)" placeholder="Cargar archivo" />
								<input type="hidden" name="tipo" id="tipo" />
								<input type="hidden" name="col3" id="col3" value="<%=col3%>" />
								<input type="hidden" name="sNombreArchivo" id="sNombreArchivo" value="<%=sNombreArchivo%>" />
								<input type="hidden" name="container" id="container" />
							</td>
						</tr>
						<tr>
							<td colspan="2" class="center">
								<br>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="center">
								<input type="hidden" id="state" name="state" value="<%=state%>" />
								<button name="cmdDescargar" id="cmdDescargar" class="buttonsBlue" onclick="download()" style="<%=cmdDescargaStyle%>">Descargar archivo</button>
								<input type="button" id="cmdCarga" name="cmdCarga" class="buttonsBlue" onclick="javascript: validarFormulario();" style="<%=cmdCargarStyle%>" value="Generar" />
							</td>
						</tr>
					</table>
				</form>
			</div>
		</center>
		<script type="text/javascript">
			const formatter = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2, });
			function validarFormulario()
			{
				var fArchivo = "";
				fArchivo = document.getElementById("container").value;
				
				if (fArchivo == "")
				{
					alert("Debe seleccionar un archivo.");
					document.getElementById("txtArchivo").focus();
				}
				else
				{
					if(document.getElementById("txtPrecio_P").value == "")
					{
						document.getElementById("txtPrecio_P").value = "0.00";
					}
					if(document.getElementById("txtPrecio_R").value == "")
					{
						document.getElementById("txtPrecio_R").value = "0.00";
					}
					document.getElementById("txtPrecio_P").value = formatter.format(document.getElementById("txtPrecio_P").value);
					document.getElementById("txtPrecio_R").value = formatter.format(document.getElementById("txtPrecio_R").value);
					document.getElementById("frmCargarArchivo").submit();
				}
			}
			function showFile(input)
			{
				var col3 = "";
				var sNombreArchivo = "";
				
				if (document.getElementById("lblMsg") != null)
				{
					document.getElementById("lblMsg").innerText = "";
				}
				let file = input.files[0];
				let reader = new FileReader();
				
				if (!file)
				{
					alert("Fallo al abrir el archivo");
				}
				else if (!file.type.match('text.*'))
				{
					document.forms[0].tipo.value = (file.name).substring((file.name).length - 3);
					reader.onload = function (e)
									{
										var data = e.target.result;
										document.forms[0].container.value = reader.result;
										
										var workbook = XLSX.read(data, { type: 'binary' });
				
										workbook.SheetNames.forEach(function (sheetName)
										{
											var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
											if (XL_row_object != "")
											{
												var json_object = JSON.stringify(XL_row_object);
												document.forms[0].container.value = json_object;
											}
										})

									};

					reader.onerror = function (ex) { console.log(ex); };
					reader.readAsBinaryString(file);
				}
				else
				{
					col3 = (file.name).substring(1,3);
					sNombreArchivo = file.name;
					
					document.getElementById("col3").value = col3;
					document.getElementById("sNombreArchivo").value = sNombreArchivo;
					
					document.forms[0].tipo.value = "txt";
					reader.onload = function ()
									{
										document.forms[0].container.value = reader.result;
										document.forms[0].tipo.value = (file.name).substring((file.name).length - 3);
									};
					reader.onerror = function () { console.log(reader.error); };
					reader.readAsText(file);
				}
			}
		</script>
	</body>
</html>
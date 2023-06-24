<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Confirmacion de los reportes
Response.Expires = 0
call check_session()

	'<- CHG-DESA-30062021-01
	dim NumCli
	NumCli = Session("cli_num")
	dim UrlListaContactos
	' CHG-DESA-30062021-01 ->

Select Case Request("Etape")
	Case ""
%>
	<html>
		<head>
			<meta http-equiv="Content-Type" content="text/html;" charset="iso-8859-1" />
			<link href="css/logis_style.css" media="all" type="text/css" rel="stylesheet" />
			<title>Gestion de correos</title>
			<script src="js/jquery-1.3.2.min.js"></script>
			<script src="js/main.js"></script>

			<!--<script type="text/javascript"> 
                var Type;
                var Url;
                var Data;
                var ContentType;
                var DataType;
                var ProcessData;
               
                function tmp_ws() {
                    const xhr = new XMLHttpRequest('select_activos');
                    var status = document.getElementById('select_activos').value.trim();
                    var usr = '<%=Session("array_user")(0,0)%>';
                    const url = urlWebService + "GetConsultaReportes?usuario=" + usr + "&status=" + status;
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
                    var info = objResult.GetConsultaReportesResult;
                    var arrayRS3 = JSON.parse(info);

                    var i = 0;
                    var htmlTable = "";
                    var SQL = "";
                    var bandera = 0;
                    $("#tbResult").empty();
                    if (arrayRS3.length == 0) {
                        htmlTable = htmlTable + "<tr class='center'>";
                        htmlTable = htmlTable + "	<td colspan='9' class='center'>";
                        htmlTable = htmlTable + "		No hay reportes registrados.";
                        htmlTable = htmlTable + "	</td>";
                        htmlTable = htmlTable + "</tr>";
                    } else {

                        for (i = 0; i < arrayRS3.length; i++) {
                            htmlTable = "";

                            
                            htmlTable = htmlTable + "</tr>"
                            $("#tbResult").append(htmlTable);
                        }
                    }
                    //hideLoading();
                }
            </script>-->
			
			<script>
				//<- CHG-DESA-30062021-01
				function validaOk()
				{ 
					var sURI = window.location.search;
					var arrParams = new URLSearchParams(sURI);
					var AllOk = arrParams.get("allOk");
					var IsPopUp = 0;

					try{if(arrParams.get("et")=="-1") localStorage.setItem('pop',0);}catch{}
					IsPopUp = localStorage.getItem('pop');
					if (IsPopUp == null) { IsPopUp = "0"; }

					document.getElementById("hdnURI").value = localStorage.getItem("sURI_list");

					if(IsPopUp = "1")
					{
						var menuBar = document.getElementById("hdnURI");
					}

					if (AllOk != undefined) {
						if (AllOk != null) {
							if(AllOk == 1) {
								var sURI_list = localStorage.getItem("sURI_list");
								window.open(sURI_list, "Lista_contactos", "toolbar=no, location=no, directories=no, status=yes, scrollbars=yes, resizable=yes, copyhistory=no, width=500, height=400, left=300, top=50");
								window.close();
							}
						}
					}
				}
				function ValidaCorreo()
				{
					var sCorreo;
					var sURI = window.location.search;	
					var arrParams = new URLSearchParams(sURI);
					var cli_num = arrParams.get("Num");
					sCorreo = document.getElementById("txtCorreo").value.trim();
					document.getElementById("txtCorreo").value = sCorreo;

					if(document.getElementById("cli_num").value != "")
					{
						if(cli_num == null)
						{
							cli_num = document.getElementById("cli_num").value;
						}
					}
			
					if(sCorreo != "")
					{
						if(sCorreo.trim().endsWith("logis.com.mx"))
						{
							document.getElementById("chkTercero").checked = false;
							document.getElementById("hdnTercero").value = false;
/* < JEMV
							document.getElementById("cli_num").value = "9929";
							document.getElementById("hdn_cli_num").value = "9929";
 JEMV > */
						}
						else
						{
							document.getElementById("chkTercero").checked = true;
							document.getElementById("hdnTercero").value = true;
							document.getElementById("cli_num").value = cli_num;
							document.getElementById("hdn_cli_num").value = arrParams.get("Num");
						}
						if(document.getElementById("ListUrl") != null)
							document.getElementById("ListUrl").value = "mail.asp?etape=2&cliente=" + cli_num;
					}
				}
				//CHG-DESA-30062021-01 ->
			</script>
		</head>
		<!-- <-CHG-DESA-30062021-01 -->
		<body  onload="lockFields();validaOk();">
		<!-- CHG-DESA-30062021-01 -> -->
			<center>
		<%
			call print_style()
			Dim SQL, nombre, correo, tercero, cliente, arrayRS 
			'<- CHG-DESA-30062021-01
			if not Request.QueryString("Num") is nothing then
				if Request.QueryString("Num") <> "" then
					cliente = SQLescape(Request.QueryString("Num"))
			response.write cliente
			response.End
					Session("cli_num") = cliente
'< JEMV
					'Response.write("<script>function lockFields() { document.getElementById('cli_num').disabled = true; document.getElementById('chkStatus').disabled = true; document.getElementById('chkTercero').disabled = true; }</script>")
					Response.write("<script>function lockFields() { }</script>") 
' JEMV >
				else
					Response.write("<script>function lockFields() { }</script>") 
				end if
			else
				Response.write("<script>function lockFields() { }</script>") 
			end if
			'CHG-DESA-30062021-01 ->
		'reitulizar ftn_lista_contactos
			SQL =   " select nombre, mail, client_num,  " & _
				" decode (tercero, 1, 'checked', ''), decode (status, 1, 'checked', '') " & _
				" from rep_mail " & _
				" where id_mail= '"&SQLEscape(Request.QueryString("mail")) & "' " 
		arrayRS = GetArrayRS(SQL)
		if IsArray(arrayRS) then
			nombre = arrayRS(0,0)
			correo = arrayRS(1,0)
			cliente = arrayRS(2,0)
			tercero = arrayRS(3,0)
			status = arrayRS(4,0)
		end if

UrlListaContactos = "mail.asp?etape=2"

if not Request.QueryString("Num") is nothing then
	if not Request.QueryString("Num") = "" then
		UrlListaContactos = "mail.asp?etape=2&cliente="+Request.QueryString("Num")
	end if
end if
if HTMLEscape(cliente) <> "" then
	UrlListaContactos = "mail.asp?etape=2&cliente=" & HTMLEscape(cliente)
end if
		%>

				<div class="contenedorMenu">
					<div class="dvMenu">
						<ul id="menu">
							<div class="logo-logis">
								<img src="images/logo-logis-s.png" alt="Logo de Logis" height="55" />
							</div>
							<li onclick="window.location.href='menu.asp';">
								Menu
							</li>
							<% if Request.QueryString("et") <> "-1" then %>
							<li onclick="window.location.href='<%=UrlListaContactos %>';">
								Lista de contactos
							</li>
							<% end if %>
						</ul>
					</div>
				</div>
				
				<hr/>
				
				<table border="0" width="350" class="tblMenu">
					
					<tbody>
						<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>" %>
					</tbody>
				</table>			
		<script LANGUAGE="JavaScript">
			function check_data()
			{
				var i, errors_radio, error_file, msg;
				var arrRequiredFields;
				errors_radio=1;
				error_file = 1 ;
				msg = "";

				arrRequiredFields = document.getElementsByClassName("required");
				
				if (arrRequiredFields != null) {
					for (i in arrRequiredFields) {
                        console.log(arrRequiredFields[i].id);
						/*
						if (document.getElementById(i.id).value == "") {
							document.getElementById(i.id).addClass("error");
                        }
						*/
                    }
                }
				

				if (document.mail_form.nombre.value == "")
					{msg = "- Entrar un nombre.\n";}
				
				if (document.mail_form.correo.value == "")
					{ msg += "- Entrar una direcion de correo."; }
	
	
				if (msg != "")
					{alert("Verifique los datos :\n"+msg);}
				else
                    tmp_ws2()
                //document.mail_form.submit();
            }
			function tmp_ws2() {
				var cte = '<%= Request.QueryString("cliente") %>';
				alert(cte);
                var NumCli = "";
                var Tercero = false;
                var id_mail
                var nombre;
                var correo;
                var hdnURI;
                if (document.getElementById("hdn_cli_num").value != "") {
                    NumCli = document.getElementById("hdn_cli_num").value;
                }
                else if (document.getElementById("cli_num").value != "") {
                    NumCli = document.getElementById("cli_num").value;
                }
                Tercero = document.getElementById("hdnTercero").value;
                id_mail = document.getElementById("id_mail").value;
                nombre = document.getElementById("txtNombre").value;
                correo = document.getElementById("txtCorreo").value;
				status = document.getElementById("chkStatus").value;
				hdnURI = document.getElementById("hdnURI").value;

                const xhr = new XMLHttpRequest();
                const url = urlWebService + "GetMail?Id_Cron=" + $("#hdn_Id_Cron").val() + "&NumCli=" + NumCli + "&id_mail=" + id_mail + "&nombre=" + nombre + "&correo=" + correo + "&Tercero=" + Tercero + "&status=" + status + "&hdnURI=" + hdnURI + ")";
                var someHandler = "ok";

                xhr.onreadystatechange = function () {
                    if (xhr.readyState == XMLHttpRequest.DONE) {
                        mostrarResultado2(xhr.responseText);
                    }
                }

                xhr.open("GET", url, true);
                xhr.send();
            }
            function mostrarResultado2(wsResponseText) {
                var objResult = JSON.parse(wsResponseText);
                var info = objResult.GetMailResult;
                window.location.href = "mail.asp?msg=" + info;
            }

        </script>
		<!-- <- CHG-DESA-30062021-01 -->
		<form name="mail_form"  action="<%=asp_self()%>?etape=1&cli_num=<%=HTMLEscape(cliente)%>" method="post">
			<table border="0" width="350" class="">
				<thead>
					<tr class="trHeader">
						<th colspan="2">
							Agregar otro contacto
						</th>
					</tr>
				</thead>
				<tbody>
					<tr>
						<td class="tdLabel">
							N&uacute;mero de cliente <i>(9929 para Logis)</i> :
						</td>
						<td class="tdField">
							<input type=text id="cli_num" name=cli_num size=5 value="<%=HTMLEscape(cliente)%>" class="required" />
							<input type="hidden" id="hdn_cli_num" name="hdn_cli_num" value="<%=HTMLEscape(cliente)%>" />
						</td>
					</tr>
					<tr>
						<td class="tdLabel">
							Direccion de correo :<br>
						</td>
						<td class="tdField" id="mail">
							<input type=text name=correo size=25 value="<%=correo%>" id="txtCorreo" onchange="ValidaCorreo();" class="required">
							<input type="hidden" id="hdnURI" name="hdnURI" value="<%= Request.QueryString("hdnURI") %>" />
						</td>
					</tr>
					<tr>
						<td class="tdLabel">
							Nombre :
						</td>
						<td class="tdField">
							<input type="text" name="nombre" size="25" id="txtNombre" value="<%=HTMLEscape(nombre)%>" class="required" />
						</td>
					</tr>
					<tr>
						<td class="tdLabel">
							Tercero :
						</td>
						<td class="tdField">
							<input type=checkbox name=tercero id="chkTercero" value=1 <%=tercero%> />
							<input type="hidden" id="hdnTercero" name="hdnTercero" />
						</td>
					</tr>
					<tr>
						<td class="tdLabel">
							Activo :
						</td>
						<td class="tdField">
							<input type=checkbox id="chkStatus" name=status value=1 
								<% if IsArray(arrayRS) then 
								Response.Write status
								else
								Response.Write "checked"
								end if
								%>>
						</td>
					</tr>
					<tr>
						<td colspan="2" align="center">
							<input type="hidden" id="id_mail" name="id_mail" value="<%=Request.QueryString("mail")%>" />
							<input type="hidden" id="hdn_Id_Cron" name="hdn_Id_Cron" value="<%=Request.QueryString("Id_Cron")%>" />
							<input type="button" id="cmdValida" onClick="check_data();" class="buttonsBlue" value="Agregar" />
							<br><br>
						</td>
					</tr>
				</tbody>
			</table>
		</form>
		<!-- CHG-DESA-30062021-01 -> -->
		
<%
case "1"
''''		dim rst, msg, status
''''	'<- CHG-DESA-30062021-01
''''	dim allOk, cte
		response.write cliente
			response.End
''''	allOk = 0
''''		
''''	dim tercer
''''	dim idMail
''''	NumCli = SQLescape(Request.Form("hdn_cli_num"))
''''	tercer = SQLescape(Request.Form("hdnTercero"))
''''	cte = Request.QueryString("cliente")
''''	'CHG-DESA-30062021-01 ->
''''
''''	set rst = Server.CreateObject("ADODB.Recordset")
''''			
''''		status = Request.Form("status") 
''''		if status = "" then Status = 0
''''		
''''		SQL = "select 1 from eclient where cliclef='" & SQLescape(Request.Form("hdn_cli_num")) & "'"
''''	
''''		'<- CHG-DESA-30062021-01
''''		Dim Id_Cron , mail_ok
''''		Id_Cron =SQLescape(Request.Form("hdn_Id_Cron"))
''''
''''		SQL = " select MAIL_OK,ID_CRON from rep_detalle_reporte where ID_CRON = '" & SQLEscape(Id_Cron) & "' "
''''
''''		arrayRS = GetArrayRS(SQL)
''''		if IsArray(arrayRS) then
''''			mail_ok = arrayRS(0,0)
''''		end if
''''	
''''		if NumCli <> "" then
''''			SQL = "select 1 from eclient where cliclef='" & NumCli & "'"
''''		else
''''			if SQLescape(Request.Form("cli_num")) <> "" then
''''				SQL = "select 1 from eclient where cliclef='" & Request.Form("cli_num") & "'"
''''			end if
''''		end if
''''		arrayRS = GetArrayRS(SQL)
''''		if not IsArray(arrayRS) then
''''			if NumCli <> "" then
''''				Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Este numero de cliente '" & NumCli & "' no existe.")
''''			else
''''				Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Este numero de cliente '" & SQLescape(Request.Form("cli_num")) & " no existe.")
''''			end if	
''''		end if
''''				
''''		if Request.Form("id_mail") <> "" then
''''			if NumCli <> "" then
''''				SQL = " update rep_mail set nombre= '" & SQLEscape(Request.Form("nombre")) &"', "& _
''''					  " mail = '" & SQLEscape(Request.Form("correo")) & "', " & _
''''					  " client_num = '" & NumCli & "', " & _
''''					  " tercero = '" & SQLEscape(Request.Form("hdnTercero")) &"', " & _
''''					  " status = '" & status &"' " & _
''''					  " where id_mail= '" & SQLEscape(Request.Form("id_mail")) &"' "
''''			else
''''				SQL = " update rep_mail set nombre= '" & SQLEscape(Request.Form("nombre")) &"', "& _
''''					  " mail = '" & SQLEscape(Request.Form("correo")) & "', " & _
''''					  " client_num = '" & SQLescape(Request.Form("cli_num")) & "', " & _
''''					  " tercero = '" & SQLEscape(Request.Form("hdnTercero")) &"', " & _
''''					  " status = '" & status &"' " & _
''''					  " where id_mail= '" & SQLEscape(Request.Form("id_mail")) &"' "
''''			end if
''''			msg = "Contacto Modificado"
''''			allOk = 1
''''		else
''''			'verificacion que el correo es unico en la base
''''			if NumCli <> "" then
''''				SQL = " select 1 from rep_mail where mail = '" & SQLEscape(Request.Form("correo")) & "' " & _
''''					  " and CLIENT_NUM = '" & NumCli & "'"
''''			else
''''				SQL = " select 1 from rep_mail where mail = '" & SQLEscape(Request.Form("correo")) & "' " & _
''''						" and CLIENT_NUM = '" & SQLescape(Request.Form("cli_num")) & "'"
''''			end if
''''			arrayRS = GetArrayRS(SQL)
''''
''''			if  IsArray(arrayRS) then
''''				Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Este correo ya existe para este cliente.")
''''			end if
''''			
''''			'verificar que no se capturen correos de Logis para otro numero de cliente que el 9929
''''		    if NumCli <> "" then
'''''<JEMV
'''''				if InStr(LCase(Request.Form("correo")), "@logis.com.mx") > 0 and NumCli <> "9929" then
'''''					Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Favor de crear los correos de Logis con el numero de cliente 9929.")
'''''				end if
'''''JEMV>
''''
''''				SQL =	"insert into rep_mail (ID_MAIL, NOMBRE, MAIL, CLIENT_NUM, TERCERO, STATUS) " & _
''''						" values  (seq_mail.nextval, '"& _
''''						SQLEscape(Request.Form("nombre")) &"', '"& SQLEscape(Request.Form("correo")) & _
''''						"', '" & NumCli & "', null,  1" & _
''''						" )"
''''			else
'''''<JEMV
'''''				if InStr(LCase(Request.Form("correo")), "@logis.com.mx") > 0 and Request.Form("cli_num") <> "9929" then
'''''					Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Favor de crear los correos de Logis con el numero de cliente 9929.")
'''''				end if
'''''JEMV>
''''
''''				SQL =	"insert into rep_mail (ID_MAIL, NOMBRE, MAIL, CLIENT_NUM, TERCERO, STATUS) " & _
''''						" values  (seq_mail.nextval, '"& _
''''						SQLEscape(Request.Form("nombre")) &"', '"& SQLEscape(Request.Form("correo")) & _
''''						"', '" & SQLescape(Request.Form("cli_num")) & "', null,  1" & _
''''						" )"
''''			end if
''''	
''''			msg = "Contacto incluido"
''''			allOk = 1
''''		end if
''''		'CHG-DESA-30062021-01 ->
''''		
''''
''''	rst.Open SQL, Connect(), 0, 1, 1
''''
''''	'<- CHG-DESA-30062021-01	
''''	SQL = " select ID_MAIL,CLIENT_NUM from rep_mail where mail = '" & SQLEscape(Request.Form("correo")) & "' " & _
''''					  " and CLIENT_NUM = '" & NumCli & "'"
''''
''''	arrayRS = GetArrayRS(SQL)
''''	if IsArray(arrayRS) then
''''		idMail = arrayRS(0,0)
''''	end if
''''	
''''	SQL = " insert into rep_dest_mail (id_dest_mail, id_dest) " & _
''''		  " values ('"& mail_ok &"','"& idMail &"' ) "
''''
''''	rst.Open SQL, Connect(), 0, 1, 1
''''
''''		if NumCli <> "" then
''''			' Cerrar ventana actual y actualizar ventana ver_lista
''''			dim urlVer
''''
''''			urlVer = SQLEscape(Request.Form("hdnURI"))& "&msg=El usuario " & SQLEscape(Request.Form("correo")) & " fue agregado correctamente."
''''			Response.Redirect urlVer
''''		else
''''			Response.Redirect asp_self & "?msg=" & Server.URLEncode (msg)
''''		end if
''''	'CHG-DESA-30062021-01 ->
		
case "2"
		dim i, filtro
	cte = Request.QueryString("cliente")
	if cte <> "" then NumCli = cte
		%>
	<html>
		<head>
			<title>Gestion de correos</title>
		<script type="text/javascript">
            //<- CHG-DESA-30062021-01	
			function select_Client() {
				var NumCli = localStorage.getItem('Cli');
				var cte = '<%= cte %>';
				if (cte != "") {
					NumCli = cte;
                    localStorage.setItem('Cli',cte);
                }

				document.getElementById('ListUrl').value = "mail.asp?etape=2&cliente=" + NumCli;
				document.getElementById('ListUrl').text = NumCli;
                document.getElementById('ListUrl').text = NumCli;
                //document.getElementById('ListUrl').disabled = true;
            }
			// CHG-DESA-30062021-01	->
        </script>
			<link href="css/logis_style.css" media="all" type="text/css" rel="stylesheet" />
		</head>
		<!-- <-CHG-DESA-30062021-01	-->
<% if NumCli <> "" then %>
		<body onload="select_Client();">
<% else %>
		<body>
<% end if %>
			<center>
<!-- CHG-DESA-30062021-01 ->	-->
			
		<SCRIPT LANGUAGE="JavaScript">
			/*
			SCRIPT EDITE SUR L'EDITEUR JAVASCRIPT
			http://www.editeurjavascript.com
			*/
			function ChangeUrl(formulaire) {
				if (formulaire.ListeUrl.selectedIndex != 0) {
					var indexTMP = formulaire.ListeUrl.selectedIndex
					formulaire.ListeUrl.selectedIndex = 0
					location.href = formulaire.ListeUrl.options[indexTMP].value;
				}
				else { alert('Please choose a destination.'); }
			}
		
			function delete_mail(nombre, direccion, id_mail) {
				// pose une question au visiteur
				var msg
				msg = "Esta seguro de borrar este contacto :"
				msg += "\n- " + nombre
				msg += "\n- " + direccion

				if (confirm(msg)) {
					/*alert("Vous avez cliqu� sur OK " + id_mail);*/
					location.href = "mail.asp?Etape=3&id_mail=" + id_mail;
				} /*else {	
				alert("Vous avez cliqu� sur Annuler");
			}*/
			}
        </SCRIPT>
<%call print_style()
	SQL = "select distinct client_num from rep_mail order by client_num "
	arrayRS = GetArrayRS(SQL)
	if not IsArray(arrayRS) then
		Response.Write "No hay contactos."
		Response.End 
	end if
	cte = Request.QueryString("cliente")
%>
				<div class="contenedorMenu">
					<div class="dvMenu">
						<ul id="menu">
							<div class="logo-logis">
								<img src="images/logo-logis-s.png" alt="Logo de Logis" height="55" />
							</div>
							<li onclick="window.location.href='menu.asp';">
								Menu
							</li>
							<!-- <- CHG-DESA-30062021-01	-->
<%
	if NumCli <> "" then
%>
							<li onclick="window.location.href='<%=asp_self() & "?Num=" & NumCli%>';">
								Agregar contacto
							</li>
<%
	elseif cte <> "" then
%>
							<li onclick="window.location.href='<%=asp_self() & "?Num=" & cte%>';">
								Agregar contacto
							</li>
<%
	else
%>
							<li onclick="window.location.href='<%=asp_self()%>';">
								Agregar contacto
							</li>
<%
	end if
%>
						</ul>
					</div>
				</div>
				
				<hr/>
				
				<form id="frmCase2">
					<table border="0" width="350" class="tblForm">
						<thead>
							<tr>
								<th colspan="6">Lista de contactos</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="tdLabel">
									Escoge numero de cliente (9929 = Logis) :
								</td>
								<td class="tdField">
									<select id="ListUrl" NAME="ListeUrl" SIZE="1" onChange="ChangeUrl(this.form)">
										<option selected value="">-Cliente Num-</option>
<%
dim NoCli_req, NoCli_bd
	
	NoCli_req = ""

	if Request.QueryString("cliente") <> "" then 
		NoCli_req = CStr(Request.QueryString("cliente"))
	end if
	
	for i=0 to UBound(arrayRS,2)
		NoCli_bd = CStr(arrayRS(0,i))
		
		if NoCli_req <> "" then
			if NoCli_req = NoCli_bd then
				Response.Write "<OPTION VALUE='" & asp_self() & "?etape=2&cliente=" & NoCli_bd & "' selected>" & NoCli_bd & "</option>"
			else
				Response.Write "<OPTION VALUE='" & asp_self() & "?etape=2&cliente=" & NoCli_bd & "'>" & NoCli_bd & "</option>"
			end if
		else
			Response.Write "<OPTION VALUE="""& asp_self() & "?etape=2&cliente=" & arrayRS(0,i) 
			Response.Write """>" & arrayRS(0,i)
		end if
	next
%>
									</select>
								</td>
							</tr>
							<tr>
								<td colspan="2" class="align-right">
									<font class="red-note">
										En rojo, el contacto esta desactivado.
									</font>
								</td>
							</tr>
						</tbody>
					</table>
				</form>
				
<%	
if Request.QueryString("cliente") <> "" then
	SQL = " select id_mail, nombre, mail, " & _
			" decode(client_num, 9929, 'Logis', client_num) as client_num, " & _
			" decode(tercero, 1, 'Si', '') as tercero, status " & _
			" From rep_mail " & _
			" Where client_num = '" & SQLescape(Request.QueryString("cliente")) & "'" & _
			" order by tercero desc, nombre "
	'Response.Write SQL
	arrayRS = GetArrayRS(SQL) 
	if not IsArray(arrayRS) then
		Response.Write "Ninguno contacto por este cliente."
		Response.End 
	end if
%>

				<table border="0" width="350" class="tblContent">
					<thead>
						<tr class="trHeader" align="center">
							<td size=15>.</td>
							<td>Nombre</td>
							<td>Correo</td>
							<td>Cliente</td>
							<td>Tercero</td>
							<td>Acci&oacute;n</td>
						</tr>
					</thead>
					<tbody>
<%
	for i = 0 to UBound(arrayRS,2)
		Response.Write "<tr"
		if arrayRS(5,i) = "0" then Response.Write " style=""color:red; "" "
		Response.Write ">" & vbCrLf & vbTab 
		Response.Write "<td>" & i & "</td>" & vbCrLf & vbTab 				
		Response.Write "<td>" & arrayRS(1,i) & "</td>"
		Response.Write vbCrLf & vbTab
		Response.Write "<td><a href=""mailto:" & arrayRS(2, i) & """>" & arrayRS(2, i) & "</a></td>" & vbCrLf & vbTab  
		Response.Write "<td>" & arrayRS(3, i) & "</td>" & vbCrLf & vbTab  
		Response.Write "<td>" & arrayRS(4, i) & "</td>" & vbCrLf & vbTab  
		Response.Write "<td><a href=""" & asp_self() & "?mail=" & arrayRS(0,i) & """>Modificar</a>&nbsp;|&nbsp;"
		Response.Write "<a href=""#"" Onclick=""delete_mail('"&JSescape(arrayRS(1, i))&"','"&JSescape(arrayRS(2, i))&"', '"&JSescape(arrayRS(0, i))&"');"">Borrar</a></td>" & vbCrLf & vbTab  '" & asp_self() & "?etape=3&id_mail=" & arrayRS(0,i) & "
		Response.Write "</tr>" & vbCrLf 
	next
end if 
%>
					</tbody>
				</table>
		<%
		
		'set rst = Server.CreateObject("ADODB.Recordset")
		'cliente = Request("cliente")
		'date_num = Request("date_num")
		'if date_num = "" or cliente = "" then Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Error")
		
		'SQL = "delete from rep_dias_libres where dia_libre = to_date('"& SQLEscape(date_num) &"', 'mm/dd/yyyy') and cliente = "&SQLEscape(cliente)
		'Response.Write SQL
		
		'rst.Open SQL, Connect(), 0, 1, 1
		'Response.Redirect asp_self

case "3"
		set rst = Server.CreateObject("ADODB.Recordset")
		
		SQL = "delete from rep_mail where id_mail='" & SQLescape(Request.QueryString ("id_mail")) & "'"
		Response.Write SQL
		rst.Open SQL, Connect(), 0, 1, 1
		Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Contacto borrado")

end select

%>
		</center>
	</body>
</html>
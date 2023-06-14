<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'pop up lista de contactos
Response.Expires = 0
call check_session()
%>
<html>
	<head>
		<link href="include/logis.css" type="text/css" rel="stylesheet" />
		<link href="css/logis_style.min.css" type="text/css" rel="stylesheet" />
		<script src="js/jquery-1.3.2.min.js"></script>
		<script src="js/main.js"></script>
		<script language="JavaScript" src="./include/tigra_tables.js"></script>
		<script type="text/javascript" src="js/reports.min.js"></script>
		<script language="JavaScript" src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
		<script type="text/javascript">
            var Type;
            var Url;
            var Data;
            var ContentType;
            var DataType;
			var ProcessData;
            var lista;
            $(document).ready(

                function () {
                    showLoading();
                    tmp_ws();
                }
            );

            function tmp_ws() {
                const xhr = new XMLHttpRequest('select_activos');
				lista = '<%=Request.QueryString("liste")%>';
                const url = urlWebService + "GetListaCorreos?lista=" + lista;
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
                var info = objResult.GetListaCorreosResult;
                var arrayRS3 = JSON.parse(info);

                var i = 0;
                var htmlTable = "";
                var SQL = "";
                var bandera = 0;
                $("#tbResult").empty();
                if (arrayRS3.length == 0) {
                    htmlTable = htmlTable + "<tr class='center'>";
                    htmlTable = htmlTable + "	<td colspan='9' class='center'>";
                    htmlTable = htmlTable + "		no contactos.....";
                    htmlTable = htmlTable + "	</td>";
                    htmlTable = htmlTable + "</tr>";
                } else {

                    for (i = 0; i < arrayRS3.length; i++) {
						htmlTable = htmlTable + arrayRS3[i].ID_DEST + ", ";                        
					}
					agregarlista(htmlTable.substr(0, htmlTable.length-2));
                }
                //hideLoading();
			}
			//sugunda funcion llena lista
            function agregarlista(grupolista) {
                const xhr = new XMLHttpRequest('select_activos');
				var lista = '<%=Request.QueryString("liste")%>';
                const url = urlWebService + "GetArmarContactos?mail_id=" + grupolista;
                var someHandler = "ok";

                xhr.onreadystatechange = function () {
                    if (xhr.readyState == XMLHttpRequest.DONE) {
                        mostrarLista(xhr.responseText);

                    }

                }

                xhr.open("GET", url, true);
                xhr.send();
            }
            function mostrarLista(wsResponseText) {
                var objResult = JSON.parse(wsResponseText);
                var info = objResult.GetArmarContactosResult;
                var arrayRS3 = JSON.parse(info);

                var i = 0;
                var htmlTable = "";
                var SQL = "";
                var bandera = 0;
                $("#tbResult").empty();
                if (arrayRS3.length == 0) {
                    htmlTable = htmlTable + "<tr class='center'>";
                    htmlTable = htmlTable + "	<td colspan='9' class='center'>";
                    htmlTable = htmlTable + "		no contactos.....";
                    htmlTable = htmlTable + "	</td>";
                    htmlTable = htmlTable + "</tr>";
                } else {

					for (i = 0; i < arrayRS3.length; i++) {
						htmlTable = "";
						htmlTable = htmlTable + "<tr>";
                        htmlTable = htmlTable + "<td>" + arrayRS3[i].NOMBRE + "</td>";
                        htmlTable = htmlTable + "<td><a href=" + "mailto:" + arrayRS3[i].MAIL + "" + ">" + arrayRS3[i].MAIL + "</a></td>";
                        if (lista == "1") {
                            htmlTable = htmlTable + "<td>" + arrayRS3[i].CLIENT_NUM + "</td>";
                            htmlTable = htmlTable + "<td>" + arrayRS3[i].TERCERO + "</td>";
						}
						
                        htmlTable = htmlTable + "</tr>"
                        $("#tbResult").append(htmlTable);
                    }

                }
                //hideLoading();
            }







        </script>


		<style type="text/css">
/* <- CHG-DESA-30062021-01 */
			.trHeader
			{
				background-color: #223F94;
				font-family: "Roboto",sans-serif;
				font-size: 12px;
				color:#FFFFFF;
				text-align: center;
			}
			.trHeader>th
			{
				font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
				font-style: normal;
				font-weight: bold;
				font-size: 12px;
			}
			.buttonsBlue
			{
				background-color: #223F94;
				color:#FFFFFF;
				font-family: "Roboto",sans-serif;
				font-size: 11px;
				font-weight: bold;
				height: 19px;
				cursor: pointer;
			}
			td { font-weight: normal !important; }
			.tblBody>tbody>tr>td
			{
				font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
				font-style: normal;
				font-size: 10px;
			}
			.tblBody>tbody>tr
			{
				cursor: default;
			}
			.trBgColor
			{
				background-color: #F6F6F6;
			}
			.tblBody>tbody>tr:hover
			{
				background-color: #E6E3DD;
			}
			.lblMSG
			{
				align-content: center;
				color: red;
				font-size: 11.5px;
				padding: 10px;
				text-align: center;
				width: 100%;
			}
	/* CHG-DESA-30062021-01 -> */
		</style>
		<% call print_style() %>
		<title>
			Lista de contactos
			<%if Request.QueryString("liste") = "error" then Response.Write " en caso de error"%>
		</title>
	<!-- <- CHG-DESA-30062021-01 -->
		<script type="text/javascript">            
			function Mostrar_AgregarCorreo() {
				var IsPopUp = 0;
                var sURI = window.location.search;
				var arrParams = new URLSearchParams(sURI);                
				var NumCte = localStorage.getItem('Cli');
				var Id_Cron = localStorage.getItem('Id_Cron');

				IsPopUp = localStorage.getItem('pop');
				if (IsPopUp == null) { IsPopUp = "0"; }
                localStorage.setItem('pop', IsPopUp);
				
				location.href = "mail.asp?Num=" + NumCte + "&Id_Cron=" + Id_Cron;
				/*
				if (NumCte != undefined) {
					if (NumCte != null) {
						window.close();
					}
				}
				*/
			}
			function saveuri() { localStorage.setItem("sURI_list", document.getElementById("hdnURI").value); }
        </script>
	<!-- CHG-DESA-30062021-01 -> -->
	</head>
	<body onload="saveuri()">
<!-- <- CHG-DESA-30062021-01 -->		
		<center>
			<table>
				<tr>
					<td colspan="2" style="text-align:right;">
						<!--<input type="button" class="buttonsBlue" onclick="Mostrar_AgregarCorreo()" value='Agregar contacto' />-->
					</td>
				</tr>
			</table>
			<table border="0" width="350" class="tblHeader">
				<tr>
					<td>
						<%if Request("msg") <> "" then	Response.Write "<label class='lblMSG'>" & Request("msg") & "</label>"%>
					</td>
				</tr>
				<tr class="trHeader">
					<td align="center">
						<%if Request.QueryString("liste") = "error" then
							Response.Write "Contactos en caso de error"
						  else
							Response.Write "Lista de contactos"
						  end if
						%>
					</td>
				</tr>
			</table>
<%
	dim Id_cron
	dim arrNombreReporte
	dim sNombreReporte
	
	Id_cron = Request.QueryString("Num")
	'arrNombreReporte = GetArrayRS("SELECT Name FROM rep_detalle_reporte repdet WHERE repdet.ID_CRON = " & Id_cron)
	
	if not IsArray(arrNombreReporte) then 
		sNombreReporte = ""
	else
		sNombreReporte=arrNombreReporte(0,0)
	end if
%>
			<input type="hidden" id="hdnURI" name="hdnURI" value="<%= asp_self()&"?"&request.QueryString %>" />
			<table border="0" width="350" class="tblBody">
				<tr>
					<td>
						Nombre del Reporte:
					</td>
					<td>
						<b>
							<% Response.Write sNombreReporte %>
						</b>
					</td>
				</tr>
				<tr>
					<td>
						No. Cliente:
					</td>
					<td>
						<b>
							<script type="text/javascript">document.write(localStorage.getItem('Cli')); </script>
						</b>
					</td>
				</tr>
			</table>
			
			
<!-- CHG-DESA-30062021-01 -> -->
<%

dim tab_contactos
select case Request.QueryString("liste")
	case "ok"
		'ver los contactos
		tab_contactos = split(Session("id_mail"), ",")
	case "error"
		'ver los contactos en caso de error
		tab_contactos = Session("id_mail_error")
	case else
		'ver los contactos por numero de lista de contactos
		'usado por modif_reporte
		if not IsNumeric(Request.QueryString("liste")) then
			Response.Write "error, el numero de la lista no es numerico."
			Response.End 
		end if
end select
	



dim i, mail_id, SQL, arrayRS

if Request.QueryString("tipo") = "grupo" then
	mail_iD= " select id_dest_lista from rep_lista_mail where id_lista = "&Request.QueryString("liste")
else
	if IsNumeric(Request.QueryString("liste")) then
		mail_id = "  select id_dest from rep_dest_mail where id_dest_mail = " & Request.QueryString("liste")
	else	
		for i = 0 to UBound(tab_contactos)
			mail_id = mail_id & "," & tab_contactos(i)
		next
		mail_id= mid(mail_id,2)
	end if 
end if 

'	<- CHG-DESA-30062021-01
if mail_id <> "" then
	SQL = " select nombre, mail, decode(client_num, 9929,'Logis',client_num) as client_num " & _
		  " , decode(tercero, 1, 'Si', '') as tercero " & _
		  " From rep_mail " & _
		  " Where id_mail  in (" & mail_id & ")" & _
		  " and status = 1 " & _
		  " order by client_num, tercero desc, nombre "
arrayRS = GetArrayRS(SQL)
'	response.Write SQL
end if
'	CHG-DESA-30062021-01 ->

if not IsArray(arrayRS) then
	Response.Write "no contactos...."
	Response.End 
end if
%>
			<br />
			<table border="0" width="350" class="tblBody">
				<thead>
					<tr class="trHeader">
						<td>Nombre</td>
						<td>Correo</td>
						<%if Request.QueryString("liste") = "1" then%>
							<td>Cliente</td>
							<td>Tercero</td>
						<%end if%>
					</tr>
				</thead>
				<tbody id="tbResult">
					
<%
'for i= 0 to UBound(arrayRS,2)
'	if i mod 2 = 0 then
'		Response.Write "<tr class='trBgColor'>" & vbCrLf & vbTab 
'	else
'		Response.Write "<tr>" & vbCrLf & vbTab 
'	end if
'	Response.Write "<td>" & arrayRS(0,i) &"</td>"  & vbCrLf & vbTab
'	Response.Write "<td><a href=""mailto:" & arrayRS(1,i) &""">" & arrayRS(1,i) &"</a></td>"  & vbCrLf & vbTab
'	if Request.QueryString("liste") = "1" then
'		Response.Write "<td>" & arrayRS(2,i) &"</td>"  & vbCrLf & vbTab
'		Response.Write "<td>" & arrayRS(3,i) &"</td>"  & vbCrLf 
'	end if	
'	Response.Write "</tr>" & vbCrLf 
'next
%>
				</tbody>
			</table>
		</center>
	</body>
</html>
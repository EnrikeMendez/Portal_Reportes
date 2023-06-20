<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Agregacion de reportes
Response.Expires = 0
call check_session()


Select Case Request.Form("Etape")
	Case ""
		dim msg, SQL, arrayRS, i, j, arrayRS2
		%>
		<html>
		<head>
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

                        tmp_ws();
                    }
                );

                function tmp_ws() {
                    const xhr = new XMLHttpRequest('select_activos');
                    var mail_list = document.getElementById('mail_list').value;
                    var id_cliente = document.getElementById('id_cliente').value;
                    const url = urlWebService + "GetArmarCorreos?mail_list=" + mail_list + "&id_client=" + id_cliente;
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
                    var info = objResult.GetArmarCorreosResult;
                    var arrayRS3 = JSON.parse(info);

                    var i = 0;
                    var htmlTable = "";
                    var SQL = "";
                    var bandera = 0;
                    $("#tbResult").empty();
                    if (arrayRS3.length == 0) {
                        htmlTable = htmlTable + "<tr class='center'>";
                        htmlTable = htmlTable + "	<td colspan='9' class='center'>";
                        htmlTable = htmlTable + "		No hay contactos.";
                        htmlTable = htmlTable + "<br>Agregar los <a href=mail.asp>aqui</a>."
                        htmlTable = htmlTable + "	</td>";
                        htmlTable = htmlTable + "</tr>";
                    } else {
                        var j = 0;
                        for (i = 0; i < arrayRS3.length; i++) {
                            htmlTable = "";
                            htmlTable = htmlTable + "<tr bgcolor='FFFFEE'>"
                            htmlTable = htmlTable + "<td><input type=checkbox name=id_mail value=" + arrayRS3[i].ID_MAIL + " " + arrayRS3[i].CHECKED + "></td>"
                            j = j + 1;
                            htmlTable = htmlTable + "<td>" + j + "</td>";
                            htmlTable = htmlTable + "<td>" + arrayRS3[i].NOMBRE + "</td>";
                            htmlTable = htmlTable + "<td><a href=" + "mailto:" + arrayRS3[i].MAIL + "" + ">" + arrayRS3[i].MAIL + "</a></td>";
                            htmlTable = htmlTable + "<td>" + arrayRS3[i].CLIENT_NUM + "</td>";
                            if (arrayRS3[i].TERCERO == null) {
                                htmlTable = htmlTable + "<td></td>";
                            } else {
                                htmlTable = htmlTable + "<td>" + arrayRS3[i].TERCERO + "</td>";
                            }

                            htmlTable = htmlTable + "</tr>";
                            $("#tbResult").append(htmlTable);
                        }
                    }
                }
                function validacion() {
                    const xhr = new XMLHttpRequest('select_activos');
                    var mail_list = document.getElementById('mail_list').value;
                    var action = document.getElementById('id_accion').value;
                    const url = urlWebService + "GetValidaCorreos?mail_list=" + mail_list;
                    var someHandler = "ok";

                    xhr.onreadystatechange = function () {
                        if (xhr.readyState == XMLHttpRequest.DONE) {
                            mostrarResultadoValidacion(xhr.responseText);
                        }

                    }

                    xhr.open("GET", url, true);
                    xhr.send();
                }

                function mostrarResultadoValidacion(wsResponseText) {
                    var objResult = JSON.parse(wsResponseText);
                    var info = objResult.GetValidaCorreosResult;
                    var arrayRS3 = JSON.parse(info);
                    var i = 0;
                    var htmlTable = "";
                    var SQL = "";
                    var bandera = 0;
                    if (arrayRS3.length == 0) {
                        location.href = ('/menu.asp?msg=' + "- escoge al menos un contacto.");
                    } else {
                        txt = "Datos del Reporte: " + arrayRS3[0].ORIGINALES;
                        traerOriginales();
                    }

                }

                function traerOriginales() {
                    const xhr = new XMLHttpRequest('select_activos');
                    var mail_list = document.getElementById('mail_list').value;
                    const url = urlWebService + "GetValidaCorreosOriginales?mail_list=" + mail_list;
                    var someHandler = "ok";

                    xhr.onreadystatechange = function () {
                        if (xhr.readyState == XMLHttpRequest.DONE) {
                            mostrarResultadoOriginales(xhr.responseText);
                        }

                    }

                    xhr.open("GET", url, true);
                    xhr.send();
                }

                function mostrarResultadoOriginales(wsResponseText) {
                    var objResult = JSON.parse(wsResponseText);
                    var info = objResult.GetValidaCorreosOriginalesResult;
                    var arrayRS3 = JSON.parse(info);
                    var i = 0;
                    var htmlTable = "";
                    var SQL = "";
                    var bandera = 0;
                    if (arrayRS3.length == 0) {
                        location.href = ('/menu.asp?msg=' + "- escoge al menos un contacto.");
                    } else {
                        txt = arrayRS3[0].ORIGINALES;
                        elimarRegistros();

                    }

                }



                function elimarRegistros() {
                    const xhr = new XMLHttpRequest('select_activos');
                    var mail_list = document.getElementById('mail_list').value;
                    const url = urlWebService + "GetEliminaOriginales?mail_list=" + mail_list;

                    var someHandler = "ok";

                    xhr.onreadystatechange = function () {
                        validaInsert();
                    }
                    xhr.open("GET", url, true);
                    xhr.send();

                }

                function insertaNuevaLista(nameobjeto, contenareglo) {
                    const xhr = new XMLHttpRequest('select_activos');
                    const url = urlWebService + "GetInsertaRegistro?nameobjeto=" + nameobjeto + "&contenareglo=" + contenareglo;
                    var someHandler = "ok";

                    xhr.onreadystatechange = function () {
                        location.href = ('/menu.asp?msg=' + "Los contactos fueron modificados.");
                    }
                    xhr.open("GET", url, true);
                    xhr.send();
                }








            </script>
		<input type="hidden" id="id_cliente" value=<%=Request.Form("id_client")%>
		<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
		<title>Modificacion de lista de correo</title>
		</head>
		<body>
		<%call print_style()
		'SQL = "  select distinct id_mail, nombre, mail, decode(client_num, 9929,'Logis',client_num) as client_num  " & VbCrLf 
		'SQL = SQL & "  , decode(tercero, 1, 'Si', '') as tercero, decode(id_dest_mail, " & SQLEscape(Request.Form("mail_list")) & ", 'checked') checked " & VbCrLf 
		'SQL = SQL & "  From rep_mail, rep_dest_mail " & VbCrLf 
		'SQL = SQL & "  Where id_dest(+) = id_mail " & VbCrLf 
		'SQL = SQL & "  and client_num in ('" & SQLEscape(Request.Form("id_client")) & "', 9929) " & VbCrLf 
		'SQL = SQL & "  and status = 1 " & VbCrLf 
		'SQL = SQL & "  order by client_num, tercero desc, nombre, checked  "
		
		'Response.Write SQL
		
		'arrayRS = GetArrayRS(SQL)
		'if not IsArray(arrayRS) then
		'	Response.Write "No hay contactos."
		'	Response.Write "<br>Agregar los <a href=mail.asp>aqui</a>."
		'	Response.End 
		'end if

		%>		
		<table border=0 >
		<tr bgcolor=goldenrod>
			<th colspan=7>Seleciona los contactos :</th>
		</tr>
		<tr><td colspan=7><br></td></tr>
		<tr bgcolor="goldenrod" align=center>
			<td colspan=2>.</td>
			<td>Nombre</td>
			<td>Correo</td>
			<td>Cliente</td>
			<td>Tercero</td>
		</tr>
			<tbody id="tbResult"></tbody>
		</table>		
		<!--<form name="form_2"  action="<%=asp_self()%>" method="post" onsubmit="return ValidateForm(this,'id_mail')">-->
		<script language = "Javascript">
		<!-- 
    /**
     * DHTML check all/clear all links script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
     */
    //var form='form_name' //Give the form name here
    function SetChecked(val, chkName, form) {
        var i = 0;
        var nameobjeto;
        nameobjeto = document.getElementsByName("id_mail");
        for (i = 0; i < nameobjeto.length; i++) {
            if (nameobjeto[i] != null) {
                nameobjeto[i].checked = true;
            }
        }
    }
    function UnChecked(val, chkName, form) {
        var i = 0;
        var nameobjeto;
        nameobjeto = document.getElementsByName("id_mail");
        for (i = 0; i < nameobjeto.length; i++) {
            if (nameobjeto[i] != null) {
                nameobjeto[i].checked = false;
            }
        }
    }

    function validaInsert() {
        var i = 0;
        var nameobjeto;
        var contenareglo = "";
        var mail_list = document.getElementById('mail_list').value;
        nameobjeto = document.getElementsByName("id_mail");
        for (i = 0; i < nameobjeto.length; i++) {
            if (nameobjeto[i].checked == true) {
                contenareglo = contenareglo + nameobjeto[i].value + ",";
            }
        }
        insertaNuevaLista(mail_list, contenareglo);
    }



    function ValidateForm(dml) {
        len = dml.elements.length;
        var i = 0;
        var mail_ok = "- escoge al menos un contacto.\n";
        for (i = 0; i < len; i++) {
            if ((dml.elements[i].name == 'id_mail') && (dml.elements[i].checked == 1)) mail_ok = "";
        }
        if (mail_ok != "") {
            alert("Verifica los contactos :\n" + mail_ok + mail_error);
            return false;
        }
        return true;
    }
		// -->

        </script>
		<%
		'j=0
		'for i = 0 to UBound(arrayRS,2)
		'	Response.Write "<tr"
		'	if j mod 3 = 0 then Response.Write  " bgcolor=""FFFFEE"""
		'	Response.Write ">" & vbCrLf & vbTab 
		'	Response.Write "<td><input type=checkbox name=id_mail value="& arrayRS(0, i) & " " & arrayRS(5, i) & "></td>"
		'	Response.Write "<td>" & j+1 & "</td>" & vbCrLf & vbTab 
		'	Response.Write "<td>" & arrayRS(1, i) & "</td>" & vbCrLf & vbTab  
		'	Response.Write "<td><a href=""mailto:" & arrayRS(2, i) & """>" & arrayRS(2, i) & "</a></td>" & vbCrLf & vbTab  
		'	Response.Write "<td>" & arrayRS(3, i) & "</td>" & vbCrLf & vbTab  
		'	Response.Write "<td>" & arrayRS(4, i) & "</td>" & vbCrLf 
		'	Response.Write "</tr>" & vbCrLf 
		'	do while i < UBound(arrayRS,2)
		'		if CInt(arrayRS(0,i)) <> CInt(arrayRS(0,i+1)) then exit do
		'		i = i + 1
		'	loop
		'	j=j+1
		'next
		%>
		<tr>
			<td colspan=6>
			<a href="javascript:SetChecked(1,'id_mail','tbResult')"><font face="Arial, Helvetica, sans-serif" size="0">Check All</font></a>
			&nbsp;&nbsp;
			<a href="javascript:UnChecked(0,'id_mail','tbResult')"><font face="Arial, Helvetica, sans-serif" size="0">Clear All</font></a>
			</td>
		</tr>
		<tr>
			<td colspan=6><br></td>
		</tr>
		<tr>
			<td colspan=6>
			<input type="hidden" name=etape value=1>
			<input type="hidden" id="mail_list" name=mail_list value=<%=Request.Form("mail_list")%> />	
				
			<input type="hidden" id="id_accion" name=id_accion value=<%=Request.Form("id_accion")%> />
			
			<button class=buttonsOrange onclick="validacion()">Validar</button>

			</td>
		</tr>
		
		</form>
		</table>
		




</body>
		</html>
		<%

end select

%>
</body>
</html>
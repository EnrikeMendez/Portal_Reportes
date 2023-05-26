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
		<title>Modificacion de lista de correo</title>
		</head>
		<body>
		<%call print_style()
		SQL = "  select distinct id_mail, nombre, mail, decode(client_num, 9929,'Logis',client_num) as client_num  " & VbCrLf 
		SQL = SQL & "  , decode(tercero, 1, 'Si', '') as tercero, decode(id_dest_mail, " & SQLEscape(Request.Form("mail_list")) & ", 'checked') checked " & VbCrLf 
		SQL = SQL & "  From rep_mail, rep_dest_mail " & VbCrLf 
		SQL = SQL & "  Where id_dest(+) = id_mail " & VbCrLf 
		SQL = SQL & "  and client_num in ('" & SQLEscape(Request.Form("id_client")) & "', 9929) " & VbCrLf 
		SQL = SQL & "  and status = 1 " & VbCrLf 
		SQL = SQL & "  order by client_num, tercero desc, nombre, checked  "
		
		'Response.Write SQL
		
		arrayRS = GetArrayRS(SQL)
		if not IsArray(arrayRS) then
			Response.Write "No hay contactos."
			Response.Write "<br>Agregar los <a href=mail.asp>aqui</a>."
			Response.End 
		end if

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
		<form name="form_2"  action="<%=asp_self()%>" method="post" onsubmit="return ValidateForm(this,'id_mail')">
		<script language = "Javascript">
		<!-- 
		/**
		 * DHTML check all/clear all links script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
		 */
		//var form='form_name' //Give the form name here
		function SetChecked(val,chkName,form) {
			dml=document.forms[form];
			len = document.forms[form].elements.length;
			var i=0;
			for( i=0 ; i<len ; i++) {
				if (dml.elements[i].name==chkName) {
					dml.elements[i].checked=val;
					}
				}
		}

		function ValidateForm(dml){
			len = dml.elements.length;
			var i=0;
			var mail_ok = "- escoge al menos un contacto.\n";
			for( i=0 ; i<len ; i++) {
				if ((dml.elements[i].name=='id_mail') && (dml.elements[i].checked==1)) mail_ok ="";
				}
			if (mail_ok != "")
				{alert("Verifica los contactos :\n" + mail_ok + mail_error);
				 return false;
				}
			return true;
		}
		// -->

		</script>
		<%
		j=0
		for i = 0 to UBound(arrayRS,2)
			Response.Write "<tr"
			if j mod 3 = 0 then Response.Write  " bgcolor=""FFFFEE"""
			Response.Write ">" & vbCrLf & vbTab 
			Response.Write "<td><input type=checkbox name=id_mail value="& arrayRS(0, i) & " " & arrayRS(5, i) & "></td>"
			Response.Write "<td>" & j+1 & "</td>" & vbCrLf & vbTab 
			Response.Write "<td>" & arrayRS(1, i) & "</td>" & vbCrLf & vbTab  
			Response.Write "<td><a href=""mailto:" & arrayRS(2, i) & """>" & arrayRS(2, i) & "</a></td>" & vbCrLf & vbTab  
			Response.Write "<td>" & arrayRS(3, i) & "</td>" & vbCrLf & vbTab  
			Response.Write "<td>" & arrayRS(4, i) & "</td>" & vbCrLf 
			Response.Write "</tr>" & vbCrLf 
			do while i < UBound(arrayRS,2)
				if CInt(arrayRS(0,i)) <> CInt(arrayRS(0,i+1)) then exit do
				i = i + 1
			loop
			j=j+1
		next
		%>
		<tr>
			<td colspan=6>
			<a href="javascript:SetChecked(1,'id_mail','form_2')"><font face="Arial, Helvetica, sans-serif" size="0">Check All</font></a>
			&nbsp;&nbsp;
			<a href="javascript:SetChecked(0,'id_mail','form_2')"><font face="Arial, Helvetica, sans-serif" size="0">Clear All</font></a>
			</td>
		</tr>
		<tr>
			<td colspan=6><br></td>
		</tr>
		<tr>
			<td colspan=6>
			<input type="hidden" name=etape value=1>
			<input type="hidden" name=mail_list value=<%=Request.Form("mail_list")%>>
			<input type=submit class=buttonsOrange value=Validar><br><br>
			</td>
		</tr>
		
		</form>
		</table>
		




</body>
		</html>
		<%

case "1"
	Dim id_mail_ok, rst, mails, txt, x
	set rst = Server.CreateObject("ADODB.Recordset")
	id_mail_ok = SQLEscape(Request.Form("mail_list"))
	
	SQL = "SELECT ID_CRON, NAME FROM REP_DETALLE_REPORTE WHERE MAIL_OK = '" & id_mail_ok & "'"
	arrayRS = GetArrayRS(SQL)
	if IsArray(arrayRS) then
		txt = "Datos del Reporte: " & arrayRS(0,0) & " - " & arrayRS(1,0)  & "	"
	end if
	
	SQL = "SELECT LISTAGG(ID_DEST,',') WITHIN GROUP (ORDER BY  ID_DEST) ORIGINALES FROM rep_dest_mail WHERE ID_DEST_MAIL = " & id_mail_ok
	arrayRS = GetArrayRS(SQL)
	if IsArray(arrayRS) then
		SQL = "SELECT ID_MAIL, NOMBRE FROM REP_MAIL WHERE ID_MAIL IN (" & SQLescape(arrayRS(0,0)) & ")"
		arrayRS = GetArrayRS(SQL)
		
		if IsArray(arrayRS) then
			txt = txt & "	||	Correos Originales: "  & "	"
			for x = 0 to UBound(arrayRS,2)
				 txt = txt & arrayRS(0,x) & " - " & arrayRS(1,x) & "	"
			next
		end if
	end if
	
	'borramos todos los contactos de la lista y luego volvemos a insertar los nuevos
	SQL = "delete from rep_dest_mail where id_dest_mail = "& id_mail_ok
	'Response.Write SQL & "<br>"
	
	rst.Open SQL, Connect(), 0, 1, 1
	
	
	'insercion de los contactos
	mails = split(Request.Form("id_mail"),",")
	for i= 0 to UBound(mails)
		SQL = " insert into rep_dest_mail (id_dest_mail, id_dest) " & _
			  " values ('"& id_mail_ok &"','"& mails(i) &"' ) "
		'Response.Write SQL & "<br>"
		rst.Open SQL, Connect(), 0, 1, 1
	next
    
	
	SQL = "SELECT LISTAGG(ID_DEST,',') WITHIN GROUP (ORDER BY  ID_DEST) ORIGINALES FROM rep_dest_mail WHERE ID_DEST_MAIL = " & id_mail_ok
	arrayRS = GetArrayRS(SQL)
	if IsArray(arrayRS) then
		SQL = "SELECT ID_MAIL, NOMBRE FROM REP_MAIL WHERE ID_MAIL IN (" & SQLescape(arrayRS(0,0)) & ")"
		arrayRS = GetArrayRS(SQL)
		
		if IsArray(arrayRS) then
			txt = txt & "	||	Correos Nuevos: " & "	"
			for x = 0 to UBound(arrayRS,2)
				 txt = txt & arrayRS(0,x) & " - " & arrayRS(1,x) & "	"
			next
		end if
	end if
	
	if txt <> "" then
		EscribeLog(txt)
	end if
	
	Response.Redirect "menu.asp?msg=" & Server.URLEncode ("Los contactos fueron modificados.")
end select

%>
</body>
</html>
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
		<title>Gestion de correos</title>
		</head>
		<body>
		<%
		call print_style()
	if Request.QueryString("accion") = "mod" then
		%>
		<table width=450 border=0 >
		<tr>
			<td colspan=2><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<tr bgcolor=goldenrod>
			<th>Lista</th>
			<th>reporte asociado</th>
		</tr>
	<%
	SQL = " select list.id_lista, list.nombre, rep.name " & _
			" from rep_lista list " & _
			" , rep_lista_reporte list_rep " & _
			" , rep_detalle_reporte rep " & _
			" where list.ID_LISTA = list_rep.ID_LISTA(+) " & _
			" and list_rep.ID_REPORTE = rep.ID_CRON(+) "
	arrayRS = GetarrayRS(SQL)
	
	if IsArray(arrayRS) then
		for i = 0 to UBound(arrayRS,2)
			Response.Write "<tr valign=top ><td><a href=""javascript:modif_list("& arrayRS(0,i) &");"">" & arrayRS(1,i) & "</a></td>" & vbCrLf
			Response.Write "<td>"
			do while i < UBound(arrayRS,2)
				Response.Write arrayRS(2,i)
				if  CInt(arrayRS(0,i)) <> CInt(arrayRS(0,i+1)) then exit do
				Response.Write "<br>"
				i = i + 1
			loop
			Response.Write "</td>"
			Response.Write "</tr>" & vbCrLf
		next
		
	
	end if
	%>
	<script language=javascript>
		function modif_list(id_list) {
			document.modif_list.id_lista.value = id_list;
			document.modif_list.submit() ;
		}
	</script>
	<form name="modif_list" action="<%=asp_self()%>" method="post">
		<input type="hidden" name="id_lista" value="">
		<input type="hidden" name="accion" value="mod">
		<input type="hidden" name="etape" value="1">
	</form>
	<%else
		%>

		<table width=450 border=0 >
		<tr>
			<td colspan=2><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
		<tr bgcolor=goldenrod>
			<th Align=left>Selecionar un cliente :</th>
		</tr>
		<script LANGUAGE="JavaScript">
		function check_data()
			{
			if (document.form_1.cli_num.value == "")
				{alert("Favor de capturar un numero de cliente.");}
			else document.form_1.submit();  		
			}
		</script>
		<tr>
			<td>Entra el numero de cliente para desplegar la lista de los contactos.
				<br>Si no agregaste los contactos, dale un clic <a href=mail.asp>aqui</a>.
				<br>
				<br>Una lista sirve para juntar contactos y facilitar la eleccion de estos a la creacion de un nuevo reporte.
			</td>
		</tr>
		<form name="form_1"  action="<%=asp_self()%>" method="post">
		<tr>
			<td>
				<br>Numero de cliente :
				<br><input type=text name=cli_num size=4 class=light>
			</td>
		</tr>
		<tr>
			<td align=left>
			<input type="hidden" name=etape value=1>
			<input type="hidden" name=accion value=<%=Request.QueryString("accion")%>>
			<input type=button onClick="check_data();" class=buttonsOrange value=Validar id=button1 name=button1><br><br>
			</td>
		</tr>
		</form>
		
		</table>
		
	<%end if
		
case "1"%>
		
		
		<html>
		<head>
		<title><%if Request.Form("accion") = "mod" then 
					Response.Write "Modificacion"
				 else
					Response.Write "Creacion"
				 end if	
					%> de lista de correo</title>
		</head>
		<body>
		<%call print_style()
		if Request.Form("accion") = "mod" then 
			SQL = "   select distinct mail.id_mail, mail.nombre, mail, decode(client_num, 9929,'Logis',client_num) as client_num   " & VbCrLf 
			SQL = SQL & "   , decode(tercero, 1, 'Si', '') as tercero , decode(listmail.ID_DEST_LISTA, mail.id_mail, 'checked') checked " & VbCrLf 
			SQL = SQL & "   , list.NOMBRE, list.ID_CLIENTE " & VbCrLf 
			SQL = SQL & "   From rep_mail mail " & VbCrLf 
			SQL = SQL & "   , rep_dest_mail destmail " & VbCrLf 
			SQL = SQL & "   , rep_lista_mail listmail " & VbCrLf 
			SQL = SQL & "   , rep_lista_reporte listrep " & VbCrLf 
			SQL = SQL & "   , rep_lista list " & VbCrLf 
			SQL = SQL & "   Where destmail.id_dest(+) = mail.id_mail  " & VbCrLf 
			SQL = SQL & "   and list.ID_LISTA = listrep.ID_LISTA(+) " & VbCrLf 
			SQL = SQL & "   and list.ID_LISTA = listmail.ID_LISTA(+) " & VbCrLf 
			SQL = SQL & "   and list.ID_LISTA = " & Request.Form("id_lista") & VbCrLf 
			SQL = SQL & "   and client_num in (list.ID_CLIENTE, 9929)  " & VbCrLf 
			SQL = SQL & "   and mail.status = 1  " & VbCrLf 
			SQL = SQL & "   order by client_num, tercero desc, mail.nombre "

		else
			SQL = "  select distinct id_mail, nombre, mail, decode(client_num, 9929,'Logis',client_num) as client_num  " & VbCrLf 
			SQL = SQL & "  , decode(tercero, 1, 'Si', '') as tercero " & VbCrLf 
			SQL = SQL & "  From rep_mail " & VbCrLf 
			SQL = SQL & "  Where client_num in ('" & SQLEscape(Request.Form("cli_num")) & "', 9929) " & VbCrLf 
			SQL = SQL & "  and status = 1 " & vbCrLf
			SQL = SQL & "  order by client_num, tercero desc, nombre  "
		end if
		'Response.Write SQL
		arrayRS = GetArrayRS(SQL)
		if not IsArray(arrayRS) then
			Response.Write "No hay contactos."
			Response.Write "<br>Agregar los <a href=mail.asp>aqui</a>."
			Response.End 
		end if

		%>
		<form name="form_2"  action="<%=asp_self()%>" method="post" onsubmit="return ValidateForm(this,'id_mail')">
		<table border=0 >
		<tr bgcolor=goldenrod>
			<th colspan=7>Nombre de la lista :</th>
		</tr>
		<tr><td colspan=7><input type=text name=lista_nombre class=light maxlength=50 size=30 value="<%if Request.Form("accion") = "mod" then Response.Write HTMLEscape(arrayRS(6,0))%>"></td></tr>
		<tr><td colspan=7><br></td></tr>
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
			var nombre_lista = "- selecionar un nombre de lista.\n";
			for( i=0 ; i<len ; i++) {
				if ((dml.elements[i].name=='id_mail') && (dml.elements[i].checked==1)) mail_ok ="";
				if ((dml.elements[i].name=='lista_nombre') && (dml.elements[i].value!='')) nombre_lista ="";
				}
			if (mail_ok != "")
				{alert("Verifica los datos :\n" + mail_ok + nombre_lista);
				 return false;
				}
			else if (mail_ok != "")
			return true;
		}
		// -->

		</script>
		<%
		j=0
		for i = 0 to UBound(arrayRS,2)
			Response.Write "<tr "
			if j mod 3 = 0 then Response.Write  " bgcolor=""FFFFEE"""
			Response.Write ">" & vbCrLf & vbTab 
			Response.Write "<td><input type=checkbox name=id_mail value="& arrayRS(0, i) 
			if Request.Form("accion") = "mod" then Response.Write " " & arrayRS(5,i)
			Response.Write "></td>"
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
			<input type="hidden" name=etape value=2>
			<input type="hidden" name=accion value=<%=Request.Form("accion")%>>
			<input type="hidden" name=id_lista value=<%=Request.Form("id_lista")%>>
			<input type="hidden" name=cli_num value=<%if Request.Form("accion") = "mod" then
														Response.Write arrayRS(7,0) 
														else
														Response.Write HTMLEscape(Request.Form("cli_num"))
														end if%>>
			<input type=submit class=buttonsOrange value=Validar><br><br>
			</td>
		</tr>
		
		</form>
		</table>



<%case "2"

	Dim rst, mails, id_lista, arrayReporte

	set rst = Server.CreateObject("ADODB.Recordset")

	'''''''''''''''''''''''''''''''''''
	
	'si pas de liste_id
	
	'borramos todos los contactos de la lista y luego volvemos a insertar los nuevos
	
	
	'_rst.Open SQL, Connect(), 0, 1, 1
	
	'''''''''''''''''''''''''''''''''''''
	
	if Request.Form("id_lista") = "" then
		SQL = "select SEQ_LISTA.nextval from dual "
		arrayRS = GetArrayRS(SQL)
		id_lista = arrayRS(0,0)
		
		'creacion de la nueva lista
		SQL = "insert into rep_lista (id_lista, nombre, created_by, date_created, id_cliente) " & _
			  " values  ("& id_lista &", '"& SQLEscape(Request.Form("lista_nombre")) &"', '"& Session("array_user")(0,0) &"', sysdate, "&(Request.Form("cli_num"))&")"
		'Response.Write SQL
		rst.Open SQL, Connect(), 0, 1, 1	


		'agregamos los contactos
		mails = split(Request.Form("id_mail"),",")
		for i= 0 to UBound(mails)
			SQL = " insert into REP_LISTA_MAIL (id_lista, id_dest_lista) " & _
				  " values ('"& id_lista &"','"& mails(i) &"' ) "
			rst.Open SQL, Connect(), 0, 1, 1
		next

		msg = "Lista creada."
	else 
		SQL = " select 1 from rep_lista " & _
			  " where id_lista = "& SQLEscape(Request.Form("id_lista")) &" " & _
			  " and nombre != '"& SQLEscape(Request.Form("lista_nombre")) & "'" 
		'Response.Write SQl
		arrayRS = GetArrayRS(SQL)
		
		if IsArray(arrayRS) then
			'el nombre de la lista se modifico, hay que actualizar los datos
			SQL = "update rep_lista set nombre = "& SQLEscape(lista_nombre) & _
				  " , DATE_MODIFIED = sysdate, MODIFIED_BY =  "& Session("array_user")(0,0) & _
				  " where id_lista = " & id_lista 
			'Response.Write "<br>Update nom liste : " & SQL
			rst.Open SQL, Connect(), 0, 1, 1
		end if	  
		
		
		'borar todos los contactos de la lista en todos los reportes que impacta
		'luego inserar otra vez los contactos por la lista y luego de todas las listas que tocan
		'este reporte
		''''''''''''''''''''''''''
		
		'seleccion de los reportes que impacta la lista
		SQL = "select id_reporte  " & _
			  " from REP_LISTA_REPORTE " & _
			  " where id_lista  = " & SQLEscape(Request.Form("id_lista"))
		arrayReporte = GetArrayRS(SQL)
		
		if IsArray(arrayReporte) then
			'uno o mas reportes usan esta lista, vamos a modificar los contactos.
			for i = 0 to UBound(arrayReporte,2) 
				'borramos los contactos incluidos en las listas que impactan estos reportes
				SQL = "delete from rep_dest_mail  " & VbCrLf 
				SQL = SQL & " where (id_dest || ',' || id_dest_mail) " & VbCrLf 
				SQL = SQL & " in ( " & VbCrLf 
				SQL = SQL & " select distinct id_dest_lista || ',' || rep.mail_ok " & VbCrLf 
				SQL = SQL & " from rep_lista_mail, rep_detalle_reporte rep " & VbCrLf 
				SQL = SQL & " where id_lista in  " & VbCrLf 
				SQL = SQL & " 	  (select id_lista  " & VbCrLf 
				SQL = SQL & " 	  from REP_LISTA_REPORTE " & VbCrLf 
				SQL = SQL & " 	  where id_reporte = "& arrayReporte(0,i)  & VbCrLf 
				SQL = SQL & " 	  ) " & VbCrLf 
				SQL = SQL & " and rep.id_cron = " & arrayReporte(0,i) & ")"		
				
				'Response.Write SQL & "<br><br>"
				rst.Open SQL, Connect(), 0, 1, 1
				
			next
	
			SQL = "delete from rep_lista_mail where id_lista = " & SQLEscape(Request.Form("id_lista"))
			'Response.Write "<br>delete nom liste : " & SQL
			rst.Open SQL, Connect(), 0, 1, 1
		
			'insercion de los contactos
			mails = split(Request.Form("id_mail"),",")
			for i= 0 to UBound(mails)
				SQL = " insert into REP_LISTA_MAIL (id_lista, id_dest_lista) " & _
					  " values ('"& SQLEscape(Request.Form("id_lista"))  &"','"& mails(i) &"' ) "
				'Response.Write SQL & "<br>"
				rst.Open SQL, Connect(), 0, 1, 1
			next			
			
			for i = 0 to UBound(arrayReporte,2) 
				'insertamos otra vez los contactos despues de haber modificado la lista
				SQL = "	insert into rep_dest_mail  " & VbCrLf 
				SQL = SQL & " select distinct mail_ok , listmail.ID_DEST_LISTA " & VbCrLf 
				SQL = SQL & " from REP_DETALLE_REPORTE r " & VbCrLf 
				SQL = SQL & " , REP_LISTA_REPORTE listrep " & VbCrLf 
				SQL = SQL & " , REP_LISTA_MAIL listmail " & VbCrLf 
				SQL = SQL & " where id_cron = " & arrayReporte(0,i)  & VbCrLf 
				SQL = SQL & " and id_cron = listrep.id_reporte " & VbCrLf 
				SQL = SQL & " and listmail.ID_LISTA = listrep.ID_LISTA"	
				SQL = SQL & "  and not exists ( " & VbCrLf 
 	            SQL = SQL & " select null " & VbCrLf 
	            SQL = SQL & " from rep_dest_mail r2 " & VbCrLf 
	            SQL = SQL & " where r2.ID_DEST =  listmail.ID_DEST_LISTA " & VbCrLf 
                SQL = SQL & "   and r2.ID_DEST_MAIL= r.mail_ok  ) " & VbCrLf 
				
				'Response.Write SQL & "<br><br>"
				rst.Open SQL, Connect(), 0, 1, 1
			
			next
		
		else
			SQL = "delete from rep_lista_mail where id_lista = " & SQLEscape(Request.Form("id_lista"))
			'Response.Write "<br>delete nom liste : " & SQL
			rst.Open SQL, Connect(), 0, 1, 1
		
			'insercion de los contactos
			mails = split(Request.Form("id_mail"),",")
			for i= 0 to UBound(mails)
				SQL = " insert into REP_LISTA_MAIL (id_lista, id_dest_lista) " & _
					  " values ('"& SQLEscape(Request.Form("id_lista"))  &"','"& mails(i) &"' ) "
				'Response.Write SQL & "<br>"
				rst.Open SQL, Connect(), 0, 1, 1
			next
		end if
			
		msg = "Lista modificada."
	end if
	

    
	Response.Redirect "menu.asp?msg=" & Server.URLEncode (msg)

end select

%>

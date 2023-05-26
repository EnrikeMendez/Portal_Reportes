<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Agregacion de reportes
Response.Expires = 0
call check_session()


Select Case Request.Form("Etape")
	Case ""
		dim msg, SQL, arrayRS, i, j, k, arrayRS2
		%>
		<html>
		<head>
		<title>Gestion de usuarios</title>
		</head>
		<body>
		<%call print_style()%>
	

		<table width=450 border=0 >
		<tr>
			<td colspan=2><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
		<tr bgcolor=goldenrod>
			<th Align=left>Selecionar un nombre de usuario (puede ser incompleto) :</th>
		</tr>
		<script LANGUAGE="JavaScript">
		function check_data()
			{
			if ((document.form_1.usuario.value == "") || (document.form_1.usuario.value.length < 3))
				{alert("Favor de capturar un nombre de usuario (mas de 3 letras).");}
			else document.form_1.submit();  		
			}
		</script>
		<tr>
			<td>Entra el nombre de usuario para ver todo sus reportes.
			</td>
		</tr>
		<form name="form_1"  action="<%=asp_self()%>" method="post">
		<tr>
			<td>
				<br>Nombre de usuario :
				<br><input type=text name=usuario size=30 class=light>
			</td>
		</tr>
		<tr>
			<td align=left>
			<input type="hidden" name=etape value=1>
			<input type="hidden" name=accion value=<%=Request.QueryString("accion")%>>
			<input type=button onClick="check_data();" class=buttonsOrange value=validar><br><br>
			</td>
		</tr>
		</form>
		
		</table>
		

		
<%case "1"%>
		
		
		<html>
		<head>
		<title>Ver los reportes por usuario.</title>
		</head>
		<body>
		<%call print_style()
		call print_popup()

		SQL = "select distinct mail.id_mail, mail.nombre, mail.mail, decode(mail.client_num, 9929,'Logis',mail.client_num) as client_num     " & VbCrLf 
		SQL = SQL & "     , decode(mail.tercero, 1, 'Si', 'No') as tercero    " & VbCrLf 
		SQL = SQL & "     , rep.name, rep.id_cron, rep.mail_ok, rep.cliente   " & VbCrLf 
		SQL = SQL & "  From rep_mail mail, rep_dest_mail destmail  " & VbCrLf 
		SQL = SQL & "     , rep_detalle_reporte rep   " & VbCrLf 
		SQL = SQL & "     , rep_chron   " & VbCrLf 
		SQL = SQL & "  Where destmail.id_dest = mail.id_mail    " & VbCrLf 
		SQL = SQL & "     and (Upper(mail.nombre) like  Upper('%"& SQLEscape(Request.Form("usuario")) &"%')    " & VbCrLf 
		SQL = SQL & "       OR Upper(mail.MAIL) like  Upper('%"& SQLEscape(Request.Form("usuario")) &"%')    " & VbCrLf 
		SQL = SQL & "     )   " & VbCrLf 
		SQL = SQL & "     and mail.status = 1    " & VbCrLf 
		SQL = SQL & "     and ID_RAPPORT(+) = id_Cron    " & VbCrLf 
		SQL = SQL & "     and active(+) = 1    " & VbCrLf 
		SQL = SQL & "     and destmail.ID_DEST_MAIL = rep.mail_ok  " & VbCrLf 
		SQL = SQL & "     order by  mail.nombre,mail.id_mail, rep.name  " 

			
		'Response.Write SQL
		arrayRS = GetArrayRS(SQL)
		if not IsArray(arrayRS) then
			Response.Write "No hay contactos."
			Response.Write "<br>Agregar los <a href=mail.asp>aqui</a>."
			Response.End 
		end if

		%>
		<script language=javascript>
		function modif_list(id_list, id_client) {
			document.modif_list.mail_list.value = id_list;
			document.modif_list.id_client.value = id_client;
			document.modif_list.submit() ;
		}
		</script>
		<script language=javascript>
		function modif_contactos(id_list) {
			document.modif_contactos.id_lista.value = id_list;
			document.modif_contactos.submit() ;
		}
		</script>
		<script language="JavaScript" src="./include/tigra_tables.js"></script>
		<form name="modif_contactos" action="lista_contactos.asp" method="post">
			<input type="hidden" name="id_lista" value="">
			<input type="hidden" name="accion" value="mod">
			<input type="hidden" name="etape" value="1">
		</form>
		
		<form name="modif_list" action="mail_modif.asp" method="post">
			<input type="hidden" name="mail_list" value="">
			<input type="hidden" name="id_client" value="">
		</form>
		<form name="form_2"  action="<%=asp_self()%>" method="post" onsubmit="return ValidateForm(this,'id_mail')">
		<table border=0 cellpadding=3 cellspacing=0 >
		<tr bgcolor=goldenrod>
			<th>Reportes / contactos :</th>
		</tr>
		<tr><td><br></td></tr>
		<tr bgcolor="goldenrod">
			<th align=Left>Nombre</th>
		</tr>
		<%
		j=0
		for i = 0 to UBound(arrayRS,2)
			Response.Write "<tr valign=top>" & vbCrLf & vbTab 
			Response.Write "<td><a href=""javascript:void(0);"" onmouseover=""return overlib('" & JSescape(arrayRS(2,i)) & "');"" onmouseout=""return nd();"">" & arrayRS(1, i) & "</td>" & vbCrLf & vbTab  
			Response.Write "</tr>" & vbCrLf
			k=0
			Response.Write "<tr><td>" & vbCrLf 
			do while i <= UBound(arrayRS,2)
				if k=0 then
				Response.Write "<table border=1 cellpadding=3 cellspacing=0 width=""100%""  id=select_reporte>" & vbCrLf & vbTab 
				Response.Write "<tr bgcolor=goldenrod><td rowspan=2>Reporte</td><td colspan=3>Listas de contactos :</td></tr>"& vbCrLf & vbTab 
				Response.Write "<tr bgcolor=goldenrod align=center><td>reporte</td><td>en caso de error</td><td>contactos agrupados</td></tr>"& vbCrLf & vbTab 
				end if
				Response.Write "<tr valign=top><td>" & arrayRS(5,i) & "</td>" & vbCrLf & vbtab
				Response.Write "<td><a href=""javascript:modif_list("& arrayRS(7,i) &","& arrayRS(8,i) &")"">Modificar</a></td>" & vbCrLf & vbtab
			
				SQL = " select mail_error, rep.cliente " & VbCrLf 
				SQL = SQL & " from rep_detalle_reporte rep " & VbCrLf 
				SQL = SQL & " 	, rep_dest_mail destmail " & VbCrLf 
				SQL = SQL & " where destmail.id_dest = " & arrayRS(0,i) & VbCrLf 
				SQL = SQL & " 	and rep.mail_error = destmail.id_dest_mail " & vbCrLf
				SQL = SQL & " 	and rep.ID_CRON = " & arrayRS(6,i)
			
				arrayRS2 = GetArrayRS(SQL)
				Response.Write "<td align=center>"
				if IsArray(arrayRS2) then
					'por la lista de contacto en caso de error, solo veemos los contactos Logis...
					Response.Write "<a href=""javascript:modif_list("& arrayRS2(0,0) &",9929);"">Modificar</a>"
				end if	
				Response.Write "&nbsp;</td>" & vbCrLf & vbTab 
			
				SQL = " select list.ID_LISTA, list.NOMBRE  " & VbCrLf 
				SQL = SQL & " from rep_lista list " & VbCrLf 
				SQL = SQL & " , rep_lista_reporte listrep " & VbCrLf 
				SQL = SQL & " , rep_lista_mail listmail " & VbCrLf 
				SQL = SQL & " where list.ID_LISTA = listrep.ID_LISTA " & VbCrLf 
				SQL = SQL & " and listmail.ID_LISTA = list.ID_LISTA " & VbCrLf 
				SQL = SQL & " and listmail.ID_DEST_LISTA = " & arrayRS(0,i) & VbCrLf 
				SQL = SQL & " and listrep.ID_REPORTE = " & arrayRS(6,i)
				arrayRS2 = GetArrayRS(SQL)
				
				Response.Write "<td>"
				
				if IsArray(arrayRS2) then
					for j = 0 to UBound(arrayRS2,2)
						Response.Write "<a href=""javascript:modif_contactos(" & arrayRS2(0,j) & ");"">" & arrayRS2(1,j) 
						if j <> UBound(arrayRS2,2) then Response.Write vbcrlf & "<br>"
					next					
				end if
				Response.Write "&nbsp;</td>" & vbCrLf & vbTab 
				Response.Write "</tr>" & vbCrLf & vbCrLf
				if i = UBound(arrayRS,2) then 
					Response.Write "</table>" & vbCrLf
					exit Do
				end if
				if CInt(arrayRS(0,i))<> CInt(arrayRS(0,i+1)) then 
					Response.Write "</table>" & vbCrLf
					exit do
				end if
				i = i +1
				k = k + 1
			loop
			Response.Write "</td></tr>" & vbCrLf & vbCrLf 
			Response.Write "<tr><td><br><br></td></tr>"
			j=j+1
		next
				%>
		<tr>
			<td><br></td>
		</tr>
		<tr>
			<td>
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

		<script language="JavaScript">
		<!--
			tigra_tables('select_reporte', 2, 0, '#ffffff', '#ffffcc', '#ffcc66', '#cccccc');
		// -->
		</script>

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
				SQL = SQL & " from REP_DETALLE_REPORTE " & VbCrLf 
				SQL = SQL & " , REP_LISTA_REPORTE listrep " & VbCrLf 
				SQL = SQL & " , REP_LISTA_MAIL listmail " & VbCrLf 
				SQL = SQL & " where id_cron = " & arrayReporte(0,i)  & VbCrLf 
				SQL = SQL & " and id_cron = listrep.id_reporte " & VbCrLf 
				SQL = SQL & " and listmail.ID_LISTA = listrep.ID_LISTA"	
				
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

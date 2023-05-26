<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Deactivacion anomalias
Response.Expires = 0
call check_session()
'date format US
Session.LCID = 1033


Select Case Request("Etape")
	Case ""
%>
		<html>
		<head>
		<title>Desactivacion anomalias</title>
		</head>
		<body>
		<%
		call print_style()
		Dim SQL, Array_anom, i
		SQL ="  select NUM_ERROR, NUM_ERROR || '. ' || DESCRIPCION " & _
			 "  from REP_ANOMALIAS_CATALOGO " & _
			 "  where NUM_ERROR <> 4 "& _
			 " order by 1"

		Array_anom= GetArrayRS(SQL)

		if not IsArray(Array_anom) then
			Response.Write "Ningun anomalia en la base."
			Response.End 
		end if
		'Response.Write "taille: " &Ubound(Array_conf,2)

		%><table border=0 align=left cellspacing=3>
		<tr>
			<td colspan=2><a href=menu.asp>Menu general</a><br>
			<a href=lista_desact.asp>Lista de folios desactivados</a><br><br></td>
		</tr>
		<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
		<tr bgcolor=goldenrod>
			<th colspan=2>Lista de anomalias :</th>
		</tr>
		<form name="anom_form" action="<%=asp_self()%>" method="post"><%
		for i = 0 to UBound(Array_anom , 2)
			%><tr>
				<td><input type=radio name=num_error value=<%=Array_anom(0,i)%>></td>
				<td><%=Array_anom(1,i)%></td>
			  </tr>
		<%next%>
		<tr><td><br></td></tr>
		<tr bgcolor=goldenrod>
			<th colspan=2>Numero de folios :</th>
		</tr>
		<tr>
			<td colspan=2><input type=text name=folio size=10 class="light">
			&nbsp;&nbsp;<input type=checkbox name=reactivar> <i>Para reactivar una anomalia desactivada.</i></td>
		</tr>
		<tr>
			<td colspan=2 align=top>Observaciones : <br><textarea name=observaciones cols=40 rows=3></textarea></td>
		</tr>
		<tr>
			<td colspan=2><input type=button value=validar class=buttonsorange onclick=javascript:validate_form();>
		<input type=hidden name=etape></td>
		</tr>
		<script language=javascript>
			function validate_form()
			{var msg = "- selecionar un error\n";
			for(var i=0 ; i<anom_form.elements.length ; i++) {
				if ((anom_form.elements[i].name=='num_error') && (anom_form.elements[i].checked==1)) msg="";
				} 	
			if (anom_form.folio.value == "")
				{msg = msg+ "- entrar un numero de folio.\n"};	
				
			if (anom_form.reactivar.checked)
				{anom_form.etape.value=3;}
			else {anom_form.etape.value=1;
				if (anom_form.observaciones.value == "")
					{msg = msg+ "- entrar una observacion.\n"};
			}	
				;
			if (msg != "")
				{alert(msg);
				return false;
				}
			else {anom_form.submit();};
			}

		</script>
		</form>
		</table>
<%
case "1"
		%><html>
		<head>
		<title>Desactivacion anomalias</title>
		</head>
		<body>
		<table border=0 align=left cellspacing=3 width=450>
		<form name="anom_form" action="<%=asp_self()%>" method="post">
		<tr bgcolor=goldenrod>
			<th colspan=2>Verificacion :</th>
		</tr>
		<%
		call print_style()
		'set date format to US
		Session.LCID = 1033
		
		SQL = "select 1 from efolios" & _
			  " where folfolio = '" & SQLEscape(Request.Form("folio")) & "' "
		Array_anom = GetArrayRS(SQL)
		if not IsArray(Array_anom) then
			%>
			<tr><td><div class=error>El folio no existe</div></td></tr>
			<tr><td><a href=<%=asp_self()%>>regresar</a></td></tr>
			<%
			Response.End 
		else
			SQL = "select NUM_ERROR || '. ' || DESCRIPCION from REP_ANOMALIAS_CATALOGO" & _
				  " where num_error = '" & SQLEscape(Request.Form("num_error")) & "' "
			Array_anom = GetArrayRS(SQL)
			if not IsArray(Array_anom) then
				%>
				<tr><td><div class=error>El error no existe</div></td></tr>
				<tr><td><a href=<%=asp_self()%>>regresar</a></td></tr>
				<%
				Response.End 
			end if
		end if
		%>
		<tr>
			<td>Folio :</td>
			<td><%=Request.Form("folio")%></td>
		</tr>
		<tr>
			<td>Error :</td>
			<td><%=Array_anom(0,0)%></td>
		</tr>
		<tr>
			<td>Observaciones :</td>
			<td><%=HTMLEscape(Request.Form("Observaciones"))%></td>
		</tr>
		<%SQL = " select des_Num_error  " & _
				" from REP_ANOMALIAS_DESACTIVADAS, efolios " & _
				" where folclave = DES_FOLCLAVE " & _
				" AND FOLFOLIO = '" & SQLEscape(Request.Form("folio")) & "'"
		 Array_anom = GetArrayRS(SQL)
		 if IsArray(Array_anom) then
			%>
			<tr><td><br></td></tr>
			<tr>
			<td valign=top><i>Folio ya desactivado <br>en otro(s) error(es) :</i></td>
			<td valign=top><%
			for i = 0 to ubound(Array_anom, 2) 
				if i <> 0 then Response.Write ", "
				Response.Write Array_anom(0,i)
				if CInt(Array_anom(0,i)) = CInt(Request.Form("num_error")) then
					Response.Write "<br><div class=error>Este error ya existe por este folio.</div>"
					Response.Write "<br><a href=javascript:history.back();>regresar</a>"
					Response.End 
				end if
			next
			%></td>
			<%
		 end if 
		%>
		<tr>
			<td colspan=2><br><input type=submit value=validar class=buttonsorange>&nbsp;&nbsp;
			<input type=hidden value="<%=HTMLEscape(Request.Form("folio"))%>" name=folio>
			<input type=hidden value="<%=HTMLEscape(Request.Form("observaciones"))%>" name=observaciones>
			<input type=hidden value="<%=HTMLEscape(Request.Form("num_error"))%>" name=num_error>
			<input type=button value=cancelar class=buttonsorange onclick=javascript:history.back();>
			<input type=hidden value=2 name=etape></td>
		</tr>
		</form>
		<%
case "2"	
	dim rst  
	set rst = Server.CreateObject("ADODB.Recordset")
	SQL = "insert into REP_ANOMALIAS_DESACTIVADAS (DES_Folclave, DES_num_error, observaciones, date_created, created_by) " & _
		  "select folclave, '"& SQLEscape(Request.Form("num_error")) &"', '" &SQLEscape(Request.Form("observaciones")) &"'  "  & _
		  " , sysdate,'" & Session("array_user")(0,0) & "' " & _
		  " from efolios where folfolio = '"& SQLEscape(Request.Form("folio")) &"' " 
	
	'Response.Write SQL
	rst.Open SQL, Connect(), 0, 1, 1

	Response.Redirect asp_self() & "?msg=" & Server.URLEncode("Anomalia Insertada.")

case "3"
	%><html>
		<head>
		<title>Reactivacion anomalias</title>
		</head>
		<%call print_style()%>
		<body>
		<table border=0 align=left cellspacing=3>
		<form name="anom_form" action="<%=asp_self()%>" method="post">
		<tr bgcolor=goldenrod>
			<th colspan=2>Verificacion :</th>
		</tr>
		<%
		'set date format to US
		Session.LCID = 1033
		
		SQL = "select folclave from efolios, REP_ANOMALIAS_DESACTIVADAS" & _
			  " where folfolio = '" & SQLEscape(Request.Form("folio")) & "' " & _
			  " and folclave = des_folclave " & _
			  " and des_num_error = '" & SQLEscape(Request.Form("num_error")) & "' " 
		Array_anom = GetArrayRS(SQL)
		'Response.Write SQL
		if not IsArray(Array_anom) then
			%>
			<tr><td colspan=2><div class=error>Este folio no esta desactivado por este error.</div>
			<br><a href=javascript:history.back();>regresar</a></td></tr>
			<%
			Response.End 
		end if
		%>
		<tr>
			<td>Folio :</td>
			<td align=left><%=Request.Form("folio")%></td>
		</tr>
		<tr>
			<td>Error :</td>
			<td align=left><%=HTMLEscape(Request.Form("num_error"))%></td>
		</tr>

		<tr>
			<td colspan=2><br><input type=submit value=reactivar class=buttonsorange>&nbsp;&nbsp;
			<input type=hidden value="<%=Array_anom(0,0)%>" name=folio>
			<input type=hidden value="<%=HTMLEscape(Request.Form("num_error"))%>" name=num_error>
			<input type=button value=cancelar class=buttonsorange onclick=javascript:history.back();>
			<input type=hidden value=4 name=etape></td>
		</tr>
		</form>
		</table>
		<%
case "4"	
	set rst = Server.CreateObject("ADODB.Recordset")
	SQL = "delete from REP_ANOMALIAS_DESACTIVADAS  " & _
		  "where des_num_error = '"& SQLEscape(Request.Form("num_error")) &"' and des_folclave = '"& SQLEscape(Request.Form("folio")) &"' " 
	
	'Response.Write SQL
	rst.Open SQL, Connect(), 0, 1, 1

	Response.Redirect asp_self() & "?msg=" & Server.URLEncode("Anomalia Reactivada.")
end select

%>

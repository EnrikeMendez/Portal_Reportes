<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Confirmacion de los reportes
Response.Expires = 0
call check_session()

Select Case Request("Etape")
	Case ""
%>
		<html>
		<head>
		<title>Confirmacion de reportes</title>
		</head>
		<body>
		<%
		call print_style()
		Dim SQL, Array_conf, i, pend_conf, hoy, cliente, filtro, hora
		hoy = Month(now)& "/" & day(now)& "/" & Year(now)
		hora = Hour(now) & ":" 
		if Minute(now) < 10 then
			hora = hora & "0" & Minute(now)
		else 
			hora = hora  & Minute(now)
		end if
		cliente = Request("cliente")
		
		if cliente = "" then
			filtro = " where dias.cliente = 0 "
		else 
			filtro = " where dias.cliente in (0, " & cliente & ") "
		end if 
		
		SQL ="  select to_char(dias.DIA_LIBRE, 'mm/dd/yyyy') dias_libre " & _
			 " , dias.cliente " &_
			 " from rep_dias_libres dias " & _
			 SQLescape(filtro) & _
			 " order by dias.DIA_LIBRE"

		'Response.Write SQL
		array_conf= GetArrayRS(SQL)
		
		if not IsArray(Array_conf) then
			Response.Write "Ninguno dia libre selecionado."
			Response.End 
		end if
		'Response.Write "taille: " &Ubound(Array_conf,2)
		%>

		<table width=350 border=0>
		<tr>
			<td colspan=2><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<tr>
			<td colspan=2>Hoy : <%=hoy%> - <%=hora%><br><br></td>
		</tr>
		<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
		<tr bgcolor=goldenrod>
			<th colspan=2>Dias libre : <%if cliente <> "" then Response.Write "(rojo por el cliente "& cliente & ")"%></span>
			&nbsp;&nbsp;<font size=0>mm/dd/yyyy</font>
			</th>
		</tr>
		<tr> <td colspan=2>
		<form name="dias_form"  action="<%=asp_self()%>" method="post">
		<table  width=200 border=0>
		<tr bgcolor=#FFFFCC>
			<td align=center width=80>Fecha</td><td>Cliente</td><td></td>
		</tr>
			<%
			for i=0 to UBound(Array_conf, 2)
				Response.Write "<tr><td>" 
				if Array_conf(1,i) <> "0" then 
					Response.Write "<font color=red>" & Array_conf(0,i) & "</font>"
				else 
					Response.Write Array_conf(0,i) 
				end if
				Response.Write "</td><td>"& Array_conf(1,i) 
				response.Write "</td><td><a href="&asp_self()&"?etape=2&date_num="&Server.URLEncode(Array_conf(0,i))&"&cliente="&Server.URLEncode(Array_conf(1,i))&">Borrar</a></td></tr>"
			next
			%>
		</table>
		</td>
		</tr>
		<tr valign=center>
			
			<td align=left><br><input type=text name=cliente size=4></td>
			<td>Escoger un numero de cliente (opcional)
			<br>0 = todos los clientes.
			</td>
		</tr>
		<tr>
			<td colspan=2 align=left><br><input type=submit class=buttonsOrange value=Validar><br><br></td>
		</tr>
		</form>
		<tr bgcolor=goldenrod>
			<th colspan=2>Insertar otro dia :</th>
		</tr>
		<LINK media=screen href="../../v2/include/dynCalendar/dyncalendar.css" type=text/css rel=stylesheet>
		<script src="../../v2/include/dynCalendar/browserSniffer.js" type="text/javascript" language="javascript"></script>
		<script src="../../v2/include/dynCalendar/dyncalendar.js" type="text/javascript" language="javascript"></script>
		<script type="text/javascript">
			<!--
				// Calendar callback. When a date is clicked on the calendar
				// this function is called so you can do as you want with it
				function ToCalendarCallback(date, month, year)
				{
					date = month + '/' + date + '/' + year;
					document.date_form.date_to.value = date;
				}
				function FromCalendarCallback(date, month, year)
				{
					date = month + '/' + date + '/' + year;
					document.date_form.date_num.value = date;
				}
			// -->
		</script>
		<form name="date_form"  action="<%=asp_self()%>?etape=1" method="post">
		<tr><td valign=middle width=120>Date<br>
		&nbsp;&nbsp;&nbsp;<input name="date_num" size=10>
			<script language="JavaScript" type="text/javascript">
				<!--
				if (is_ie5up || is_nav6up || is_gecko){
					FromCalendar = new dynCalendar('FromCalendar', 'FromCalendarCallback');
					FromCalendar.setOffset(10, 5);
					}
				//-->
			</script>
		</td>
		</tr>
		<tr>
		<td><br>Cliente :
		<br>
		&nbsp;&nbsp;&nbsp;<input type=text name=cliente size=4></td>
		<td><br>(opcional, si no lo pones,<br> se aplicara a todos los clientes)
		</td>
		</tr>
		<tr>
			<td colspan=2 align=left><br><input type=submit class=buttonsOrange value=Validar id=submit1 name=submit1><br><br></td>
		</tr>
		</form>
		</table>
<%
case "1"
		dim rst, date_num 
		set rst = Server.CreateObject("ADODB.Recordset")
		
		cliente = Request.Form("cliente")
		date_num = Request.Form("date_num")
		if date_num = "" then Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Fecha vacia")
		if cliente = "" then cliente = 0
		SQL = "insert into rep_dias_libres values (to_date('"& SQLEscape(date_num) &"', 'mm/dd/yyyy'),"&SQLEscape(cliente)&")"
		'Response.Write SQL
		
		rst.Open SQL, Connect(), 0, 1, 1
		Response.Redirect asp_self & "?cliente=" & cliente 
		
case "2"
		set rst = Server.CreateObject("ADODB.Recordset")
		cliente = Request("cliente")
		date_num = Request("date_num")
		if date_num = "" or cliente = "" then Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Error")
		
		SQL = "delete from rep_dias_libres where dia_libre = to_date('"& SQLEscape(date_num) &"', 'mm/dd/yyyy') and cliente = "&SQLEscape(cliente)
		'Response.Write SQL
		
		rst.Open SQL, Connect(), 0, 1, 1
		'Response.Write asp_self & "?cliente=" & cliente 
		Response.Redirect asp_self & "?cliente=" & cliente 
end select


				
%>


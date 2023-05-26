<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
call check_session()
%>

<HTML>
<HEAD>
<title>
Lista de folios desactivados
</title>
</head>
<body>
<%call print_style()

Select Case Request("Etape")
	Case ""
		select case Request("nivel")
			Case ""
%>		
<LINK media=screen href="include/dyncalendar.css" type=text/css rel=stylesheet>
<script src="include/browserSniffer.js" type="text/javascript" language="javascript"></script>
<script src="include/dyncalendar.js" type="text/javascript" language="javascript"></script>
<div id="menu" style="position: absolute; left: 0; top: -15; z-index:1;">
<div id="anom" name="anom_l" style="LEFT: 50px; VISIBILITY: visible; POSITION: absolute; TOP: 50px">
	<form name="anom_form" action=<%=asp_self()%> method="post">
	<table border=0 cellspacing=3>
	<tr bgcolor=goldenrod valign=center align=center>
		<td colspan=3>Otra Busqueda</td>
	</tr>
	<tr><td><a href=lista_desact.asp?nivel=2>Por aduana</a></td></tr>
	<tr><td><a href=lista_desact.asp?nivel=1>Por numero de error</a></td></tr>
		<tr><td><hr><a href=menu.asp>menu general</a><br><br></td></tr>
	</table>
	<table width=450 cellspacing=0 cellpadding=3 border=1>
	<tr bgcolor=goldenrod valign=center align=left><td colspan=2><font color=#ffffff style="FONT-SIZE: 10pt"><b>.</font> Buscar por fecha de desactivacion :</B>     </td></tr>
	<tr>
	<script type="text/javascript">
	<!--
		// Calendar callback. When a date is clicked on the calendar
		// this function is called so you can do as you want with it
		function ToCalendarCallback(date, month, year)
		{
			date = month + '/' + date + '/' + year;
			document.anom_form.anom_to.value = date;
		}
		function FromCalendarCallback(date, month, year)
		{
			date = month + '/' + date + '/' + year;
			document.anom_form.anom_num.value = date;
		}
	// -->
	</script>	
		
		
		<td valign=center>
		From<br>
		&nbsp;&nbsp;&nbsp;<input size=15 name="anom_num" readonly>
			<script language="JavaScript" type="text/javascript">
				<!--
				if (is_ie5up || is_nav6up || is_gecko){
					FromCalendar = new dynCalendar('FromCalendar', 'FromCalendarCallback');
					FromCalendar.setOffset(-45, -40);
					}
				//-->
			</script>			
		<br>To<br>
		&nbsp;&nbsp;&nbsp;<input size=15 name="anom_to" readonly>
			<script language="JavaScript" type="text/javascript">
				<!--
				if (is_ie5up || is_nav6up || is_gecko){
					ToCalendar = new dynCalendar('ToCalendar', 'ToCalendarCallback');
					ToCalendar.setOffset(-45, -40);
					}
				//-->
			</script>
		</td>
	</tr>
	</table>
		<input type=hidden name=Etape value=1>
		<input TYPE=submit  VALUE="Search" class="buttonsorange" onclick="javascript:check_data();" id=submit1 name=submit1>
	</form>
	<script language="javascript">
	function check_data()
	{if (document.anom_form.anom_num.value != "")
		{if (document.anom_form.anom_to.value != "")
			return true}
	 
	 alert ('Please select 2 date.');
	 return false
	
	}
	</script> 
</div>
</div>

<%	
	case "1"
%>

<div id="menu" style="position: absolute; left: 0; top: -15; z-index:1;">
<div id="anom_l" name="anom_l" style="LEFT: 50px; VISIBILITY: visible; POSITION: absolute; TOP: 50px">
	<form name="anom_form" action=<%=asp_self()%> method="post">
	<table border=0 cellspacing=3>
	<tr bgcolor=goldenrod valign=center align=center>
		<td colspan=3>Otra Busqueda</td>
	</tr>
	<tr><td><a href=lista_desact.asp>Por fecha</a></td></tr>
	<tr><td><a href=lista_desact.asp?nivel=2>Por aduana</a></td></tr>
		<tr><td><hr><a href=menu.asp>menu general</a><br><br></td></tr>
	</table>
	<table width=450 cellspacing=0 cellpadding=3 border=1>
	<tr bgcolor=goldenrod valign=center align=left><td colspan=2><font color=#ffffff style="FONT-SIZE: 10pt"><b>.</font> Buscar por tipo de error :</B>     </td></tr>
	<tr>
		<td valign=center>
		<%dim i, SQL, array_sql
		SQL = "select distinct des_num_error from rep_anomalias_desactivadas " & _
				" order by 1 "
		array_sql = GetArrayRS(SQL)
		if IsArray(array_sql) then
			Response.Write "<select name=error>"
			for i = 0 to Ubound(array_sql,2)
				Response.Write "<option value=" & array_sql(0,i) & ">"& array_sql(0,i)
			next	
			Response.Write "</select>"
			Response.Write "<br><br><input TYPE=submit  VALUE=""Search"" class=""buttonsorange"">"
		end if
		%>
			
		</td>
	</tr>
	<input type=hidden value=1 name=etape>
	</table>
		
	</form>

</div>
</div>
<%	
	case "2"
%>

<div id="menu" style="position: absolute; left: 0; top: -15; z-index:1;">
<div id="anom_l" name="anom_l" style="LEFT: 50px; VISIBILITY: visible; POSITION: absolute; TOP: 50px">
	<form name="anom_form" action=<%=asp_self()%> method="post">
	<table border=0 cellspacing=3>
	<tr bgcolor=goldenrod valign=center align=center>
		<td colspan=3>Otra Busqueda</td>
	</tr>
	<tr><td><a href=lista_desact.asp>Por fecha</a></td></tr>
	<tr><td><a href=lista_desact.asp?nivel=1>Por Numero de error</a></td></tr>
		<tr><td><hr><a href=menu.asp>menu general</a><br><br></td></tr>
	</table>
	<table width=450 cellspacing=0 cellpadding=3 border=1>
	<tr bgcolor=goldenrod valign=center align=left><td colspan=2><font color=#ffffff style="FONT-SIZE: 10pt"><b>.</font> Buscar por aduana :</B>     </td></tr>
	<tr>
		<td valign=center>
		<%
		SQL = "select distinct fol_douclef from efolios " & _
				" , rep_anomalias_desactivadas " & _
				" where des_folclave=folclave " & _
				" order by 1 "
		array_sql = GetArrayRS(SQL)
		if IsArray(array_sql) then
			Response.Write "<select name=ad>"
			for i = 0 to Ubound(array_sql,2)
				Response.Write "<option value=" & array_sql(0,i) & ">"& array_sql(0,i)
			next	
			Response.Write "</select>"
			Response.Write "<br><br><input TYPE=submit  VALUE=""Search"" class=""buttonsorange"">"
		end if
		%>
			
		</td>
	</tr>
	<input type=hidden value=1 name=etape>
	</table>
		
	</form>

</div>
</div>
		<%end select%>

<%
	case "1"	


	dim filtro
	if Request.Form("error") <> "" then
		filtro = "and des_num_error = '"& Request.Form("error")&"'"
	end if
	
	if Request.Form("ad") <> "" then
		filtro = "and fol_douclef = '"& Request.Form("ad") &"'"
	end if
	
	if Request.Form("anom_to") <> "" and Request.Form("anom_num") <> ""  then
		filtro = "and D.Date_created >= to_date('"&Request.Form("anom_num")&"', 'mm/dd/yyyy') " & _
				 "and D.Date_created < to_date('"&Request.Form("anom_to")&"', 'mm/dd/yyyy') +1" 
	end if
'test si la 1ere qry a ete executee si non on la lance
	SQL = "select des_num_error,folfolio, fol_douclef, " & _
		  " to_char(D.Date_created, 'mm/dd/yyyy'), observaciones " & _
		  " from efolios, rep_anomalias_desactivadas D " & _
		  " where des_folclave=folclave " & _
		  filtro & _
		  " order by D.Date_created desc "
	
	Dim tab_fin
	'Response.Write SQL
	tab_fin = GetArrayRS(SQL)



'Response.End 
'initialisation des num de page
Dim PageSize, PageNum
PageSize = 15
PageNum = Request("Page_Num")
if Not IsNumeric(PageNum) or Len(PageNum) = 0 then
   PageNum = 1
else
   PageNum = CInt(PageNum)
end if

if not IsArray (tab_fin) then
	response.write "No records found !"
	response.end
end if

Dim iRows, iCols, iRowLoop, iColLoop, iStop, iStart
Dim iRows2, iCols2
 iRows = UBound(tab_fin , 2)
 iCols = UBound(tab_fin , 1) 



If iRows > (PageNum * PageSize ) Then
   iStop = PageNum * PageSize - 1
Else
   iStop = iRows
End If 

iStart = (PageNum -1 )* PageSize
If iStart > iRows then iStart = iStop - PageSize  'inutile en principe... mais bon si on modifie la variable pagenum...
 

%>
<table border=0 cellspacing=3>
<tr bgcolor=goldenrod valign=center align=center>
	<td colspan=3>Otra Busqueda</td>
</tr>
<tr><td><a href=lista_desact.asp>Por fecha</a></td></tr>
<tr><td><a href=lista_desact.asp?nivel=2>Por Aduana</a></td></tr>
<tr><td><a href=lista_desact.asp?nivel=1>Por Numero de error</a></td></tr>
<tr><td><hr><a href=menu.asp>menu general</a><br><br></td></tr>
</table>
<table BORDER="1" width=700 cellpadding=2 cellspacing=0>
 <tr bgcolor=goldenrod valign=center align=center>
        
        <td>Num Error</td>
        <td>Folio</td>
        <td>Aduana</td>
        <td>Fecha</td>
        <td>Oberservaciones</td>
 </tr>
 
 <%
 For iRowLoop = iStart to iStop 
 	%>
 	
 	<tr align=center valign=center>	
	  <td><%=tab_fin(0,iRowLoop)%></td>
	  <td><%=tab_fin(1,iRowLoop)%></td>
	  <td><%=tab_fin(2,iRowLoop)%></td>
	  <td><%=tab_fin(3,iRowLoop)%></td>
	  <td align=left><%=tab_fin(4,iRowLoop)%></td>
	</tr>
<%
  next
  
%>
<form name="next_page" action="<%call asp_self()%>" method="post">
	<input type="hidden" name="ad" value="<%=Request.Form("ad")%>">
	<input type="hidden" name="error" value="<%=Request.Form("error")%>">
	<input type="hidden" name="anom_to" value="<%=Request.Form("anom_to")%>">
	<input type="hidden" name="anom_num" value="<%=Request.Form("anom_num")%>">
	<input type="hidden" name="page_num" value="<%=TAGEscape(Request.Form("page_num"))%>">
	<input type="hidden" name="etape" value="1">
</form>

<script language="javascript">
function next_page(page_num)
{document.next_page.page_num.value = page_num;
document.next_page.submit();
}
</script>
</table>
</div> 	

<%
'NB : iRows contient le dernier indice du tableau donc nb_lignes -1 !
call BuildNav2(PageNum, PageSize, iRows +1,"next_page")
%>
<%end select%>

</body>
</html>
<%
'html functions :

sub print_style()
%>
<style type="text/css">
td{
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	font-style: normal;
	font-weight: bold;
	font-size: 10px;
}

th {
	background-color: goldenrod;
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	font-style: normal;
	font-weight: bold;
	font-size: 12px;
}

.title {
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	font-style: normal;
	font-weight: bold;
	font-size: 11px;
}

.light {
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	font-style: normal;
	font-size: 9px;
}

a:visited {
	color: Blue;
}

.buttonsOrange { background-color: goldenrod; color:#000000; font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; height: 19px; font-color: #000000 }

.error {
	font-size: 12pt;
	background-color: #ffffff;
	font-family: Arial, sans-serif;
	font-weight: bold;
	color: #ff0000;
}
</style>
<%
end sub

'Source : http://jenda.krynicky.cz/Texts/Escape.html
' !!! NEVER EVER INSERT A STRING INTO A HTML, JAVASCRIPT OR EVEN SQL CODE WITHOUT ESCAPING IT PROPERLY!!!

' HTMLescape : if the variable is to be showed on the page
'	<td>< %=HTMLescape( variable )% ></td>

' TAGescape : if its used to set the value of a field
'	<input type="text" name="Foo" value="< %=TAGescape( variable )% >">

' JSescape : if it's used as a parameter to a JavaScript function
'	<a href="JavaScript:foo( '< %=JSescape(variable)% >' );">

' URLescape : if it's to become part of a query string
'	<a href="Foo.asp?name=< %=URLescape(variable)% >"

' SQLescape : if it's to become part of an SQL query
'	sql = "select * from Table where name = '" & SQLescape( name) & "'"

' SQLquote : if it's to become part of an SQL query, adds the quotes aroung the value, works correct with NULLs
'	sql = "select * from Table where name = " & SQLquote( name)

Function HTMLescape (s)
	if isEmpty(s) or isNull(s) then
		HTMLescape = ""
	else
		HTMLescape = Server.HTMLEncode(s)
	end if
End Function

Function TAGescape (s)
	if isEmpty(s) or isNull(s) then
		TAGescape = ""
	else
		TAGescape = Replace(Replace(Server.HTMLEncode(s),"""", "&dblquote;" ), "'", "&#39;")
	end if
End Function

Function URLescape (s)
	if isEmpty(s) or isNull(s) then
		URLescape = ""
	else
		URLescape = Replace(Replace(Server.URLEncode(s),"'","%27"), """", "%22" )
	end if
End Function

Function JSescape (s)
	if isEmpty(s) or isNull(s) then
		JSescape = ""
	else
		JSescape = TAGescape( Replace( Replace( Replace( s, "\", "\\"), """", "\""" ) , "'", "\'" ))
	end if
End Function

Function SQLescape (s)
	if isEmpty(s) or isNull(s) then
		SQLescape = ""
	else
		SQLescape = Replace( s, "'", "''")
	end if
End Function

Function SQLquote (s)
	if isEmpty(s) or isNull(s) then
		SQLquote = "NULL"
	else
		SQLquote = "'" & Replace( s, "'", "''") & "'"
	end if
End Function
%>

<%sub print_popup()%>
<!-- Calque de la bulle d'aide -->
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<script language="JavaScript" src="include/overlib.js">
//<!-- overLIB (c) Erik Bosrup -->
</script> 

<%end sub%>


<%
Function display_mail(servidor , warning_message ) 
'crea el cuerpo del correo
'- servidor : permite de poner las direccion de rede interna o foranea :
'             192.168.100.10 o www.logisconcept.com
'- warning_message : desplega el contenido del aviso si existe

display_mail = "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf
display_mail = display_mail & "<html>" & vbCrLf
display_mail = display_mail & "<head>" & vbCrLf
display_mail = display_mail & "<meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"">" & vbCrLf
display_mail = display_mail & "<title>Logis Web Site</title>" & vbCrLf
display_mail = display_mail & "<STYLE TYPE=""text/css"">" & vbCrLf
display_mail = display_mail & ".link:active {color:""#0000FF""}" & vbCrLf
display_mail = display_mail & ".link:link {color:""#0000FF""}" & vbCrLf
display_mail = display_mail & ".link:hover {color:""#0000FF""}" & vbCrLf
display_mail = display_mail & ".link:visited {color:""#0000FF""}" & vbCrLf
display_mail = display_mail & "</STYLE>" & vbCrLf
display_mail = display_mail & "</head>" & vbCrLf
display_mail = display_mail & "<body>" & vbCrLf
display_mail = display_mail & "" & vbCrLf
display_mail = display_mail & "<center>" & vbCrLf
display_mail = display_mail & "<TABLE WIDTH=""500"" CELLSPACING=""0"" CELLPADDING=""0"" BORDER=""1"" bgcolor=""#ffffff"">" & vbCrLf
display_mail = display_mail & "<tr><td>" & vbCrLf
display_mail = display_mail & "<TABLE WIDTH=""500"" CELLSPACING=""0"" CELLPADDING=""0"" BORDER=""0"" bgcolor=""#336699"">" & vbCrLf
display_mail = display_mail & "<tr>" & vbCrLf
display_mail = display_mail & "<td align=""left""><a href=""" & servidor & "/""><IMG SRC=""" & servidor & "/images/pixel.gif"" WIDTH=""1"" HEIGHT=""45"" border=""0"" alt=""""><img src=""" & servidor & "/images/LetrasLogis.gif"" border=""0"" alt=""logo logis""></a></td>" & vbCrLf
display_mail = display_mail & "</tr>" & vbCrLf
display_mail = display_mail & "</table>" & vbCrLf
display_mail = display_mail & "<TABLE WIDTH=""500"" CELLSPACING=""0"" CELLPADDING=""0"" BORDER=""0"" bgcolor=""#ffffff"">" & vbCrLf
display_mail = display_mail & "<tr>" & vbCrLf
display_mail = display_mail & "    <td align=""left""><IMG SRC=""" & servidor & "/images/pixel.gif"" WIDTH=""1"" HEIGHT=""5"" alt=""""></td>" & vbCrLf
display_mail = display_mail & "</tr>" & vbCrLf
display_mail = display_mail & "</table>" & vbCrLf
display_mail = display_mail & "<TABLE WIDTH=""500"" CELLSPACING=""0"" CELLPADDING=""3"" BORDER=""0"" bgcolor=""#ffffff"">" & vbCrLf
display_mail = display_mail & "<tr bgcolor=""#C69633"">" & vbCrLf
display_mail = display_mail & "    <td height=""25""  align=""left"" valign=bottom><IMG SRC=""" & servidor & "/images/pixel.gif"" WIDTH=""20"" HEIGHT=""1"" alt=""""><FONT FACE=""Arial,Helvetica"" SIZE=""3"" COLOR=""#FFFFFF""><B>Logis Web Site :</B></FONT></td>" & vbCrLf
display_mail = display_mail & "</tr>" & vbCrLf
display_mail = display_mail & "<tr>" & vbCrLf
display_mail = display_mail & "    <td><br>" & vbCrLf
display_mail = display_mail & warning_message & vbCrLf
display_mail = display_mail & "    </td>" & vbCrLf
display_mail = display_mail & "</tr>" & vbCrLf
display_mail = display_mail & "<tr>" & vbCrLf
display_mail = display_mail & "    <td><IMG SRC=""" & servidor & "/images/pixel.gif"" WIDTH=""1"" HEIGHT=""30"" alt=""""></td>" & vbCrLf
display_mail = display_mail & "</tr>" & vbCrLf
display_mail = display_mail & "<tr>" & vbCrLf
display_mail = display_mail & "    <td align=""left""><hr>" & vbCrLf
display_mail = display_mail & "    <table border=""0""><tr><td><FONT SIZE=""2"" FACE=""Arial,Helvetica"" COLOR=""#000000"">This is a message automatically generated, please contact" & vbCrLf
display_mail = display_mail & "<a href=""mailto:" & Get_mail("webmaster") & """ class=""link"">" & Get_mail("webmaster") & "</a> for any question or to unsubscribe. </FONT></td>" & vbCrLf
display_mail = display_mail & "    <td align=""right"">" & vbCrLf
display_mail = display_mail & "        <p><img border=""0"" src=""http://www.w3.org/Icons/valid-html401""  alt=""Valid HTML 4.01!"" height=""31"" width=""88""></p>" & vbCrLf
display_mail = display_mail & "    </td></tr>" & vbCrLf
display_mail = display_mail & "	   </table>" & vbCrLf
display_mail = display_mail & "</tr>" & vbCrLf
display_mail = display_mail & "<tr bgcolor=""#336699"">" & vbCrLf
display_mail = display_mail & "    <td><IMG SRC=""" & servidor & "/images/pixel.gif"" WIDTH=""1"" HEIGHT=""15"" alt=""""></td>" & vbCrLf
display_mail = display_mail & "</tr>" & vbCrLf
display_mail = display_mail & "</table>" & vbCrLf
display_mail = display_mail & "" & vbCrLf
display_mail = display_mail & "</td></tr>" & vbCrLf
display_mail = display_mail & "</table>" & vbCrLf
display_mail = display_mail & "</center>" & vbCrLf
display_mail = display_mail & "</BODY>" & vbCrLf
display_mail = display_mail & "</HTML>"

end function
%>
<%sub print_bodega(form_name, id_param)
	dim i, i2, SQL,array_bodega
	SQL = "SELECT EAL.ALLCLAVE, EAL.ALLCODIGO || ' - ' || InitCap(EAL.ALLNOMBRE) FROM EALMACENES_LOGIS EAL WHERE ALLCLAVE > 0 ORDER BY EAL.ALLNOMBRE"
	array_bodega = GetArrayRS(SQL)
	i2=0
	Response.Write "<table cellspacing=0 cellpadding=2 border=0 width=400>" & vbCrLf
	for i = 0 to Ubound(array_bodega, 2)
		if (i2  mod 2) = 0 then Response.Write "<tr>" & vbCrLf & vbtab
			Response.Write "<td width=20><input type=checkbox name=" & id_param & " value="& array_bodega(0,i) &"></td>" & vbCrLf
			Response.Write "<td>"& array_bodega(1,i) &"</td>" & vbCrLf & vbtab
			i2=i2+1
		if (i2  mod 4) = 0 then Response.Write "</tr>" & vbCrLf 
	next
	%><script language = "Javascript">
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
	// -->

	</script>
	<tr><td colspan="2">
		<a href="javascript:SetChecked(1,'<%=id_param%>','<%=form_name%>')"><font face="Arial, Helvetica, sans-serif" size="0">Selecionar todo</font></a>
		&nbsp;&nbsp;<a href="javascript:SetChecked(0,'<%=id_param%>','<%=form_name%>')"><font face="Arial, Helvetica, sans-serif" size="0">Quitar todo</font></a>
	 </td></tr>
	<%
	Response.Write "</table>" & vbCrLf
end sub%>
<%sub print_sucursal(form_name, id_param)
	dim i, i2, SQL,array_bodega
	SQL = "SELECT '''' || SUCCLEF || '''', SUCCLEF || ' - ' || INITCAP(SUCNOM)  " & VbCrlf
    SQL = SQL & " FROM ESUCCURSALE " & VbCrlf
    SQL = SQL & " WHERE EXISTS ( " & VbCrlf
    SQL = SQL & " 	SELECT NULL " & VbCrlf
    SQL = SQL & " 	FROM ECLIENT " & VbCrlf
    SQL = SQL & " 	WHERE CLI_SUCURSAL = SUCCLEF " & VbCrlf
    SQL = SQL & " 	AND CLISTATUS = 0) " & VbCrlf
    SQL = SQL & " ORDER BY 2 "
	array_bodega = GetArrayRS(SQL)
	i2=0
	Response.Write "<table cellspacing=0 cellpadding=2 border=0 width=400>" & vbCrLf
	for i = 0 to Ubound(array_bodega, 2)
		if (i2  mod 2) = 0 then Response.Write "<tr>" & vbCrLf & vbtab
			Response.Write "<td width=20><input type=checkbox name=" & id_param & " value="""& array_bodega(0,i) &"""></td>" & vbCrLf
			Response.Write "<td>"& array_bodega(1,i) &"</td>" & vbCrLf & vbtab
			i2=i2+1
		if (i2  mod 4) = 0 then Response.Write "</tr>" & vbCrLf 
	next
	%><script language = "Javascript">
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
	// -->

	</script>
	<tr><td colspan="2">
		<a href="javascript:SetChecked(1,'<%=id_param%>','<%=form_name%>')"><font face="Arial, Helvetica, sans-serif" size="0">Selecionar todo</font></a>
		&nbsp;&nbsp;<a href="javascript:SetChecked(0,'<%=id_param%>','<%=form_name%>')"><font face="Arial, Helvetica, sans-serif" size="0">Quitar todo</font></a>
	 </td></tr>
	<%
	Response.Write "</table>" & vbCrLf
end sub%>
<%sub print_estado(form_name, id_param)
	dim i, i2, SQL,array_bodega
	SQL = "SELECT EST.ESTESTADO, InitCap(EST.ESTNOMBRE) FROM EESTADOS EST WHERE EST_PAYCLEF = 'N3' ORDER BY EST.ESTNOMBRE"
	array_bodega = GetArrayRS(SQL)
	i2=0
	Response.Write "<table cellspacing=0 cellpadding=2 border=0 width=400>" & vbCrLf
	for i = 0 to Ubound(array_bodega, 2)
		if (i2  mod 2) = 0 then Response.Write "<tr>" & vbCrLf & vbtab
			Response.Write "<td width=20><input type=checkbox name=" & id_param & " value="""& array_bodega(0,i) &"""></td>" & vbCrLf
			Response.Write "<td>"& array_bodega(1,i) &"</td>" & vbCrLf & vbtab
			i2=i2+1
		if (i2  mod 4) = 0 then Response.Write "</tr>" & vbCrLf 
	next
	%><script language = "Javascript">
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
	// -->

	</script>
	<tr><td colspan="2">
		<a href="javascript:SetChecked(1,'<%=id_param%>','<%=form_name%>')"><font face="Arial, Helvetica, sans-serif" size="0">Selecionar todo</font></a>
		&nbsp;&nbsp;<a href="javascript:SetChecked(0,'<%=id_param%>','<%=form_name%>')"><font face="Arial, Helvetica, sans-serif" size="0">Quitar todo</font></a>
	 </td></tr>
	<%
	Response.Write "</table>" & vbCrLf
end sub%>
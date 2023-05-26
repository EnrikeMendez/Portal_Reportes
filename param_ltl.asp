<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
Response.Expires = 0
call check_session()


Select Case Request.Form("Etape")
	Case ""
		dim msg, SQL, arrayRS, i, rst
		%>
		<html>
		<head>
		<title>Parametros LTL</title>
		</head>
		<body>
		<%call print_style()%>
	
        <form name="form_1"  action="<%=asp_self()%>" method="post">
		<table width=680 BORDER="1" cellpadding="2" cellspacing="0">
		<tr>
			<td colspan=7><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
		<tr bgcolor=goldenrod align="center">
			<th width="50">N° Cliente</th>
			<th>Razon Social</th>
			<th width="85">Dimensiones<br>por caja</th>
			<th width="85">Peso<br>por caja</th>
			<th width="85">Dimensionar<br>Encabezados</th>
			<th width="85">Dimensionar<br>Referencias</th>
			<th width="85">Captura de dimensiones web<br>en documentacion por encabezados</th>
			<th width="10">&nbsp;</th>
		</tr>
		</table>
		<div style="width:680px; height: 400px; overflow : auto;">
		<table width=100% BORDER="1" cellpadding="2" cellspacing="0">
		<%SQL = "SELECT /*+use_nl(wcp cli)*/  DISTINCT CLICLEF  " & VbCrlf
        SQL = SQL & "    , InitCap(CLINOM)  " & VbCrlf
        SQL = SQL & "    , DECODE(WCPCAJA_DIMENSIONES, 'S', 'checked', NULL)  " & VbCrlf
        SQL = SQL & "    , DECODE(WCPCAJA_PESO, 'S', 'checked', NULL)  " & VbCrlf
        SQL = SQL & "    , DECODE(WCPENCABEZADO, 'S', 'checked', NULL)  " & VbCrlf
        SQL = SQL & "    , DECODE(WCP_REFERENCIA, 'S', 'checked', NULL)  " & VbCrlf
        SQL = SQL & "    , DECODE(WCPCAPTURA_BULTOS_WEB, 'S', 'checked', NULL)  " & VbCrlf
        
        SQL = SQL & "  FROM ECLIENT CLI " & VbCrlf
        SQL = SQL & "    , WEB_CAPTURA_PARAMETROS WCP " & VbCrlf
        SQL = SQL & "  WHERE WCP_CLICLEF(+) = CLICLEF  " & VbCrlf
        SQL = SQL & "    AND EXISTS (  " & VbCrlf
        SQL = SQL & "    	  SELECT NULL  " & VbCrlf
        SQL = SQL & "  	  FROM WEB_LTL WEL  " & VbCrlf
        SQL = SQL & "  	  WHERE WEL_CLICLEF = CLICLEF  " & VbCrlf
        SQL = SQL & "  	    AND WEL.DATE_CREATED > SYSDATE - 30  " & VbCrlf
        SQL = SQL & "  		AND rownum = 1 " & VbCrlf
        SQL = SQL & " 	  UNION ALL " & VbCrlf
        SQL = SQL & " 	  SELECT /*+INDEX(WCD IDX_WCD_CLICLEF)*/ NULL " & VbCrlf
        SQL = SQL & " 	  FROM WCROSS_DOCK WCD " & VbCrlf
        SQL = SQL & " 	  WHERE WCD_CLICLEF = CLICLEF  " & VbCrlf
        SQL = SQL & " 	    AND WCD.DATE_CREATED > SYSDATE - 30  " & VbCrlf
        SQL = SQL & "  		AND rownum = 1 )  " & VbCrlf
        SQL = SQL & " ORDER BY 1 " 

		'Response.Write SQL
		arrayRS = GetArrayRS(SQL)
		if not IsArray(arrayRS) then
			Response.Write "No hay clientes."
			Response.End 
		end if
		
		for i = 0 to UBound(arrayRS, 2)%>
		  <tr>
		    <td width="50"><%=arrayRS(0,i)%></td>
		    <td><%=arrayRS(1,i)%></td>
		    <td Align=center width="85"><input type="checkbox" name="dim_<%=arrayRS(0,i)%>" value="S" <%=arrayRS(2,i)%>></td>
		    <td Align=center width="85"><input type="checkbox" name="peso_<%=arrayRS(0,i)%>" value="S" <%=arrayRS(3,i)%>></td>
		    <td Align=middle width="85"><input type="checkbox" name="enc_<%=arrayRS(0,i)%>" value="S" <%=arrayRS(4,i)%>></td>
		    <td Align=middle width="85"><input type="checkbox" name="ref_<%=arrayRS(0,i)%>" value="S" <%=arrayRS(5,i)%>></td>
		    <td Align=middle width="85"><input type="checkbox" name="web_<%=arrayRS(0,i)%>" value="S" <%=arrayRS(6,i)%>></td>
		  </tr>
		<%next%>

		</table>
		</div>
		<table width=680 BORDER="1" cellpadding="2" cellspacing="0">
		  <tr>
			<td align=left colspan=7>
			<input type="hidden" name=etape value=1>
			<input type=submit class=buttonsOrange value=validar id=submit1 name=submit1><br><br>
			</td>
		</tr>
		</form>

		
<%case "1"
	dim clef
	
	'SQL = "DELETE WEB_CAPTURA_PARAMETROS"
	'set rst = Server.CreateObject("ADODB.Recordset")
    'rst.Open SQL, Connect(), 0, 1, 1
    
	for each clef in Request.Form
	    'Response.Write clef & " - " & Request.Form(clef) & "<br>"
	    if left(clef, 3) = "dim" or left(clef, 3) = "pes" or left(clef, 3) = "enc" or left(clef, 3) = "ref"  or left(clef, 3) = "web" then
	        SQL = "SELECT COUNT(0) FROM WEB_CAPTURA_PARAMETROS WHERE WCP_CLICLEF = " & split(clef, "_")(1)
	        arrayRS = GetArrayRS(SQL)
	        if arrayRS(0,0) = "0" then
	            SQL = "INSERT INTO WEB_CAPTURA_PARAMETROS (WCP_CLICLEF " & VbCrlf
                if left(clef, 3) = "dim" then
                    SQL = SQL & "    , WCPCAJA_DIMENSIONES " & vbCrLf
                elseif left(clef, 3) = "enc" then
                    SQL = SQL & "    , WCPENCABEZADO " & vbCrLf
                elseif left(clef, 3) = "ref" then
                    SQL = SQL & "    , WCP_REFERENCIA " & vbCrLf
                elseif left(clef, 3) = "web" then
                    SQL = SQL & "    , WCPCAPTURA_BULTOS_WEB " & vbCrLf
                else
                    SQL = SQL & "    , WCPCAJA_PESO " & vbCrLf
                end if
                SQL = SQL & "    , CREATED_BY, DATE_CREATED)  " & VbCrlf
                SQL = SQL & " VALUES (" & split(clef, "_")(1) & VbCrlf
                SQL = SQL & "   , 'S' " & VbCrlf
                SQL = SQL & "  , USER, SYSDATE)  "
	        else
	            SQL = " UPDATE WEB_CAPTURA_PARAMETROS " & vbCrLf
	            SQL = SQL & " SET  "
	            if left(clef, 3) = "dim" then
                    SQL = SQL & " WCPCAJA_DIMENSIONES = 'S' " & vbCrLf
                elseif left(clef, 3) = "enc" then
                    SQL = SQL & " WCPENCABEZADO = 'S' " & vbCrLf
                elseif left(clef, 3) = "ref" then
                    SQL = SQL & " WCP_REFERENCIA = 'S' " & vbCrLf
                elseif left(clef, 3) = "web" then
                    SQL = SQL & " WCPCAPTURA_BULTOS_WEB = 'S' " & vbCrLf
                else
					SQL = SQL & " WCPCAJA_PESO = 'S' " & vbCrLf
                end if
                SQL = SQL & " , MODIFIED_BY = USER " & vbCrLf
                SQL = SQL & " , DATE_MODIFIED = SYSDATE " & vbCrLf
                SQL = SQL & " WHERE WCP_CLICLEF = " & split(clef, "_")(1)
                
	        end if
	        'Response.Write SQL & "<br>" & "<br>"
	        set rst = Server.CreateObject("ADODB.Recordset")
            rst.Open SQL, Connect(), 0, 1, 1
	    end if
	next
	
	Response.Redirect "menu.asp?msg=" & Server.URLEncode ("Parametros de LTL actualizados.")

end select

%>

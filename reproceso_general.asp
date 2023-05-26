<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'modificacion de reportes
Response.Expires = 0
'call check_session()
dim SQL, arrayRS, SQL_02, arrayRS2, i, rst
set rst = Server.CreateObject("ADODB.Recordset")

Function NVL(str)
	if IsNull(str) then
		NVL = "" 
	else 
		NVL = str
	end if
End Function

Select Case Request.Form("Etape")
	Case ""
	
%>
		<html>
		<head>
		<title>Reproceaso general</title>
		
		</head>
		<body>
			<div style="border: 1px solid #C00; padding-left: 20px;">
				<strong style="color:#C00;"><p>Antes de reprocesar, es necesario verificar que:</p>
				<p>El reporte "Pedimento Instantaneo" exista para el cliente solicitado.</p>
				<p>Se debe de agrear como destinatarios a desarrollo_web@logis.com.mx para asegurarse que haya llegado el correo</p></strong>
			</div>
			<br>
			<br>
			<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><h2 color=red size=2>" & Request("msg") & "</h2></td></tr>"
					%>
			<form action="reproceso_general.asp" method="post">
				<label>Pedimiento</label>
				<input type="text" name="pedimento" value="">
				<label>Aduana</label>
				<input type="text" name="aduana" value="">
				<label>A&ntilde;o</label>
				<input type="text" name="an_o" value="">
				<button type="submit"> Reprocesar</button>
				
				<input type="hidden" name="etape" value="1">
				<input type="hidden" name="id_reporte" value="">
				<input type="hidden" name="accion" value="">	
			</form>
<%
case "1"
		


		SQL = "select trunc(PEDDATE) PEDDATE from  epedimento " 
		SQL = SQL &" WHERE pednumero = '" & SQLEscape(Request.Form("pedimento")) & "'  " & vbCrLf 
		SQL = SQL &" AND peddouane = " & SQLEscape(Request.Form("aduana")) & " " & vbCrLf 
		SQL = SQL &" and pedanio = " & SQLEscape(Request.Form("an_o")) & " "
		arrayRS = GetArrayRS(SQL)
		'Response.Write SQL
		if not IsArray(arrayRS) then 
			Response.Redirect "reproceso_general.asp" & "?msg=" & Server.URLEncode ("No se encontraron registros en la base de datos.")
		end if
		
		SQL = " update epedimento  " 
		SQL = SQL &" set PEDDATE = null " & vbCrLf 
		SQL = SQL &" WHERE pednumero = '" & SQLEscape(Request.Form("pedimento")) & "'  " & vbCrLf 
		SQL = SQL &" AND peddouane = " & SQLEscape(Request.Form("aduana")) & " " & vbCrLf 
		SQL = SQL &" and pedanio = " & SQLEscape(Request.Form("an_o")) & " "
		'Response.Write SQL
		rst.Open SQL, Connect(), 0, 1, 1
		
		SQL = " update epedimento  " 
		SQL = SQL &" set PEDDATE = to_date('" & SQLEscape(arrayRS(0,0)) & "', 'mm/dd/yyyy') " & vbCrLf 
		'SQL = SQL &" set PEDDATE = to_date('04/03/2021', 'mm/dd/yyyy') " & vbCrLf 
		SQL = SQL &" WHERE pednumero = '" & SQLEscape(Request.Form("pedimento")) & "'  " & vbCrLf 
		SQL = SQL &" AND peddouane = " & SQLEscape(Request.Form("aduana")) & " " & vbCrLf 
		SQL = SQL &" and pedanio = " & SQLEscape(Request.Form("an_o")) & " "
		'Response.Write SQL
		rst.Open SQL, Connect(), 0, 1, 1
		
		
		
		Response.Redirect "reproceso_general.asp" & "?msg=" & Server.URLEncode ("Se reproceso.")

end select %>
</body>
</html>

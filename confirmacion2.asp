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
		<title>Confirmacion de aduanas</title>
		</head>
		<body>
		<LINK media=screen href="./include/dyncalendar.css" type=text/css rel=stylesheet>
		<script src="./include/browserSniffer.js" type="text/javascript" language="javascript"></script>
		<script src="./include/dyncalendar.js" type="text/javascript" language="javascript"></script>

		<%
		call print_style()
		Dim SQL, aduanas_usuario, array_conf, i
		Dim hora, hoy, dia_conf
		hora = Hour(Now)
		if Minute(Now) < 10 then
			hora = hora & ":0" & Minute(Now)
		else 
			hora = hora & ":" & Minute(Now)
		end if
		hoy = Month(now)& "/" & day(now)& "/" & Year(now)


		'no sirve, a lo menos el usuario tiene acceso con aduana : 0 (sin aduana) pero no le deja el menu.asp
		if Trim(Session("array_user")(1,0)) <> "" then
		'el usuario puede validar a lo menos una aduana...
			aduanas_usuario = "'" & Trim(Session("array_user")(1,0)) & "'"
			for i = 1 to UBound(Session("array_user"),2)
				aduanas_usuario = aduanas_usuario & ", '" & Trim(Session("array_user")(1,i)) & "'"
			next
		end if
	
	SQL = " select to_char(decode(sign(to_number(to_char(sysdate, 'HH24')) - 7), -1, sysdate-1, sysdate), 'mm/dd/yyyy') " & _
		  " from dual"
	array_conf= GetArrayRS(SQL)
	dia_conf = array_conf(0,0)
	
	SQL = "select RCA.RCA_DOUCLEF || ' - ' || Initcap(dou.douabreviacion) || ' - ' || to_char(RCA.RCA_FECHA_CONF, 'mm/dd/yyyy') " & _
		  " , RCA.RCA_DOUCLEF, to_char(RCA.RCA_FECHA_CONF, 'mm/dd/yyyy') " & _
		  " from REP_CONFIRMACION_ADUANAS RCA " & _
		  " , EDOUANE DOU " & _
		  " where RCA.RCA_DOUCLEF in ("& aduanas_usuario &") " & _
		  " and trunc(RCA.RCA_FECHA_CONF) = trunc(decode(sign(to_number(to_char(sysdate, 'HH24')) - 7), -1, sysdate-1, sysdate)) " & _
		  " and dou.douclef = RCA.RCA_DOUCLEF " & _
		  " order by dou.douabreviacion "
		'recuperamos las confirmaciones que pudo haber confirmado el usuario
		'la confirmaciones validas son :
		'si la hora es entre las 7 et las 24 -> sysdate
		'si es entre medianoche y 7 -> sysdate - 1	
		'Response.Write SQL
		array_conf= GetArrayRS(SQL)
		
		
		if Request.QueryString("msg") <> "" then
		%>
		<center><div class=error><%=Request.QueryString("msg")%></div></center>
		<%end if%>
		<table width=450 border=0>
		<tr>
			<td colspan=3><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<tr>
			<td colspan=3>Hoy : <%=hoy%> - <%=hora%><br><br></td>
		</tr>
		<form name="delete_conf" action="<%=asp_self()%>" method="post">
		<SCRIPT language="javascript">
		   function deleteItem(aduana, fecha) {
		      if (confirm("¿ Estas seguro de borrar esta confirmacion ?"))
				{
				document.delete_conf.aduana.value = aduana;
				document.delete_conf.fecha.value = fecha;
				document.delete_conf.submit();
				}
		      }
		   
		</SCRIPT>
		<input type=hidden name=aduana>
		<input type=hidden name=fecha>
		<input type=hidden name=etape value=2>
		</form>
		<tr bgcolor=goldenrod>
			<th colspan=3>Confirmaciones de aduanas recibidas :</th>
		</tr>
		<%if IsArray(array_conf) then
			for i = 0 to UBound(array_conf, 2)
				Response.Write "<tr"
				if i+1 mod 2 = 0 then Response.Write " bgcolor=""FFFFEE"""
				Response.Write ">" & vbCrLf & vbTab  & "<td align=center><li></td>" & vbCrLf 
				Response.Write vbTab & "<td>" & array_conf(0,i) & "</td>"  & vbCrLf & vbTab 
				Response.Write "<td><a href=""javascript:deleteItem('" & array_conf(1,i) & "','" & array_conf(2,i) & "');"">Borrar</a></td>"
				Response.Write "</tr>" & vbCrLf 
			next
		  else
			Response.Write "<tr><td colspan=3>No hay confirmaciones</td></tr>"
		  end if
		%>
		<tr>
			<td colspan=3>&nbsp;</td>
		</tr>
		<form name="valid_conf" action="<%=asp_self()%>" method="post">
		<SCRIPT language="javascript">
		   function testerRadio(radio) {
		      var ok=0;
		      var aduana;
		      for (var i=0; i<radio.aduana.length;i++) 
		      {//alert (radio.aduana[i].checked)
		         if (radio.aduana[i].checked)  {
		            ok=1
		         }
		      }
		      if (ok==1) {
				//alert('ok');
				radio.submit();
		      }
		      else {
				if (ok==0 && radio.aduana.checked) {
				//en caso que solo haya una aduana, radio.aduana.length regresa undefined.
					radio.submit();
					}
				else {
					alert("Favor de selecionar una aduana.");
					}
				}
		   }
		   
		</SCRIPT>
		<tr bgcolor=goldenrod>
			<th colspan=3>Confirmaciones pendientes : Aduanas</th>
		</tr>
		<%'selecionar la confirmaciones de aduanas que todavia se pueden hacer.
		  SQL = "  select dou.douclef, dou.douclef || ' - ' || InitCap(dou.douabreviacion)  " & _
				" from edouane dou  " & _
				" where douclef in ("& aduanas_usuario &")  " & _
				" and not exists ( " & _
				" select RCA.RCA_DOUCLEF  " & _
				" from REP_CONFIRMACION_ADUANAS RCA  " & _
				" where RCA.RCA_DOUCLEF = DOU.DOUCLEF " & _
				" and trunc(RCA.RCA_FECHA_CONF) = trunc(decode(sign(to_number(to_char(sysdate, 'HH24')) - 7), -1, sysdate-1, sysdate))  " & _
				" )" 
		  array_conf= GetArrayRS(SQL)
		
		if IsArray(array_conf) then
			for i = 0 to UBound(array_conf, 2)
				Response.Write "<tr"
				if i mod 2 = 0 then Response.Write " bgcolor=""FFFFEE"""
				Response.Write ">" & vbCrLf & vbTab  & _
					"<td align=center><input type=radio name=aduana value="&array_conf(0,i)&"></td>" & vbCrLf & vbTab & _
					"<td>" & array_conf(1,i) & "</td>" & vbCrLf & vbTab & _
					"<td>"& dia_conf &" - <i>(mm/dd/yyyy)</i><input type=hidden value=" & hoy & " name=fecha_" & array_conf(0,i) & ">" & _
					"</td>" & vbCrLf & vbTab & _
					"</tr>"		
			next	
		else
			Response.Write "<tr><td colspan=3>No se puede confirmar aduanas.</td></tr>"
		end if
		
		%>
		<tr>
			<td colspan=3>&nbsp;</td>
		</tr>
		<tr>
			<input type=hidden name=etape value=1>
			<td colspan=3 align=left><br><input type=button onclick="testerRadio(this.form);" class=buttonsOrange value=Validar></td>
		</tr>

		</form>
		</table>
		<%
		'function display_cal (num)
		'		display_cal = display_cal & vbCrLf & "<script type=""text/javascript"">" & vbCrLf 
		'		display_cal = display_cal & "<!--"& vbCrLf & vbTab
		'		display_cal = display_cal & "function FromCalendarCallback"& num &"(date, month, year)" & vbCrLf & vbTab
		'		display_cal = display_cal & "{" & vbCrLf & vbTab
		'		display_cal = display_cal & "date = month + '/' + date + '/' + year;" & vbCrLf & vbTab
		'		display_cal = display_cal & "document.valid_conf.fecha_"&num&".value = date;" & vbCrLf & vbTab
		'		display_cal = display_cal & "}"& vbCrLf 
		'		display_cal = display_cal & "// -->"& vbCrLf 
		'		display_cal = display_cal & "</script>" &  vbCrLf  &  vbCrLf 
		'		display_cal = display_cal & "<script language=""JavaScript"" type=""text/javascript"">" & vbCrLf 
		'		display_cal = display_cal & "<!--"& vbCrLf & vbTab
		'		display_cal = display_cal & "if (is_ie5up || is_nav6up || is_gecko){"& vbCrLf & vbTab
		'		display_cal = display_cal & "FromCalendar"&num&" = new dynCalendar('FromCalendar"&num&"', 'FromCalendarCallback"&num&"');"& vbCrLf & vbTab
		'		display_cal = display_cal & "FromCalendar"&num&".setOffset(10, 5);	} " & vbCrLf 
		'		display_cal = display_cal & "// -->"& vbCrLf 
		'		display_cal = display_cal & "</script>" &  vbCrLf  &  vbCrLf 

		'end function


case "1"
		'set date format to US
		Session.LCID = 1033
		'dim fecha_conf
		dim aduana, fecha_conf, rst, msg, ArrayRS, array_reporte, reporte_error, reporte_ok, j
		aduana = trim(Request.Form("aduana"))
		fecha_conf = Request.Form("fecha_" & Request.Form("aduana"))
		
		if aduana = "" or fecha_conf = "" then
			msg = "Por favor selecione una aduana y una fecha."
			Response.Redirect asp_self() & "?msg=" & msg
		end if
		
		'verificacion de los datos :
		'fecha:
		'verificamos que se confirmo la fecha segun los criterios definidos arriba.
		'
		'eso ya no debe de servir ya que restringimos las confirmaciones al dia correcto ;)
		SQL = "select 1 from dual" & _
			  " where to_date('" & Request.Form("fecha_" & aduana) & "', 'mm/dd/yyyy') = trunc(decode(sign(to_number(to_char(sysdate, 'HH24')) - 7), -1, sysdate-1, sysdate))"
		array_conf = GetArrayRS(SQL)
		
		if not IsArray(array_conf) then
			Response.Redirect asp_self() & "?msg=" & Server.URLEncode("Solo se puede confirmar el dia corriente<br>o, despues de medianoche, el dia pasado.")
		end if
		
		SQL = "select 1 from REP_CONFIRMACION_ADUANAS " & _
			  " where RCA_DOUCLEF = '"& aduana &"' and RCA_FECHA_CONF = trunc(to_date('"& fecha_conf &"','mm/dd/yyyy')) "
		'Response.Write SQL
		array_conf = GetArrayRS(SQL)
		
		if IsArray(array_conf) then
			Response.Redirect asp_self() & "?msg=" & Server.URLEncode("Esta fecha ha sido confirmada por esta aduana.")
		end if		

		SQL = "INSERT INTO REP_CONFIRMACION_ADUANAS ( " & _
			  " RCA_DOUCLEF, RCA_FECHA_CONF, DATE_CREATED, CREATED_BY)  " & _
			  " VALUES ('"& aduana &"', to_date('"& fecha_conf &"','mm/dd/yyyy') , sysdate , '" & Session("array_user")(0,0) & "' ) "
		
		set rst = Server.CreateObject("ADODB.Recordset")
		
		'insercion ok de la confirmacion
		rst.Open SQL, Connect(), 0, 1, 1
		msg = "Confirmacion insertada."
		
		'
		SQL = "select repdet.ID_CRON, repdet.ID_CRON || ' - ' || repdet.name " & VbCrLf 
		SQL = SQL & " from REP_DETALLE_REPORTE repdet " & VbCrLf 
		SQL = SQL & " where repdet.confirmacion = 2 " & VbCrLf 
		SQL = SQL & " order by 1"
		
		array_reporte = GetArrayRS(SQL)
		
		if not IsArray (array_reporte) then
			msg = msg & "<br>" & "No hay reportes que generar."
			Response.Redirect asp_self() & "?msg=" &server.URLEncode(msg)
		end if
		
		for i = 0 to UBound(array_reporte,2)
			SQL = " select RDA.RDA_DOUCLEF " & _
				  " from REP_CONF_ADUANAS_DETALLE RDA " & _
				  " where RDA.RDA_REPORTE_ID = '" & array_reporte(0,i) & "'" & _
				  " and RDA.RDA_DOUCLEF not in " & _
				  "   (select RCA.RCA_DOUCLEF from REP_CONFIRMACION_ADUANAS RCA " & _
				  "    where RCA.RCA_FECHA_CONF = trunc(decode(sign(to_number(to_char(sysdate, 'HH24')) - 7), -1, sysdate-1, sysdate)))" 
			'Response.Write SQL
			array_conf = GetArrayRS(SQL)
			
			if not IsArray(array_conf) then
				'como son generados como reportes temporales, aqui (last_conf_date_1, last_conf_date_2) debemos de poner la fecha del intervalo
				
				SQL = "update rep_detalle_reporte set last_conf_date_1 = trunc(decode(sign(to_number(to_char(sysdate, 'HH24')) - 7), -1, sysdate-1, sysdate)) " & _
					  " , last_conf_date_2 = trunc(decode(sign(to_number(to_char(sysdate, 'HH24')) - 7), -1, sysdate-1, sysdate)) " & _
					  " where id_cron = '" & array_reporte(0,i) & "'"
				'Response.Write SQL
				rst.Open SQL, Connect(), 0, 1, 1 
				
				SQL = "INSERT INTO REP_CHRON (ID_CHRON, ID_RAPPORT, ACTIVE, TEST) " & _
					  " VALUES ( seq_chron.nextval,'" & array_reporte(0,i) & "', 1,0) "
   
				'Response.Write SQL 
				rst.Open SQL, Connect(), 0, 1, 1 
				reporte_ok = reporte_ok & "<br><li> " & array_reporte(1,i)
			else
				reporte_error = reporte_error & "<br><li>Falta la confirmacion de la(s) aduana(s) <font color=red>" 
				for j = 0 to UBound(array_conf,2)
					reporte_error = reporte_error & array_conf (0,j)
					if j <> UBound(array_conf,2) then reporte_error = reporte_error & ", "
				next 
				reporte_error = reporte_error & "</font><br>para generar el reporte " & array_reporte(1,i) & "."
			end if
		
		next
		
		%>
		<html>
		<head>
		<title>Estado de reportes :</title>
		</head>
		<body>
		<%call print_style()%>
		<table width=400>
			<tr>
				<th>Reportes pendiente de generar :</th>
			</tr>
			<tr>
				<td><%if reporte_error <> "" then 
						Response.Write reporte_error
					  else
						Response.Write "Ningun reporte pendiente."
					  end if%></td>
			</tr>
			<tr><td><br></td></tr>
						<tr>
				<th>Reportes en generacion :</th>
			</tr>
			<tr>
				<td><%if reporte_ok <> "" then 
						Response.Write reporte_ok
					  else
						Response.Write "Ningun reporte por generar."
					  end if%></td>
			</tr>
			<tr><td><br><a href=menu.asp>Regresar</a></td></tr>
		</table>
		
		
		<%
		
case "2"
	set rst = Server.CreateObject("ADODB.Recordset")
	SQL = "delete from REP_CONFIRMACION_ADUANAS " & _
		  " where RCA_DOUCLEF = '"&Request.Form("aduana")&"'" & _
		  " and RCA_FECHA_CONF=to_date('"& Request.Form("fecha") &"','mm/dd/rrrr') "  
	
	'Response.Write SQL
	rst.Open SQL, Connect(), 0, 1, 1

	Response.Redirect asp_self() & "?msg=" & Server.URLEncode("Confirmacion borrada.")
end select
%>


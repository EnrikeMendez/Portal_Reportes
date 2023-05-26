<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Confirmacion de los reportes
Response.Expires = 0
call check_session()
'date format US
Session.LCID = 1033


Select Case Request("Etape")
	Case ""
%>
		<html>
		<head>
		<title>Confirmacion de reportes</title>
		</head>
		<body>
		<LINK media=screen href="./include/dyncalendar.css" type=text/css rel=stylesheet>
		<script src="./include/browserSniffer.js" type="text/javascript" language="javascript"></script>
		<script src="./include/dyncalendar.js" type="text/javascript" language="javascript"></script>
		<SCRIPT language="javascript">
		   function AlimFecha(tipo,num) {
		      var fecha_1 = "fecha_1_" + num;
		      var fecha_2 = "fecha_2_" + num;
		      alert (fecha_1);
		      //document.valid_conf.fecha_1.value = document.valid_conf.elements[fecha_1].value
		      
		   }
		   
		</SCRIPT>
		<%
		call print_style()
		Dim SQL, Array_conf, i, j, pend_conf, hoy, Array_det_conf, tab_confirmed, tab_pending_to_conf
		Dim hora, Array_det_conf2, SQL_02, k, l
		hora = Hour(Now)
		if Minute(Now) < 10 then
			hora = hora & ":0" & Minute(Now)
		else 
			hora = hora & ":" & Minute(Now)
		end if
		hoy = Month(now)& "/" & day(now)& "/" & Year(now)

		SQL ="  select id_cron from REP_DETALLE_reporte, rep_chron " & _
			 " where confirmacion=1  and REP_DETALLE_reporte.test = 0 and rep_chron.active = 1  " & _
			 " and rep_chron.id_rapport = REP_DETALLE_reporte.id_cron "
		' rajouter le param de restriction des confirmations
		
		'Response.Write SQL
		array_conf= GetArrayRS(SQL)

		if not IsArray(Array_conf) then
			Response.Write "Ninguno reporte para confimar."
			Response.End 
		end if
		'Response.Write "taille: " &Ubound(Array_conf,2)
		
		function display_cal (num, num_2)
				display_cal = display_cal & vbCrLf & "<script type=""text/javascript"">" & vbCrLf 
				display_cal = display_cal & "<!--"& vbCrLf & vbTab
				display_cal = display_cal & "function ToCalendarCallback"& num &"(date, month, year)" & vbCrLf & vbTab
				display_cal = display_cal & "{" & vbCrLf & vbTab
				display_cal = display_cal & "date = month + '/' + date + '/' + year;" & vbCrLf & vbTab
				display_cal = display_cal & "document.valid_conf.fecha_2_"&num&".value = date;" & vbCrLf & vbTab
				display_cal = display_cal & "}"& vbCrLf & vbTab
				display_cal = display_cal & "function FromCalendarCallback"& num &"(date, month, year)" & vbCrLf & vbTab
				display_cal = display_cal & "{" & vbCrLf & vbTab
				display_cal = display_cal & "date = month + '/' + date + '/' + year;" & vbCrLf & vbTab
				display_cal = display_cal & "document.valid_conf.fecha_1_"&num&".value = date;" & vbCrLf & vbTab
				display_cal = display_cal & "}"& vbCrLf 
				display_cal = display_cal & "// -->"& vbCrLf 
				display_cal = display_cal & "</script>" &  vbCrLf  &  vbCrLf 
				
			if num_2 = 1 then
				display_cal = display_cal & "<script language=""JavaScript"" type=""text/javascript"">" & vbCrLf 
				display_cal = display_cal & "<!--"& vbCrLf & vbTab
				display_cal = display_cal & "if (is_ie5up || is_nav6up || is_gecko){"& vbCrLf & vbTab
				display_cal = display_cal & "FromCalendar"&num&" = new dynCalendar('FromCalendar"&num&"', 'FromCalendarCallback"&num&"');"& vbCrLf & vbTab
				display_cal = display_cal & "FromCalendar"&num&".setOffset(10, 5);	} " & vbCrLf 
				display_cal = display_cal & "// -->"& vbCrLf 
				display_cal = display_cal & "</script>" &  vbCrLf  &  vbCrLf 
			else
				display_cal = display_cal & "<script language=""JavaScript"" type=""text/javascript"">" & vbCrLf 
				display_cal = display_cal & "<!--"& vbCrLf & vbTab
				display_cal = display_cal & "if (is_ie5up || is_nav6up || is_gecko){"& vbCrLf & vbTab
				display_cal = display_cal & "ToCalendar"&num&" = new dynCalendar('ToCalendar"&num&"', 'ToCalendarCallback"&num&"');"& vbCrLf & vbTab
				display_cal = display_cal & "ToCalendar"&num&".setOffset(10, 5);	} " & vbCrLf 
				display_cal = display_cal & "// -->"& vbCrLf 
				display_cal = display_cal & "</script>" &  vbCrLf  &  vbCrLf 
			end if
		
		end function
		
		k=0
		l=0
		for i = 0 to UBound(Array_conf, 2)
			'old query
			SQL =  "   SELECT REPDET.NAME,to_char(conf.CONF_DATE, 'mm/dd/yyyy') as conf_date  " & _ 
				"  ,to_char(conf.CONF_DATE_2, 'mm/dd/yyyy') as conf_date_2 " & _ 
				" , check_fecha_confirmacion(detconf.TIPO_confirmacion,conf.CONF_DATE) as Check_fecha  " & _ 
				" , display_fecha_confirmacion2(detconf.TIPO_confirmacion,conf.CONF_DATE,conf.CONF_DATE_2) as next_fecha  " & _ 
				" , detconf.ID_CONF, detconf.TIPO_confirmacion    " & _ 
				" , TIPO.NUM_COLUMNS  " & _ 
				" FROM REP_DETALLE_CONFIRMACION detconf, rep_confirmacion conf  " & _ 
				" , REP_DETALLE_REPORTE REPDET    " & _ 
				" , REP_TIPO_FRECUENCIA TIPO    " & _ 
				" WHERE detconf.ID_CONF = '"&Array_conf(0,i)&"'  " & _ 
				" and tipo.ID_TIPO_FREC = detconf.TIPO_CONFIRMACION  " & _ 
				" and  detconf.ID_CONF = conf.ID_CONF (+)   " & _ 
				" AND detconf.ID_CRON = REPDET.ID_CRON   " & _ 
				" order by conf.conf_date desc"
			'no longer use of rep_confirmacion_detalle
			SQL = " SELECT REPDET.NAME,to_char(conf.CONF_DATE, 'mm/dd/yyyy') as conf_date    " & VbCrLf 
			SQL = SQL & "    ,to_char(conf.CONF_DATE_2, 'mm/dd/yyyy') as conf_date_2    " & VbCrLf 
			SQL = SQL & "    , check_fecha_confirmacion2(repdet.FRECUENCIA,conf.CONF_DATE,conf.CONF_DATE_2) as Check_fecha    " & VbCrLf 
			SQL = SQL & "    , display_fecha_confirmacion4(repdet.FRECUENCIA,conf.CONF_DATE,conf.CONF_DATE_2,decode(conf.CONF_DATE,null,1,0)) as next_fecha    " & VbCrLf 
			SQL = SQL & "    , repdet.ID_CRON, repdet.FRECUENCIA , TIPO.NUM_COLUMNS    " & VbCrLf 
			
			'SQL = SQL & "    , decode(sign(conf.conf_date - repdet.LAST_CONF_DATE_1),1, 'Borrar') as borrar  " & VbCrLf 
			SQL = SQL & "   , decode(sign(to_date(to_char(conf.conf_date, 'dd/mm/yyyy') || ' ' || to_char(NVL(CRON.LAST_EXECUTION,sysdate-0.5), 'hh24:mi'),'dd/mm/yyyy hh24:mi') - decode(repdet.FRECUENCIA,1,sysdate-1, 5,sysdate-1,sysdate)),1, 'Borrar') as borrar  "  & VbCrLf 
			'NVL(CRON.LAST_EXECUTION,sysdate-0.5) en caso que no existe la fecha de ultima execucion, tomar una fecha mas vieja.
			'hay un decode sobre sysdate, porque si estamos en un tipo de confirmacion 1 (por el dia anteriror)
			'se necesita quitarle un dia a sysdate porque la confirmacion viene con el dia anterior del dia de la generacion del reporte.
			
			SQL = SQL & "    FROM rep_confirmacion conf    " & VbCrLf 
			SQL = SQL & "    , REP_DETALLE_REPORTE REPDET    " & VbCrLf 
			SQL = SQL & "    , REP_TIPO_FRECUENCIA TIPO    " & VbCrLf 
			SQL = SQL & "    , REP_CHRON CRON   " & VbCrLf 
			SQL = SQL & "    WHERE repdet.ID_CRON = '"&Array_conf(0,i)&"'    " & VbCrLf 
			SQL = SQL & "    and repdet.ID_CRON = conf.ID_CONF(+)   " & VbCrLf 
			SQL = SQL & "    and repdet.FRECUENCIA = tipo.ID_TIPO_FREC  " & VbCrLf 
			SQL = SQL & "    and ( check_fecha_confirmacion2(repdet.FRECUENCIA,conf.CONF_DATE,conf.CONF_DATE_2) ='ok'   " & VbCrLf 
			SQL = SQL & "    or NVL(conf_date_2, conf_date) >= trunc(sysdate) )  " & VbCrLf 
			SQL = SQL & "    and cron.ID_RAPPORT = REPDET.ID_CRON  " & VbCrLf 
			SQL = SQL & "    and cron.ACTIVE = 1  " & VbCrLf 
			SQL = SQL & "    order by check_fecha desc,conf.CONF_DATE desc  "

'Response.Write SQL


			SQL_02 = " SELECT * FROM (SELECT REPDET.NAME,to_char(conf.CONF_DATE, 'mm/dd/yyyy') as conf_date   " & _ 
				"   ,to_char(conf.CONF_DATE_2, 'mm/dd/yyyy') as conf_date_2   " & _ 
				"   , check_fecha_confirmacion2(repdet.FRECUENCIA,conf.CONF_DATE,conf.CONF_DATE_2) as Check_fecha   " & _ 
				"   , display_fecha_confirmacion4(repdet.FRECUENCIA,conf.CONF_DATE,conf.CONF_DATE_2,decode(conf.CONF_DATE,null,1,0)) as next_fecha   " & _ 
				"   , repdet.ID_CRON, repdet.FRECUENCIA , TIPO.NUM_COLUMNS   " & _ 
				"   FROM rep_confirmacion conf   " & _ 
				"   , REP_DETALLE_REPORTE REPDET   " & _ 
				"   , REP_TIPO_FRECUENCIA TIPO   " & _ 
				"   , REP_CHRON CRON   " & _ 
				"   WHERE repdet.ID_CRON = '"&Array_conf(0,i)&"'   " & _ 
				"   and repdet.ID_CRON = conf.ID_CONF(+)  " & _ 
				"   and repdet.FRECUENCIA = tipo.ID_TIPO_FREC " & _ 
				"   and cron.ID_RAPPORT = REPDET.ID_CRON " & _ 
				"   and cron.ACTIVE = 1 " & _ 
				"   order by conf.CONF_DATE desc "	 & _
				"   ) WHERE rownum = 1 "			
			'" WHERE detconf.ID_CRON = '"&Array_conf(0,i)&"'  " & _ 
			'Response.Write "<br>ok"&Array_conf(0,i)
			Response.Write SQL_02
			Array_det_conf = GetArrayRS(SQL)
			Array_det_conf2 = GetArrayRS(SQL_02)
			'Response.Write "<br>ok" & i&Array_det_conf(3,0)
	
		
			
			if IsArray(Array_det_conf) then
				for j = 0 to UBound(Array_det_conf,2) 
					tab_confirmed = tab_confirmed & "<tr" 
						if k mod 3 = 0 then tab_confirmed = tab_confirmed &  " bgcolor=""FFFFEE"""
					tab_confirmed = tab_confirmed & ">" & vbCrLf & vbTab  & _
									"<td align=right><li></td>" & vbCrLf & vbTab & _
									"<td>" & Array_det_conf(0,j) & " - " & Array_det_conf(1,j) 
									if Array_det_conf(7,j)="2" then
										tab_confirmed = tab_confirmed & "&nbsp;hasta&nbsp;" & Array_det_conf(2,j) 
									end if
					tab_confirmed = tab_confirmed & "</td>" & vbCrLf & vbTab & _
									"<td><a href=""#"" onclick=""deleteItem('"&Array_det_conf(5,j)&"','"&Array_det_conf(1,j)&"','0');"">"& Array_det_conf(8,j) &"</a></td> " & vbCrLf & _
									"</tr>" & vbCrLf 
					k = k+1
				next
			end if
			tab_pending_to_conf = tab_pending_to_conf & "<tr" 
						if l mod 3 = 0 then tab_pending_to_conf = tab_pending_to_conf &  " bgcolor=""FFFFEE"""
					tab_pending_to_conf = tab_pending_to_conf & ">" & vbCrLf & vbTab  & _
								"<td><input type=radio name=id_conf value="&Array_det_conf2(5,0)&"><input type=hidden name=tipo_conf_"&Array_det_conf2(5,0)&" value="&Array_det_conf2(6,0)&"></td>" & vbCrLf & vbTab & _
								"<td>" & Array_det_conf2(0,0) & " </td><td align=center><input type=text readonly size=11 class=""light"" value=""" & Left(Array_det_conf2(4,0),10) & """ name=fecha_1_"&Array_det_conf2(5,0)&">" & _
								display_cal(Array_det_conf2(5,0),1) 
								
								if Array_det_conf2(7,0)="2" then
									tab_pending_to_conf = tab_pending_to_conf & _
									"&nbsp;hasta <input type=text readonly size=11 class=""light"" value=" & Right(Array_det_conf2(4,0),10) & " name=fecha_2_"&Array_det_conf2(5,0)&">" & vbCrLf & _
									display_cal(Array_det_conf2(5,0),0) & "</td>" & vbCrLf 
								end if
								tab_pending_to_conf = tab_pending_to_conf & "</td></tr>" & vbCrLf 
			
			l = l+1
		next
		
		if Request.QueryString("msg") <> "" then
		%>
		<center><div class=error><%=Request.QueryString("msg")%></div></center>
		<%end if%>
		<table border=0 align=left cellspacing=3>
		<tr>
			<td colspan=2><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<tr>
			<td colspan=3>Hoy : <%=hoy%> - <%=hora%><br><br></td>
		</tr>
		<tr bgcolor=goldenrod>
			<th colspan=3>Reportes confirmados :</th>
		</tr>
		<%=tab_confirmed%>
		<form name="delete_conf" action="<%=asp_self()%>" method="post">
		<SCRIPT language="javascript">
		   function deleteItem(id_conf, fecha_1, param) {
		      if (confirm("� Estas seguro de borrar esta confirmacion ?"))
				{
				document.delete_conf.id_conf.value = id_conf;
				document.delete_conf.fecha_1.value = fecha_1;
				document.delete_conf.param.value = param;
				//alert (document.delete_conf.id_conf.value + "\n" + document.delete_conf.fecha_1.value + "\n" + document.delete_conf.param.value);
				document.delete_conf.submit();
				}
		      }
		   
		</SCRIPT>
		<input type=hidden name=fecha_1>
		<input type=hidden name=param>
		<input type=hidden name=id_conf>
		<input type=hidden name=etape value=2>
		</form>
		<tr>
			<td>&nbsp;</td>
		</tr>
		<form name="valid_conf" action="<%=asp_self()%>?etape=1" method="post">
		<SCRIPT language="javascript">
		   function testerRadio(radio) {
		      var ok=0;
		      for (var i=0; i<radio.id_conf.length;i++) 
		      {
		         if (radio.id_conf[i].checked) {
		            ok=1
		         }
		      }
		      if (ok==1) {
				//alert(radio.fecha_1_1.value);
				radio.submit();
		      }
		      else {
		      alert("Favor de selecionar una confirmacion.");
		      }
		   }
		   
		</SCRIPT>
		<tr bgcolor=goldenrod>
			<th colspan=3>Reportes pendientes de confirmacion :</th>
		</tr>
		<%=tab_pending_to_conf%>
		<tr>
			<td colspan=3 align=left><br>
			<input type=hidden name=fecha_1>
			<input type=hidden name=fecha_2>
			<input type=button onclick="testerRadio(this.form);" class=buttonsOrange value=Validar name=validar></td>
		</tr>
		
		</form>
		</table>
<%
case "1"
		'set date format to US
		Session.LCID = 1033
		dim id_conf, clef, rst, msg
		dim fecha_1, fecha_2, fecha_fin, fecha
		fecha_1 = Request.Form("fecha_1_"&Request.Form("id_conf"))
		fecha_2 = Request.Form("fecha_2_"&Request.Form("id_conf"))
		'Response.Write "<br>" & Request.Form("tipo_conf")
		
		'	Response.Write "<table>"
		'	for each i in Request.Form 
		'		Response.Write "<tr><td>" & i & "</td><td>" & Request.Form(i) & "</td></tr>"
		'	next
		'Response.Write "</table>"
		
		'get last day of month
		'Response.Write "fecha_1_2" & Request.Form("fecha_1_2")&"<br>"
		SQL = " select to_char(last_day(to_date('"&fecha_2&"', 'mm/dd/yyyy')), 'mm/dd/yyyy') from dual"
		array_conf= GetArrayRS(SQL)
		fecha_fin = array_conf(0,0)
		'Response.Write fecha_fin &"<br>"
		
		'Verificacion por tipo de confirmacion
		select case Request.Form("tipo_conf_"&Request.Form("id_conf")) 
			case "2" 
			'verificacion que las fechas son de una semana...
				if DateDiff("d", CDate(fecha_1), CDate(fecha_2)) <> 6 then
					msg= "La semana debe tener 7 dias."
				end if
			
				if Weekday (CDate(fecha_1)) <> 2 or Weekday (CDate(fecha_2)) <> 1 then
				'1->domingo, 2->Lunes
					msg = msg & "<br>La semana debe ir de Lunes a Domingo."
				end if
			
			case "3" 
			'verificacion que las fechas son de una quicena...
				if DateDiff("d", CDate(fecha_1), CDate(fecha_2)) > 15 then
					msg = "Debes escoger una quicena de intervalo."	
				end if
				Response.Write "<br>" & DateDiff("d", CDate(fecha_1), CDate(fecha_2))
				'Response.Write 
				
				select case Day(CDate(fecha_1))
					case 1
						if Day(CDate(fecha_2)) <> 15 then
							msg = msg & "<br>Una quicena es de 1 a 15."
						end if
					case 16
						if Day(CDate(fecha_2)) <> Day(CDate(fecha_fin)) then
							msg = msg & "<br>Una quicena es de 16 al fin del mes ("&Day(CDate(fecha_fin))&")."
						end if
					case else 
						msg = msg & "<br>Verificen sus datos."
				end select
			
			case "4" 
			'verificacion que las fechas son de un mes...
				if DateDiff("d", CDate(fecha_1), CDate(fecha_2)) > 31 then
					msg = "Debes escoger un mes de intervalo."	
				end if
				
				if Day(CDate(fecha_1)) <> "1" then
					msg = msg & "<br>El mes debe empezar el 1�."
				end if
				if Day(CDate(fecha_2)) <> Day(CDate(fecha_fin)) then
					msg = msg & "<br>El mes debe acabar al fin ("&Day(CDate(fecha_fin))&")."
				end if
			'case else
			
		'elseif Request.Form("tipo_conf_"&Request.Form("id_conf")) = "2" then
		end select 
		
		if msg = "" then
		
			set rst = Server.CreateObject("ADODB.Recordset")
			SQL = " insert into rep_confirmacion (ID_CONF,CONF_DATE,DATE_CREATED,CONF_DATE_2,PARAM,CREATED_BY) values ('"&Request.Form("id_conf") & _
				  "' ,to_date('"& fecha_1 &"','mm/dd/yyyy'), sysdate, to_date('"& fecha_2 &"','mm/dd/yyyy'), '0','" & Session("array_user")(0,0) & "')"
		
		
			'Response.Write "<br>"&SQL
			'Response.End 
			
			Session("SQL") = SQL
			rst.Open SQL, Connect(), 0, 1, 1
			
			
			SQL = " select to_char(last_execution, 'mm/dd/yyyy hh24:mi') " & _
				  " from rep_detalle_reporte, rep_chron  " & _
				  " where id_cron = '"& Request.Form("id_conf") &"' " & _
				  " and id_cron = id_rapport "
			Array_conf = GetArrayRS(SQL)
			Response.Write SQL
			Response.End
			
		'Response.Write SQL & "<br>" 
			'penser ou pas de conf arrivee
			if IsArray(Array_conf) then
				fecha = Array_conf(0,0)
			
				if fecha_2 = "" then
					if fecha = "" then
						fecha_2 = CDate( fecha_1 & " 00:00")
					else
						fecha_2 = CDate( fecha_1 & " " & Hour(fecha) & ":" & Minute(fecha))
					end if
				else 
					if fecha = "" then
						fecha_2 = CDate( fecha_2 & " 00:00")
					else
						fecha_2 = CDate( fecha_2 & " " & Hour(fecha) & ":" & Minute(fecha))
					end if
				end if
				if Request.Form("tipo_conf_"&Request.Form("id_conf")) = "1" or Request.Form("tipo_conf_"&Request.Form("id_conf")) = "5" then
					fecha_2 = fecha_2 + 1
				end if
				'verificamos si la fecha de confirmacion llega tarde o no
				'si no llega a tiempo :
				' o es una confirmacion por el proximo periodo de fecha (no vamos a confirmar el ultimo reporte que tiene error)
				'CDate(fecha_2) < now es falso porque a lo menos es el ultimo rango de fecha del periodo y debe de ser anterior a now
				'Response.Write "<br> "& fecha_2& "<br>" & fecha
				'Response.Write fecha_2 & " " & fecha
				dim tarde
				tarde = 0
				
				'SQL = "select 1 " & _
				'	  "from dual " & _
				'	  "where CHECK_FECHA_CONFIRMACION2 ('"& Request.Form("tipo_conf_" & Request.Form("id_conf")) & "',to_date('"&FormatDateTime (fecha_1,2) &"', 'mm/dd/rrrr'),to_date('"&FormatDateTime (fecha_2,2) &"', 'mm/dd/rrrr')) = 'ok' "
				'Array_conf = GetArrayRS(SQL)
				
				SQL = "select 1 from rep_chron " & _
					  "where last_execution < to_date('"&FormatDateTime (fecha_1,2) &"', 'mm/dd/rrrr hh24:mi') " & _
					  " and id_rapport = "&Request.Form("id_conf")&" "
				'Response.Write SQL 
				Array_conf = GetArrayRS(SQL)
dim maintenant  
				if Weekday(now) = 2 and Request.Form("tipo_conf_"&Request.Form("id_conf")) = "1" then
					maintenant = now - 2
				else 
					maintenant = now
				end if				
'Response.End 
				if CDate(fecha_2) <= maintenant then 
					if not IsArray(Array_conf) then
					'Response.Write "tard"
					'Response.End 
					'''''''''''''
					''''''''''''
					''''''''''
					'attetion au test = 1
					SQL = "insert into rep_chron (id_chron, id_rapport, priorite, test) " & _
						  " values (seq_chron.nextval, '"& Request.Form("id_conf") &"', 4, 0 ) "
					'Response.Write sql
					rst.Open SQL, Connect(), 0, 1, 1
					tarde = 1
					
					'regresamos a la fecha_2 original en caso que fue cambiado
					if Request.Form("tipo_conf_"&Request.Form("id_conf")) = "1" or Request.Form("tipo_conf_"&Request.Form("id_conf")) = "5" then
						fecha_2 = fecha_2 - 1
					end if
					
					SQL = "update rep_detalle_reporte set last_conf_date_1 = to_date('"& FormatDateTime (fecha_1,2) &"', 'mm/dd/rrrr'), last_conf_date_2 = to_date('"& FormatDateTime (fecha_2,2) &"', 'mm/dd/rrrr') " & _
						  " where id_cron = '"& Request.Form("id_conf") &"' "
					'Response.Write "<br>" & SQl
					rst.Open SQL, Connect(), 0, 1, 1
				end if 
				end if 
			else 
				'confirmacion tarde :
				SQL = "insert into rep_chron (id_chron, id_rapport, priorite) " & _
					  "values (seq_chron.nextval, '"& Request.Form("id_conf") &"', 4 ) "
					  'Response.End 
				rst.Open SQL, Connect(), 0, 1, 1
			end if 

			set msg = Server.CreateObject( "JMail.Speedmailer" )
			msg.SendMail Get_mail("web_reports"), Get_mail("IT_1"), "confirmation recue : " & Request.Form("id_conf"), "fecha: " & FormatDateTime (fecha_1,2), Get_IP("mail_server")
			'Response.End 

			'Response.Write fecha & "<br>" & fecha_2
			dim msg_url
			msg_url = Server.URLEncode("Confirmacion insertada.")
			if tarde = 1 then msg_url = msg_url & Server.URLEncode("<br>Favor de esperar de recibir el correo de este reporte antes de confirmar otro.") 
			Response.Redirect asp_self() & "?msg=" & Server.URLEncode("Confirmacion insertada.")
	

		else
			Response.Redirect asp_self() & "?msg=" & Server.URLEncode(msg)
		end if



	
case "2"	
	set rst = Server.CreateObject("ADODB.Recordset")
	SQL = "delete from rep_confirmacion " & _
		  " where id_conf = '"&Request.Form("id_conf")&"'" & _
		  " and conf_date=to_date('"& Request.Form("fecha_1") &"','mm/dd/rrrr') "  & _
		  " and param='"& Request.Form("param") &"' "  
	
	rst.Open SQL, Connect(), 0, 1, 1

	Response.Redirect asp_self() & "?msg=" & Server.URLEncode("Confirmacion borrada.")
end select



%>


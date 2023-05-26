<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Agregacion de reportes ponctuales
Response.Expires = 0
dim qa
	qa = "_qa"



' <JEMV: ESTO PERMITE EL ACCESO LOCAL, descomentar antes de liberar el módulo:
if qa = "" then
	if Left(Request.ServerVariables("REMOTE_ADDR"),7) <> "192.168" then
	'avoid external acceses 
		Response.Write "<center class=error><H1>Access Prohibited</h1></center>"
		Response.End 
	end if
end if
' JEMV>

'call check_session()
dim SQL, arrayRS,i,j
Select Case Request.Form("Etape")
	Case ""

%>
		<html>
		<head>
		<title>Seleccion del reporte</title>
		</head>
		<body>
		<%
		call print_style()
		
		%>
		<style>
		td {
		    font-size: 10px;
		    line-height: 2em;    
		}
			.lblMSG {
				align-content: center;
				font-size: 11.5px;
				padding: 10px;
				text-align: center;
				width: 100%;
			}
			.green {
				border: solid 1px #DDF0DD;
				background-color: #EBFFEB;
			}
		</style>
		<%if Request("msg") <> "" then	Response.Write "<table width='900px'><tr><td align='center'><font class='lblMSG green' size='2'>" & Request("msg") & "</font></td></tr></table><br/>"%>
		<table width=1200 border=0>
		<form name="valid_conf" action="<%=asp_self()%>" method="post">
		<tr bgcolor=goldenrod>
			<th colspan="4">Seleciona un reporte :</th>
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<tr>
		  <th>Aduana</th>
		  <th colspan="2">Trading</th>
		  <th>Aduana-Trading</th>
		</tr>
		<tr valign="top">
		  <td>
				<%
				'Reportes Aduana 
				dim tab_tmp, cons
				cons = 1
				SQL = " select rep.id_rep, rep.name, rep.num_of_param " & _
					  " from rep_reporte rep" & _
					  " where id_rep in (74,59,182,143,104 ,81,24,23,22,21,20,17,13,187,189,192,193,194,198,216,219,220,223,224,225,229,134,230,232,239,241,245,246,255,258,225,238,267,272,274,273,275,277,205,210,231,269,296,162,298) order by name "
				arrayRS = GetArrayRS(SQL)
				
				for i=0 to UBound(arrayRS,2)
					Response.Write "<input type=radio name=id_rep value="&arrayRS(0,i)&">&nbsp;" & cons & ". " & arrayRS(1,i) & "<br>" & vbCrLf  & vbTab 
				    cons = cons + 1
				next
				
				%>
		</td>
		  <td>
				<%
'< JEMV: Agrego el ID 334 (Reservacion de Guias CD)
				'Reportes Trading parte 1
				SQL = " select rep.id_rep, rep.name, rep.num_of_param, 1 " & _
					  " from rep_reporte rep" & _
				  " where id_rep in (196,172,170,169,159,158,157,155,153,151,150,149,147,139,176," & _
				  " 138,137,121,112,110,100,99,98,96,94,89,82,43,42,37,171,191,80,197,213,218,222, " & _ 
				  " 226,234,235,237,249,256,250,251,259,261,262,266,271,175,265,303,304,334, 335) " & _
				  " union all " & _
				  " select rep.id_rep, rep.name, rep.num_of_param, 2 " & _
					  " from rep_reporte rep" & _
				  " where id_rep in (320) " & _
				  " order by 4, 2 "
' JEMV >
					  
				arrayRS = GetArrayRS(SQL)
				cons = 1
				for i=0 to UBound(arrayRS,2)
					Response.Write "<input type=radio name=id_rep value="&arrayRS(0,i)&">&nbsp;" & cons & ". " & arrayRS(1,i) & "<br>" & vbCrLf  & vbTab 
				    cons = cons + 1
				    if i = CInt(UBound(arrayRS,2) / 2) then
				       Response.Write " </td><td>"
				    end if
				next
				%>
		</td>
		<td>
				<%
				'Reportes Aduana-Trading
				SQL = " select rep.id_rep, rep.name, rep.num_of_param " & _
					  " from rep_reporte rep" & _
					  " where id_rep in (174) order by  name "
					  
				arrayRS = GetArrayRS(SQL)
				cons = 1
				for i=0 to UBound(arrayRS,2)
					Response.Write "<input type=radio name=id_rep value="&arrayRS(0,i)&" >&nbsp;" & cons & ". " & arrayRS(1,i) & "<br>" & vbCrLf  & vbTab 
				    cons = cons + 1
				next
				
				%>
		</td>
		</tr>
		<tr>
		
		<tr>
			<td align=left colspan=6>
			<!--<br>
			Nombre de la lista de contactos :<br>
			<input type=text name=list_name>-->
			<br><br>
			<input type="hidden" name=etape value=1>
			<input type="hidden" name=test value=1>
			<input type=submit class=buttonsOrange ><br><br>
			</td>
		</tr>
		
		</form>
		</table>
<%
	Case "1"
%>	

		<html>
		<head>
		<title>Seleccion de los parametros</title>
		<LINK media=screen href="include/dyncalendar.css" type=text/css rel=stylesheet>
		<script src="include/browserSniffer.js" type="text/javascript" language="javascript"></script>
		<script src="include/dyncalendar.js" type="text/javascript" language="javascript"></script>
		</head>
		<body>
		<%call print_style()%>
			<%if Request("msg") <> "" then	Response.Write "<table width='600px'><tr><td align='center' colspan='2'><font class='lblMSG' size='2'>" & Request("msg") & "</font></td></tr></table><br/>"%>
		<table width=600 border=0 >
		<form name="valid_conf" action="<%=asp_self()%>" method="post">
		<tr bgcolor=goldenrod>
		<th colspan=2>Captura los parametros : </th>
	</tr>

	<%
	SQL = " select num_of_param " & _
		  " from rep_reporte " & _
		  " where id_rep='"& Request.Form("id_rep") &"' " 
'		response.write SQL
'		response.End
	arrayRS = GetArrayRS(SQL)
	'Response.Write arrayRS(0,0)
	dim num_param
	num_param = arrayRS(0,0)
	'Response.Write num_param & "<br>" 
	if num_param <> "0" then
		SQL = "select  "
		for i=1 to CInt(num_param)
			if i <> 1 then SQL = SQL & ","
			SQL = SQL & "name_param_" & i
			SQL = SQL & ", opcion_" & i
		next
		SQL = SQL & " from rep_reporte where id_rep='"& Request.Form("id_rep") &"' "  
	
		'Response.Write sql
		arrayRS = GetArrayRS(SQL)
		if IsArray(arrayRS) then
			
			for i=0 to UBound(arrayRS,1) step 2
			    
			    if Left(Trim(UCase(arrayRS(i,0))), 5) = "CEDIS" then
                    Response.Write "<tr>" & vbCrLf & vbTab
				    Response.Write "<td valign=top>"
				    if arrayRS(i+1,0) = "1" then Response.Write "<i>"
				    Response.Write (i/2) + 1  &".&nbsp;&nbsp;" & arrayRS(i,0) & "&nbsp;:&nbsp;&nbsp;"
				    if arrayRS(i+1,0) = "1" then Response.Write "</i>"
				    Response.Write "</td><td>"
				    call print_bodega("valid_conf", "param_"& (i/2) + 1)
				    Response.Write "<input type=hidden name=opcion_param_" & ((i/2) + 1) & " value="& arrayRS(i+1,0) & "></td></tr>" & vbCrLf
			    elseif Left(Trim(UCase(arrayRS(i,0))), 6) = "ESTADO" then
                    Response.Write "<tr>" & vbCrLf & vbTab
				    Response.Write "<td valign=top>"
				    if arrayRS(i+1,0) = "1" then Response.Write "<i>"
				    Response.Write (i/2) + 1  &".&nbsp;&nbsp;" & arrayRS(i,0) & "&nbsp;:&nbsp;&nbsp;"
				    if arrayRS(i+1,0) = "1" then Response.Write "</i>"
				    Response.Write "</td><td>"
				    call print_estado("valid_conf", "param_"& (i/2) + 1)
				    Response.Write "<input type=hidden name=opcion_param_" & ((i/2) + 1) & " value="& arrayRS(i+1,0) & "></td></tr>" & vbCrLf
			    elseif Trim(UCase(arrayRS(i,0))) = "SUCURSAL" then
                    Response.Write "<tr>" & vbCrLf & vbTab
				    Response.Write "<td valign=top>"
				    if arrayRS(i+1,0) = "1" then Response.Write "<i>"
				    Response.Write (i/2) + 1  &".&nbsp;&nbsp;" & arrayRS(i,0) & "&nbsp;:&nbsp;&nbsp;"
				    if arrayRS(i+1,0) = "1" then Response.Write "</i>"
				    Response.Write "</td><td>"
				    call print_sucursal("valid_conf", "param_"& (i/2) + 1)
				    Response.Write "<input type=hidden name=opcion_param_" & ((i/2) + 1) & "  value="& arrayRS(i+1,0) & "></td></tr>" & vbCrLf
				elseif display_parametros(Request.Form("id_rep"), (i/2) + 1, arrayRS(i+1,0)) <> "" then
				    Response.Write "<tr>" & vbCrLf & vbTab
				    Response.Write "<td>"
				    if arrayRS(i+1,0) = "1" then Response.Write "<i>"
				    Response.Write (i/2) + 1  &".&nbsp;&nbsp;" & arrayRS(i,0) & "&nbsp;:&nbsp;&nbsp;"
				    if arrayRS(i+1,0) = "1" then Response.Write "</i>"
				    Response.Write "</td><td>"& display_parametros(Request.Form("id_rep"), (i/2) + 1, arrayRS(i+1,0)) &"<input type=hidden name=opcion value="& arrayRS(i+1,0) &"> </td>" & vbCrLf 
				    Response.Write "</tr>" & vbCrLf				
				else
				    Response.Write "<tr>" & vbCrLf & vbTab
				    Response.Write "<td>"
				    if arrayRS(i+1,0) = "1" then Response.Write "<i>"
				    Response.Write (i/2) + 1  &".&nbsp;&nbsp;" & arrayRS(i,0) & "&nbsp;:&nbsp;&nbsp;"
				    if arrayRS(i+1,0) = "1" then Response.Write "</i>"
				    Response.Write "</td><td><input type=text name=param_"& (i/2) + 1 &" size=10 class=light><input type=hidden name=opcion value="& arrayRS(i+1,0) &"> </td>" & vbCrLf 
					Response.Write "</tr>" & vbCrLf
					'<<CHG-DESA-23032023-01:Agregar el campo de la clave del pedimento en el módulo reporte_anomalia
					if Request.Form("id_rep") = "104" then
						if ((i/2) + 1) = CInt(num_param) then
							Response.Write "<tr><td><i>4.&nbsp;&nbsp;Clave de pedimento&nbsp;:&nbsp;&nbsp;</i>"
							Response.Write "</td><td><input type=text name=param_4 size=10 class=light><input type=hidden name=clavePed value="& arrayRS(i+1,0) &"> </td></tr>" & vbCrLf 
						end if 
					end if
					'CHG-DESA-23032023-01>>
				end if
			next
			Response.Write "<tr><td colspan=2><i>En italico, los parametros son opcional.</i></td></tr>" & vbCrLf
		end if
	end if
	%>
	<input type=hidden name=num_param value=<%=num_param%>>
	<tr>
	<tr>	
		<script type="text/javascript">
	<!--
		// Calendar callback. When a date is clicked on the calendar
		// this function is called so you can do as you want with it
		function ToCalendarCallback(date, month, year)
		{
			date = date + '/' + month + '/' + year;
			document.valid_conf.fecha_2.value = date;
		}
		function FromCalendarCallback(date, month, year)
		{
			date = date + '/' + month + '/' + year;
			document.valid_conf.fecha_1.value = date;
		}
	// -->
	</script>	
	<%dim mi_fecha
	if Request.Form("id_rep") = "13" then 
	    mi_fecha= "01/01/2003"
	    mi_fecha = day(now) & "/" & month(now) & "/" & year(now) -1
	else
	    mi_fecha = day(now) & "/" & month(now) & "/" & year(now)
	end if%>
	
	<%if Request.Form("id_rep")<> 271 and Request.Form("id_rep")<> 176 then	%>

		<td>Fecha 1 :</td>
		<td><input type=text name=fecha_1 class=light size=11 readonly onclick="javascript:{if (this.value=='dd/mm/yyyy') {this.value=''}};" value="<%=mi_fecha%>"  maxlength=12> <%'onblur="Remplace(this.form.file_name);"%>
			<script language="JavaScript" type="text/javascript">
				<!--
				//if (is_ie5up || is_nav6up || is_gecko){
					FromCalendar = new dynCalendar('FromCalendar', 'FromCalendarCallback');
					FromCalendar.setOffset(5, 5);
				//	}
				//-->
			</script>	
		</td>
	<%else%>	
	<input type="hidden" name=fecha_1  value="<%=mi_fecha%>" > 
	<%end if	%>	
	</tr>
	</tr>
	<tr>
		<%if Request.Form("id_rep")<> 271 and Request.Form("id_rep")<> 176  then	%>
		<td valign=top>Fecha 2 :</td>
		<td><input type=text name=fecha_2 class=light size=11 readonly onclick="javascript:{if (this.value=='dd/mm/yyyy') {this.value=''}};" value="<%Response.Write day(now) & "/" & month(now) & "/" & year(now)%>"  maxlength=12> <%'onblur="Remplace(this.form.file_name);"%>
			<script language="JavaScript" type="text/javascript">
				<!--
				//if (is_ie5up || is_nav6up || is_gecko){
					ToCalendar = new dynCalendar('ToCalendar', 'ToCalendarCallback');
					ToCalendar.setOffset(5, 5);
				//	}
				//-->
			</script>
		</td>
		<%end if%>
	</tr>
	<%if Request.Form("id_rep") = "108" then 
	    'para el reporte de evolucion agregamos 11 rangos de fecha mas
	    
	    for i = 1 to 12
	%>
	<tr>
	  <td>Rango<%=i%> :</td>
		<td>De <input type='text' name='rango_<%=i%>_from' class='light' size='11' readonly>
			<script language='JavaScript' type='text/javascript'>
				<!--
				function Rango_<%=i%>_FromCalendarCallback(date, month, year)
		        {
		        	date = date + '/' + month + '/' + year;
		        	document.valid_conf.rango_<%=i%>_from.value = date;
		        }
				//if (is_ie5up || is_nav6up || is_gecko){
					Rango_<%=i%>_FromCalendar = new dynCalendar('Rango_<%=i%>_FromCalendar', 'Rango_<%=i%>_FromCalendarCallback');
					Rango_<%=i%>_FromCalendar.setOffset(5, 5);
				//	}
				//-->
			</script>	
		A <input type='text' name='rango_<%=i%>_to' class='light' size='11' readonly>
			<script language='JavaScript' type='text/javascript'>
				<!--
				function Rango_<%=i%>_ToCalendarCallback(date, month, year)
		        {
		        	date = date + '/' + month + '/' + year;
		        	document.valid_conf.rango_<%=i%>_to.value = date;
		        }
				//if (is_ie5up || is_nav6up || is_gecko){
					Rango_<%=i%>_ToCalendar = new dynCalendar('Rango_<%=i%>_ToCalendar', 'Rango_<%=i%>_ToCalendarCallback');
					Rango_<%=i%>_ToCalendar.setOffset(5, 5);
				//	}
				//-->
			</script>	
		</td>
	  </tr>
	<%  next
	
	end if%>
	<!--<tr>
		<td colspan=2>
		<br>
		<%SQL = " select distinct carpeta from rep_archivos " & _
				" Where id_rep= '"& Request.Form("id_rep") &"' " 
		dim carpeta
		'arrayRS = GetArrayRS(SQL)
		if IsArray(arrayRS) then
			carpeta = arrayRS(0,0)
		end if
		%>
		Carpeta : <input type=text class=light name=carpeta value="<%=carpeta%>" onblur="Remplace(this.form.carpeta);" maxlength=30>
		<br>(no poner espacio en el nombre y eligir uno sencillo)<br>-->
		<!--<br>Nombre del archivo :<br>
		<input type=text name=file_name class=light size=30 onblur="Remplace(this.form.file_name);"  maxlength=30><br>
		Nombre del reporte :<br>
		<input type=text name=report_name class=light size=30 maxlength=50><br>
		</td>
	</tr>-->
	<script language="javascript">
function CheckDate(d) {
      // Cette fonction vérifie le format JJ/MM/AAAA saisi et la validité de la date.
      // Le séparateur est défini dans la variable separateur
      var amin=1999; // année mini
      var amax=<%=year(now)%>; // année maxi
      var separateur="/"; // separateur entre jour/mois/annee
      //var j=(d.substring(0,2));
      //var m=(d.substring(3,5));
      //var a=(d.substring(6));
      var j = (d.split("/")[0]); // jour
      var m = (d.split("/")[1]); // mois
      var a = (d.split("/")[2]); // année
      
      var ok=1;
      if ( ((isNaN(j))||(j<1)||(j>31)) && (ok==1) ) {
         alert("El dia no es corecto."); ok=0;
      }
      if ( ((isNaN(m))||(m<1)||(m>12)) && (ok==1) ) {
         alert("El mes no es corecto."); ok=0;
      }
      if ( ((isNaN(a))||(a<amin)||(a>amax)) && (ok==1) ) {
         alert("El año no es corecto."); ok=0;
      }
      /*if ( ((d.substring(2,3)!=separateur)||(d.substring(5,6)!=separateur)) && (ok==1) ) {
         alert("Usar los separadores "+separateur); ok=0;
      }*/
      if (ok==1) {
         var d2=new Date(a,m,j);
         j2=d2.getDate();
         m2=d2.getMonth()+1;
         a2=d2.getYear();
         if (a2<=100) {a2=1900+a2}
         if ( (j!=j2)||(m!=m2)||(a!=a2) ) {
            alert("La fecha "+d+" no existe !");
            ok=0;
         }
      }
      return ok;
   }

// Checks a string to see if it in a valid date format
// of (D)D/(M)M/(YY)YY and returns true/false
function isValidDate(s) {
    // format D(D)/M(M)/(YY)YY
    var dateFormat = /^\d{1,4}[\.|\/|-]\d{1,2}[\.|\/|-]\d{1,4}$/;

    if (dateFormat.test(s)) {
        // remove any leading zeros from date values
        s = s.replace(/0*(\d*)/gi,"$1");
        var dateArray = s.split(/[\.|\/|-]/);
      
        // correct month value
        dateArray[1] = dateArray[1]-1;

        // correct year value
        if (dateArray[2].length<4) {
            // correct year value
            dateArray[2] = (parseInt(dateArray[2]) < 50) ? 2000 + parseInt(dateArray[2]) : 1900 + parseInt(dateArray[2]);
        }

        var testDate = new Date(dateArray[2], dateArray[1], dateArray[0]);
        if (testDate.getDate()!=dateArray[0] || testDate.getMonth()!=dateArray[1] || testDate.getFullYear()!=dateArray[2]) {
            return false;
        } else {
            return true;
        }
    } else {
        return false;
    }
}

function check_opcion(param, op, num_param) {
   var error = "";
		//alert (op.length);
		//alert (op.value);
	if (num_param.value == 1)
		//caso que solo hay un unico parametro entonces op, param no son arrays
		{if ((op.value == 0) && (param.value == "")) 
		   {return 1; }
		 else {return 0;}
		 }
	else
	{
      for (var i=0; i<op.length;i++) 
       {
         if ((op[i].value == 0) && (param[i].value == "")) 
          { if (error != "")
             {error = error + "," +  (i+1); }
             else {error = error + (i+1); }
		  }
		if (error == "")
		 {return 0;}
		else 
		 {return error;}
	   }
    }
}

function Trim(s) 
{
  // Remove leading spaces and carriage returns
  
  while ((s.substring(0,1) == ' ') || (s.substring(0,1) == '\n') || (s.substring(0,1) == '\r'))
  {
    s = s.substring(1,s.length);
  }

  // Remove trailing spaces and carriage returns

  while ((s.substring(s.length-1,s.length) == ' ') || (s.substring(s.length-1,s.length) == '\n') || (s.substring(s.length-1,s.length) == '\r'))
  {
    s = s.substring(0,s.length-1);
  }
  return s;
}


function ValidateForm()
{
	var msg = "";
	/*if (CheckDate(document.valid_conf.fecha_1.value) == 0) 
		{msg+="- la fecha 1 no es correcta.\n"};
	if (CheckDate(document.valid_conf.fecha_2.value) == 0 )
		{msg+="- la fecha 2 no es correcta.\n"};*/	
	/*if (document.valid_conf.carpeta.value == "")
		{msg+="- no hay nombre de carpeta.\n"};
	if (document.valid_conf.file_name.value == "")
		{msg+="- el archivo no tiene nombre.\n"};
	if (document.valid_conf.report_name.value == "")
		{msg+="- el reporte no tiene nombre.\n"};*/
	if (document.valid_conf.correo.value == "")
		{msg+="- no hay correo electronico.\n"};
	if (valid_conf.num_param.value >= 1)
	    {
			var check
			check = check_opcion(document.valid_conf.param, document.valid_conf.opcion, document.valid_conf.num_param) 
			if (check != 0 )
				{msg+="- los siguientes parametros : "+check+" son necesarios.\n"};
		}  
   
	
	if (msg == "") 
		{valid_conf.submit();}
	else
		{alert ("Verifica los datos : \n"+msg);
		 return false;}
}

   function Remplace(expr) {
      var new_name = expr.value;
      var Forbidden_char = "\\/:*?\"\'<>|;,.~ &";
      var i=0;
      for (var i=0; i<= new_name.length; i++)
		{for (var j=0; j < Forbidden_char.length; j++ )
			{if (new_name.charAt(i) == Forbidden_char.charAt(j))
				{//alert (Forbidden_char.charAt(j));
				 new_name=new_name.substring(0,i)+'_'+new_name.substring(i+1);
				}
			}
		 }
		expr.value = new_name;
   }
</script>
<tr>
		<td colspan=2>
		<br>
		Correo electronico : <input type=text class=light size=80 name=correo maxlength=200 onblur="javascript:this.value=Trim(this.value);">
		<br>(se puede poner varios separados con ";" o "," )<br>
		</td>
	</tr>
<tr>
	<td>
		<input type="hidden" name="id_rep" value="<%=Request.Form("id_rep")%>">
		<input type="hidden" name="num_param" value="<%=Request.Form("num_param")%>">
		<input type="hidden" name="etape" value="2"><br>
		<input type="hidden" name=test value=1>
		<input type=button class="buttonsOrange" value="Validar"  onclick="javascript:ValidateForm();" id=button1 name=button1><br><br>
	</td>
</tr>

</form>
<br></table>
<%	
	case "2"
%>
<html>
<head>
<title>Agregar nuevo reporte - Verificacion</title>
</head> 
<body>
<%call print_style()%>
	<%if Request("msg") <> "" then	Response.Write "<table width='350px'><tr><td align='center' colspan='2'><font class='lblMSG' size='2'>" & Request("msg") & "</font></td></tr></table><br/>"%>
<table border=0 cellpadding=2 cellspacing=0 width=350>
<form action=<%=asp_self()%> method="post" name=valid_conf  onsubmit='document.valid_conf.valid_button.disabled=true'>


<tr bgcolor=goldenrod>
	<th colspan=2>Verifica los datos :</th>
</tr>
<tr>
	<td>Tipo del reporte</td>
	<td>
	<%SQL = "Select id_rep || ' - ' || name from rep_reporte " & _
			"Where id_rep ='"&Request.Form("id_rep")&"' " 
	arrayRS = GetArrayRS(SQL)
	Response.Write arrayRS(0,0)
	%></td>
</tr>
<!--<tr>
	<td>Nombre del archivo</td>
	<td><%=TAGescape(Request.Form("file_name"))%></td>
</tr>
<tr>
	<td>Nombre del reporte</td>
	<td><%=TAGescape(Request.Form("report_name"))%></td>
</tr>
<tr>
	<td>Carpeta</td>
	<td><%=TAGescape(Request.Form("carpeta"))%></td>
</tr>-->
	<%dim correo_HTML
	correo_HTML = check_mail(Request.Form("correo")) '''''''''''''''''''''''''''''''''''''''
	
	if correo_HTML = "" then
		'no hay errores...
%><tr>
	<td>Correo</td>
	<td><%=TAGescape(Request.Form("correo"))%></td>
</tr>
<%
	else
		Response.Write "<tr><td bgcolor=red valign=top><font size=""2"" color=white>Error :</font></td>" 
		Response.Write vbTab & "<td bgcolor=red valign=top><font color=white>" 
		for i = 0 to UBound(split(correo_HTML, "|"))
			Response.Write "<li>" & split(correo_HTML, "|")(i) & "<br>" & vbCrLf 
		next
		Response.Write "</td></tr>"
	end if
	%>

<tr>

<tr>
	<td>Fecha 1</td>
	<td><%=TAGescape(Request.Form("fecha_1"))%></td>
</tr>
<tr>
	<td>Fecha 2</td>
	<td><%=TAGescape(Request.Form("fecha_2"))%></td>
</tr>

	<%dim params, params_form, opciones, opciones_form, param_HTML
	params_form = Request.Form("param_1") & "|" & Request.Form("param_2") & "|" & Request.Form("param_3") & "|" & Request.Form("param_4")
	opciones_form = Request.Form("opcion_param_1") & "|" & Request.Form("opcion_param_2") & "|" & Request.Form("opcion_param_3") & "|" & Request.Form("opcion_param_4")
	params = split(params_form, "|")
	opciones = split(opciones_form, "|")

	param_HTML = verif_parametros(params, opciones, Request.Form("id_rep"))
	
	if param_HTML = "" then
		'no hay errores...
	    for i = 0 to UBound(params)
	    	Response.Write "<tr>" & vbCrLf & vbTab & "<td>Parametro " & i+1 & "</td><td>" & _
	    					TAGescape(params(i)) & "</td>" & vbCrLf & "</tr>"
	    next
	else
		Response.Write "<tr><td bgcolor=red valign=top><font size=""2"" color=white>Error :</font></td>" 
		Response.Write vbTab & "<td bgcolor=red valign=top><font color=white>" 
		for i = 0 to UBound(split(param_HTML, "|"))
			Response.Write "<li>" & split(param_HTML, "|")(i) & "<br>" & vbCrLf 
		next
		Response.Write "</td></tr>"
	end if
	
	dim clef, rangos
	for each clef in Request.Form
	    if left(clef, 6) = "rango_" and right(clef, 4) = "from" then
	        
	        if Request.Form("rango_" & split(clef, "_")(1) & "_from") <> ""  and Request.Form("rango_" & split(clef, "_")(1) & "_to") <> "" then   'split(clef, "_")(1) para recuperar el numero
	            if rangos <> "" then rangos = rangos & "|"
	            rangos = rangos & Request.Form("rango_" & split(clef, "_")(1) & "_from") & _
	                     "-" & Request.Form("rango_" & split(clef, "_")(1) & "_to")
	            
	        end if
	    end if
	next
	%>

<tr>
	<td colspan=2><br></td>
</tr>
<tr>
	<td colspan=2>
		<%if param_HTML = "" and correo_HTML = "" then%>
		<font color=green>Los datos son correctos ?</font>
		<%else%>
		<font color=red>Los datos no son correctos, favor de verificar los.</font>
		<br><a href="javascript:history.back();">Regresar</a>
		<%end if%>
	</td>
</tr>
<tr>
	<td colspan=2>
	<input type="hidden" name="file_name" value="<%=TAGescape(SQLEscape(Request.Form("file_name")))%>">
	<input type="hidden" name="report_name" value="<%=TAGescape(SQLEscape(Request.Form("report_name")))%>">
	<input type="hidden" name="carpeta" value="<%=TAGescape(SQLEscape(Request.Form("carpeta")))%>">
	<input type="hidden" name="param_1" value="<%=TAGescape(SQLEscape(Request.Form("param_1")))%>">
	<input type="hidden" name="param_2" value="<%=TAGescape(SQLEscape(Request.Form("param_2")))%>">
	<%Dim param_3
	if Request.Form("id_rep") = "108" then
	    param_3 = rangos
	  else
	    param_3 = TAGescape(SQLEscape(Request.Form("param_3")))
	  end if
	
	%>
	<input type="hidden" name="param_3" value="<%=param_3%>">
	<input type="hidden" name="param_4" value="<%=TAGescape(SQLEscape(Request.Form("param_4")))%>">
	<input type="hidden" name="fecha_1" value="<%=Request.Form("fecha_1")%>">
	<input type="hidden" name="fecha_2" value="<%=Request.Form("fecha_2")%>">
	<input type="hidden" name="correo" value="<%=Request.Form("correo")%>">
	<input type="hidden" name="id_rep" value="<%=Request.Form("id_rep")%>">
	<input type="hidden" name="num_param" value="<%=Request.Form("num_param")%>">
	<input type="hidden" name="etape" value="3">
	<%if param_HTML = "" then%>
	<input type=submit value=Validar name=valid_button class=buttonsOrange  >
	<input type=button value=Cancelar class=buttonsOrange onclick="javascript:history.back();" id>
	<%end if%>
	</td>
</tr>
</table>	
	<%
	case "3"
'insercion de los datos en la base...
	Dim  id_rep, rst, mail_message
	set rst = Server.CreateObject("ADODB.Recordset")
	'Response.Write "<br><br>Detail rapport : <br>"
	SQL = "select SEQ_REPORTE_DETALLE.nextval from dual"
	arrayRS = GetArrayRS(SQL)
	id_rep = arrayRS(0,0)
		'<--- CHG-DESA-07122021-
	if Request.Form("id_rep") = 334 or Request.Form("id_rep") = 335 then 
		SQL = "select replace(name, ' ', '_') || '_' || '" &  Request.Form("param_" & 1) & "_' || to_char(to_date('"& Request.Form("fecha_1") &"', 'dd/mm/yyyy'), 'dd-mm-yyyy') "
		if Request.Form("fecha_2") <> "" and Request.Form("fecha_2") <> Request.Form("fecha_1") then
			SQL = SQL & "|| '_to_'  || to_char(to_date('"& Request.Form("fecha_2") &"', 'dd/mm/yyyy'), 'dd-mm-yyyy') "
		end if	
	else 
		SQL = "select replace(name, ' ', '_') || '_' || to_char(to_date('"& Request.Form("fecha_1") &"', 'dd/mm/yyyy'), 'dd-mm-yyyy') "
		if Request.Form("fecha_2") <> "" and Request.Form("fecha_2") <> Request.Form("fecha_1") then
			SQL = SQL & "|| '_to_'  || to_char(to_date('"& Request.Form("fecha_2") &"', 'dd/mm/yyyy'), 'dd-mm-yyyy') "
		end if	
	end if
	'CHG-DESA-07122021-	-->
	
	SQL = SQL & "  as nombre from rep_reporte where id_Rep='" & Request.Form("id_rep") & "'"
	arrayRS = GetArrayRS(SQL)
	
	'Response.Write SQL
	dim file_name
	file_name = replace(arrayRS(0,0), " ", "_")
	
	
	SQL = " insert into rep_detalle_reporte (id_cron, id_rep, name, " & _
		  " file_name, carpeta " 

	params = split(Request.Form("param"), ",")
	for i = 1 to Request.Form("num_param")
		SQL = SQL & ", param_" & i
	next  
	
	if Request.ServerVariables("REMOTE_ADDR") <> "" then
		SQL = SQL & ", IP_ADDRESS"
	end if
	
	 'Response.Write(SQL)
	
	SQL = SQL & " , dest_mail, LAST_CONF_DATE_1, LAST_CONF_DATE_2, DAYS_DELETED, last_created ) "  & _
		" values ('" & id_rep &"', '" & Request.Form("id_rep") & "' " & _
		" , '"& file_name &"' "  & _
		" , '"& file_name &"', 'temp' " 
		
	'params = split(Request.Form("param"), ",")
	for i = 1 to Request.Form("num_param")
	'<<CHG-DESA-23032023-01:Se concatena el valor de la clave del pedimento al tipo de Exportación o Importación
		'SQL = SQL & ", '" & Request.Form("param_" & i) & "' "
		if Request.Form("id_rep") = "104" then
			if i = CInt(Request.Form("num_param")) then
				SQL = SQL & ", '" & Request.Form("param_" & i)& "|" & Request.Form("param_4") & "' "
			else
				SQL = SQL & ", '" & Request.Form("param_" & i) & "' "
			end if
		else
			SQL = SQL & ", '" & Request.Form("param_" & i) & "' "
		end if
	'CHG-DESA-23032023-01>>
	next
	
	if Request.ServerVariables("REMOTE_ADDR") <> "" then
		SQL = SQL & ", '" & Request.ServerVariables("REMOTE_ADDR") & "'"
	end if
	
	dim days_deleted
	
	' se modifica la cantidad de dias el 2014/06/30
	if Request.Form("id_rep") = "96" then
	    days_deleted = 7		' tenia 90
	elseif Request.Form("id_rep") = "226" or Request.Form("id_rep") = "158" then
	    days_deleted = 5		' tenia 90
	elseif Request.Form("id_rep") = "298" or Request.Form("id_rep") = "134" or Request.Form("id_rep") = "229" or Request.Form("id_rep") = "267" or Request.Form("id_rep") = "239" or Request.Form("id_rep") = "219" then
	    days_deleted = 4		' tenia 90
	else
	    days_deleted = 7		' tenia 30
	end if
	SQL = SQL & " , '"& Request.Form("correo") &"' , " & _
	"to_date('"& trim(Request.Form("fecha_1")) &"', 'dd/mm/yyyy'), to_date('"& trim(Request.Form("fecha_2")) &"', 'dd/mm/yyyy'), "& days_deleted &", sysdate)"
Response.Write SQL
		Response.end
	rst.Open SQL, Connect(), 0, 1, 1
	'Response.Write SQL
	'Response.Write SQL & "<br><br>SQL Cron :<br>"
	'Response.End 
	
	SQL = "insert into rep_chron (id_chron, id_rapport " & _
		  ", priorite, test, active) values (SEQ_CHRON.nextval, '"& id_rep &"', 1,0, 1) " 

	
	'Response.Write SQL 
	
	rst.Open SQL, Connect(), 0, 1, 1

''''''''''''''''''''''''''' Cambio de prioridad a pedimentos Instantáneos
	SQL = "update rep_chron set priorite = 5 " & VbCrLf
	SQL = SQL & "	where id_chron in ( "  & VbCrLf
	SQL = SQL & "		select id_chron from rep_chron where id_rapport in ( "  & VbCrLf
	SQL = SQL & "				select rep_det.id_cron "  & VbCrLf
	SQL = SQL & "				from rep_detalle_reporte rep_det "  & VbCrLf
	SQL = SQL & "				join rep_chron chron on chron.id_rapport = rep_det.id_cron "  & VbCrLf
	SQL = SQL & "				join rep_reporte reporte on reporte.id_rep = rep_det.id_rep "  & VbCrLf
	SQL = SQL & "				where chron.active = 1 " & VbCrLf
	SQL = SQL & "				and chron.MINUTES is null and chron.HEURES is null and chron.JOURS is null and chron.MOIS is null and chron.JOUR_SEMAINE is null and chron.LAST_EXECUTION is null "  & VbCrLf
	SQL = SQL & "				and reporte.id_rep = 173 "  & VbCrLf
	SQL = SQL & "			) "  & VbCrLf
	SQL = SQL & "		)"
	
	rst.Open SQL, Connect(), 0, 1, 1


		''''''''''''''''''''''''''' Cambio de prioridad a pedimentos Expediente Aduanal Antolin (16/08/2021)
		SQL = "update rep_chron set priorite = 6 " & VbCrLf
	SQL = SQL & "	where id_chron in ( "  & VbCrLf
	SQL = SQL & "		select id_chron from rep_chron where id_rapport in ( "  & VbCrLf
	SQL = SQL & "				select rep_det.id_cron "  & VbCrLf
	SQL = SQL & "				from rep_detalle_reporte rep_det "  & VbCrLf
	SQL = SQL & "				join rep_chron chron on chron.id_rapport = rep_det.id_cron "  & VbCrLf
	SQL = SQL & "				join rep_reporte reporte on reporte.id_rep = rep_det.id_rep "  & VbCrLf
	SQL = SQL & "				where chron.active = 1 " & VbCrLf
	SQL = SQL & "				and chron.MINUTES is null and chron.HEURES is null and chron.JOURS is null and chron.MOIS is null and chron.JOUR_SEMAINE is null and chron.LAST_EXECUTION is null "  & VbCrLf
	SQL = SQL & "				and reporte.id_rep = 311 "  & VbCrLf
	SQL = SQL & "			) "  & VbCrLf
	SQL = SQL & "		)"
	
	rst.Open SQL, Connect(), 0, 1, 1
		'''''''''''''''''''''''''''''''''''''''''''''



	
	dim mail, mail_server, mail_footer, mail_fecha
	set mail = Server.CreateObject( "JMail.Message" )
 
	mail_server = Get_IP ("mail_server")
	mail_footer = vbCrLf & vbCrLf & vbCrLf & _
              "*********************************************************" & vbCrLf & _
              "This is a message automatically generated, please contact " & vbCrLf & _
              Get_Mail("webmaster") & " for any question or to unsubscribe."

	
	
	'manda correo de bienvenido
	SQL = "select name from rep_reporte where id_rep = '"& Request.Form("id_rep") &"'"
	'Response.Write "<br> SQL mail " & SQL 
	arrayRS = GetArrayRS(SQL)
	
	if IsArray(arrayRS) then
		mail.Charset = "UTF-8"
		mail.ISOEncodeHeaders = False
		if qa <> "" then
			mail.Subject = "Notification : report " & Request.Form("file_name") & " < " & arrayRS(0,0) &  " > (Q.A.)"
		else
			mail.Subject = "Notification : report " & Request.Form("file_name") & " < " & arrayRS(0,0) &  " >"
		end if
		mail.From = Get_Mail("web_reports")
		mail.FromName = Get_Mail("web_reports_name")
		
		'hacemos que estan todas las direcciones si hay varias separadas con "," o ";"
		for i=0 to UBound(split(replace(Request.Form("correo"), ",", ";"), ";"))
			mail.AddRecipient trim(split(replace(Request.Form("correo"), ",", ";"), ";")(i))
		next
		
		mail_message = "<FONT FACE=""Arial,Helvetica"" SIZE=""2""><b>Notification</b><br><br>" & _
					"Thank you for choosing Logis products and services.<br><br>" & vbCrLf & _
					"We are processing your request on line and  we'll do  our best to answer your questions based on the information you provided." & vbCrLf  & _
					"Please wait, you will receive another mail to give you the link." & vbCrLf  & _
					"For any question, please contact the Webmaster." & vbCrLf  & _
					"<br><br>Regards<br><br>Logis Web Site.</FONT>"

		
		'body en HTML
		mail.HTMLBody = display_mail("http://" & Get_IP ("WEB_1"), mail_message )

		mail.body = notag(Replace(mail_message , "<br>", vbCrLf )) & _
					 vbCrLf & vbCrLf & mail_footer
					 
		
		mail.Send mail_server
	end if	

	

	Response.Redirect "reporte_anomalia" & qa & ".asp?msg=" & Server.URLEncode("El reporte se esta generando. Checa tu correo.")
end select
%>




</body>
</html>

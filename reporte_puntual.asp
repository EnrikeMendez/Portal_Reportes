<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Agregacion de reportes ponctuales
Response.Expires = 0
call check_session()
dim SQL, arrayRS,i,j, array_tmp
Select Case Request.Form("Etape")
	Case ""
		if Request.Form("reporte_num") = "" then Response.Redirect "menu.asp"
%>
		<html>
		<head>
		<title>Seleccion del reporte</title>
		</head>
		<body>
		<%
		call print_style()
		
		%>
		<table width=450 border=0 >
		<form name="valid_conf" action="<%=asp_self()%>" method="post">
		<tr bgcolor=goldenrod>
			<th>Seleciona un reporte :</th>
		</tr>
				<%
				'Response.Write Request.Form("id_mail") 
				dim tab_tmp
				SQL = " select rep.id_rep, rep.name, rep.num_of_param " & _
					  " from rep_reporte rep" & _
					  " where id_rep in ("& Request.Form("reporte_num") &") "
					  '7,8,11,17,18,14,13, 6
				'_ a changer .... limitacion al reporte de indice de calidad
				'Response.Write SQL
				arrayRS = GetArrayRS(SQL)
				
				for i=0 to UBound(arrayRS,2)
				    Response.Write "<tr><td>" & vbCrLf & vbTab 
				    if arrayRS(0,i) = "202" then
				        'verificar si se esta generando el reporte para bloquearlo
				        SQL = "SELECT COUNT(0) " & VbCrlf
                        SQL = SQL & "   FROM REP_DETALLE_REPORTE " & VbCrlf
                        SQL = SQL & "     , REP_CHRON " & VbCrlf
                        SQL = SQL & "   WHERE ID_RAPPORT = ID_CRON " & VbCrlf
                        SQL = SQL & "   AND ID_REP = 202 "
                        array_tmp = GetArrayRS(SQL)
                        if array_tmp(0, 0) <> "0" then
                            Response.Write "<font color='red'>Un reporte de pedimentos se esta generando, favor de esperar a que se termina.</font>" & vbCrLf  & vbTab 
                        else
					        Response.Write "<input type=radio name=id_rep value="&arrayRS(0,i)&" checked>&nbsp;" & arrayRS(1,i) & vbCrLf  & vbTab 
                        end if
					else
					    Response.Write "<input type=radio name=id_rep value="&arrayRS(0,i)&" checked>&nbsp;" & arrayRS(1,i) & vbCrLf  & vbTab 
					end if
					for j=0 to CInt(arrayRS(2,i))
						'Response.Write "<br>"
						'a voir ou on mets les parametres... + nom rapport
					next
					Response.Write "</td>" & vbCrLf 
					Response.Write "</tr>" & vbCrLf
				next
				
				%>
		<tr>
		
		<tr>
			<td align=left colspan=6>
			<!--<br>
			Nombre de la lista de contactos :<br>
			<input type=text name=list_name>-->
			<br><br>
			<input type="hidden" name=etape value=1>
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
		</head>
		<body>
		<%call print_style()%>
		<table width=450 border=0 >
		<form name="valid_conf" action="<%=asp_self()%>" method="post">
		<tr bgcolor=goldenrod>
		<th colspan=2>Captura los parametros : </th>
	</tr>

	<%
	SQL = " select num_of_param " & _
		  " from rep_reporte " & _
		  " where id_rep='"& Request.Form("id_rep") &"' " 

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
			
			'for i=0 to UBound(arrayRS,1) step 2
			'	Response.Write "<tr>" & vbCrLf & vbTab
			'	Response.Write "<td>"
			'	if arrayRS(i+1,0) = "1" then Response.Write "<i>"
			'	Response.Write (i/2) + 1  &".&nbsp;&nbsp;" & arrayRS(i,0) & "&nbsp;:&nbsp;&nbsp;"
			'	if arrayRS(i+1,0) = "1" then Response.Write "</i>"
			'	Response.Write "</td><td><input type=text name=param size=10 class=light onblur=""Remplace(this.form.param" 
			'	if UBound(arrayRS,1) >= 2 then Response.Write "["& i/2 &"]" 
			'	'en caso que no haya varios elementos, me da un error JS si uso this.form.param[0] y que param no es un array
			'	Response.Write ");""><input type=hidden name=opcion value="& arrayRS(i+1,0) &" </td>" & vbCrLf 
			'	Response.Write "</tr>" & vbCrLf
			'next
			'Response.Write "<tr><td colspan=2><i>En italico, los parametros son opcional.</i></td></tr>" & vbCrLf
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
						if Request.Form("id_rep")="260" and ((i/2) + 1)=4 then
								Response.Write "</td><td><input type=text name=param_"& (i/2) + 1 &" size=10 class=light value=clientes><input type=hidden name=opcion value="& arrayRS(i+1,0) &"> </td>" & vbCrLf 
						else
				    Response.Write "</td><td><input type=text name=param_"& (i/2) + 1 &" size=10 class=light><input type=hidden name=opcion value="& arrayRS(i+1,0) &"> </td>" & vbCrLf 
						end if
						'Response.Write "</td><td><input type=text name=param_"& (i/2) + 1 &" size=10 class=light><input type=hidden name=opcion value="& arrayRS(i+1,0) &"> </td>" & vbCrLf 
				    Response.Write "</tr>" & vbCrLf
				end if
			next
			Response.Write "<tr><td colspan=2><i>En italico, los parametros son opcional.</i></td></tr>" & vbCrLf

		end if
	end if
	
	%>
	<input type=hidden name=num_param value=<%=num_param%>>
	<tr>
	<tr>
		<td>Fecha 1 :</td>
		<td><input type=text name=fecha_1 class=light size=11 onclick="javascript:{if (this.value=='dd/mm/yyyy') {this.value=''}};" value="dd/mm/yyyy" onblur="javascript:isValidDate(this.value);"  maxlength=12><br> <%'onblur="Remplace(this.form.file_name);"%>
		</td>
	</tr>
	</tr>
	<tr>
		<td valign=top>Fecha 2 :</td>
		<td><input type=text name=fecha_2 class=light size=11 onclick="javascript:{if (this.value=='dd/mm/yyyy') {this.value=''}};" value="dd/mm/yyyy" onblur="javascript:isValidDate(this.value);"  maxlength=12><br> <%'onblur="Remplace(this.form.file_name);"%>
		formato : dd/mm/yyyy - (poner la misma fecha en 2 si se necesita un unico dia)</td>
	</tr>
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
	   }
	  if (error == "")
		 {return 0;}
	  else 
		 {return error;}
    }
}
   
function ValidateForm()
{
	var msg = "";
	if (!isValidDate(document.valid_conf.fecha_1.value)) 
		{msg+="- la fecha 1 no es correcta.\n"};
	if (!isValidDate(document.valid_conf.fecha_2.value))
		{msg+="- la fecha 2 no es correcta.\n"};	
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
      var Forbidden_char = "\\/:*?\"\'<>|;.~ & ";
      var i=0;
      for (var i=0; i<= new_name.length; i++)
		{for (var j=0; j < Forbidden_char.length; j++ )
			{if (new_name.charAt(i) == Forbidden_char.charAt(j))
				{//alert (Forbidden_char.charAt(j));
				 new_name=new_name.substring(0,i)+new_name.substring(i+1);
				}
			}
		 }
		expr.value = new_name;
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
<table border=0 cellpadding=2 cellspacing=0 width=400>
<form name="valid_conf" action="<%=asp_self()%>" method="post">
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
	correo_HTML = check_mail(Request.Form("correo"))
	
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
	<%Dim param_3, rangos
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
	<input type=button value=Cancelar class=buttonsOrange onclick="javascript:history.back();"  id=button1 name=button1>
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
	
	SQL = "select replace(name, ' ', '_') || '_' || to_char(to_date('"& Request.Form("fecha_1") &"', 'dd/mm/yyyy'), 'dd-mm-yyyy') "
	if Request.Form("fecha_2") <> "" then
		SQL = SQL & "|| '_to_'  || to_char(to_date('"& Request.Form("fecha_2") &"', 'dd/mm/yyyy'), 'dd-mm-yyyy') "
	end if
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
	
	 
	
	SQL = SQL & " , dest_mail, LAST_CONF_DATE_1, LAST_CONF_DATE_2, DAYS_DELETED, last_created ) "  & _
		" values ('" & id_rep &"', '" & Request.Form("id_rep") & "' " & _
		" , '"& file_name &"' "  & _
		" , '"& file_name &"', 'temp' " 
		
	'params = split(Request.Form("param"), ",")
	for i = 1 to Request.Form("num_param")
		SQL = SQL & ", '" & Request.Form("param_" & i) & "' "
	next
	SQL = SQL & " , '"& Request.Form("correo") &"' , " & _
	"to_date('"& trim(Request.Form("fecha_1")) &"', 'dd/mm/yyyy'), to_date('"& trim(Request.Form("fecha_2")) &"', 'dd/mm/yyyy'), 7, sysdate)"
	
	rst.Open SQL, Connect(), 0, 1, 1
	
	'Response.Write SQL & "<br><br>SQL Cron :<br>"
	
	SQL = "insert into rep_chron (id_chron, id_rapport " & _
		  ", priorite, test, active) values (SEQ_CHRON.nextval, '"& id_rep &"', 5,0,1) " 

	
	'Response.Write SQL 
	
	rst.Open SQL, Connect(), 0, 1, 1
	
	dim mail, mail_server, mail_footer, mail_fecha
	set mail = Server.CreateObject( "JMail.Message" )
 
	mail_server = Get_IP("MAIL_SERVER")
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
		mail.Subject = "Notification : report " & Request.Form("file_name") & " < " & arrayRS(0,0) &  " (4)>"
		mail.From = Get_Mail("web_reports")
		mail.FromName = Get_Mail("web_reports_name")
		'para debug, estoy en los contactos ;)
		mail.AddRecipientBCC Get_Mail("IT_1")
		
		'hacemos que estan todas las direcciones si hay varias separadas con "," o ";"
		for i=0 to UBound(split(replace(Request.Form("correo"), ",", ";"), ";"))
			mail.AddRecipient trim(split(replace(Request.Form("correo"), ",", ";"), ";")(i))
		next
		
		mail_message = "<FONT FACE=""Arial,Helvetica"" SIZE=""2""><b>Notification : </b> Report : " & arrayRS(0,0) & " is being generated.<br><br>" & _
					"Thank you for choosing Logis products and services.<br><br>" & vbCrLf & _
					"We are processing your request on line and  we'll do  our best to answer your questions based on the information you provided." & vbCrLf  & _
					"Please wait, you will receive another mail to give you the link." & vbCrLf  & _
					"For any question, please contact the Webmaster." & vbCrLf  & _
					"<br><br>Regards<br><br>Logis Web Site.</FONT>"

		
		'body en HTML
		mail.HTMLBody = display_mail("http://" & Get_IP("web_1"), mail_message )

		mail.body = notag(Replace(mail_message, "<br>", vbCrLf )) & _
					 vbCrLf & vbCrLf & mail_footer
		
		mail.Send mail_server
	end if	

	

	Response.Redirect "menu.asp?msg=" & Server.URLEncode("El reporte se esta generando.<br>Checka su correo.")
end select
%>

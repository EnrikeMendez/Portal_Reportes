<%@ Language=VBScript %>
<% option explicit 
on error goto 0
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Agregacion de reportes
Response.Expires = 0
call check_session()


Select Case Request.Form("Etape")
	Case ""

'borrar todos los contactos si hay...
Session("id_mail") = ""
Session("id_mail_error") = ""
Session("id_lista") = ""
%>
		<html>
		<head>
		<title>Gestion de correos</title>
		</head>
		<body>
		<%
		call print_style()
		
		%>

		<table width=450 border=0 >
		<tr>
			<td colspan=2><a href=menu.asp>Menu general</a><br><br></td>
		</tr>
		<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
		<tr bgcolor=goldenrod>
			<th Align=left>Etapa 1 :</th>
		</tr>
		<script LANGUAGE="JavaScript">
		function check_data()
			{
			if (document.form_1.cli_num.value == "")
				{alert("Favor de capturar un numero de cliente.");}
			else document.form_1.submit();  		
			}
		</script>
		<tr>
			<td>
				Hay que checkar los contactos necesarios para enviar el reporte.
				<br>Entra el numero de cliente para desplegar la lista de los contactos.
				<br>Si no agregaste los contactos, dale un clic <a href=mail.asp>aqui</a>.<br>
			</td>
		</tr>
		<form name="form_1"  action="<%=asp_self()%>" method="post">
		<tr>
			<td>
				<br>Numero de cliente :
				<br><input type=text name=cli_num size=8 class=light>
			</td>
		</tr>
		<tr>
			<td align=left>
			<input type="hidden" name=etape value=1>
			<input type=button onClick="check_data();" class=buttonsOrange value=Validar><br><br>
			</td>
		</tr>
		</form>
		
		</table>
<%
case "1"
		dim msg, SQL, arrayRS, i, j, arrayRS2, arrayRS3
		
		SQL = "select 1 from eclient where cliclef='" & CLng(SQLescape(Request.Form("cli_num"))) & "'"
		arrayRS = GetArrayRS(SQL)
		if not IsArray(arrayRS) then
			Response.Redirect asp_self & "?msg=" & Server.URLEncode ("Cliente no existante.")
		end if

		'dim i, filtro
		%>
		<html>
		<head>
		<title>Agregar nuevo reporte - Etapa 2</title>
		</head>
		<body>
		<%call print_style()
		SQL = " select id_mail, nombre, mail, decode(client_num, 9929,'Logis',client_num) as client_num " & _
				  " , decode(tercero, 1, 'Si', '') as tercero " & _
				  " From rep_mail " & _
				  " Where client_num in ('" & SQLescape(Request.Form("cli_num")) & "', 9929)" & _
				  " and status = 1 " & _
				  " order by client_num, tercero desc, nombre "
		arrayRS = GetArrayRS(SQL)
		if not IsArray(arrayRS) then
			Response.Write "No hay contactos."
			Response.Write "<br>Agregar los <a href=mail.asp>aqui</a>."
			Response.End 
		end if
		SQL = " select id_mail, nombre, mail " & _
				  " From rep_mail " & _
				  " Where client_num = 9929 " & _
				  " and status = 1 " & _
				  " order by nombre "
		arrayRS2 = GetArrayRS(SQL)
		
		SQL = " select ID_LISTA, NOMBRE " & _
			  " from REP_LISTA " & _
			  " where id_cliente in ("& SQLescape(Request.Form("cli_num")) & ",9929)" & _
			  " order by 1"
		arrayRS3 = GetArrayRS(SQL)		
		%>
		<table border=0 >
		<form name="form_2"  action="<%=asp_self()%>" method="post" onsubmit="return ValidateForm(this,'id_mail')">
		<tr bgcolor=goldenrod>
			<th colspan=7>Seleciona los contactos :</th>
		</tr>
		<tr><td colspan=7><br></td></tr>
		<%if IsArray(arrayRS3) then%>
		<tr><th colspan=7>Listas de contactos :</th></tr>
		<%for i = 0 to UBound(arrayRS3,2)
			Response.Write "<tr"
			if i mod 3 = 0 then Response.Write  " bgcolor=""FFFFEE"""
			Response.Write ">" & vbCrLf & vbTab 
			Response.Write "<td><input type=checkbox name=id_lista value="& arrayRS3(0, i) &"></td>"
			Response.Write "<td>" & i+1 & "</td>" & vbCrLf & vbTab 
			Response.Write "<td colspan=4><a href=""javascript:void(0);"" onclick=""ver_lista('"& arrayRS3(0, i) &"');"">" & arrayRS3(1, i) & "</a></td>" & vbCrLf & vbTab  
			Response.Write "</tr>" & vbCrLf 
		next%>
		<tr><td colspan=7><br></td></tr>
		<%end if%>
		<tr bgcolor="goldenrod" align=center>
			<td colspan=2>.</td>
			<td>Nombre</td>
			<td>Correo</td>
			<td>Cliente</td>
			<td>Tercero</td>
		</tr>
		<script language = "Javascript">
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

		function ValidateForm(dml){
			len = dml.elements.length;
			var i=0;
			var mail_ok = "- escoge al menos un contacto.\n";
			var mail_error = "- escoge al menos un contacto en caso de error.";
			for( i=0 ; i<len ; i++) {
				if ((dml.elements[i].name=='id_mail') && (dml.elements[i].checked==1)) mail_ok ="";
				if ((dml.elements[i].name=='id_mail_error') && (dml.elements[i].checked==1)) mail_error ="";
				}
			if ((mail_ok != "") || (mail_error != ""))
				{alert("Verifica los contactos :\n" + mail_ok + mail_error);
				 return false;
				}
			return true;
		}
		// -->
		<!--
		function ver_lista(lista)
		{
            // <- CHG-DESA-30062021-01
			window.open("ver_lista.asp?liste=" + lista + "&Num=" + <% Request.Form("cli_num") %> + "&tipo=grupo", "Lista_contactos", "toolbar=no, location=no, directories=no, status=yes, scrollbars=yes, resizable=yes, copyhistory=no, width=500, height=400, left=300, top=50");
			// CHG-DESA-30062021-01 ->
		}
		//-->
        </script>
		<%
		for i = 0 to UBound(arrayRS,2)
			Response.Write "<tr"
			if i mod 3 = 0 then Response.Write  " bgcolor=""FFFFEE"""
			Response.Write ">" & vbCrLf & vbTab 
			Response.Write "<td><input type=checkbox name=id_mail value="& arrayRS(0, i) &"></td>"
			Response.Write "<td>" & i+1 & "</td>" & vbCrLf & vbTab 
			Response.Write "<td>" & arrayRS(1, i) & "</td>" & vbCrLf & vbTab  
			Response.Write "<td><a href=""mailto:" & arrayRS(2, i) & """>" & arrayRS(2, i) & "</a></td>" & vbCrLf & vbTab  
			Response.Write "<td>" & arrayRS(3, i) & "</td>" & vbCrLf & vbTab  
			Response.Write "<td>" & arrayRS(4, i) & "</td>" & vbCrLf 
			Response.Write "</tr>" & vbCrLf 
		next
		%>
		<tr>
			<td colspan=6>
			<a href="javascript:SetChecked(1,'id_mail','form_2')"><font face="Arial, Helvetica, sans-serif" size="0">Check All</font></a>
			&nbsp;&nbsp;
			<a href="javascript:SetChecked(0,'id_mail','form_2')"><font face="Arial, Helvetica, sans-serif" size="0">Clear All</font></a>
			</td>
		</tr>
		<tr>
			<td colspan=6><br></td>
		</tr>
		<tr bgcolor=goldenrod>
			<th colspan=7>Seleciona los contactos en caso de error :</th>
		</tr>	
		<tr><td colspan=7><br></td></tr>
		<tr bgcolor="goldenrod" align=center>
			<td colspan=2>.</td>
			<td>Nombre</td>
			<td colspan=3>Correo</td>
		</tr>
		<%
		for i = 0 to UBound(arrayRS2,2)
			Response.Write "<tr"
			if i mod 3 = 0 then Response.Write  " bgcolor=""FFFFEE"""
			Response.Write ">" & vbCrLf & vbTab 
			Response.Write "<td><input type=checkbox name=id_mail_error value="& arrayRS2(0, i) &"></td>"
			Response.Write "<td>" & i+1 & "</td>" & vbCrLf & vbTab 
			Response.Write "<td>" & arrayRS2(1, i) & "</td>" & vbCrLf & vbTab  
			Response.Write "<td colspan=3><a href=""mailto:" & arrayRS2(2, i) & """>" & arrayRS2(2, i) & "</a></td>" & vbCrLf & vbTab  
			Response.Write "</tr>" & vbCrLf 
		next
		%>
		<tr>
			<td colspan=6>
			<a href="javascript:SetChecked(1,'id_mail_error','form_2')"><font face="Arial, Helvetica, sans-serif" size="0">Check All</font></a>
			&nbsp;&nbsp;
			<a href="javascript:SetChecked(0,'id_mail_error','form_2')"><font face="Arial, Helvetica, sans-serif" size="0">Clear All</font></a>
			</td>
		</tr>
		<tr>
			<td align=left colspan=6>
			<!--<br>
			Nombre de la lista de contactos :<br>
			<input type=text name=list_name>-->
			<br><br>
			<input type="hidden" name=etape value=2>
			<input type="hidden" name=cli_num value=<%=Request.Form("cli_num")%>>
			<input type=submit class=buttonsOrange value=Validar><br><br>
			</td>
		</tr>
		
		</form>
		</table>
</body>
		</html>
		<%

case "2"
'almanecer los contactos en una tabla porque pueden ser numerosos
'Session("id_mail") = split(Request.Form("id_mail"),",")
Session("id_mail_error") = split(Request.Form("id_mail_error"),",")
'Session("id_lista") = split(Request.Form("id_lista"),",")
Session("id_mail") = Request.Form("id_mail")
'agregar los contactos de las listas a los contactos generales
if Request.Form("id_lista") <> "" then
	SQL = " select id_dest_lista from REP_LISTA_MAIL " & _
		  " where id_lista in ("&Request.Form("id_lista")&") "
			
	arrayRS = GetArrayRS(SQL)
	for i = 0 to ubound(arrayRS,2)
		Session("id_mail") = Session("id_mail") & "," & arrayRS(0,i)
	next
end if



		%>
<html>
<head>
<title>Agregar nuevo reporte - Etapa 3<%=Request.Form("id_lista")%></title>
</head>
<body>
<%call print_style()%>
<table border=0 width=350>
<form name="valid_conf" action="<%=asp_self()%>" method="post">
<tr bgcolor=goldenrod>
	<th>Seleciona un reporte :</th>
</tr>
		<%
		'Response.Write Request.Form("id_mail") 
		dim tab_tmp
		' <- CHG-DESA-30062021-01
		dim arr1(3), arr2(9)
		i=0

			'14,173,24
			arr1(0) = "ALEJANDROLE"
			arr1(1) = "JAVIERD"
			arr1(2) = "ALMALFS"

			'174
			arr2(0) = "YAZMINCC"
			arr2(1) = "EVELINGB"
			arr2(2) = "ELIZABETHBM"
			arr2(3) = "LUISFR"
			arr2(4) = "DULCELO"
			arr2(5) = "ALOURDESC"
			arr2(6) = "ALEXANDRAMM"
			arr2(7) = "LGABRIELAM"
			arr2(8) = "MLOURDESB"

		SQL = " select rep.id_rep, rep.name, rep.num_of_param " & _
			  " from rep_reporte rep where 1=1 "
			
		for i=0 to UBound(arr1)
			if Session("array_user")(0,0) = arr1(i) then
				SQL = SQL + " and rep.id_rep in (14,173,24) "
				exit for
			end if
		next

		for i=0 to UBound(arr2)
			if Session("array_user")(0,0) = arr2(i) then
				SQL = SQL + " and rep.id_rep = 174 "
				exit for
			end if
		next
	
		SQL = SQL + " order by 2"
		' CHG-DESA-30062021-01 ->

		arrayRS = GetArrayRS(SQL)
		
		Response.Write "<tr><td><select name=id_rep class=light>"
		for i=0 to UBound(arrayRS,2)
			Response.Write "<option value="&arrayRS(0,i)&">" & arrayRS(1,i) & vbCrLf  & vbTab 

		next
			Response.Write "</td>" & vbCrLf 
			Response.Write "</tr>" & vbCrLf
		%>
<tr>
	<td><br></td>
</tr>
<tr>
	<td>Cliente :<br><font color=red> <%=Request.Form("cli_num")%></font>
	<%SQL = "select clinom from eclient where cliclef='"& Request.Form("cli_num") &"'"
	arrayRS = GetArrayRS(SQL)
	Response.Write " : " & arrayRS(0,0)
	%>
	</td>

</tr>
<tr>
	<td><br></td>
</tr>
<tr bgcolor=goldenrod>
	<th>Seleciona una periodicidad :</th>
</tr>
<tr>
	<td>
	Cada...
	<br>
	<script type="text/javascript">
	<!--
		function display_conf(num)
		{if (num == 1) 
			{//alert('ok');
			document.all["dia_semana_l"].style.visibility = 'hidden';
			document.all["dia_mes_l"].style.visibility = 'hidden';
			document.all["dia_mismo_l"].style.visibility = 'visible';
			document.all["dia_quincena_l"].style.visibility = 'hidden';
			//alert  (document.valid_conf.mismo_dia.length);
			if (document.valid_conf.mismo_dia[1].checked ) 
				{document.all["dia_laboral_l"].style.visibility = 'visible';}
			else
				{document.all["dia_laboral_l"].style.visibility = 'hidden';}
			}
		 else if (num == 2)
			{document.all["dia_semana_l"].style.visibility = 'visible';
			 document.all["dia_mes_l"].style.visibility = 'hidden';
			 document.all["dia_mismo_l"].style.visibility = 'hidden';
			 document.all["dia_quincena_l"].style.visibility = 'hidden';
			 document.all["dia_laboral_l"].style.visibility = 'hidden';
			}
		 else if (num == 3)
			{document.all["dia_semana_l"].style.visibility = 'hidden';
			 document.all["dia_mismo_l"].style.visibility = 'hidden';
			 document.all["dia_mes_l"].style.visibility = 'hidden';
			 document.all["dia_quincena_l"].style.visibility = 'visible';
			 document.all["dia_laboral_l"].style.visibility = 'hidden';
			    if (document.valid_conf.mismo_dia[1].checked ) 
			    	{document.all["quincena_curso_l"].style.visibility = 'visible';}
			    else
			    	{document.all["quincena_curso_l"].style.visibility = 'hidden';}
			}
		 
		 else if (num == 4)
			{document.all["dia_semana_l"].style.visibility = 'hidden';
			 document.all["dia_mes_l"].style.visibility = 'visible';
			 document.all["dia_mismo_l"].style.visibility = 'hidden';
			 document.all["dia_quincena_l"].style.visibility = 'hidden';
			 document.all["dia_laboral_l"].style.visibility = 'hidden';
			}
		}
	function show_laboral(view)
		{//alert (view);
		 if (view == 1) 
		   {document.all["dia_laboral_l"].style.visibility = 'visible';
		    //alert (document.all["dia_laboral_l"].style.visibility);
		    }
		 else 
		   {document.all["dia_laboral_l"].style.visibility = 'hidden';}
		return true;
		}
	// -->
	</script>
	<input type=radio name=periodicidad value=1 onclick="(1);" checked>&nbsp;dia
	<input type=radio name=periodicidad value=2 onclick="display_conf(2);">&nbsp;semana
	<input type=radio name=periodicidad value=3 onclick="display_conf(3);">&nbsp;quincena
	<input type=radio name=periodicidad value=4 onclick="display_conf(4);">&nbsp;mes
	</td>
</tr>
<tr>
	<td><br></td>
</tr>
<tr bgcolor=goldenrod>
	<th>Seleciona una fecha de confirmacion :</th>
</tr>
<tr>
	<td>
	<div><input type=checkbox name=con_conf value=1> Con confirmacion.<br><br></div>
	<!--div name="conf_l" id="conf_l"-->
	Fecha de envio (y de confirmacion si hay) :<br><br>
	<!--div name="hour_l" id="hour_l" style="POSITION: relative;"-->
	Hora : <select name=hora class=light>
	<%
	for i=0 to 23
		Response.Write vbTab & vbTab &"<option value=" & i 
		if i=10 then Response.Write " selected "
		Response.Write  ">" & i
	next
	%>		
	</select>
	&nbsp;&nbsp;
	Minutos : 	<select name=minutos class=light>	
	<%
	for i=0 to 55 step 5
		Response.Write vbTab & vbTab &"<option value=" & i & ">" & i
	next
	%>
	</select>		
	<div name="dia_semana_l" id="dia_semana_l" style="position: relative; visibility:hidden;">
	&nbsp;&nbsp;&nbsp;
	<select name=dia_semana class=light>
		<option value=1>Lunes
		<option value=2>Martes
		<option value=3>Miercoles
		<option value=4>Jueves
		<option value=5>Viernes
		<option value=6>Sabado
		<option value=7>Domingo
	</select>
	</div>
	<div name="dia_mismo_l" id="dia_mismo_l" style="position: relative; visibility:visible;">
	<input type=radio name=mismo_dia value=1 checked onfocus="javascript:show_laboral(0);"> Mismo Dia <input type=radio name=mismo_dia value=0 onfocus="javascript:show_laboral(1);"> Dia pasado
	<!--div name="dia_laboral_l" id="dia_laboral_l" style="position: relative; visibility:hidden;"-->
	<input type=checkbox name=laboral value=1> Tomar en cuenta los dias laborales.<br>
	<i>Ejemplo : si hay palomita, el dia pasado del lunes es el viernes.</i> 
	<!--/div-->
	</div>
	
	<div name="dia_mes_l" id="dia_mes_l" style="position: relative; left=0px; top=5px; visibility:hidden;">

	Fecha mes : 
	<select name=fecha_mes class=light>
		<%for i=1 to 30
			Response.Write "<option value=" & i & ">" & i
		next
		%>
	</select>
	</div>
	

	<div name="dia_quincena_l" id="dia_quincena_l" style="position: relative; left=-5px; top=-5px; visibility:hidden;">
	1<sup>a</sup>Fecha : 
	<select name=fecha15_1 class=light>
		<%for i=16 to 30
			Response.Write "<option value=" & i & ">" & i
		next
		%>
	</select>
	&nbsp;
	&nbsp;&nbsp;&nbsp;&nbsp;2<sup>nda</sup> Fecha : 
	<select name=fecha15_2 class=light>
		<%for i=1 to 15
			Response.Write "<option value=" & i & ">" & i
		next
		%>
	</select>
	<!--div name="quincena_curso_l" id="quincena_curso_l" style="position: relative; visibility:hidden;"-->
	<input type=checkbox name=quincena_curso value=1> Tomar la quincena en curso.<br>
	<!--/div-->
	</div>
	
	</td>
</tr>
<tr>
	<td>
<input type=hidden name=id_mail value="<%=Request.Form("id_mail")%>">
<input type=hidden name=id_lista value="<%=Request.Form("id_lista")%>">
<input type=hidden name=etape value="3">
<input type=hidden name=cli_num value="<%=Request.Form("cli_num")%>">
<input type=submit class=buttonsOrange value=Validar><br><br>

</form>
<br></td>
</tr></table>
</body>
</html>
<%
case "3"
%>
<html>
<head>
<title>Agregar nuevo reporte - Etapa 4</title>
</head>
<body>
<%call print_style()%>
<script language="javascript">
function ValidateForm()
{
	var msg = "";
    /*var patt = new RegExp(/$^[A-Za-z0-9\-\_\%]+/g);
	if (document.valid_conf.file_name.value == "")
		{msg="- el archivo no tiene nombre.\n"};
	if (document.valid_conf.report_name.value == "")
		{msg+="- el reporte no tiene nombre.\n"};	
	if (document.valid_conf.carpeta.value == "")
		{ msg += "- no hay nombre de carpeta.\n" };
	if (patt.test(document.valid_conf.report_name.value) == false)
		{ msg += "- el nombre del reporte no es v·lido.\n" };
	if (patt.test(document.valid_conf.file_name.value) == false)
		{ msg += "- el nombre del archivo no es v·lido.\n" };
	if (patt.test(document.valid_conf.carpeta.value) == false)
		{ msg += "- el nombre de la carpeta no es v·lido.\n" };*/
	
	if (document.valid_conf.file_name.value == "")
		{msg="- el archivo no tiene nombre.\n"};
	if (document.valid_conf.report_name.value == "")
		{msg+="- el reporte no tiene nombre.\n"};	
	if (document.valid_conf.carpeta.value == "")
		{msg+="- no hay nombre de carpeta."};
	var check
	check = check_opcion(document.valid_conf.param, document.valid_conf.opcion) 
	if (check != 0 && check != undefined)
		{msg+="- los siguientes parametros : "+check+" son necesarios.\n"};
	
	
   
	/*var dml = document.valid_conf
	var len = dml.elements.length;
			var i=0;
			var j=1;
			for( i=0 ; i<len ; i++) {
				if ((dml.elements[i].name=="param") && (dml.elements[i].value=="")) 
				{msg+="- falta parametro " + j;
				 j++;}
				}*/

	if (msg == "") return true;

	alert ("Verifica los datos : \n"+msg);
	return false;
}



   function Remplace(expr) {
      var new_name = expr.value;
      var Forbidden_char = "\\/:*?\"\'<>|;,.~ &Ò—";
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

function check_opcion(param, op) {
   var error = "";
		//alert (op.length);
		//alert (op.value);
	if (op != undefined) {
		if (op.length == undefined)
		//caso que solo hay un unico parametro entonces op, param no son arrays
		{//alert("ok");
			if ((op.value == 0) && (param.value == "")) { return "1"; }
			else { return 0; }
		}
		else {
			for (var i = 0; i < op.length; i++) {
				if ((op[i].value == 0) && (param[i].value == "")) {
					if (error != "") { error = error + "," + (i + 1); }
					else { error = error + (i + 1); }
				}
				if (error == "") { return 0; }
				else { return error; }
			}
		}
	}
}

</script>

<table border=0 width=350>
<form name="valid_conf" action="<%=asp_self()%>" method="post" onsubmit="return ValidateForm();">
<tr bgcolor=goldenrod>
	<th>Seleciona un nombre de reporte :</th>
</tr>
<tr>
	<td>
	Nombre del reporte :<br>
	<input type=text name=report_name class=light size=40 maxlength=100><br>
	<br>Nombre del archivo :<br>
	<input type=text name=file_name class=light size=40 onblur="Remplace(this.form.file_name);"  maxlength=100><br>
	- Puedes usar signos especiales :<br>
	&nbsp;&nbsp;&nbsp;%P -> el rango de fecha de los datos (Mar-01-2003_to_Mar-07-2003)<br>
	&nbsp;&nbsp;&nbsp;%p -> la fecha de los datos 
	<br>
	<br>
	</td>
</tr>
<tr bgcolor=goldenrod>
	<th>Captura los parametros :</th>
</tr>

<%
'SQL = " select num_of_param " & _
'	  " from rep_reporte " & _
'	  " where id_rep='"& Request.Form("id_rep") &"' " 
'
'arrayRS = GetArrayRS(SQL)
'dim num_param
'num_param = arrayRS(0,0)
'
'SQL = "select "
'for i = 1 to Cint(num_param)
'	SQL = SQL & "name_param_" & i
'	if i <>  Cint(num_param) then  SQL = SQL & ", "
'next

'SQL = SQL & " from rep_reporte " & _
'	  " where id_rep='"& Request.Form("id_rep") &"' " 
'Response.Write SQL
'arrayRS = GetArrayRS(SQL)	  

'for i=1 to Cint(num_param)
'	Response.Write "<tr>" & vbCrLf & vbTab	
'	Response.Write "<td>"&i&".&nbsp;&nbsp;" & arrayRS(i-1,0) & "&nbsp;:&nbsp;&nbsp;<input type=text name=param size=10 class=light></td>" & vbCrLf 
'	Response.Write "</tr>" & vbCrLf
'next


%>

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
			
			for i=0 to UBound(arrayRS,1) step 2
				Response.Write "<tr>" & vbCrLf & vbTab
				Response.Write "<td>"
				if arrayRS(i+1,0) = "1" then Response.Write "<i>"
				Response.Write (i/2) + 1  &".&nbsp;&nbsp;" & arrayRS(i,0) & "&nbsp;:&nbsp;&nbsp;"
				if arrayRS(i+1,0) = "1" then Response.Write "</i>"
				Response.Write "&nbsp;&nbsp;&nbsp;<input type=text name=param size=10 class=light><input type=hidden name=opcion value="& arrayRS(i+1,0) &" </td>" & vbCrLf 
				Response.Write "</tr>" & vbCrLf
			next
			Response.Write "<tr><td colspan=2><i>En italico, los parametros son opcional.</i></td></tr>" & vbCrLf
		end if
	end if
	%>


<tr>
	<td>
	<br>
	<%SQL = " select distinct carpeta from rep_detalle_reporte " & _
			" Where cliente= '"& Request.Form("cli_num") &"' and test = 0 " 
	dim carpeta
	arrayRS = GetArrayRS(SQL)
	if IsArray(arrayRS) then
		carpeta = arrayRS(0,0)
	end if
	%>
	Carpeta : <input type=text class=light name=carpeta value="<%=carpeta%>" onblur="Remplace(this.form.carpeta);" maxlength=30>
	<br>(no poner espacio en el nombre y eligir uno sencillo)<br>
	</td>
</tr>
<tr>
	<td><br>
	Tiempo que se va a quedar el reporte en el servidor : <br>
	<%
	dim dias_delete, fecha_1, fecha_2
	select case Request.Form("periodicidad")
		case "1"
			dias_delete = 30
			fecha_1 = ""
			fecha_2 = ""
		case "2"
			dias_delete = 30
			fecha_1 = Request.Form("dia_semana")
			fecha_2 = ""
		case "3"
			dias_delete = 30
			if Request.Form("quincena_curso") <> "1" then
			    fecha_1 = Request.Form("fecha15_1")
			    fecha_2 = Request.Form("fecha15_2")
			else
			    fecha_2 = Request.Form("fecha15_1")
			    fecha_1 = Request.Form("fecha15_2")
			end if
		case "4"
			dias_delete = 60
			fecha_1 = Request.Form("fecha_mes")
			fecha_2 = ""
	end select 
	Response.Write dias_delete
	%> 
	&nbsp;dias.
	</td>
</tr>

<tr>
	<td>
		<input type=hidden name="dias_delete" value="<%=dias_delete%>">
		<input type="hidden" name="periodicidad" value="<%=Request.Form("periodicidad")%>">
		<input type="hidden" name="fecha_1" value="<%=fecha_1%>">
		<input type="hidden" name="fecha_2" value="<%=fecha_2%>">
		<input type="hidden" name="hora" value="<%=Request.Form("hora")%>">
		<input type="hidden" name="minutos" value="<%=Request.Form("minutos")%>">
		<input type="hidden" name="id_mail" value="<%=Request.Form("id_mail")%>">
		<input type="hidden" name="id_lista" value="<%=Request.Form("id_lista")%>">
		<input type="hidden" name="cli_num" value="<%=Request.Form("cli_num")%>">
		<input type="hidden" name="id_rep" value="<%=Request.Form("id_rep")%>">
		<input type="hidden" name="con_conf" value="<%=Request.Form("con_conf")%>">	
		<input type="hidden" name="mismo_dia" value="<%=Request.Form("mismo_dia")%>">
		<input type="hidden" name="laboral" value="<%=Request.Form("laboral")%>">	
		<input type="hidden" name="quincena_curso" value="<%=Request.Form("quincena_curso")%>">	
		<input type="hidden" name="etape" value="4"><br>
		<input type=submit class="buttonsOrange" value="Validar"><br><br>
	</td>
</tr>

</form>
 
<br></table>
</body>
</html>
<%

case "4"
%>
<html>
<head>
<title>Agregar nuevo reporte - Etapa 5</title>
</head>
<script language="javascript">
<!--
function ver_lista(lista)
{
    // <- CHG-DESA-30062021-01
	window.open("ver_lista.asp?liste=" + lista + "&Num=" + <%=Request.Form("cli_num")%> , "Lista_contactos", "toolbar=no, location=no, directories=no, status=yes, scrollbars=yes, resizable=yes, copyhistory=no, width=500, height=400, left=300, top=50");
	// CHG-DESA-30062021-01 ->
}
//-->
</script> 
<body>
<%call print_style()%>
<table border=0 cellpadding=2 cellspacing=0 width=350>
<form name="valid_conf" action="<%=asp_self()%>" method="post">
<tr bgcolor=goldenrod>
	<th colspan=2>Verifica los datos :</th>
</tr>
<tr>
	<td>Cliente</td>
	<td><%=Request.Form("cli_num")%></td>
</tr>
<tr>
	<td>Tipo del reporte</td>
	<td>
	<%SQL = "Select name from rep_reporte " & _
			"Where id_rep ='"&Request.Form("id_rep")&"' " 
	arrayRS = GetArrayRS(SQL)
	Response.Write arrayRS(0,0)
	%></td>
</tr>
<tr>
	<td>Nombre del reporte</td>
	<td><%=TAGescape(Request.Form("report_name"))%></td>
</tr>
<tr>
	<td>Nombre del archivo</td>
	<td><%=TAGescape(Request.Form("file_name"))%></td>
</tr>
<tr>
	<td>Carpeta</td>
	<td><%=TAGescape(Request.Form("carpeta"))%></td>
</tr>
	<%dim params, opciones, param_HTML
	params = split(Request.Form("param"), ",")
	opciones = split(Request.Form("opcion"), ",")

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
	<td>Dias de disponibilidad</td>
	<td><%=Request.Form("dias_delete")%></td>
</tr>
<tr>
	<td>Peridodicidad</td>
	<td>Cada 
	<%select case Request.Form("periodicidad")
		case "1"
			Response.Write "dia"
		case "2"
			Response.Write "semana"
		case "3"
			Response.Write "quincena"
		case "4"
			Response.Write "mes"
	end select
	%></td>
</tr>
<tr>	
	<td>Contactos</td>
	<td><a href="javascript:void(0);" onclick="ver_lista('ok');">Ver la lista</a></td>
</tr>
<tr>	
	<td>Contactos en caso de error</td>
	<td><a href="javascript:void(0);" onclick="ver_lista('error');">Ver la lista</a></td>
</tr>
<tr>
	<td><br></td>
</tr>
<tr>
	<th colspan=2>Fecha de envio (y confirmacion)</th>
</tr>
<tr>
	<td>Confirmacion</td>
	<td><%
	select case Request.Form("con_conf")
		case "1"
			Response.Write "Si"
		case else 
			Response.Write "No"
	end select
	%></td>
</tr>
<tr>
	<td valign=top>Fecha</td>
	<td>
	<%dim frecuen
	frecuen = ""
	select case Request.Form("periodicidad")
		case "1"
			Response.Write "Cada dia, "
			if Request.Form("con_conf") = "1" then
				Response.Write "<br>Confirmacion el "
			end if
			if Request.Form("mismo_dia") = "1" then
				Response.Write "mismo dia."
				frecuen = 0
			else 
				Response.Write "dia pasado"
				if Request.Form("laboral") = "1" then 
					Response.Write " (laboral)="
					frecuen = 1
				else 
					frecuen = 5
				end if
				Response.Write "."
					
			end if
			'end if
		case "2"
			Response.Write "Cada " 
			select case Request.Form("fecha_1")
				case "1"
					Response.Write "lunes"
				case "2"
					Response.Write "martes"
				case "3"
					Response.Write "miercoles"
				case "4"
					Response.Write "jueves"
				case "5"
					Response.Write "viernes"
				case "6"
					Response.Write "sabado"
				case "7"
					Response.Write "domingo"
			end select
			'Response.Write "<br>&nbsp;"
		case "3"
			Response.Write Request.Form("fecha_1") & " y " & Request.Form("fecha_2") & " de cada mes"
            if Request.Form("quincena_curso") = "1" then 
				frecuen = 6
			else 
				frecuen = 3
			end if
		case "4"
			Response.Write Request.Form("fecha_1") & " de cada mes"
	end select
	if frecuen = "" then frecuen = Request.Form("periodicidad")
	%></td>
</tr>
<tr>
	<td>Hora</td>
	<td><%=Request.Form("hora")%>:<%if Request.Form("minutos") < 10 then 
		Response.Write "0" & Request.Form("minutos")
	  else
		Response.Write Request.Form("minutos")
	  end if
	%></td>
</tr>
<tr>
	<td colspan=2><br></td>
</tr>
<tr>
	<td colspan=2>
		<%if param_HTML = "" then%>
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
	<input type="hidden" name="param" value="<%=TAGescape(SQLEscape(Request.Form("param")))%>">
	<input type="hidden" name="dias_delete" value="<%=Request.Form("dias_delete")%>">
	<input type="hidden" name="periodicidad" value="<%=frecuen%>">
	<input type="hidden" name="fecha_1" value="<%=Request.Form("fecha_1")%>">
	<input type="hidden" name="fecha_2" value="<%=Request.Form("fecha_2")%>">
	<input type="hidden" name="hora" value="<%=Request.Form("hora")%>">
	<input type="hidden" name="minutos" value="<%=Request.Form("minutos")%>">
	<input type="hidden" name="id_mail" value="<%=Request.Form("id_mail")%>">
	<input type="hidden" name="id_lista" value="<%=Request.Form("id_lista")%>">
	<input type="hidden" name="cli_num" value="<%=Request.Form("cli_num")%>">
	<input type="hidden" name="id_rep" value="<%=Request.Form("id_rep")%>">
	<input type="hidden" name="con_conf" value="<%=Request.Form("con_conf")%>">	
	<input type="hidden" name="mismo_dia" value="<%=Request.Form("mismo_dia")%>">
	<input type="hidden" name="laboral" value="<%=Request.Form("laboral")%>">
	<input type="hidden" name="quincena_curso" value="<%=Request.Form("quincena_curso")%>">
	<input type="hidden" name="etape" value="5">
	<%if param_HTML = "" then%>
	<input type=submit value=Validar class=buttonsOrange >
	<input type=button value=Cancelar class=buttonsOrange onclick="javascript:location.href='menu.asp';">
	<%end if%>
	</td>
</tr>
</table>

<%
case "5"
	'insercion de los datos en la base...
	Dim id_mail_ok, id_mail_error, id_rep, rst, id_mail_list
	
	''''''''''''''''''''''''''''''
	'' insercion contactos mail ''
	''''''''''''''''''''''''''''''

	SQL = " select distinct id_mail " & _
		  " from rep_mail " & _
		  " where id_mail in ("& Session("id_mail")& ") "  & _
		  " and status = 1 " 
	id_mail_list = GetArrayRS(SQL)

	
	'creacion nueva clave
	SQL = "select SEQ_DEST_MAIL.nextval from dual"
	arrayRS = GetArrayRS(SQL)
	id_mail_ok = arrayRS(0,0)
	
	'Response.Write "mail ok : <br>"
	
	set rst = Server.CreateObject("ADODB.Recordset")
	
	for i= 0 to UBound(id_mail_list,2)
		SQL = " insert into rep_dest_mail (id_dest_mail, id_dest) " & _
			  " values ('"& id_mail_ok &"','"& id_mail_list(0,i) &"' ) "
		'Response.Write SQL
		Session("SQL") = SQL
		rst.Open SQL, Connect(), 0, 1, 1
	next
	
	'insercion de los contactos en caso de error
	'Session("id_mail_error")
	SQL = "select SEQ_DEST_MAIL.nextval from dual"
	arrayRS = GetArrayRS(SQL)
	id_mail_error = arrayRS(0,0)
	
	'Response.Write "<br>mail error : <br>"
	for i= 0 to UBound(Session("id_mail_error"))
		SQL = " insert into rep_dest_mail (id_dest_mail, id_dest) " & _
			  " values ('"& id_mail_error &"','"& Session("id_mail_error")(i) &"' ) "
		'Response.Write SQL & "<br>"
		Session("SQL") = SQL
		rst.Open SQL, Connect(), 0, 1, 1
	next

	
	'Insercion del detalle del reporte
	SQL = "select SEQ_REPORTE_DETALLE.nextval from dual"
	arrayRS = GetArrayRS(SQL)
	id_rep = arrayRS(0,0)
		
	
	
	SQL = " insert into rep_detalle_reporte (id_cron, id_rep, mail_ok, mail_error, name, " & _
		  " cliente, frecuencia, file_name, carpeta, days_deleted, confirmacion " 

	params = split(Request.Form("param"), ",")
	for i = 0 to UBound(params)
		SQL = SQL & ", param_" & i+1 
	next  
	
	
	SQL = SQL & " , created_by, date_created) "  & _
		" values ('" & id_rep &"', '" & Request.Form("id_rep") & "' " & _
		" , '" & id_mail_ok  & "' , '" & id_mail_error  & "', '"& Request.Form("report_name") &"' "  & _
		" , '"& Request.Form("cli_num") &"', '"& Request.Form("periodicidad") &"' " & _
		" , '"& Request.Form("file_name") &"', '"& Request.Form("carpeta") &"' " & _
		" , '"& Request.Form("dias_delete") &"', '"& Request.Form("con_conf") &"' "
		
	params = split(Request.Form("param"), ",")
	for i = 0 to UBound(params)
		SQL = SQL & ", '" & params(i) & "' "
	next
	SQL = SQL & ", '"&  Session("array_user")(0,0) &"', sysdate)"
	
	Session("SQL") = SQL
	'Response.Write SQL & "<br>SQL 2 :<br> " & SQL
	'Response.End 
	rst.Open SQL, Connect(), 0, 1, 1
	
	SQL = "insert into rep_chron (id_chron, id_rapport, minutes, heures, jours " & _
		  ", mois, jour_semaine, priorite) values (SEQ_CHRON.nextval, '"& id_rep &"' " & _
		  ", '"& Request.Form("minutos") &"', '"& Request.Form("hora") &"' " 
	
	select case Request.Form("periodicidad")
		case "0"
			SQL = SQL & ", '', '', '1-5', '5')"
		case "1"
			SQL = SQL & ", '', '', '1-5', '5')"
		case "2"
			SQL = SQL & ", '', '', '"& Request.Form("fecha_1") &"', '5')"
		case "3"
			SQL = SQL & ", '"& Request.Form("fecha_1") &","& Request.Form("fecha_2") &"', '', '', '5')"
		case "6"
			SQL = SQL & ", '"& Request.Form("fecha_1") &","& Request.Form("fecha_2") &"', '', '', '5')"
		case "4"
			SQL = SQL & ", '"& Request.Form("fecha_1") &"', '', '', '5')"
		case "5"
			SQL = SQL & ", '"& Request.Form("fecha_1") &"', '', '', '5')"
	end select
	
	'Response.Write SQL & "<br>SQL 2 :<br> " & SQL & "<br>" & Request.Form("periodicidad")
	Session("SQL") = SQL
	rst.Open SQL, Connect(), 0, 1, 1
	'Response.End 
	
	'insercion de los numeros de lista usados por este reporte (los guardamos para luego 
	'usar los en caso de actualisacion de una lista
	if Request.Form("id_lista") <> "" then
		for i= 0 to UBound(split(Request.Form("id_lista"), ","))
			SQL = " insert into REP_LISTA_REPORTE (ID_LISTA, ID_REPORTE) " & _
				  " values ('"& split(Request.Form("id_lista"), ",")(i) &"','"& id_rep &"' ) "

			'Response.Write SQL & "<br>"
			Session("SQL") = SQL
			rst.Open SQL, Connect(), 0, 1, 1
		next
	end if
	
	
	dim mail, mail_server, mail_footer, mail_fecha, mail_message
	set mail = Server.CreateObject( "JMail.Message" )

 
	mail_server = Get_IP("MAIL_SERVER")
	mail_footer = vbCrLf & vbCrLf & vbCrLf & _
              "*********************************************************" & vbCrLf & _
              "This is a message automatically generated, please contact " & vbCrLf & _
              Get_Mail("webmaster") & " for any question or to unsubscribe."


	select case Request.Form("periodicidad")
		case "2"
			select case Request.Form("fecha_1")
				case "1"
					mail_fecha = "monday"
				case "2"
					mail_fecha = "tuesday"
				case "3"
					mail_fecha = "wednesday"
				case "4"
					mail_fecha = "thursday"
				case "5"
					mail_fecha = "friday"
				case "6"
					mail_fecha = "saturday"
				case "7"
					mail_fecha = "sunday"
			end select
		case "3"
			mail_fecha = "two weeks"
		case "6"
			mail_fecha = "two weeks"
		case "4"
			mail_fecha = "month"
		case else
			mail_fecha = "day"
	end select
	
	
	'manda correo de bienvenido
	mail.ISOEncodeHeaders = False
	mail.Subject = "Welcome to Logis Report server (4)."
	mail.From = Get_Mail("web_reports")
	mail.FromName = Get_Mail("web_reports_name")
	'para debug, estoy en los contactos ;)
	mail.AddRecipientBCC Get_Mail("IT_1")

	SQL = " SELECT MAIL.NOMBRE, MAIL.MAIL " & _ 
		" , REPORTE.NAME " & _
		" FROM REP_DETALLE_REPORTE REP   " & _ 
		"     , REP_DEST_MAIL DEST  " & _ 
		"     , REP_MAIL MAIL  " & _ 
		"     , REP_REPORTE REPORTE  " & _ 
		" WHERE REP.ID_CRON ='"& id_rep &"'  " & _ 
		"     AND MAIL.ID_MAIL = DEST.ID_DEST  " & _ 
		"     AND DEST.ID_DEST_MAIL = REP.MAIL_OK " & _
		"	  and status = 1 " & _ 
		"	  AND REPORTE.ID_REP = REP.ID_REP "
	arrayRS = GetArrayRS(SQL)
	'Response.Write SQL
	
	
	if IsArray(arrayRS) then
		for i=0 to UBound(arrayRS,2)
			mail.AddRecipient arrayRS(1,i), arrayRS(0,i)
		next
		
		mail_message = "<FONT FACE=""Arial,Helvetica"" SIZE=""2""><b>Welcome</b><br><br>" & _
					"This message will be sent only once, to unsubscribe see at the end.<br><br>" & vbCrLf & _
					"Begining today, you will receive each " & mail_fecha & vbCrLf & _
					" this report : <br> &lt; " & arrayRS(2,0) & " &gt; as requested." & _
					"For any question, please contact the Webmaster." & vbCrLf  & _
					"<br><br>Regards<br><br>Logis Web Site.</FONT>"

		
		'body en HTML
		mail.HTMLBody = display_mail("http://" & Get_IP("web_1"), mail_message )

		mail.body = notag(Replace(mail_message, "<br>", vbCrLf )) & _
					 vbCrLf & vbCrLf & mail_footer
		
		mail.Send mail_server
		

	end if	

	

	if Request.Form("con_conf") = "1" then
		'manda correo de bienvenido a la lista de error
		mail.Clear 
		mail.ISOEncodeHeaders = False
		mail.Subject = "Welcome to Logis Report server. Error list."
		mail.From = Get_Mail("web_reports")
		mail.FromName = Get_Mail("web_reports_name")
		'para debug, estoy en los contactos ;)
		mail.AddRecipientBCC Get_Mail("IT_1")
		
    	SQL = " SELECT MAIL.NOMBRE, MAIL.MAIL " & _ 
			" , REPORTE.NAME " & _
			" FROM REP_DETALLE_REPORTE REP   " & _ 
			"     , REP_DEST_MAIL DEST  " & _ 
			"     , REP_MAIL MAIL  " & _ 
			"     , REP_REPORTE REPORTE  " & _ 
			" WHERE REP.ID_CRON ='"& id_rep &"'  " & _ 
			"     AND MAIL.ID_MAIL = DEST.ID_DEST  " & _ 
			"     AND DEST.ID_DEST_MAIL = REP.MAIL_ERROR " & _
			"	  and status = 1 " & _
			"	  AND REPORTE.ID_REP = REP.ID_REP "
		arrayRS = GetArrayRS(SQL)
	
		if IsArray(arrayRS) then
			for i=0 to UBound(arrayRS,2)
				mail.AddRecipient arrayRS(1,i), arrayRS(0,i)
			next
			
			mail_message = "<FONT FACE=""Arial,Helvetica"" SIZE=""2""><b>Welcome to error list.</b><br><br>" & _
					"This message will be sent only once, to unsubscribe see at the end.<br><br>" & vbCrLf & _
					"Begining today, the errors for missing confirmation for " & vbCrLf & _
					" this report : <br> &lt; " & arrayRS(2,0) & " &gt; as requested." & vbCrLf & _
					"<br>For any question, please contact the Webmaster." & vbCrLf  & _
					"<br><br>Regards<br><br>Logis Web Site.</FONT>"

		
		'body en HTML
		mail.HTMLBody = display_mail("http://" & Get_IP("web_1"), mail_message )

		mail.body = notag(Replace(mail_message, "<br>", vbCrLf )) & _
					 vbCrLf & vbCrLf & mail_footer

		mail.Send mail_server
    
    end if
    
    end if
    
	Response.Redirect "menu.asp?msg=" & Server.URLEncode ("Reporte agregado, un correo fue mandado a los destinatarios.")
end select

%>
</body>
</html>

<%@ Language=VBScript %>
<%Response.Expires = 0%>
<HTML>
<HEAD>
<title>Client Login</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<style type="text/css">
.login {
	font-size: 10pt;
	background-color: #ffffff;
	font-family: Arial, sans-serif;
	border: 1 #000000 solid;
}

.error {
	font-size: 12pt;
	background-color: #ffffff;
	font-family: Arial, sans-serif;
	font-weight: bold;
	color: #ff0000;
}

.rowHead {
	background-color: goldenrod;
	color: #000000;
	font-family: Arial, Serif;
	font-size: 12pt;
	font-weight: bold;
}

td{
font-family: verdana;
font-style: normal;
font-weight: bold;
font-size: .8em;
}
.buttonsOrange { background-color: goldenrod; color:#000000; font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; height: 19px; font-color: #000000 }
/* <- CHG-DESA-30062021-01 */
		.trHeader
		{
			background-color: #223F94;
			font-family: "Roboto",sans-serif;
			font-size: 12px;
			color:#FFFFFF;
			text-align: center;
		}
		.trHeader>td
		{
			font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
			font-style: normal;
			font-weight: bold;
			font-size: 12px;
		}
		.buttonsBlue
		{
			background-color: #223F94;
			color:#FFFFFF;
			font-family: "Roboto",sans-serif;
			font-size: 11px;
			font-weight: bold;
			height: 19px;
		}
/* CHG-DESA-30062021-01 -> */
</style>
</HEAD>
<BODY>
<%
Session.Abandon 
'kill previous session if exists

dim er, logoff
logoff = request("logoff")
er=REQUEST("err")
'IF er<>"" THEN
'   if er=1 then%>
      <!--<SCRIPT LANGUAGE="JavaScript">
      alert ("Error, ID not Found")
      </SCRIPT><%
'   ELSE%>
      <SCRIPT LANGUAGE="JavaScript">
      alert ("Error, Incorrect Password")
      </SCRIPT>--><%
'   END IF
'end if

if er<>"" then
	if er=1 then
		%><center><p class="error">Error, ID not Found</p></center><%
	else%>
		<center><p class="error">Error, Incorrect Password</p></center><%
	end if
end if
if logoff = 1 then 
	%><center><p class="error">Your session is over, please reconnect</p></center><%
	session("time_out") = 1
	Session.Abandon ()
end if
%>
<center>
<form action="checkLogin.asp" method="post" name=login>
<table width=100% cellspacing=0 cellpadding=5 border=0>
 <tr>
  <td valign=center align=center>
      <table width=350 cellspacing=5 cellpadding=3 border=0>
<!-- <- CHG-DESA-30062021-01 -->
	<tr>
		<td align="center" colspan="2">
			<img src="images/logo-logis-s.png" alt="Logo de Logis" style="width:300px;" />
		</td>
	</tr>
    <tr>
    <tr>
     <td colspan=2 valign=center align=center><br></td>
    </tr>
    <tr>
    
    <tr class="trHeader">
     <td colspan=2 valign=center align=center>Por favor ingrese su Usuario y Contrase&ntilde;a:</td>
    </tr>
    <tr>
     <td align=right>Usuario:</td>
     <td><input class="login" type="text" name="username" value=""></td>
    </tr>
    <script>
		document.login.username.focus();
	</script>    
    <tr>
     <td align=right>Contrase&ntilde;a:</td>
     <td><input class="login" type="password" name="password" value=""></td>
    </tr>
    <tr>
     <td colspan=2 align=right>
      <input type="submit" name="submit" value="Enviar" class="buttonsBlue">
     </td>
    </tr>
    <tr class="trHeader">
     <td colspan=2>&nbsp;</td>
    </tr>
<!-- CHG-DESA-30062021-01 -> -->
   </table>
  </td>
 </tr>
</table>
</form>
<br><br>
 <%=request.serverVariables("LOCAL_ADDR")%> 
<!--Disculpa por la molestia, estoy Haciendo una actualizacion.
En unos minutos estara lista la pantalla.
Nicolas-->
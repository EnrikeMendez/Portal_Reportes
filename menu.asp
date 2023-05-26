<%@ Language=VBScript %>
<% option explicit 
'admin of logis web site :
' access to site stats, check client passwords....
Response.Expires = 0
call check_session()
%>
<html>
<head>
	<!--#include file="include/include.asp"--><%

call print_style()

%>
<title>Administracion Logis Web Site</title>
	<!-- <- CHG-DESA-30062021-01 -->
	<style type="text/css">
		.trHeader
				{
					background-color: #223F94;
					font-family: "Roboto",sans-serif;
					font-size: 14px;
					color:#FFFFFF;
					text-align: center;
				}
		.trHeader>td
		{
			font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
			font-style: normal;
			font-weight: bold;
			font-size: 14px;
		}
		td{
			font-size: 12px;
		}
	</style>
	<!-- CHG-DESA-30062021-01 -> -->
</head>
<body>

<center>
<table width=350>
	<!-- <- CHG-DESA-30062021-01 -->
	<tr>
		<td align="center">
			<img src="images/logo-logis-s.png" alt="Logo de Logis" style="height:50px;" />
		</td>
	</tr>
	<!-- CHG-DESA-30062021-01 -> -->
<tr style="text-align: right;">
	<td><a href=login.asp?logoff=1>Log out</a><br><br></td>
</tr>
<%if Request("msg") <> "" then	Response.Write "<tr><td align=center colspan=2><font color=red size=2>" & Request("msg") & "</font></td></tr>"
		%>
<!-- <- CHG-DESA-30062021-01 -->
<tr class="trHeader">
	<td>
		Generaci&oacute;n de Reportes
	</td>
</tr>
<!-- CHG-DESA-30062021-01 -> -->
<tr>
	<td>
		<a href=confirmacion.asp>Confirmacion de reportes autom&aacute;ticos</a><br><br>
		<%dim i,j
		for i=0 to UBound(Session("array_user"),2)
			if Session("array_user")(1,i) <> "0" then
				Response.Write "<a href=""confirmacion2.asp"">Confirmacion de aduanas</a><br><br>"
				exit for
			end if
		next
		%>
		<%
		' <- CHG-DESA-30062021-01
		if Session("array_user")(0,0) = "HECTORRR" or Session("array_user")(0,0) = "CESARRP" or Session("array_user")(0,0)="ELYANGV" then
			if Session("array_user")(0,0)="NICOLAST" or Session("array_user")(0,0)="CHRISTELLE" or Session("array_user")(0,0)="ADMIN" or _ 
			Session("array_user")(0,0)="NAJIBS" or Session("array_user")(0,0)="NAJIBS" or Session("array_user")(0,0)="OLIVIERD" or _
			Session("array_user")(0,0)="CESARRP"  then
			%>
					<a href="javascript:select_reporte('7,8,11,17,18,14,13,6,24,36,185,202,184,192,260,315,327,108<%if Session("array_user")(4,0) = "1"then Response.Write ",188"%>');">Entrar (reporte puntual)</a><br><br>

			<%else%>
				<a href="javascript:select_reporte('7,8,11,17,18,14,13,6,24,36,185,202,184,192,315,327,108<%if Session("array_user")(4,0) = "1" then Response.Write ",188"%>');">Entrar (reporte puntual)</a><br><br>
			<%end if
		else
			if Session("array_user")(0,0)="NICOLAST" or Session("array_user")(0,0)="CHRISTELLE" or Session("array_user")(0,0)="ADMIN" or _ 
			Session("array_user")(0,0)="NAJIBS" or Session("array_user")(0,0)="NAJIBS" or Session("array_user")(0,0)="OLIVIERD" or _
			Session("array_user")(0,0)="CESARRP"  then
			%>
					<a href="javascript:select_reporte('7,8,11,17,18,14,13,6,24,36,185,202,184,192,260,315,327<%if Session("array_user")(4,0) = "1"then Response.Write ",188"%>');">Entrar (reporte puntual)</a><br><br>

			<%else%>
				<a href="javascript:select_reporte('7,8,11,17,18,14,13,6,24,36,185,202,184,192,315,327<%if Session("array_user")(4,0) = "1" then Response.Write ",188"%>');">Entrar (reporte puntual)</a><br><br>
			<%end if%>
		<% 
		end if
		%>
		
		<a  href=reporte_anomalia.asp>Reportes Anomal&iacute;as</a><br><br>
		<%
			if Session("array_user")(0,0) = "HECTORRR" or Session("array_user")(0,0) = "CESARRP" then
		%>
			<a  href="cambio_prioridad.asp">Reportes en progreso - Cambio de prioridades</a><br><br>
			<a  href="consulta_errores.asp">Consulta de errores</a><br><br>
		<%
			end if
		' CHG-DESA-30062021-01 ->
		%>
	</td>
</tr>
	<tr>
		<td class="bold" style="font-weight:bold;font-size:13pt;color:green;">
			<a  href="mail_pruebas_ws.asp">
				Prueba web Service
			</a>
		</td>
	</tr>
<!-- <- CHG-DESA-30062021-01 -->
<tr class="trHeader">
	<td>Configuraci&oacute;n de Reportes</td>
</tr>
<tr>
	
	<td>
	<a href=reporte.asp>Nuevo reporte</a><br><br>
		<a href=modif_reporte.asp>Modificar / Reprocesar un reporte</a><br><br>
		
		<%
			if Session("array_user")(0,0) = "HECTORRR" or Session("array_user")(0,0) = "CESARRP" then
		%>
			<a href="reproceso_general.asp">Reproceso de Pedimentos inst&aacute;ntaneos</a><br><br>
			<a href=lista_usuarios.asp>Ver los reportes por usuario</a><br><br>
			<a href=monitoreo_reportes.asp>Monitoreo de reportes</a><br><br>
		<%end if%>
	</td>
</tr>

<%  if Session("array_user")(0,0) = "HECTORRR" or Session("array_user")(0,0) = "CESARRP" then	%>
	<tr class="trHeader">
		<td>Contactos</td>
	</tr>
	<tr>
		<td>
			<a href="mail.asp?et=-1">Crear contacto</a><br><br>
			<a href=lista_contactos.asp>Crear lista de contactos</a><br><br>
			<a href=lista_contactos.asp?accion=mod>Modificar una lista de contactos</a><br><br>
		</td>
	</tr>
<%end if%>
<!-- CHG-DESA-30062021-01 -> -->

	
	

<tr>
	<td><!--modif lista de contacto<a href=reporte_puntual.asp>Entrar (reporte puntual)</a><br><br>--></td>
</tr>
<!-- <- CHG-DESA-30062021-01 -->
<tr class="trHeader">
	<td>Configuraciones</td>
</tr>
<!-- CHG-DESA-30062021-01 -> -->
<tr>
	<td>
		<a href=dias.asp>Dias libres de reportes</a><br><br>
		<a href="param_ltl.asp">Par&aacute;metros LTL/CrossDock</a><br><br>
		
		

	</td>
</tr>
<%if Session("array_user")(3,0) = "1" then%>
<!-- <- CHG-DESA-30062021-01 -->
<tr class="trHeader">
	<td>Deactivacion anomalias :</td>
</tr>
<!-- CHG-DESA-30062021-01 -> -->
<tr>
	<td><a href=desactivacion.asp>Entrar</a><br><br></td>
</tr>
<%end if%>
<%if Session("array_user")(4,0) = "1" then%>
<!-- <- CHG-DESA-30062021-01 -->
<tr class="trHeader">
	<td>Confrmacion Estado de resultados :</td>
</tr>
<!-- CHG-DESA-30062021-01 -> -->
<tr>
	<td><a href="javascript:select_reporte('29,30,31');">Entrar</a><br><br></td>
</tr>
<%end if%>
</table>

<!-- << CHG-DESA-06072022-01 -->
<% if Session("array_user")(0,0) = "HECTORRR" then %>
	<table width="350">
		<tr class="trHeader">
			<td>
				Prevalidaci&oacute;n de formatos
			</td>
		</tr>
		<tr>
			<td>
				<a href="formato_carga.asp">Generar formato de carga</a>
			</td>
		</tr>
	</table>
<% end if %>
<!--    CHG-DESA-06072022-01 >> -->
<script language="javascript">
function select_reporte(num) {
	document.reporte.reporte_num.value = num;
	document.reporte.submit() ;
}
</script>
<form name="reporte" action="reporte_puntual.asp" method="post">
	<input type="hidden" name="reporte_num" value="">
</form>
</body>
</html>
<%
'este archivo se encuentra a fuera de la raiz del sitio 
'por razones evidentes de seguridad
dim errorVar
function Get_Conn_string(what)
'return important string of connexion
'LOGIN
'PASS
'SERVER
	what = UCASE(what)
	if what = "LOGIN" then
		Get_Conn_string = "web_cliente"
	elseif what = "PASS" then
		Get_Conn_string = "rwJ8XnFvM"
	elseif what = "LOGIN_ADM" then
		Get_Conn_string = "web_adm"
	elseif what = "PASS_ADM" then
		Get_Conn_string = "ciPn4aFGf"
	elseif what = "LOGIN_PORTEO" then
		Get_Conn_string = "web_porteo"
	elseif what = "PASS_PORTEO" then
		Get_Conn_string = "xme3XB7Vi"
	elseif what = "SERVER" then
		Get_Conn_string = "logis"
	elseif what = "SERVER_ADM" then
		Get_Conn_string = "LOGIS"
	elseif what = "CLON" then
		Get_Conn_string = "clon"
	elseif what = "LOGIN_ADM_SRV" then
		Get_Conn_string = "logis"
	elseif what = "PASS_ADM_SRV" then
		Get_Conn_string = "x41ygjRT6"
	elseif what = "SERVER10" then
		Get_Conn_string = "logis10"
	else
		errorVar = "Get_Conn_string(" & what & ")"
	end if 
end function

function Get_IP(what)
'devuelve direcciones IP de servidores que pueden cambiar
	what = UCase(what)
	if what = "MAIL_SERVER" then
		Get_IP = "192.168.100.6"
	elseif what = "WEB_1" then
		Get_IP = "192.168.100.4"
	elseif what = "WEB_2" then
		Get_IP = "192.168.100.5"
	else
		errorVar = "Get_IP(" & what & ")"
	end if
end function

function Get_Path(what)
'devuelve direcciones IP de servidores que pueden cambiar
	what = UCase(what)
	if what = "EXCEL_TEMP" then
		Get_Path = "D:\Logis_Web\wwwroot\@_interno\test2\excel\"
	
	elseif what = "EXCEL_TEMP_V2" then
		Get_Path = "D:\Logis_Web\wwwroot\@_interno\v2\excel\"	
	
	elseif what = "EXCEL_SIMPLE" then
		Get_Path = "excel/"
	
	elseif what = "EXCEL_GASTOS" then
		Get_Path = "D:\Logis_Web\wwwroot\@_interno\gastos\excel\"	
			
	elseif what = "EXCEL_VENTA" then
		Get_Path = "D:\Logis_Web\wwwroot\@_interno\venta\excel\"
	
	elseif what = "PDF_CACHE" then
		Get_Path = "D:\Logis_WebReports\cache\"
	
	elseif what = "PHOTOS_TEMP" then
		Get_Path = "/photos/temp/"
		
	elseif what = "CODIGOS_BARRA_PEDIM" then
		Get_Path = "D:\web_photos\CodigosBarra\pedim\"
	
	else
		errorVar = "Get_Path(" & what & ")"

	end if
end function

function Get_Mail(what)
'devuelve correo electronicos de contacto
	what = UCase(what)
	if what = "WEBMASTER" then
		Get_Mail = "web-master@logis.com.mx"
	
	elseif what = "IT_1" then
		Get_Mail = "monitoreo_web@logis.com.mx"
	
	elseif what = "IT_2" then
		Get_Mail = "sylvaind@logis.com.mx"		
	
	elseif what = "IT_3" then
		Get_Mail = "monitoreo_web@logis.com.mx"
	
	elseif what = "IT_BOSS" then
		Get_Mail = "monitoreo_web@logis.com.mx"
	
	elseif what = "IT_ALL" then
		Get_Mail = "monitoreo_web@logis.com.mx"
		
	elseif what = "LTL_BOSS" then
		Get_Mail = "danielae@logis.com.mx"
	
	elseif what = "LOGIS_BOSS" then
		Get_Mail = "olivierd@logis.com.mx"

	elseif what = "LOGIS_SALES_BOSS" then
		Get_Mail = "javierd@logis.com.mx"
		
	elseif what = "LOGIS_SALES_TRADING" then
		Get_Mail = "myrnalf@logis.com.mx,monitoreo_web@logis.com.mx"
		
	elseif what = "LOGIS_OCP" then
		'Get_Mail = "laurap@logis.com.mx, gabrielacc@logis.com.mx"
		Get_Mail = "elyangv@logis.com.mx, myrnalf@logis.com.mx, olivierd@logis.com.mx"
		
	
	elseif what = "LOGIS_MX" then
		Get_Mail = "logis.com.mx"
	
	elseif what = "WEB_REPORTS" then
		Get_Mail = "web-reports@logis.com.mx"

	elseif what = "WEB_REPORTS_NAME" then
		Get_Mail = "Logis report server"	
	else
		errorVar = "Get_Mail(" & what & ")"
	end if
	'if what = "LOGIN" then
end function

if errorVar <> "" then
	set jmail = Server.CreateObject( "JMail.Speedmailer" )
	jmail.SendMail Get_mail("webmaster"), Get_mail("IT_1"), "errorVar definicion variable", "La funcion : " & errorVar & " no existe, verificar el codigo fuente.", Get_IP("mail_server")
end if
%>

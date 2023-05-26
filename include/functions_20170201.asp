<%
function Connect()
	dim strCon, obj_conn 
	set obj_conn=Server.CreateObject("ADODB.connection")
	Dim CONN_STRING, CONN_USER, CONN_PASS	
	CONN_STRING = Get_Conn_string("server")
	CONN_USER = Get_Conn_string("login_adm")
	CONN_PASS = Get_Conn_string("pass_adm")
	obj_conn.ConnectionTimeout = 1000	'timeout for connection
	obj_conn.CommandTimeout = 1000		' timeout for SQL commands
	obj_conn.Open CONN_STRING, CONN_USER, CONN_PASS	
	Connect = obj_conn
end function
%>


<%
function GetArrayRS (strSQL)
'return an array from a query
	dim strCon, obj_conn 
	set obj_conn=Server.CreateObject("ADODB.connection")
	Dim CONN_STRING, CONN_USER, CONN_PASS	
	CONN_STRING = Get_Conn_string("server")
	CONN_USER = Get_Conn_string("login_adm")
	CONN_PASS = Get_Conn_string("pass_adm")
	obj_conn.ConnectionTimeout = 1000	'timeout for connection
	obj_conn.CommandTimeout = 1000		' timeout for SQL commands
	obj_conn.Open CONN_STRING, CONN_USER, CONN_PASS	
	
	Dim rst
	set rst = Server.CreateObject("ADODB.Recordset")
	rst.Open strSQL, obj_conn, 0, 1, 1 'cursortype: forwardonly
	
	if not (rst.EOF) then
		GetArrayRS = rst.GetRows 
	else GetArrayRS = ""
	end if
	'clean
	set rst = nothing
	obj_conn.Close 
	set obj_conn = nothing
	'response.end
end function

function NBSP(count)
	dim i
	for i = 1 to count
		NBSP = NBSP + "&nbsp;"
	next
end function

'idem que BuildNav, pero con el uso de los formularios
'asi cada clic va modificar la variable PageNum del formulario
'pasar el nombre de la funcion Javascript para modificar el numero de pagina
Sub BuildNav2 (PageNum, Offset, TotalRecord, JS_func )
Dim TotalP
TotalP = int ((TotalRecord-1) \ (Offset)) +1 
Response.Write "<br>" 
	Dim counter, counterEnd
	Response.Write "<font face=verdana size=1><b>"
	
	' If the previous page id is not 0, then let's add the 'prev' link.
	if PageNum  > 1 then
		if PageNum > Offset then	
			With Response
				.Write  "<a href=""javascript:"& JS_func &"('"& PageNum - Offset &"');""><< prev " & Offset & "</a>"
				.Write  NBSP(5)
			End With
		end if
		With Response
			.Write "<a href=""javascript:"& JS_func &"('"& PageNum - 1 &"');"">< prev</a>"
			.Write  NBSP(5)
		End With		
	end if

	'This section displays the list of page numbers as links, separated with the pipe ( | )
	counter = PageNum
	if PageNum + Offset > TotalP then 
		counterEnd = TotalP - PageNum
	else counterEnd = counter + Offset -1
	end if
		
	Response.Write  "<br>jump to page: <br>"
	Dim I, C
	C=0
	I = (INT((PageNum-1)/Offset) * Offset)+1 
	
	do while I <> TotalP+1 AND C<>Offset
		C=C+1
		If I = PageNum Then
			Response.Write " " & I & " "
		Else
			Response.Write  "<a href=""javascript:"& JS_func &"('"& I &"');"">" & I & "</a>"
		End If
		
		if C <> Offset AND I <> TotalP then
			Response.Write  " | "
		end if
		I = I + 1
		
	loop

	if (PageNum < TotalP ) then
		With Response
			.Write  NBSP(5)
			.Write  "<br><a href=""javascript:"& JS_func &"('"& PageNum + 1 &"');"">next ></a>"
		End With
		if (PageNum + Offset <= TotalP  ) then
			With Response
				.Write  NBSP(5)
				.Write  "<br><a href=""javascript:"& JS_func &"('"& PageNum + Offset &"');"">next " & Offset & " >></a>"
			End With
		end if
	end if
	Response.Write "<br><br>" & TotalP  & " page"
	if TotalP >1 then
		Response.Write "s"
	end if
	Response.Write " found ( " & TotalRecord  &" record"
	if TotalRecord >1 then
		Response.Write "s"
	end if
	Response.Write " ) </b></font>"
end Sub
	
%>



<%
sub check_session()
'check if session is still alive, if not send to login page
	if not IsArray(Session("array_user")) then
		response.redirect ("login.asp?logoff=1")
	end if
end sub

function asp_self()
'send the name of the current script like Php_self do !
	dim long_URL, array_tmp
	long_URL  = Request.ServerVariables("SCRIPT_NAME") 
	array_tmp = Split (CStr(long_URL), "/",-1,1 )
	asp_self = array_tmp (UBound(array_tmp)) 
	set array_tmp = nothing
end function

%>
<%
Function notag(txt) 'Quita el HTML de txt.
   dim regEx
   Set regEx = New RegExp
   regEx.Global = True
   ' ^ significa todo menos
   ' + significa 1, porque a lo menos hay 1 "<"
   regEx.Pattern = "<[^>]+>"
   'lo cambia por nada
   notag = regEx.replace(txt,"")
end function
%>
<%
function verif_parametros(params, opciones, id_reporte)
	'verifica los parametros de los reportes segun la tabla REP_PARAMETROS_DESCRIPCION y REP_PARAMETROS
	dim SQL, arrayRS, i, arrayRS2
	on error resume next
	'porque en caso que el cliente sea un varchar, nos da un error SQL
	'lo comprobamos con  or not IsNumeric(params(i))...
	for i = 0 to UBound(params)
		if opciones(i) = "0" or Trim(params(i)) <> "" then
		SQL = "select rep.id_rep, rep.name, par.NUM_PARAM, prd.ID_PARAM, prd.DESCRIPCION_PARAM " & VbCrLf 
		SQL = SQL & " from rep_reporte rep " & VbCrLf 
		SQL = SQL & " 	 , REP_PARAMETROS par " & VbCrLf 
		SQL = SQL & " 	 , REP_PARAMETROS_DESCRIPCION prd " & VbCrLf 
		SQL = SQL & " where rep.id_rep = " & id_reporte & VbCrLf 
		SQL = SQL & " and par.NUM_PARAM = " & i+1 & VbCrLf 
		SQL = SQL & " and rep.id_rep = par.id_rep " & VbCrLf 
		SQL = SQL & " and par.ID_DESCR_PARAM = prd.ID_PARAM " & VbCrLf 
		SQL = SQL & " order by par.NUM_PARAM "
	'Response.Write opciones(i) & i & "<br>" & SQL
	arrayRS = GetArrayRS(SQL)
	
		if IsArray(arrayRS) then
		'hay parametros que verificar, vamos a checarlos:
			select case arrayRS(3,0)
				case "1"
					'aduana
					SQL = "select 1 from edouane where douclef = '" & Trim(SQLEscape(params(i))) & "'"
					'Response.Write SQL & "<br>"
					arrayRS2 = GetArrayRS(SQL)
					if not IsArray(arrayRS2) then
						if verif_parametros <> "" then verif_parametros = verif_parametros & "|"
						verif_parametros = verif_parametros & "La aduana no existe : " & params(i)
					end if
					set arrayRS2 = nothing
					
				case "2"
					'cliente
					SQL = "select 1 from eclient where cliclef = '" & Trim(SQLEscape(params(i))) & "'"
					'Response.Write SQL & "<br>"
					arrayRS2 = GetArrayRS(SQL)
					if not IsArray(arrayRS2) or not IsNumeric(params(i)) then
						if verif_parametros <> "" then verif_parametros = verif_parametros & "|"
						verif_parametros = verif_parametros & "El cliente no es valido : " & params(i)
					end if	
					set arrayRS2 = nothing
								
				case "3"
					'empresa de logis
					SQL = "select 1 from eempresas where empclave = '" & Trim(SQLEscape(params(i))) & "'"
					'Response.Write SQL & "<br>"
					arrayRS2 = GetArrayRS(SQL)
					if not IsArray(arrayRS2) or not IsNumeric(params(i)) then
						if verif_parametros <> "" then verif_parametros = verif_parametros & "|"
						verif_parametros = verif_parametros & "La empresas Logis no esta correcta : " & params(i)
					end if	
					set arrayRS2 = nothing
									
				case "4"
					'grupo de cotizacion
					SQL = "select 1 from egroupecotiz where grcgrupo = '" & Trim(SQLEscape(params(i))) & "'"
					'Response.Write SQL & "<br>"
					arrayRS2 = GetArrayRS(SQL)
					if not IsArray(arrayRS2) then
						if verif_parametros <> "" then verif_parametros = verif_parametros & "|"
						verif_parametros = verif_parametros & "La empresas Logis no esta correcta : " & params(i)
					end if	
					set arrayRS2 = nothing				

				case "5"
					'sucursal o lista de sucursal
					Dim j, lista_suc
'					For j = 0 To UBound(Split(params(i), ","))
'					    lista_suc = lista_suc & "'" & Trim(Split(params(i), ",")(j)) & "'"
'					    If j <> UBound(Split(params(i), ",")) Then lista_suc = lista_suc & ","
'					Next
					lista_suc = params(i)
					for j = 0 to UBound(Split(lista_suc, ","))
						SQL = "select 1 from esuccursale where sucnumerocontable = '" & Trim(SQLEscape(Split(lista_suc, ",")(j))) & "'"
						'Response.Write "SQL : " & SQL & "<br>"
						arrayRS2 = GetArrayRS(SQL)
						if not IsArray(arrayRS2) then
							if verif_parametros <> "" then verif_parametros = verif_parametros & "|"
							verif_parametros = verif_parametros & "El numero de sucursal no esta correcto : " & Split(lista_suc, ",")(j)
						end if	
					next
					set arrayRS2 = nothing

				case "6"
					'area o lista de areas contables
					Dim lista_area
					lista_area = params(i)
					for j = 0 to UBound(Split(lista_area, ","))
						SQL = "select 1 from edepartamentos where substr(depnumerocontable, 1,2)  = '" & Trim(SQLEscape(Split(lista_area, ",")(j))) & "'"
						'Response.Write SQL & "<br>"
						arrayRS2 = GetArrayRS(SQL)
						if not IsArray(arrayRS2) then
							if verif_parametros <> "" then verif_parametros = verif_parametros & "|"
							verif_parametros = verif_parametros & "El numero de area no esta correcto : " & Split(lista_area, ",")(j)
						end if	
					next	
				set arrayRS2 = nothing		

				case "7"
					'CEDIS
					Dim lista_cedis
					lista_cedis = params(i)
					for j = 0 to UBound(Split(lista_cedis, ","))
						SQL = " select 1 from EALMACENES_LOGIS where ALLCLAVE = '" & Trim(SQLEscape(Split(lista_cedis, ",")(j))) & "'"
						'Response.Write SQL & "<br>"
						arrayRS2 = GetArrayRS(SQL)
						if not IsArray(arrayRS2) then
							if verif_parametros <> "" then verif_parametros = verif_parametros & "|"
							verif_parametros = verif_parametros & "El numero de CEDIS no esta correcto : " & Split(lista_cedis, ",")(j)
						end if	
					next	
				set arrayRS2 = nothing		
			end select
		end if 
	end if
	next
	on error goto 0
end function


function check_mail(correo)
	'verifica si los correo son de logis y si asi es, si son validos...
	dim lista_correo, i, SQL, arrayRS
	if right(correo,1)=";" or right(correo,1)="," then
		correo=Mid(correo, 1, Len(correo) - 1)
	end if
	lista_correo = Split(Replace(correo, ";", ","), ",")
	for i = 0 to UBound(lista_correo)
		if UBound(Split(lista_correo(i), "@")) > 0 then
		   if LCase(Trim(Split(lista_correo(i), "@")(1))) = Get_Mail("logis_mx") then
			 if LCase(Trim(Split(lista_correo(i), "@")(0))) <> "christelle" and LCase(Trim(Split(lista_correo(i), "@")(0))) <> "guillehe" _
			    and LCase(Trim(Split(lista_correo(i), "@")(0))) <> "jessica" and LCase(Trim(Split(lista_correo(i), "@")(0))) <> "alejandrol"  _
			    and LCase(Trim(Split(lista_correo(i), "@")(0))) <> "crossdocktln" and LCase(Trim(Split(lista_correo(i), "@")(0))) <> "teresam"     and LCase(Trim(Split(lista_correo(i), "@")(0))) <> "erikapm" and  LCase(Trim(Split(lista_correo(i), "@")(0))) <> "cedischihuahua"	then
				'excepcion para Christelle porque no tiene su correo en esta tabla...
				'tambien por guillermina Hernandez... que buena tabla.... :))
				'y Jessica Gonzalez; cedischihuahua
				SQL = "select 1 from usuarios where cdusuario ='" & UCase(Trim(Split(lista_correo(i), "@")(0))) & "'"
				arrayRS = GetArrayRS(SQL)
				if not IsArray(arrayRS) then
					if check_mail <> "" then check_mail = check_mail & "|"
					check_mail = check_mail & "El correo : "& Trim(lista_correo(i)) &" no existe."
				end if
			end if
		  else
		    check_mail = check_mail & "El correo : "& Trim(lista_correo(i)) &" tiene que pertenecer a Logis."
		  end if
		else
		  if LCase(Trim(Split(lista_correo(i), "@")(1))) = Get_Mail("logis_mx") then
			if LCase(Trim(lista_correo(0))) <> "christelle" and LCase(Trim(lista_correo(0))) <> "guillehe" and LCase(Trim(lista_correo(0))) <> "jessica"  and LCase(Trim(Split(lista_correo(i), "@")(0))) <> "crossdocktln" and LCase(Trim(lista_correo(0))) <> "cedischihuahua" then
				SQL = "select 1 from usuarios where cdusuario ='" & UCase(Trim(lista_correo(i))) & "'"
				arrayRS = GetArrayRS(SQL)
				if not IsArray(arrayRS) then
					if check_mail <> "" then check_mail = check_mail & "|"
					check_mail = check_mail & "El correo : "& Trim(lista_correo(i)) &" no existe."
				end if			
			end if
		  else
		    check_mail = check_mail & "El correo : "& Trim(lista_correo(i)) &" tiene que pertenecer a Logis."
		  end if
		end if
	next
	
	
end function

function display_parametros(id_reporte, param_id, opcion_param)
	'verifica en la tabla REP_PARAMETROS_SELECT si hay parametros fijos para ponerlos en un SELECT
	dim SQL, arrayRS, i
	
	SQL = " SELECT RPSTIPO, RPSVALUE, RPSTEXT, RPSSELECTED " & vbCrLf
	SQL = SQL & " FROM REP_PARAMETROS_SELECT " & vbCrLf
	SQL = SQL & " WHERE RPS_ID_REP = " & id_reporte & vbCrLf
	SQL = SQL & "   AND RPS_NUM_PARAM = " & param_id & vbCrLf
	SQL = SQL & " ORDER BY RPSTEXT" & vbCrLf
	
	arrayRS = GetArrayRS(SQL)
	if IsArray(arrayRS) then
	    if arrayRS(0,0) = "SELECT" then
	        display_parametros = "<select name=""param_" & param_id & """ class=light>" & vbCrLf 
	        if opcion_param = "1" then
	            display_parametros = display_parametros & vbTab & "<option value="""">--" & vbCrLf
	        end if
	        for i = 0 to UBound(arrayRS, 2)
	            display_parametros = display_parametros & vbTab & "<option value="""& arrayRS(1, i) &""" "& arrayRS(3, i) &">" & arrayRS(2, i) & vbCrLf
	        next
	        display_parametros = display_parametros & "</select>"
	        
	    end if
	end if
	
end function
%>
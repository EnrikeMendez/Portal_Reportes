<%@ Language=VBScript %>
<% option explicit 
%><!--#include file="include/include.asp"--><%
'admin of logis web site :
'Agregacion de reportes ponctuales
Response.Expires = 0
call check_session()
dim SQL, i, arrayRS, secuenciaEmpleado, pernumero, secuenciaPerfil, upo_uorclave, uorusuario, etape
etape = 1

	

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>Alta usuarios</title></head>
<body>
	<style>
	    table thead tr 
		{
			background-color: #223F94;
			font-family: "Roboto",sans-serif;
			color:#FFFFFF;
	    }
		tbody>tr:hover 
		{
			background-color: #E6E3DD;
		}
		 tr:nth-child(even) 
		{
			  background-color: #F6F6F6;
		}
	    button 
		{
			background-color: #223F94;
			color: #FFFFFF;
			font-family: "Roboto",sans-serif;
			font-size: 11px;
			font-weight: bold;
			height: 19px;
	    }
		.center
		{
			    display: flex;
				justify-content: center;
	    }
	    .center-recuadro 
		{
			border-color: gainsboro;
			border-radius: 10px;
			border-style: solid;
			border-width: thin;
			margin-bottom: 2%;
			margin-top: 3%;
			padding: 1%;
			box-shadow: 6px -5px lightgrey;
			font-size: 15px;
	    }
		.left, .right { 
		  width: 40%; 
		  margin:5px; 
		  padding: 1em; 
		  background: white; 
		} 
		.left  { float:left;  }
		.right { float:right; } 
	</style>
	<div class="center">
        <div class="center-recuadro" style="width: 50%"> 
			<%Select Case Request.Form("Etape")
				Case "1"
					etape = 2
					SQL = "select uorclave, uorusuario, date_created, date_modified, created_by "
					SQL = SQL & "from EUSUARIOS_ORFEO "
					If IsNumeric(Request.Form("nombre_empleado_usuario")) Then
							SQL = SQL &  "WHERE uor_pernumero = " & Request.Form("nombre_empleado_usuario")
					else
						SQL = SQL &  "WHERE UORUSUARIO = upper('" & Request.Form("nombre_empleado_usuario") & "') "
					End If	
					
					arrayRS = GetArrayRS(SQL)
					if IsArray(arrayRS) then 
						upo_uorclave = arrayRS(0,0) 
						uorusuario = arrayRS(1,0) 
						Response.Write "<p style='color:red; text-align:center;'><strong> El usuario " & arrayRS(1,0) & " ya tiene persimso. Creado por " & arrayRS(4,0) & "</strong></p> " & upo_uorclave
					else
						Response.Write "<p style='color:red; text-align:center;'><strong> El usuario " & Request.Form("nombre_empleado_usuario") & " no tiene persmisos se agregara el prermiso 3 </strong></p> "
					end if
				Case "2"					
					dim objParam_p_Codigo, objParam_p_Mensaje
					dim p_Codigo, p_Mensaje
					dim Conn, rst1
					dim objCommand, objParam, rsYourRecordSet, objParm

					Set objCommand = Server.CreateObject("ADODB.Command")
					Set objParam = Server.CreateObject("ADODB.Parameter")
					Set rsYourRecordSet = Server.CreateObject("ADODB.Recordset")
    
					Set Conn = CreateObject("ADODB.Connection")
					Conn.ConnectionTimeout = 4000
					Conn.CommandTimeout = 4000					
					upo_uorclave = Request.Form("datos_empleado")
					
					'Conn.open Get_Conn_string("SERVER"), Get_Conn_string("LOGIN"), Get_Conn_string("PASS")        
					Conn.open Get_Conn_string("SERVER"), Get_Conn_string("LOGIN_ADM"), Get_Conn_string("PASS_ADM")        
					
					'Conn.BeginTrans
					objCommand.ActiveConnection = Conn

					objCommand.CommandType = 4
					objCommand.commandtext = "LOGIS.PR_ALTA_USUARIO_PERFIL"
					if IsNull(upo_uorclave) or upo_uorclave = "" then ' 1
						Set objParm = objCommand.CreateParameter("@p_uorclave", 5, 1, 7, 0)
					else 
						Set objParm = objCommand.CreateParameter("@p_uorclave", 5, 1, 7, upo_uorclave)
					end if
					
					
					
					objCommand.Parameters.Append objParm
					Set objParm = objCommand.CreateParameter("@p_peoclave", 5, 1, 7, Request.Form("nuevo_permisos")) ' 2
					objCommand.Parameters.Append objParm
					Set objParm = objCommand.CreateParameter("@p_usuario", 200, 1, 30, Request.Form("usuario")) ' 3
					objCommand.Parameters.Append objParm
					Set objParm = objCommand.CreateParameter("@p_pernumero", 5, 1, 8, Request.Form("numero_empleado")) ' 4
					objCommand.Parameters.Append objParm

    
					set objParam_p_Codigo = Server.CreateObject("ADODB.Parameter")    ' 5
					objParam_p_Codigo.Name = "@p_Codigo"
					objParam_p_Codigo.Direction = 2
					objParam_p_Codigo.Type = 4
					objParam_p_Codigo.Size = 3
					objCommand.Parameters.Append objParam_p_Codigo

					set objParam_p_Mensaje = Server.CreateObject("ADODB.Parameter")     ' 6
					objParam_p_Mensaje.Name = "@p_Mensaje"
					objParam_p_Mensaje.Direction = 2
					objParam_p_Mensaje.Type = 201
					objParam_p_Mensaje.Size = 200
					objCommand.Parameters.Append objParam_p_Mensaje
					objCommand.Execute()
					p_Codigo = objCommand.Parameters(objCommand.Parameters(4).Name).Value
					p_Mensaje  = objCommand.Parameters(objCommand.Parameters(5).Name).Value    
					Response.Write  p_Mensaje & "<br/><br/>"					
			end select%>
			<%if Request.QueryString("msg") <> "" AND Request.Form("nombre_empleado_usuario") = "" then%>
				<strong style='color:red;'><%=Request.QueryString("msg")%> </strong>
				<br />
				<br />
			<%end if%>
			<form name="valid_conf" method="post">
				<label>Número de empleado o usuario</label>
				<input type="text" name="nombre_empleado_usuario" 
					value="<%=Request.Form("nombre_empleado_usuario") %>" 
					<%if Request.Form("nombre_empleado_usuario") <> "" then %>readonly style="background: #F6F6F6;"<% end if %>/>
				<input type="hidden" name="datos_empleado" id="datos_empleado" value="<%=upo_uorclave %>" />
				<input type="hidden" name="Etape" value="<%=etape %>">
				<% if Request.Form("Etape") = "" then %>
			
					<button > Buscar</button>
				<%end if %>
		
				<%
					Select Case Request.Form("Etape")
						Case "1", "2"
				%>
					<br />
					<br />
					<%if upo_uorclave = "" then %>
						<label>Usuario </label>
						<input type="text" name="usuario"/>
						<br />
						<br />
						<label>Número de empleado: </label>
						<input type="text" name="numero_empleado"/>
						<br />
						<br />
						<label>Nuevo permiso</label>
						<input type="text" name="nuevo_permisos" value="3" readonly style="background: #F6F6F6;"/>
					<% else%>
						<label>Nuevo Perfile</label>
						<input type="text" name="nuevo_permisos" value="3" />
						<input type="hidden" name="usuario" value="<%=uorusuario %>"/>
						<input type="hidden" name="numero_empleado" value="<%=upo_uorclave %>"/>
					<%end if %>
			
					<button > Guardar</button>
					<a href="alta_perfil_usuarios.asp">Regresar</a>
				<%end select %>
			</form>

			<%
		
				Select Case Request.Form("Etape")
					Case "1", "2"
						if upo_uorclave <> "" then 
							'SQL = "select upo_peoclave, upo_peoclave, upo_uorclave, created_by "
							'SQL = SQL & "from EUSUARIOS_PERFILES_ORFEO where upo_uorclave = " & upo_uorclave 
							SQL = "select estatus.upo_peoclave, perfiles.peonombre "
							SQL = SQL & " from EUSUARIOS_PERFILES_ORFEO estatus "
							SQL = SQL & " join eperfiles_orfeo  perfiles on estatus.upo_peoclave = perfiles.peoclave "
							SQL = SQL & " where estatus.upo_uorclave = " & upo_uorclave 
							
					
							arrayRS = GetArrayRS(SQL)
							if IsArray(arrayRS) then 
							%>
							<table width=40%>
								<thead>
									<tr bgcolor="goldenrod" >
										<td rowspan=2 style="text-align: right;">Id</td>
										<td rowspan=2>Perfile</td>
									</tr>
								</thead>
								<tbody>
									<%
										for i = 0 to UBound(arrayRS,2)
											Response.Write "<tr>"
											Response.Write "<td style='text-align: right;'>" & arrayRS(0,i) & "</td>" 
											Response.Write "<td>" & arrayRS(1,i) & "</td>" 
											Response.Write "</tr>"
										next
									%>
								</tbody>
							</table>
							<%
							end if
						end if
				end select %>
		</div>
	</div>
</body>
</html>

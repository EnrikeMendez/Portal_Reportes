
<!--#include file="include\include.asp"-->
<%
'CHG-DESA-17082021-01 <--
DIM USR, PSW, SQL,i, rst, array_usr, array_VILCLEF
    
set rst = Server.CreateObject("ADODB.Recordset")
USR = UCase(REQUEST("username"))
PSW = UCase(REQUEST("password"))
PROCEDENCIA = UCase(REQUEST("procedencia"))

'acceso_desact_anom : permite a estos usuarios de desactivar las anomalias en la pantalla   
'contabilidad : permite confirmar los reportes de estado de resultados.  
    
 if PROCEDENCIA = "TEST2_CIUDADES" then     
    CP = UCase(REQUEST("cp-enviado"))
    CPCIUDADCOLONIA = UCase(REQUEST("cp-ciudad-colonia"))
    TIPOENTREGA = UCase(REQUEST("tipo-entrega-enviado"))
    if SQLescape(USR) = "HECTORRR"  OR SQLescape(USR) = "ESTEPHANIAGH" OR SQLescape(USR) = "SANDYJM" then 
        
        SQL = "select DISTINCT cdusuario from USUARIOS where UPPER(cdusuario) = UPPER('" & SQLescape(USR) & "') " 
        SQL = SQL & "AND UPPER(DSUSUARIO) = '" & SQLescape(PSW) & "'"
        
        array_usr = GetArrayRS(SQL)
        
        
        IF not IsArray(array_usr) then
            Response.Redirect("/test2/ciudades.asp?ERR=1&CP=" & CP & "&CPCIUDADCOLONIA=" & CPCIUDADCOLONIA & "&TIPOENTREGA=" & TIPOENTREGA)
        else
            SQL = "SELECT DISTINCT VILCLEF " 
            SQL = SQL & " FROM ECODIGOS_POSTALES, ECIUDADES, EESTADOS , EDESTINOS_POR_RUTA, EALMACENES_LOGIS "
            SQL = SQL & " WHERE (ECOP_D_CODIGO = '" & CP & "')"
            SQL = SQL & " AND VILCLEF = ECOP_VILCLEF "
            SQL = SQL & " AND ESTESTADO = VIL_ESTESTADO "
            SQL = SQL & " AND DER_VILCLEF(+) = VILCLEF "
            SQL = SQL & " AND ALLCLAVE(+) = DER_ALLCLAVE "
            SQL = SQL & " ORDER BY 1"
        
            array_VILCLEF = GetArrayRS(SQL)
	        SQL = "UPDATE EDESTINOS_POR_RUTA SET DATE_MODIFIED = SYSDATE ,MODIFIED_BY = 'SISTEMAS',DER_TIPO_ENTREGA ='" & TIPOENTREGA & "'  WHERE DER_VILCLEF = " & array_VILCLEF(0,0)
            'Response.Write SQL 
            rst.Open SQL, Connect(), 0, 1, 1
            response.redirect("/test2/ciudades.asp?ERR=0&mensage=Se guardo " & CPCIUDADCOLONIA & " con el tipo de entrega " & TIPOENTREGA & "&CP=" & CP & "&CPCIUDADCOLONIA=" & CPCIUDADCOLONIA & "&TIPOENTREGA=" & TIPOENTREGA)
        end if
    end if
    Response.Redirect("/test2/ciudades.asp?ERR=2&errormensaje=El usuario " & SQLescape(USR) & " no tiene permisos.&CP=" & CP & "&CPCIUDADCOLONIA=" & CPCIUDADCOLONIA & "&TIPOENTREGA=" & TIPOENTREGA)
    
else 
    
    SQL = "  select DISTINCT US.USUARIO, US.ADUANA"
    'SQL = SQL &  "  , InitCap(DOU.DOUABREVIACION), decode(US.USUARIO, 'SILVIAVI', 1, 'NICOLAST', 1, 'CHRIS', 1, 'PATYC', 1, 'ROSALBAC', 1, 'MAGALYCF', 1, 'JUDITHGV', 1, 'GABRIELACD', 1, 'VIRGINIAS', 1, 'ALEJANDRAAS', 1, 'MAYRAGO', 1, 0) acceso_desact_anom"
    SQL = SQL &  "  , InitCap(DOU.DOUABREVIACION), decode(US.USUARIO, 'SILVIAVI', 1, 'NICOLAST', 1, 'CHRISTELLE', 1, 'PATYC', 1, 'ROSALBAC', 1, 'MAGALYCF', 1, 'JUDITHGV', 1, 'GABRIELACD', 1, 'VIRGINIAS', 1, 'ALEJANDRAAS', 1, 'MAYRAGO', 1, 'MEUGENIAC', 1, 'RAULCZ', 1, 0) acceso_desact_anom"
    SQL = SQL &  "  , CONTABILIDAD "
    SQL = SQL &  "  from  rep_usuarios_aduana US"
    SQL = SQL &  "  , USUARIOS US2"
    SQL = SQL &  "  , Edouane  DOU"
    SQL = SQL &  " where UPPER(US.USUARIO) = UPPER('"&SQLescape(USR)&"') "
    SQL = SQL &  " AND UPPER(US2.DSUSUARIO) = UPPER('"&SQLescape(PSW)&"') "
    SQL = SQL &  "  AND US2.CDUSUARIO = US.USUARIO"
    SQL = SQL &  "  AND US.ADUANA = DOU.DOUCLEF(+)"
    SQL = SQL &  "  order by InitCap(DOU.DOUABREVIACION)"

    'Response.Write SQL
    

    array_usr = GetArrayRS(SQL)
    IF not IsArray(array_usr) then
        'Response.Write ("not array")
        Response.Redirect("login.asp?ERR=1")
    else
	    Session("array_user")= array_usr 
        response.redirect("menu.asp")
    end if
end if

'CHG-DESA-17082021-01 -->
%>
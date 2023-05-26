<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                              xmlns:oxml='http://www.ventanillaunica.gob.mx/cove/ws/oxml/'
                              xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/">
<xsl:output method="xml" omit-xml-declaration="yes"/>  
<xsl:param name="param1"/>

<xsl:template match="soapenv:Header">
</xsl:template>

<xsl:template match="oxml:solicitarRecibirRelacionFacturasNoIAServicio/oxml:comprobantes">
  <h2>Comprobante de valor electronico</h2>
  
  <h3>Datos del comprobante</h3>
    <table align="center" BORDER="1" width="600" cellpadding="2" cellspacing="0">
      <tr>
        <td bgcolor="goldenrod">Tipo de Operación</td><td><xsl:value-of select="oxml:tipoOperacion"/></td>
      </tr>
      <tr>
        <td bgcolor="goldenrod">Relacion de factura</td><td><xsl:value-of select="oxml:numeroRelacionFacturas"/></td>
        <td bgcolor="goldenrod">Fecha Exp.</td><td><xsl:value-of select="oxml:fechaExpedicion"/></td>
      </tr>
      <tr>
        <td bgcolor="goldenrod">Patente</td><td><xsl:value-of select="oxml:patenteAduanal"/></td>
        <td bgcolor="goldenrod">RFC Consulta</td><td><xsl:value-of select="oxml:rfcConsulta"/></td>
      </tr>
      <tr>
        <td bgcolor="goldenrod">Tipo de figura</td><td colspan="3"><xsl:value-of select="oxml:tipoFigura"/> 
        <i>(1 - Agente Aduanal, 4 - Exportador, 5 - Importador)</i></td>
      </tr>
      <tr>
        <td bgcolor="goldenrod">eDocument</td><td><xsl:value-of select="$param1"/></td>
      </tr>
      <tr>
        <td bgcolor="goldenrod">Observaciones</td><td colspam="3"><xsl:value-of select="oxml:observaciones"/></td>
      </tr>
    </table>

	<h3>Facturas en la relacion</h3>

	<xsl:for-each select="oxml:facturas">
		<h3>Datos de facturas</h3>
		<table align="center" BORDER="1" width="600" cellpadding="2" cellspacing="0">
		  <tr>
			<td bgcolor="goldenrod">No. de factura</td><td><xsl:value-of select="oxml:numeroFactura"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Subdivisión</td><td><xsl:value-of select="oxml:subdivision"/> </td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Certificado de origen</td><td><xsl:value-of select="oxml:certificadoOrigen"/> </td>
		  </tr>
		</table>

		<h4>Datos del emisor</h4>
		<table align="center" border="1" width="600" cellpadding="2" cellspacing="0">
		  <tr>
			<td bgcolor="goldenrod">Tipo de identificador</td><td><xsl:value-of select="oxml:emisor/oxml:tipoIdentificador"/>
			<i> (0 - Tax ID, 1 - RFC)</i>
			</td>
			<td bgcolor="goldenrod">Identificación</td><td><xsl:value-of select="oxml:emisor/oxml:identificacion"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Razón Social</td><td colspan="3"><xsl:value-of select="oxml:emisor/oxml:nombre"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">No. exterior</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:numeroExterior"/></td>
			<td bgcolor="goldenrod">No. interior</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:numeroInterior"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Calle</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:calle"/></td>
			<td bgcolor="goldenrod">Colonia</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:colonia"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Municipio</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:municipio"/></td>
			<td bgcolor="goldenrod">Entidad federativa</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:entidadFederativa"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">País</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:pais"/></td>
			<td bgcolor="goldenrod">Código postal</td><td><xsl:value-of select="oxml:emisor/oxml:domicilio/oxml:codigoPostal"/></td>
		  </tr>
		</table> 		


		<h4>Datos del destinatario</h4>
		<table align="center" border="1" width="600" cellpadding="2" cellspacing="0">
		  <tr>
			<td bgcolor="goldenrod">Tipo de identificador</td><td><xsl:value-of select="oxml:destinatario/oxml:tipoIdentificador"/>
			<i> (0 - Tax ID, 1 - RFC)</i>
			</td>
			<td bgcolor="goldenrod">Identificación</td><td><xsl:value-of select="oxml:destinatario/oxml:identificacion"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Razón Social</td><td colspan="3"><xsl:value-of select="oxml:destinatario/oxml:nombre"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">No. exterior</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:numeroExterior"/></td>
			<td bgcolor="goldenrod">No. interior</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:numeroInterior"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Calle</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:calle"/></td>
			<td bgcolor="goldenrod">Colonia</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:colonia"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">Municipio</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:municipio"/></td>
			<td bgcolor="goldenrod">Entidad federativa</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:entidadFederativa"/></td>
		  </tr>
		  <tr>
			<td bgcolor="goldenrod">País</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:pais"/></td>
			<td bgcolor="goldenrod">Código postal</td><td><xsl:value-of select="oxml:destinatario/oxml:domicilio/oxml:codigoPostal"/></td>
		  </tr>
		</table>  
  
		<h4>Mercancias</h4>
		<table align="center" border="1" width="800" cellpadding="2" cellspacing="0"> 
		  <tr bgcolor="goldenrod">
			<td>Descripción</td>
			<td>Clave UMC</td>
			<td>Cantidad UMC</td>
			<td>Valor unitario</td>
			<td>Valor total</td>
			<td>Tipo moneda</td>
			<td>Valor total USD</td>        

            <td>Marca</td>
            <td>Modelo</td>
            <td>SubModelo/Lote</td>
            <td>Numero Serie</td>

		  </tr>
		  <xsl:for-each select="oxml:mercancias">
		  <tr>
			<td><xsl:value-of select="oxml:descripcionGenerica"/></td>
			<td><xsl:value-of select="oxml:claveUnidadMedida"/></td>
			<td><xsl:value-of select="oxml:cantidad"/></td>
			<td><xsl:value-of select="oxml:valorUnitario"/></td>
			<td><xsl:value-of select="oxml:valorTotal"/></td>
			<td><xsl:value-of select="oxml:tipoMoneda"/></td>
			<td><xsl:value-of select="oxml:valorDolares"/></td>
		  </tr>


		  <xsl:for-each select="oxml:descripcionesEspecificas">
		  <tr>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td><xsl:value-of select="oxml:marca"/></td>
			<td><xsl:value-of select="oxml:modelo"/></td>
			<td><xsl:value-of select="oxml:subModelo"/></td>
			<td><xsl:value-of select="oxml:numeroSerie"/></td>
		  </tr>
		  </xsl:for-each>


		  </xsl:for-each>
		</table>
		

	</xsl:for-each>
	
</xsl:template>
</xsl:stylesheet>

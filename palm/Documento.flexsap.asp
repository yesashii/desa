<!--#include Virtual="/includes/conexion.asp"-->
<% 'Server.ScriptTimeOut= %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Consulta Documento</TITLE>
<META NAME="Generator"   CONTENT="EditPlus">
<META NAME="Author"      CONTENT="Simon Hernandez">
<META NAME="Keywords"    CONTENT="">
<META NAME="Description" CONTENT="consulta Documentos">
<style type="text/css">
	#noprint {
		display:none;
	}
	body{
		background-color: #F4F4F4;
		font-family: verdana;
		font-size: 12px;
	}
	th{
		font-family: verdana;
		font-size: 12px;
		color: #000033;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		background-color: #000066*/
	}
	td{
		font-family: verdana;
		font-size: 12px;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		padding: 2px;*/
	}
	table{
		border-collapse: collapse;
		/*border: 1px;
		*border-color: #000066;
		background-color: #FFFFFF;*/
	}
</style>
</HEAD>

<BODY onLoad="borraespere()">
<% 
'---------------------------------------------------------------------------------------------
'Private sub Main()
	set r = response
	dim basex
	mibusca  =recuperavalor("mibusca"  )
	documento=recuperavalor("factura"  )
	buspor   =recuperavalor("buspor"   )
	empresa  =recuperavalor("empresa"  )
	tipodocto=recuperavalor("tipodocto")
	mitop    =recuperavalor("mitop"    )
	base     =recuperavalor("base"     )
    correlati=recuperavalor("correlati")
	nuser    =recuperavalor("nuser"    )

	'response.write(correlati)
	if correlati="0" then
		documento="Aun No existe Documento"
		%>
		<CENTER>
			Aun No existe Documento
		</CENTER>
		<%		
	end if
	if len(correlati)>1 then
		sql="select numero from serverdesa.BDFlexline.flexline.documento as d " & _
		"where tipodocto in ('facturas palm','facturas movil','fact. afecta elec') and empresa='" &  empresa & "' and correlativo=" & correlati & _
		" UNION SELECT RIGHT(CAST('0000' + [Numero Factura] AS NVARCHAR),10) AS numero FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURAENC " & _
		"WHERE [Numero Factura] =  "& correlati
		
		'r.write(sql)
		Set rs=oConn.execute(sql)
		documento=rs(0)
		'response.write(documento)
		
	end if
	'r.write "Base : " & base

	if empresa="DESA"  then idempresa=1
	if empresa="LACAV" then idempresa=4

	if len(mitop)=0     then mitop="50"
	if len(documento)=0 then documento=recuperavalor("documento")

	'valida numero
	if len(documento)>0 then
		if ucase(left(documento,1))="X" then
			if not len(documento)=11 then
				documento="X" & right("0000000000" & right(documento,len(documento)-1),10)
			end if
		else
			documento=right("0000000000" & documento,10)
		end if
	end if
'r.write(documento)
	if len(documento)>0 then
		if len(empresa)>0 and len(tipodocto)>0 then
			mostrardocumento()
		else
			call ListaDocumentos()
		end if
	else
		if len(mibusca)>0 and len(buspor)>0 then
			call ListaDocumentos()
		else
			call buscasimple()
		end if
	end if

	'paso=4
	'if paso=0 then call buscasimple()
	'if paso=1 then call buscaavanzada()
	'if paso=2 then call ListaDocumentos()
	'if paso=3 then call ListaDocumentos()
	'if paso=4 then call mostrardocumento()

'End sub 'Main()
'---------------------------------------------------------------------------------------------
private sub buscasimple()
	%>
	<CENTER>
	<FONT SIZE="2" COLOR="#000066">Busca documentos por su numero</FONT>
	<BR>
	<FORM METHOD=POST ACTION="Documento.asp">
		<INPUT TYPE="text" NAME="factura" size="10">
		<INPUT TYPE="submit" value="Buscar">
		<INPUT TYPE="hidden" name="nuser" value="<%=nuser%>">
	</FORM>
	<HR>
	<FONT SIZE="2" COLOR="#000066">Busca documentos por Nombre o Rut</FONT>
	<FORM METHOD=POST ACTION="Documento.asp">
		<TABLE>
		<TR>
			<TD><INPUT TYPE="text" NAME="mibusca"></TD>
		</TR>
		<TR>
			<TD>
				<INPUT TYPE="radio" NAME="buspor" id="buscapor1" value="rs">
				<LABEL for='buscapor1'>Razon Social</LABEL>
			</TD>
		</TR>
		<TR>
			<TD>
				<INPUT TYPE="radio" NAME="buspor" id="buscapor2" value="rut" CHECKED>
				<LABEL for='buscapor2'>Codigo legal</LABEL>
			</TD>
		</TR>
		<TR>
			<TD><INPUT TYPE="submit" value="Buscar"></TD>
		</TR>
		</TABLE>
		<INPUT TYPE="hidden" name="nuser" value="<%=nuser%>">
	</FORM>
	<HR>
	<input type="button" value="<< Atras" onClick="history.back()"> 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<!-- <A HREF="">Busqueda Avanzada >></A> -->
	</CENTER>
	<%
End sub 'buscasimple()
'---------------------------------------------------------------------------------------------
private sub buscaavanzada()
End sub 'buscaavanzada()
'---------------------------------------------------------------------------------------------
Private sub ListaDocumentos()
	on error resume next
	%><CENTER><div id="espere"><B>Espere&nbsp;mientras&nbsp;carga&nbsp;...&nbsp;</B></div><%
	r.flush()
	%><FORM METHOD=POST ACTION="">
	<B><FONT SIZE="2" COLOR="#808080">Base Activa</FONT></B><%
	sql="SELECT TOP " & mitop & " * FROM (SELECT documento.numero as Numero, documento.tipodocto, ctacte.razonsocial as Nombre, " & _
		"documento.fecha, documento.cliente, documento.vigencia as V, documento.empresa, ctacte.sigla, ctacte.ejecutivo " & _ 
		"from serverdesa.BDFlexline.flexline.documento as documento " & _ 
		"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte on ctacte.empresa=documento.empresa " & _
		"and ctacte.ctacte=documento.cliente and ctacte.tipoctacte=documento.tipoctacte " & _ 
		"inner join serverdesa.BDFlexline.flexline.tipodocumento as tipodocumento on tipodocumento.empresa=documento.empresa and tipodocumento.tipodocto=documento.tipodocto " & _ 
		"where documento.empresa in ('DESA','LACAV','UNDURRAGA') and tipodocumento.sistema='VENTAS' and tipodocumento.clase = 'Factura (v)' and documento.factormonto<>0 " & _
		"UNION " & _
		"SELECT RIGHT('00000' + [Numero Factura], 10) AS Numero, 'FACT. AFECTA ELEC' AS tipodocto, Nombre,  " & _
		"[Fecha Emision] AS Fecha, CtaCte AS Cliente, 'V' AS Vigencia, 'DESA' AS Empresa, sigla, [Vendedor Factura] AS ejecutivo " & _
		"FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURAENC " & _
		") AS FACTURAS WHERE "
	if len(documento) then
		sql=sql & "numero='" & documento & "'"
		sql=sql & " or (tipodocto ='GUIA DESPACHO ELEC' and numero='" & documento & "')"
	else
		if buspor="rut" then
			sql=sql & "cliente like '%" & mibusca & "%'"
		else
			sql=sql & "nombre + sigla like '%" & mibusca & "%'"
		end if
	end if
		
	'response.write sql

	if len(nuser)>0 then
		SW_CONSULTA=consultarapida("select SW_CONSULTA from handheld.dbo.dim_vendedores where idEMPRESA='1' AND idvendedor=" & cint(nuser) )
		nombrevendedor=consultarapida("select NOMBRE   from handheld.dbo.dim_vendedores where idEMPRESA='1' AND idvendedor=" & cint(nuser))
		'RESPONSE.WRITE ("<BR>nuser : " &  nuser)
		'RESPONSE.WRITE ("<BR>sw_consulta : " &  SW_CONSULTA)
		'RESPONSE.WRITE ("<BR>NombreVendedor : " &  nombrevendedor)
		'sql=sql & " and ctacte.ejecutivo='" & nombrevendedor & "' "
		IF UCASE(SW_CONSULTA)="N" then
			sql=sql & " and ejecutivo='" & nombrevendedor & "' "
		end if
	end if
	'r.write(sql)
	'call tablasimple(sql,"color")
	sql=sql & " order by fecha desc, numero desc "
	set rs=oConn.execute(sql)
if err<>0 then
	'response.write ("<BR>Error Desc : " &  err.description )'
	if err=-2147217900 then 
		response.write ("" )'
		%>
		<HR>
		<H3>
		<BR>No Existe comunicacion con el servidor SERVERDESA, 
		<BR>intente la consulta en otro momento
		</H3>
		<INPUT TYPE="button" value="<< Atras"      onclick="history.back();history.back()">
		<%
	end if
	%>	
		<INPUT TYPE="checkbox" NAME="ckerr" id="ckerr"
		onclick=" 
		if (document.getElementById('ckerr').checked){
			document.getElementById('ctlerr').style.display = '';
		}else{
			document.getElementById('ctlerr').style.display = 'none';
		}
		"
		><label for="ckerr"><FONT SIZE="2" face="verdana" COLOR="#808080">Mostrar detalle error</FONT></label>
		<DIV id="ctlerr" name="ctlerr" style="display:none">
		<TABLE bgcolor="#C0C0C0"  width='300px' border='1' cellspacing='0' cellpadding='0' style="border-collapse: collapse">
		<TR>
			<TD><%=err.number%></TD>
		</TR>
		<TR>
			<TD><%=err.description%></TD>
		</TR>
		</TABLE>
		</DIV>
	<%
	exit sub
end if
	doctos=0
	'doctox=0
	do until rs.eof
		empresa  =rs("empresa"  )
		tipodocto=rs("tipodocto")
		doctos   =doctos+1
		basex="BDFlexline"
		documento=rs("numero")
	'	doctox=doctox+1
	rs.movenext
	loop
	
	call tablasimple(rs,"color")
	r.flush()

	'r.Write("<HR><B><FONT SIZE="2" COLOR="#808080">Base Historica</FONT></B>") 'Activar Para Base Histórica
	sql=replace(sql,"BDFlexline","BDHistorica")
	'r.write(sql)
	set rs=oConn.execute(sql)
	'doctox=0
	do until rs.eof
		empresa  =rs("empresa"  )
		tipodocto=rs("tipodocto")
		doctos   =doctos+1
		basex="BDHistorica"
		documento=rs("numero")
	'	doctox=doctox+1
	rs.movenext
	loop
	
'	call tablasimple(rs,"color") 'Activar Para Base Histórica

	'call tablasimple(sql,"color")
	'tablasimple(sql,"color")

	'r.write doctos
	if doctos=0 then
		%><BR><BR>La busqueda no obtuvo resultados<%
		if len(nombrevendedor)>0 then 
			response.write("<BR>Para el Vendedor " & nombrevendedor & "<BR>")
		end if 
		'r.write("<BR>documento : " & documento)
		'r.write("<BR>mibusca : " & mibusca)
		'r.write("<BR>buspor : " & buspor)
		if buspor="rs"      then r.write("<BR>Razonsocial : " & mibusca)
		if buspor="rut"     then r.write("<BR>Rut cliente : " & mibusca)
		if len(documento)>0 then r.write("<BR>Documento : " & documento)
	end if
	%><HR>
	<INPUT TYPE="hidden" name="doctos"    id="doctos"    value="<%=doctos%>"   >
	<INPUT TYPE="hidden" name="empresa"   id="empresa"   value="<%=empresa%>"  >
	<INPUT TYPE="hidden" name="documento" id="documento" value="<%=documento%>">
	<INPUT TYPE="hidden" name="tipodocto" id="tipodocto" value="<%=tipodocto%>">
	<INPUT TYPE="hidden" name="basex"     id="basex"     value="<%=basex%>"    >
	<INPUT TYPE="hidden" name="buspor"    id="buspor"    value="<%=buspor%>"   >
	<INPUT TYPE="hidden" name="mibusca"   id="buspor"    value="<%=mibusca%>"  >

	<TABLE>
	<TR>
		<TD><INPUT TYPE="button" value="<< Atras" onClick="history.back()"></TD>
		<TD>&nbsp;&nbsp;&nbsp;</TD>
		<TD>
		<TABLE bordercolor='#C0C0C0' border='1'>
			<TD>&nbsp;Mostrar&nbsp;las&nbsp;ultimas&nbsp;</TD>
			<TD>
				<SELECT NAME="mitop">
					<OPTION value="50" > 50</OPTION>
					<OPTION value="100">100</OPTION>
					<OPTION value="200">200</OPTION>
					<OPTION value="100 PERCENT ">Todas</OPTION>
				</SELECT>
			</TD>
			<TD><INPUT TYPE="submit" Value="Aplicar" onClick="document.getElementById('documento').value=''"></TD>
		</TABLE>
		</TD>
	</TR>
	</TABLE><%
	r.flush()
	%><SCRIPT LANGUAGE="JavaScript">
	<!--
		document.getElementById("espere").style.display='none';
		var doctos   =document.getElementById("doctos"   ).value;
		var empresa  =document.getElementById("empresa"  ).value;
		var documento=document.getElementById("documento").value;
		var tipodocto=document.getElementById("tipodocto").value;
		var basex    =document.getElementById("basex"    ).value;

		if (doctos==1 ){window.open('documento.asp?empresa='+empresa+'&documento='+documento+'&tipodocto='+tipodocto+'&base='+basex+'&nocache='+Math.random() ,'_self')}
	//-->
	</SCRIPT></CENTER>
	</FORM><%
	if err<>0 then
	'response.write err.description
		
	end if
End sub 'ListaDocumentos()
'---------------------------------------------------------------------------------------------
Private sub mostrardocumento()
'on error resume next
'response.write tipodocto 
	%><CENTER><div id="espere"><B>Espere&nbsp;mientras&nbsp;carga&nbsp;...&nbsp;</B></div><%
	dim matriz(10)
	matriz( 1)="Nombre. .   |ctacte.razonsocial"
	matriz( 2)="id Cliente  |ctacte.ctacte, ctacte.sigla, documento.fecha"
	matriz( 3)="Direccion   |ctacte.Direccionenvio , ctacte.comuna"
	matriz( 4)="Cond Pago   |documentop.codigopago"
	matriz( 5)="Vend Cliente|ctacte.ejecutivo"
	matriz( 6)="Vend Factura|documento.vendedor"
	matriz( 7)="Local-bodega|documento.local, documento.bodega"
	matriz( 8)="Referencia  |documento.referenciaexterna"
	miselect=""
	for x=1 to 9
		a=split(matriz(x),"|")
		if ubound(a)=1 then
			'r.write "<BR>u(a):" & ubound(a)
	'		r.write "<BR>a(0):" & a(0)
	'		r.write "<BR>a(1):" & a(1)
			miselect=miselect & ", " & a(1)
		end if
	next
	
	miselect= "SELECT * FROM (Select documento.vigencia, documento.total, documento.correlativo " & miselect
	miselect=replace(miselect,"ctacte.Direccionenvio","handheld.flexline.ctactedireccion(documento.cliente) as Direccionenvio")
	miselect= miselect + ", documento.EMPRESA, documento.NUMERO, documento.TIPODOCTO"

	sqldatos=" from serverdesa.BDFlexline.flexline.documento as Documento " & _
	"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte " & _
	"on ctacte.empresa=documento.empresa and ctacte.tipoctacte=documento.tipoctacte and documento.cliente=ctacte.ctacte " & _
	"inner join serverdesa.BDFlexline.flexline.documentop documentop " & _
	"on documentop.empresa=documento.empresa and documento.tipodocto=documentop.tipodocto and documento.correlativo=documentop.correlativo " & _
	"UNION " & _
"SELECT CASE WHEN Vigencia = 'N' THEN 'S' ELSE 'N' END AS VIGENCIA, " & _
"TOTAL, CAST([Numero factura] AS NUMERIC(18, 0)) AS CORRELATIVO, Nombre AS RAZONSOCIAL, " & _
"CTACTE, SIGLA, [FECHA EMISION] AS FECHA, [Direccion Despacho] AS DIRECCIONENVIO, COMUNA, " & _
"[Condicion de Pago] AS CODIGOPAGO, [Vendedor Cliente] AS EJECUTIVO, [Vendedor Factura] AS VENDEDOR, " & _
"Sucursal AS [LOCAL], BODEGA,'0000' AS REFERENCIAEXTERNA, 'DESA' AS EMPRESA,  " & _
"RIGHT('0000' + [Numero Factura],10) AS NUMERO, 'FACT. AFECTA ELEC' AS TIPODOCTO " & _
"FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURAENC  " & _
")ENCAB " & _
	"where empresa='" & empresa & "' and numero='" & documento & "' and TipoDocto='" & tipodocto & "'"
	
	if tipodocto="GUIA DESPACHO ELEC" then 
	sqldatos=" from serverdesa.BDFlexline.flexline.documento as Documento " & _
	"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte " & _
	"on ctacte.empresa=documento.empresa and ctacte.tipoctacte=documento.tipoctacte and documento.cliente=ctacte.ctacte " & _
	"inner join (select 'desa' as empresa, '" & tipodocto & "' as tipodocto, '" & documento & "' as correlativo, '' as codigopago) documentop " & _
	"on documentop.empresa=documento.empresa and documento.tipodocto=documentop.tipodocto and documento.correlativo=documentop.correlativo " & _
		"UNION " & _
"SELECT CASE WHEN Vigencia = 'N' THEN 'S' ELSE 'N' END AS VIGENCIA, " & _
"TOTAL, CAST([Numero factura] AS NUMERIC(18, 0)) AS CORRELATIVO, Nombre AS RAZONSOCIAL, " & _
"CTACTE, SIGLA, [FECHA EMISION] AS FECHA, [Direccion Despacho] AS DIRECCIONENVIO, 'SIN DATO' AS COMUNA, " & _
"[Condicion de Pago] AS CODIGOPAGO, [Vendedor Cliente] AS EJECUTIVO, [Vendedor Factura] AS VENDEDOR, " & _
"Sucursal AS [LOCAL], BODEGA, referencia AS REFERENCIAEXTERNA, 'DESA' AS EMPRESA,  " & _
"RIGHT('0000' + [Numero Factura],10) AS NUMERO, 'FACT. AFECTA ELEC' AS TIPODOCTO " & _
"FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURAENC  " & _
")ENCAB " & _
	"where empresa='" & empresa & "' and numero='" & documento & "' and TipoDocto='" & tipodocto & "'"
	end if

	sql=miselect & sqldatos

	sql=replace(ucase(sql),"BDFlexline",base)
'r.write(sql)
	set rs=oConn.execute(sql)

	'if rs.eof then exit sub
	if rs("Vigencia")="A" then 
		%><FONT COLOR="#CC0000"><B>Documento NULO</B></FONT><%
	End if
	
	if rs("Vigencia")="N" then 
		%><FONT COLOR="#CC0000"><B>Documento No Vigente</B></FONT><%
	End if
	mitotal= FormatNumber(cdbl(rs.fields("total")),0)
'	sapFlete =
'	sapImpuesto =
	correlativo=rs.fields("correlativo")

	dempresa="Distribuci&oacute;n y Excelencia S.A."
	if empresa="DESA"      then dempresa="Distribucion y Excelencia"
	if empresa="LACAV"     then dempresa="Distribuidora LA CAV Ltda."
	if empresa="UNDURRAGA" then dempresa="Viña Undurraga"

	%><CENTER>
	<TABLE width='400px'>
	<TR>
		<TD><FONT SIZE="2" COLOR="#6E6E6E"><B><%=dempresa%></B></FONT></TD>
		<TD>
			<table border='2' width='180px' height='60px'  bordercolorlight='#008000' bordercolordark='#008000' bordercolor='#008000'>
			<TR>
				<TD bordercolor='#008000' bordercolorlight='#008000' bordercolordark='#008000'>
				<CENTER>
				<FONT SIZE="1" COLOR="#008000"><%=tipodocto%></FONT><BR><BR>
				<FONT SIZE="2" COLOR="#008000"><B><%=documento%></B></FONT><BR>
				<FONT SIZE="1" COLOR="#C0C0C0"><%=correlativo%></FONT>
				</CENTER>
				</TD>
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan="2"><%
':: Datos cliente
	for x=1 to 9
		a=split(matriz(x),"|")
		if ubound(a)=1 then
			'r.write "<BR>u(a):" & ubound(a)
			'res=a(1)
			res=split(a(1),",")
			'respuesta=ubound(res)
			respuesta=""
			for u=0 to ubound(res)
				resx=split(res(u),".")
				if ucase(trim(resx(1)))="RAZONSOCIAL" then
					'respuesta=respuesta & "&nbsp;&nbsp;&nbsp;<B>" & pc(rs(trim(cisnull(resx(1),""))),0) & "</B>"
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;<B>" & rs(trim(cisnull(resx(1),""))) & "</B>"
				else
					'respuesta=respuesta & "&nbsp;&nbsp;&nbsp;"    & pc(rs(trim(cisnull(resx(1),""))),0)
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;"    & rs(trim(cisnull(resx(1),"")))
				end if
			next

			r.write "<BR><FONT COLOR='#6E6E6E'>" & trim(a(0)) & "&nbsp;:</FONT>" & respuesta
			'r.write "<BR>a(1):" & a(1)
		end if
	next
		'forsar datos
		sql="SELECT numero_pedido " & _
			"FROM   sqlserver.desaerp.dbo.PED_PEDIDOSENC " & _
			"WHERE  ((correlativo_flex = " & correlativo & " or factura_sap = " & correlativo & ") and idempresa=" & idempresa & ") " & _
			"order by fecha_pedido DESC"
		respuesta=consultarapida(sql)
		'response.write(sql)
		respuesta="<A HREF='http://pda.desa.cl/palm/buscanota.asp?np=" & respuesta & "'>" & respuesta & "</A>"
		'respuesta=tipodocto
		if ucase(left(tipodocto,3))="FAC" then
			r.write "<BR><FONT COLOR='#6E6E6E'>Nota&nbsp;Pedido&nbsp;:</FONT>&nbsp;&nbsp;" & respuesta
		end if

		if ucase(left(tipodocto,3))="NOT" then
			'r.write "<BR><FONT COLOR='#6E6E6E'>Nota&nbsp;Pedido&nbsp;:</FONT>&nbsp;&nbsp;" & respuesta
		end if
		%>
		
		</TD>
	</TR>
	</TABLE><BR><%

':: Grilla producto
sql="SELECT PRODUCTO, DESCRIPCION, CANT, PRECIO, VALOR, [%DESC] FROM ( " & _
"SELECT DOC.PRODUCTO, LEFT(P.GLOSA,45) AS DESCRIPCION, DOC.CANTIDAD AS CANT, " & _
"DOC.PRECIO, DOC.NETO AS VALOR, DOC.PORCENTAJEDR AS [%DESC], DOC.EMPRESA, DOC.TIPODOCTO, DOC.CORRELATIVO, DOC.LINEA  " & _
"FROM SERVERDESA.BDFlexline.FLEXLINE.DOCUMENTOD AS DOC INNER JOIN " & _
"SERVERDESA.BDFlexline.FLEXLINE.PRODUCTO AS P ON DOC.EMPRESA = P.EMPRESA AND DOC.PRODUCTO = P.PRODUCTO " & _
"WHERE LINEA >0 " & _
"UNION " & _
"SELECT Producto, LEFT(descripcion, 45) AS DESCRIPCION, cantidad,Precio, Valor, Descuento AS [%DESC], 'DESA' AS Empresa, " & _
"'FACT. AFECTA ELEC' AS TIPODOCTO, CAST([Numero Factura] AS NUMERIC(18,0)) AS CORRELATIVO, linea " & _
"FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURADET " & _
") AS DETALLE " & _
"WHERE (EMPRESA = '" & empresa & "') AND (TIPODOCTO = '" & tipodocto & "') AND (CORRELATIVO='" & correlativo & "') " & _
"ORDER BY LINEA "

	sql=replace(ucase(sql),"BDFlexline",base)
	'r.write sql
	set rs=oConn.execute(sql)
	call tablasimple(rs,"negro")

':: Totales
dim flete, afecto, exento, iva, ila1, ila2, ila3, ila4, ila6
flete=0
afecto=0
Exento=0
iva=0
ila1=0
ila2=0
ila3=0
ila4=0
ila6=0

sql="SELECT DocumentoV.TipoDocto, DocumentoV.Correlativo, DocumentoV.Nombre, DocumentoV.Orden, DocumentoV.Factor, DocumentoV.Monto, DocumentoV.MontoIngreso " & _
"FROM serverdesa.BDFlexline.flexline.Documento Documento, serverdesa.BDFlexline.flexline.DocumentoV DocumentoV " & _
"WHERE DocumentoV.Correlativo = Documento.Correlativo AND DocumentoV.Empresa = Documento.Empresa AND DocumentoV.TipoDocto = Documento.TipoDocto AND ((Documento.Empresa='" & empresa & "') AND (Documento.TipoDocto='" & tipodocto & "') AND (Documento.Numero='" & documento & "'))"

'r.write(documento)

sql=replace(ucase(sql),"BDFlexline",base)
'r.write sql
Set rs=oConn.execute(Sql)
if rs.eof then Set rs=oConn.execute(replace(Sql,"BDFlexline","BDHistorica"))

If rs.eof and rs.bof Then
Set rSAP = Server.CreateObject("ADODB.Recordset")
sSQL = "SELECT Flete, Afecto, Impuesto, Total FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURAENC WHERE [Numero Factura] = CAST('" & documento & "' AS NUMERIC)"
rSAP=oConn.Execute(sSQL)
%>
<br>
  <center>
  <table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse'>
    <tr>
	  <td>Flete:</td>
	  <td align="right"><%= FormatNumber(cdbl(rSAP(0)),0) %></td>
	</tr>
	<tr>
	  <td>Afecto:</td>
  	  <td align="right"><%= FormatNumber(cdbl(rSAP(1)),0) %></td>
	</tr>
	<tr>
	  <td>Impuestos:</td>
	  <td align="right"><%= FormatNumber(cdbl(rSAP(2)),0) %></td>
	</tr>
	<tr>
	  <td>Total:</td>
	  <td align="right"><%= FormatNumber(cdbl(mitotal),0) %></td>
	</tr>
  </table>
  </center>
<br>
<%
Else
do until rs.eof
if rs.fields("Nombre")="FleteTotal" then flete =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="AfectoIVA"  then afecto=FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="Exento"     then exento=FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="IVA"        then iva   =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA1"       then ila1  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA2"       then ila2  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA3"       then ila3  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA4"       then ila4  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA6"       then ila6  =FormatNumber(cdbl(rs.fields("monto")),0)
rs.movenext
loop
response.write("<BR><CENTER><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 style='border-collapse: collapse'>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Flete</FONT  ></TD><TD ALIGN=right>" & flete  & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Afecto</FONT ></TD><TD ALIGN=right>" & afecto & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Exento</FONT ></TD><TD ALIGN=right>" & exento & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>IVA</FONT    ></TD><TD ALIGN=right>" & iva    & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Vinos</FONT  ></TD><TD ALIGN=right>" & ila1   & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Cerv.</FONT  ></TD><TD ALIGN=right>" & ila2   & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Licor</FONT  ></TD><TD ALIGN=right>" & ila3   & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Whisky</FONT ></TD><TD ALIGN=right>" & ila4   & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2>Bebidas</FONT></TD><TD ALIGN=right>" & ila6   & "</TD></TR>" & _
"<TR><TD><FONT face='Arial' SIZE=2><B>Total</FONT></B></TD ><TD ALIGN=right><B>" & mitotal & "</B></TD></TR>" & _
"</TABLE></CENTER>")
End If
%>
<TABLE>
<TR>
	<TD valign="top"><INPUT TYPE="button" value="<< Atras"      onclick="history.back();history.back()"></TD>
	<TD valign="top"><input type="button" value="Imprimir"      onClick="window.print()"></TD>
	<!-- <TD valign="top"><input type="button" value="Propiedades"   onClick=""></TD> -->
	<TD valign="top">
	<FORM METHOD=POST ACTION='/palm/ruta.asp'>
	 <INPUT TYPE='hidden' name='documento' value='<%=documento%>'>
	 <INPUT TYPE='hidden' name='empresa' value='<%=empresa%>'>
	 <INPUT TYPE='hidden' name='tipodocto' value='<%=tipodocto%>'>
	
	 <INPUT TYPE='submit' value='Info Despacho'>
	</FORM>
	</TD>
	<TD valign="top">
	<FORM METHOD=POST ACTION='/palm/referencia.asp'>
	 <INPUT TYPE='hidden' name='correlativo' value='<%=correlativo%>'>
	 <INPUT TYPE='hidden' name='empresa' value='<%=empresa%>'>
	 <INPUT TYPE='hidden' name='tipodocto' value='<%=tipodocto%>'>
	
	 <INPUT TYPE='submit' value='Referencia'>
	</FORM>
	</TD>
	

</TR>
</TABLE>


	</CENTER>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.getElementById("espere").style.display='none';
		//-->
	</SCRIPT><%
End sub 'mostrardocumento()
'---------------------------------------------------------------------------------------------
Private sub tablasimple(rs, Especial)
	on error resume next
	bordercolor="#C0C0C0"
	if Especial="negro" then bordercolor="#000033"
	rs.movefirst
	'set rs=oConn.execute(sql)
	if rs.eof then exit sub
	%><TABLE border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse;' bordercolor='<%=bordercolor%>'><TR><%
	For x=0 to rs.Fields.count-1
		%><TH><%=pc(rs.Fields(x).Name,0)%></TH><%
	next
	%></TR><%
	do until rs.eof
if Especial="color" then
	if micolor="#CCFFFF" then
		micolor="#FFFF99"
	else
		micolor="#CCFFFF"
	end if 
end if
		%><TR bgcolor="<%=micolor%>"><%
		For x=0 to rs.Fields.count-1
			if isnull(rs.Fields(x)) then
				mivalor="&nbsp;"
			else
				mivalor=PC(rs.Fields(x),0)
				if rs.Fields(x).Name="Numero" then 
					'mivalor=replace(replace(mivalor,"<b>",""),"</b>","")
					mivalor="<A HREF='documento.asp?documento=" & rs("numero") & "&empresa=" & rs("empresa") & "&tipodocto=" & rs("tipodocto") & "&base=" & basex & "'><B>" & mivalor & "</B></A>"
				end if
			end if
			%><TD><%=mivalor%></TD><%
		next
		%></TR><%
	rs.movenext
	loop
	%></TABLE><%
End sub 'tablasimple(sql)
'---------------------------------------------------------------------------------------------
function cisnull(yvalor,xvalor)
	if isnull(yvalor) then
		cisnull=xvalor
	else
		cisnull=yvalor
	end if
end function
'---------------------------------------------------------------------------------------------
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function borraespere(){
		//document.getElementById("espere").style.display='none'
	}
//-->
</SCRIPT>

</BODY>
</HTML>

<!--#include Virtual="/includes/conexion_SAP.asp"-->
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

Set oConn2 = server.createobject("ADODB.Connection")
oConn2.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=HANDHELD;User Id=sa;Password=desakey;"

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
	idempresa= request.querystring("idempresa"   )

	if idempresa=1 then empresa="DESA"
	if idempresa=4 then empresa="LACAV"
	
	if correlati="0" then
		documento="Aun No existe Documento"
		%>
		<CENTER>
			Aun No existe Documento
		</CENTER>
		<%		
	end if
	if empresa="DESA" or idempresa=1 then
		if len(correlati)>1 then
			sql="select docnum from HIS_DESA.DESA_DOCENC where  empresa='"& empresa &"' and correlativo=" & correlati
			Set rs=oConn.execute(sql)
			documento=rs(0)
		end if
	end if

	if empresa="LACAV" or idempresa=4 then 
		if len(correlati)>1 then
			sql="select numero from serverdesa.BDFlexline.flexline.documento as d where tipodocto in ('facturas palm','facturas movil','fact. afecta elec') and empresa='" &  empresa & "' and correlativo=" & correlati & "and fecha>='20160301'"
			Set rs2=oConn2.execute(sql)
			documento=rs2(0)
		end if
	end if

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
'---------------------------------------------------------------------------------------------
private sub buscasimple()
	%>
	<CENTER>
	<FONT SIZE="2" COLOR="#000066">Busca documentos por su numero</FONT>
	<BR>
	<FORM METHOD=POST ACTION="">
		<INPUT TYPE="text" NAME="factura" size="10">
		<INPUT TYPE="submit" value="Buscar">
		<INPUT TYPE="hidden" name="nuser" value="<%=nuser%>">
	</FORM>
	<HR>
	<FONT SIZE="2" COLOR="#000066">Busca documentos por Nombre o Rut</FONT>
	<FORM METHOD=POST ACTION="">
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
	if empresa="DESA" or idempresa=1 then
		sql="SELECT TOP "& mitop &" RIGHT('00000'||CAST(DOCNUM AS VARCHAR),10) AS Numero, TIPODOCUMENTO AS Tipodocto, RAZONSOCIAL AS Nombre, " & _ 
		"FECHAEMISION AS Fecha, CTACTE AS Cliente, EMPRESA, SIGLA, VENDEDORCLIENTE AS Ejecutivo, FECHAOC FROM ""HIS_DESA"".""DESA_DOCENC"" WHERE  (""DOCNUM""!=0) AND "
		if len(documento) then
			sql=sql & """DOCNUM""='" & documento & "'"
			sql=sql & " or (""TIPODOCUMENTO"" ='GUIA DESPACHO ELEC' and ""DOCNUM""='" & documento & "')"
		else
			if buspor="rut" then
				sql=sql & """CTACTE"" like '%" & mibusca & "%'"
			else
				sql=sql & """RAZONSOCIAL""||""SIGLA"" like '%" & ucase(mibusca) & "%'"
			end if
		end if
		sql=sql & " order by fechaemision desc, docnum desc "
		set rs=oConn.execute(sql)
	end if
	
if empresa="LACAV" or idempresa=4 then
sql="select top " & mitop & " documento.numero as Numero, " & _
	"documento.tipodocto, ctacte.razonsocial as Nombre, documento.fecha, documento.cliente, documento.vigencia as V, documento.empresa " & _
	"from serverdesa.BDFlexline.flexline.documento as documento " & _
	"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte " & _
	"on ctacte.empresa=documento.empresa and ctacte.ctacte=documento.cliente and ctacte.tipoctacte=documento.tipoctacte " & _
	"inner join serverdesa.BDFlexline.flexline.tipodocumento as tipodocumento " & _
	"on tipodocumento.empresa=documento.empresa and tipodocumento.tipodocto=documento.tipodocto " & _
	"where documento.empresa in ('" & empresa & "') and documento.Fecha>='20160301' and tipodocumento.sistema='VENTAS' " & _
	"and tipodocumento.clase = 'Factura (v)' and documento.factormonto<>0 and "
	if len(documento) then
		sql=sql & "documento.numero='" & documento & "'"
		sql=sql & " or (tipodocumento.tipodocto ='GUIA DESPACHO ELEC' and documento.numero='" & documento & "')"
	else
		if buspor="rut" then
			sql=sql & "documento.cliente like '%" & mibusca & "%'"
		else
			sql=sql & "ctacte.razonsocial + ctacte.sigla like '%" & mibusca & "%'"
		end if
	end if
	sql=sql & " order by documento.fecha desc, documento.numero desc "

	set rs2=oConn2.execute(sql)
end if	
	'response.write(sql)
if err<>0 then
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
if empresa="DESA" THEN
	do until rs.eof
		empresa  =rs("EMPRESA"  )
		tipodocto=rs("TIPODOCUMENTO")
		doctos   =doctos+1
		documento=rs("DOCNUM")
	rs.movenext
	loop
	
	call tablasimple(rs,"color")
	r.flush()

	set rs=oConn.execute(sql)

end if

if empresa="LACAV" then
do until rs2.eof
		empresa  =rs2("empresa"  )
		tipodocto=rs2("tipodocto")
		doctos   =doctos+1
		basex="BDFlexline"
		documento=rs2("numero")
	rs2.movenext
	loop
	
	call tablasimple(rs2,"color")
	r.flush()
	
	set rs2=oConn2.execute(sql)
	
end if

	if doctos=0 then
		%><BR><BR>La busqueda no obtuvo resultados<%
		if len(nombrevendedor)>0 then 
			response.write("<BR>Para el Vendedor " & nombrevendedor & "<BR>")
		end if 
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
					<OPTION value="2000">Todas</OPTION>
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

	</SCRIPT></CENTER>
	</FORM><%
	if err<>0 then
		
	end if
End sub 'ListaDocumentos()
'---------------------------------------------------------------------------------------------
Private sub mostrardocumento()
	%><CENTER><div id="espere"><B>Espere&nbsp;mientras&nbsp;carga&nbsp;...&nbsp;</B></div><%
	dim matriz(11)
if empresa="DESA" THEN	

	matriz( 1)="Nombre. .   |RAZONSOCIAL"
	matriz( 2)="id Cliente  |CTACTE, SIGLA, FECHAEMISION"
	matriz( 3)="Direccion   |DIRECCIONDESPACHO , COMUNA"
	matriz( 4)="Cond Pago   |CONDICIONPAGO"
	matriz( 5)="Vend Cliente|VENDEDORCLIENTE"
	matriz( 6)="Vend Factura|VENDEDORFACTURA"
	matriz( 7)="Local-bodega|SUCURSAL, BODEGA"
	matriz( 8)="Referencia  |REFERENCIA"
	matriz( 9)="Fecha OC    |FECHAOC"
	miselect=""
	for x=1 to 10
		a=split(matriz(x),"|")
		if ubound(a)=1 then
			'r.write "<BR>u(a):" & ubound(a)
	'		r.write "<BR>a(0):" & a(0)
	'		r.write "<BR>a(1):" & a(1)
			miselect=miselect & ", " & a(1)
		end if
	next
	
	miselect= "SELECT VIGENCIA, TOTAL, CORRELATIVO " & miselect
	miselect= miselect + ", EMPRESA, RIGHT('00000'||CAST(DOCNUM AS VARCHAR),10) AS NUMERO, TIPODOCUMENTO"

	sqldatos=" FROM ""HIS_DESA"".""DESA_DOCENC"" " & _
		"WHERE EMPRESA='" & empresa & "' AND DOCNUM='" & documento & "' and TIPODOCUMENTO='" & tipodocto & "'"

	sql=miselect & sqldatos

	set rs=oConn.execute(sql)

	if rs("Vigencia")="A" then 
		%><FONT COLOR="#CC0000"><B>Documento NULO</B></FONT><%
	End if
	
	if rs("Vigencia")="N" then 
		%><FONT COLOR="#CC0000"><B>Documento No Vigente</B></FONT><%
	End if
	mitotal= FormatNumber(cdbl(rs.fields("TOTAL")),0)
	correlativo=rs.fields("CORRELATIVO")
end if


if empresa="LACAV" THEN	
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
		miselect=miselect & ", " & a(1)
	end if
	next

	miselect= "Select documento.vigencia, documento.total, documento.correlativo " & miselect
	miselect=replace(miselect,"ctacte.Direccionenvio","handheld.flexline.ctactedireccion(documento.cliente) as Direccionenvio")

	sqldatos=" from serverdesa.BDFlexline.flexline.documento as Documento " & _
	"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte " & _
	"on ctacte.empresa=documento.empresa and ctacte.tipoctacte=documento.tipoctacte and documento.cliente=ctacte.ctacte " & _
	"inner join serverdesa.BDFlexline.flexline.documentop documentop " & _
	"on documentop.empresa=documento.empresa and documento.tipodocto=documentop.tipodocto and documento.correlativo=documentop.correlativo " & _
	"where documento.empresa='" & empresa & "' and documento.numero='" & documento & "' and Documento.TipoDocto='" & tipodocto & "'"
	
	if tipodocto="GUIA DESPACHO ELEC" then 
	sqldatos=" from serverdesa.BDFlexline.flexline.documento as Documento " & _
	"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte " & _
	"on ctacte.empresa=documento.empresa and ctacte.tipoctacte=documento.tipoctacte and documento.cliente=ctacte.ctacte " & _
	"inner join (select 'desa' as empresa, '" & tipodocto & "' as tipodocto, '" & documento & "' as correlativo, '' as codigopago) documentop " & _
	"on documentop.empresa=documento.empresa and documento.tipodocto=documentop.tipodocto and documento.correlativo=documentop.correlativo " & _
	"where documento.empresa='" & empresa & "' and documento.numero='" & documento & "' and Documento.TipoDocto='" & tipodocto & "'"
	end if

	sql=miselect & sqldatos
	sql=replace(ucase(sql),"BDFlexline",base)
	'response.write(sql)
	
	set rs2=oConn2.execute(sql)

	'if rs.eof then exit sub
	if rs2("Vigencia")="A" then 
		%><FONT COLOR="#CC0000"><B>Documento NULO</B></FONT><%
	End if
	
	if rs2("Vigencia")="N" then 
		%><FONT COLOR="#CC0000"><B>Documento No Vigente</B></FONT><%
	End if
	mitotal= FormatNumber(cdbl(rs2.fields("total")),0)
	correlativo=rs2.fields("correlativo")
	
end if

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
if empresa="DESA" THEN 
	for x=1 to 9
		a=split(matriz(x),"|")
		if ubound(a)=1 then
			res=split(a(1),",")
			respuesta=""
			for u=0 to ubound(res)
				if ucase(trim(res(u)))="RAZONSOCIAL" then
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;<B>" & pc(rs(trim(cisnull(res(u),""))),0) & "</B>"
				else
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;"    & rs(trim(cisnull(res(u),"")))
				end if
			next

			r.write "<BR><FONT COLOR='#6E6E6E'>" & trim(a(0)) & "&nbsp;:</FONT>" & respuesta
		end if
	next
		'forsar datos
		sql="SELECT NOTAPEDIDO " & _
			"FROM   ""HIS_DESA"".""DESA_DOCENC"" " & _
			"WHERE  ((CORRELATIVO = " & correlativo & " ) AND TIPODOCUMENTO= '" & tipodocto & "' ) "
		respuesta=consultarapida(sql)
		respuesta="<A HREF='http://pda.desa.cl/palm/buscanota.asp?np=" & respuesta & "'>" & respuesta & "</A>"
		if ucase(left(tipodocto,3))="FAC" then
			r.write "<BR><FONT COLOR='#6E6E6E'>Nota&nbsp;Pedido&nbsp;:</FONT>&nbsp;&nbsp;" & respuesta
		end if
		if ucase(left(tipodocto,3))="NOT" then
		end if

		%>
		
		</TD>
	</TR>
	</TABLE><BR><%

':: Grilla producto
sql="SELECT PRODUCTO, DESCRIPCIN AS GLOSA, CANTIDAD, PRECIO, NETO, DESCUENTO FROM ""HIS_DESA"".""DESA_DOCDET"" " & _
"WHERE (EMPRESA = '" & empresa & "') AND (TIPODOCUMENTO = '" & tipodocto & "') AND (CORRELATIVO='" & correlativo & "') " & _
"ORDER BY LINEA "

	set rs=oConn.execute(sql)
	call tablasimple(rs,"negro")

':: Totales

Set rSAP = Server.CreateObject("ADODB.Recordset")
sSQL = "SELECT FLETE, AFECTO, IMPUESTO FROM ""HIS_DESA"".""DESA_DOCENC"" WHERE DOCNUM = '" & documento & "'"
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
<TABLE>
<TR>
	<TD valign="top"><INPUT TYPE="button" value="<< Atras"      onclick="history.back();history.back()"></TD>
	<TD valign="top"><input type="button" value="Imprimir"      onClick="window.print()"></TD>
	<TD valign="top">
	<FORM METHOD=POST ACTION='/palm/ruta.asp'>
	 <INPUT TYPE='hidden' name='documento' value='<%=documento%>'>
	 <INPUT TYPE='hidden' name='empresa' value='<%=empresa%>'>
	 <INPUT TYPE='hidden' name='tipodocto' value='<%=tipodocto%>'>
	
	 <INPUT TYPE='submit' value='Info Despacho'>
	</FORM>
	</TD>
	<TD valign="top">
	
	</TD>
	

</TR>
</TABLE>


	</CENTER>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.getElementById("espere").style.display='none';
		//-->
	</SCRIPT><%
end if
if empresa="LACAV" THEN
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
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;<B>" & pc(rs2(trim(cisnull(resx(1),""))),0) & "</B>"
				else
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;"    & pc(rs2(trim(cisnull(resx(1),""))),0)
				end if
			next

			r.write "<BR><FONT COLOR='#6E6E6E'>" & trim(a(0)) & "&nbsp;:</FONT>" & respuesta
			'r.write "<BR>a(1):" & a(1)
		end if
	next
		'forsar datos
		sql="SELECT numero_pedido " & _
			"FROM   sqlserver.desaerp.dbo.PED_PEDIDOSENC " & _
			"WHERE  (correlativo_flex = " & correlativo & " and idempresa=" & idempresa & ") " & _
			"order by fecha_pedido DESC"
		set rs3=oConn2.execute(sql)
		if not rs3.eof then respuesta  = rs3.fields(0)
	
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
sql="SELECT  doc.Producto, left(p.GLOSA,45) AS Descripcion, doc.Cantidad as Cant, doc.Precio, doc.Neto AS Valor, doc.PorcentajeDR AS [%Desc] " & _
"FROM SERVERDESA.BDFlexline.flexline.documentoD AS doc INNER JOIN " & _
"SERVERDESA.BDFlexline.flexline.producto AS p ON doc.Empresa = p.EMPRESA AND doc.Producto = p.PRODUCTO " & _
"WHERE (doc.Empresa = '" & empresa & "') AND (doc.TipoDocto = '" & tipodocto & "') AND (doc.correlativo='" & correlativo & "') " & _
"and linea >0 " & _
"order by doc.linea"
	sql=replace(ucase(sql),"BDFlexline",base)
	'p.write sql
	set rs2=oConn2.execute(sql)
	call tablasimple(rs2,"negro")

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

sql=replace(ucase(sql),"BDFlexline",base)
'r.write sql
Set rs2=oConn2.execute(Sql)
if rs2.eof then Set rs2=oConn2.execute(replace(Sql,"BDFlexline","BDHistorica"))

do until rs2.eof
if rs2.fields("Nombre")="FleteTotal" then flete =FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="AfectoIVA"  then afecto=FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="Exento"     then exento=FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="IVA"        then iva   =FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="ILA1"       then ila1  =FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="ILA2"       then ila2  =FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="ILA3"       then ila3  =FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="ILA4"       then ila4  =FormatNumber(cdbl(rs2.fields("monto")),0)
if rs2.fields("Nombre")="ILA6"       then ila6  =FormatNumber(cdbl(rs2.fields("monto")),0)
rs2.movenext
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
%>

<TABLE>
<TR>
	<TD valign="top"><INPUT TYPE="button" value="<< Atras"      onclick="history.back();history.back()"></TD>
	<TD valign="top"><input type="button" value="Imprimir"      onClick="window.print()"></TD>
	<!-- <TD valign="top"><input type="button" value="Propiedades"   onClick=""></TD> -->
	<TD valign="top">
	<FORM METHOD=POST ACTION='/palm/ruta.flex.asp'>
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
END IF

	End sub 'mostrardocumento()
'---------------------------------------------------------------------------------------------
Private sub tablasimple(rs, Especial)
	on error resume next
	bordercolor="#C0C0C0"
	if Especial="negro" then bordercolor="#000033"
	rs.movefirst
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
				if rs.Fields(x).Name="NUMERO" then 
					mivalor="<A HREF='?documento=" & rs("numero") & "&empresa=" & rs("empresa") & "&tipodocto=" & rs("tipodocto") & "'><B>" & mivalor & "</B></A>"
				end if
				if rs.Fields(x).Name="Numero" then 
					mivalor="<A HREF='?documento=" & rs("numero") & "&empresa=" & rs("empresa") & "&tipodocto=" & rs("tipodocto") & "&base=" & basex & "'><B>" & mivalor & "</B></A>"
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
	}
//-->
</SCRIPT>

</BODY>
</HTML>

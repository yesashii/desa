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
oConn2.ConnectionTimeOut = 0
oConn2.CommandTimeout = 0
oConn2.open "Provider=SQLOLEDB;Data Source=SERVERDESA;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

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
			sql="select docnum from HIS_DESA.DESA_DOCENC where  empresa='"& empresa &"' and correlativo=" & correlati
			Set rs=oConn.execute(sql)
			documento=rs(0)
			'response.write(documento)
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
		
	'response.write sql

'	if len(nuser)>0 then
'		SW_CONSULTA=consultarapida("select SW_CONSULTA from handheld.dbo.dim_vendedores where idEMPRESA='1' AND idvendedor=" & cint(nuser) )
'		nombrevendedor=consultarapida("select NOMBRE   from handheld.dbo.dim_vendedores where idEMPRESA='1' AND idvendedor=" & cint(nuser))
'		'RESPONSE.WRITE ("<BR>nuser : " &  nuser)
'		'RESPONSE.WRITE ("<BR>sw_consulta : " &  SW_CONSULTA)
'		'RESPONSE.WRITE ("<BR>NombreVendedor : " &  nombrevendedor)
'		'sql=sql & " and ctacte.ejecutivo='" & nombrevendedor & "' "
'		IF UCASE(SW_CONSULTA)="N" then
'			sql=sql & " and ""vendedorcliente""='" & nombrevendedor & "' "
'		end if
'	end if
	'r.write(sql)
	'call tablasimple(sql,"color")
	sql=sql & " order by fechaemision desc, docnum desc "
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
		empresa  =rs("EMPRESA"  )
		tipodocto=rs("TIPODOCUMENTO")
		doctos   =doctos+1
		'basex="BDFlexline"
		documento=rs("DOCNUM")
	'	doctox=doctox+1
	rs.movenext
	loop
	
	call tablasimple(rs,"color")
	r.flush()

	'r.Write("<HR><B><FONT SIZE="2" COLOR="#808080">Base Historica</FONT></B>") 'Activar Para Base Histórica
	'sql=replace(sql,"BDFlexline","BDHistorica")
	'r.write(sql)
	set rs=oConn.execute(sql)
	'doctox=0
'	do until rs.eof
'		empresa  =rs("empresa"  )
'		tipodocto=rs("tipodocto")
'		doctos   =doctos+1
'		basex="BDHistorica"
'		documento=rs("docnum")
'	'	doctox=doctox+1
'	rs.movenext
'	loop
	
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

		//if (doctos==1 ){window.open('?empresa='+empresa+'&documento='+documento+'&tipodocto='+tipodocto+'&nocache='+Math.random() ,'_self')}
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
	dim matriz(11)
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
	
'	if tipodocto="GUIA DESPACHO ELEC" then 
'	sqldatos=" from serverdesa.BDFlexline.flexline.documento as Documento " & _
'	"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte " & _
'	"on ctacte.empresa=documento.empresa and ctacte.tipoctacte=documento.tipoctacte and documento.cliente=ctacte.ctacte " & _
'	"inner join (select 'desa' as empresa, '" & tipodocto & "' as tipodocto, '" & documento & "' as correlativo, '' as codigopago) documentop " & _
'	"on documentop.empresa=documento.empresa and documento.tipodocto=documentop.tipodocto and documento.correlativo=documentop.correlativo " & _
'		"UNION " & _
'"SELECT CASE WHEN Vigencia = 'N' THEN 'S' ELSE 'N' END AS VIGENCIA, " & _
'"TOTAL, CAST([Numero factura] AS NUMERIC(18, 0)) AS CORRELATIVO, Nombre AS RAZONSOCIAL, " & _
'"CTACTE, SIGLA, [FECHA EMISION] AS FECHA, [Direccion Despacho] AS DIRECCIONENVIO, 'SIN DATO' AS COMUNA, " & _
'"[Condicion de Pago] AS CODIGOPAGO, [Vendedor Cliente] AS EJECUTIVO, [Vendedor Factura] AS VENDEDOR, " & _
'"Sucursal AS [LOCAL], BODEGA, referencia AS REFERENCIAEXTERNA, 'DESA' AS EMPRESA,  " & _
'"RIGHT('0000' + [Numero Factura],10) AS NUMERO, 'FACT. AFECTA ELEC' AS TIPODOCTO " & _
'"FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURAENC  " & _
'")ENCAB " & _
'	"where empresa='" & empresa & "' and numero='" & documento & "' and TipoDocto='" & tipodocto & "'"
'	end if

	sql=miselect & sqldatos

'	sql=replace(ucase(sql),"BDFlexline",base)
'r.write(sql)
	set rs=oConn.execute(sql)

	'if rs.eof then exit sub
	if rs("Vigencia")="A" then 
		%><FONT COLOR="#CC0000"><B>Documento NULO</B></FONT><%
	End if
	
	if rs("Vigencia")="N" then 
		%><FONT COLOR="#CC0000"><B>Documento No Vigente</B></FONT><%
	End if
	mitotal= FormatNumber(cdbl(rs.fields("TOTAL")),0)
'	sapFlete =
'	sapImpuesto =
	correlativo=rs.fields("CORRELATIVO")

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
				'resx=split(res(u),".")
				'response.Write(res(u))
				if ucase(trim(res(u)))="RAZONSOCIAL" then
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;<B>" & pc(rs(trim(cisnull(res(u),""))),0) & "</B>"
				else
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;"    & rs(trim(cisnull(res(u),"")))
				end if
			next

			r.write "<BR><FONT COLOR='#6E6E6E'>" & trim(a(0)) & "&nbsp;:</FONT>" & respuesta
			'r.write "<BR>a(1):" & a(1)
		end if
	next
		'forsar datos
		sql="SELECT NOTAPEDIDO " & _
			"FROM   ""HIS_DESA"".""DESA_DOCENC"" " & _
			"WHERE  ((CORRELATIVO = " & correlativo & " ) ) "
		respuesta=consultarapida(sql)
'		'response.write(sql)
		respuesta="<A HREF='http://pda.desa.cl/palm/buscanota.asp?np=" & respuesta & "'>" & respuesta & "</A>"
'		'respuesta=tipodocto
		if ucase(left(tipodocto,3))="FAC" then
			r.write "<BR><FONT COLOR='#6E6E6E'>Nota&nbsp;Pedido&nbsp;:</FONT>&nbsp;&nbsp;" & respuesta
		end if
'
		if ucase(left(tipodocto,3))="NOT" then
			'r.write "<BR><FONT COLOR='#6E6E6E'>Nota&nbsp;Pedido&nbsp;:</FONT>&nbsp;&nbsp;" & respuesta
		end if
		%>
		
		</TD>
	</TR>
	</TABLE><BR><%

':: Grilla producto
sql="SELECT PRODUCTO, DESCRIPCIN AS GLOSA, CANTIDAD, PRECIO, NETO, DESCUENTO FROM ""HIS_DESA"".""DESA_DOCDET"" " & _
"WHERE (EMPRESA = '" & empresa & "') AND (TIPODOCUMENTO = '" & tipodocto & "') AND (CORRELATIVO='" & correlativo & "') " & _
"ORDER BY LINEA "

'	sql=replace(ucase(sql),"BDFlexline",base)
	'r.write sql
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
<!--	<FORM METHOD=POST ACTION='/palm/referencia.asp'>
	 <INPUT TYPE='hidden' name='correlativo' value='<%=correlativo%>'>
	 <INPUT TYPE='hidden' name='empresa' value='<%=empresa%>'>
	 <INPUT TYPE='hidden' name='tipodocto' value='<%=tipodocto%>'>
	
	 <INPUT TYPE='submit' value='Referencia'>
	</FORM> -->
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
				'response.write(rs(x).Name)
				if rs.Fields(x).Name="NUMERO" then 
					'mivalor=replace(replace(mivalor,"<b>",""),"</b>","")
					mivalor="<A HREF='?documento=" & rs("numero") & "&empresa=" & rs("empresa") & "&tipodocto=" & rs("tipodocto") & "'><B>" & mivalor & "</B></A>"
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

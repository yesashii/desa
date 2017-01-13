<html>
<head>
<title>Ingreso Productos</title>
</head>
<body topmargin="0">
<% 'on error resume next
Private tipodoc, mitotal, OBS, oConn, rs, sql, rs1
public p01, c01, d01, p02, c02, d02, p03, c03, d03, p04, c04, d04, p05, c05, d05, p06, c06, d06
public p07, c07, d07, p08, c08, d08, p09, c09, d09, p10, c10, d10, p11, c11, d11, p12, c12, d12
public dia, descuentocliente, fechaentrega, bgcolor, id_Nro_ck, OC
dim afeciva, toila, totfact, toflete, toneto, apago, iva, ila1, ila2, ila3, ila4, ila5, ila6	

Set oConn = server.createobject("ADODB.Connection")
Set rsl = Server.CreateObject("ADODB.Recordset")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0
oConn.open "Provider=SQLOLEDB;Data Source=sqlserver;Initial Catalog=handheld;User Id=sa;Password=desakey;"

vend = ucase(trim(request.querystring("vend")))
vend=replace(vend,"?","Ñ")'Error blackberry
'vend=replace(vend,"NN","Ñ")'Error blackberry

'if left(ucase(vend),15)="JOSE LUIS CARRE" then vend="JOSE LUIS CARREÑO"
cliente = trim(request.querystring("cliente"))
PN = trim(request.querystring("PN"))
fechaentrega=trim(request.querystring("fechaentrega"))
id_porfolio=trim(request.QueryString("id_porfolio"))
if len(fechaentrega)=0 then fechaentrega=request.form("fechaentrega2")

misql="select idvendedor,idlocal,pedido_minimo  from desaerp.dbo.DIM_VENDEDORES where nombre='" & vend & "'"
	set rs1=oConn.execute(misql)

	if not rs1.eof then 
		nuser  = rs1.fields(0)
		idlocal= rs1.fields(1)
		pedmin=rs1.fields(2)
		'response.write("<BR>nuser : " & nuser)
		'response.write("<BR>idlocal : " & idlocal)
	end if

'response.write("<FONT SIZE='1' COLOR='#C0C0C0'>fecha entrega : " & fechaentrega & "</FONT>")

'empresa="DESA"
empresa = trim(request.querystring("empresa"))
sw_iva		= request.QueryString("sw_iva")
tipodocto   = request.QueryString("tipodocto")
'response.write("<BR>empresa : " & empresa)
if len(empresa)=0 then empresa="DESA"
'response.write("<BR>empresa : " & empresa)

if len(id_porfolio) > 5 then
	if id_porfolio="porfolio1" then empresa="DESA"
	if id_porfolio="porfolio2" then empresa="LACAV"
end if

if empresa="DESA" then idempresa="1"
if empresa="LACAV" then idempresa="4"
if empresa="DESAZOFRI" then idempresa="3"

'response.write("<BR>empresa : " & empresa)
'SQL="select * from handheld.flexline.fx_vende_man where nombre='" & vend & "'"
'Set rs=oConn.execute(Sql)
'if not rs.eof then
'	empresa=rs.fields("empresa")
'end if

idmegacanal=""
'sw_ocd=0
sql="select idmegacanal , idcanal " & _
"from handheld.Flexline.CtaCte as cl inner join desaerp.dbo.DIM_CANALES as cn  " & _
"	on cast(rtrim(ltrim(left(isnull(cl.analisisctacte1,0),2))) as numeric)=idcanal  " & _
"where empresa='" & empresa & "' and ctacte='" & cliente & "'"
'response.write(sql)
set rsp=oConn.execute(sql)
if not rsp.eof then idmegacanal=rsp(0)
if not rsp.eof then idcanal=rsp(1)


SQL="SELECT C.CodLegal, C.RazonSocial, C.ListaPrecio, D.Direccion, C.CtaCte, C.Giro, " & _
	"C.CondPago, C.Comuna, C.Telefono, C.LimiteCredito, C.PorcDr1, c.tipo, c.analisisctacte12 " & _
	"FROM flexline.CtaCte C INNER JOIN flexline.CtaCteDirecciones d " & _
	"ON C.CtaCte = d.CtaCte AND C.Empresa = d.Empresa AND c.TipoCtaCte = d.TipoCtaCte " & _
	"WHERE     (C.Empresa = '" & empresa & "') AND (C.TipoCtaCte = 'CLIENTE') AND (d.CtaCte = '" & cliente & "') AND (d.Principal <> 's')"
	
Set rs=oConn.execute(Sql)
if rs.eof then
	if id_porfolio="porfolio2" then 
		response.write("<BR><B>Falta Crear Cliente en Maestras de la CAV</B><BR><BR>")
	else
		response.write("<BR><B>Datos inconsistentes en cliente</B><BR><BR>")
	end if
	response.write("<FONT SIZE='1' face='arial' COLOR='#808080'>" & sql & "</FONT>")
end if
ListPrec = rs.fields("ListaPrecio")
ctactetipo=rs.fields("tipo")
idtipocliente = rs.fields("analisisctacte12")
if len(idtipocliente)=0 then idtipocliente="02"

miboton =trim(cstr(request.form("boton" )))
micantidad=trim(cstr(request.form("cantidad" )))
midescuent=trim(cstr(request.form("descuento" )))
midescuent=replace(midescuent,".",",")
opobs=trim(cstr(request.form("optobs" )))
OBS=trim(cstr(request.form("obs" )))
'1
p01 = trim(request.querystring("p01"))
c01 = trim(request.querystring("c01"))
d01 = trim(request.querystring("d01"))
'2
p02 = trim(request.querystring("p02"))
c02 = trim(request.querystring("c02"))
d02 = trim(request.querystring("d02"))
'3
p03 = trim(request.querystring("p03"))
c03 = trim(request.querystring("c03"))
d03 = trim(request.querystring("d03"))
'4
p04 = trim(request.querystring("p04"))
c04 = trim(request.querystring("c04"))
d04 = trim(request.querystring("d04"))
'5
p05 = trim(request.querystring("p05"))
c05 = trim(request.querystring("c05"))
d05 = trim(request.querystring("d05"))
'6
p06 = trim(request.querystring("p06"))
c06 = trim(request.querystring("c06"))
d06 = trim(request.querystring("d06"))
'7
p07 = trim(request.querystring("p07"))
c07 = trim(request.querystring("c07"))
d07 = trim(request.querystring("d07"))
'8
p08 = trim(request.querystring("p08"))
c08 = trim(request.querystring("c08"))
d08 = trim(request.querystring("d08"))
'9
p09 = trim(request.querystring("p09"))
c09 = trim(request.querystring("c09"))
d09 = trim(request.querystring("d09"))
'10
p10 = trim(request.querystring("p10"))
c10 = trim(request.querystring("c10"))
d10 = trim(request.querystring("d10"))
'11
p11 = trim(request.querystring("p11"))
c11 = trim(request.querystring("c11"))
d11 = trim(request.querystring("d11"))
'12
p12 = trim(request.querystring("p12"))
c12 = trim(request.querystring("c12"))
d12 = trim(request.querystring("d12"))

if len(p01)=0 then CuentaProductos= 0
if len(p01)>0 then CuentaProductos= 1
if len(p02)>0 then CuentaProductos= 2
if len(p03)>0 then CuentaProductos= 3
if len(p04)>0 then CuentaProductos= 4
if len(p05)>0 then CuentaProductos= 5
if len(p06)>0 then CuentaProductos= 6
if len(p07)>0 then CuentaProductos= 7
if len(p08)>0 then CuentaProductos= 8
if len(p09)>0 then CuentaProductos= 9
if len(p10)>0 then CuentaProductos=10
if len(p11)>0 then CuentaProductos=11
if len(p12)>0 then CuentaProductos=12

':::::::::::::::::: conexion :::::::::::::::::
Set oConn = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=sqlserver;Initial Catalog=handheld;User Id=sa;Password=desakey;"

%><TABLE bgcolor="#666633" width="100%">
<TR>
	<Th align="left"  ><FONT SIZE="2" face="verdana" COLOR="#FFFFFF"><%=nuser%></FONT></Th>
	<Th align="center"><FONT SIZE="2" face="verdana" COLOR="#FFFFFF"><%=pc(vend,0)%></FONT></Th>
	<Th align="right" ><FONT SIZE="2" face="verdana" COLOR="#FFFFFF"><A HREF="../default.asp?nuser=<%=nuser%>&p=ok" style='text-decoration:none; color:#FFFFFF'>[Inicio]</A></FONT></Th>
</TR>
</TABLE><%

if len(trim(cstr(request.form("familia" ))))<>0 then ' Si elijo familia en producto
	miboton="Agregar"
	nFamilia=trim(cstr(request.form("familia" )))
	'if len(nFamilia)=0 then nFamilia=request.cookies("NFamilia")
	Response.cookies("NFamilia")=nFamilia
	Response.cookies("NFamilia").expires=date+1
end if

if len(trim(cstr(request.form("buscaprod" ))))<>0 then ' Si elijo familia en producto
	miboton="Agregar"
end if


if len(miboton) =0 then 
	if len(PN)>0 then
		if len(micantidad)>0 then
			sumaprod()
		else
			prodetalle()
		end if
	else
		mustraventa()
	end if
end if

if miboton="Agregar"     then Agregarprod()
if miboton="Agregarmult" then Agregarmultprod()
if miboton="Edit"        then editaprod()
if miboton="Borrar"      then sacaprod()
if miboton="Finalizar"   then Finalizar()

'response.write("<BR>miboton : " & miboton)
'call Agregarmultprod()

Response.cookies("fx_pedidorep")=""
'---------------------------------------------------------------------------------------------------------
sub Agregarmultprod()
	
	id_Nro_ck=request.Form("Nro_ck")
	if len(id_Nro_ck)=0 then id_Nro_ck=0
	CuentaProdsel=0
	validarepetido=request.Form("sw_validarepetido")
	if len(validarepetido)=0 then validarepetido=1
	msgerror=""
	for each campo in request.Form
		if left(campo,3)="ck_" then 
			CuentaProdsel=CuentaProdsel+1
			productoact=right(campo,7)
			if validarepetido=1 then
				if p01=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p02=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p03=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p04=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p05=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p06=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p07=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p08=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p09=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p10=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p11=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"
				if p12=productoact then msgerror=msgerror & "&nbsp;&nbsp;Producto " & productoact & " Repetido &nbsp;&nbsp;<BR>"

				
			end if
		end if
		
	next

'if si ya termine
if (cint(id_Nro_ck)+1) > cint(CuentaProdsel) and id_Nro_ck > 0 then
	'response.write("<BR>id_Nro_ck: " & id_Nro_ck) 
	if cint(id_Nro_ck)=cint(CuentaProdsel) then sumaprod()
	mustraventa()
	exit sub
end if

'response.write("<BR>CuentaProdsel : " & CuentaProdsel)
'exit sub
'muestra productos ya seleccionados	
		if len(msgerror)>0 then
			%>
				<CENTER>
				<TABLE border="0" style='border-collapse: collapse'>
				<TR>
					<TD bgcolor="#666633"><FONT SIZE="1" face="verdana" COLOR="#FFFFFF"><B>Info</B></FONT></TD>
				</TR>
				<TR>
					<TD bgcolor="#C0C0C0"><BR><FONT SIZE="2" face="verdana"  COLOR=""><%=msgerror%></FONT><BR></TD>
				</TR>
				</TABLE>
				</CENTER>
			<%
			call  Agregarprod()
			exit sub
		end if
'response.write("<BR>CuentaProdsel : " & CuentaProdsel)
	if CuentaProdsel=0 then
		%>
		<CENTER>
		<TABLE border="0" style='border-collapse: collapse'>
		<TR>
			<TD bgcolor="#666633"><FONT SIZE="1" face="verdana" COLOR="#FFFFFF"><B>Info</B></FONT></TD>
		</TR>
		<TR>
			<TD bgcolor="#C0C0C0"><BR>&nbsp;&nbsp;<FONT SIZE="2" face="verdana"  COLOR="">No selecciono ningun producto</FONT>&nbsp;&nbsp;<BR><BR></TD>
		</TR>
		</TABLE>
		</CENTER>
		<%
		call  Agregarprod()
		exit sub
	end if
	%><TABLE>
	<TR>
		<TD bgcolor="#666633">&nbsp;&nbsp;<FONT SIZE="2" face="verdana" COLOR="#FFFFFF">Seleccion Multiple <B><%=cint(id_Nro_ck)+1%></B> de <B><%=CuentaProdsel%></B></FONT>&nbsp;&nbsp;</TD>
	</TR>
	</TABLE><%

	if id_Nro_ck>0 then sumaprod()
	'PN=productoact
	id_prod_ck=0
	for each campo in request.Form
		if left(campo,3)="ck_" then
			if cint(id_prod_ck)=cint(id_Nro_ck) then
				productoact=right(campo,7)
				PN=productoact
				
			end if
			'response.write("<BR>id_prod_ck:" & id_prod_ck & " - id_Nro_ck:" & id_Nro_ck )
			id_prod_ck=id_prod_ck+1
		end if
	next
	'else
	'if (cint(id_Nro_ck)+1) > cint(CuentaProdsel) then
	'	mustraventa()
	'else
	'	prodetalle()
	'end if

	prodetalle()
end sub 'Agregarmultprod()
'---------------------------------------------------------------------------------------------------------
sub finalizarlinea(Productox,Cantidadx,Descuentox)
if len(Descuentox)=0 then Descuentox=0
on error resume next
SQL="SELECT C.PorcDr1 " & _
"FROM flexline.CtaCte C " & _
"WHERE ((C.Empresa='" & empresa & "') AND (C.TipoCtaCte='CLIENTE') AND (C.CtaCte='" & cliente & "') )"
Set rs=oConn.execute(Sql)
porcDR1=rs.fields("PorcDr1")

if isnumeric(replace(trim(porcDR1),"-","")) then 
	'response.write("numerico")
	porcDR1=replace(trim(porcDR1),"-","")*1
end if
	SQL="SELECT P.PRODUCTO, P.GLOSA, P.FAMILIA, D.Valor " & _
	"FROM flexline.ListaPrecio L, flexline.ListaPrecioD D, flexline.PRODUCTO P " & _
	"WHERE D.Empresa = L.Empresa AND D.Empresa = P.EMPRESA AND D.IdLisPrecio = L.IdLisPrecio AND D.Producto = P.PRODUCTO AND ((P.EMPRESA='" & empresa & "') AND (P.TIPOPRODUCTO='EXISTENCIA') AND (P.PRODUCTO='" & Productox & "') AND (L.LisPrecio='" & listprec & "'))"

Set rs=oConn.execute(Sql)
'response.write(sql)

'Escribe
response.write("<TR>")
response.write("<TD align='center'>" & MF & Productox & "</FONT></TD>")
response.write("<TD align='left'  >" & MF & pc(left(rs.fields("Glosa")     , 30),0) & "</FONT></TD>")
response.write("<TH align='center'>" & MF & replace(right("        " & formatnumber(Cantidadx,0) , 9 )," ","&nbsp;") & "</FONT></TH>")
response.write("<TD align='center'>" & MF & replace(right("        " & Descuentox, 6 )," ","&nbsp;") & "</FONT></TD>")
response.write("<TD align='center'>" & MF & replace(right("        " & formatnumber(rs("valor"),0), 9)," ","&nbsp;") & "</FONT></TD>")
response.write("</TR>")
'*****_____Totalizador_____*****
'.:: Flete Neto ::.
SQL0="SELECT fleteneto as VALOR1 FROM Flexline.lpfletes WHERE (Codigo = '" & Productox & "')"
'response.write len(idtipocliente)'idtipocliente
if len(idtipocliente)<>2 then idtipocliente="02"
'sql0="SELECT *, valorflete as Valor1 FROM serverdesa.BDFlexline.flexline.PE_TarifasFlete where idtipocliente='" & idtipocliente & "' and idproducto='" & Productox & "'" '****
sql0="SELECT *, valorflete as Valor1 FROM HandHeld.flexline.PE_TarifasFlete where idtipocliente='" & idtipocliente & "' and idproducto='" & Productox & "'" '****
if empresa="DESAZOFRI" then SQL0="SELECT 0 as VALOR1"

SQL1="SELECT FX_PRODUCTO_FACTOR.PRODUCTO, FX_IMPUESTOS_ILAS.TEXTO, FX_IMPUESTOS_ILAS.VALOR1 FROM todo.flexline.FX_IMPUESTOS_ILAS FX_IMPUESTOS_ILAS, flexline.FX_PRODUCTO_FACTOR FX_PRODUCTO_FACTOR WHERE FX_IMPUESTOS_ILAS.CODIGO = FX_PRODUCTO_FACTOR.Factor AND ((FX_PRODUCTO_FACTOR.EMPRESA='" & EMPRESA & "') AND (FX_PRODUCTO_FACTOR.PRODUCTO='" & Productox & "'))"

set rs1=oConn.execute(SQL0)

if rs1.eof then
	vflete=cdbl(70)
else
	vflete=cdbl(rs1.fields("VALOR1"))
end if
if isnull(porcDR1) then porcDR1=0
cantxprec= (cdbl(rs.fields("valor")) * cdbl(Cantidadx))
desclin = cdbl(Descuentox)
if desclin > 0 then desclin = desclin * -1
midesc 	= porcDR1 + (desclin * -1) * (100 - porcDR1) / 100
drglobal= (cantxprec*midesc/100)
neto 	= round(cantxprec - drglobal)
valflete= round(vflete * Cantidadx)
afecivap= round(neto + valflete)
'response.write("<BR>Producto:" & Productox & " - porcDR1:" & porcDR1 & " | " & len(trim(porcDR1)) & " - desclin:" & desclin &  "<BR>")
'response.write("<BR>Producto:" & Productox & " - cantxprec:" & cantxprec & " - midesc:" & midesc & "<BR>")
'response.write("<BR>Producto:" & Productox & " - cantxprec:" & cantxprec & " - drglobal:" & drglobal & "<BR>")
'response.write("<BR>Producto:" & Productox & " - Neto:" & neto & " - valflete:" & valflete & "<BR>")
PN=Productox
	ilap=15
	if Ucase(left(PN,1))="C" then ilap=20.5
	if Ucase(left(PN,1))="V" then ilap=20.5
	if Ucase(left(PN,1))="W" then ilap=31.5
	if Ucase(left(PN,1))="L" then ilap=31.5
	if Ucase(left(PN,1))="P" then ilap=31.5
	if Ucase(left(PN,1))="B" then ilap=18
	if Ucase(PN)="BE80061" then ilap=10
	if Ucase(PN)="BE80066" then ilap=10
	if Ucase(PN)="BE80067" then ilap=10
	if Ucase(PN)="BE80068" then ilap=10
	if Ucase(PN)="BE80102" then ilap=10
	if Ucase(left(PN,1))="J" then ilap=0
	if Ucase(left(PN,1))="A" then ilap=0
	if Ucase(left(PN,2))="MP" then ilap=0
	if Ucase(left(PN,2))="GN" then ilap=0
if ilap = 0 then
	ilas = 0
else
	ilas = (cdbl(neto)*cdbl(ilap)/100)
End if

afeciva	= (cdbl(afeciva) + cdbl(afecivap))
toila 	= (cdbl(toila) + cdbl(ilas))
totfact	= (cdbl(totfact) + cdbl(afeciva))
toflete = (cdbl(toflete) + cdbl(valflete))
toneto	= (cdbl(toneto) + cdbl(neto))
end sub 'finalizarlinea
'---------------------------------------------------------------------------------------------------------
sub Finalizar()
'***************
'on error resume next
'*****************

if len(P01) = 0 then 
call MIERROR()
else

SQL="SELECT CtaCte.CodLegal, CtaCte.RazonSocial, CtaCte.ListaPrecio, CtaCteDirecciones.Direccion, CtaCte.CtaCte, CtaCte.analisisctacte1 + ' ('+ isnull(CtaCte.analisisctacte12,'N')+')' as giro, CtaCte.CondPago, CtaCte.Comuna, CtaCte.Telefono, CtaCte.LimiteCredito, CtaCte.PorcDr1, ctacte.analisisctacte1 " & _
"FROM flexline.CtaCte CtaCte, flexline.CtaCteDirecciones CtaCteDirecciones " & _
"WHERE CtaCteDirecciones.CtaCte = CtaCte.CtaCte AND CtaCteDirecciones.Empresa = CtaCte.Empresa AND CtaCteDirecciones.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='" & empresa & "') AND (CtaCte.TipoCtaCte='CLIENTE') AND (CtaCteDirecciones.CtaCte='" & cliente & "') AND (CtaCteDirecciones.Principal<>'s'))"
'response.write(sql)

Set rs=oConn.execute(Sql)
ListPrec = rs.fields("ListaPrecio")
analisisc1 = rs.fields("analisisctacte1")
'formulario
response.write("<FORM METHOD=POST ACTION='Pedido_PDA.asp" & _
		"?vend=" & ucase(vend) & "&cliente=" & cliente & "&fechaentrega=" & fechaentrega & "&id_porfolio=" & id_porfolio & _
		"&empresa=" & empresa & "&tipodocto=" & tipodocto & "&sw_iva=" & sw_iva &  _
		"&p01=" & p01 & "&c01=" & c01 & "&d01=" & d01 & _
		"&p02=" & p02 & "&c02=" & c02 & "&d02=" & d02 & _
		"&p03=" & p03 & "&c03=" & c03 & "&d03=" & d03 & _
		"&p04=" & p04 & "&c04=" & c04 & "&d04=" & d04 & _
		"&p05=" & p05 & "&c05=" & c05 & "&d05=" & d05 & _
		"&p06=" & p06 & "&c06=" & c06 & "&d06=" & d06 & _
		"&p07=" & p07 & "&c07=" & c07 & "&d07=" & d07 & _
		"&p08=" & p08 & "&c08=" & c08 & "&d08=" & d08 & _
		"&p09=" & p09 & "&c09=" & c09 & "&d09=" & d09 & _
		"&p10=" & p10 & "&c10=" & c10 & "&d10=" & d10 & _
		"&p11=" & p11 & "&c11=" & c11 & "&d11=" & d11 & _
		"&p12=" & p12 & "&c12=" & c12 & "&d12=" & d12 & "'>")

response.write("<FONT face='Courier' SIZE='1' COLOR=''>")
response.write("Resumen Pedido")
response.write("<HR>")
response.write("<B>" & replace(left(rs.fields("RazonSocial"),30)," ","&nbsp;") & "</B>")
response.write("<BR>Rut. : " & cliente )
response.write("<BR>Dire : " & replace(left(rs.fields("Direccion"),30)," ","&nbsp;") )
response.write("<BR>Tipo : " & replace(left(rs.fields("Giro"),30)," ","&nbsp;") )
response.write("<BR>Pago : " & replace(left(rs.fields("CondPago")& " ",30)," ","&nbsp;") )
response.write("<BR>D.Cli : " & rs.fields("PorcDr1") )
response.write("<HR>")

MF="<FONT face='Courier' SIZE='1' COLOR=''>"

'detalle
response.write("<TABLE>")
response.write("<TR>")
	response.write("<TD align='center'>" & MF & "<B>Codigo</B></TD></FONT>"     )
	response.write("<TD align='center'>" & MF & "<B>Descripcion</B></TD></FONT>")
	response.write("<TD align='center'>" & MF & "<B>Cant</B></TD></FONT>"       )
	response.write("<TD align='center'>" & MF & "<B>Desc</B></TD></FONT>"       )
	response.write("<TD align='center'>" & MF & "<B>Precio</B></TD></FONT>"     )
response.write("</TR>")

iva=0
if len(p01)>0 then call finalizarlinea(p01,c01,d01)
if len(p02)>0 then call finalizarlinea(p02,c02,d02)
if len(p03)>0 then call finalizarlinea(p03,c03,d03)
if len(p04)>0 then call finalizarlinea(p04,c04,d04)
if len(p05)>0 then call finalizarlinea(p05,c05,d05)
if len(p06)>0 then call finalizarlinea(p06,c06,d06)
if len(p07)>0 then call finalizarlinea(p07,c07,d07)
if len(p08)>0 then call finalizarlinea(p08,c08,d08)
if len(p09)>0 then call finalizarlinea(p09,c09,d09)
if len(p10)>0 then call finalizarlinea(p10,c10,d10)
if len(p11)>0 then call finalizarlinea(p11,c11,d11)
if len(p12)>0 then call finalizarlinea(p12,c12,d12)

response.Write("</TD>")
response.Write("</TR>")
response.write("</TABLE>")
'response.write("<BR>empresa: " & empresa & " / tipodocto : " & tipodocto & " / sw_iva : " & sw_iva)
'final
response.write("<HR>")
txt_iva="19"
iva = (cdbl(afeciva)* 19 /100)
if empresa="DESAZOFRI" then 
	if tipodocto="008" then 
		'iva = (cdbl(afeciva)* 19 /100)
		if sw_iva ="on" then
			iva = (cdbl(afeciva)* 19 /100)
		else
			iva = (cdbl(afeciva)* 0.06 /100)
			txt_iva="0,06"
		end if
	End if
end if

apago = (cdbl(afeciva) + cdbl(toila) + cdbl(iva))

response.Write( MF & "&nbsp;&nbsp;Flete&nbsp;&nbsp;&nbsp;&emsp;:&nbsp;" & formatnumber(toflete,0) & "</FONT><BR>")
response.Write( MF & "&nbsp;&nbsp;Afecto&nbsp;&nbsp;&emsp;:&nbsp;" & formatnumber(afeciva,0) & "</FONT><BR>")
response.Write( MF & "&nbsp;&nbsp;iva&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&emsp;:&nbsp;" & formatnumber(iva,0) & "&nbsp;&nbsp;&nbsp;(" & txt_iva & "%)</FONT><BR>")
response.Write( MF & "&nbsp;&nbsp;ilatotal&emsp;:&nbsp;" & formatnumber(toila,0) & "</FONT><BR>")
response.Write( "<B>" & MF & "&nbsp;&nbsp;Total&nbsp;&nbsp;&emsp;:&nbsp;" & formatnumber(apago,0) & "</FONT></B>")
response.write("<HR>")
response.Write("OC : ")
response.Write("<INPUT TYPE='text' NAME='oc' value='" & OC & "'>")
response.Write("<br>")
response.Write( MF & "Fecha&nbsp;Entrega:&nbsp;")
fechaentregaf=right(fechaentrega,2) & "/" & right(left(fechaentrega,6),2) & "/" & left(fechaentrega,4)
%><SELECT NAME="fechaentrega">
<OPTION value="<%=fechaentrega%>"><%=fechaentregaf%></OPTION>
</SELECT>
<BR>

<%
'*************************************************************************************************
'********************     Selecciona Bodega         **********************************************
'*************************************************************************************************

If nuser = 145 or nuser = 76 or nuser = 73 or nuser = 327 then 'Lambert, Rubio, Urra, Molnar, Condell

With Response
	.Write("Bodega: ")
	.Write("<select name='idbodega'>")
	.Write("<option value='1' selected>Lampa</option>")
	.Write("<option value='43' >Lautaro</option>")
	.Write("<option value='46' >San Martin</option>")
	.Write("</select>")

End With


Else

BodSQL="select idbodega from dim_vendedores where idvendedor='"& nuser &"'"
Set rs=oConn.Execute(BodSQL)

Response.Write("<input type='hidden' name='idbodega' value='"& rs(0) &"'/>")

End If



%>

<SCRIPT LANGUAGE="JavaScript">
<!--
function limitalargo(objeto){
	if (objeto.value.length>99){
		alert('El Texto a superado el maximo permitido');
	};
}
//-->
//
</SCRIPT>
<BR>Observacion <BR>
Gerencia : <BR>
<TEXTAREA id='OBS' NAME='OBS' ROWS='3' COLS='20' maxlength='100' onKeypress="limitalargo(this);if ((event.keyCode > 32 && event.keyCode < 48) || (event.keyCode > 57 && event.keyCode < 65) || (event.keyCode > 90 && event.keyCode < 97)){ event.returnValue = false;}">
</TEXTAREA>
<%
monto_minimo=0
if(idempresa=1 or idempresa=4) then
	sql="SELECT [pedido_minimo] FROM [DesaERP].[dbo].[DIM_VENDEDORES] Where idvendedor="& nuser &" AND idempresa=" & idempresa
	Set rs=oConn.execute(sql)

	'if not rs.eof then monto_minimo=cdbl(rs("monto_minimo"))
	'if instr(ucase(vend), "OCD")    then monto_minimo=0
	'if instr(ucase(vend), "GEO")    then monto_minimo=0
	'if len(id_porfolio)>0           then monto_minimo=0
	'if analisisc1="39 PARTICULARES" then monto_minimo=0
	'response.write(analisisc1)
	'monto_minimo=0'nula drop said

	if not rs.eof then 
		if rs(0)="S" then 
			monto_minimo=15000
		else
			monto_minimo=0
		end if	
	end if
end if

with response
  .write("<BR>")
  .Write(chr(13) & "<BR>Observacion <BR>")
  .Write(chr(13) & "Factura : <BR>")
  .Write "<TEXTAREA NAME='OBS2' ROWS='3' COLS='20' onKeypress="&chr(34)&"if ((event.keyCode > 32 && event.keyCode < 48) || (event.keyCode > 57 && event.keyCode < 65) || (event.keyCode > 90 && event.keyCode < 97)) event.returnValue = false;"&chr(34)
  .write "></TEXTAREA>"'parche simbolos
  .write("<BR>")
  if cdbl(apago)<monto_minimo then
	.write("<BR><H2>El monto mínimo del pedido es de: " & formatnumber(monto_minimo,0) & " pesos</H2>")
  else
  .write("<BR><INPUT TYPE='submit' value='Guardar Pedido'>")
  end if
  .write("</FONT>")
  .write("</FORM>")
end with

end if
end sub 'Finalizar()
'--------------------------------------------------------------------------------------------------------
private function Cantidadminima(CantPedida,CantMinima,sku)
	'on error resume next
	'response.write("<BR>CantPedida : " & CantPedida)
	if (CantPedida)<CantMinima then
		Cantidadminima=CantMinima
	%>
	<HR><H3>Cantidad Minima de venta <%=CantMinima%> unidades</H3><HR>
	<%
	else
		resultado=CantPedida mod CantMinima
		'response.write("<BR>resultado : " & resultado)
		if resultado = 0 then
			Cantidadminima=CantPedida
		else
			%>
				<HR><H3>(<%=sku%>) Venta en Caja Cerrada</H3><HR>
			<%
			Cantidadminima=CantPedida+(CantMinima-resultado)
		end if
	end if
	'response.write("<BR>Cantidadminima : " & Cantidadminima)
end function
'--------------------------------------------------------------------------------------------------------
sub sumaprod()

if left(midescuent,1)="-" then midescuent=replace(midescuent,"-","")
EntDes=midescuent
if instr(midescuent,",")>0 then
EntDes=left(midescuent,instr(midescuent,",")-1)
end if

if len(EntDes)>2 then midescuent=right(EntDes,2)
'midescuent=left(midescuent,instr(midescuent,",")-1)

if len(p01)=0 then
	if PN="VN32130" then
		p01="LI47909"
		c01=micantidad
		d01=0
	End if
	if PN="LI47909" then
		p01="VN32130"
		c01=micantidad
		d01=0
	End if
Else
	if PN="LI47909" or PN="VN32130" then
		PN=""
		micantidad=""
		midescuent=""
	end if
end if
'BE80068
if empresa <>"DESAZOFRI" then
	if PN="BE80051" then micantidad=Cantidadminima(micantidad, 24, "BE80051") 'Venta en caja cerrada
	'if PN="BE80068" then micantidad=Cantidadminima(micantidad,  2, "BE80068") 'Venta en caja cerrada
	if PN="BE80061" then micantidad=Cantidadminima(micantidad, 24, "BE80061") 'Venta en caja cerrada
end if 
'response.write(micantidad)
'No duplicado

sql="select * from desaerp.dbo.COM_SKUPACK_ENC where idSKU='" & PN & "'"
set rs=oConn.execute(sql)
Do until rs.eof
	SKU_PACK=PN
	UNICO=rs("productounico")
	sql="select * from desaerp.dbo.COM_SKUPACK_DET where SKU_PACK='" & SKU_PACK & "'"
	set rs1=oConn.execute(sql)
	x1=1
	Do until rs1.eof
		if x1=1 then
			p01=rs1("SKU_RECETA")
			c01=cdbl(rs1("cantidad1")) 
			d01=0
		end if
		if x1=2 then
			p02=rs1("SKU_RECETA")
			c02=cdbl(rs1("cantidad1")) 
			d02=0
		end if
		if x1=3 then
			p03=rs1("SKU_RECETA")
			c03=cdbl(rs1("cantidad1")) 
			d03=0
		end if
		if x1=4 then
			p04=rs1("SKU_RECETA")
			c04=cdbl(rs1("cantidad1")) 
			d04=0
		end if
		if x1=5 then
			p05=rs1("SKU_RECETA")
			c05=cdbl(rs1("cantidad1")) 
			d05=0
		end if
		if x1=6 then
			p06=rs1("SKU_RECETA")
			c06=cdbl(rs1("cantidad1")) 
			d06=0
		end if
		if x1=7 then
			p07=rs1("SKU_RECETA")
			c07=cdbl(rs1("cantidad1")) 
			d07=0
		end if
		if x1=8 then
			p08=rs1("SKU_RECETA")
			c08=cdbl(rs1("cantidad1")) 
			d08=0
		end if
		if x1=9 then
			p09=rs1("SKU_RECETA")
			c09=cdbl(rs1("cantidad1")) 
			d09=0
		end if
		if x1=10 then
			p10=rs1("SKU_RECETA")
			c10=cdbl(rs1("cantidad1")) 
			d10=0
		end if
		if x1=11 then
			p11=rs1("SKU_RECETA")
			c11=cdbl(rs1("cantidad1")) 
			d11=0
		end if
		if x1=12 then
			p12=rs1("SKU_RECETA")
			c12=cdbl(rs1("cantidad1")) 
			d12=0
		end if
		x1=x1+1
	rs1.movenext
	loop
	PN=""
	micantidad=""
	midescuent=""
	OBS=SKU_PACK
	OC=SKU_PACK
	Finalizar()
	exit sub 
rs.movenext
loop


if p01=PN or p02=PN or p03=PN or p04=PN or p05=PN or p06=PN or p07=PN or p08=PN or p09=PN or p10=PN or p11=PN or p12=PN then
	PN=""
	micantidad=""
	midescuent=""
	response.write("<FONT SIZE=3 COLOR='#003366' face='Arial'><BR><BR><CENTER><B>No Se Permite ingresar dos veces el mismo producto</B></CENTER><BR></FONT>")
end if


if len(p01)=0 then
p01=PN
c01=cdbl(micantidad)
d01=midescuent
else
	if len(p02)=0 then
	p02=PN
	c02=micantidad
	d02=midescuent
	else
		if len(p03)=0 then
		p03=PN
		c03=micantidad
		d03=midescuent
		else
			if len(p04)=0 then
			p04=PN
			c04=micantidad
			d04=midescuent
			else
				if len(p05)=0 then
				p05=PN
				c05=micantidad
				d05=midescuent
				else
					if len(p06)=0 then
					p06=PN
					c06=micantidad
					d06=midescuent
					else
						if len(p06)=0 then
						p06=PN
						c06=micantidad
						d06=midescuent
						else
							if len(p07)=0 then
							p07=PN
							c07=micantidad
							d07=midescuent
							else
								if len(p08)=0 then
								p08=PN
								c08=micantidad
								d08=midescuent
								else
									if len(p09)=0 then
									p09=PN
									c09=micantidad
									d09=midescuent
									else
										if len(p10)=0 then
										p10=PN
										c10=micantidad
										d10=midescuent
										else
											if len(p11)=0 then
											p11=PN
											c11=micantidad
											d11=midescuent
												if empresa="DESAZOFRI" then
													%>
													<HR><H3>limite 10 Productos</H3><HR>
													<%
													p11=""
													c11=""
													d11=""
												end if
											else
												if len(p12)=0 then
												p12=PN
												c12=micantidad
												d12=midescuent
												else

												end if
											end if
										end if
									end if
								end if
							end if
						end if
					end if
				end if
			end if
		end if
	end if
end if


if miboton="Agregarmult" then exit sub
	mustraventa()
end sub 'sumaprod()

'-------------------------------------------------------------------------------------------------
Sub Agregarprod()
session("echo")= false
'+++++++++++++++++++'
'idempresa=1
if empresa="LACAV" then idempresa=4
SQL="SELECT Familia, Count(PRODUCTO) 'Contar de PRODUCTO' " & _
"FROM Flexline.FB_PRODUCTO_OK " & _
"GROUP BY Familia" 'lista familias
set rs=oConn.execute(Sql)

%>
<BR>
<FORM name="buscafrm" id="buscafrm" METHOD=POST ACTION=''>
	<FONT SIZE=2 face='arial' COLOR='#000066'><B>Buscar Producto.</B></FONT>&nbsp;
	<INPUT TYPE='text' size='8' NAME='buscaproducto' onKeyPress="if(event.keyCode==13){document.buscafrm.submit();return 0;}">
	<INPUT TYPE='submit' name='buscaprod' value='?'>
</FORM>


<FORM METHOD=POST ACTION=''>
<% do until rs.eof
	if not left(lcase(trim(rs.fields("familia"))),8)="material" then
		%><INPUT TYPE='submit' Name='Familia' Value='<%=pc(rs.fields("familia"),0)%>'><%
	end if
  rs.movenext
  loop
%>

</FORM>
<%

nFamilia=trim(cstr(request.form("familia" )))
if len(nFamilia)=0 then nFamilia=request.cookies("NFamilia")

'response.write("<FORM METHOD=POST ACTION=''>")
if len(trim(cstr(request.form("buscaprod" )))) then ' si hay busqueda
	SQL="select * from flexline.FB_PRODUCTO_OK  where empresa='" & empresa & "' and busca like'%" & trim(cstr(request.form("buscaproducto" ))) & "%'"
	'SQL="select * from Flexline.PRODUCTO where glosa like '%" & trim(cstr(request.form("buscaproducto" ))) & "%'"
	if lcase(id_porfolio)="porfolio1" then
		SQL="select * from flexline.FB_PRODUCTO_OK " & _
		"where  empresa='" & empresa & "' and busca like '%" & trim(cstr(request.form("buscaproducto"))) & "%' " & _
		"and porfoliodesa1='S' "
	end if
	if lcase(id_porfolio)="porfolio2" then
		SQL="select * from flexline.FB_PRODUCTO_OK " & _
		"where  empresa='" & empresa & "' and busca like '%" & trim(cstr(request.form("buscaproducto"))) & "%' " & _
		"and porfoliodesa2='S' " 
	end if
else
	':: lista producto ::
	if len(nfamilia)=0 then nfamilia="A"
	SQL="select * from flexline.FB_PRODUCTO_OK  where  empresa='" & empresa & "' and Familia='" & nfamilia & "' " & _
	"order by tipo, glosa" ' 

	if lcase(id_porfolio)="porfolio1" then
		SQL="select * from flexline.FB_PRODUCTO_OK  where  empresa='" & empresa & "' and Familia='" & nfamilia & "' " & _
		"and porfoliodesa1='S' " & _
		"order by tipo, glosa" ' 
	end if

	if lcase(id_porfolio)="porfolio2" then
		SQL="select * from flexline.FB_PRODUCTO_OK  where  empresa='" & empresa & "' and Familia='" & nfamilia & "' " & _
		"and porfoliodesa2='S' " & _
		"order by tipo, glosa" ' 
	end if
	'response.write(sql)
	'if ucase(nfamilia)=ucase("vinos") then SQL=replace(SQL,"order by glosa","order by tipo, glosa")
	'if ucase(nfamilia)=ucase("vinos") then response.write(sql)
end if

': filtro UNIMARC - pernod ricard
if left(cliente,8)="81537600" or left(cliente,8)="76012795"  then
	sql=replace(ucase(sql),"WHERE", "WHERE subtipo <> 'PERNOD RICARD' and ") 
end if

set rs=oConn.execute(Sql)
'response.write(sql)
'+++++++++++++++++++'
'nuser=17
	misql="select idvendedor from desaerp.dbo.DIM_VENDEDORES where nombre='" & vend & "'"
	set rs1=oConn.execute(misql)
	if rs1.eof then
		response.write("<BR><BR><BR><BR>********************************************<BR>No se encontro registro en la tabla : Dim_Vendedores<BR>********************************************")
		exit sub
	end if
	nuser  = rs1.fields(0)

'response.write(nuser)
misql="select idproducto from handheld.flexline.pda_skunoven where idempresa='" & idempresa & "' and idvendedor=" & nuser & " UNION SELECT idsku FROM [HandHeld].[dbo].[PDA_SKUCANAL] Where idempresa='" & idempresa & "' and idcanal=" & idcanal  & " and idsucursal=" & idlocal
set rs1=oConn.execute(misql)
'response.write(sql)

%>
	<INPUT TYPE="hidden" NAME="nprod" value="<%=CuentaProductos%>">
	<FORM METHOD=POST name="frm_xprod" id="frm_xprod" ACTION="">
	<TABLE>
<%

familiant="familia"
tipoant="tipo"
skunoven=""

'.:: Indice ::.
if ucase(nfamilia)=ucase("vinos") then
	%>	
	<TR>
		<TD COLSPAN="2">
			<a name='indice'>
	<%
	do until rs.eof
		if tipoant<>rs.fields("tipo") then
			%><a href="#<%=rs.fields("tipo")%>"><FONT SIZE="2" face="verdana" COLOR=""><B>>>&nbsp;<%=pc(rs.fields("tipo"),0)%></B></FONT></a><BR>
			<%
		end if
		tipoant=rs.fields("tipo")
	rs.movenext
	loop
	tipoant="tipo"
	rs.movefirst
	%>
	</TD>
	</TR>
	<%
end if

do until rs.eof
	if not rs1.bof then rs1.movefirst
	do until rs1.eof
		'response.write(rs1.fields("idproducto"))
		if lcase(trim(rs.fields("producto")))=lcase(trim(rs1.fields("idproducto"))) then
		skunoven=lcase(trim(rs1.fields("idproducto")))
		'response.write(rs1.fields("idproducto"))
		exit do
		end if
	rs1.movenext
	loop
if not lcase(trim(rs.fields("producto")))=skunoven then ' si el producto no esta an la lista skunoven
'escribe marca
	if familiant<>rs.fields("familia") then
	response.write("<TR>")
		response.write("<TD></TD><TD><FONT SIZE='2' face='Arial' COLOR='#C0C0C0'><B>" & _
		replace(rs.fields("familia")," ","&nbsp;") & _
		"</B></FONT></TD>")
	response.write("</TR>")
	end if 
'escribe tipo
	if tipoant<>rs.fields("tipo") then
	response.write("<TR bgcolor='#FFFFFF' border='1' style='border: solid'><TD valign='bottom' >")
	if ucase(nfamilia)=ucase("vinos") then response.write("<a href='#indice'><IMG SRC='../../Includes/up.gif' border='0' style='text-decoration: none'></a>")
	response.write("</TD>")
		response.write("<TD bgcolor='#FFFFFF' border='1' style='border: solid'><HR><FONT SIZE='2' face='verdana' COLOR='#000066'><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.::&nbsp;" & _
		replace(rs.fields("tipo")," ","&nbsp;") & _
		"&nbsp;::.</B></FONT><a name='" & rs.fields("tipo") & "'></a></TD>")
	response.write("</TR>")
	end if 
'escribe glosa
	response.write("<TR>")
		%><TD><INPUT TYPE='checkbox' NAME="ck_<%=rs.fields("producto")%>" onClick="
			//alert(this.checked);
			var xnprod=document.getElementById('nprod').value*1;
			if(this.checked==true){
				if(xnprod>11){
					alert('12 Productos, limite por pedido')
					this.checked=false;
					if(confirm('¿Confirma los productos seleccionados?')){
						document.getElementById('frm_xprod').submit();
						//alert(submit);
					}
				}else{
					document.getElementById('nprod').value=xnprod+1;
				}
			}else{
				document.getElementById('nprod').value=xnprod-1;
			}
		"></TD><%
		response.write("<TD>" & _
		"<a href='Pedido02.asp?vend=" & ucase(vend) & "&cliente=" & cliente & "&fechaentrega=" & fechaentrega & "&id_porfolio=" & id_porfolio & _
		"&empresa=" & empresa & "&tipodocto=" & tipodocto & "&sw_iva=" & sw_iva & "&PN=" & rs.fields("producto") & _
		"&p01=" & p01 & "&c01=" & c01 & "&d01=" & d01 & _
		"&p02=" & p02 & "&c02=" & c02 & "&d02=" & d02 & _
		"&p03=" & p03 & "&c03=" & c03 & "&d03=" & d03 & _
		"&p04=" & p04 & "&c04=" & c04 & "&d04=" & d04 & _
		"&p05=" & p05 & "&c05=" & c05 & "&d05=" & d05 & _
		"&p06=" & p06 & "&c06=" & c06 & "&d06=" & d06 & _
		"&p07=" & p07 & "&c07=" & c07 & "&d07=" & d07 & _
		"&p08=" & p08 & "&c08=" & c08 & "&d08=" & d08 & _
		"&p09=" & p09 & "&c09=" & c09 & "&d09=" & d09 & _
		"&p10=" & p10 & "&c10=" & c10 & "&d10=" & d10 & _
		"&p11=" & p11 & "&c11=" & c11 & "&d11=" & d11 & _
		"&p12=" & p12 & "&c12=" & c12 & "&d12=" & d12 & _
		"' style='text-decoration: none'>")
with response		
		.write "<FONT SIZE='2' face='Arial' COLOR='#000066'>"
		.write replace(replace(rs.fields("glosa"),"RED BULL SIX","RED BULL (GEOCARGO) SIX")," ","&nbsp;")
		.write "<a></FONT></TD>"
end with		
	response.write("</TR>")
	familiant=rs.fields("familia")
	tipoant=rs.fields("tipo")
end if ' fin: si el producto no esta an la lista skunoven

rs.movenext
loop
'****************************


%>
		<TR>
			<TD></TD>
			<TD><FONT COLOR='#FFFFFF'>____________________________________________________</FONT></TD>
		</TR>
	</TABLE>
	<INPUT TYPE="submit" value=">> Confirmar los productos seleccionados">
	<INPUT TYPE="hidden" name="boton" id="miboton" value="Agregarmult">
</FORM>
<%
end Sub 'Agregarprod()
'---------------------------------------------------------------------------------------------------
Sub prodetalle()
'on error resume next
sql="SELECT analisisctacte1 FROM flexline.ctacte WHERE (ctacte='" & cliente & "') AND (TipoCtaCte='CLIENTE') and (left(analisisctacte1,2) not in('04','36','43') )"
set mRs=oConn.execute(sql)
if mRs.eof then
	LSQL="select 'OFF PREMISE' as Grupo"
else
	LSQL="select 'ON PREMISE' as Grupo"
end if

set mRs=oConn.execute(LSQL)
		%>
		<script language="javascript">
		function fijaDesc()
		{
		misel=document.getElementById('cantidad').selectedIndex;
		switch (misel)
		 {
			case 0:
				document.getElementById('descuento').value='6.0';
				break;
			case 1:
				document.getElementById('descuento').value='6.9';
				break;
			case 2:
				document.getElementById('descuento').value='9.07';
				break;
			case 3:
				document.getElementById('descuento').value='10.29';
				break;
		 }
		}
		</script>
		<%
valprecio=""
	SQL="SELECT P.PRODUCTO, P.GLOSA, P.FAMILIA, D.Valor " & _
	"FROM flexline.ListaPrecio L, flexline.ListaPrecioD D, flexline.PRODUCTO P " & _
	"WHERE D.Empresa = L.Empresa AND D.Empresa = P.EMPRESA AND D.IdLisPrecio = L.IdLisPrecio AND D.Producto = P.PRODUCTO AND ((P.EMPRESA='" & empresa & "') AND (P.TIPOPRODUCTO='EXISTENCIA') AND (P.PRODUCTO='" & PN & "') AND (L.LisPrecio='" & listprec & "'))"
	'response.write(sql)
set rs=oConn.execute(Sql)
if not rs.eof then
	valprecio=rs("valor")
	familia=rs("FAMILIA")
end if

if left(PN, 5)="IPACK" then PN="L" & PN

SQL="select * from Flexline.FB_PRODUCTO_OK where producto='" & PN & "'"
	'ucaja=rsu("factoralt")
	'ubase=rsu("analisisproducto8")
	'upallet=rsu("analisisproducto9")
set rs=oConn.execute(Sql)
'response.write(sql)
'exit sub 

'response.write(PN)
if isnumeric(trim(rs.fields("factoralt"))) then
	ucaja=cdbl(rs.fields("factoralt"))
else
	ucaja=0
end if

if isnumeric(trim(rs.fields("factorbase"))) then
	ubase=cdbl(rs.fields("factorbase")) 
else
	ubase=0 
end if

if isnumeric(trim(rs.fields("factorPallet"))) then
	upall=cdbl(rs.fields("factorPallet"))
else
	upall=0
end if

'on error resume next

response.write("<FONT SIZE='1' face='Arial' COLOR='#808080'>Producto&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Precio unitario&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Descuento Máximo</FONT><BR>")
response.write("<FONT SIZE='2' face='Arial' COLOR='#000066'><B>" & rs.fields("producto") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$&nbsp;" & valprecio & "</B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<label id=""dsct""></label>%</FONT><BR>")
response.write("<FONT SIZE='1' face='Arial' COLOR='#808080'>Nombre</FONT><BR>")
response.write("<FONT SIZE='2' face='Arial' COLOR='#000066'><B>" & rs.fields("glosa") & "</B></FONT><BR>")
response.write("<FONT SIZE='1' face='Arial' COLOR='#808080'>Empaque</FONT><BR>")
response.write("<FONT SIZE='2' face='Arial' COLOR='#000066'><B>" & ucaja & "</B>&nbsp;x&nbsp;Caja&nbsp;")
response.write("/&nbsp;<B>" & (ucaja * ubase ) & "</B>&nbsp;x&nbsp;Base&nbsp;/&nbsp;<B>" & (ucaja * ubase * upall) & "</B>&nbsp;x&nbsp;Pallet&nbsp;</FONT><BR>")
productoactivo=rs.fields("producto")
factoralt=rs.fields("factoralt")

response.write("<FORM METHOD=POST name='frmpn' id='frmpn' ACTION='?vend=" & ucase(vend) & "&cliente=" & cliente & "&fechaentrega=" & fechaentrega & "&id_porfolio=" & id_porfolio & _
		"&empresa=" & empresa & "&tipodocto=" & tipodocto & "&sw_iva=" & sw_iva & "&PN=" & rs.fields("producto") & _
		"&p01=" & p01 & "&c01=" & c01 & "&d01=" & d01 & _
		"&p02=" & p02 & "&c02=" & c02 & "&d02=" & d02 & _
		"&p03=" & p03 & "&c03=" & c03 & "&d03=" & d03 & _
		"&p04=" & p04 & "&c04=" & c04 & "&d04=" & d04 & _
		"&p05=" & p05 & "&c05=" & c05 & "&d05=" & d05 & _
		"&p06=" & p06 & "&c06=" & c06 & "&d06=" & d06 & _
		"&p07=" & p07 & "&c07=" & c07 & "&d07=" & d07 & _
		"&p08=" & p08 & "&c08=" & c08 & "&d08=" & d08 & _
		"&p09=" & p09 & "&c09=" & c09 & "&d09=" & d09 & _
		"&p10=" & p10 & "&c10=" & c10 & "&d10=" & d10 & _
		"&p11=" & p11 & "&c11=" & c11 & "&d11=" & d11 & _
		"&p12=" & p12 & "&c12=" & c12 & "&d12=" & d12 & _
		"'>")

Minetox=valprecio

response.write("</TR>")
response.write("</TABLE>")
'response.write("<INPUT TYPE='submit' Value='Guardar'>")
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ INI
banderas=1'1 ' Imprime detalles en pantalla descuentos

'idempresa=1
'if id_porfolio="porfolio2" then idempresa=4

fechaanalisis=year(date) & right("00" & month(date),2) & right("00" & day(date),2)
if len(idcanal)=0 then idcanal="0"
descuento=9
descuentomaximo=9
desctipo="variable"
':: GRUPO VENTAS ::
id_gurpoventas="1"

sql="select v.idmgrupoventa, v.idgrupoventa from sqlserver.desaerp.dbo.DIM_vendedores as v " & _
	"where v.idempresa='" & idempresa & "' and v.nombre = '" & vend & "'"
'set rsp=oConn.execute(sql)
'if rsp(0)="25" then grupo_ocd="si"

	'sql="SELECT  g.idgrupoventa " & _
	'"FROM sqlserver.desaerp.dbo.DIM_GRUPOVTADET as G INNER JOIN " & _
	'" sqlserver.desaerp.dbo.DIM_VENDEDORES as V ON g.idvendedor = v.idvendedor " & _
	'"WHERE v.idempresa=" & idempresa & " and (v.nombre = '" & vend & "')"
	set rsp=oConn.execute(sql)
	do until rsp.eof
		id_gurpoventas=rsp("idgrupoventa")
		id_mgurpoventas=rsp("idmgrupoventa")
	rsp.movenext
	loop

if 	id_mgurpoventas="25" then
	tipoventa=2
else
	tipoventa=1
end if
	
	
	
	
	
	
'if banderas=1 then response.write("<BR>TIPODEVENTA : " & tipoventa)	

':: DESCUENTO BASE ::
sql="select replace(isnull(analisisctacte4,9),'.',',') as Descuento from handheld.Flexline.CtaCte as c " & _
"where empresa='" & empresa & "' and tipoctacte='cliente' and ctacte='" & cliente & "' "
set rsp=oConn.execute(sql)
if not rsp.eof then ' cliente
	descuento_base=rsp.fields("Descuento")
	if len(trim(descuento))=0 then descuento_base=9
	'if banderas=1 then response.write("<BR>DESCUENTO BASE : " & descuento_base)
end if

'*********************************
'****DESCUENTO POR CLIENTE********
'*********************************

'Se inicia las variables a utilizar en la busqueda del descuento en el grupo de DESCUENTO POR CLIENTE.
descuento_cli=0
desctipo_cli="variable"


'.::COM_CLIPROMDESC ::.
'Sentencia SQL que trae el descuento.
sql="select * from sqlserver.desaerp.dbo.COM_CLIPROMDESC " & _
"where idempresa=" & idempresa & " and vb_aprobacion='S' and idproducto='" & PN & "' and idcliente='" & cliente & "' and fecha_desde<=" & fechaanalisis & " and fecha_hasta>=" & fechaanalisis & " and idtipoventa=" & tipoventa
set rsp=oConn.execute(sql)
	if not rsp.eof then
	'Si extrae un descuento se aplica el valor a descuento_cli.
		descuento_cli=rsp.fields("Descuento")
		trato=rsp.fields("sw_trato")
		if trato="F" Then desctipo_cli="fijo" Else desctipo_cli="variable"
		'if banderas=1 then response.write("<BR>COM_CLIPROMDESC : " & descuento_cli & "Tipo: " &  desctipo_cli)
	'Si no encuentra, va a buscar otro descuento.
	else
		'.::COM_MAXPAGDESC ::.
		'Sentencia SQL que trae el descuento.
		clientepag=trim(left(cliente, instr(1,cliente," ")-1))
		sql="Select isnull(d.descuento,0) as descuento, d.sw_trato " & _
		"from sqlserver.desaerp.dbo.COM_MAXPAGDESC as d " & _
		"WHERE (d.vb_aprobacion = 's') AND (d.fecha_hasta >=" & fechaanalisis & ") AND (d.fecha_desde <=" & fechaanalisis & ") AND (d.idpagador = '" & clientepag & "') AND (d.idproducto = '" & PN & "') AND (d.idempresa = " & idempresa & ") AND (idtipoventa=" & tipoventa & ")"
		set rsp=oConn.execute(sql)
			if not rsp.eof then 
			'Si extrae un descuento se aplica el valor a descuento_cli.
					descuento_cli=rsp.fields("Descuento")
					trato=rsp.fields("sw_trato")
					if trato="F" Then desctipo_cli="fijo" Else desctipo_cli="variable"
					'if banderas=1 then response.write("<BR>COM_MAXPAGDESC : " & descuento_cli & "Tipo: " &  desctipo_cli)
					'Si no encuentra, va a buscar otro descuento.
			else
				'.:: COM_PAGPROMDESC ::.
				'Variable que se utiliza en la busqueda del SQL
				clientepag=trim(left(cliente, instr(1,cliente," ")-1))
				'Sentencia SQL que trae el descuento.
				sql="select * from sqlserver.desaerp.dbo.COM_PAGPROMDESC " & _
				"where idempresa=" & idempresa & " and vb_aprobacion='S' and idproducto='" & PN & "' and idpagador='" & clientepag & "' and fecha_desde<=" & fechaanalisis & " and fecha_hasta>=" & fechaanalisis & " and idtipoventa=" & tipoventa
				set rsp=oConn.execute(sql)
					if not rsp.eof then 
					'Si extrae un descuento se aplica el valor a descuento_cli.
						descuento_cli=rsp.fields("Descuento")
						trato=rsp.fields("sw_trato")
						if trato="F" Then desctipo_cli="fijo" Else desctipo_cli="variable"
						'if banderas=1 then response.write("<BR>COM_PAGPROMDESC : " & descuento_cli& "Tipo: " &  desctipo_cli)
						'Si no encuentra, va a buscar otro descuento.
					else	
						'.:: COM_GRUPROMDESC ::.
						'Sentencia SQL que trae el descuento.
						sql="select g.descuento, g.sw_trato from sqlserver.desaerp.dbo.com_grupromdesc as g inner join " & _
						"sqlserver.desaerp.dbo.dim_grupoclirel as c on  " & _
						"c.idempresa=g.idempresa and c.idgrupocliente=g.idgrupocliente " & _
						"where c.idempresa=" & idempresa & " and vb_aprobacion='S'  and c.idcliente='" & clientepag & "' and g.idproducto='" & PN & "' " & _
						"and fecha_desde<=" & fechaanalisis & " and fecha_hasta>=" & fechaanalisis & " and idtipoventa=" & tipoventa
						set rsp=oConn.execute(sql)
							if not rsp.eof then 
							'Si extrae un descuento se aplica el valor a descuento_cli.
								descuento_cli=rsp.fields("Descuento")
								trato=rsp.fields("sw_trato")
								if trato="F" Then desctipo_cli="fijo" Else desctipo_cli="variable"
								'if banderas=1 then response.write("<BR>com_grupromdesc : " & descuento_cli & "Tipo: " &  desctipo_cli)
							end if 'fin, grupromdesc
					end if 'fin, PAGPROMDESC	
			end if 'fin, MAXPAGDESC
	end if 'fin, CLIPROMDESC
	
	
'*********************************
'****DESCUENTO POR SEGMENTO*******
'*********************************	
		
'Se inicia las variables a utilizar en la busqueda del descuento en el grupo de DESCUENTO POR SEGMENTO.
descuento_seg=0
desctipo_seg="varible"

'.:: COM_SUCPROMDESC ::.
'Sentencia SQL que trae el descuento.
sql="SELECT	SD.Descuento, SD.Sw_trato,SD.Idcanal,DV.idlocal " & _
	"FROM	[DesaERP].[dbo].[COM_SUCCPROMDESC] as SD " & _
	"INNER JOIN	[HandHeld].[Flexline].[CtaCte] as C " & _
		"ON	RIGHT('00' + CAST(SD.idcanal AS NVARCHAR),2)=RTRIM(LEFT(ISNULL(C.Analisisctacte1,'0'),2)) " & _
	"INNER JOIN	[HandHeld].[dbo].[DIM_VENDEDORES] AS DV " & _
		"ON	SD.idsucursal=DV.idlocal AND DV.nombre=C.Ejecutivo and SD.IdSucursal=DV.Idlocal " & _
	"WHERE C.Empresa='DESA' AND SD.Idempresa=" & idempresa & " AND C.TipoCtacte='CLIENTE' " & _
	"AND SD.Fecha_desde<= " & fechaanalisis & "  AND SD.Fecha_hasta >= " & fechaanalisis & " " & _
	"AND SD.Vb_aprobacion='S' AND DV.idempresa=" & idempresa & " AND SD.idproducto='" & PN & "' " & _
	"AND C.Ctacte='" & cliente & "' AND SD.Idtipoventa=" & tipoventa
set rsp=oConn.execute(sql)
if not rsp.eof then
	'Si extrae un descuento se aplica el valor a descuento_seg.
		descuento_seg=rsp.fields("Descuento")
		trato=rsp.fields("sw_trato")
		if trato="F" Then desctipo_seg="fijo" Else desctipo_seg="variable"
		'if banderas=1 then response.write("<BR>COM_SUCPROMDESC : " & descuento_seg & "Tipo: " &  desctipo_seg)
		'response.write(sql)		
		'Si no encuentra, va a buscar otro descuento.
	else
		'.:: COM_CANPROMDESC ::.
		'Sentencia SQL que trae el descuento.
		sql="SELECT     m.descuento, m.sw_trato " & _
		"FROM         sqlserver.desaerp.dbo.COM_CANPROMDESC as m  " & _
		"inner join handheld.Flexline.CtaCte as c " & _
		"	on right('00' + cast(m.idcanal as nvarchar),2)=rtrim(left(isnull(c.analisisctacte1,'0'),2)) " & _
		"where c.empresa='DESA' and  m.idempresa=" & idempresa & " and c.tipoctacte='cliente' " & _
		"and m.fecha_desde<=" & fechaanalisis & " and m.fecha_hasta >= " & fechaanalisis & " " & _
		"and m.vb_aprobacion='S' and m.idproducto='" & PN & "' " & _
		"and c.ctacte='" & cliente & "' and idtipoventa=" & tipoventa
		set rsp=oConn.execute(sql)
			if not rsp.eof then
			'Si extrae un descuento se aplica el valor a descuento_seg.
				descuento_seg=rsp.fields("Descuento")
				trato=rsp.fields("sw_trato")
				if trato="F" Then desctipo_seg="fijo" Else desctipo_seg="variable"
				'if banderas=1 then response.write("<BR>COM_CANPROMDESC : " & descuento_seg & "Tipo: " &  desctipo_seg)	
				'Si no encuentra, va a buscar otro descuento.
			else
				'.:: COM_MAXMCANDESC ::.
				'Sentencia SQL que trae el descuento.
				sql="  SELECT isnull(MMCD.descuento,0) as descuento, MMCD.sw_trato " & _
				"FROM sqlserver.DesaERP.dbo.DIM_CANALES as canal " & _
				"inner join handheld.Flexline.CtaCte as cta " & _
				"ON canal.idcanal=rtrim(left(isnull(cta.analisisctacte1,'0'),2))"  & _
				"inner join SQLSERVER.DesaERP.dbo.COM_MAXMCANDESC AS MMCD " & _
				"ON canal.idmegacanal=MMCD.idmegacanal " & _
				"Where cta.tipoctacte='cliente' and " & _
				"MMCD.fecha_desde<=" & fechaanalisis & " and MMCD.fecha_hasta >=" & fechaanalisis & " " & _
				"and MMCD.vb_aprobacion='S' and MMCD.idproducto='" & PN & "' " & _
				"and cta.ctacte='" & cliente & "' and MMCD.idempresa=" & idempresa & " and idtipoventa=" & tipoventa
				set rsp=oConn.execute(sql)
					if not rsp.eof then 
					'Si extrae un descuento se aplica el valor a descuento_seg.
						descuento_seg=rsp.fields("Descuento")
						trato=rsp.fields("sw_trato")
						if trato="F" Then desctipo_seg="fijo" Else desctipo_seg="variable"
						'if banderas=1 then response.write("<BR>COM_MAXMCANDESC : " & descuento_seg & "Tipo: " &  desctipo_seg)
						'Si no encuentra, va a buscar otro descuento.
					else
						'.:: COM_MEGPROMDESC ::.
						'Sentencia SQL que trae el descuento.
						sql="SELECT     isnull(descuento,0) as descuento, sw_trato " & _
						"FROM  sqlserver.desaerp.dbo.COM_MEGPROMDESC " & _
						"WHERE     (vb_aprobacion = 's') AND (fecha_hasta >=" & fechaanalisis & ") AND (fecha_desde <=" & fechaanalisis & ") AND (idproducto = '" & PN & "') AND (idmegacanal = " & idmegacanal & ") AND (idempresa = " & idempresa & ") AND (idtipoventa=" & tipoventa & ")"
						set rsp=oConn.execute(sql)
							if not rsp.eof then 
							'Si extrae un descuento se aplica el valor a descuento_seg.
									descuento_seg=rsp.fields("Descuento")
									trato=rsp.fields("sw_trato")
									if trato="F" Then desctipo_seg="fijo" Else desctipo_seg="variable"
									'if banderas=1 then response.write("<BR>Desc_MEGPROMDESC : " & descuento_seg & "Tipo: " &  desctipo_seg)
							end if 'fin, MEGPROMDESC
					end if 'fin, MAXMCANDESC
			end if 'fin, CANPROMDESC
	end if 'fin, COM_SUCPROMDESC
	
	'Compite el descuento del grupo de CLIENTES y del grupo de SEGMENTO
	'luego se guarda en el descuento principal llamado descuento, descuentomaximo
	'y el tipo del descuento
	
	'Este es el valor original del descuento (0)'
		descuento=0
	
	if (cdbl(descuento_cli)>=cdbl(descuento_seg)) then
		descuento=descuento_cli
		descuentomaximo=descuento_cli
		descuentotipo=desctipo_cli
		'if banderas=1 then response.write("<BR>**CLIENTE : " & descuento & "/ descuentoMaximo: " & descuentomaximo & "/descuentotipo: " & descuentotipo)
		else
			descuento=descuento_seg
			descuentomaximo=descuento_seg
			descuentotipo=desctipo_seg
			'if banderas=1 then response.write("<BR>**SEGMENTO : " & descuento & "/ descuentoMaximo: " & descuentomaximo & "/descuentotipo: " & descuentotipo)
	end if
	
	
'*********************************
'****DESCUENTO POR VENDEDOR*******
'*********************************	

'Se inicia las variables a utilizar en la busqueda del descuento en el grupo de DESCUENTO POR VENDEDOR.
descuento_ven=0
desctipo_ven="variable"

'.:: COM_VENPROMDESC ::.
'Sentencia SQL que trae el descuento.
sql="SELECT     isnull(d.descuento,0) as descuento, d.sw_trato " & _
	"FROM  sqlserver.desaerp.dbo.COM_VENPROMDESC as D " & _
	"inner join sqlserver.desaerp.dbo.dim_vendedores as V on d.idvendedor=v.idvendedor " & _
	"WHERE     (d.vb_aprobacion = 's') AND (d.fecha_hasta >=" & fechaanalisis & ") AND (d.fecha_desde <=" & fechaanalisis & ") AND (d.idproducto = '" & PN & "') AND (d.idempresa = " & idempresa & ") and (v.nombre='" & vend & "') AND (idtipoventa=" & tipoventa & ")"
set rsp=oConn.execute(sql)
	if not rsp.eof then 
	'Si extrae un descuento se aplica el valor a descuento_seg.
			descuento_ven=rsp.fields("Descuento")'descuentomaximo
			trato=rsp.fields("sw_trato")
			if trato="F" Then desctipo_ven="fijo" Else desctipo_ven="variable"
			'if banderas=1 then response.write("<BR>COM_VENPROMDESC : " & descuento_ven & "Tipo: " &  desctipo_ven)
			'Si no encuentra, va a buscar otro descuento.
	else
		'.:: COM_GVTPROMDESC ::.
		'Sentencia SQL que trae el descuento.
		sql=" Select descuento,sw_trato " & _
		"from sqlserver.desaerp.dbo.COM_GVTPROMDESC " & _
		"where idempresa=" & idempresa & " and idgrupoventa=" & id_gurpoventas & " " & _
		"and (vb_aprobacion = 's') AND (fecha_hasta >=" & fechaanalisis & ") AND (fecha_desde <=" & fechaanalisis & ") AND (idproducto = '" & PN & "') AND (idtipoventa=" & tipoventa & ")"
		set rsg=oConn.execute(sql)		
			if not rsg.eof then
			'Si extrae un descuento se aplica el valor a descuento_seg.
				descuento_ven=rsg.fields("Descuento")
				trato=rsg.fields("sw_trato")
				if trato="F" Then desctipo_ven="fijo" Else desctipo_ven="variable"
				'if banderas=1 then response.write("<BR>Desc_GVTPROMDESC. : " & descuento_ven & "Tipo: " &  desctipo_ven)
			else
				'.:: COM_MGVTPROMDESC ::.
				'Sentencia SQL que trae el descuento.
				sql="select descuento, sw_trato " & _
				"from sqlserver.desaerp.dbo.COM_MGVTPROMDESC " & _
				"where idempresa=" & idempresa & " and idmgrupoventa=" & id_mgurpoventas & " " & _
				"and (vb_aprobacion = 's') AND (fecha_hasta >=" & fechaanalisis & ") AND (fecha_desde <=" & fechaanalisis & ") AND (idproducto = '" & PN & "') AND (idtipoventa=" & tipoventa & ")"
				set rsg=oConn.execute(sql)		
					if not rsg.eof then
					'Si extrae un descuento se aplica el valor a descuento_seg.
						descuento_ven=rsg.fields("Descuento")
						trato=rsg.fields("sw_trato")
						if trato="F" Then desctipo_ven="fijo" Else desctipo_ven="variable"
						'if banderas=1 then response.write("<BR>Desc_MGVTPROMDESC. : " & descuento_ven & "Tipo: " &  desctipo_ven)
					end if 'fin, MGVTPROMDESC
			end if 'fin, GVTPROMDESC
	end if 'fin, VENPROMDESC

	'compite el descuento que arrastro anteriormente con el descuento del
	'grupo de vendedores.
	if cdbl(descuento_ven)>=cdbl(descuento) Then 
		descuento=descuento_ven
		descuentomaximo=descuento_ven
		descuentotipo=desctipo_ven
		'if banderas=1 then response.write("<BR>**VENDEDOR : " & descuento & "/ descuentoMaximo: " & descuentomaximo & "/descuentotipo: " & descuentotipo)
	end if
	

'*********************************
'*******DESCUENTO POR SKU*********
'*********************************	

'Se inicia las variables a utilizar en la busqueda del descuento en el grupo de DESCUENTO POR SKU.
descuento_sku=0
desctipo_sku="variable"

'.:: COM_MAXSKUDESC ::.
'Sentencia SQL que trae el descuento.
sql="Select isnull(d.descuento,0) as descuento, d.sw_trato " & _
"from sqlserver.desaerp.dbo.COM_MAXSKUDESC as d " & _
"WHERE (d.vb_aprobacion = 's') AND (d.fecha_hasta >=" & fechaanalisis & ") AND (d.fecha_desde <=" & fechaanalisis & ") AND (d.idproducto = '" & PN & "') AND (d.idempresa = " & idempresa & ") AND (idtipoventa=" & tipoventa & ")"
set rsp=oConn.execute(sql)
	if not rsp.eof then 
	'Si extrae un descuento se aplica el valor a descuento_sku.
		descuento_sku=rsp.fields("Descuento")
		trato=rsp.fields("sw_trato")
		if trato="F" Then desctipo_sku="fijo" Else desctipo_sku="variable"
		'if banderas=1 then response.write("<BR>COM_MAXSKUDESC : " & descuento_sku & "Tipo: " &  desctipo_sku)
	end if

	'compite el descuento que arrastro anteriormente con el descuento del
	'grupo de SKU.
	if cdbl(descuento_sku)>=cdbl(descuento) then 
	descuento=descuento_sku
	descuentomaximo=descuento_sku
	descuentotipo=desctipo_sku
	'if banderas=1 then response.write("<BR>**SKU : " & descuento & "/ descuentoMaximo: " & descuentomaximo & "/descuentotipo: " & descuentotipo)
	end if
	
'if banderas=1 then response.write("<BR>vend : " & vend)
'if banderas=1 then response.write("<BR>id_Mgurpoventas : " & id_mgurpoventas)
'if banderas=1 then response.write("<BR>id_gurpoventas : " & id_gurpoventas)
'if banderas=1 then response.write("<BR>descuentomaximo : " & descuentomaximo)
'if banderas=1 then response.write("<BR>idmegacanal : " & idmegacanal)
'if banderas=1 then response.write("<BR>idcanal : " & idcanal)	

'*********************************
'*****DESCUENTO POR VOLUMEN*******
'*********************************	

'Descuento volumen
sidescvolumen=1 'ESTE ES EL ELEMENTO QUE PERMITE UTILIZAR EL DESCUENTO POR VOLUMEN
sql="SELECT d.descuento, d.cantidad_desde, d.cantidad_hasta " & _
	"FROM sqlserver.desaerp.dbo.COM_ESCDESCENC as e INNER JOIN " & _
    " sqlserver.desaerp.dbo.COM_ESCDESCDET as d ON e.idempresa = d.idempresa AND  " & _
    " e.idcanal = d.idcanal AND e.idproducto = d.idproducto and e.idtipoventa=d.idtipoventa " & _
	"WHERE     (e.idempresa = " & idempresa & ") AND (e.idcanal = " & idcanal & ") AND (e.idproducto = '" & PN & "') " & _
	" and (vb_aprobacion = 's') AND (fecha_hasta >=" & fechaanalisis & ") AND (fecha_desde <=" & fechaanalisis & ") AND (e.idtipoventa=" & tipoventa & ")"
'response.write(sql)
'exit sub

set rsp=oConn.execute(sql)
'response.write(sql)
simatriz=1
if rsp.eof then simatriz=0

'Muestra la Informacion
i=0
do until rsp.eof
	'if banderas=1 then response.write("<BR>" & rsp(0) & " : " & rsp(1) & " : " & rsp(2) )
i=i+1
rsp.movenext
loop

response.write(chr(10) & "<SCRIPT LANGUAGE=" & chr(34) & "JavaScript" & chr(34) & " TYPE=" & chr(34) & "text/javascript" & chr(34) & "> " & chr(10) )
response.write("var Matriz2D=new Array(" & i & ");" & chr(10) )
response.write("for(i=0;i<=" & i & ";i++){Matriz2D[i]=new Array(3);}" & chr(10) )

response.write("sidescvolumen='" &sidescvolumen & "';" & chr(10) )
if simatriz=1 and sidescvolumen=1 then 
	rsp.moveFirst
end if
a=0
do until rsp.eof
	response.write("Matriz2D[" & a & "][0]='" & rsp.fields(0) & "';" & chr(10) )
	response.write("Matriz2D[" & a & "][1]='" & rsp.fields(1) & "';" & chr(10) )
	response.write("Matriz2D[" & a & "][2]='" & rsp.fields(2) & "';" & chr(10) )
a=a+1
rsp.movenext
loop
response.write("var factoralt='" & factoralt & "';" & chr(10) )
response.write("</SCRIPT>" & chr(10) )


	'En ultima instancia compiten el descuento que arrastro anteriormente
	'con el descuento base.
	
	if(cdbl(descuento)=0) then 
		if cdbl(descuento_base)>=cdbl(descuento) then 
			descuento=descuento_base
			descuentomaximo=descuento_base
			descuentotipo="variable"
			'if banderas=1 then response.write("<BR>**BASE : " & descuento & "/ descuentoMaximo: " & descuentomaximo & "/descuentotipo: " & descuentotipo)
		end if
	end if

	'Se identifica que tipo de descuento es (variable o fijo), en caso de ser variable
	'se muestra la calculadora y el descuento maximo NO se muestra, ocurre lo contrario
	'en caso de ser descuento fijo
	if descuentotipo="variable" then
		v1="none"
		descuentotxt=""
		'if banderas=1 then response.write("<BR> tipo: "& descuentotipo)
	else
		v1=""
		descuentotxt=descuento
		'if banderas=1 then response.write("<BR> Else tipo:"& descuentotipo)
	end if
	
	'Se muestra el descuento máximo
	'if banderas=1 then response.write("<BR>" & descuentomaximo)
	
%>
<SCRIPT LANGUAGE="JavaScript">
document.getElementById("dsct").innerHTML='<%= descuento %>';
<!--
//-------------------------------------------------------------------------
function valescalar(){

}
//-------------------------------------------------------------------------
function escalar(){
	cantidad=document.getElementById('cantidad' ).value*1;
	descmax =document.getElementById('descmaxx' ).value;
	descuen =document.getElementById('descuento').value;
	descmax=replace(replace(descmax,',', '.'),'-','');
	descuen=replace(replace(descuen,',', '.'),'-','');
	esc=Matriz2D.length-1;
	if(isNaN(factoralt)){factoralt=1}
	//alert(factoralt);
	factoralt=1;
	for (e=0;e<esc;e++){
		if (cantidad>=(Matriz2D[e][1]*factoralt) && cantidad<=(Matriz2D[e][2]*factoralt) ){
			descuentoadd=(Matriz2D[e][0]);
			descuentoadd=replace(replace(descuentoadd,',', '.'),'-','');
			descuentoadd=descuentoadd*1;
			if(descuentoadd>=descmax)
			{
			descpond=sumadesc(descmax,descuentoadd);
			}
			else
			{
			descpond=descmax;
			}
			
			return descpond;
		}
	}
	//alert(descmax);
	return descmax;
}
//-------------------------------------------------------------------------
function sumadesc(desc_1,desc_2){
//	if (isNaN(desc_1)){
//		alert('NaN Desc1');
//		return 0;
//	}else{
//		if (isNaN(desc_2)){
//			alert('NaN Desc2');
//			return 0;
//		}else{
//			desc_1=desc_1*1
//			desc_2=desc_2*1
//			return desc_1+(desc_2*(100-desc_1)/100);
//		}
//	}
	return desc_2; // devuelve el descuento maximo por producto
	entelpcs

}
//-------------------------------------------------------------------------
function solonumeros(){
		//       0  1  2  3  4  5  6  7  8  9                (Enter)
	if ( (event.keyCode > 47 && event.keyCode < 58) || (event.keyCode==13) ){
		//alert(event.keyCode);
	}else{
		event.returnValue = false;
	}
}
//-------------------------------------------------------------------------
function solonumeros_desc(){
		//       0 1 2 3 4 5 6 7 8 9      -  .  ,            (Enter)
	if ( (event.keyCode > 43 && event.keyCode < 58) || (event.keyCode==13) ){
		//alert(event.keyCode);
	}else{
		event.returnValue = false;
	}
}
//-------------------------------------------------------------------------
function validadescuentos(){
	//alert(factoralt);
	cantida=document.getElementById('cantidad' ).value;
	descmax=document.getElementById('descmaxx' ).value;
	descuen=document.getElementById('descuento').value;
	descmax=replace(replace(descmax,',', '.'),'-','');
	descuen=replace(replace(descuen,',', '.'),'-','');
	
	//alert(descuen);
	if (cantida==''){
		alert("Tiene que Dijitar una cantidad");
		return 0;
	}
	if (descuen*1){descuen=descuen*1;}
	if (isNaN(descuen)){ 
		alert('Descuento Mal Dijitado \n');
		return 0;
	}else{
		descuen=Math.abs(descuen);
	}
	if (isNaN(cantida)){ 
		alert('Cantidad invalida');
		return 0;
	}
	if (descuen<0){descuen=descuen*-1}
	//alert(descmax);
	descmax=escalar();
	//alert('descuen:'+Math.abs(descuen)+' - descmax:'+descmax);
	//alert(descmax);
	//alert(descmax);
	if ( (Math.abs(descuen)*1)>(descmax*1) ) {
		alert("El descuento supera el maximo permitido");
		return 0;
	}
	if (validacantidad()=='OK'){
		//ok
	}else{
		alert(validacantidad() );
		return 0;
	}

	document.frmpn.submit()
}
//-------------------------------------------------------------------------
function validacantidad(){
	mireturn="OK"
	cantidad=document.getElementById('cantidad').value*1;
	uempaque=document.getElementById('uempaque').value*1;
	ucaja   =document.getElementById('ucaja'   ).value*1;
	ubase   =document.getElementById('ubase'   ).value*1;
	upallet =document.getElementById('upallet' ).value*1;
	ccerrada=document.getElementById('ucajacerrada' ).value;	
	upedminimo=document.getElementById('upedminimo' ).value;	
	
	//alert(ucaja);
	//alert(ubase);
	//nada
	if ( uempaque ==0 ){
		mireturn="OK";
	}
	//unidad
	if ( uempaque ==1 ){
		mireturn="OK";
	}
	//caja
	if ( uempaque ==2 && ucaja>0){
		if (cantidad<(ucaja) ){
			mireturn="Minimo Venta "+(ucaja)+" Unidades";
		}else{
			dif = cantidad % (ucaja);
			if (dif==0){
				mireturn="OK";
			}else{
				mireturn="Venta en multiplos "+(ucaja)+" Unidades";
			}
		}
	}
	//base
	if ( uempaque ==3 && ubase>0 ){
		if (cantidad<(ubase*ucaja) ){
			mireturn="Minimo Venta "+(ubase*ucaja)+" Unidades";
		}else{
			dif = cantidad % (ubase*ucaja);
			if (dif==0){
				mireturn="OK";
			}else{
				mireturn="Venta en multiplos "+(ubase*ucaja)+" Unidades";
			}
		}
	}
	//pallet
	if ( uempaque ==4 && upallet>0 ){
		if (cantidad<(ubase*ucaja*upallet) ){
			mireturn="Minimo Venta "+(ubase*ucaja*ubase)+" Unidades";
		}else{
			dif = cantidad % (ubase*ucaja*ubase);
			if (dif==0){
				mireturn="OK";
			}else{
				mireturn="Venta en multiplos "+(ubase*ucaja*ubase)+" Unidades";
			}
		}
	}
	
	if(upedminimo=="S")
	{
		if(ccerrada=="S")
		{
			if(cantidad%ucaja!=0)
			{
				mireturn="La cantidad debe ser multiplo de "+ucaja ;
			}
		}
	}
	//mireturn="OK";
	return mireturn;
}
//-------------------------------------------------------------------------
function CalcEmpaque(){
	cantidad=document.getElementById('cantempaque').value;
	uempaque=document.getElementById('uempaque').value;
	ucaja   =document.getElementById('ucaja'   ).value;
	ubase   =document.getElementById('ubase'   ).value;
	upallet =document.getElementById('upallet' ).value;
	var factor=1;
	if (uempaque==1){factor=1};
	if (uempaque==2){factor=ucaja};
	if (uempaque==3){factor=ubase*ucaja};
	if (uempaque==4){factor=upallet*ubase*ucaja};
	document.getElementById('cantidad').value=factor*cantidad;//Muestra unidades de venta
	unidades=factor*cantidad;
	//----
	var msg ='';
	var dif =0 ;
	var pallc=0; //pallets cerrados
	var basec=0; //bases cerradas
	var cajac=0; //cajas cerradas
	var saldo=0;

	if (unidades>=(upallet*ubase*ucaja)) { //si es mas que un pallet
		dif=unidades%(upallet*ubase*ucaja)
		pallc=parseInt(unidades/(upallet*ubase*ucaja));
		msg=(pallc + ' Pallets');
		saldo=unidades-(pallc*upallet*ubase*ucaja);
	}else{
		saldo=unidades;
	}
	
	if (saldo>=(ubase*ucaja)) { //Si es mas que una base
		if (saldo<unidades){msg=msg+" , "}
		basec=parseInt(saldo/(ubase*ucaja)) 
		msg = msg + basec + " Bases";
	}
	if ((saldo%(ubase*ucaja))!=0){//Si es mas que una caja
		saldo=saldo-(basec*ubase*ucaja);
		if (basec>0){msg=msg+" , "}
		msg = msg + (saldo/ucaja) + " Cajas";
	}
	document.getElementById('cextra').value=msg;

}
//-------------------------------------------------------------------------
function replace(str,txta,txtb){
	str=String(str).replace(txta,txtb);
	str=String(str).replace(txta,txtb);
	str=String(str).replace(txta,txtb);
	str=String(str).replace(txta,txtb);
	return String(str).replace(txta,txtb);
}
//-------------------------------------------------------------------------
function Length(str){return String(str).length;}
//-------------------------------------------------------------------------

//-->
</SCRIPT>
<%
'.:: . . . . . . . V A L I D A . . C A N T I D A D E S . . . . . . . . . ::.
'productoactivo
idcl=split(cliente," ")
idcliente=idcl(0)
'iduempaque=consultarapida("select iduempaque from sqlserver.desaerp.dbo.com_cliuempaque where idempresa='1' and idcliente='" & idcliente & "' and idsku='" & NP & "'")
sqlu="select iduempaque from sqlserver.desaerp.dbo.com_cliuempaque where idempresa='" & idempresa & "' and idcliente='" & idcliente & "' and idsku='" & PN & "'"
'response.write(sqlu)
iduempaque=0
set rsu=oConn.execute(sqlu)
if rsu.eof then
	sqlu="select iduempaque from sqlserver.desaerp.dbo.com_cliuempaque where idempresa='" & idempresa & "' and idcliente='*' and idsku='" & PN & "'"
	set rsu=oConn.execute(sqlu)
end if
do until rsu.eof
	iduempaque=rsu("iduempaque")
rsu.movenext
loop

sqlu="select factoralt, isnull(analisisproducto8,0) as analisisproducto8, isnull(analisisproducto9,0) as analisisproducto9 from serverdesa.BDFlexline.flexline.producto where empresa='desa' and producto='" & PN & "'"
sqlu="select factoralt, isnull(analisisproducto8,0) as analisisproducto8, isnull(analisisproducto9,0) as analisisproducto9 from handheld.flexline.producto where empresa='desa' and producto='" & PN & "'"
'iduempaque=0
set rsu=oConn.execute(sqlu)
do until rsu.eof
	ucaja=rsu("factoralt")
	ubase=rsu("analisisproducto8")
	upallet=rsu("analisisproducto9")
rsu.movenext
loop
'response.write(ucaja)

if iduempaque=1 then empaque="Unidades"
if iduempaque=2 then empaque="Cajas"
if iduempaque=3 then empaque="Bases"
if iduempaque=4 then empaque="Pallets"

if(idempresa=1 or idempresa=4) then
sqlcajacerrada="SELECT [caja_cerrada] FROM [DesaERP].[dbo].[DIM_FAMILIAS] WHERE nombre='"&familia&"' and idempresa=" & idempresa
'response.write(sqlcajacerrada)
rcc=oConn.execute(sqlcajacerrada)
cajacerrada=rcc("caja_cerrada")
'response.write(cajacerrada)
else
cajacerrada="N"
end if


%>
<TABLE>
<%
if len(empaque)>0 then
	%>
	<TR>
		<TD><FONT SIZE='2' face='Arial' COLOR='#000066'><B><%=empaque%></B></FONT></TD>
		<TD align="left"><INPUT TYPE="cantempaque" NAME="cantempaque" id="cantempaque" size="8" onKeypress="solonumeros()" onKeyup="valescalar();CalcEmpaque()"></TD>
		<TD><INPUT TYPE="text" NAME="cextra" id="cextra" size="27" style="border-width:0;background-color:#F2F2F7"></TD>
	</TR>
	<%
end if
%>
<TR>
	<TD><FONT SIZE='2' face='Arial' COLOR='#000066'>Cantidad</FONT></TD>
	<TD align="right"><INPUT TYPE="text" NAME="cantidad" id="cantidad" size="8" onKeypress="solonumeros()" onKeyup="valescalar()"></TD>
	<TD></TD>
</TR>
<TR>
	<TD><FONT SIZE='2' face='Arial' COLOR='#000066'>Descuento</FONT></TD>
	<TD align="right">
		<INPUT TYPE="hidden" id="descfijo"  NAME="descfijo" value="<%=desctipo%>">
		<INPUT TYPE="text"   id="descuento" NAME="descuento" size="8" value="<%=descuentotxt%>" onKeypress="solonumeros_desc()">
	</TD>
	<TD><INPUT TYPE="button" value=">" onClick="calldescuento2()"></TD>
</TR>
<TR>
	<TD><INPUT TYPE="hidden" id="descmaxx" value="<%=descuentomaximo%>"></TD>
	<TD><INPUT TYPE="button" id="guardalinea" name="guardalinea" value="Guardar" onClick="validadescuentos()"></TD>
	<TD></TD>
</TR>
</TABLE>

<!--  <HR>
<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
<TABLE>
	<TR>
		<TD><FONT SIZE='2' face='Arial' COLOR='#000066'><B><%=empaque%></B></FONT></TD>
		<TD align="left"><INPUT TYPE="cantempaque" NAME="cantempaque" id="cantempaque" size="8" onKeypress="solonumeros()" onKeyup="valescalar();CalcEmpaque()"></TD>
		<TD><INPUT TYPE="text" NAME="cextra" id="cextra" size="27"></TD>
	</TR>
<TR>
	<TD><FONT SIZE='2' face='Arial' COLOR='#000066'>Unidades</FONT></TD>
	<TD align="left"><INPUT TYPE="text" NAME="cantidad1" id="cantidad1" size="8" onKeypress="solonumeros()" onKeyup="valescalar()"></TD>
	<TD></TD>
</TR>
</TABLE>   -->
<%
'.:: . . . . . . . V A L I D A . . C A N T I D A D E S . . . . . . . . . ::. parte B

%>
<INPUT TYPE="hidden" name="uempaque" id="uempaque" value="<%=iduempaque%>" >
<INPUT TYPE="hidden" name="ucaja"    id="ucaja"    value="<%=ucaja%>" >
<INPUT TYPE="hidden" name="ubase"    id="ubase"    value="<%=ubase%>" >
<INPUT TYPE="hidden" name="upallet"  id="upallet"  value="<%=upallet%>" >
<INPUT TYPE="hidden" name="ucajacerrada"  id="ucajacerrada"  value="<%=cajacerrada%>" >
<INPUT TYPE="hidden" name="upedminimo"  id="upedminimo"  value="<%=pedmin%>" >

<%
if miboton="Agregarmult" then
	for each campo in request.Form
		if left(campo,3)="ck_" then
			%><INPUT TYPE="hidden" name="<%=campo%>" id="<%=campo%>" value="<%=request.form(campo)%>" ><%
		end if
	next
	'id_Nro_ck
	%><INPUT TYPE="hidden" name="Nro_ck" id="Nro_ck" value="<%=cint(id_Nro_ck)+1%>" ><%
	%><INPUT TYPE="hidden" name="boton" id="boton" value="Agregarmult" ><%
	%><INPUT TYPE="hidden" name="sw_validarepetido" id="sw_validarepetido" value="0" ><%
end if
%>
</FORM><!-- fin formulario productos -->

<%
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ FIN

'.:: Flete Neto ::.

'********** Buscar fletes en el servidor ********************************
'SQL="SELECT * FROM Flexline.lpfletes WHERE (Codigo = '" & PN & "')"
'sql="SELECT *, valorflete as Fleteneto FROM serverdesa.BDFlexline.flexline.PE_TarifasFlete where idtipocliente='" & idtipocliente & "' and idproducto='" & PN & "'"'
sql="SELECT *, valorflete as Fleteneto FROM HandHeld.flexline.PE_TarifasFlete where idtipocliente=" & idtipocliente & " and idproducto='" & PN & "'"
'response.write(sql)
set rs=oConn.execute(sql)

	if not rs.eof then
		fleteneto=cdbl(rs.fields("fleteneto"))
	else
		if empresa<>"DESAZOFRI" then
			'fleteneto=cdbl(70)
			%>
			<TABLE border="0">
			<TR>
				<TD  bgcolor="#336699" align="Center"><FONT face="VERDANA" SIZE="1" COLOR="#FFFFFF">.:: Alerta ::.</FONT></TD>
			</TR>
			<TR>
				<TD bgcolor="#EEEEEE" align="Center">&nbsp;<B><FONT face="VERDANA" SIZE="1" COLOR="#330000">PRODUCTO SIN FLETE ASIGNADO</FONT></B>&nbsp;</TD>
			</TR>
			<TR>
				<TD bgcolor="#EEEEEE" align="Center"><INPUT TYPE="button" value="<< Volver" onClick="history.back()" ></TD>
			</TR>
			</TABLE>
			
			<SCRIPT LANGUAGE="JavaScript">
			<!--
			document.getElementById('guardalinea').style.display='none';
			//-->
			</SCRIPT>
			<%
		end if
		fleteneto=cdbl(70)
		'exit sub
	end if
	if empresa="DESAZOFRI" then fleteneto=0
	'response.write("<BR>Fleteneto : " & rs.fields("fleteneto"))
	ilap=15
	if Ucase(left(PN,1))="C" then ilap=20.5
	if Ucase(left(PN,1))="V" then ilap=20.5
	if Ucase(left(PN,1))="W" then ilap=31.5
	if Ucase(left(PN,1))="L" then ilap=31.5
	if Ucase(left(PN,1))="P" then ilap=31.5
	if Ucase(left(PN,1))="B" then ilap=18
	if Ucase(PN)="BE80061" then ilap=10
	if Ucase(PN)="BE80066" then ilap=10
	if Ucase(PN)="BE80067" then ilap=10
	if Ucase(PN)="BE80068" then ilap=10
	if Ucase(PN)="BE80102" then ilap=10
	if Ucase(left(PN,1))="J" then ilap=0
	if Ucase(left(PN,1))="A" then ilap=0
	if Ucase(left(PN,2))="MP" then ilap=0
	if Ucase(left(PN,2))="GN" then ilap=0
on error resume next

	mitotalp=Minetox
	tneto=cdbl(Minetox) + fleteneto + ((cdbl(Minetox) + fleteneto)*0.19)+((cdbl(Minetox) )*(ilap/100)) 
	Mitotalp=tneto
	Factor_IVA="19"
	
if empresa="DESAZOFRI" then 
	if tipodocto="008" then 
		'iva = (cdbl(afeciva)* 19 /100)
		if sw_iva ="on" then
			'iva = (cdbl(afeciva)* 19 /100)
		else
			tneto=cdbl(Minetox) + fleteneto + ((cdbl(Minetox) + fleteneto)*0.006)+((cdbl(Minetox) )*(ilap/100)) 
			Mitotalp=tneto
			Factor_IVA="006"
			'iva = (cdbl(afeciva)* 0.06 /100)
			'txt_iva="0,6"
		end if
	End if
end if


	%>
	
<script language=JavaScript type=text/JavaScript>
<!--
	function calldescuento() {
		var cero='', total=0, x=0, descuento=0, fleteneto=<%=fleteneto%>, ilap=<%=replace(ilap, ",", ".")%>, neto=<%=minetox%>;
		total=document.getElementById('totalventa').value;
		//alert(total);
		x=(total-(fleteneto*1.<%=Factor_IVA%>))*100/(1<%=Factor_IVA%>+ilap);
		//alert(fleteneto);
		if(Math.round(<%=Factor_IVA%>)==6){
			x=(total)*100/((1<%=Factor_IVA%>/10)+ilap);
			//alert(ilap);
			//alert (1<%=Factor_IVA%>);
			//alert ((1<%=Factor_IVA%>/10)+ilap);
			//alert (x);
		}
	//	if (Math.round(fleteneto)==0){
		//	x=(total)*100/(1<%=Factor_IVA%>+ (ilap));
		//	alert(ilap)
	//		alert(x);
	//		alert(neto);
	//	}
		//x=(total-(fleteneto*1.19))*100/(119+ilap);
		//alert(<%=Factor_IVA%>);
		//alert(x);
		descuento=100-(x *100/neto);
		document.getElementById('descuento').value=Math.round(descuento*100)/100 ;
	}
	//--------------------------------------------------------------------------------------
	function calldescuento2() {
		var cero='', total=0, x=0, descuento=0, descuentop='eeee', fleteneto=<%=fleteneto%>, ilap=<%=replace(ilap, ",", ".")%>, neto=<%=minetox%>;
		descuento=document.getElementById('descuento').value;
		descuento=descuento.replace(',','.');
		descuento=descuento.replace('-','');
		x=((neto-(neto*descuento/100))*1.<%=Factor_IVA%>)+(fleteneto*1.<%=Factor_IVA%>)+((neto-(neto*descuento/100))*(ilap/100));
		//x=((neto-(neto*descuento/100))*1.19)+(fleteneto*1.19)+((neto-(neto*descuento/100))*(ilap/100));
		document.getElementById('totalventa').value=Math.round(x) ;
		//document.getElementById('descuentod').style.display='none';
		//document.getElementById('descuento').disabled = 'disabled';
		//disabled=”disabled”
		//document.getElementById('descuento').type = 'hidden';
	}
	//--------------------------------------------------------------------------------------
	function detalle() {
		var cero='', total=0, x=0, descuento=0, descuentop='', fleteneto=<%=fleteneto%>, ilap=<%=replace(ilap, ",", ".")%>, neto=<%=minetox%>;
		descuento=document.getElementById('descuento').value;
		descuento=descuento.replace(',','.');
		descuento=descuento.replace('-','');
		x=(neto-(neto*descuento/100));
		//fiva=<%=Factor_IVA%>;
		fiva=<%=Factor_IVA%>;
		//varcero='';
		
		if (Math.round(fiva)==6){fiva='0.06'}
		//descuentop=descuentop + 'Neto.....      \t: ' + neto+'\n';
		descuentop=descuentop + 'Neto-desc..     \t: ' + Math.round(x)+'\n';
		descuentop=descuentop + 'Flete.....      \t: ' + fleteneto+'\n';
		descuentop=descuentop + 'ILA.(' + ilap + '%) \t: ' + Math.round((x*(ilap/100)))+'\n';
		//alert(Math.round(fiva));
		if (fiva=='0.06'){
			descuentop=descuentop + 'IVA.(' + fiva + '%) \t: ' + Math.round((x+fleteneto)*0.0<%=Factor_IVA%>)+'\n';
		}else{
			descuentop=descuentop + 'IVA.(' + fiva + '%) \t: ' + Math.round((x+fleteneto)*0.<%=Factor_IVA%>)+'\n';
		}
		
		descuentop=descuentop + '===================\n';
		window.alert (descuentop);
		//window.alert(x);
		//window.alert(fleteneto);
		//window.alert(0.<%=Factor_IVA%>);
	}
//-->
</SCRIPT>

<%
'response.write(empresa)
if empresa="DESAZOFRI" then v1="none"
if v1="none" then 'Ocultar calculadora%>
	<CENTER><HR>
	<TABLE border="0" bgcolor="#CCCCCC" cellspacing="1" cellpadding="1" style="border: 2px #23214E solid;">
	<TR>
		<TD><B><FONT SIZE="2" face="verdana" COLOR="#000033">&nbsp;Calculadora Descuento&nbsp;</FONT></B></TD>
	</TR>
	<TR>
		<TD><CENTER><FONT SIZE="2" face="verdana" COLOR="#000033">Precio Venta Unitario</FONT></CENTER></TD>
	</TR>
	<TR>
		<TD><CENTER><INPUT ID="totalventa" TYPE="text" NAME="" value="<%=round(Mitotalp)%>"></CENTER></TD>
	</TR>
	<TR>
		<TD><CENTER><input type="button" value="Calcular" name="B3" onClick="calldescuento()">
		<input type="button" value="Detalle" name="B5" onClick="detalle()"></CENTER></TD>
	</TR>
	</TABLE>
	<INPUT id="tneto"  TYPE="hidden" value="<%=tneto%>">
	</CENTER><BR><BR><BR><BR><BR><%
	End if 'Ocultar calculadora
	'end if 'if eof flete

'end if 'fin comprobacion por coronita
end Sub 'prodetalle()

'--------------------------------------------------------------------------------------------------
Sub editaprod()
x=0
For Each elemento In request.form
	IF LEFT(elemento,1) = "k" THEN
		if elemento="k01" then PN=p01
		if elemento="k02" then PN=p02
		if elemento="k03" then PN=p03
		if elemento="k04" then PN=p04
		if elemento="k05" then PN=p05
		if elemento="k06" then PN=p06
		if elemento="k07" then PN=p07
		if elemento="k08" then PN=p08
		if elemento="k09" then PN=p09
		if elemento="k10" then PN=p10
		if elemento="k11" then PN=p11
		if elemento="k12" then PN=p12
	x=x+1
	END IF
Next

if x>0 then 
	call borraprod()
	call prodetalle()
else
	call mustraventa()
end if
end Sub 'editaprod()

'---------------------------------------------------------------------------------------------------
Sub sacaprod()
	call borraprod()
	call mustraventa()
end Sub 'sacaprod()

'---------------------------------------------------------------------------------------------------
Sub borraprod()
For Each elemento In request.form
	IF LEFT(elemento,1) = "k" THEN
		if elemento = "k01" then
			p01=""
			c01=""
			d01=""
		end if
		if elemento = "k02" then
			p02=""
			c02=""
			d02=""
		end if
		if elemento = "k03" then
			p03=""
			c03=""
			d03=""
		end if
		if elemento = "k04" then
			p04=""
			c04=""
			d04=""
		end if
		if elemento = "k05" then
			p05=""
			c05=""
			d05=""
		end if
		if elemento = "k06" then
			p06=""
			c06=""
			d06=""
		end if
		if elemento = "k07" then
			p07=""
			c07=""
			d07=""
		end if
		if elemento = "k08" then
			p08=""
			c08=""
			d08=""
		end if
		if elemento = "k09" then
			p09=""
			c09=""
			d09=""
		end if
		if elemento = "k10" then
			p10=""
			c10=""
			d10=""
		end if
		if elemento = "k11" then
			p11=""
			c11=""
			d11=""
		end if
		if elemento = "k12" then
			p12=""
			c12=""
			d12=""
		end if
	end if
next

end Sub 'borraprod()
'--------------------------------------------------------------------------------------------------
sub MIERROR()
%>
<p align="center"><font size="+2">ERROR: Ingrese Productos a nota de Venta</font></p>
<form method="post" action="">
		<p align="center"><input name="boton" type="submit" value="Agregar"></p>
</form>
<form method="post" action="../Default.asp">
		<p align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;Salir&nbsp;&nbsp;&nbsp;&nbsp;"></p>
</form>
<form method="post" action="noventa/noventa.asp?vend=<%= ucase(vend) %>&cliente=<% =cliente%>">
		<p align="center"><input type="submit" value="Realizar No Venta"></p>
</form>
<%
end sub 'MIERROR()
'--------------------------------------------------------------------------------------------------
Sub mustraventa()
 
 SQL="SELECT CtaCte.CodLegal, CtaCte.RazonSocial, CtaCteDirecciones.Direccion, CtaCte.CtaCte, CtaCte.Giro, CtaCte.CondPago, CtaCte.Comuna, CtaCte.Telefono, CtaCte.LimiteCredito, CtaCte.PorcDr1, ctacte.analisisctacte1, ctacte.analisisctacte4 FROM flexline.CtaCte CtaCte, flexline.CtaCteDirecciones CtaCteDirecciones WHERE CtaCteDirecciones.CtaCte = CtaCte.CtaCte AND CtaCteDirecciones.Empresa = CtaCte.Empresa AND CtaCteDirecciones.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='" & empresa & "') AND (CtaCte.TipoCtaCte='CLIENTE') AND (CtaCteDirecciones.CtaCte='" & cliente & "') AND (CtaCteDirecciones.Principal<>'s')) "
'on error resume next

Set rs=oConn.execute(Sql)
'************************************
if isnull(rs.fields("porcDR1")) then
	porcDR1=0
else
	porcDR1=cdbl(replace(rs.fields("porcDR1"),".",","))
end if
'************************************

%>
<FORM METHOD=POST ACTION="?vend=<%=ucase(vend)%>&cliente=<%=cliente%>&fechaentrega=<%=fechaentrega%>&id_porfolio=<%=id_porfolio%>&empresa=<%=empresa%>&tipodocto=<%=tipodocto%>&sw_iva=<%=sw_iva%>&p01=<%=p01%>&c01=<%=c01%>&d01=<%=d01%>&p02=<%=p02%>&c02=<%=c02%>&d02=<%=d02%>&p03=<%=p03%>&c03=<%=c03%>&d03=<%=d03%>&p04=<%=p04%>&c04=<%=c04%>&d04=<%=d04%>&p05=<%=p05%>&c05=<%=c05%>&d05=<%=d05%>&p06=<%=p06%>&c06=<%=c06%>&d06=<%=d06%>&p07=<%=p07%>&c07=<%=c07%>&d07=<%=d07%>&p08=<%=p08%>&c08=<%=c08%>&d08=<%=d08%>&p09=<%=p09%>&c09=<%=c09%>&d09=<%=d09%>&p10=<%=p10%>&c10=<%=c10%>&d10=<%=d10%>&p11=<%=p11%>&c11=<%=c11%>&d11=<%=d11%>&p12=<%=p12%>&c12=<%=c12%>&d12=<%=d12%>">  
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-width: 0" width="100%">
  <tr>
    <td width="100%" height="1" align="center" style="border-style: none; border-width: medium">
    <font face="Arial" size="2" color="#333333"><B>Sistema Preventa</B></font></td>
  </tr>
    <tr>
    <td width="100%" height="1" align="center" style="border-style: none; border-width: medium">
    <font size="2" face="Arial"></font></td>
  </tr>
  <tr>
    <td width="100%" height="13" align="left" style="border-style: none; border-width: medium">
<font size="1" face="Arial" COLOR="#808080">Cliente:</FONT><BR>
<font size="2" face="Arial" COLOR="#000000"><% = replace(rs.fields("razonsocial")," ","&nbsp;") %></FONT></b><BR>
<font size="1" face="Arial" COLOR="#808080">Rut:</FONT><BR>
<font size="2" face="Arial" COLOR="#000000"><% = replace(rs.fields("ctacte")," ","&nbsp;") %></FONT><BR>
<font size="1" face="Arial" COLOR="#808080">Direccion:</FONT><BR>
<font size="2" face="Arial" COLOR="#000000"><% = replace(rs.fields("direccion")," ","&nbsp;") %></FONT><BR>
<font size="1" face="Arial" COLOR="#808080">Canal:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cond Pago:</FONT><BR>
<font size="2" face="Arial" COLOR="#000000"><% = replace(rs.fields("analisisctacte1") & "&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;&nbsp;" & rs.fields("CondPago")," ","&nbsp;") %></FONT><BR>	
	
<!-- Botones -->
<INPUT TYPE="submit" value="Agregar"   name="boton">
<INPUT TYPE="submit" value="Edit"     name="boton">
<INPUT TYPE="submit" value="Borrar"    name="boton">
<INPUT TYPE="submit" value="Finalizar" name="boton">

	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td width="100%" colspan="4" align="center" bgcolor="#666633"><b>
        <font size="2" face="Arial" color="#FFFFFF">Glosa Producto</font></b></td>
      </tr>
      <tr>
        <td width="10%" align="center" bgcolor="#666633"><b>
			<font size="2" face="Arial" color="#FFFFFF">ck</font></b>
		</td>
        <td width="30%" align="center" bgcolor="#666633"><b>
			<font size="2" face="Arial" color="#FFFFFF">Cant.</font></b>
		</td>
        <td width="30%" align="center" bgcolor="#666633"><b>
			<font size="2" face="Arial" color="#FFFFFF">Desc%</font></b>
		</td>
        <td width="30%" align="center" bgcolor="#666633"><b>
			<font size="2" face="Arial" color="#FFFFFF">Precio</font></b>
		</td>
      </tr>
<%
bgcolor="#99CCFF"
if len(p01)>0 then call muestralineaproducto(p01,c01,d01,listprec,"01")
if len(p02)>0 then call muestralineaproducto(p02,c02,d02,listprec,"02")
if len(p03)>0 then call muestralineaproducto(p03,c03,d03,listprec,"03")
if len(p04)>0 then call muestralineaproducto(p04,c04,d04,listprec,"04")
if len(p05)>0 then call muestralineaproducto(p05,c05,d05,listprec,"05")
if len(p06)>0 then call muestralineaproducto(p06,c06,d06,listprec,"06")
if len(p07)>0 then call muestralineaproducto(p07,c07,d07,listprec,"07")
if len(p08)>0 then call muestralineaproducto(p08,c08,d08,listprec,"08")
if len(p09)>0 then call muestralineaproducto(p09,c09,d09,listprec,"09")
if len(p10)>0 then call muestralineaproducto(p10,c10,d10,listprec,"10")
if len(p11)>0 then call muestralineaproducto(p11,c11,d11,listprec,"11")
if len(p12)>0 then call muestralineaproducto(p12,c12,d12,listprec,"12")
%>
	</table>
</font></td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    </tr>
  </table>
  <INPUT TYPE="reset" value="&nbsp;">
  <font size="1" face="Arial" COLOR="#000099">Borrar seleccion multiple  </FONT>
</FORM>
<p align="center"><font face="Arial" size="1" color="#808080"></font></p>

</body>

</html>
<%
end sub 'mustraventa()
'--------------------------------------------------------------------
sub muestralineaproducto(vx_producto,vx_cantidad,vx_descuento,listprecio,id)
on error resume next
	if len(vx_descuento)=0 then vx_descuento=0
	if bgcolor<>"#99CCFF" then
		bgcolor="#99CCFF"
	else
		bgcolor="#ECF1FB"
	end if
	sql="SELECT P.PRODUCTO, P.GLOSA, P.FAMILIA, D.Valor FROM flexline.ListaPrecio L, flexline.ListaPrecioD D, flexline.PRODUCTO P WHERE D.Empresa = L.Empresa AND D.Empresa = P.EMPRESA AND D.IdLisPrecio = L.IdLisPrecio AND D.Producto = P.PRODUCTO AND ((P.EMPRESA='" & empresa & "') AND (P.TIPOPRODUCTO='EXISTENCIA') AND (P.PRODUCTO='" & vx_producto & "') AND (L.LisPrecio='" & listprecio & "'))"
	Set rs=oConn.execute(Sql)
	if rs.eof then
		response.write("<tr><tD colspan='4'>Producto no esta en Lista precio " & listprecio & "</td></tr>")
		exit sub
	end if
	'response.write("<tr><tD colspan='4'>Producto Lista precio " & listprecio & "</td></tr>")
	%><tr>
	<td width='100%' colspan='4' align='center' bgcolor='<%=bgcolor%>'><%=replace(rs.fields("glosa")," ","&nbsp;")%></td>
	</tr>
	<tr>
	<td width='10%' align='center' bgcolor='<%=bgcolor%>'><INPUT TYPE='radio' NAME='k<%=id%>'></td>
	<td width='30%' align='center' bgcolor='<%=bgcolor%>'><%=vx_cantidad %></td>
	<td width='30%' align='center' bgcolor='<%=bgcolor%>'><%=vx_descuento%></td>
	<td width='30%' align='center' bgcolor='<%=bgcolor%>'><%=rs("valor") %></td>
	</tr>
	<tr><%
End sub 'muestralineaproducto
'--------------------------------------------------------------------
sub asignafecha(DIA)

'aqui modifica fecha
'Jueves
FC2=dateadd("d",4,date) 'normal 3
'Viernes
FC3=dateadd("d",3,date) 'normal 2
with response
select case DIA
   case 1
 .write "<td><input id='lun' name='fE' type='radio' value='"&_
   year(cdate(FC2)) & right("00" & month(cdate(FC2)),2) &_
    right("00" & day(cdate(FC2)),2) & "' checked></td>"
 .write "<td><label style=""text-transform:capitalize"">"& weekdayname(weekday(FC2)) &"</label></td>" 

   case 2
 .write "<td><input id='lun' name='fE' type='radio' value='"&_
   year(cdate(FC3)) & right("00" & month(cdate(FC3)),2) &_
    right("00" & day(cdate(FC3)),2) & "' checked></td>"
 .write "<td><label style=""text-transform:capitalize"">"& weekdayname(weekday(FC3)) &"</label></td>"

end select
end with
end sub 
'--------------------------------------------------------------------------------------------
Function PC(sString, mzise) 
  Dim sWhiteSpace, bCap, iCharPos, sChar 
  sWhiteSpace = Chr(32) & Chr(9) & Chr(13) 
  sString = LCase(sString)
  bCap = True 
  For iCharPos = 1 to Len(sString) 
    sChar = Mid(sString, iCharPos, 1) 
    If bCap = True Then 
      sChar = UCase(sChar) 
    End If 
    ProperCase = ProperCase + sChar 
    If InStr(sWhiteSpace, sChar) Then 
      bCap = True 
    Else 
      bCap = False 
    End If 
  Next  
PC="<FONT SIZE='" & Mzise & "' face='arial' COLOR='#000033'>" & replace(ProperCase," ","&nbsp;") & "</FONT>"
PC=replace(replace(PC,"%20"," "),"***",chr(34))
if mzise=0 then PC=ProperCase
End Function 
'--------------------------------------------------------------------------------------------

%>
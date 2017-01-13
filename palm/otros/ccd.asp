<%':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql, oconn1, rs1

Set oConn = server.createobject("ADODB.Connection")
Set oConn1 = server.createobject("ADODB.Connection")

oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey"

oConn1.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey"

'recupera datos
vendedor=request.querystring("nuser")
if len(vendedor)=0 then vendedor =trim(cstr(request.form("vendedor")))
B1 =trim(cstr(request.form("B1" )))
%>
<html>
<head>
<meta name="description" content="vers 3"/>
<title>Ingresos CCD</title>
</head>
<style type="text/css">
<!--
.Estilo1 {
	font-family: Arial;
	font-size: 12px;
}
.Estilo2 {color: #FFFFFF}
.Estilo4 {font-weight: bold}
.Estilo5 {font-weight: bold}
-->
</style>
<body>
<form method="post" action="">
<%
if B1 =	""					then call paso01()
if B1 =	"Cancelar"         	then call paso01()
if B1 =	"Agregar"	        then call paso02()
if B1 =	"Ingresa Doc Pago"	then call Ingresopago()
if B1 =	"Aceptar"     	    then call paso03()
if B1 =	"Acepta Pago"   	then call resumen()
if B1 =	"Traspasar"        	then call paso04()      'Guarda datos en servidor

':::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::
Sub Paso04()

'---------------------------------------
'--------Verifica Cheque al Día --------
'---------------------------------------
Cuenta = request.form("Cuenta")
if request.form("Tipopago") = "Cheque" then
	if cdate(request.form("vence")) <= date then'DateAdd("d", 5, Date) Then
'response.write(cdate(request.form("vence")))
	Cuenta = "1104005"
	end if
end if
if  request.form("Tipopago") = "Factura" then
	Cuenta = "1104001"
end if
'---------------------------------------
'---------------------------------------
Tipopago = request.form("Tipopago")
Sql2="SELECT RazonSocial FROM CtaCte "&_
"WHERE (TipoCtaCte = 'cliente') AND (Empresa = 'DESA') AND (CodLegal = '"& request.form("Cliente") &"')"
set rs2=oConn.execute(Sql2)

Sql1="SELECT nombre FROM Flexline.FX_vende_man WHERE (numero = "& trim(vendedor) & ")"
set rs1=oConn1.execute(Sql1)
'response.write(Sql1)

Sql0="SELECT MAX(Operacion) AS Operacion "&_
"FROM PDA_CCD_DOC"
set rs0=oConn.execute(Sql0)
OPER = rs0.fields("Operacion")
OPER = cdbl(OPER) + 1
Sql="SELECT PDA_CCD_DOC.* "&_
"FROM PDA_CCD_DOC"
set rs=oConn.execute(Sql)

dim Numdoc()
dim PNumdoc()
dim PMondoc()
dim Mondoc()
x=0
y=0
for each elemento in request.form
Redim Preserve Numdoc(y)
Redim Preserve Mondoc(y)
Redim Preserve PNumdoc(x)
Redim Preserve PMondoc(x)

if left(elemento,4) = "sfac" then
PNumdoc(x) = right(elemento,6)
PMondoc(x) = cdbl(request.form(elemento))*-1
x = x + 1
end if
if left(elemento,4) = "fact" then
Numdoc(y) = right(elemento,6)
Mondoc(y) = request.form(elemento)
y = y + 1
end if
next
linea = 1

%>
	<span class="Estilo5">Detalle Pago</span></p>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
	<tr>
		<td bgcolor="#000099">
		<font size="-1" face="arial"><span class="Estilo2"><strong>Numero</strong></span></font></td>
		<td bgcolor="#000099">
		<span class="Estilo2"><strong><font size="-1" face="arial">Debe</font></strong></span></td>
		<td bgcolor="#000099">
		<span class="Estilo4 Estilo2"><font size="-1" face="arial">Haber</font></span></td>
	</tr>
	<font face="arial" size="-1">
	<%
rs.close
for i = lbound(PNumdoc) to ubound(PNumdoc)
if len(PNumdoc(0))= 0 then exit for
saldo = cdbl(PMondoc(i))
t = lbound(Numdoc)
for x = t to ubound(Numdoc)
'---
'----if not cdbl(PMondoc(i)) = 0 and  saldo = 0 then
'------------------------------------------------
rs.open Sql, oConn, 1, 3
rs.addnew

rs.fields("Empresa"			) = "DESA"
rs.fields("Operacion"		) = OPER
rs.fields("ID_Vendedor"		) = vendedor
rs.fields("vendedor"		) = rs1.fields("nombre")
rs.fields("Linea"			) = linea
rs.fields("fecha"			) = date
rs.fields("TipoDocto"		) = "Factura"
rs.fields("Referencia"		) = Numdoc(x)
rs.fields("TipoDoctoPago"	) = "Factura"
rs.fields("NroDoctoPago"	) = PNumdoc(i)
rs.fields("Entidad"			) = ""
rs.fields("fechaVcto"		) = ""
rs.fields("Cliente"			) = request.form("Cliente")
rs.fields("razonsocial"		) = rs2.fields("razonsocial")
rs.fields("Traspasado"		) = 0
rs.fields("CuentaPago"		) = Cuenta
%>
	<tr>
		<td>&nbsp;Factura N&ordm;<%= Numdoc(x) %>&nbsp;&nbsp;&nbsp;&nbsp;N&ordm; Documento: 
			<%=  PNumdoc(i) %></td>
		<td><%
if cdbl(saldo) >= cdbl(Mondoc(x))then
response.write(Mondoc(x))
rs.fields("monto")= cdbl(Mondoc(x))
end if
if cdbl(saldo) <= cdbl(Mondoc(x)) then
response.write(formatnumber(cdbl(saldo),0))
rs.fields("monto")= cdbl(saldo)
end if
%>
			&nbsp;</td>
		<td><% if x = 0 then response.write(PMondoc(x)) %>&nbsp;</td>
	</tr>
	<%
rs.update
rs.close
'---
'end if
saldo = cdbl(saldo) - cdbl(Mondoc(x))
linea = linea + 1
if cdbl(saldo) < cdbl(Mondoc(x)) then
mondoc(x)= cdbl(saldo)*-1
t = x
exit for
end if
next
next
saldo = cdbl(request.form("monto"))
for x = t to ubound(Numdoc)

rs.open Sql, oConn, 1, 3
rs.addnew

rs.fields("Empresa"			) = "DESA"
rs.fields("Operacion"		) = OPER
rs.fields("ID_Vendedor"		) = vendedor
rs.fields("vendedor"		) = rs1.fields("nombre")
rs.fields("Linea"			) = linea
rs.fields("fecha"			) = date
rs.fields("TipoDocto"		) = "Factura"
rs.fields("Referencia"		) = Numdoc(x)
rs.fields("TipoDoctoPago"	) = Tipopago
rs.fields("NroDoctoPago"	) = request.form("Cheque")
rs.fields("Entidad"			) = request.form("nbanco")
rs.fields("fechaVcto"		) = request.form("vence")
rs.fields("Cliente"			) = request.form("Cliente")
rs.fields("razonsocial"		) = rs2.fields("razonsocial")
rs.fields("Traspasado"		) = 0
rs.fields("CuentaPago"		) = Cuenta
%>
	<tr>
		<td>&nbsp;Factura N&ordm;<%= Numdoc(x) %>&nbsp;&nbsp;&nbsp;&nbsp;N&ordm; Documento:&nbsp;
			<%= request.form("cheque") %></td>
		<td><%
if cdbl(saldo) >= cdbl(Mondoc(x))then
response.write(Mondoc(x))
rs.fields("monto")= cdbl(Mondoc(x))
end if
if cdbl(saldo) <= cdbl(Mondoc(x)) then
response.write(saldo)
rs.fields("monto")= cdbl(saldo)
end if
%>
			&nbsp;</td>
		<td><% if x = t then response.write(request.form("monto")) %>&nbsp;</td>
	</tr>
	<%
rs.update
rs.close
linea = linea + 1
saldo = cdbl(saldo) - cdbl(Mondoc(x))
'response.write(saldo)
if cdbl(saldo) < cdbl(Mondoc(x)) then
exit for
end if
next
if cdbl(saldo) > 0 then 

rs.open Sql, oConn, 1, 3
rs.addnew

rs.fields("Empresa"			) = "DESA"
rs.fields("Operacion"		) = OPER
rs.fields("ID_Vendedor"		) = vendedor
rs.fields("vendedor"		) = rs1.fields("nombre")
rs.fields("Linea"			) = linea
rs.fields("fecha"			) = date
rs.fields("TipoDocto"		) = "Factura"
rs.fields("Referencia"		) = Numdoc(x)
rs.fields("TipoDoctoPago"	) = Tipopago
rs.fields("NroDoctoPago"	) = request.form("cheque")
rs.fields("Entidad"			) = request.form("nbanco")
rs.fields("fechaVcto"		) = request.form("vence")
rs.fields("Cliente"			) = request.form("Cliente")
rs.fields("razonsocial"		) = rs2.fields("razonsocial")
rs.fields("Traspasado"		) = 0
rs.fields("CuentaPago"		) = Cuenta
%>
	<tr>
		<td>&nbsp;Factura N&ordm; <%= Numdoc(x) %>&nbsp;&nbsp;&nbsp;&nbsp;N&ordm; Documento: 
		<%= request.form("cheque") %></td>
		<td><%response.write(formatnumber(cdbl(saldo),0))
		      rs.fields("monto") = cdbl(saldo) 
		 %>
			&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<%
	rs.update
	rs.close
end if
%>
	</font>
</table>
	<p align="center">Datos Almacenados Correctamente</p>
	<form method="post" action="">
<p align="center">
		<input type="submit" value="Ingresar Nuevo Pago"></p>
		<input type="hidden" name="B1" value="Agregar">
		<input type="hidden" name="nuser" value="<%= right("000" & vendedor,3)%>">
		<input type="hidden" name="vendedor" value="<%= right("000" & vendedor,3)%>">
	</form>
	<form method="post" action="../Default.asp?nuser=<%= right("000" & vendedor,3)%>">
<p align="center">
		<input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Salir&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"></p>
<p>
	<%
end sub 'paso04()
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub resumen()

'call rescatavalor()

':::: validando pago ::::

' Si el monto no es valido
if len(request.form("monto"))=0  or not IsNumeric(request.form("monto")) then
		response.write("<BR><BR>Monto No valido<BR><BR>")
		Exit sub
else ' si el cheque esta mal ingresado
	if request.form("Tipopago")="Cheque" and not IsNumeric(request.form("cheque")) then
		response.write("<BR><BR>Ingrese Numero de Cheque<BR><BR>")
		Exit sub
	end if
end if

SQL="SELECT FX1_CtaCte_Deudas.Nreferencia, Min(FX1_CtaCte_Deudas.FechaVcto) as 'FechaVcto', Sum(FX1_CtaCte_Deudas.Debe_Ingreso-FX1_CtaCte_Deudas.Haber_ingreso) 'Saldo', FX1_CtaCte_Deudas.Ejecutivo, max(FX1_CtaCte_Deudas.AUX_VALOR6) as 'AUX_VALOR6', FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, FX1_CtaCte_Deudas.Direccion, FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, FX1_CtaCte_Deudas.Referencia " & _
"FROM master.dbo.FX_Saldos_CtaCte FX_Saldos_CtaCte, master.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas, BDFlexline.dbo.PALM_VENDEDOR PALM_VENDEDOR " & _
"WHERE FX_Saldos_CtaCte.Referencia = FX1_CtaCte_Deudas.Referencia AND PALM_VENDEDOR.DESCRIPCION = FX1_CtaCte_Deudas.Ejecutivo " & _
"GROUP BY FX1_CtaCte_Deudas.Nreferencia, FX1_CtaCte_Deudas.Ejecutivo, FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, FX1_CtaCte_Deudas.Direccion, FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, FX1_CtaCte_Deudas.Referencia, PALM_VENDEDOR.NV " & _
"HAVING (PALM_VENDEDOR.NV='" & right(Vendedor,2) & "') " & _
"AND (FX1_CtaCte_Deudas.AUX_VALOR3='" & trim(cstr(request.form("cliente" ))) & "') " & _
"AND (SUM(FX1_CtaCte_Deudas.Debe_Ingreso - FX1_CtaCte_Deudas.Haber_ingreso) <> 0) " & _
"ORDER BY FX1_CtaCte_Deudas.FechaVcto, FX1_CtaCte_Deudas.Nreferencia"

Set rs=oConn.execute(Sql)
'response.write(SQL)

%>
	<font color="#000000">
	<strong><font color="#000080">Resumen CCD</font></strong></font></p>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="139">
	<tr>
		<td align="center" bgcolor="#000080"><b><font face="Arial" SIZE=2 color="#FFFFFF">TipoDocto</font></b></td>
		<td align="center" bgcolor="#000080"><b><font face="Arial" SIZE=2 color="#FFFFFF">Numero</font></b></td>
		<td align="center" bgcolor="#000080"><b><font face="Arial" SIZE=2 color="#FFFFFF">Vencimiento</font></b></td>
		<td align="center" bgcolor="#000080"><b><font face="Arial" SIZE=2 color="#FFFFFF">Debe</font></b></td>
		<td align="center" bgcolor="#000080"><b><font face="Arial" SIZE=2 color="#FFFFFF">Haber</font></b></td>
	</tr>
	<tr>
		<td align="center">
			<font face="Arial"><%= request.Form("Tipopago") %></font>
		</td>
		<input type="hidden" name="Tipopago" value="<% =request.Form("Tipopago") %>">
		<td align="center"><font face="Arial">
			<% =request.Form("cheque") %>
			</font></td>
		<input type="hidden" name="cheque" value="<% =request.Form("cheque") %>">
		<td align="center"><font face="Arial">
			<%
if request.Form("Tipopago") = "Efectivo" then
		response.write("&nbsp;")
else
		response.write(request.Form("dia") & "/" & request.Form("mes") & "/" & request.Form("ano"))
end if		
%>
			</font></td>
		<input type="hidden" name="vence" value="<% =right("00" & request.Form("dia"),2) & "/" & right("00" & request.Form("mes"),2) & "/" & request.Form("ano") %>">
		<td align="center"><font face="Arial">
			.</font></td>
		<td align="center"><font face="Arial">
			<% =formatnumber(cdbl(request.Form("monto")),0) %>
			</font></td>
		<input type="hidden" name="monto" value="<% =request.Form("monto") %>">
		<input type="hidden" name="nbanco" value="<%= request.form("Banco")%>">
		<input type="hidden" name="Cliente" value="<%= request.form("Cliente")%>">
		<input type="hidden" name="vendedor" value="<%= request.form("vendedor")%>">
		<input type="hidden" name="Cuenta" value="<%= request.form("Cuenta")%>">
	</tr>
	<%
 saldo2 = 0
Saldopago=cdbl(request.form("Monto" ))
Sichecked="checked"
do until rs.eof
'********************************
'*******
Numero=cdbl(trim(rs.fields("nreferencia")))

For Each elemento In request.form
	if elemento = ("fact" & Numero) then 
		if cdbl(rs.fields("saldo")) < 0 then 
		saldopago = saldopago - cdbl(rs.fields("saldo")) 
		saldo2 = cdbl(rs.fields("saldo"))* -1
'		response.Write(rs.fields("saldo") & "&nbsp;" & saldopago)
response.write("<input type='hidden' name='sfac"& Numero &"' value='"& rs.fields("saldo") &"'>")
		end if
		if cdbl(rs.fields("saldo")) > 0 then
response.write("<input type='hidden' name='fact"& numero &"' value='"& rs.fields("saldo") &"'>")
			if saldopago >= cdbl(rs.fields("saldo")) then
				pago = cdbl(rs.fields("saldo"))
				Saldopago = Saldopago - cdbl(rs.fields("saldo"))
				saldfin=saldopago
				tmf=tmf + cdbl(rs.fields("saldo"))
			else
				pago = saldopago
				saldopago = saldopago - cdbl(rs.fields("saldo"))
				saldfin=saldopago
				tmf=(cdbl(tmf) + cdbl(pago))
			end if
		end if
		if pago < 0 then exit do
%>
	<tr>
		<td width="25%" align="center"><font face="Arial">Factura</font></td>
		<td width="25%" align="center"><font face="Arial">
			<% =right(elemento,6) %>
			</font></td>
		<td width="25%" align="center"><font face="Arial">
			<% =rs.fields("FechaVcto")%>
			</font></td>
		<td width="25%" align="center"><font face="Arial">
			&nbsp;
			<%=formatnumber(cdbl(pago),0)%>
			</font></td>
		<td width="25%" align="center">&nbsp;
			<%if cdbl(rs.fields("saldo")) < 0 then response.Write(cdbl(rs.fields("saldo")) * -1)%></td>
	</tr>
	<%
	end if

Next
'*********************************
rs.movenext
loop
if saldfin >0 then
tmf=(cdbl(tmf)+ cdbl(saldfin))
%>
	<tr>
		<td width="25%" align="center"><font face="Arial">Factura</font></td>
		<td width="25%" align="center"><font face="Arial">888888</font></td>
		<td width="25%" align="center"><font face="Arial">
			<%= date %>
			</font></td>
		<td width="25%" align="center"><font face="Arial">
			<%= saldfin %>
			</font></td>
		<td width="25%" align="center"><font face="Arial">&nbsp;</font></td>
	</tr>
	<%
end if
%>
	<tr>
		<td colspan="3" width="25%" align="right"><font face="Arial"><b>Total:
			&nbsp;
			</b></font></td>
		<td width="25%" align="center"><font face="Arial"><b>
			<%= formatnumber(tmf,0) %>
			</b></font></td>
		<td width="25%" align="center"><font face="Arial"><b>
			<%
		response.write(formatnumber(cdbl(request.form("Monto"))+ saldo2,0))%>
			</b></font></td>
	</tr>
</table>
<p align="center">
	<center>
		<input type="button" value="Editar" onClick="history.back()">
		&nbsp;
		<input type="Submit" name="B1" value="Traspasar">
	</center>
	<%
'---------------------
End sub 'resumen()
'::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub paso03()

call rescatavalor()

':::: validando pago ::::
'IsNumeric(
' Si el monto no es valido
if len(request.form("monto"))=0  or not IsNumeric(request.form("monto")) then
		response.write("<BR><BR><p align='center'>Monto No valido<BR><BR>")
				call vuelve()
		Exit sub
' si el cheque esta mal ingresado
end if
if request.form("Tipopago") = "Efectivo" then
	if len(request.form("Cheque")) <> 0 then
		response.write("<br><br><p align='center'>Indique Documento (Cheque, Factura Canje)</p>")
		response.write("<br><br><p align='center'>El numero solo es requerido para Cheque o Factura Canje</p>")
				call vuelve()
		exit sub
	end if
end if
if request.form("Tipopago")="Cheque" then
		if len(request.form("Cheque"))=0 and not IsNumeric(request.form("cheque")) then
		response.write("<BR><BR>Ingrese Numero de Cheque<BR><BR>")
				call vuelve()
		Exit sub
		end if
		if len(request.Form("Banco"))=0 then
		response.Write("<BR><BR><strong>Seleccione Banco</strong><BR><BR>")
				call vuelve()
 		exit sub 
		end if
end if
if request.form("Tipopago")="Factura Canje" then
		if not isnumeric(request.form("Cheque")) then
		response.write("<br><br><p align='center'>Ingrese Numero de Factura Canje</p>")
				call vuelve()
		exit sub
		end if
end if
SQL="SELECT FX1_CtaCte_Deudas.Nreferencia, min(FX1_CtaCte_Deudas.FechaVcto) as 'FechaVcto', Sum(FX1_CtaCte_Deudas.Debe_Ingreso-FX1_CtaCte_Deudas.Haber_ingreso) 'Saldo', FX1_CtaCte_Deudas.Ejecutivo, MAX(FX1_CtaCte_Deudas.AUX_VALOR6) as 'AUX_VALOR6', FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, FX1_CtaCte_Deudas.Direccion, FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, FX1_CtaCte_Deudas.Referencia " & _
"FROM master.dbo.FX_Saldos_CtaCte FX_Saldos_CtaCte, master.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas, BDFlexline.dbo.PALM_VENDEDOR PALM_VENDEDOR " & _
"WHERE FX_Saldos_CtaCte.Referencia = FX1_CtaCte_Deudas.Referencia AND PALM_VENDEDOR.DESCRIPCION = FX1_CtaCte_Deudas.Ejecutivo " & _
"GROUP BY FX1_CtaCte_Deudas.Nreferencia, FX1_CtaCte_Deudas.Ejecutivo, FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, FX1_CtaCte_Deudas.Direccion, FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, FX1_CtaCte_Deudas.Referencia, PALM_VENDEDOR.NV " & _
"HAVING (PALM_VENDEDOR.NV='" & right(Vendedor,2) & "') " & _
"AND (FX1_CtaCte_Deudas.AUX_VALOR3='" & trim(cstr(request.form("cliente" ))) & "') " & _
"AND (SUM(FX1_CtaCte_Deudas.Debe_Ingreso - FX1_CtaCte_Deudas.Haber_ingreso) <> 0) " & _
"ORDER BY FX1_CtaCte_Deudas.FechaVcto, FX1_CtaCte_Deudas.Nreferencia"
'response.write(SQL)

'Sql1="SELECT * FROM Todo.Flexline.PDA_CCD_DOC PDA_CCD_DOC " & _
'"WHERE (PDA_CCD_DOC.estado='0') AND (PDA_CCD_DOC.tipo='Factura')"

Set rs=oConn.execute(Sql)

if request.form("Tipopago") = "Efectivo" then 
	response.write("<input type='hidden' name='Cuenta' value='1104005'>")
end if
if request.form("Tipopago") = "Cheque" then
stp = "SELECT zona FROM Flexline.FX_vende_man WHERE (NU_2 = N'"& vendedor &"')"
set rtp =oConn1.execute(stp)
	if rtp.fields("zona") = "Z Santiago" then
	response.write("<input type='hidden' name='Cuenta' value='1104003'>")
	else
	response.write("<input type='hidden' name='Cuenta' value='1104004'>")
	end if
end if

%>
<p align="center">
	<b><font face="Arial" size="2" color="#000080">Listado
	&nbsp;
	Facturas</font></b></p>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="8%" bgcolor="#000080" align="center"><b>
			<font face="Arial" size="2" color="#FFFFFF">CK</font></b></td>
		<td width="24%" bgcolor="#000080" align="center"><b>
			<font face="Arial" size="2" color="#FFFFFF">Numero</font></b></td>
		<td width="34%" bgcolor="#000080" align="center"><b>
			<font face="Arial" size="2" color="#FFFFFF">Vencimiento</font></b></td>
		<td width="121%" bgcolor="#000080" align="center"><b>
			<font face="Arial" size="2" color="#FFFFFF">Monto</font></b></td>
	</tr>
	<%
Saldopago=cdbl(request.form("Monto" ))
Sichecked="checked"
do until rs.eof
			if cdate(rs.fields("FechaVcto")) <= date() then
				Colorven="#CC3300"
			else
				Colorven="#000080"
			end if
		if Saldopago < 1 then Sichecked=""
		
	response.write("<tr>")
	response.write("<td width='8%' align='center' height='20'>")
	response.write("<input type='checkbox' name='fact" & _
	cdbl(rs.fields("nreferencia")) & "' " & Sichecked & "></td>")
	response.write("<td width='24%' align='center' height='20'><font face='Arial' size='2'>")
	response.write( cdbl(rs.fields("nreferencia")) & "</font></td>")
	response.write("<td width='34%' align='center' height='20'><b>")
	response.write("<font face='Arial' size='2' color='" & Colorven & "'>")
	response.write( rs.fields("FechaVcto") & "</font></b></td>")
	response.write("<td width='121%' align='center' height='20'><font face='Arial' size='2'>")
	response.write( formatnumber(cdbl(rs.fields("saldo")),0) & "</font></td>")
	response.write("</tr>")
	
Saldopago = Saldopago - cdbl(rs.fields("saldo"))
'end if
rs.movenext
loop
%>
</table>
<p align="center">
	<input type="submit" value="Acepta Pago" name="B1">
	<input type="Submit" value="Cancelar" name="B1">
</p>
<%
end sub 'paso03()
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub Ingresopago()

call rescatavalor()

%>
<p align="center"><b><font face="Arial" size="2" color="#000080">Ingreso
	Pago</font></b></p>
<hr>
<table border="1" align="center" cellpadding="0" 
cellspacing="0" bordercolor="#111111" 
style="border-collapse: collapse; border-width: 0">
	<tr>
		<td align="center" style="border-style: none; border-width: medium"><p align="right">
		<input type="radio" value="Efectivo" checked name="Tipopago"></p></td>
		<td align="center" style="border-style: none; border-width: medium"><p align="left">
		<font face="Arial" size="2"><b>&nbsp;Efectivo</b></font></p></td>
	</tr>
	<tr>
		<td align="center" style="border-style: none; border-width: medium"><p align="right">
		<font face="Arial"><input type="radio" name="Tipopago" value="Cheque"></font></p></td>
		<td align="center" style="border-style: none; border-width: medium"><p align="left">
		<font face="Arial" size="2"><b>&nbsp;Cheque</b></font></p></td>
	</tr>
	<tr>
		<td align="right" style="border-style: none; border-width: medium">
		<input name="Tipopago" type="radio" value="Factura Canje"></td>
		<td align="left" style="border-style: none; border-width: medium">
		<b>Factura Canje</b></td>
	</tr>
	<tr>
		<td align="center" style="border-style: none; border-width: medium">&nbsp;</td>
		<td align="center" style="border-style: none; border-width: medium">&nbsp;</td>
	</tr>
	<tr>
		<td align="center" style="border-style: none; border-width: medium">
		<b><font face="Arial" size="2">Monto</font></b></td>
		<td align="center" style="border-style: none; border-width: medium">
		<input type="text" name="Monto" size="20"></td>
	</tr>
	<tr>
		<td align="center" style="border-style: none; border-width: medium">
		<b><font face="Arial" size="2">* Numero</font></b></td>
		<td align="center" style="border-style: none; border-width: medium">
		<input name="Cheque" style="speak-numeral:digits" type="text" size="20"></td>
	</tr>
	<tr>
		<td height="27" align="center" style="border-style: none; border-width: medium">
		<b><font face="Arial" size="2">* Banco</font></b></td>
<%
Sql_BCO="SELECT * FROM GEN_TABCOD WHERE (EMPRESA = 'DESA') AND (TIPO = 'gen_bancos')"
set rs_BCO=oConn.execute(Sql_BCO)
%>
		<td align="center" nowrap style="border-style: none; border-width: medium">
		<select name="Banco" id="Banco"  style="height:20" >
				<option value="" selected>&nbsp;&nbsp;</option>
<%
do until rs_BCO.eof
%>
	<option value="<%= rs_BCO.fields("codigo") %>"><%= left(mid(rs_BCO.fields("descripcion"),6),15) %></option>
<%
rs_BCO.movenext
loop
%>
	</select></td>
	</tr>
	<tr>
		<td align="center" style="border-style: none; border-width: medium">
		<b><font face="Arial" size="2">* Vencim</font></b></td>
		<td align="center" style="border-style: none; border-width: medium">
		<select size="1" name="Dia">
<% x=1
do until x=32
	if x=day(date) then
		response.write("<OPTION selected value=" & x & ">" & x )
	else
		response.write("<OPTION value=" & x & ">" & x )
	end if
		x=x+1
loop %>
		</select>
		<select size="1" name="Mes">
<% x=1
do until x=13
	if x=month(date) then
		response.write("<OPTION selected value=" & x & ">" & x )
	else
		response.write("<OPTION value=" & x & ">" & x )
	end if
		x=x+1
loop %>
		</select>
		<select size="1" name="Ano">
			<option selected><%= year(date)%></option>
			<option><%= year(dateadd("yyyy",1,date))%></option>
		</select></td>
	</tr>
</table>
<p align="center"><font size="1" color="#808080" face="Arial">* Numero, Banco y Vencimiento solo Cheque</font></p>
<p align="center">
	<input type="submit" value="Aceptar" name="B1">
	<input type="reset" value="Cancelar">
</p>
<%
End Sub 'Ingresopago()
'::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::
sub paso02()

call rescatavalor()

SQL="SELECT FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, Sum(FX1_CtaCte_Deudas.Debe_Ingreso-FX1_CtaCte_Deudas.Haber_ingreso) 'saldo1' " & _
"FROM master.dbo.FX_Saldos_CtaCte FX_Saldos_CtaCte, master.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas, BDFlexline.dbo.PALM_VENDEDOR PALM_VENDEDOR " & _
"WHERE FX_Saldos_CtaCte.Referencia = FX1_CtaCte_Deudas.Referencia AND PALM_VENDEDOR.DESCRIPCION = FX1_CtaCte_Deudas.Ejecutivo AND ((PALM_VENDEDOR.NV='" & right(Vendedor,2) & "')) " & _
"GROUP BY FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial " & _
"ORDER BY FX1_CtaCte_Deudas.RazonSocial"
Set rs=oConn.execute(Sql)
%>
<table width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="3" align="center"><b><font size="2" face="Arial" color="#000080">Ingreso
		Caja CCD</font></b></td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	<tr align="center" bgcolor="#000080">
		<td width="30" bgcolor="#000080"><b><font face="Arial" size="2" color="#FFFFFF">CK</font></b></td>
		<td><b><font face="Arial" size="2" color="#FFFFFF">Cliente</font></b></td>
		<td><b><font face="Arial" size="2" color="#FFFFFF">Deuda</font></b></td>
	</tr>
<%
Do until rs.eof
%>
	<tr>
		<td style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin"
		align="center"><input type="radio" value="<%= rs.fields("AUX_VALOR3") %>" name="Cliente"></td>
<a href="cc.asp?nuser=<%= right(Vendedor,2) %>&Cliente=<%= rs.fields("AUX_VALOR3") %>" style="text-decoration: none">
		<td	align="center" nowrap style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin;cursor:pointer"><%= left(rs.fields("RazonSocial"),20) %></td>
		<td style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin;cursor:pointer"	align="right"><%= formatnumber(cdbl(rs.fields("saldo1")),0) %></td>
</a>
	</tr>
<%
rs.movenext
loop
%>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
</table>
<p align="center">
	<input type="submit" value="Ingresa Doc Pago" name="B1">
</p>
<%
end sub 'paso02()
'::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::
Sub paso01()
call rescatavalor()

Sql="SELECT TOP 10 Operacion AS ID, TipoDoctoPago AS Tipopago, NroDoctoPago AS Refer, Entidad AS banco, SUM(monto) AS Monto, Traspasado AS Estado "&_
"FROM PDA_CCD_DOC WHERE (ID_Vendedor = '" & right(vendedor,2) &"') "&_
"GROUP BY Operacion, TipoDoctoPago, NroDoctoPago, Entidad, Traspasado "&_
"ORDER BY Operacion DESC, COUNT(Linea)"

Set rs=oConn.execute(Sql)
'response.write(Sql)
if len(vendedor)<>0 then Response.Write("<INPUT TYPE='hidden' name='Vendedor' value='" & vendedor & "'>") 
%>
<table width="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="4" align="center"><b><font size="2" face="Arial" color="#000080">Ingreso
		Caja CCD</font></b></td>
	</tr>
	<tr>
		<td colspan="4" align="center">&nbsp;</td>
	</tr>
	<tr bgcolor="#000080">
		<td align="center" bgcolor="#000080"><b>
		<font color="#FFFFFF" face="Arial" size="2">N
			&ordm;
		</font></b></td>
		<td colspan="2" align="center" bgcolor="#000080"><b>
		<font color="#FFFFFF" face="Arial" size="2">Descripcion</font></b></td>
		<td align="center" bgcolor="#000080"><b><font face="Arial" size="2" color="#FFFFFF">Estado</font></b></td>
	</tr>
	<%
Do until rs.eof
'Miestado
if rs.fields("Estado") = 0 then Miestado = "Pendiente"
if rs.fields("Estado")=1 then Miestado = "Ingresado"
%>
	<tr align="center">
		<td><span class="Estilo1">
			<%= rs.fields("ID")  %>
		</span></td>
		<td><span class="Estilo1"><strong class="Estilo1">Tipo: </strong>
				<%= rs.fields("TipoPago") %>
		</span></td>
		<td><span class="Estilo1"><strong class="Estilo1">Entidad: </strong>
				<%= rs.fields("Banco") %>
		</span></td>
		<td><span class="Estilo1">
			<%= miestado %>
		</span></td>
	</tr>
	<tr align="center">
		<td style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin">
			<span class="Estilo2">
			&nbsp;
		<%'=  %></span></td>
		<td style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin">
		<span class="Estilo1"><strong class="Estilo1">N&ordm;</strong>&nbsp;
		<%= rs.fields("Refer") %></span></td>
		<td style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin">
		<b><font face="Arial" size="2">Monto: <span class="Estilo1">
		<%= formatnumber(cdbl(rs.fields("monto")),0) %>
		</span></font></b></td>
		<td style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin">
		&nbsp;
		<span class="Estilo2"><%'=  %></span></td>
<%
rs.movenext
loop
%>
	</tr>
</table>
<p align="center">
	<input type="submit" value="Agregar" name="B1">
</p>
<%
End Sub'Paso01()
%>
<% 'call buscavalor()%>
</form>
</body>
</html>
<%
sub rescatavalor()

For Each elemento In request.form
	if elemento <> "B1" then 
	Valorelemento=Request.form(elemento)
	if elemento="Cheque" then Valorelemento=right(trim(Valorelemento),4)
	Response.Write("<INPUT TYPE='hidden' name='" & elemento & "' Value='" & Valorelemento & "'>")
	end if
Next
end sub 'rescatavalor()
'::::::::::::::::::::::::::::::::::::::::::::::::::
sub buscavalor()
'Response.Write("<FONT SIZE=1 COLOR='#FFFFFF'>")
For Each elemento In request.form
	Response.Write("<BR>" & elemento & " : " & Request.form(elemento))
Next
'Response.Write("</FONT>")
end sub 'buscavalor()
sub vuelve()
%>
<p align="center">
<input type="button" value="&nbsp;&nbsp;&nbsp;Volver&nbsp;&nbsp;&nbsp;" onClick="history.back()">
</p>
<%
end sub
%>
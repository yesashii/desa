<%
set oConn = server.CreateObject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<!--
*************************************
*	Página desarrollada por:		*
*	Cristian Palma Roubillart		*
*	Reservados todos los derechos	*
*	Contacto:						*
*	developer_systems@hotmail.com	*
*************************************
-->
<meta http-equiv="Content-Type" content="text/html">
<title>Pago Documentos - Distribuidora Err&aacute;zuriz</title>
</head>
<body topmargin="0" style="font-family:Arial; font-size:x-small">
<%
step = request.form("step")

if step = "" then call buscador()
if step = "Buscar" then call listaclientes()
if step = "Ver Ultimos Ingresos" then call listapagos()
'•••••••••••••••••••••••••••••••••••••••••
'•••••••••••••••••••••••••••••••••••••••••
sub buscador()
%>
<form method="post" action="">
<table width="100%" cellspacing="0" cellpadding="0">
	<tr>
	  <td align="center" bgcolor="#000080"><font face="Arial" color="#FFFFFF" size="-1"><strong>Pago Cuenta Corriente</strong></font></td>
  </tr>
	<tr>
	  <td align="center"><br></td>
    </tr>
	<tr>
	  <td align="center"><strong>Buscar Cliente:</strong>&nbsp;</td>
  </tr>
	<tr>
	  <td align="center"><br></td>
    </tr>
	<tr>
	  <td align="center"><input name="busca" type="text" value="" size="20" style="font-family:Arial; font-size:x-small"></td>
  </tr>
	<tr>
	  <td align="center"><br></td>
    </tr>
	<tr>
	  <td align="center"><input name="step" type="submit" value="Buscar"></td>
  </tr>
 </table>
</form>	
<form method="post" action="">
<p align="center"><input type="submit" name="step" value="Ver Ultimos Ingresos"></p>
</form>
<%
end sub 'buscador
'••••••••••••••••••••••••••••••••••••••••••
'••••••••••••••••••••••••••••••••••••••••••
sub listaclientes()
vendedor = request.querystring("nuser")
cliente = request.form("busca")
SQL="SELECT FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, Sum(FX1_CtaCte_Deudas.Debe_Ingreso-FX1_CtaCte_Deudas.Haber_ingreso) 'saldo1' " & _
"FROM master.dbo.FX_Saldos_CtaCte FX_Saldos_CtaCte, master.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas, BDFlexline.dbo.PALM_VENDEDOR PALM_VENDEDOR " & _
"WHERE FX_Saldos_CtaCte.Referencia = FX1_CtaCte_Deudas.Referencia AND PALM_VENDEDOR.DESCRIPCION = FX1_CtaCte_Deudas.Ejecutivo AND ((PALM_VENDEDOR.NV='" & right(Vendedor,2) & "')) " & _
"GROUP BY FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial " & _
"HAVING      (FX1_CtaCte_Deudas.AUX_VALOR3 + FX1_CtaCte_Deudas.RazonSocial LIKE '%"& cliente &"%') "&_
"ORDER BY FX1_CtaCte_Deudas.RazonSocial"
'response.write(SQL)
Set rs=oConn.execute(Sql)
%>
<table width="100%" cellspacing="0" cellpadding="0">
<tr>
	  <td colspan="2" align="center" bgcolor="#000080"><font face="Arial" color="#FFFFFF" size="-1"><strong>Pago Cuenta Corriente</strong></font></td>
  </tr>
	<tr>
		<td colspan="2"><b><font color="#000080">Seleccione
		Cliente : </font></b></td>
	</tr>
<%
if rs.eof then 
response.write("</table><p align='center'>No Hay Coincidencias</p>")
response.write("<input type='button' value='volver' onclick='history.back()'")
end if
Do until rs.eof
%>
	<tr bgcolor="#000080">
		<td bgcolor="#DFDFDF"><font face="Arial" size="2"><b>Cliente</b></font></td>
		<td align="center" bgcolor="#DFDFDF"><b><font face="Arial" size="2">Deuda</font></b></td>
    </tr>

	<tr>
		<td
		align="center" style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin"><%= left(rs.fields("RazonSocial"),20) %></td>
        <td
		align="center" style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin">$<%= formatnumber(cdbl(rs.fields("saldo1")),0) %></td>
  </tr>
	<tr>
	  <td
		align="center" style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin">
		<form method="post" action="ccd.asp?nuser=<%= vendedor %>">
		<input type="hidden" value="<%= rs.fields("AUX_VALOR3") %>" name="Cliente">
		<input type="submit" value="Ingresa Doc Pago" name="B1" style="font-size:x-small">
		</form>
		</td>
	  <td align="center" style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin"><form method="get" action="cc.asp">
        <input type="hidden" name="nuser" value="<%= right(Vendedor,2) %>">
        <input type="hidden" name="Cliente" value="<%= rs.fields("AUX_VALOR3")%>">
        <input name="submit" type="submit" style="font-size:x-small" value="Ver Facturas">
      </form></td>
  </tr>
	<tr>
		<td colspan="2"><br>&nbsp;</td>
	</tr>
<%
rs.movenext
loop
%>
</table>
<%
end sub 'lista clientes
'???????????????????????????????????????
'???????????????????????????????????????
sub listapagos()
vendedor = request.querystring("nuser")

Sql="SELECT TOP 10 Operacion AS ID, TipoDoctoPago AS Tipopago, NroDoctoPago AS Refer, Entidad AS banco, SUM(monto) AS Monto, Traspasado AS Estado "&_
"FROM PDA_CCD_DOC WHERE (ID_Vendedor = '" & right(vendedor,2) &"') "&_
"GROUP BY Operacion, TipoDoctoPago, NroDoctoPago, Entidad, Traspasado "&_
"ORDER BY Operacion DESC, COUNT(Linea)"

Set rs=oConn.execute(Sql)
'response.write(Sql)
%>
<table width="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="4" align="center"><b><font size="2" face="Arial" color="#000080">Ingresos
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
tipo	= rs.fields("TipoPago")
banco	= rs.fields("Banco")
numdoc	= rs.fields("Refer")
monto	= rs.fields("monto")
if  tipo = "Efectivo" then
banco	= "General"
numdoc	= "EFE"
end if
if tipo = "Factura" then banco = "Canje"
%>
	<tr align="center">
		<td><span class="Estilo1">
			<%= rs.fields("ID")  %>
		</span></td>
		<td><span class="Estilo1"><strong class="Estilo1">Tipo: </strong>
				<%= tipo %>
		</span></td>
		<td><span class="Estilo1"><strong class="Estilo1">Entidad: </strong>
				<%= banco %>
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
		<%= numdoc %></span></td>
		<td style="border-bottom-color:#CCCCCC; border-bottom-style:solid; border-bottom-width:thin">
		<b><font face="Arial" size="2">Monto: <span class="Estilo1">
		<%= formatnumber(cdbl(monto),0) %>
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
<p align="center"><input type="button" value="Volver" onClick="history.back()"></p>
<%
end sub 'lista pagos
%>
<p align="center" style="font-family:Arial;font-size:xx-small">Distribuci&oacute;n y Excelencia S.A.</p>
</body></html>
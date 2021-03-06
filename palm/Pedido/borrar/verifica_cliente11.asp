<%
private oConn, oConn2, rs, rs2, sql, sql2, SQL1, rs1, rs0, SQL0 
dim vend

vend	= trim(request.QueryString("vend"))
cliente = trim(request.QueryString("cliente"))
nuser	=right("0" & request.cookies("pdausr"),3)

set oConn = server.CreateObject("ADODB.connection")

oConn.open "provider=SQLOLEDB; Data Source=serverdesa; initial catalog=BDFlexline; user id=flexline; password=corona;"

SQL="SELECT CtaCte.CodLegal, CtaCte.RazonSocial, CtaCteDirecciones.Direccion, CtaCte.CtaCte, CtaCte.Giro, CtaCte.CondPago, CtaCte.Comuna, CtaCte.Telefono, CtaCte.LimiteCredito, CtaCte.PorcDr1 " & _
"FROM BDFlexline.flexline.CtaCte CtaCte, BDFlexline.flexline.CtaCteDirecciones CtaCteDirecciones " & _
"WHERE CtaCteDirecciones.CtaCte = CtaCte.CtaCte AND CtaCteDirecciones.Empresa = CtaCte.Empresa AND CtaCteDirecciones.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='DESA') AND (CtaCte.TipoCtaCte='CLIENTE') AND (CtaCte.Ejecutivo='" & vend & "') AND (CtaCteDirecciones.CtaCte='" & cliente & "') AND (CtaCteDirecciones.Principal<>'s'))"

set rs=oConn.execute(sql)
if rs.eof then
	response.write(sql)
	on error resume next
	sql=""
	set rs=oConn.execute(sql)
end if

sql1="SELECT FX1_CtaCte_Deudas.AUX_VALOR3, Sum(Debe_Ingreso-Haber_ingreso) AS 'Saldo' FROM master.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas WHERE (FX1_CtaCte_Deudas.FechaVcto>=getdate()) GROUP BY FX1_CtaCte_Deudas.AUX_VALOR3 HAVING (FX1_CtaCte_Deudas.AUX_VALOR3='" & rs.fields("codlegal") & "')"

sql2="SELECT FX1_CtaCte_Deudas.AUX_VALOR3, Sum(Debe_Ingreso-Haber_ingreso) AS 'Saldo' FROM master.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas WHERE (FX1_CtaCte_Deudas.FechaVcto<=getdate()) GROUP BY FX1_CtaCte_Deudas.AUX_VALOR3 HAVING (FX1_CtaCte_Deudas.AUX_VALOR3='" & rs.fields("codlegal") & "')"

SQL0="SELECT AUX_VALOR3, SUM(SALDO) AS saldo FROM dbo.[el_SALDOS CHEQUES SANTIAGO] GROUP BY AUX_VALOR3 HAVING (AUX_VALOR3 = '" & rs.fields("codlegal") & "')"

set rs0=oConn.execute(SQL0)
if rs0.eof then
saldo=0
else
saldo = rs0.fields("saldo")
end if

set rs1=oConn.execute(sql1)

set rs2=oConn.execute(sql2)

ctacte	= rs.fields("ctacte")

'------------------------------------
'saldo linea de credito
lmcr=formatnumber(cdbl(rs.fields("LimiteCredito")),0)

if rs2.eof then
vartmp1=0
else
vartmp1=formatnumber(cdbl(rs2.fields("saldo")),0)
end if
if rs1.eof then
vartmp2=0
else
vartmp2=formatnumber(cdbl(rs1.fields("saldo")),0)
end if

vartmp3  = (CLNG(vartmp1) + CLNG(vartmp2))
saldocta = (CLNG(lmcr) - CLNG(vartmp3)+ clng(saldo))

%>
<html>
<head>
<title>Verifica estado cliente</title>
<style type="text/css">
<!--
.Estilo1 {
	font-weight: bold;
	color: #FFFFFF;
}
.Estilo3 {font-size: xx-small}
.Estilo4 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style></head>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0">
	<tr>
		<td colspan="4" bgcolor="#000080"><div align="center" class="Estilo1 Estilo1">Verificacion de estado cliente</div></td>
	</tr>
	<tr>
		<td bordercolor="#999999" class="Estilo3">Nombre Cliente </td>
		<td colspan="2" bordercolor="#999999">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="3" bordercolor="#999999"><strong><%= rs.fields("razonsocial")%></strong>
			<hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Rut		</td>
		<td colspan="2"><strong><%= rs.fields("ctacte") %>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Giro</td>
		<td colspan="2"><strong>
			<% = rs.fields("Giro") %>
		</strong></span></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Monto Credito </td>
		<td colspan="2"><strong><% = formatnumber(cdbl(rs.fields("LimiteCredito")),0) %>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Cond. Pago </td>
		<td colspan="2"><strong><%= rs.fields("condpago")%>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Mnto Vencido </td>
		<td colspan="2"><span class="Estilo4">
		<%
		if rs2.eof then
		response.Write("0") 
		else 
		response.Write( formatnumber(cdbl(rs2.fields("saldo")),0)) 
		end if
%>
		&nbsp;</span></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Mnto. A Vencer </td>
		<td colspan="2"><strong><%
		if rs1.eof then
		response.Write("0") 
		else 
		response.Write( formatnumber(cdbl(rs1.fields("saldo")),0)) 
		end if
%>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Cheque en cartera </td>
		<td colspan="2"><strong><% response.Write( formatnumber(cdbl(saldo),0))  %>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo3">Cred. Disponible </td>
		<td colspan="2"><strong>
<%
if saldocta<0 then
response.Write("<font color='#FF0000'>" & formatnumber(cdbl(saldocta),0) & "</font>")
else
response.Write( formatnumber(cdbl(saldocta),0))
end if
%>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
		<td colspan="3"><span class="Estilo3"><a href="../otros/cc.asp?cliente=<%= rs.fields("codlegal")%>">Ver detalle</a></span></td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	<tr>
		<td><form method="post" action="../otros/ccd.asp?nuser=<%= nuser %>&cliente=<%= rs.fields("codlegal")%>"><input style="font-weight:bold" type="submit" value="Ingreso Pago">
		</form></td>
		
		<% response.Write("<td colspan='2'><form method='post' action='pedido02.asp?vend=" & vend & "&cliente=" & ctacte & "'><input style='font-weight:bold' type='submit' value='Hacer Pedido'></Form>") %>
		</td>
	</tr>
</table>
</body>
</html>

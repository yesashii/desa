<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Reporte CCD Por Vendedor</title>
</head>
<body topmargin="0" leftmargin="0">
<%
vendedor=request.form("vendedor")

set oConn=server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

SQL="SELECT flexline.PDA_CCD_DOC.* " & _
"FROM flexline.PDA_CCD_DOC " & _
"WHERE (Vendedor = '" & trim(vendedor) & "') AND (prep = 0) " & _
"ORDER BY Operacion, Linea"
'response.write(sql)
'on error resume next
set rs=oConn.execute(SQL)

if rs.eof then
	on error resume next

End if
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td><i><font color="#000080"><b>Distribuci&oacute;n y Excelencia S. A.</b></font></i></td>
    <td>
      <p align="center"><b><font face="Arial Narrow">I<u>NGRESO A CAJA</u></font></b></td>
    <td>
      <p align="right"><font face="Arial Narrow">Vendedor : <%=rs.fields("vendedor")%></font></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
    <td>
      <p align="right"><font face="Arial Narrow">fecha:<%=date() & " " & time() %></font></td>
  </tr>
</table>

<table border="1" width="100%" style="border-collapse: collapse" bordercolor="#000066" bordercolorlight="#000066" bordercolordark="#000066" height="29" cellspacing="0" cellpadding="0" >
  <tr>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">Factura</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">Razon_Social</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">OP</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">F.Ingreso</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">Rut</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">Valor</font></b></td>
    <td colspan="4" align="center" bgcolor="#BABED8" height="1" ><b><font size="1" face="Arial" color="#000080">Cheque A Fecha</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font size="1" face="Arial" color="#000080">Cheque<BR>al Dia</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">Efectivo</font></b></td>
    <td rowspan="2" align="center" bgcolor="#BABED8" height="1" ><b><font face="Arial Narrow" color="#000080">Total</font></b></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#BABED8" height="8" ><b><font size="1" face="Arial" color="#000080">Numero</font></b></td>
    <td align="center" bgcolor="#BABED8" height="8" ><b><font size="1" face="Arial" color="#000080">Banco</font></b></td>
    <td align="center" bgcolor="#BABED8" height="8" ><b><font size="1" face="Arial" color="#000080">Venc</font></b></td>
    <td align="center" bgcolor="#BABED8" height="8" ><b><font size="1" face="Arial" color="#000080">Monto</font></b></td>
  </tr>
<%
pago=0
MTotal  = 0
UltOperacion=""
UltNroDoctoPago=""
operacion=""
x=0
totalfacturas=0
totalfact=0
tot1=0
tot2=0
tot3=0
do until rs.eof
Montoafecha = "&nbsp;"
MontoalDia  = "&nbsp;"
MontoEfectivo="&nbsp;"
MontoTotal  = "&nbsp;"
NroDoctoPago= "&nbsp;"
FechaVcto="&nbsp;"
operacion=trim(rs.fields("operacion"))
if UltOperacion <> operacion then
	if x>0  then
	%>
	  <tr>
		<td colspan="5" align="right"><font size="1">
		<%="Total Operacion " & UltOperacion & " : $ " & formatnumber(totalfacturas,0)%>
		&nbsp;</font></td>
		<td align="center"><font size="1">&nbsp;</font></td>
		<td colspan="7"><font size="1">&nbsp;</font></td>
	  </tr>
	<%
	totalfacturas=0
	End If
	UltNroDoctoPago="X"'.....
End if

if UltNroDoctoPago<>trim(rs.fields("NroDoctoPago")) then
	NroDoctoPago=rs.fields("NroDoctoPago")
	if isnumeric(rs.fields("Entidad")) then MontoTotal  =formatnumber(rs.fields("Entidad"),0)
	FechaVcto=cdate(rs.fields("fechaVcto"))
End if
pago=cdbl(rs.fields("monto"))
'montos
if Ucase(rs.fields("TipoDoctoPago"))="EFECTIVO" then 
	MontoEfectivo=MontoTotal
	if MontoTotal<> "&nbsp;" then  tot1=tot1+MontoTotal
End if
if Ucase(rs.fields("TipoDoctoPago"))="CHEQUE" then
	if FechaVcto <= date() then
		MontoalDia  = MontoTotal
		if MontoTotal<> "&nbsp;" then tot2=tot2+MontoTotal
	Else
		Montoafecha=MontoTotal
		if MontoTotal<> "&nbsp;" then tot3=tot3+MontoTotal
	end if
'FechaVcto=cdate(FechaVcto)
End if
%>
  <tr>
    <td align="center"><font face="Arial" size="2"><%=rs.fields("referencia")%></font></td>
    <td align="left"  ><font face="Arial" size="2">&nbsp;<%=ProperCase(rs.fields("razonsocial"))%></font></td>
    <td align="center"><font face="Arial" size="2"><%=rs.fields("operacion")%></font></td>
    <td align="center"><font face="Arial" size="2"><%=replace(rs.fields("fecha"),"2007","07")%></font></td>
    <td align="center"><font face="Arial" size="2"><%=rs.fields("cliente")%></font></td>
    <td align="right" ><font face="Arial" size="2"><%=formatnumber(rs.fields("monto"),0)%>&nbsp;</font></td>
    <td align="center"><font face="Arial" size="2"><%=NroDoctoPago%></font></td>
    <td align="center"><font face="Arial" size="2">&nbsp;</font></td>
    <td align="center"><font face="Arial" size="2"><%=FechaVcto    %>&nbsp;</font></td>
    <td align="right" ><font face="Arial" size="2"><%=Montoafecha  %>&nbsp;</font></td>
    <td align="right" ><font face="Arial" size="2"><%=MontoalDia   %>&nbsp;</font></td>
    <td align="right" ><font face="Arial" size="2"><%=MontoEfectivo%>&nbsp;</font></td>
    <td align="right" ><font face="Arial" size="2"><%=MontoTotal   %>&nbsp;</font></td>
  </tr>


<%
UltNroDoctoPago=trim(rs.fields("NroDoctoPago"))
totalfacturas=totalfacturas+pago
totalfact=totalfact+pago
UltOperacion=trim(rs.fields("operacion"))
MTotal=MTotal+pago
x=x+1
rs.movenext
Loop

%>
<tr>
	<td colspan="5" align="right"><font size="1">
	<%="Total Operacion " & UltOperacion & " : $ " & formatnumber(totalfacturas,0)%>
	&nbsp;</font></td>
	<td align="center"><font size="1">&nbsp;</font></td>
	<td colspan="7"><font size="1">&nbsp;</font></td>
</tr>
<tr>
	<td colspan="5" align="right"><font size="2">
	<%="Total CCD : $ "%>
	&nbsp;</font></td>
	<td align="center"><font size="2"><B><%=formatnumber(totalfact,0)%>&nbsp;</B></font></td>
	<td><font size="2">&nbsp;</font></td>
	<td><font size="2">&nbsp;</font></td>
	<td><font size="2">&nbsp;</font></td>
	<td><font size="2"><%=formatnumber(tot3,0)%>&nbsp;</font></td>
	<td><font size="2"><%=formatnumber(tot2,0)%>&nbsp;</font></td>
	<td><font size="2"><%=formatnumber(tot1,0)%>&nbsp;</font></td>
	<td align="center"><font size="2"><B><%=formatnumber(MTotal,0)%>&nbsp;</B></font></td>

</tr>
<%
response.write(chr(10) & "</table>")

'------------------------------------ 
Function ProperCase(sString) 
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
 ProperCase=replace(ProperCase," ","&nbsp;")
End Function
%>
</body>
</html>

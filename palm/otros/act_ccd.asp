<%
set oConn = server.CreateObject("ADODB.Connection")
oConn.open "provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"
%>
<html>
<head>
<title>Actualiza Ingresos CCD PDA - Distribuidora Err&aacute;zuriz</title>
</head>
<!-- Desarrollado por Cristian Palma -->
<body>
<%
step = request.querystring("step")
if step = "" then call paso01()
if step = "01" then call paso02()
sub paso01()

SQL="SELECT MAX(id) AS ID FROM PDA_CCD_DOC"

set rs = oConn.execute(SQL)
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
     <td align="center" bgcolor="#000080"><b><font color="#FFFFFF">Actualizacion Tabla Pagos CCD </font></b></td>
  </tr>
  <tr>
    <td><br></td>
  </tr>
  <tr>
    <td align="center">Ultimo registro en tabla:&nbsp;<b><%= rs.fields("ID")%></b></td>
  </tr>
  <tr>
    <td align="center"><br></td>
  </tr>
  <tr>
    <td align="center">
	   <form method="get" action="/palm/otros/actualiza.asp">
	   <input type="hidden" name="ID" value="<%= rs.fields("ID")%>">
	   <input type="submit" value="Actualizar Tabla Pagos" style="width:250px">
	   </form>
	   </td>
  </tr>
  <tr>
    <td align="center"><br></td>
  </tr>
  <tr>
    <td align="center"><font color="#CCCCCC" face="Arial" size="-3">Distribuci&oacute;n y Excelencia S.A.</font></td>
  </tr>
</table>
<%
end sub 'paso01
'--------------------------
sub paso02()

SQL="SELECT MAX(id) AS ID FROM PDA_CCD_DOC"
set rs = oConn.execute(SQL)
Miid=cdbl(rs.fields("ID"))+1

SQL="select * FROM PDA_CCD_DOC"
tt=request.form("tt")
set rs=oConn.execute(SQL)

rs.close

redim elem(tt,18)

if request.form( right( "0000000000" & Miid, 10) & "-0") = ""  then exit sub
if isnull(request.form(right("0000000000" & Miid,10) & "-0")) then exit sub

for x = Miid to (Miid+tt-1)
x1=right("0000000000" & x,10)
x2=x-Miid
	elem(x2,0) =request.form( x1 &"-0")
	elem(x2,1) =request.form( x1 &"-1")
	elem(x2,2) =request.form( x1 &"-2")
	elem(x2,3) =request.form( x1 &"-3")
	elem(x2,4) =request.form( x1 &"-4")
	elem(x2,5) =request.form( x1 &"-5")
	elem(x2,6) =request.form( x1 &"-6")
	elem(x2,7) =request.form( x1 &"-7")
	elem(x2,8) =request.form( x1 &"-8")
	elem(x2,9) =request.form( x1 &"-9")
	elem(x2,10)=request.form( x1 &"-10")
	elem(x2,11)=request.form( x1 &"-11")
	elem(x2,12)=request.form( x1 &"-12")
	elem(x2,13)=request.form( x1 &"-13")
	elem(x2,14)=request.form( x1 &"-14")
	elem(x2,15)=request.form( x1 &"-15")
	elem(x2,16)=request.form( x1 &"-16")
	elem(x2,17)=request.form( x1 &"-17")
	elem(x2,18)=request.form( x1 &"-18")
next

for x=0 to (tt-1)
SQL2=SQL & " where id=" & elem(x,0)

rs.Open SQL2 , oConn, 1, 3

if rs.eof then rs.addnew

'	rs.fields(0)  = elem(x,0)  'ID
	rs.fields(1)  = elem(x,1)  'Empresa
	rs.fields(2)  = elem(x,2)  'Operacion
	rs.fields(3)  = elem(x,3)  'ID Vendedor
	rs.fields(4)  = elem(x,4)  'Vendedor
	rs.fields(5)  = elem(x,5)  'Lnea
	rs.fields(6)  = elem(x,6)  'Fecha
	rs.fields(7)  = elem(x,7)  'Tipo Documento
	rs.fields(8)  = elem(x,8)  'Refefrencia
	rs.fields(9)  = elem(x,9)  'Monto
	rs.fields(10) = elem(x,10) 'Tipo documento de pago
	rs.fields(11) = elem(x,11) 'Numero documento pago
	rs.fields(12) = elem(x,12) 'Entidad
	if not len(elem(x,13))  < 1 then elem(x,13)=cdate(elem(x,13)) 'formatea fecha vencimiento
	rs.fields(13) = elem(x,13) 'fecha vencimiento
	rs.fields(14) = elem(x,14) 'Cliente
	rs.fields(15) = elem(x,15) 'Razon social
	rs.fields(16) = elem(x,16) 'Traspasado
	rs.fields(17) = elem(x,17) 'Cuentapago
	rs.fields(18) = elem(x,18) 'prep
	rs.update
  rs.close
next
with response
 .write("<p align='center'>Datos Actualizados</p>")
 .write("<p align='center'><input type='button' value='Salir' onclick="&_
    chr(34)&"location='/palm/otros/act_ccd.asp'"&chr(34)&"></p>")
end with
end sub 'paso02
%>
</body>
</html>

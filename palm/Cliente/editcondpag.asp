<%
set oConn = server.CreateObject("ADODB.Connection")
desa1 = "provider=sqloledb;data source=serverdesa;initial catalog=BDFlexline;User Id=sa;Password=desakey"
desa2 = "provider=sqloledb;data source=SQLSERVER;initial catalog=todo;User Id=sa;Password=desakey"
%>
<html>
<head>
<title>Editor Condici&oacute;n de Pago</title>
<style type="text/css">
body {
	margin:0;
	font:Arial 12px;
}
#title {
	font-weight:bold;
	text-align:center;
	background-color:#000080;
	color:#FFFFFF;
}
</style>
</head>

<body>
<div id="title">Editor Condici&oacute;n de Pago</div>
<%
'inicio programacion
paso=request.querystring("paso")
if len(paso)=0 then paso=request.form("paso")
if paso="" then call cliente()
if paso="Buscar" then call listar()
if paso="cambiar" then call cambiar()
if paso="guardar" then call guardar()
Sub cliente()
%>
<form method="post" action="">
<table width="100%" align="center" cellpadding="2" cellspacing="0">
  <tr>
    <td align="center"><b>Ingrese</b> RUT, Codigo Cliente &oacute; </td>
  </tr>
  <tr>
    <td align="center">Raz&oacute;n Social </td>
  </tr>
  <tr>
    <td align="center">
	   <input name="q" type="text" id="q" style="width:150px"></td>
  </tr>
  <tr>
    <td align="center">
	   <input name="paso" type="submit" id="paso" value="Buscar" 
	   style="width:150px; font-weight:bold;"></td>
  </tr>
</table>
</form>
<%
end sub 'cliente()
sub listar()
oConn.open desa1
buscar=request.form("q")
SQL="SELECT CtaCte + RazonSocial + Sigla AS busqueda, "&_
    "CtaCte, RazonSocial, Sigla, CondPago "&_
    "FROM CtaCte WHERE (Empresa = 'DESA') AND (TipoCtaCte = 'CLIENTE') "&_
	"AND (CtaCte + RazonSocial + Sigla LIKE '%"& buscar &"%') "&_
    "ORDER BY RazonSocial"
set rs = oConn.execute(SQL)
if rs.eof and rs.bof then
response.write "<p align='center'> No se encontro "& buscar &"</p>"
exit sub
end if
%>
<table width="100%" cellspacing="0">
  <tr>
    <th height="19" bgcolor="#669966">C&oacute;digo Cliente </th>
	<td bgcolor="#669966">&nbsp;</td>
  </tr>
<%
do until rs.eof 
%>
  <tr onMouseOut="this.style.backgroundColor='transparent';" onMouseOver="this.style.backgroundColor='#CCCCCC';" onClick="location='editcondpag.asp?paso=cambiar&ctacte=<%= rs.fields("ctacte") %>';" style="cursor:pointer">
  
    <td align="center"><small>
	<a href="editcondpag.asp?paso=cambiar&ctacte=<%= rs.fields("ctacte") %>">
	<%= replace(rs.fields("ctacte")," ","&nbsp;") %></a>
	</small></td>
    <td><small><b>Raz&oacute;n Social :</b> <%= rs.fields("razonsocial") %><br>
    <b>Sigla : </b><%= rs.fields("sigla") %><br>
    <b>Cond : </b><%= rs.fields("condpago") %></small></td>
  </tr>  
  <tr><td colspan="2"><hr></td></tr>
<%
rs.movenext
loop
%>
</table>
<%
oConn.close
set oConn = nothing
end sub 'listar()
sub cambiar()
vend=request.QueryString("vend")
SQL="SELECT CtaCte, RazonSocial, CondPago FROM CtaCte "&_
    "WHERE (Empresa = 'DESA') AND (TipoCtaCte = 'CLIENTE') "&_
    "AND (ctacte = '"& request.querystring("ctacte") &"')"
oConn.open desa1
set rs = oConn.execute(SQL)
'set request.QueryString = empty
%>
<form method="post" action="?vend=<%= vend %>">
  <input type="hidden" name="ctacte" value="<%=  rs.fields("ctacte")%>">
  <input type="hidden" name="condant" value="<%= rs.fields("condpago")%>">
  <input type="hidden" name="condnew" value="EFECTIVO CAMION">
  <input type="hidden" name="paso" value="guardar">
<table width="100%" cellspacing="0">
  <tr>
    <td><b>Cliente : </b></td>
    <td align="center"><%= rs.fields("ctacte")%>&nbsp;</td>
  </tr>
  <tr>
    <td><b>Raz&oacute;n Social </b></td>
    <td align="center"><%= rs.fields("razonsocial") %>&nbsp;</td>
  </tr>
  <tr>
    <td><b>Cond. Pago Anterior </b></td>
    <td align="center"><%= rs.fields("condpago") %>&nbsp;</td>
  </tr>
  <tr>
    <td><b>Nueva Cond. Pago </b></td>
    <td align="center">EFECTIVO CAMION</td>
  </tr>
  <tr>
    <td colspan="2" align="center"><br>
      <input type="submit" value="Aceptar"> <input type="button" value="Cancelar" onClick="history.back()"></td>
    </tr>
</table>
</form>
<%
oConn.close
set oConn = nothing
end sub 'cambiar()
sub guardar()
if len(request.QueryString("vend"))=0 then
saleA="/palm"
else
saleA="/palm/pedido/verifica_cliente.asp"
saleA=saleA&"?vend="& request.QueryString("vend") &"&cliente="& request.form("ctacte")
end if
oConn.open desa2
MiFecha = right(year(date),2) & right("00" & month(date),2) & right("00" & day(date),2)
MiHora  = right("00"& hour(time),2) & right("00" & minute(time),2) & right("00" & second(time),2)

SQLa="SELECT * FROM Flexline.FX_CambiaCondPago WHERE (CTACTE = '"& request.form("ctacte") &"')"
set rs = oConn.execute(SQLa)
if not(rs.eof and rs.bof) then
with response
 .write "<p align='center'>Cliente ya Fue Modificado<br><a href='" & saleA & "'>Salir</a></p>"
end with

exit sub
end if
SQL="INSERT INTO Flexline.FX_CambiaCondPago (ctacte, condpago_ant, fecha, hora) "&_
    "VALUES ('"& request.form("ctacte") &"', "&_
	"'"& request.form("condant") &"', '"& MiFecha &"', '"& MiHora &"')"
oConn.execute(SQL)
oConn.close
oConn.open desa1
SQL="UPDATE CtaCte "&_
    "SET CondPago = '"& request.form("condnew") &"' "&_
    "WHERE (Empresa = 'DESA') AND (TipoCtaCte = 'CLIENTE') AND "&_
	"(CtaCte = '"& request.form("ctacte") &"')"
oConn.execute(SQL)

%>
<p align="center">
Condición pago Modificada<br>
<a href='<%= saleA %>'>Salir</a>
</p>
<%
end sub 'guardar()
'fin programacion
'response.write(replace(request.form,"&","<br>")&"<br>")
'response.write(replace(request.querystring,"&","<br>"))
%>
</body>
</html>

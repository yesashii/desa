<% Dim oConn, Vnd
set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey"
Nuser = Request.Querystring("nuser")
SQL = "SELECT nombre AS vendedor, NU_2 AS Nuser "&_
	  "FROM Flexline.FX_vende_man "&_
	  "WHERE (NU_2 = '" & Nuser & "')"
Set rs0 = oConn.Execute(SQL)

Vnd = rs0.fields("vendedor")
%>
<html>
<head>
<title>Editar Pedido</title>
<style type="text/css">
<!--
body {
	margin:0px;
	font:Arial;
}
td {
	font-size:xx-small;
}
-->
</style>
<script type="text/javascript" language="javascript">
function borraconf(Nnota){
var Nnota, enlace;
if(confirm('¿Realmente desea eliminar la nota '+ Nnota +'?')){
 alert("eliminar");
 enlace="?nuser=<%= Nuser %>&step=del&nota=" + Nnota ;
 document.location=enlace;
}
}
</script>
</head>
<body>
<div align="center"	style="background-color:#000080; color:#FFFFFF"><b>Edici&oacute;n Pedidos</b></div>
<%
If request.QueryString("step") = "" Then Call Lista_Notas()
If request.QueryString("step") = "edit" Then Call Editar()
if request.querystring("step") = "Actualizar" Then Call Actualizar()
if request.querystring("step") = "del" Then Call Eliminar()
Sub Lista_Notas()

fmd = right(year(date), 2) & right("00" & month(date), 2) & right("00" & day(date), 2)

SQL0 =  "SELECT handheld.Flexline.FX_PEDIDO_PDA.nota, handheld.Flexline.FX_PEDIDO_PDA.Cliente, " & _
		"Flexline.CtaCte.RazonSocial, handheld.Flexline.FX_PEDIDO_PDA.oc, " & _
		"handheld.Flexline.FX_PEDIDO_PDA.fechaentrega " & _
		"FROM handheld.Flexline.FX_PEDIDO_PDA LEFT OUTER JOIN " & _
		"Flexline.CtaCte ON handheld.Flexline.FX_PEDIDO_PDA.Cliente = Flexline.CtaCte.CtaCte " & _
		"WHERE (handheld.Flexline.FX_PEDIDO_PDA.vendedor = '"& Vnd &"') AND "&_
		"(handheld.Flexline.FX_PEDIDO_PDA.fecha = '"& fmd &"') AND "&_
		"(handheld.Flexline.FX_PEDIDO_PDA.estado = 'np') AND (Flexline.CtaCte.Empresa = 'DESA') "&_
		"AND (Flexline.CtaCte.TipoCtaCte = 'CLIENTE')" 
'response.write(SQL0)
SQL0 = "select " & _
"right(numero_pedido,4) as Nota , " & _
"idcliente+' '+cast(idsucursal as nvarchar) as Cliente, " & _
"'' as RazonSocial, " & _
"Orden_compra as OC, " & _
"fecha_facturacion as Fechaentrega " & _
"from sqlserver.desaerp.dbo.ped_pedidosenc " & _
"where idvendedor=" & cint(Nuser) & " and sw_estado='I'"
%>
<br>
<center>
<table border="1" cellpadding="0" cellspacing="0" bordercolor="#999999" 
      width="100%" style="border:outset thin #CCCCCC; border-collapse:collapse">
 <tr>
    <td>Del&nbsp;</td>
	<td>Edt&nbsp;</td>
	<td nowrap><b>N&ordm;&nbsp;Nota</b></td>
	<td><b>Cliente</b></td>
 </tr>
<%
Set rs1 = oConn.Execute(SQL0)
if rs1.eof and rs1.bof then
with response
	.write "<tr><td colspan='4'>"
	.write "No hay Notas sin Procesar"
	.write "</td></tr>"
end with
end if

do until rs1.eof
%>
 <tr>
    <td style="cursor:pointer;"><a onClick="borraconf('<%= rs1.fields("Nota") %>')"><img src="delete.jpg" alt="Borrar" width="10" height="10" border="0" /></a>&nbsp;</td>
	<td>
	<!-- <a href="?nuser=<%= Nuser %>&step=edit&nota=<%= rs1.fields("Nota") %>"><img src="edit.jpg" alt="Editar" width="10" height="10" border="0" /></a> -->
	&nbsp;</td>
	<td align="center"><%= rs1.fields("Nota") %></td>
	<td><%= left(rs1.fields("Cliente"), 15) %></td>
 </tr>
<%
rs1.movenext
loop
%> <tr>
   <td colspan="4"><hr></td>
 </tr>
</table>
<%
end sub 'lista_notas
'::::::::::::::::::::::::::::::::::::::::
Sub Editar()
nota = request.QueryString("nota")

SQL="SELECT * FROM handheld.Flexline.FX_PEDIDO_PDA "&_
"WHERE (vendedor = '"& Vnd &"') AND (nota = "& nota &")"
'response.write(SQL)
set rs=oConn.execute(SQL)
%>
<script language="javascript" type="text/javascript">
<!--
function Right(s, n){
	// Devuelve los n últimos caracteres de la cadena
	var t=s.length;
	if(n>t)
		n=t;
		
	return s.substring(t-n, t);
}

function valdesc(){
var x;
   for(x=1 ; x <=12 ; x++ ){
	    z='0'.concat(x); z=Right(z,2);
			 //alert(z);
       if(parseInt(document.getElementById('desc'+z).value)>99){
   			alert('Porcentaje no valido')
			return false;
       }
   }
return true;
}
//-->
</script>
<form method="get" action="" onSubmit="return valdesc()">
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2" align="center" nowrap bgcolor="#E3E9F1"><b>Detalle Pedido Modificar </b></td>
  </tr>
  <tr>
    <td colspan="2" nowrap>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" nowrap><b>Cliente</b></td>
<%
sql2="SELECT RazonSocial FROM Flexline.CtaCte "&_
     "WHERE (CtaCte LIKE '%"& rs.fields("cliente") &"%')"
	 'response.write sql2
set rs1=oConn.execute(sql2)
%>    <% if rs.fields("estado")="OK" then modd = "disabled" %> 
  </tr>
  <tr>
    <td colspan="2" nowrap><%= rs1.fields("razonsocial") %> &nbsp;
    <input type="button" value="Editar" disabled></td>
  </tr>
  <tr>
    <td colspan="2" nowrap><b>Fecha Entrega </b></td>
  </tr>
  <tr>
    <td colspan="2" nowrap>
      <% if isnull(rs.fields("fechaentrega")) or len (rs.fields("fechaentrega")) = 0 then 
	     response.write "20" & rs.fields("fecha") + 1 
		 else
		 response.write rs.fields("fechaentrega")
		 end if
		 %>    </td>
  </tr>
  <tr>
    <td colspan="2" nowrap><hr></td>
  </tr>
<% for x=1 to 12 
if not len(rs.fields("producto"& right("00"& x,2)))=0 then
%>  <tr>
    <td colspan="2" nowrap><b>Producto</b></td>
  </tr>
  <tr>
    <td colspan="2" nowrap><%= rs.fields("producto"& right("00"& x,2)) %> - 
      <%
Sql2="SELECT GLOSA FROM flexline2.dbo.producto "&_
"WHERE empresa='desa' and    (PRODUCTO = '"& rs.fields("producto"& right("00"& x,2)) &"')"
set rs2=oConn.execute(Sql2)
	response.write left(rs2.fields("GLOSA"), 25)
rs2.close	%></td>
  </tr>
  <tr>
    <td nowrap><b>Cantidad</b></td>
    <td nowrap><b>Descuento</b></td>
  </tr>
  <tr>
    <td nowrap><input name="cant<%= right("00"& x,2)%>" type="text" id="cant<%= right("00"& x,2)%>" style="width:30px;" value="<%= rs.fields("cantidad"& right("00"& x,2)) %>" size="4" <%= modd %>></td>
    <td nowrap><input name="desc<%= right("00"& x,2)%>" type="text" id="desc<%= right("00"& x,2)%>" style="width:30px;" value="<%= rs.fields("descuento"& right("00"& x,2)) %>" size="4" <%= modd %>></td>
  </tr>
  <tr>
    <td colspan="2"><hr></td>
  </tr>
<% 
end if
next 
%>
  
  <tr>
    <td colspan="2" align="center"><input name="step" type="submit" value="Actualizar" <%= modd %>></td>
  </tr>
</table>
<input type="hidden" name="nuser" value="<%=  Nuser %>">
<input type="hidden" name="nota" value="<%= nota %>">
</form>
<%
End Sub 'Editar
'::::::::::::::::::::::::::::::::::::::::
Sub Actualizar()
'request.form("nota")
SQLG="SELECT * FROM handheld.Flexline.FX_PEDIDO_PDA "&_
"WHERE (vendedor = N'"& Vnd &"') AND (nota = "& request.querystring("nota") &")"
'response.write(SQLG&"<br>")
set rsG=oConn.execute(SQLG)
rsG.close
rsG.open SQLG, oConn, 1, 3
for each elemento in request.querystring
	if left(elemento,4)="cant" then
		rsG.fields("cantidad"&right(elemento,2))=request.querystring(elemento)
	end if
	if left(elemento,4)="desc" then
		rsG.fields("descuento"&right(elemento,2))=replace(request.querystring(elemento), ".", ",")
	end if
next
'rsG.fields("fechaentrega")=request.form("FENTREGA")
'rsG.fields("estado")=request.form("estado")
rsG.update
rsG.close
%><p align='center'>Nota Actualizada<br><br>
<input type="button" value="Salir" onClick="document.location='?nuser=<%= Nuser %>'"></p>
<%
End Sub 'actualizar
Sub Eliminar()

SQL = "UPDATE handheld.Flexline.FX_PEDIDO_PDA "&_
	  "SET ESTADO = 'noventa'" &_
	  "WHERE (vendedor = N'"& Vnd &"') AND (nota = "& request.querystring("nota") &")"
SQL="update sqlserver.desaerp.dbo.ped_pedidosenc " & _
	"SET SW_ESTADO='N' " & _
	"WHERE idvendedor=" & cint(nuser) & " and sw_estado='I' and right(Numero_pedido,4)='" & request.querystring("nota") & "'" 
oConn.execute(SQL)

%>
<p align="center">
		Nota Eliminada
		<br>
		    <br>
<input type="button" value="Salir" onClick="document.location='?nuser=<%= Nuser %>'">
</p>
<%
End Sub 'Eliminar
%>
<%
dim vend, cliente

vend    = request.querystring("vend")
cliente = trim(request.QueryString("cliente"))

%>
<html>
<head>
<title>noventa</title>
</head>
<body text="#000000" style="font:arial">
<%
response.Write("<form method='post' action='noventa_t.asp?vend=" & vend & "&cliente=" & cliente& "'>")
%>
<table width="100%" border="0">
<tr>
  <td height="29" align="center" valign="middle"><strong>MOTIVO VENTA NO EFECTUADA</strong></td>
</tr>
<tr>
  <td height="29" align="center" valign="middle"><strong><%=cliente%></strong></td>
</tr>
<tr>

    <td height="37" align="center">
		<select name="select" size="0">
			<option value="OK Con Producto"      >Centralizado - OK Con Producto</option>
			<option value="Sin Producto"         >Centralizado - Sin Producto</option>
			<option value="Atraso"               >Credito - Atraso</option>
			<option value="Sobrepasado"          >Credito - Sobrepasado</option>
			<option value="Sin Telefono"         >No Ubicado - Sin Telefono</option>
			<option value="Equivocado"           >No Ubicado - Equivocado</option>
			<option value="No Contestan"         >No Ubicado - No Contestan</option>
			<option value="No Se Encuentra Encargado">No Ubicado - No Se Enc. Encargado</option>
			<option value="N.Celular/L.Distancia">N.Celular/L.Distancia</option>
			<option value="Con Stock"            >No Venta - Con Stock</option>
			<option value="No Compra Por Precio" >No Venta - No Compra Por Precio</option>
			<option value="Compra a Vendedor"    >No Venta - Compra a Vendedor</option>
        </select></td>
  </tr>
  <tr align="center" valign="middle">
    <td height="34"><strong>Observaciones</strong> : </td>
  </tr>
  <tr align="center" valign="middle">
    <td height="67"><textarea name="textarea" cols="25" rows="3"></textarea></td>
  </tr>
  <tr>
    <td height="45" align="center" valign="middle"><input type="submit" value="Enviar">
      </td>
    </tr>
</table>
</form></br>
</body>
</html>
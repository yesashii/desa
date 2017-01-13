<%
dim vend, cliente

vend    = request.querystring("vend")
vend=replace(vend,"NN","Ñ")'Error blackberry
vend=replace(vend,"?","Ñ")'Error blackberry
cliente = trim(request.QueryString("cliente"))

%>
<html>
<head>
<title>noventa</title>
</head>
<body text="#000000" style="font:arial">
<%
response.Write("<form method='post' action='noventa_t.asp?vend=" & replace(vend,"Ñ","NN") & "&cliente=" & cliente& "'>")
%>
<table width="100%" border="0">
<tr>
  <td height="29" align="center" valign="middle"><strong><FONT SIZE="2" face="verdana" COLOR="#000066">MOTIVO VENTA NO EFECTUADA</FONT></strong></td>
<tr>
    <td height="37" align="center"><select name="select" size="0">
        <option value="">SELECCIONE</option>
        <option value="Local Cerrado">Local Cerrado</option>
        <option value="Tiene Stock">Tiene Stock</option>
        <option value="No esta el encargado o du">No esta el encargado o dueño</option>
        <option value="Cliente Bloqueado por ced">Cuenta corriente Bloqueada</option>
		<option value="Por diferencia de precio ">Por Dif. precio Compra a Dist.</option>
		<option value="Pendiente por pate de DES">Pend. Generar N/C o Canje Publicit.</option>
		<option value="SIN ESPECIFICAR">SIN ESPECIFICAR</option>
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
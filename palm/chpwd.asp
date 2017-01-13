<% dim usuario, empresa, oConn

set oConn = server.CreateObject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey;"
usuario = request.QueryString("usuario")
empresa = request.QueryString("empresa")

pwd = request.form("pwd")

if len(pwd) > 0 then

sql="UPDATE V " &_
"SET V.psw_pedidos = " & pwd & " " &_
"FROM SQLSERVER.desaerp.dbo.dim_vendedores AS V INNER JOIN " &_
"SQLSERVER.desaerp.dbo.dim_empresas AS E ON V.idempresa = E.idempresa " &_
"WHERE     (V.idvendedor = "& usuario &") AND (E.nombre = '" & empresa & "')"

sql2="UPDATE V " &_
"SET V.psw_pedidos = " & pwd & " " &_
"FROM dbo.dim_vendedores AS V INNER JOIN " &_
"dbo.dim_empresas AS E ON V.idempresa = E.idempresa " &_
"WHERE     (V.idvendedor = '" & usuario & "') AND (E.nombre = '" & empresa & "')"

oConn.execute(sql)
oConn.execute(sql2)

with response
	.Write "<script language='javascript'>"
	.Write "alert('Password Cambiada');"
	.write "window.location='./'"
	.write "</script>"
response.Flush()
'response.redirect("./")
end with
end if

%><html>
<head>
<title>Sistema Ventas</title>
</head>
<style type="text/css">
html,body{
height:100%;
font-size-adjust:inherit;
}
#wrap{
min-height:90%;
}
#footer{
margin-top:-20px;
height: 20px;
}
</style>
<script type="text/javascript" language="javascript">
function validarfrm(){
pwd  = document.getElementById('pwd').value;
rpwd = document.getElementById('rpwd').value;
if(pwd == '' || rpwd == ''){
document.getElementById('msgcheck').innerHTML='el campo esta vacio'
return false;
}
if(pwd != rpwd){
document.getElementById('msgcheck').innerHTML='Las contrase&ntilde;as no coinciden'
return false;
}else{
return true;}
}
</script>
<body>
<br/>
<div id="wrap">
<table width="800px" align="center" style="border-collapse: collapse;">
  <tr height="20px">
	<td width="20px">&nbsp;</td>
	<td width="150px" bgcolor="#666633" align="center" style="color:#FFFFFF; font-size:small">Sistema Ventas DESA</td>
	<td width="*" align="right" style="font-size:small">&nbsp;</td>
  </tr>
  <tr bgcolor="#666633" height="10px">
	<td colspan="3"></td>
</tr>
</table>
<div>
<br/>
<br/>
<form method="post" onSubmit="return validarfrm()">
<table align="center">
  <tr>
    <td>Nueva Contrase&ntilde;a</td>
	<td><input name="pwd" id="pwd" type="password"></td>
  </tr>
  <tr>
    <td>Repita Contrase&ntilde;a</td>
	<td><input name="rpwd" id="rpwd" type="password"></td>
  </tr>
    <tr>
    <td colspan="2" id="msgcheck" align="center" style="color:#FF0000">&nbsp;</td>
  </tr>
    <tr>
    <td></td>
	<td><input type="submit" value="Cambiar"></td>
  </tr>
</table>
</form>
</div>
</div>
<div id="footer">
<table align="center">
<tr>
<td>&copy; 2014 - Distribuci&oacute;n y Excelencia S.A.</td>
</tr>
</table>
</div>
</body>
</html>

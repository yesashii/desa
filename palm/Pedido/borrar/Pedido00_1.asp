<html>
<%
vend = trim(request.querystring("vend"))
':::::::::::::::::: conexion :::::::::::::::::
Private tipodoc, mitotal, oConn, rs, sql, rs1
Set oConn2 = server.createobject("ADODB.Connection")
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
'oConn.open "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=todo;User Id=sa;Password=desakey;"
oConn2.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"


SQL="SELECT     Flexline.Rutas.Territorio, Flexline.Rutas.CodSem, Flexline.Rutas.Dia, Flexline.Rutas.RutFlex, Flexline.Rutas.RazonSocial, Flexline.CtaCte.CodLegal " & _
"FROM         Flexline.FX_CODGEN INNER JOIN " & _
"Flexline.Rutas ON Flexline.FX_CODGEN.Valor = Flexline.Rutas.CodSem INNER JOIN " & _
"Flexline.CtaCte ON Flexline.Rutas.RutFlex = Flexline.CtaCte.CtaCte " & _
"WHERE (Flexline.FX_CODGEN.CodGen = N'Semana Ruta') AND (Flexline.FX_CODGEN.Codigo = DATEPART(wk, GETDATE())) AND " & _
"(Flexline.Rutas.Vendedor = N'MAXIMILIANO MARTINEZ') AND (Flexline.Rutas.CodDia = DATEPART(q, GETDATE())) AND " & _ 
"(Flexline.CtaCte.Empresa = 'desa')"

Set rs=oConn2.execute(Sql)

%>
<head>
<meta http-equiv="Content-Language" content="es-cl">
<title>Seleccionar Cliente</title>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-width: 0" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%" height="58" align="left" style="border-style: none; border-width: medium">
      <div align="center">
        <p><b><font face="Arial" size="2" color="#333333">
		<!-- <img src="file://///Brio/wwwroot/Flexmovil/logo.JPG" width="32" height="20"> -->
		Sistema Preventa</font></b></p>
      </div></td>
  </tr>
  <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="center">
    <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
      <tr>
        <td width="100%" align="center" valign="middle" bgcolor="#000080">
          <font face="Arial" size="2" color="#FFFFFF">Listado Clientes</font></td>
      </tr>
<%
do until rs.eof
	response.write("<tr>")
	response.write("<td width='100%' align='center'><FONT face='Arial' SIZE='2'>&nbsp;" & _
	"<a href='Pedido01.asp?vend=" & vend & "&cliente=" & rs.fields("CodLegal") & "' style='text-decoration: none'>" & _
	replace(rs.fields("RazonSocial")," ","&nbsp;") & _ 
	"</a></FONT></td>")
	response.write("</tr>")
rs.movenext
loop
%>
    </table>
    </td>
  </tr>
</table>
<p align="center"><font face="Arial" size="1" color="#808080"></font></p>

</body>

</html>
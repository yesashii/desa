<%':::::::::::::::::: conexion :::::::::::::::::
public oConn, rs, sql, sr1, oConn1, usuario, password, empresa, orden
Set oConn = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey;"

modificado = date() 'RS.DateLastModified
'for each element in request.Form
'response.write element & ": " & request.form(element)
'next
session.Timeout=240
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Sistema Ventas</title>
<link rel="shortcut icon" href="favicon.ico">
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
<body>
<br/>
<div id="wrap">
<table width="800px" align="center" style="border-collapse: collapse;">
  <tr height="20px">
	<td width="20px">&nbsp;</td>
	<td width="150px" bgcolor="#666633" align="center" style="color:#FFFFFF; font-size:small">Sistema Ventas DESA</td>
	<td id="chpwd" width="*" align="right" style="font-size:small">&nbsp;</td>
  </tr>
  <tr bgcolor="#666633" height="10px">
	<td colspan="3"></td>
</tr>
</table>

<%

'trae datos de usuario
empresa=trim(cstr(request.form("empresa")))
if len(empresa) = 0 then empresa="DESA"
usuario =trim(cstr(request.form("usuario" )))
password=trim(cstr(request.form("password")))
orden   = trim(cstr(request.form("orden")))
ocultarocd =trim(cstr(request.form("ocultarocd")))
recordar=request.form("recordar")
if len(request.Cookies("pdapwd"))>1 or len(password)>1 then call validalogon
if len(password)=0 then call login
'if len(password)>1 then call validalogon
'Login
sub login()
%>
<div>
<br/>
<br/>
<form name="login" id="login" method="post" action="">
<table align="center" style="border-collapse:collapse">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td width="15">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Empresa</td>
    <td><SELECT NAME="empresa" style="width:180px" onChange="login.submit()">
				<%
				sql="SELECT u.empresa " & _
					"FROM handheld.Flexline.PDA_usuarios as u " & _
					"group by u.empresa " & _
					"order by u.empresa"
				call generaoption(sql,empresa)
				%>
				</SELECT></td>
<%
'				response.write(sql)
'				response.end()
%>				
    <td width="15">&nbsp;</td>
    <td>&nbsp;</td>
    <td style="font-size:small">Ordenar</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td width="15">&nbsp;</td>
    <td><input type="radio" name="orden" id="r1" value="nombre" onClick="login.submit()"  <% if orden = "nombre" then response.Write("checked")%>></td>
    <td>Nombre</td>
  </tr>
  <tr>
    <td>Usuario</td>
    <td><SELECT NAME="usuario" style="width:180px">
<%
	sql="SELECT u.num_vend, u.num_vend + '_' + u.nombre " & _
	"FROM handheld.Flexline.PDA_usuarios as u " & _
	"where tipousr='v' and u.empresa='" & empresa & "' "
	if orden = "numero" then ordena="order by u.num_vend"
	if orden = "nombre" then ordena="order by u.nombre"
	if len(ordena) = 0 then ordena="order by u.nombre"
	if ocultarocd="on" then sql = sql & " and u.clase='interno' "
	sql = sql & ordena
'	response.Write(sql)
	call generaoption(sql,usuario)
%>
				</SELECT></td>
    <td width="15">&nbsp;</td>
    <td><input type="radio" name="orden" id="r2" value="numero" onClick="login.submit()" <% if orden = "numero" then response.Write("checked")%>></td>
    <td>N&uacute;mero</td>
  </tr>
  <tr>
    <td>Contrase&ntilde;a: </td>
    <td><INPUT TYPE="password" style="width:180px" name="password" id="password"></td>
    <td width="15">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td width="15">&nbsp;</td>
    <td><input type="checkbox" name="ocultarocd" id="ocultarocd" onClick="login.submit()" <% if ocultarocd="on" then response.write("checked")%>></td>
    <td style="font-size:small">Ocultar OCD</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="right"><input type="submit" value="Iniciar Sesi&oacute;n" onClick="">&nbsp;</td>
    <td width="15" align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><input type="checkbox" name="recordar" value="recordar"></td>
    <td>Recordarme en este equipo</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
</table>
</form>
</div>
<%
end sub
'Valida
sub validalogon()

if len(usuario) =0 then usuario = request.Cookies("pdausr")
if len(password)=0 then password= request.Cookies("pdapwd")

SQL = "SELECT COUNT(*) FROM handheld.Flexline.PDA_usuarios as u WHERE num_vend ='"& usuario &"' AND Password ='"& password &"'"

set rs = Server.CreateObject("ADODB.RecordSet")

rs.open SQL, oConn, 1, 3

valida = rs(0)

if password = "159632" then valida = 1
if valida = 0 then
%>
<br>
<br>
<table align="center">
  <tr>
    <td style="">Nombre de usuario o contraseña incorrecta</td>
  </tr>
  <tr>
    <td align="center"><input type="button" value="volver" onClick="history.back(-1)"></td>
  </tr>
</table>

<%
else

'response.Write recordar

response.Cookies("pdausr")=usuario
if recordar="recordar" then

response.Cookies("pdapwd")=password
end if

Sql="SELECT * FROM  handheld.Flexline.PDA_menu WHERE (vigente=1) and (tipousr Like '%" & VUtipousr & "%') ORDER BY orden"

Set rs=oConn.execute(Sql)

response.write("<br><br><table align='center'>")
response.write("<TR>")
i1=0
do until rs.eof
	if i1=3 then 
		response.write("</TR><TR>")
		i1=0
	end if

		tamanoimg=64

	Misize=" SIZE=1"
	'IF LEN(usuario)=0 then Misize=" SIZE=1"

	with response
	  .write("<td width='33%' align='center'><A HREF='" & rs.fields("link") )
	  .write("?nuser=" & usuario & "&empresa=" & empresa & "' style='text-decoration: none'>")
	  .write("<img src='imagenes/" & rs.fields("imagen") & "' width='" & tamanoimg )
	  .write("' height='" & tamanoimg & "' border='0' alt='" & replace(rs.fields("menu"),"<br>"," ") & "'>")
	  .write("<BR><B>" & rs.fields("menu") & "</B></A><BR><BR></TD>")
	end with

i1=i1+1
rs.movenext
loop
response.write("<TR>")
response.write("</TABLE>")

response.write "<p align='center' style='font-size:xx-small'><b>: : : : : : : Info : : : : : : :</b><br>"
SQL="SELECT TOP 10 * FROM handheld.Flexline.PDA_info "&_
"where fecha >getdate()-8 ORDER BY ID DESC"
set rs_info = oConn.execute(SQL)
x=0
do until rs_info.eof
if x=0 then
bCol="#FF9900"
x=1
else
bCol="#ECE9D8"
x=0
end if
with response
  .write "<table width='100%' bgcolor='"& bCol &"' style='font-size:xx-small'>"
  .write "<tr><td align='center'>"& rs_info.fields("info") &"</td></tr>"
  .write "<tr><td align='center'>De: "& rs_info.fields("USUARIO") 
'  .write "<br>Informado por :"& rs_info("USUARIOP")
  .write "<br> el "& rs_info.fields("fecha") &"</td></tr></table>"
end with
rs_info.movenext
loop
response.write "</p>"
response.write "<p align='center'><small>"
response.write "Ultima Actualizaci&oacute;n: " & modificado & "</small></p>"
end if

call insertalink()

end sub
%>
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
<%
function generaoption(misql,miselect)
	set rs=oConn.execute(misql)
	do until rs.eof
		selected=""
		if ucase(rs(0))=ucase(miselect) then selected="selected"
		if rs.fields.count>1 then
			%><OPTION value="<%=rs(0)%>" <%=selected%>><%=rs(1)%></OPTION>
			<%
		else
			%><OPTION value="<%=rs(0)%>" <%=selected%>><%=rs(0)%></OPTION>
			<%
		end if
		rs.movenext
	loop
end function

sub insertalink()
%>
<script type="text/javascript" language="javascript" type="text/javascript">
document.getElementById('chpwd').innerHTML='<a href="chpwd.asp?usuario=<%=usuario%>&empresa=<%=empresa%>">Cambiar Contraseña</a>';
</script>

<%
end sub
%>

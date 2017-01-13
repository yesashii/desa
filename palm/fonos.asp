<%
Dim tipodoc, mitotal, oConn, rs, sql, oconn1, rs1
Set oConn = server.createobject("ADODB.Connection")
oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<TITLE>LISTADO CELULARES</TITLE>
</HEAD>
<BODY>
<CENTER>
<FONT SIZE="1" COLOR="#727272">Lista De Celulares</FONT>

<%
if request.querystring("zona")="vendedores" then
SQL="select * from fonos where zona <> 'oficina' order by nombre"
n1="<A HREF='fonos.asp' >Oficina</a>"
n2="<B>Vendedores</B>"
else
SQL="select * from fonos where zona='oficina' order by nombre"
n1="<B>Oficina</B>"
n2="<A HREF='fonos.asp?zona=vendedores' >Vendedores</a>"
end if


Set rs=oConn.execute(SQL)

dim micolor
Micolor="#FFFF99"

response.write("<TABLE>")

response.write("<TR>")
	response.write("<TD>" & n1 & "</TD>")
	response.write("<TD>" & n2 & "</TD>")
response.write("</TR>")

response.write("<TR>")
	response.write("<TD></TD>")
	response.write("<TD></TD>")
response.write("</TR>")

response.write("<TR>")
	response.write("<TD>Numero</TD>")
	response.write("<TD>Nombre</TD>")
response.write("</TR>")

response.write("<TR>")
do until rs.eof

response.write("<TR>")
	response.write("<TD BGCOLOR=" & micolor & ">" & rs.fields("Numero") & "</TD>")
	response.write("<TD BGCOLOR=" & micolor & ">" & rs.fields("nombre") & "</TD>")
response.write("</TR>")

	if micolor="#FFFF99" then
		micolor="#CCFFFF"
	else
		micolor="#FFFF99"
	end if 
rs.movenext
loop
response.write("</TR>")
response.write("</TABLE>")

%>
<FORM METHOD=POST ACTION="Correos.asp">
<INPUT TYPE="submit" value="Listado de Correos">
</FORM>

</BODY>
</CENTER>
</HTML>

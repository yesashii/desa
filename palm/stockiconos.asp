<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Stock Iconos</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY bgcolor="#C0C0C0">
<%
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql
Set oConn = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=BDgestionweb;User Id=sa;Password=desakey;"

nusuario=request.querystring("nuser")

SQL="SELECT     TOP 100 PERCENT flexline.FX_Stock_Dia1.Bodega, flexline.FX_Stock_Dia1.Producto, dbo.Productos.GLOSA, flexline.FX_Stock_Dia1.Stock, " & _
"flexline.FX_Stock_Dia1.Fecha " & _
"FROM         flexline.FX_Stock_Dia1 INNER JOIN " & _
"dbo.Productos ON flexline.FX_Stock_Dia1.Producto = dbo.Productos.PRODUCTO INNER JOIN " & _
"todo.Flexline.FX_vende_man ON flexline.FX_Stock_Dia1.Bodega = Todo.Flexline.FX_vende_man.bodega " & _
"WHERE (flexline.FX_Stock_Dia1.Producto LIKE 'vn34%') AND (dbo.Productos.EMPRESA = 'desa') AND (flexline.FX_Stock_Dia1.Stock > 0) AND " & _ 
"(Todo.Flexline.FX_vende_man.NU_2 = N'" & nusuario & "') " & _
"ORDER BY flexline.FX_Stock_Dia1.Bodega DESC, flexline.FX_Stock_Dia1.Producto"
'response.write(sql)
Set rs=oConn.execute(Sql)

if not rs.eof or Not rs.bof then
	response.write("<B>Stock a la Fecha : " & rs.fields("fecha")-2 & "</B><HR>")
End if

%>

<CENTER>
<TABLE border="1" bordercolor="#004080">
<TR bgcolor="#FFFFFF">
	<TD><CENTER><B>Bodega</B     ></CENTER></TD>
	<TD><CENTER><B>Codigo</B     ></CENTER></TD>
	<TD><CENTER><B>Descripcion</B></CENTER></TD>
	<TD><CENTER><B>Botellas</B   ></CENTER></TD>
</TR>
<%
do until rs.eof
	%>
	<TR>
		<TD><%=rs.fields("bodega"          )%></TD>
		<TD><CENTER><%=rs.fields("Producto")%></CENTER></TD>
		<TD><%=rs.fields("glosa"           )%></TD>
		<TD><CENTER><%=rs.fields("stock"   )%></CENTER></TD>
	</TR>


	<%
rs.movenext
loop
%>
</TABLE>
<BR>
<input type="button" value="<< Menu" onClick="history.back()">
</CENTER>
</BODY>
</HTML>

<!--#include Virtual="/includes/migrilla.asp"-->
<!--#include Virtual="/includes/pivot.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
pivot
<%
sql="select 'hola' as saludo, 'chao' as adios, 20 as dato"
sql = "select top 100 bodega, producto, glosa, total from serverdesa.BDFlexline.flexline.DESA_STOC�"

', 1 AS [MUnidades], 1/p.FACTORALT AS [MCajas], p.VOLUMEN / 9000 AS MCajas9L 

sql="select  s.bodega , p.Producto , s.Glosa , s.saldo " & _
"from bdgestionweb.dbo.Productos as p INNER JOIN " & _
"bdgestionweb.flexline.TMP_SALDOS_STOCK as s " & _
"	ON p.PRODUCTO = s.Producto INNER JOIN bdgestionweb.flexline.Usuarios_marca as u " & _
"	ON p.TIPO = u.Marca " & _
"WHERE (u.Usuario = 'pernod') and (p.FACTORALT <> 0) AND (p.VOLUMEN <> 0) and s.bodega<>'quebrazones' " & _
"group by s.bodega, s.Glosa , s.saldo, p.FACTORALT, p.VOLUMEN, p.producto " & _
"order by s.Glosa, s.bodega"
call simplepivot(sql)
'response.write(sql)
'call Migrilla(sql,"", "", "", "")
%>
</BODY>
</HTML>
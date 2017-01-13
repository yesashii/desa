<!--#include Virtual="/includes/migrilla.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Detalle</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY><FONT SIZE="2" face="verdana" COLOR="#000066">
<%
nota = trim(cstr(request.form("nota")))
if len(nota)=0 then nota=request.querystring("nota")

sql="SELECT * FROM sqlserver.desaerp.dbo.PED_PEDIDOSENC WHERE (numero_pedido = N'" & nota & "')"
formato="border='1' cellpadding='0' cellspacing='0' background='fondo.GIF' style='border-collapse: collapse; border-left-width: 0; border-top-width: 0' bordercolor='#808080'"
response.write("<B>PED_PEDIDOSENC</B>")
call migrilla(sql, 0, 0,formato ,"")

sql="SELECT * FROM sqlserver.desaerp.dbo.PED_PEDIDOSDET WHERE (numero_pedido = N'" & nota & "')"
formato="border='1' cellpadding='0' cellspacing='0' background='fondo.GIF' style='border-collapse: collapse; border-left-width: 0; border-top-width: 0' bordercolor='#808080'"
response.write("<B>PED_PEDIDOSDET</B>")
call migrilla(sql, 0, 0,formato ,"")

%></FONT>

</BODY>
</HTML>

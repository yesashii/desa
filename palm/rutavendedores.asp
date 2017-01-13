<!--#include Virtual="/includes/migrilla.asp"-->
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

<%
idlocal=request.querystring("idlocal")
sql="SELECT     r.vendedor, r.RutFlex, r.sigla,r.RazonSocial, r.direccionenvio , r.Telefono " & _
"FROM         handheld.Flexline.PDA_RUTA_HOY AS r INNER JOIN " & _
"                      handheld.dbo.DIM_VENDEDORES AS v ON r.vendedor = v.nombre " & _
"WHERE     (r.empresa = 'desa') AND (v.idempresa = 1) AND (v.idlocal = " & idlocal & ") and (v.idmgrupoventa<>25) " & _
"ORDER BY r.vendedor"

call migrilla(sql, 2, 0,"","")

%>

</BODY>
</HTML>

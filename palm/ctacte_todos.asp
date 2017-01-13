<html>
<head>
<title>Mantenedor Clientes (Editor)</title>
<style type="text/css">
<!--
body, th {
	font-family: Arial;
	font-size: 14px;
	background-color: #E5E5E5;
	cursor: default;
	margin-top: 20px;
}
.Barra {
	background-color: #000080;
	color: #FFFFFF;
}
input {
	font-family: arial;
	font-size: 10px;
	text-transform: capitalize;
}
.tit1 {
	font-size: 11px;
	text-transform: capitalize;
	font-weight: bold;
}
select {
	font-size: 10px;
	text-transform: capitalize;
}
.fra1 {
	text-transform: capitalize;
}
-->
</style>
</head>
<%
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql
Set oConn = server.createobject("ADODB.Connection")
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=HANDHELD;User Id=sa;Password=desakey;"

SQL="SELECT nombre FROM Flexline.FX_vende_man where empresa<>'no' group by nombre order by nombre"
'PALM_VENDEDOR.DESCRIPCION
Set rs=oConn.execute(Sql)
%>
<FORM METHOD=POST ACTION="ctacte.asp">
Empresa 
<SELECT NAME="empresa" id="empresa">
	<OPTION SELECTED value='DESA'>Distribuci&oacute;n y Excelencia S.A.</option>
	<OPTION         value='LACAV'>Distribuidora LA CAV</option>
	<OPTION     value='UNDURRAGA'>Undurraga</option>
</SELECT>
Vendedor <SELECT NAME="vendedor" id="vendedor">
<OPTION value='OFICINA'          >OFICINA</OPTION>
<OPTION value='VENDEDOR PERSONAL'>VENDEDOR PERSONAL</OPTION>
<OPTION value='SIN MOVIMIENTO'   >SIN MOVIMIENTO</OPTION>
<%
do until rs.eof
	response.write("<OPTION value='" & rs.fields("nombre") & "'>" & rs.fields("nombre") & "</OPTION>")
	rs.movenext
loop
%>
</SELECT>  

<INPUT TYPE="submit" value="ver reporte">
&nbsp;&nbsp;<INPUT TYPE="checkbox" NAME="excel"><FONT SIZE="2" face="arial" COLOR="#000040">Exportar&nbsp;a&nbsp;Excel</FONT>
</FORM>
<div style="page-break-before:always"></div>
<HR>
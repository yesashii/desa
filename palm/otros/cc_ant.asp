<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> ccd </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="Simon Hernandez">
<META NAME="Keywords" CONTENT="Distribuci&oacute;n y Excelencia">
<META NAME="Description" CONTENT="PDA">
</HEAD>

<BODY>
<CENTER>

<%'---------------------------------------------------------------------------------------------------------
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql, oconn1, rs1
'DIM BdNombre, BDdireccion, BDFono, BDServer, BDuser, BDpass, BDimagen, bdbase
'Dim VUNombre, VUtipousr, VUnum_vend
Set oConn = server.createobject("ADODB.Connection")
'Set oConn1 = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey"

'recupera datos
vendedor=request.querystring("nuser")
Cliente=request.querystring("cliente")
Convend=""
if len(vendedor)<>0 then 
	'vendedor =trim(cstr(request.form("vendedor")))
	 Convend="(FX_vende_man.NU_2='" & right(Vendedor,2) & "') AND " 
end if


B1 =trim(cstr(request.form("B1" )))

%><form method="POST" action="ccd.asp" ><% 

call paso03()
	

call rescatavalor()
%></form><% 


call buscavalor()
'---------------------------------------------------------------------------------------------------------
Sub Paso03()
'SQL="SELECT FX1_CtaCte_Deudas.Nreferencia, FX1_CtaCte_Deudas.FechaVcto, 'Sum(FX1_CtaCte_Deudas.Debe_Ingreso-FX1_CtaCte_Deudas.Haber_ingreso) 'Saldo', FX1_CtaCte_Deudas.Ejecutivo, 'FX1_CtaCte_Deudas.AUX_VALOR6, FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, FX1_CtaCte_Deudas.Direccion, 'FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, FX1_CtaCte_Deudas.Referencia " & _
'"FROM master.dbo.FX_Saldos_CtaCte FX_Saldos_CtaCte, master.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas, 'BDFlexline.dbo.PALM_VENDEDOR PALM_VENDEDOR " & _
'"WHERE FX_Saldos_CtaCte.Referencia = FX1_CtaCte_Deudas.Referencia AND PALM_VENDEDOR.DESCRIPCION = 'FX1_CtaCte_Deudas.Ejecutivo " & _
'"GROUP BY FX1_CtaCte_Deudas.Nreferencia, FX1_CtaCte_Deudas.FechaVcto, FX1_CtaCte_Deudas.Ejecutivo, 'FX1_CtaCte_Deudas.AUX_VALOR6, FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, FX1_CtaCte_Deudas.Direccion, 'FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, FX1_CtaCte_Deudas.Referencia, PALM_VENDEDOR.NV " & _
'"HAVING " & Convend & "(FX1_CtaCte_Deudas.AUX_VALOR3='" & Cliente & "') " & _
'"ORDER BY FX1_CtaCte_Deudas.FechaVcto, FX1_CtaCte_Deudas.Nreferencia"
'response.write(sql)
SQL="SELECT FX1_CtaCte_Deudas.Nreferencia, min(FX1_CtaCte_Deudas.FechaVcto) as 'FechaVcto', " & _
"SUM(FX1_CtaCte_Deudas.Debe_Ingreso - FX1_CtaCte_Deudas.Haber_ingreso) AS Saldo, FX1_CtaCte_Deudas.Ejecutivo, " & _ 
"MAX(FX1_CtaCte_Deudas.AUX_VALOR6) AS AUX_VALOR6, FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, " & _
"FX1_CtaCte_Deudas.Direccion, FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, FX1_CtaCte_Deudas.Referencia " & _
"FROM master2.dbo.FX_Saldos_CtaCte FX_Saldos_CtaCte INNER JOIN "&_
"master2.dbo.FX1_CtaCte_Deudas FX1_CtaCte_Deudas ON FX_Saldos_CtaCte.Referencia = "&_
"FX1_CtaCte_Deudas.Referencia INNER JOIN Flexline.FX_vende_man ON "&_
"FX1_CtaCte_Deudas.Ejecutivo = Flexline.FX_vende_man.nombre "&_
"GROUP BY FX1_CtaCte_Deudas.Nreferencia, FX1_CtaCte_Deudas.Ejecutivo, "&_
"FX1_CtaCte_Deudas.AUX_VALOR3, FX1_CtaCte_Deudas.RazonSocial, "&_
"FX1_CtaCte_Deudas.Direccion, FX1_CtaCte_Deudas.Condpago, FX1_CtaCte_Deudas.Giro, "&_
"FX1_CtaCte_Deudas.Referencia, Flexline.FX_vende_man.NU_2 " & _
"HAVING " & Convend & "(FX1_CtaCte_Deudas.AUX_VALOR3 = '" & Cliente & "') " & _
"ORDER BY FX1_CtaCte_Deudas.FechaVcto, FX1_CtaCte_Deudas.Nreferencia"
'response.write(SQL)
Set rs=oConn.execute(Sql)

%>
<input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" endspan i-checksum="43374" --><p align="center">
  <b><font face="Arial" size="2" color="#000080">ListadoFacturas</font></b></p>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="139">
    <tr>
      <td width="24%" bgcolor="#000080" align="center" height="16"><b>
      <font face="Arial" size="2" color="#FFFFFF">Numero</font></b></td>
      <td width="34%" bgcolor="#000080" align="center" height="16"><b>
      <font face="Arial" size="2" color="#FFFFFF">Vencimiento</font></b></td>
      <td width="121%" bgcolor="#000080" align="center" height="16"><b>
      <font face="Arial" size="2" color="#FFFFFF">Monto</font></b></td>
    </tr>
<%
'MiNreferencia = cdbl(rs.fields("nreferencia"))
'MiFechavcto = rs.fields("FechaVcto")
'Saldofac=0

'FacturaAnt=rs.fields("nreferencia")
Saldopago=cdbl(request.form("Monto" ))
Sichecked="checked"
do until rs.eof
'saldofac=saldofac + cdbl(rs.fields("Debe_Ingreso")) - cdbl(rs.fields("Haber_ingreso")) 
	if cdate(rs.fields("FechaVcto")) <= date() then
		Colorven="#CC3300"
	else
		Colorven="#000080"
	end if
'if FacturaAnt<>rs.fields("nreferencia") then
if Saldopago < 1 then Sichecked=""
	response.write("<tr>")
	  ' value='ON'
	  response.write("<td width='24%' align='center' height='20'><font face='Arial' size='2'>")
	  response.write( cdbl(rs.fields("nreferencia")) & "</font></td>")
	  response.write("<td width='34%' align='center' height='20'><b>")
	  response.write("<font face='Arial' size='2' color='" & Colorven & "'>")
	  response.write( rs.fields("FechaVcto") & "</font></b></td>")
	  response.write("<td width='121%' align='center' height='20'><font face='Arial' size='2'>")
	  response.write( rs.fields("saldo") & "</font></td>")
	response.write("</tr>")
	'saldofac=0
'end if

Saldopago = Saldopago - cdbl(rs.fields("saldo"))
'FacturaAnt=rs.fields("nreferencia")
rs.movenext
loop
%>

  </table>

  <BR>
  <INPUT TYPE="hidden" name="vendedor" Value="<% =vendedor %>">
  <INPUT TYPE="submit" value="Agregar" name="B1">

<%

end sub 'paso03()
'---------------------------------------------------------------------------------------------------------
sub buscavalor()
For Each elemento In request.form

	Response.Write("<BR>" & elemento & " : " & Request.form(elemento))

Next
end sub 'buscavalor()
'---------------------------------------------------------------------------------------------------------
sub rescatavalor()
For Each elemento In request.form

	if elemento <> "B1" then 
	Response.Write("<INPUT TYPE='hidden' name='" & elemento & "' Value='" & Request.form(elemento) & "'>")
	end if

Next
end sub 'buscavalor()
'---------------------------------------------------------------------------------------------------------
%>

</BODY>
<CENTER>
<FONT SIZE="1" COLOR="#C3C3C3">
Distribuci&oacute;n y Excelencia
</FONT>
</CENTER>
</HTML>

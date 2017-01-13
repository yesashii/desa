<html>
<%
vend 		= trim(request.querystring("vend"))
tipPed 		= trim(request.QueryString("tipPed"))
buscadnom 	= trim(request.form("buscar"))
':::::::::::::::::: conexion :::::::::::::::::
Private tipodoc, mitotal, oConn, rs, sql, rs1
Set oConn = server.createobject("ADODB.Connection")
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
'oConn.open "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=todo;User Id=sa;Password=desakey;"
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"

':::Busca por cliente
'anterior localhost
SQL1="SELECT CodLegal, RazonSocial, COUNT(CodLegal) 'Contar de CodLegal'" & _
"FROM todo.flexline.CtaCte CtaCte " & _
"WHERE (Empresa = 'DESA') AND (TipoCtaCte = 'CLIENTE') AND (Ejecutivo = '" & vend & "') AND (CodLegal + RazonSocial LIKE '%" & buscadnom & "%')" & _
"GROUP BY CodLegal, RazonSocial" '& _
'"ORDER BY RazonSocial"
response.write(tipPed)

'::: Busca por ruta
SQL2="SELECT *, rutflex as 'CodLegal', getdate() as fecha FROM Flexline.FX_RUTA_HOY1 " & _
"WHERE (Vendedor = '" & vend & "')"
':::: supercadoA
SQL3="SELECT Flexline.FX_RUTA_VENDEDOR_CLON.Destinatario, Flexline.CtaCte.CodLegal, Flexline.CtaCte.RazonSocial " & _
"FROM Flexline.FX_RUTA_VENDEDOR_CLON INNER JOIN Flexline.CtaCte ON Flexline.FX_RUTA_VENDEDOR_CLON.Origen = Flexline.CtaCte.Ejecutivo " & _
"WHERE (Flexline.CtaCte.Empresa = 'desa') AND (Flexline.CtaCte.TipoCtaCte = 'cliente') " & _
"GROUP BY Flexline.FX_RUTA_VENDEDOR_CLON.Destinatario, Flexline.CtaCte.CodLegal, Flexline.CtaCte.RazonSocial " & _
"HAVING (Flexline.FX_RUTA_VENDEDOR_CLON.Destinatario = N'CARLOS VALDEBENITO')"
'Detalle del listado
on error resume next

'Estado pedido Dia
nfecha=right(year(date),2) & right("00" & month(date),2) & right("00" & day(date),2)
SQLX="SELECT vendedor, estado, fecha, nota, Cliente " & _
"FROM Flexline.FX_PEDIDO_PDA " & _
"WHERE (fecha = N'" & nfecha & "') AND (vendedor = N'" & vend & "')"
Set rs1=oConn.execute(SQLX)

if len(buscadnom)=0 then
	Set rs=oConn.execute(SQL2)
	if not rs.eof then
		Midia="&nbsp;"
		if rs.fields("coddia")=1 then Midia="Lun"
		if rs.fields("coddia")=2 then Midia="Mar"
		if rs.fields("coddia")=3 then Midia="Mie"
		if rs.fields("coddia")=4 then Midia="Jue"
		if rs.fields("coddia")=5 then Midia="Vie"
		if rs.fields("coddia")=6 then Midia="Sab"
		if rs.fields("coddia")=7 then Midia="Dom"
		DetList= Midia & "-Semana:" & rs.fields("codsem")
	else
		Set rs=oConn.execute(SQL1)
		DetList= "todos"
	end if
Else
	Set rs=oConn.execute(SQL1)
	DetList= "busca:" & buscadnom
end if 
'response.write(SQL)

'77
if vend="CARLOS VALDEBENITO" then
	Set rs=oConn.execute(SQL3)
	DetList= "todos"
end if

%>
<head>
<title>Seleccionar Cliente</title>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-width: 0" bordercolor="#111111" width="100%">
  <tr>
    <td style="border-style: none; border-width: medium">
		<CENTER><B>
		<!-- <img src="../logo.JPG" width="32" height="20"> -->
			<font face="Arial" size="2" color="#333333">
				Sistema Preventa
			</font>
		</B></CENTER>
     </td>
  </tr>
  <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="center">
    <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
	  <tr>
	  	<td width="100%" align="right" valign="middle">
			<form method="post" action="">
				<FONT SIZE="2" COLOR="#808080">Busqueda</FONT>
				<input name="buscar" type="text" size="20">
	  			<input type="submit" name="buscanom" value="buscar">
			</form>
		</td>
	</tr>
	<tr>
        <td width="100%" align="center" valign="middle" bgcolor="#000080">
          <font face="Arial" size="2" color="#FFFFFF">
			List.&nbsp;Clientes&nbsp;(<%=DetList %>)
		  </font></td>
      </tr>
<%
'instr
'if len(trim(cstr(request.form("buscanom")))) then
'SQL1="SELECT CodLegal + RazonSocial AS busca FROM flexline.CtaCteCtaCte.CodLegal, CtaCte.RazonSocial where busca like'%" & trim(cstr(request.form("buscar"))) & "%'"
'end if
If rs.bof then response.write("<BR>La busqueda no produjo ninguna Concidencia")

do until rs.eof
	Estadoped=""

	if instr(rs.fields("CodLegal")," ")=0 then
		Milink="Pedido01.asp"
		LocalCliente=""
	else
		Milink="verifica_cliente.asp"
		LocalCliente="(" & mid(rs.fields("CodLegal"),instr(rs.fields("CodLegal")," ")+1) & ")"
		'busca estado
		Do until rs1.eof or rs1.bof
			If trim(rs.fields("CodLegal"))=trim(rs1.fields("Cliente"))then 
				if rs1.fields("estado")="noventa" then
					Estadoped="<B>NV</B>"
				else
					Estadoped="<B>PD</B>"
				end if 
			end if
		rs1.movenext
		loop
		rs1.MoveFirst
	end if
	
	if tipPed="N" then Milink="noventa/noventa.asp"

	CodLegal = rs.fields("CodLegal")
	response.write("<tr>")
	response.write("<td width='100%' align='center'><FONT face='Arial' SIZE='2'>&nbsp;" & _
	"<a href='" & Milink & "?vend=" & vend & "&cliente=" & CodLegal & "&tipPed=" & tipPed & "' style='text-decoration: none'>" & Estadoped & LocalCliente &  _
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
	<font face="Arial" size="1" color="#808080">
		&nbsp;NV=No Venta<BR>
		&nbsp;PD=Pedido<BR>
		&nbsp;* Mientras Menos letras pongas en el buscados, mas concidencias encuentras

	</font>
</body>
</html>
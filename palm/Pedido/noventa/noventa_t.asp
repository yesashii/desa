<html>
<head>
<title>traspaso no venta</title>
</head>
<%
on error resume next
dim opcion, txobs, cliente, vend
private oConn2, rs, sql, rs1
	Mifecha	=right(year(date),2) & right("00" & month(date),2) & right("00" & day(date),2)
	mihora	=right("00" & hour(time),2) & right("00" & minute(time),2) & right("00" & second(time),2)
	cliente =request.querystring("cliente")
	opcion  =trim(request.form("select"))
	txobs 	=trim(request.form("textarea"))
	vend 	=request.Querystring("vend")
	nuser=right(request.cookies("pdausr"),3)
	'response.write("<BR>" & nuser)
	if not isnumeric (nuser) then nuser=right("0" & request.cookies("pdausr"),3)
	'response.write("<BR>" & nuser)
	if not isnumeric (nuser) then nuser=1
	'response.write("<BR>" & nuser)
	IF LEN(opcion)=0 then opcion = left(request.querystring("opcion"),25)
	if isnumeric(vend) then
		
	end if

'================================================================='
'=========================Conexion servidor======================='
'================================================================='

	Set oConn2 = server.createobject("ADODB.Connection")

	oConn2.open "Provider=SQLOLEDB;Data Source=SQLSERVER; Initial Catalog=handheld; User Id=flexline; Password=corona;"

		sql="select * FROM handheld.flexline.FX_PEDIDO_PDA FX_PEDIDO_PDA" 
		
'================================================================='
'response.write len(vend)
	if len(vend)<4 then
		'response.write " menos de 4 "
		set rs=oConn2.execute("select top 1 nombre, idvendedor from handheld.dbo.Dim_vendedores where idvendedor=" & vend)
		if not rs.eof then vend=rs(0)
	'else
	'	set rs=oConn2.execute("select top 1 nombre, idvendedor from handheld.dbo.Dim_vendedores where idvendedor=" & vend)
	'	if not rs.eof then nuser=rs(1)
	end if

	set rs=oConn2.execute(sql)
		rs.close
	
rs.Open SQL, oConn2, 1, 3
	rs.AddNew
	rs.Fields("vendedor"            ) = vend
	rs.Fields("fecha"               ) = mifecha
	rs.Fields("hora"                ) = mihora
	rs.Fields("cliente"             ) = cliente
	rs.fields("obs"                 ) = txobs
	rs.fields("estado"              ) = "noventa"
	rs.fields("opcion"				) = opcion
	
	rs.Update
rs.Close

%>
<body text="#000000" style="font:arial" link="#000000" vlink="#000000">
<table width="100%" border="0" bordercolor="#000000">
  <tr>
    <td align="center"><strong>Registro No Venta</strong></td>
  </tr>
  <tr>
    <td align="center">---------------------------------------</td>
  </tr>
  <tr>  </tr>
  <tr>
    <td align="center"><strong>Opcion Seleccionada : </strong>
      <%Response.Write(opcion)%>	</td>
  </tr>
  <tr>
    <td align="center">---------------------------------------</td>
  </tr>
  <tr>      </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td align="center"><strong>Observaciones : </strong>
    <%response.Write(txobs)%></td>
  </tr>
  <tr>
    <td align="center">---------------------------------------</td>
  </tr>
  <tr>
    <td align="center"><FORM METHOD="POST" ACTION="../PEDIDO.ASP?nuser=<%=  nuser  %>">
	<INPUT TYPE="submit" VALUE="Aceptar">
	</FORM></td>
  </tr>
  <tr>
    <td align="center"></td>
  </tr>
</table>
</body>
</html>
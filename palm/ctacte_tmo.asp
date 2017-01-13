<CENTER>
<% 'on error resume next
PUBLIC tipodoc, mitotal, oConn, rs, sql
dim Mirut, nusuario, vendedor
':::::::::::::::::: conexion :::::::::::::::::
on error resume next

Set oConn = server.createobject("ADODB.Connection")
Set oConn2 = server.createobject("ADODB.Connection")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

oConn2.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"

nusuario=request.querystring("nuser"   )
vendedor=request.querystring("vendedor")
cliente =request.querystring("cliente" )
empresa =request.querystring("empresa" )

idempresa=1
if empresa="DESA"  then idempresa=1
if empresa="LACAV" then idempresa=4

if len(cliente)>0 then
	'response.write("cliente:" & cliente)
	sql="select ejecutivo from handheld.flexline.ctacte where ctacte like '" & cliente & "%' and empresa='" & empresa & "'"
	'response.write(sql)
	Set rs=oConn2.execute(Sql)
	if not rs.eof then
		vendedor=rs.fields(0)
		sql="select idvendedor from SQLSERVER.Desaerp.dbo.DIM_vendedores where nombre='" & vendedor & "' and idempresa=" & idempresa
		Set rs=oConn2.execute(Sql)
		nusuario=rs.fields(0)

		'response.write(cliente & "<br>" & vendedor & "<br>" & nusuario)
	end if
end if

if len(vendedor)=0 then vendedor=trim(cstr(request.form("vendedor")))
if len(empresa )=0 then empresa =trim(cstr(request.form("empresa" )))

response.write("<FONT SIZE='1' face='verdana'>empresa : " & empresa & "</FONT><BR>")

public N1, totdebe, tothaber, SaldoCliente
'response.write(trim(cstr(request.form("excel"))))

if trim(cstr(request.form("excel")))= "on" then Response.ContentType="application/vnd.ms-excel"
if request.querystring("excel")=1 then Response.ContentType="application/vnd.ms-excel"

call ctacte()
'response.write(nusuario)

'---------------------------------------------------------------------------------------------
sub ctacte() 
'on error resume next
if len(vendedor)=0 then 
	if len(nusuario)=3 then

	SQL="SELECT DESCRIPCION " & _
	"FROM GEN_TABCOD " & _
	"WHERE (EMPRESA = 'desa') AND (TIPO = 'GEN_VENDEDORES_PALM') AND (CODIGO = '" & right("000" & cint(nusuario),3) & "')"

	Set rs=oConn.execute(Sql)
		if not rs.eof then vendedor=rs.fields(0)
	end if
end if
nusuario=vendedor

sqlpagador="SELECT D.Nreferencia, D.Ejecutivo, D.AUX_VALOR6, D.AUX_VALOR3, D.RazonSocial, D.Direccion, D.Condpago, D.Giro, D.Correlativo, D.Tipo, D.Fecha, D.Glosa, D.Debe_Ingreso, D.Haber_ingreso, D.Referencia, D.FechaVcto " & _
"FROM BDFlexline.flexline.FX1_CtaCte_Deudas as D " & _
"WHERE ((D.rtmo='" & nusuario & "')) and D.AUX_VALOR3 like '%" & cliente & "%' and D.empresa='DESA' and D.Nreferencia not like '%888888' " & _
"ORDER BY D.AUX_VALOR3,D.NReferencia,D.Glosa, D.FechaVcto"


sqldistribuido="SELECT " & _
"D.Nreferencia, D.Ejecutivo, D.AUX_VALOR6, D.AUX_VALOR3, D.RazonSocial, D.Direccion, D.Condpago, D.Giro,  " & _
"D.Correlativo, D.Tipo, D.Fecha, D.Glosa, D.Debe_Ingreso, D.Haber_ingreso, D.Referencia, D.FechaVcto  " & _
"FROM BDFlexline.flexline.FX_Saldos_CtaCte as S " & _
"inner join BDFlexline.flexline.FX1_CtaCte_Deudas as D  " & _
"on S.Referencia = D.Referencia and s.empresa=d.empresa and s.codlegal=D.aux_valor3 " & _
"inner join ( " & _
"select codlegal " & _
"from BDFlexline.flexline.ctacte " & _
"where empresa='DESA' and tipoctacte='cliente' and Ejecutivo='" & nusuario & "' " & _
"group by codlegal " & _
") as C " & _
"on c.codlegal=D.aux_valor3 " & _
"WHERE D.AUX_VALOR3 like '%" & cliente & "%' and d.empresa='" & empresa & "' " & _
"ORDER BY D.AUX_VALOR3, D.FechaVcto, D.Referencia"

if nusuario="RODRIGO AVILA" then
	sql=sqldistribuido
else
	sql=sqlpagador
end if

'response.write(SQL)
'EXIT SUB

if nusuario="Nada" then 
	Set rs=oConn.execute("select 'ctacte sin datos', 'AUX_VALOR3'")
else
	Set rs=oConn.execute(sql)
'	response.write(cliente & "<br>" & vendedor & "<br>" & nusuario)
'response.write(sql)
end if

if rs.eof then
	'response.write("Cliente Sin Deuda<BR><BR><BR>")
	%>
	<CENTER>
	<FONT face="verdana" SIZE="2" COLOR="#000066">
	<BR><BR>Cliente Sin Deuda<BR><BR>
	<!-- <INPUT TYPE="button" onclick="history.back()" value="<< Atras"> -->
	</FONT>
	</CENTER>
	<%
	exit sub
end if

call titulo()
call sicliente()
dim ClienteAnt, Midebe, Mihaber


n1=2
totdebe = 0
tothaber = 0
clienteant=ucase(rs.fields("AUX_VALOR3"))
facturaant=rs.fields("Nreferencia")
Saldocliente = 0
do until rs.eof
	if ucase(rs.fields("AUX_VALOR3"))<>ucase(ClienteAnt) then call clientenuevo


Midebe = cdbl(rs.fields("DEBE_INGRESO"))
Mihaber = cdbl(rs.fields("HABER_INGRESO"))
totdebe = totdebe + Midebe
tothaber = tothaber + mihaber

if facturaant=rs.fields("Nreferencia") then
	SaldoCliente = SaldoCliente + Midebe - Mihaber
else
	SaldoCliente = Midebe - Mihaber
end if
response.write("<TR>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("CORRELATIVO") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("tipo") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("fecha") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & left(replace(rs.fields("glosa"),"-","&nbsp;"),27) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("DEBE_INGRESO"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("HABER_INGRESO"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(SaldoCliente,0) & "</FONT></TD>")

Numeroreferencia="<a style='FONT-WEIGHT: 700; TEXT-DECORATION: none' href='/palm/documento.asp?factura=" & rs.fields("Nreferencia") & "'" & _
"target='_blank'>" & right(rs.fields("Nreferencia"),7) & "</a>"

	response.write("<TD ALIGN='center' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & Numeroreferencia & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("FECHAVCTO") & "</FONT></TD>")
response.write("</TR>")

facturaant=rs.fields("Nreferencia")
clienteant=rs.fields("AUX_VALOR3")
rs.movenext
'salto pagina
	if n1>42 then
		call saltopag()
		call titulo()
		n1=1
	else
		n1=n1+1
	end if
loop

call fininfor()
if len(cliente)=0 then call ctacteporfolio()
end sub 'ctacte
'--------------------------------------------------------------------------------------------------------------
sub ctacteporfolio()
on error resume next
if len(vendedor)=0 then 
	if len(nusuario)=3 then

	SQL="SELECT DESCRIPCION " & _
	"FROM GEN_TABCOD " & _
	"WHERE (EMPRESA = 'desa') AND (TIPO = 'GEN_VENDEDORES_PALM') AND (CODIGO = '" & cint(nusuario) & "')"

	Set rs=oConn.execute(Sql)
		if not rs.eof then vendedor=rs.fields(0)
	end if
end if
nusuario=vendedor

SQL="SELECT D.Nreferencia, D.Ejecutivo, D.AUX_VALOR6, D.AUX_VALOR3, D.RazonSocial, D.Direccion, D.Condpago, D.Giro, D.Correlativo, D.Tipo, D.Fecha, D.Glosa, D.Debe_Ingreso, D.Haber_ingreso, D.Referencia, D.FechaVcto " & _
"FROM BDFlexline.flexline.FX_Saldos_CtaCte as S, BDFlexline.flexline.FX1_CtaCte_Deudas as D " & _
"WHERE S.Referencia = D.Referencia " & _
"AND ((D.Ejecutivo='" & nusuario & "')) and D.AUX_VALOR3 like '%" & cliente & "%' and D.Nreferencia not like '%888888' " & _
"ORDER BY D.AUX_VALOR3, D.Referencia"

SQL="SELECT D.Nreferencia, D.Ejecutivo, D.AUX_VALOR6, D.AUX_VALOR3, D.RazonSocial, D.Direccion, D.Condpago, D.Giro, D.Correlativo, D.Tipo, D.Fecha, D.Glosa, D.Debe_Ingreso, D.Haber_ingreso, D.Referencia, D.FechaVcto " & _
"FROM serverdesa.BDFlexline.flexline.FX_Saldos_CtaCte as S inner join " & _
"	serverdesa.BDFlexline.flexline.FX1_CtaCte_Deudas as D on S.Referencia = D.Referencia and d.empresa=d.empresa inner join " & _
"	handheld.Flexline.CtaCte as p on D.AUX_VALOR3 =left(rtrim(ltrim(p.ctacte)),len(ltrim(rtrim(D.AUX_VALOR3)))) " & _
"WHERE porfolio2='" & nusuario & "' and d.empresa='" & empresa & "' and D.Nreferencia not like '%888888' " & _
"group by D.Nreferencia, D.Ejecutivo, D.AUX_VALOR6, D.AUX_VALOR3, D.RazonSocial, D.Direccion, D.Condpago, D.Giro, D.Correlativo, D.Tipo, D.Fecha, D.Glosa, D.Debe_Ingreso, D.Haber_ingreso, D.Referencia, D.FechaVcto " & _
"ORDER BY D.AUX_VALOR3, D.Referencia " 

'response.write(SQL)
'EXIT SUB

if nusuario="Nada" then 
	Set rs=oConn2.execute("select 'ctacte sin datos', 'AUX_VALOR3'")
else
	Set rs=oConn2.execute(Sql)
end if

if rs.eof then
	'response.write("Cliente Sin Deuda<BR><BR><BR>")
	%>
	<!-- <CENTER>
	<FONT face="verdana" SIZE="2" COLOR="#000066">
	<BR><BR>Cliente Sin Deuda<BR><BR>
	 <INPUT TYPE="button" onclick="history.back()" value="<< Atras">
	</FONT>
	</CENTER> -->
	<%
	exit sub
end if

call titulo()
call sicliente()
dim ClienteAnt, Midebe, Mihaber


n1=2
totdebe = 0
tothaber = 0
clienteant=ucase(rs.fields("AUX_VALOR3"))
facturaant=rs.fields("Nreferencia")
Saldocliente = 0
do until rs.eof
	if ucase(rs.fields("AUX_VALOR3"))<>ucase(ClienteAnt) then call clientenuevo


Midebe = cdbl(rs.fields("DEBE_INGRESO"))
Mihaber = cdbl(rs.fields("HABER_INGRESO"))
totdebe = totdebe + Midebe
tothaber = tothaber + mihaber

if facturaant=rs.fields("Nreferencia") then
SaldoCliente = SaldoCliente + Midebe - Mihaber
else
SaldoCliente = Midebe - Mihaber
end if
response.write("<TR>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("CORRELATIVO") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("tipo") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("fecha") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & left(replace(rs.fields("glosa"),"-","&nbsp;"),27) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("DEBE_INGRESO"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("HABER_INGRESO"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(SaldoCliente,0) & "</FONT></TD>")

Numeroreferencia="<a style='FONT-WEIGHT: 700; TEXT-DECORATION: none' href='/palm/documento.asp?factura=" & rs.fields("Nreferencia") & "'" & _
"target='_blank'>" & right(rs.fields("Nreferencia"),7) & "</a>"

	response.write("<TD ALIGN='center' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & Numeroreferencia & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("FECHAVCTO") & "</FONT></TD>")
response.write("</TR>")

facturaant=rs.fields("Nreferencia")
clienteant=rs.fields("AUX_VALOR3")
rs.movenext
'salto pagina
	if n1>42 then
		call saltopag()
		call titulo()
		n1=1
	else
		n1=n1+1
	end if
loop

call fininfor()
end sub 'ctacteporfolio
'--------------------------------------------------------------------------------------------------------------
sub titulo()
	%>
	<U><B><FONT SIZE="2" COLOR="#000099" FACE="Arial" >INFORME DE CUENTAS CORRIENTES</FONT></B></U>
	<BR><FONT FACE="Arial" SIZE="1" COLOR="#5F5F5F">Vendedor: <%=nusuario %></FONT>
	<HR>
	<TABLE>
	<TR>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Numero</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Tipo</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>&nbsp;&nbsp;Fecha&nbsp;&nbsp;&nbsp;</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Glosa</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Debe</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Haber</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Saldo</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Docum</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Vencimi&nbsp;</B></CENTER></FONT></TD>
	</TR>
	<%
end sub 'titulo()
'--------------------------------------------------------------------------------------------------------------
sub saltopag()
%>
</TABLE>
<div style="page-break-before:always"></div>
<%
end sub 'saltopag()

'--------------------------------------------------------------------------------------------------------------
sub sicliente()
'on error resume next
if nusuario="Nada" then exit sub 
if rs.eof then exit sub
%>
<TR>
	<TD></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'>Rut</FONT></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'><% =ucase(rs.fields("AUX_VALOR3")) %></FONT></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'><% =left(rs.fields("RazonSocial"),28) %></FONT></TD>
	<TD></TD>
	<TD></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'><% =left(rs.fields("CondPago"),15) %></FONT></TD>
	<TD></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
</TR>
<%
end sub 'sicliente()
'--------------------------------------------------------------------------------------------------------------
sub Clientenuevo()
%>

<TR>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium" ><FONT FACE='Arial' SIZE=1 COLOR=''><CENTER><B>Total Cliente ------></B></CENTER></FONT></TD>
	<TD ALIGN="right" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium" ><FONT FACE='Arial' SIZE=1 COLOR=''><B><% =FormatNumber(totdebe,0) %></B></FONT></TD>
	<TD ALIGN="right" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium" ><FONT FACE='Arial' SIZE=1 COLOR=''><B><% =FormatNumber(tothaber,0) %></B></FONT></TD>
	<TD ALIGN="right" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium" ><FONT FACE='Arial' SIZE=1 COLOR=''><B><% =FormatNumber((totdebe-tothaber),0) %></B></FONT></TD>
	<TD></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
</TR>
<TR>
	<TD>____</TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
</TR>
<%
totdebe = 0
tothaber = 0
SaldoCliente = 0
n1 = n1 + 3 'mas lines
call sicliente()
end sub 'Clientenuevo()
'--------------------------------------------------------------------------------------------------------------
sub fininfor()
%>

<TR>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR=''><B>Total Cliente</B></FONT></TD>
	<TD ALIGN="right" ><FONT FACE='Arial' SIZE=1 COLOR=''><B><% =totdebe %></B></FONT></TD>
	<TD ALIGN="right" ><FONT FACE='Arial' SIZE=1 COLOR=''><B><% =tothaber %></B></FONT></TD>
	<TD ALIGN="right" ><FONT FACE='Arial' SIZE=1 COLOR=''><B><% =(totdebe-tothaber) %></B></FONT></TD>
	<TD></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
	<TD></TD>
</TR>
</TABLE>
<%
totdebe = 0
tothaber = 0
SaldoCliente = 0
n1 = n1 + 2 'mas lines
'call sicliente()
end sub 'fininfor()


'rs.close
'oConn.close
'Set rs=Nothing
'Set oConn = Nothing

%>
</CENTER>
<CENTER>
<% 'on error resume next
PUBLIC tipodoc, mitotal, oConn, rs, sql
dim Mirut, nusuario, vendedor
':::::::::::::::::: conexion :::::::::::::::::
'on error resume next

Set oConn = server.createobject("ADODB.Connection")
Set oConn2 = server.createobject("ADODB.Connection")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=bdflexline;User Id=sa;Password=desakey;"

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

	sql="select nombre from SQLSERVER.Desaerp.dbo.DIM_vendedores where idvendedor='" & nusuario & "' and idempresa=" & idempresa
	
	Set rs=oConn.execute(Sql)
		if not rs.eof then vendedor=rs.fields(0)
	end if
end if

'response.write(sql)

sqlpagador= "SELECT * FROM ( " & _
			"SELECT T0.REFERENCIA as NUM_DOCUMENTO,T1.Ejecutivo AS VENDEDOR, " & _
			"T0.AUX_VALOR3 AS RUT,T1.RazonSocial AS CLIENTE,T1.CondPago AS COND_PAGO_FACT,  " & _
			"T0.REFERENCIA AS TRANSACCION, 'FACTURA' AS DOCUMENTO, " & _
			"(Select CONVERT(VARCHAR(10),T2.Fecha,105) FROM [BDFlexline].[flexline].[Documento] AS T2 Where T2.Empresa='LACAV' and  " & _
			"T2.TipoDocto='FACT. AFECTA ELEC' and T2.Numero=RIGHT(('0000000000'+T0.REFERENCIA),10) ) as FECHA_CONTAB,  " & _
			"('FACT. ' +  T0.REFERENCIA + ' ' +T0.EMPRESA) AS GLOSA,  " & _
			"case when ( SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO)>0) THEN  SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO) else 0 end as debe, " & _
			"case when ( SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO)<0) THEN  -1*SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO) else 0 end as haber, " & _
			"(SELECT CONVERT(VARCHAR(10),T4.FechaVcto,105) FROM [BDFlexline].[flexline].[Documento] AS T3  " & _
			"INNER JOIN [BDFlexline].[flexline].[DocumentoP] T4 ON T3.Empresa=T4.Empresa AND T3.TipoDocto=T4.TipoDocto and T3.Correlativo=T4.Correlativo  " & _
			"WHERE T3.Empresa='LACAV' and T3.TipoDocto='FACT. AFECTA ELEC' and T3.Numero=RIGHT(('0000000000'+T0.REFERENCIA),10) ) AS FECHA_VENC " & _
		    "FROM SERVERDESA.BDFlexline.flexline.CON_MOVCOM AS T0 " & _
		    "INNER JOIN [BDFlexline].[flexline].[CtaCte] AS T1 " & _
		    "ON T0.AUX_VALOR3+' 1'=T1.CtaCte and T0.EMPRESA=T1.Empresa AND T1.TipoCtaCte='CLIENTE' " & _
		    "WHERE (T0.EMPRESA = 'LACAV') AND (T0.PERIODO IN (2016, 2017) ) AND (T0.CUENTA = '010103010000')  " & _
		    "AND (T0.AUX_VALOR3 like '%" & cliente & "%') AND (T0.ESTADO = 'A') AND T1.Ejecutivo='" & vendedor & "' " & _
		    "GROUP BY T0.EMPRESA, T0.PERIODO, T0.CUENTA, T0.AUX_VALOR3, T0.ESTADO, T0.REFERENCIA,  " & _
		    "T1.Ejecutivo,T1.RazonSocial,T1.CondPago " & _
		    "HAVING (Sum(T0.DEBE_INGRESO - T0.HABER_INGRESO) <> 0) " & _
		    "UNION " & _
		    "SELECT D.NUM_DOCUMENTO, D.VENDEDOR, D.RUT, D.CLIENTE, D.COND_PAGO_FACT, D.TRANSACCION, D.DOCUMENTO, D.FECHA_CONTAB, D.GLOSA,  " & _
		    "case when (D.[SALDO_PENDIENTE]>0) THEN D.[SALDO_PENDIENTE] else 0 end as debe, " & _
		    "case when (D.[SALDO_PENDIENTE]<0) THEN -1*D.[SALDO_PENDIENTE] else 0 end as haber,  " & _
		    "D.FECHA_VENC  " & _
		    "FROM SQLSERVER.[DataWarehouse].[dbo].[Deuda_SAP] as D  " & _
		    "WHERE ((D.VENDEDOR='" & vendedor & "')) and D.RUT like '%" & cliente & "%' and D.TRANSACCION not like '%888888'  " & _
		    ") AS T99 " & _
		    "ORDER BY T99.RUT,T99.TRANSACCION,T99.GLOSA,T99.FECHA_VENC"

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
	Set rs=oConn.execute("select 'ctacte sin datos', 'RUT'")
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
clienteant=ucase(rs.fields("RUT"))
facturaant=rs.fields("NUM_DOCUMENTO")
Saldocliente = 0
do until rs.eof
	if ucase(rs.fields("RUT"))<>ucase(ClienteAnt) then call clientenuevo


Midebe = cdbl(rs.fields("DEBE"))
Mihaber = cdbl(rs.fields("HABER"))
totdebe = totdebe + Midebe
tothaber = tothaber + mihaber

if facturaant=rs.fields("NUM_DOCUMENTO") then
	SaldoCliente = SaldoCliente + Midebe - Mihaber
else
	SaldoCliente = Midebe - Mihaber
end if
response.write("<TR>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("TRANSACCION") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("DOCUMENTO") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("FECHA_CONTAB") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & left(replace(rs.fields("GLOSA"),"-","&nbsp;"),27) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("DEBE"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("HABER"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(SaldoCliente,0) & "</FONT></TD>")

	if RIGHT(rs.fields("GLOSA"),5)="LACAV" THEN
		linkemp="LACAV"
	ELSE
		linkemp="DESA"
	END IF
	
	if rs.fields("DOCUMENTO")="FACTURA" then
		Numeroreferencia="<a style='FONT-WEIGHT: 700; TEXT-DECORATION: none' href='/palm/documento.asp?factura=" & rs.fields("NUM_DOCUMENTO") & "&empresa="& linkemp &"'" & _
		"target='_blank'>" & right(rs.fields("NUM_DOCUMENTO"),7) & "</a>"
		response.write("<TD ALIGN='center' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & Numeroreferencia & "</FONT></TD>")
	else
		response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("NUM_DOCUMENTO") & "</FONT></TD>")
	end if
	
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("FECHA_VENC") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("COND_PAGO_FACT") & "</FONT></TD>")
response.write("</TR>")

facturaant=rs.fields("NUM_DOCUMENTO")
clienteant=rs.fields("RUT")
rs.movenext
'salto pagina
	if n1>42 then
		call saltopag()
		call titulo()
		n1=1
	else
		n1=n1+1
	end if
response.Flush()
loop

call fininfor()
'if len(cliente)=0 then call ctacteporfolio()
end sub 'ctacte
'--------------------------------------------------------------------------------------------------------------
sub ctacteporfolio()
on error resume next
if len(vendedor)=0 then 
	if len(nusuario)=3 then

	sql="select nombre from SQLSERVER.Desaerp.dbo.DIM_vendedores where idvendedor='" & nusuario & "' and idempresa=" & idempresa

	Set rs=oConn.execute(Sql)
		if not rs.eof then vendedor=rs.fields(0)
	end if
end if

SQL="SELECT * FROM ( " & _
			"SELECT T0.REFERENCIA as NUM_DOCUMENTO,T1.Ejecutivo AS VENDEDOR, " & _
			"T0.AUX_VALOR3 AS RUT,T1.RazonSocial AS CLIENTE,T1.CondPago AS COND_PAGO_FACT,  " & _
			"T0.REFERENCIA AS TRANSACCION, 'FACTURA' AS DOCUMENTO, " & _
			"(Select CONVERT(VARCHAR(10),T2.Fecha,105) FROM [BDFlexline].[flexline].[Documento] AS T2 Where T2.Empresa='LACAV' and  " & _
			"T2.TipoDocto='FACT. AFECTA ELEC' and T2.Numero=RIGHT(('0000000000'+T0.REFERENCIA),10) ) as FECHA_CONTAB,  " & _
			"('FACT. ' +  T0.REFERENCIA + ' ' +T0.EMPRESA) AS GLOSA,  " & _
			"case when ( SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO)>0) THEN  SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO) else 0 end as debe, " & _
			"case when ( SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO)<0) THEN  -1*SUM(T0.DEBE_INGRESO - T0.HABER_INGRESO) else 0 end as haber, " & _
			"(SELECT CONVERT(VARCHAR(10),T4.FechaVcto,105) FROM [BDFlexline].[flexline].[Documento] AS T3  " & _
			"INNER JOIN [BDFlexline].[flexline].[DocumentoP] T4 ON T3.Empresa=T4.Empresa AND T3.TipoDocto=T4.TipoDocto and T3.Correlativo=T4.Correlativo  " & _
			"WHERE T3.Empresa='LACAV' and T3.TipoDocto='FACT. AFECTA ELEC' and T3.Numero=RIGHT(('0000000000'+T0.REFERENCIA),10) ) AS FECHA_VENC " & _
		    "FROM SERVERDESA.BDFlexline.flexline.CON_MOVCOM AS T0 " & _
		    "INNER JOIN [BDFlexline].[flexline].[CtaCte] AS T1 " & _
		    "ON T0.AUX_VALOR3+' 1'=T1.CtaCte and T0.EMPRESA=T1.Empresa AND T1.TipoCtaCte='CLIENTE' " & _
		    "WHERE (T0.EMPRESA = 'LACAV') AND (T0.PERIODO IN (2016, 2017) ) AND (T0.CUENTA = '010103010000')  " & _
		    "AND (T0.AUX_VALOR3 like '%" & cliente & "%') AND (T0.ESTADO = 'A') AND T1.Ejecutivo='" & vendedor & "' " & _
		    "GROUP BY T0.EMPRESA, T0.PERIODO, T0.CUENTA, T0.AUX_VALOR3, T0.ESTADO, T0.REFERENCIA,  " & _
		    "T1.Ejecutivo,T1.RazonSocial,T1.CondPago " & _
		    "HAVING (Sum(T0.DEBE_INGRESO - T0.HABER_INGRESO) <> 0) " & _
		    "UNION " & _
		    "SELECT D.NUM_DOCUMENTO, D.VENDEDOR, D.RUT, D.CLIENTE, D.COND_PAGO_FACT, D.TRANSACCION, D.DOCUMENTO, D.FECHA_CONTAB, D.GLOSA,  " & _
		    "case when (D.[SALDO_PENDIENTE]>0) THEN D.[SALDO_PENDIENTE] else 0 end as debe, " & _
		    "case when (D.[SALDO_PENDIENTE]<0) THEN -1*D.[SALDO_PENDIENTE] else 0 end as haber,  " & _
		    "D.FECHA_VENC  " & _
		    "FROM SQLSERVER.[DataWarehouse].[dbo].[Deuda_SAP] as D  " & _
		    "WHERE ((D.VENDEDOR='" & vendedor & "')) and D.RUT like '%" & cliente & "%' and D.TRANSACCION not like '%888888'  " & _
		    ") AS T99 " & _
		    "ORDER BY T99.RUT,T99.TRANSACCION,T99.GLOSA,T99.FECHA_VENC"

SQL="SELECT D.NUM_DOCUMENTO, D.VENDEDOR, D.VENDEDOR, D.RUT, D.CLIENTE, D.COND_PAGO_FACT, D.TRANSACCION, D.DOCUMENTO, D.FECHA_CONTAB, D.GLOSA, case when (D.[SALDO_PENDIENTE]>0) THEN D.[SALDO_PENDIENTE] else 0 end as debe,case when (D.[SALDO_PENDIENTE]<0) THEN -1*D.[SALDO_PENDIENTE] else 0 end as haber, D.TRANSACCION, D.FECHA_VENC " & _
"FROM SQLSERVER.[DataWarehouse].[dbo].[Deuda_SAP] as D " & _
"WHERE ((D.VENDEDOR='" & nusuario & "')) and D.RUT like '%" & cliente & "%' and D.TRANSACCION not like '%888888' " & _
"ORDER BY D.RUT,D.TRANSACCION,D.GLOSA, D.FECHA_VENC"

response.write(SQL)
'EXIT SUB

if nusuario="Nada" then 
	Set rs=oConn2.execute("select 'ctacte sin datos', 'RUT'")
else
	Set rs=oConn2.execute(Sql)
end if

if rs.eof then
'response.write(sql)
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
clienteant=ucase(rs.fields("RUT"))
facturaant=rs.fields("NUM_DOCUMENTO")
Saldocliente = 0
do until rs.eof
	if ucase(rs.fields("RUT"))<>ucase(ClienteAnt) then call clientenuevo


Midebe = cdbl(rs.fields("DEBE"))
Mihaber = cdbl(rs.fields("HABER"))
totdebe = totdebe + Midebe
tothaber = tothaber + mihaber

if facturaant=rs.fields("NUM_DOCUMENTO") then
SaldoCliente = SaldoCliente + Midebe - Mihaber
else
SaldoCliente = Midebe - Mihaber
end if
response.write("<TR>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("TRANSACCION") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("DOCUMENTO") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("FECHA_CONTAB") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & left(replace(rs.fields("GLOSA"),"-","&nbsp;"),27) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("DEBE"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(rs.fields("HABER"),0) & "</FONT></TD>")
	response.write("<TD ALIGN='right' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & FormatNumber(SaldoCliente,0) & "</FONT></TD>")

	if RIGHT(rs.fields("GLOSA"),5)="LACAV" THEN
		linkemp="LACAV"
	ELSE
		linkemp="DESA"
	END IF
	
if rs.fields("DOCUMENTO")="FACTURA" then
	Numeroreferencia="<a style='FONT-WEIGHT: 700; TEXT-DECORATION: none' href='/palm/documento.asp?factura=" & rs.fields("NUM_DOCUMENTO") & "&empresa=" & linkemp & "'" & _
	"target='_blank'>" & right(rs.fields("NUM_DOCUMENTO"),7) & "</a>"
	response.write("<TD ALIGN='center' ><FONT FACE='Arial' SIZE=1 COLOR=''>" & Numeroreferencia & "</FONT></TD>")
else
    response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("NUM_DOCUMENTO") & "</FONT></TD>")
end if

	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("FECHA_VENC") & "</FONT></TD>")
	response.write("<TD><FONT FACE='Arial' SIZE=1 COLOR=''>" & rs.fields("COND_PAGO_FACT") & "</FONT></TD>")
response.write("</TR>")

facturaant=rs.fields("NUM_DOCUMENTO")
clienteant=rs.fields("RUT")
rs.movenext
'salto pagina
	if n1>42 then
		call saltopag()
		call titulo()
		n1=1
	else
		n1=n1+1
	end if
response.Flush()
loop

call fininfor()
end sub 'ctacteporfolio
'--------------------------------------------------------------------------------------------------------------
sub titulo()
	%>
	<U><B><FONT SIZE="2" COLOR="#000099" FACE="Arial" >INFORME DE CUENTAS CORRIENTES</FONT></B></U>
	<BR><FONT FACE="Arial" SIZE="1" COLOR="#5F5F5F">Vendedor: <%=vendedor %></FONT>
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
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Vencimi</B></CENTER></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#000099" FACE="Arial" ><CENTER><B>Cond. Factura&nbsp;</B></CENTER></FONT></TD>
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
'response.write(SQL)

SQL3=" SELECT top 1 [CondPago] FROM [BDFlexline].[flexline].[ctacte] " & _
"Where CodLegal like '%" & ucase(rs.fields("RUT")) & "%' and Empresa='DESA' and tipoctacte='CLIENTE' Order by ctacte "

'response.write(SQL3)

Set rscondpago=oConn.execute(Sql3)


%>
<TR>
	<TD></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'>Rut</FONT></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'><% =ucase(rs.fields("RUT")) %></FONT></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'><% =left(rs.fields("CLIENTE"),28) %></FONT></TD>
	<TD></TD>
	<TD></TD>
	<TD><FONT FACE='Arial' SIZE=1 COLOR='#000099'><% =(rscondpago.fields("CondPago")) %></FONT></TD>
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
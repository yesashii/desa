<body topmargin="0" >
<%
'------------------------------------------------------------------------------------------------
On error resume next
dim oConn, oConn2, rs, rs2, sql, sql2, SQL1, rs1, rs0, SQL0 
dim vend, id_porfolio

vend	 = trim(request.QueryString("vend"))
'vend     = replace(vend,"NN","�")'Error blackberry
cliente  = trim(request.QueryString("cliente"))
nuser	 = right("0" & request.cookies("pdausr"),3)
empresa  = trim(request.QueryString("empresa"))
sw_iva	 = request.QueryString("sw_iva")
tipodocto= request.QueryString("tipodocto")




if request.form("id_porfolio")="DESA"   then id_porfolio="porfolio1"
if request.form("id_porfolio")="LA CAV" then id_porfolio="porfolio2"


if len(id_porfolio)<1 then id_porfolio=request.form("id_porfolio"   )
if len(id_porfolio)<1 then id_porfolio=request.querystring("id_porfolio")

'response.write "ID_porfolio : " & id_porfolio & "<BR>"
'testCheck = "BAT123khg" Like "B?T*"

'response.write instr(ucase(vend),"CALL") 'testCheck

'response.write "ID_porfolio : " & id_porfolio
':: coneccion ::
set oConn = server.CreateObject("ADODB.connection")
oConn.open "provider=SQLOLEDB; Data Source=SQLSERVER; initial catalog=handheld; user id=sa; password=desakey;"





':: valida Ruta call
if instr("CALL",ucase(vend))=0 then
	'response.write "Es vendedor"
	'IdDia = datepart("w", date-1)
	'response.write IdDia
	sql="SELECT * FROM handheld.Flexline.PDA_RUTA_HOY WHERE (RutFlex = '" & cliente & "') and vendedor like '%CALL%'"
	set rs=oConn.execute(sql)
	if not rs.eof then
		%><BR>
		<CENTER>
		<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-width: 0">
		<TR bgcolor="#666633">
			<TD>.::</TD>
			<TD align="center"><FONT SIZE="2" face="verdana" COLOR="#FFFFFF"><B>Cliente en ruta call Center</B></FONT></TD>
			<TD>::.</TD>
		</TR>
		<TR  bgcolor="#C0C0C0">
			<TD></TD>
			<TD>&nbsp;</TD>
			<TD></TD>
		</TR>
		<TR bgcolor="#C0C0C0">
			<TD></TD>
			<TD><FONT SIZE="2" face="verdana" COLOR="#000033">Advertencia <BR>Hoy este cliente sera llamado por <%=rs("vendedor")%></FONT></TD>
			<TD></TD>
		</TR>
		<TR  bgcolor="#C0C0C0">
			<TD></TD>
			<TD>&nbsp;</TD>
			<TD></TD>
		</TR>
		</TABLE>
		</CENTER>
		<%
	end if
end if


':: valida Porfolio ::
accion=1
If len(id_porfolio)<1 Then
	sql="select porfolio1, porfolio2 from handheld.flexline.ctacte where empresa='" & empresa & "' and ctacte = '" & cliente & "'"
	set rs1=oConn.execute(sql)
	if rs1.eof then
		response.write("EOF")
	else
		'response.write("<BR>porfolio1 : " & rs1.fields(0))
		'response.write("<BR>porfolio2 : " & rs1.fields(1))
		'response.write("<BR> V E N D .: " & vend)
		'response.write("<BR>rs.fields:porfolio2" & rs1.fields("porfolio2") )

		if len(rs1.fields(0))>1 then
			if len(rs1.fields(1))>1 then
				if rs1.fields(1)<>"NO ASIGNADO" then accion=0
			end if
		end if
		'response.write("<BR>ACCION : " & accion)
		
		if ucase(rs1.fields("porfolio1"))=ucase(vend) then
			accion = 1
			id_porfolio="porfolio1"
			if len(trim(rs1.fields("porfolio2")))=0  then id_porfolio="" 'si solo es del vendedor desa
			if isnull(rs1.fields("porfolio2"))       then id_porfolio="" 'si solo es del vendedor desa
			if rs1.fields("porfolio2")="NO ASIGNADO" then id_porfolio=""
			
			response.write("-ID porfolio : " & id_porfolio)

		'	response.write("<BR>id_porfolio : " & id_porfolio )
			
		end if
		'response.write("<BR>ACCION : " & accion)
	end if
End If
id_porfolio="" 'anula asignacion porfolios

	'if ucase(vend)="PEDRO MORALES"        then id_porfolio="porfolio2"
	'if ucase(vend)="CRISTIAN VALDEBENITO" then id_porfolio="porfolio2"
	'if ucase(vend)="MARIO JARA"           then id_porfolio="porfolio2"



': Bloqueo cliente
if len(cliente)>7 then
	msgbloqueo=""
	varsplit=split(cliente," ")
	codlegal=varsplit(0)
	sql="select top 1 comentario1 from Flexline.CtaCte where empresa='" & trim(empresa) & "' and codlegal='" & trim(codlegal) &"' and vigencia='B' order by len(comentario1) desc"
	'response.write(sql)
	Set rs=oConn.execute(Sql)
	do until rs.eof
		msgbloqueo="Comentario : " & rs(0)
	rs.movenext
	loop
	if len(msgbloqueo) >0 then
		%><CENTER><BR><BR><B>Cliente Bloqueado</B>
		<BR>
		<BR>
		<%=msgbloqueo%>
		<BR>
		<BR>
		<INPUT TYPE="button" value="<< Volver" onClick="history.back();history.back()" >
		</CENTER><%
		'response.write(msgbloqueo)
		'accion = 999
	end if 
end if




if accion = 1 then call verificacliente()
if accion = 0 then call selectporfolio()

'----------------------------------------------------------------------------------
function selectporfolio()
%>
	<CENTER>
	<FORM METHOD=POST ACTION="">
	<TABLE>
	<TR>
		<TD colspan="2">&nbsp;</TD>
	</TR>
	<TR>
		<TD colspan="2"><B><FONT SIZE="2" face="arial" COLOR="#666633">Seleccion Empresa</FONT></B></TD>
	</TR>
	<TR>
		<TD>
		<INPUT TYPE="submit" value="DESA" name="id_porfolio"><!-- Porfolio1 -->
		</TD>
		<TD>
		<INPUT TYPE="submit" value="LA CAV" name="id_porfolio"><!-- Porfolio2 -->
		</TD>
	</TR>
	</TABLE>
	</FORM>
	</CENTER>
<%
end function 'selectporfolio()
'----------------------------------------------------------------------------------

function validarut(rut)
validarut="ok"
if instr(rut," ")<>0 then validarut="mal"
if instr(rut,"-") =0 then validarut="mal"
if instr(rut,".")<>0 then validarut="mal"
ruts=split(lcase(rut),"-")
if ubound( ruts )<>1 then validarut="mal"
if validarut="mal" then exit function
rut=ruts(0)
dig=ruts(1)
'rut= trim(request.form("rut"))
'dig=trim(request.form("dig"))
tur=strreverse(rut) 
mult = 2 

for i = 1 to len(tur) 
   if mult > 7 then 
      mult = 2 
   end if 
   suma = mult * mid(tur,i,1) + suma 
   mult = mult +1 
next 

valor = 11 - (suma mod 11) 

if valor = 11 then 
   codigo_veri = "0" 
elseif valor = 10 then 
   codigo_veri = "k" 
else 
   codigo_veri = valor 
end if

if Cstr(dig)=Cstr(codigo_veri) then
   validarut="ok"
else
   validarut="mal"
end if
end function 'validarut()

'----------------------------------------------------------------------------------
function verificacliente()
on error resume next

sql="SELECT C.CodLegal, C.RazonSocial, C.Sigla, C.EMAIL, "&_
	"C.Contacto, D.Direccion, C.CtaCte, C.Giro, C.CondPago, "&_
	"C.Comuna, C.Telefono, C.LimiteCredito, C.PorcDr1 "&_
	"FROM handheld.flexline.CtaCte C, handheld.flexline.CtaCteDirecciones D "&_
	"WHERE D.CtaCte = C.CtaCte AND D.Empresa = C.Empresa  "&_
	"AND D.TipoCtaCte = C.TipoCtaCte AND ((C.Empresa='DESA')  "&_
	"AND (C.TipoCtaCte='CLIENTE') AND (C.Ejecutivo='" & vend & "')  "&_
	"AND (D.CtaCte='" & cliente & "') AND (D.Principal<>'s'))"
if ucase(empresa)="DESAZOFRI" then sql=replace(sql,"'DESA'","'DESAZOFRI'")
set rs=oConn.execute(sql)
'response.write sql
if rs.eof then
	'response.write("NX")
	'on error resume next
sql="SELECT C.CodLegal, C.RazonSocial, D.Direccion, C.CtaCte, C.Giro, "&_
	"C.CondPago, C.Comuna, C.Telefono, C.LimiteCredito, C.PorcDr1, C.vigencia "&_
	"FROM flexline.CtaCte C INNER JOIN flexline.CtaCteDirecciones D  "&_
	"ON C.CtaCte = D.CtaCte AND C.Empresa = D.Empresa  "&_
	"AND C.TipoCtaCte = D.TipoCtaCte "&_
	"WHERE (C.Empresa = 'DESA') AND (C.TipoCtaCte = 'CLIENTE')  "&_
	"AND (D.CtaCte = '" & cliente & "') AND (D.Principal <> 's')"

	set rs=oConn.execute(sql)
end if
rutok = validarut(rs.fields("codlegal"))
if len(vend)=0 then rutok="x"
'Mnto. A Vencer
'"FROM serverdesa.BDFlexline.Flexline.FX1_CtaCte_Deudas Deudas "&_'
sql1="SELECT AUX_VALOR3, SUM(Debe_Ingreso - Haber_ingreso) AS SALDO "&_
     "FROM [SQLSERVER].HandHeld.Flexline.FX1_CtaCte_Deudas Deudas "&_
     "WHERE (FechaVcto >= GETDATE()) GROUP BY AUX_VALOR3 "&_
     "HAVING (AUX_VALOR3 = '" & rs.fields("codlegal") & "')"
'sql1="select '' as AUX_VALOR3, 0 AS SALDO"
'Mnto Vencido
'"FROM serverdesa.BDFlexline.Flexline.FX1_CtaCte_Deudas Deudas "&_'
sql2="SELECT AUX_VALOR3, SUM(Debe_Ingreso - Haber_ingreso) AS SALDO "&_
     "FROM [SQLSERVER].HandHeld.Flexline.FX1_CtaCte_Deudas Deudas "&_
     "WHERE (FechaVcto <= GETDATE()) GROUP BY AUX_VALOR3 "&_
     "HAVING (AUX_VALOR3 = '" & rs.fields("codlegal") & "')"
'sql2="select '' as AUX_VALOR3, 0 AS SALDO"

'saldos CH cartera
'"FROM serverdesa.BDFlexline.dbo.[el_SALDOS CHEQUES SANTIAGO] "&_'
'SQL0="SELECT AUX_VALOR3, SUM(SALDO) AS SALDO "&_
'     "FROM [SQLSERVER].HandHeld.dbo.[el_SALDOS CHEQUES SANTIAGO] "&_
'     "GROUP BY AUX_VALOR3 "&_
'	 "HAVING (AUX_VALOR3 = '" & rs.fields("codlegal") & "')"
	 
SQL0="SELECT [codlegal] AS AUX_VALOR3,[monto_chcartera] as SALDO "&_
	 "FROM [HANA].[dbo].[HANA_SALDO] "&_
	 "Where codLegal = '" & rs.fields("codlegal") & "'"
	 
'sql0="select '' as AUX_VALOR3, 0 AS SALDO"
'response.write(SQL0)
'response.write SQL0

set rs0=oConn.execute(SQL0)
if rs0.eof and rs0.bof then
saldo=0
else
saldo = rs0.fields("saldo")
end if

set rs1=oConn.execute(sql1)

set rs2=oConn.execute(sql2)

ctacte	= rs.fields("ctacte")

'*************************
'81201000-K',
'78627210-6',
'81537600-5',
'96829710-4',
'96618540-6',
'76027291-4',
'76012833-3')

response.write(empresa)
if ucase(empresa)="DESA" then 

	if (rs.fields("codlegal")="81201000-K" OR rs.fields("codlegal")="78627210-6" OR rs.fields("codlegal")="81537600-5" OR rs.fields("codlegal")="96829710-4") and nuser<>359 and nuser<>004 and nuser<>367 then
		%><BR>
		<CENTER><TABLE bgcolor="#EAF8D6">
		<TR>
			<TD bgcolor="#FFFFFF"><B>Alerta</B></TD>
		</TR>
		<TR>
			<TD>
			<FONT SIZE="3" face="verdana" COLOR="#000066">Todos los pedidos que se bajan desde Comercionet.cl 
			<BR><B>Solo</B> pueden ser integrados directamente en el sistema, mediante "DESAERP"
			<BR>(Escritorio Remoto)
			<BR>
			<BR>Todas las consultas por favor hacer llegar a 
			<BR>Juan Pablo Mu�oz
			<BR>Fono : (02) 4891580
			<BR>Correo: jpmunoz@desa.cl</FONT>
			</TD>
		</TR>
		<TR>
			<TD></TD>
		</TR>
		</TABLE>
		<BR><input type='button' value='<< Atras' onClick='history.back()'>
		</CENTER>
		<%
		exit function
	end if
end if
'*************************
'------------------------------------
'saldo linea de credito
lmcr=formatnumber(cdbl(rs.fields("LimiteCredito")),0)

if rs2.eof then
vartmp1=0
else
vartmp1=formatnumber(cdbl(rs2.fields("saldo")),0)
end if
if rs1.eof then
vartmp2=0
else
vartmp2=formatnumber(cdbl(rs1.fields("saldo")),0)
end if

vartmp3  = (CLNG(vartmp2) + CLNG(vartmp1))
saldocta = (CLNG(lmcr) - (CLNG(vartmp3)+ clng(saldo)))

%>
<html>
<head>
<title>Verifica estado cliente</title>
<style type="text/css">
<!--
.Estilo1 {
	font-weight: bold;
	color: #FFFFFF;
}
.Estilo3 {font-size: xx-small}
.Estilo4 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style></head>
<body>
<table width="100%" border="0" cellspacing="0">
	<tr>
		<td colspan="4" bgcolor="#666633"><div align="center" class="Estilo1 Estilo1">Verificacion de estado cliente</div></td>
	</tr>
	<tr>
		<td colspan="3" bordercolor="#999999"><strong><%= left(rs.fields("razonsocial"),28)%></strong>
			<hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Rut		</td>
		<td colspan="2"><strong><%= rs.fields("codlegal") %>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo2"><hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Giro</td>
		<td colspan="2"><strong>
			<% = rs.fields("Giro") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;."%>
		</strong></span></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo2"><hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Monto Credito </td>
		<td colspan="2"><strong><% = formatnumber(cdbl(rs.fields("LimiteCredito")),0) %>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo2"><hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Cond. Pago </td>
		<td colspan="2"><strong><%= rs.fields("condpago")%>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Mnto Vencido </td>
		<td colspan="2"><span class="Estilo4">
		<%
		if rs2.eof then
		response.Write("0") 
		else 
		response.Write( formatnumber(cdbl(rs2.fields("saldo")),0)) 
		end if
%>
		&nbsp;</span></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo2"><hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Mnto. A Vencer </td>
		<td colspan="2"><strong><%
		if rs1.eof then
		response.Write("0") 
		else 
		response.Write( formatnumber(cdbl(rs1.fields("saldo")),0)) 
		end if
%>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo2"><hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Cheque en cartera </td>
		<td colspan="2"><strong><% response.Write( formatnumber(cdbl(saldo),0))  %>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3" class="Estilo2"><hr></td>
	</tr>
	<tr>
		<td class="Estilo2">Cred. Disponible </td>
		<td colspan="2"><strong>
<%
if saldocta<0 then
response.Write("<font color='#FF0000'>" & formatnumber(cdbl(saldocta),0) & "</font>")
else
response.Write( formatnumber(cdbl(saldocta),0))
end if
%>&nbsp;</strong></td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
		<td colspan="3"><span class="Estilo2"><a href="../otros/detalle.asp?cliente=<%= rs.fields("codlegal")%>">Ver detalle</a></span></td>
	</tr>
	<tr>
		<td><form method="post" action="/palm/Cliente/Cliente.asp?nuser=<%=nuser%>&empresa=<%=empresa%>">
			<INPUT TYPE="hidden" name="STEP" Value="EDITAR">
			<INPUT TYPE="hidden" name="RUT" Value="<%=rs.fields("codlegal")%>">
			<input style="font-weight:bold" type="submit" value="Act. Cliente">
		</form></td>
		<td><!-- <form method="post" action="../otros/ccd_pagos.asp">
		<input type="hidden" name="vendedor" value="<%= nuser %>">
		<input type="hidden" name="cliente" value="<%= rs.fields("codlegal") %>">
		<input style="font-weight:bold" type="submit" value="Ingreso Pago">
		</form> --></td>
  </tr>
  <tr>
    <td>
<%

if len(ctacte)=0 then ctacte=cliente

if rutok="ok" then 
	SQL="SELECT vigencia " & _
	"FROM todo.flexline.CtaCte " & _
	"WHERE empresa='DESA' and tipoctacte='cliente' and ctacte='" & ctacte & "'"
	if ucase(empresa)="DESAZOFRI" then sql=replace(sql,"'DESA'","'DESAZOFRI'")
	set rs=oConn.execute(sql)

	if ucase(trim(rs.fields("vigencia")))<>"S" then
	response.write("Cliente NO VIGENTE")
	else

with response
	.write "<form method='post' action='progrecepcion.asp"
	.write "?vend=" & vend & "&cliente=" & ctacte & "&id_porfolio=" & id_porfolio & "&empresa=" & empresa & "&tipodocto=" & tipodocto & "&sw_iva=" & sw_iva & "'>"
	.write "<input style='font-weight:bold;' type='submit' value='Hacer Pedido'>"
	.write "</Form>"
end with
	end if
else
	response.write("Problemas con la creacion del cliente<BR>No se puede generar Pedido")
	response.write("<BR><input type='button' value='<< Atras' onClick='history.back()'>")
end if
	
if left(ucase(vend),4)=ucase("call") then
	response.write "<form method='post' action='./noventa/noventa_call.asp"
	response.write "?vend=" & vend & "&cliente=" & ctacte & "&id_porfolio=" & id_porfolio & "'>"
	response.write "<input style='font-weight:bold;' type='submit' value='Hacer No Venta'>"
	response.write "</Form>"
end if
%>  </td> <td><%
with response
	'.write "<form method='post' action='http://164.77.199.68/palm/cliente/editcondpag.asp"
	'.write "?vend=" & vend & "&ctacte=" & ctacte & "&paso=cambiar'>"
	'.write "<input style='font-weight:bold' type='submit' value='Camb Cnd Pag'>"
	'.write "</Form>"
end with %>&nbsp;
      </td>
	</tr>
</table>
</body>
</html>
<%
end function 'verificacliente()
'--------------------------------------------------------------------------------
%>
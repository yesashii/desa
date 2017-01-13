<% 'on error resume next
private oConn, oConn2, rs, rs2, sql, sql2, SQL1, rs1, rs0, SQL0 
dim vend

vend	= trim(request.QueryString("vend"))
cliente = trim(request.QueryString("cliente"))
nuser	=right("0" & request.cookies("pdausr"),3)

set oConn = server.CreateObject("ADODB.connection")

oConn.open "provider=SQLOLEDB; Data Source=SQLSERVER; initial catalog=todo; user id=flexline; password=corona;"

SQL="SELECT     CodLegal, RazonSocial, CtaCte, Giro, CondPago, Comuna, Telefono, LimiteCredito, PorcDr1, Texto3, Sigla, telefono,  email, Contacto " & _
"FROM         Flexline.CtaCte " & _
"WHERE     (Empresa = 'DESA') AND (TipoCtaCte = 'CLIENTE') AND (CtaCte = '" & cliente & "')"
'================
'0|Dia Visita
'1|Desde Mañana
'2|Hasta Mañana
'3|Clase
'4|Desde Tarde
'5|Hasta Tarde
'6|Descuento
'7|Dias Recibe
'8|Nada


set rs=oConn.execute(sql)

if rs.eof then
	on error resume next
end if

response.write(rs.fields("RazonSocial") & "<BR>")
response.write("Giro&nbsp;:&nbsp;" & rs.fields("Giro") & "<BR>")
If isnull(rs.fields("Texto3")) then
	a = Split("|09:00|11:59|D|12:00|18:00||LUN, MAR, MIE, JUE, VIE|", "|")
	'For i = LBound(a) To UBound(a)
	'	response.write(i & " = " & a(i) & "<BR>")
	'Next
Else
	if len("hola")=0 then
		a = Split("|09:00|11:59|D|12:00|18:00||LUN, MAR, MIE, JUE, VIE|", "|")
	Else
		if instr(1,rs.fields("Texto3"),"|")>0 then
			a = Split(rs.fields("Texto3"), "|")
		else
			a = Split("|09:00|11:59|D|12:00|18:00||LUN, MAR, MIE, JUE, VIE|", "|")
		End if
	End if
End if

'response.write(trim(rs.fields("Texto3")))
'response.write(sql)
if UBound(a) < 7 then 
	on error resume next
End if

if instr(1,a(7),"SAB")=0 then
	SiSabado="No"
Else
	SiSabado="SI"
end if
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" colspan="2" bgcolor="#C0C0C0">
      <p align="center"><font face="Arial"><b>Horario Recepción Semanal</b></font></td>
  </tr>
  <tr>
    <td width="33%" rowspan="2">


<TABLE border="1">
<%
	e = Split(a(7),",")
	For i = LBound(e) To UBound(e)
		response.write("<TR><TD>" & _
		replace(replace(replace(replace(replace(e(i),"LUN","Lunes"),"MAR","Martes"),"MIE","Miercoles"),"JUE","Jueves"),"VIE","Viernes") & _
		"</TD></TR>")
	Next
%>	
</TABLE>


	</td>
    <td width="67%">

	<TABLE border="1">
	<TR>
		<td width="33%" rowspan="2">Mañan</TD>
		<TD>Desde</TD>
		<TD><%=a(1)%></TD>
	</TR>
	<TR>
		<TD>Hasta</TD>
		<TD><%=a(2)%></TD>
	</TR>
	</TABLE>

	</td>
  </tr>
  <tr>
    <td width="67%">
	
	<TABLE border="1">
	<TR>
		<td width="33%" rowspan="2">Tarde</TD>
		<TD>Desde</TD>
		<TD><%=a(4)%></TD>
	</TR>
	<TR>
		<TD>Hasta</TD>
		<TD><%=a(5)%></TD>
	</TR>
	</TABLE>

	</td>
  </tr>
  <tr>
    <td width="100%" colspan="2"><BR><font face="Arial" color="#808080"><b>Sábado</b></font>
      <p><font face="Arial">Cliente <b><%=SiSabado%></b> recibe en día </font><font face="Arial">sábado</font></td>
  </tr>
  <tr>
    <td width="33%"></td>
    <td width="67%">
	
	<TABLE border="1">
	<TR>
		<td width="33%" rowspan="2">Sabado</TD>
		<TD>Desde</TD>
		<TD>00:00</TD>
	</TR>
	<TR>
		<TD>Hasta</TD>
		<TD>00:00</TD>
	</TR>
	</TABLE>

	</td>
  </tr>
</table>

<table>
	<tr>
		<td colspan="3">
		


		</td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
	<!-- http://164.77.199.68/palm/otros/ccd.asp?nuser=<%= nuser %>&cliente=<%= rs.fields("codlegal")%> -->
		<td>
		<form method="post" action="http://164.77.199.68/palm/Cliente/Cliente.asp?nuser=<%=nuser%>">
			<INPUT TYPE="hidden" name="paso" Value="EDITAR">
			<INPUT TYPE="hidden" name="ctacte" Value="<%=Cliente%>">
			<INPUT TYPE="hidden" name="sigla" Value="<%=rs.fields("sigla")%>">
			<INPUT TYPE="hidden" name="fono" Value="<%=rs.fields("telefono")%>">
			<INPUT TYPE="hidden" name="mail" Value="<%=rs.fields("email")%>">
			<INPUT TYPE="hidden" name="encarg" Value="<%=rs.fields("Contacto")%>">
			<input style="font-weight:bold" type="submit" value="Act. Cliente">
		</form>
		</td>
		
		<% response.Write("<td colspan='2'><form method='post' action='pedido02.asp?vend=" & vend & "&cliente=" & cliente & "'><input style='font-weight:bold' type='submit' value='Hacer Pedido'></Form>") %>
		</td>
	</tr>
</table>
<FONT SIZE="1" face="Arial" COLOR="#7E7E7E">*Las actualizaciones se generan directamente en el servidor central y al dia sigiente en el servidor de trabajo</FONT>
</body>
</html>

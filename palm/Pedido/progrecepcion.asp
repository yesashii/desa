<!--#include Virtual="/includes/pda.asp"-->
<SCRIPT LANGUAGE="JavaScript">
<!--
function optioncambia2(){
	var fechaentrega='';
	if(document.getElementById('fecha1').checked){fechaentrega=document.getElementById('fecha1').value;};
	if(document.getElementById('fecha2').checked){fechaentrega=document.getElementById('fecha2').value;};
	document.getElementById('fechaentrega').disabled='';
	document.getElementById('fechaentrega').value=fechaentrega;
	document.getElementById('fechaentrega2').value=fechaentrega;
}
function optioncambia(){
	var fechaentrega='';
	if(document.getElementById('fecha1').checked){fechaentrega=document.getElementById('fecha1').value;};
	if(document.getElementById('fecha2').checked){fechaentrega=document.getElementById('fecha2').value;};
	//document.getElementById('fechaentrega').disabled='';
	document.getElementById('fechaentrega').value=fechaentrega;
	document.getElementById('fechaentrega2').value=fechaentrega;
}
//-->
</SCRIPT>
<body topmargin="0" bgcolor="#F2F2F7">
<% on error resume next
private oConn, oConn2, rs, rs2, sql, sql2, SQL1, rs1, rs0, SQL0 
dim vend

vend	   = trim(request.QueryString("vend"))
'vend       = replace(vend,"NN","Ñ")'Error blackberry
cliente    = trim(request.QueryString("cliente"))
nuser	   = right("0" & request.cookies("pdausr"),3)
id_porfolio= trim(request.QueryString("id_porfolio"))
empresa    = trim(request.QueryString("empresa"))
sw_iva	   = request.QueryString("sw_iva")
tipodocto  = request.QueryString("tipodocto")

set oConn = server.CreateObject("ADODB.connection")
oConn.open "provider=SQLOLEDB; Data Source=sqlserver; initial catalog=handheld; user id=sa; password=desakey;"
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0

tipocab="PDA"
call encabezado(nuser, tipocab)

SQL="SELECT     CodLegal, RazonSocial, CtaCte, Giro, CondPago, Comuna, Telefono, LimiteCredito, PorcDr1, Texto3, Sigla, telefono,  email, Contacto, analisisctacte2 " & _
"FROM         Flexline.CtaCte " & _
"WHERE     (Empresa = 'DESA') AND (TipoCtaCte = 'CLIENTE') AND (CtaCte = '" & cliente & "')"
'================
'0|Dia Recepcion
'1|Desde Mañana
'2|Hasta Mañana
'3|Desde Tarde
'4|Hasta Tarde
'5|Desde Sabado
'6|Hasta Sabado
set rs=oConn.execute(sql)

texto3=split(rs.fields("texto3"),"|")

if rs.eof then
	on error resume next
end if

response.write(pc("  Cliente :<B> " & rs.fields("RazonSocial") & "</B><BR>",2))
response.write(pc("  Giro : <B>" & rs.fields("Giro") & "</B><BR>",2))

If isnull(rs.fields("Texto3")) then
	texto3 = Split("|L,M,R,J,V|08:30|11:59|12:00|18:00|00:00:|00:00|", "|")
End If
'response.write(trim(rs.fields("Texto3")))
'response.write(sql)

mayoristaoffvv=""
if not isnull(rs.fields("analisisctacte2")) then mayoristaoffvv=rs.fields("analisisctacte2")


if instr(1,texto3(0),"S")=0 then
	SiSabado="No"
Else
	SiSabado="SI"
end if
%>


<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" colspan="2" bgcolor="#C0C0C0">
      <p align="center"><% =pc("<b> Horario Recepción Cliente</b>",2)%></td>
  </tr>
</TABLE>


<TABLE border="1">
<!-- Dias de atencion -->
<% dias=pc("<B> Dias</B>",2)
SiSabado="NO"
if instr(1,texto3(0),"L")<>0 then 
	dia=pc("Lunes",2)%>
	<TR>
		<TD><%=dias%></TD>
		<TD colspan="2"><%=dia%></TD>
	</TR>
	<% dias=""
End If
if instr(1,texto3(0),"M")<>0 then 
	dia=pc("Martes",2)%>
	<TR>
		<TD><%=dias%></TD>
		<TD colspan="2"><%=dia%></TD>
	</TR>
	<% dias=""
End If
if instr(1,texto3(0),"R")<>0 then 
	dia=pc("Miercoles",2)%>
	<TR>
		<TD><%=dias%></TD>
		<TD colspan="2"><%=dia%></TD>
	</TR>
	<% dias=""
End If
if instr(1,texto3(0),"J")<>0 then 
	dia=pc("Jueves",2)%>
	<TR>
		<TD><%=dias%></TD>
		<TD colspan="2"><%=dia%></TD>
	</TR>
	<% dias=""
End If
if instr(1,texto3(0),"V")<>0 then 
	dia=pc("Viernes",2)%>
	<TR>
		<TD><%=dias%></TD>
		<TD colspan="2"><%=dia%></TD>
	</TR>
	<% dias=""
End If
if instr(1,texto3(0),"S")<>0 then 
	dia=pc("<B>Sabado</B>",2)%>
	<TR>
		<TD><%=dias%></TD>
		<TD colspan="2"><%=dia%></TD>
	</TR>
	<% SiSabado="SI"
End If
'response.write(texto3(0))
%>
	<TR>
		<td colspan="3"><%=pc("<B> Horarios</B>",2)%></TD>
	</TR>
<!-- Hora mañana -->
	<TR>
		<td width="33%" rowspan="2"><%=pc("Lun-Vier<BR> Mañana",2)%></TD>
		<TD><%=pc("Desde",2)%></TD>
		<TD align="center"><%=pc(left(texto3(1),5),2)%></TD>
	</TR>
	<TR>
		<TD><%=pc("Hasta",2)%></TD>
		<TD align="center"><%=pc(left(texto3(2),5),2)%></TD>
	</TR>
<!-- Hora Tarde -->
	<TR>
		<td width="33%" rowspan="2"><%=pc("Lun-Vier<BR> Tarde",2)%></TD>
		<TD><%=pc("Desde",2)%></TD>
		<TD align="center"><%=pc(left(texto3(3),5),2)%></TD>
	</TR>
	<TR>
		<TD><%=pc("Hasta",2)%></TD>
		<TD align="center"><%=pc(left(texto3(4),5),2)%></TD>
	</TR>
<!-- Hora Sabado -->
	<TR>
		<td width="33%" rowspan="2"><%=pc("Solo<BR> Sabado",2)%></TD>
		<TD><%=pc("Desde",2)%></TD>
		<TD align="center"><%= pc(left(texto3(5),5),2)%></TD>
	</TR>
	<TR>
		<TD><%=pc("Hasta",2)%></TD>
		<TD align="center"><%=pc(left(texto3(6),5),2) %></TD>
	</TR>
</TABLE>

<%=pc("Cliente <b> " & SiSabado & "</b> recibe en día sábado",2) %>


<form method="post" action="/palm/Cliente/Cliente.asp?nuser=<%=nuser%>&empresa=<%=empresa%>">
	<INPUT TYPE="hidden" name="STEP" Value="EDITAR">
	<INPUT TYPE="hidden" name="RUT" Value="<%=rs.fields("codlegal")%>">
	<input style="font-weight:bold" type="submit" value="Act. Cliente">
</form>
<HR>
<%
'configuracion sabado lunes
fecha1="L;09:00;M"
fecha2="L;12:00;R"
'sql="select codigo, valor from handheld.flexline.fx_codgen where codgen='cambiofecha'"
fechahoy=year(date) & right("00" & month(date),2) & right("00" & day(date),2)
'fechahoy=year(date)
sql="SELECT    Codigo, Valor, LEFT(RIGHT(Valor, 17), 8) AS desde, RIGHT(Valor, 8) AS Hasta " & _
"FROM         Flexline.FX_CODGEN " & _
"WHERE      codgen='cambiofecha' and (LEFT(RIGHT(Valor, 17), 8) <= '" & fechahoy & "') AND (RIGHT(Valor, 8) >= '" & fechahoy & "') " & _
"ORDER BY desde, Codigo"
'response.write(sql)
set rs=oConn.execute(sql)
do until rs.eof
	'response.write(rs.fields(0) & " : " & rs.fields(1) & "<BR>")
	if rs.fields(0)="fecha1" then fecha1=rs.fields(1)
	if rs.fields(0)="fecha2" then fecha2=rs.fields(1)
rs.movenext
loop
'response.write("<BR>fecha1 : " & fecha1)
'response.write("<BR>fecha2 : " & fecha2)
fecha1s=split(ucase(fecha1),";")
fecha2s=split(ucase(fecha2),";")
'ifdia1
if fecha1s(0)="L" then ifdia1=2
if fecha1s(0)="M" then ifdia1=3
if fecha1s(0)="R" then ifdia1=4
if fecha1s(0)="J" then ifdia1=5
if fecha1s(0)="V" then ifdia1=6
if fecha1s(0)="S" then ifdia1=7
if fecha1s(0)="D" then ifdia1=1
'ifdia2
if fecha2s(0)="L" then ifdia2=2
if fecha2s(0)="M" then ifdia2=3
if fecha2s(0)="R" then ifdia2=4
if fecha2s(0)="J" then ifdia2=5
if fecha2s(0)="V" then ifdia2=6
if fecha2s(0)="S" then ifdia2=7
if fecha2s(0)="D" then ifdia2=1
'Fecha destino 1
if fecha1s(2)="L" then desdia1="Lunes"
if fecha1s(2)="M" then desdia1="Martes"
if fecha1s(2)="R" then desdia1="Miercoles"
if fecha1s(2)="J" then desdia1="Jueves"
if fecha1s(2)="V" then desdia1="Viernes"
if fecha1s(2)="S" then desdia1="Sabado"
if fecha1s(2)="D" then desdia1="Domingo"
'Fecha destino 2
if fecha2s(2)="L" then desdia2="Lunes"
if fecha2s(2)="M" then desdia2="Martes"
if fecha2s(2)="R" then desdia2="Miercoles"
if fecha2s(2)="J" then desdia2="Jueves"
if fecha2s(2)="V" then desdia2="Viernes"
if fecha2s(2)="S" then desdia2="Sabado"
if fecha2s(2)="D" then desdia2="Domingo"

ifhora1=fecha1s(1)
ifhora2=fecha2s(1)
'Mayoristas top clientes con fuesza de ventas
if mayoristaoffvv="CLIENTES F.Z.B.S." or mayoristaoffvv="MAYORISTA TOP" or nuser = 91 then
	enablefechabox=""
else
	'" readonly='true' disabled='disabled' "
	enablefechabox=" readonly='true' disabled='disabled' "
end if
'response.write(ifdia1 & " : " & ifhora1)
fEhora = hour(time) & right("00" & minute(time),2) & right("00" & second(time),2) 'Hora actual
%>
<form method="post" 
action="pedido02.asp?vend=<%=ucase(vend)%>&cliente=<%=cliente%>&id_porfolio=<%=id_porfolio%>&empresa=<%=empresa%>&tipodocto=<%=tipodocto%>&sw_iva=<%=sw_iva%>">
	<TABLE>
	<TR>
		<TD colspan="2" align="center"><%=pc("<B> Fecha Entrega</B>",2)%></TD>
	</TR>
	<TR>
		<TD colspan="2" align="center">

		<SELECT NAME="fechaentrega" id="fechaentrega" <%=enablefechabox %> >
		
		<% for fecha=cdbl(date) to (cdbl(date)+6)
			selected=""
			xd=1
			if weekday(date)>ifdia1 or ( weekday(date)=ifdia1 AND cdbl(fEhora)>cdbl(replace(ifhora1,":","") & "00") ) then 
				if fecha1s(2)="L" then cdia1=2
				if fecha1s(2)="M" then cdia1=3
				if fecha1s(2)="R" then cdia1=4
				if fecha1s(2)="J" then cdia1=5
				if fecha1s(2)="V" then cdia1=6
				if fecha1s(2)="S" then cdia1=7
				if fecha1s(2)="D" then cdia1=1
				if cdia1<weekday(date) then cdia1+7 'si el dia es de la proxima semana
				yd=cdia1-weekday(date)
				if fecha2s(2)="L" then cdia2=2
				if fecha2s(2)="M" then cdia2=3
				if fecha2s(2)="R" then cdia2=4
				if fecha2s(2)="J" then cdia2=5
				if fecha2s(2)="V" then cdia2=6
				if fecha2s(2)="S" then cdia2=7
				if fecha2s(2)="D" then cdia2=1
				'response.write(cdia2 & " : " & weekday(date) )
				xfecha3=cdia2 & " : " & weekday(date)
				if cint(cdia2)<cint(weekday(date)) then cdia2=cdia2+7 'si el dia es de la proxima semana
				xfecha3=cdia2 & " : " & weekday(date)
				xd=cdia2-weekday(date)
			end if
			xfecha=year(cdate(fecha)) & right("00" & month(cdate(fecha)),2) & right("00" & day(cdate(fecha)),2)
			xfecha1=year(cdate(date+yd)) & right("00" & month(cdate(date+yd)),2) & right("00" & day(cdate(date+yd)),2)
			xfecha2=year(cdate(date+xd)) & right("00" & month(cdate(date+xd)),2) & right("00" & day(cdate(date+xd)),2)
			if xfecha=xfecha2 then selected="selected"
			%><option <%=selected%> value="<%=xfecha%>"><%=cdate(fecha)%></option><%
		   next %>
		</SELECT>
		<INPUT TYPE="hidden" name="fechaentrega2" id="fechaentrega2" value="<%=xfecha2%>">
		</TD>
	</TR>
<% if weekday(date)>ifdia1 or ( weekday(date)=ifdia1 AND cdbl(fEhora)>cdbl(replace(ifhora1,":","") & "00") ) then 
'and cdbl(fEhora)>cdbl(replace(ifhora1,":","") & "00")
' if weekday(date)>=ifdia1 and (iif(weekday(date)>ifdia1,true,cdbl(fEhora)>cdbl(replace(ifhora1,":","") & "00") ))
'weekday(date)>=6 AND cdbl(fEhora)>160000 
'if weekday(date)>=ifdia1 AND  then 
	if weekday(date)>=ifdia2 AND cdbl(fEhora)>cdbl(replace(ifhora2,":","") & "00") then 
		selfecha1=" readonly='true' disabled='disabled' "
	else
		selfecha1=""
	end if
	%>
	<TR>
		<TD align="center">
		<INPUT TYPE="radio" NAME="opfecha" id="fecha1" onClick="optioncambia()" value="<%=xfecha1%>" <%=selfecha1%>>
		</TD>
		<TD><label for="fecha1"><B><%=pc(desdia1,2)%></B></label></TD>
	</TR>
	<TR>
		<TD align="center">
		<INPUT TYPE="radio" NAME="opfecha" id="fecha2" onClick="optioncambia()" value="<%=xfecha2%>" checked >
		</TD>
		<TD><label for="fecha2"><B><%=pc(desdia2,2)%></B></label></TD>
	</TR>
<% end if %>
	<TR>
		<TD colspan="2">&nbsp;</TD>
	</TR>
	<TR>
		<TD colspan="2"><input style="font-weight:bold" type="submit" value="Hacer Pedido" onClick="document.getElementById('fechaentrega').disabled=''">
		&nbsp;&nbsp;
		</TD>
	</TR>
	</TABLE>
	
</form>
<!-- <FONT SIZE="3" COLOR="#CC0000"><B>Todos los pedidos seran Facturados con fecha 31 de octubre y el despacho para el proximo Lunes</B></FONT> -->
<% 'response.Write("<form method='post' action='pedido02.asp?vend=" & replace(vend,"Ñ","NN") & "&cliente=" & cliente & "'><input style='font-weight:bold' type='submit' value='Hacer Pedido'></Form>")

%>

<BR><FONT SIZE="1" face="Arial" COLOR="#7E7E7E">*Las actualizaciones se generan directamente en el servidor central y al dia sigiente en el servidor de trabajo</FONT>
</body>
</html>
<%
'--------------------------------------------------------------------------------------------
Function PC(sString, mzise) 
  Dim sWhiteSpace, bCap, iCharPos, sChar 
  sWhiteSpace = Chr(32) & Chr(9) & Chr(13) 
  sString = LCase(sString)
  bCap = True 
  For iCharPos = 1 to Len(sString) 
    sChar = Mid(sString, iCharPos, 1) 
    If bCap = True Then 
      sChar = UCase(sChar) 
    End If 
    ProperCase = ProperCase + sChar 
    If InStr(sWhiteSpace, sChar) Then 
      bCap = True 
    Else 
      bCap = False 
    End If 
  Next  
PC="<FONT SIZE='" & Mzise & "' face='arial' COLOR='#000033'>" & replace(ProperCase," ","&nbsp;") & "</FONT>"
PC=replace(replace(PC,"%20"," "),"***",chr(34))
if mzise=0 then PC=ProperCase
End Function 
'--------------------------------------------------------------------------------------------
Function iif(consulta, valortrue, valorfalse)
	if consulta then
		iif=valortrue
	else
		iff=valorfalse
	end if
End Function
'--------------------------------------------------------------------------------------------
%>
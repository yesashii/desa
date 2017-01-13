<!--#include Virtual="/includes/migrilla.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Ctacte Cliente</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>
<%
nuser      =cint(trim(request.querystring("nuser")))
boton      =trim(cstr(request.form("boton"       )))

codlegal   =trim(cstr(request.form("codlegal"    )))
razonsocial=trim(cstr(request.form("razonsocial" )))
sigla      =trim(cstr(request.form("sigla"       )))
empresa    =trim(cstr(request.form("empresa"     )))

if empresa="DESA"  then nempresa=1
if empresa="LACAV" then nempresa=4

%>
<BODY>
<% 

if len(boton)=0 then
	call inicio()
else
	call busqueda()
end if
%>

</BODY>
</HTML>
<%
'--------------------------------------------------------------------------
sub inicio()
%>
<CENTER>
<FORM METHOD=POST ACTION="" name="frm1" id="frm1">
<TABLE>
<TR>
	<TD align="center"><FONT face="verdana" SIZE="2" COLOR="#000066"><B>Seleccione Empresa</B></FONT></TD>
</TR>
<TR>
	<TD align="center">
		<SELECT NAME="empresa" id="empresa">
			<option value="DESA"     >Distribucion y Excelencia S.A.</option>
			<option value="LACAV"    >Distribuidora LA CAV LTDA.  </option>
			<option value="UNDURRAGA">VIÑA UNDURRAGA              </option>
		</SELECT>
	</TD>
</TR>
<TR>
	<TD align="center"><FONT face="verdana" SIZE="2" COLOR="#000066"><B>Seleccione Tipo de Consulta</B></FONT></TD>
</TR>
<TR>
	<TD align="center"><INPUT name="boton" id="boton" TYPE="submit" value="Consulta por Clientes"></TD>
</TR>
<TR>
	<TD align="center"><INPUT TYPE="button" value="Reporte Completo" 
	onclick="window.open('ctacte.asp?empresa='+document.frm1.empresa.value+'&nuser=<%=right("000" & nuser,3)%>')"></TD>
</TR>
</TABLE>
</FORM>
</CENTER>
<%
end sub' inicio
'--------------------------------------------------------------------------
sub busqueda()
nuser=cint(nuser)
nvendedor=""
sql="select nombre from SQLSERVER.Desaerp.dbo.DIM_vendedores where idvendedor=" & nuser & " and idempresa=" & nempresa
Set rs=oConn.execute(Sql)
if not rs.eof then nvendedor=rs.fields(0)

%>
<FONT SIZE="2" face="verdana" COLOR="#000066"><B>Busqueda de Cliente</B></FONT>
<FORM METHOD=POST ACTION="">
<TABLE>
<TR>
	<TD><FONT face="verdana" SIZE="2" COLOR="#000066">Empresa</FONT></TD>
	<TD><FONT face="verdana" SIZE="2" COLOR="#000066"><%=trim(Empresa)%></FONT></TD>
</TR>
<TR>
	<TD><FONT face="verdana" SIZE="2" COLOR="#000066">Codigo legal</FONT></TD>
	<TD><INPUT TYPE="text" NAME="codlegal" value="<%=trim(codlegal)%>"></TD>
</TR>
<TR>
	<TD><FONT face="verdana" SIZE="2" COLOR="#000066">Razon Social</FONT></TD>
	<TD><INPUT TYPE="text" NAME="razonsocial" value="<%=trim(razonsocial)%>"></TD>
</TR>
<TR>
	<TD><FONT face="verdana" SIZE="2" COLOR="#000066">Sigla</FONT></TD>
	<TD><INPUT TYPE="text" NAME="sigla" value="<%=trim(sigla)%>"></TD>
</TR>
<TR>
	<TD colspan="2">
		<INPUT id="boton" name="boton" TYPE="submit" value="Aplicar Filtro">
	</TD>
</TR>
</TABLE>
<INPUT TYPE="hidden" name="empresa" value="<%=empresa%>">
</FORM>
<%

SQL="select top 100 codlegal , left(min(razonsocial),40) as RazonSocial, min(sigla) from handheld.flexline.ctacte " & _
"where empresa='" & empresa & "' and tipoctacte='cliente' " & _
"group by codlegal " & _
"Order by min(razonsocial)"

if len(razonsocial)>0 then sql=replace(sql,"empresa='" & empresa & "'","empresa='" & empresa & "' and razonsocial like '%" & razonsocial & "%'")
if len(codlegal)   >0 then sql=replace(sql,"empresa='" & empresa & "'","empresa='" & empresa & "' and codlegal    like '%" & codlegal & "%'"   )
if len(sigla)      >0 then sql=replace(sql,"empresa='" & empresa & "'","empresa='" & empresa & "' and sigla       like '%" & sigla & "%'"      )

if len(razonsocial)=0 and len(codlegal)=0 and len(sigla)=0 then 
	sql=replace(sql,"empresa='" & empresa & "'","empresa='" & empresa & "' and ejecutivo='" & nvendedor & "' ")
End if

'response.write(sql)

%><FONT SIZE="1" face="verdana" COLOR="#000066"><B>Click en cliente para ver Detalle</B></FONT><%
call migrilla(SQL, 2, 0,"","ctacte.asp?empresa=" & empresa & "&vendedor=" & nvendedor & "&cliente=rs.fields(0)" & chr(34) & "," & chr(34) & "_self")

end sub 'busqueda
'--------------------------------------------------------------------------
%>
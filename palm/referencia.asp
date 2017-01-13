<!--#include Virtual="/includes/conexion.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
<CENTER><%
	correlativo=recuperavalor("correlativo")
	empresa    =recuperavalor("empresa"    )
	tipodocto  =recuperavalor("tipodocto"  )



	sql="SELECT   Numero, TipoDocto, Correlativo, Local, total " & _
	"FROM         serverdesa.BDFlexline.flexline.Documento as E " & _
	"WHERE empresa='" & empresa & "' and RefTipoDocto = '" & tipodocto & "' and RefCorrelativo = '" & correlativo & "'"
'response.write(sql)
'response.write(consultarapida(sql))

set rs=oConn.execute(sql)

%>
Documentos relacionados
<TABLE border="1">
<TR>
	<TH>Numero</TH>
	<TH>Tipodocto</TH>
	<TH>Correlativo</TH>
	<TH>Local</TH>
	<TH>Total</TH>
</TR>
<%
Do until rs.eof
	%><TR>
		<TD><A HREF="http://pda.desa.cl/palm/documento.asp?empresa=<%=empresa%>&documento=<%=rs(0)%>&tipodocto=<%=rs(1)%>&base=BDFlexline"><%=rs(0)%></A></TD>
		<TD><%=rs(1)%></TD>
		<TD><%=rs(2)%></TD>
		<TD><%=rs(3)%></TD>
		<TD><%=rs(4)%></TD>
	</TR><%
rs.movenext
loop
%></TABLE><%
%>
<BR>
<BR>
<INPUT TYPE="button" value="<< Atras"      onclick="history.back();history.back()">
</CENTER>
</BODY>
</HTML>

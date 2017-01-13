<!--#include Virtual="/includes/migrilla.asp"-->
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
<% 
cliente= request.querystring("cliente")
if len(cliente)=0 then cliente=request.form("cliente")

if len(cliente)=0 then
	call inicio()
else
	call busqueda()
end if

'----------------------------------------------------------------------------
sub busqueda()
	SQL="Select top 100 '<A%20HREF=***http://pda.desa.cl/palm/otros/detalle.asp?cliente='+codlegal+'***>'+codlegal+'</A>' as Codlegal, " & _
	"sigla, Razonsocial  from handheld.flexline.ctacte where empresa='DESA' and ctacte+razonsocial+sigla like '%" & cliente & "%'"
	call migrilla(SQL, 2, 0,"","")
End sub 
'----------------------------------------------------------------------------
sub inicio()
%>
<CENTER>
	<BR>
	<FORM METHOD=POST ACTION="">
		<B>Buscar cliente</B><BR>
		(rut, nombre o sigla)<BR>
		<INPUT TYPE="text" NAME="cliente"><BR>
		<INPUT TYPE="submit" value="Buscar">
	</FORM>
</CENTER>
<%
End sub
'----------------------------------------------------------------------------
%>

<!-- http://pda.desa.cl/palm/otros/detalle.asp?cliente=16735365-7 -->
</BODY>
</HTML>

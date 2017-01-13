<!--#include Virtual="/includes/migrilla.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Cobranza </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY><FONT SIZE="2" face="verdana" COLOR="#000066">
<%
nuser=request.querystring("nuser")
sql="select nombre from Flexline.FX_vende_man where nu_2='" & nuser & "'"
Set rs=oConn.execute(Sql)
vend=rs("nombre")
response.write(vend)
sql="select d.codigo_legal as Cliente, c.razonsocial, facturas_vencidas as Monto_Vencido , V0_30 as De_0a30, " & _
	"v31_90 as de_31a90, v_91_180 as de_91a180, v_180 as Sobre_180, Ch_Protesto " & _
	"from handheld.flexline.PDA_Cobranza as d inner join handheld.flexline.ctacte as c " & _
	"on c.codlegal=d.codigo_legal " & _
	"where d.vendedor_antiguo='" & vend & "' " & _
	"and c.tipoctacte='cliente' and empresa='desa' and right(ctacte,2)=' 1'"
call migrilla(SQL, 4, 0,"","formatnumber")
%></FONT>
</BODY>
</HTML>

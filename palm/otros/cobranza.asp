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

'<A%20HREF=***cobranzadetalle.asp?cl='+d.AUX_VALOR3+'***>'+cast(d.AUX_VALOR3 as nvarchar)+'</A>' as Pedido

sql="select '<A%20HREF=***cobranzadetalle.asp?cl='+d.AUX_VALOR3+'***>'+cast(d.AUX_VALOR3 as nvarchar)+'</A>' as  Cliente, " & _
	"max(d.RazonSocial) RazonSocial,max(d.condpago) condpago,sum(d.debe_ingreso-d.haber_ingreso) Saldo " & _
	"from serverdesa.BDFlexline.flexline.FX1_CtaCte_Deudas as d inner join handheld.Flexline.PDA_Cobranza as c " & _
	"	on d.aux_valor3=c.codigo_legal " & _
	"	inner join (select aux_valor3 ,referencia  " & _
	"		from serverdesa.BDFlexline.flexline.FX1_CtaCte_Deudas " & _
	"		where fechaVcto<'2009-01-04' and referencia<>888888 " & _
	"		group by aux_valor3, referencia " & _
	"	) as f on d.aux_valor3=F.aux_valor3 and d.referencia=f.referencia " & _
	"where c.vendedor_antiguo='" & vend & "' " & _
	"group by d.AUX_VALOR3 " & _
	"having sum(d.debe_ingreso-d.haber_ingreso) >0"
call migrilla(SQL, 4, 0,"","formatnumber")
%></FONT>
</BODY>
</HTML>

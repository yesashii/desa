<!--#include Virtual="/includes/conexion.asp"-->
<!--#include Virtual="/includes/forms2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> clientes </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<style type="text/css">
	<!--
	body {
		background-color: #FFFFFF;
		font-family: Arial;
		font-size: 12px;
	}
	td {
		font-size: 12px;
		border-collapse: collapse;
		border-style: solid;
		border-width:1;
	}
	th {
		background-color: #000033;
		font-family: Arial;
		font-size: 12;
		color: #FFFFFF;
		border-style: solid;
		border-width:1;
	}
	table {
		border: 0;
		border-width:0;
		border-collapse: collapse;
		border-style: solid;
	}
	-->
</style>
</HEAD>

<BODY>
<%
'-------------------------------------------------------------------------------------
'void main
	idempresa    =recuperavalor("idempresa" )
	idvendedor   =recuperavalor("nuser")
	sw_guardar   =recuperavalor("Guardar")
	
	if sw_guardar="Guardar" then
		call guardardatos()
	end if

	call listado()
'end void man
'-------------------------------------------------------------------------------------
sub listado()
	sql="select c.ctacte, c.razonsocial, c.sigla, c.analisisctacte1, c.analisisctacte8, c.ejecutivo, vigencia " & _
	"from serverdesa.BDFlexline.flexline.ctacte as c " & _
	"inner join handheld.flexline.fx_vende_man as v " & _
	"on v.nombre=c.ejecutivo " & _
	"where c.empresa='DESA' and c.tipoctacte='Cliente' " & _
	"and v.numero=" & cint(idvendedor) & " " & _
	"order by razonsocial"
	'"and c.vigencia='S' " & _
	
	set rs=oConn.execute(sql)
	%>
	<B>Cartera vendedor : <%=rs("Ejecutivo")%></B>
	<FORM METHOD=POST ACTION="">
	<!-- <INPUT TYPE="checkbox" NAME="">Ocultar clientes ya clasificados -->
	<TABLE>
		<TR>
			<TH>Patente</TH>
			<TH>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cliente&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TH>
			<TH>Nombre</TH>
			<TH>Sigla</TH>
			<TH>Canal</TH>
			<TH>Vig</TH>
		<TR><%
	do until rs.eof
		sel0=""
		sel1=""
		sel2=""
		if rs(4)=""  then sel0="Selected"
		if rs(4)="S" then sel1="Selected"
		if rs(4)="N" then sel2="Selected"

		%><TR>
			<TD align="center">
				<SELECT NAME="<%=rs(0)%>">
					<option value=""  <%=sel0%>></option>
					<option value="S" <%=sel1%>>S</option>
					<option value="N" <%=sel2%>>N</option>
				</SELECT>
			</TD>
			<TD><%=pc(rs(0),2)%></TD>
			<TD><%=pc(left(rs(1),45),2)%></TD>
			<TD><%=pc(rs(2),2)%></TD>
			<TD><%=pc(rs(3),2)%></TD>
			<TD><%=pc(rs(6),2)%></TD>
		<TR><%
	rs.movenext
	loop
	%></TABLE>
	<INPUT TYPE="submit" value='Guardar' name='Guardar'>
	</FORM>
	<%
end sub'listado
'-------------------------------------------------------------------------------------
sub guardardatos()
	sql="select '' as nada"
	set rs=oConn.execute(sql)
	n=0
	for each campo in request.Form
		if len(request.form(campo))>0 and campo <> "Guardar" then
			'response.write("<BR>" & campo & " : " & request.form(campo) )
			ctacte=campo
			sql="update serverdesa.BDFlexline.flexline.ctacte " & _
				"set analisisctacte8='" & request.form(campo) & "' " & _
				"where tipoctacte='cliente' and ctacte='" & ctacte & "'"
			oConn.execute(sql)
			n=n+1
		end if
	next
	%>Se Modificaron <B><%=n%></B> Registros<BR><%
End sub 'guardardatos()
'-------------------------------------------------------------------------------------
%>
</BODY>
</HTML>

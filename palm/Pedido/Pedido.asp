<%'=====================================================================
'|  PEDIDOS | Ver 2.0 | Simon Hernandez, 20090318
'=======================================================================
':: includes ::
%><!--#include file="prepara.asp" //--><%

':: Variables ::
public id_vendedor, id_producto, nro_producto, nb_vendedor, id_porfolio

if len(id_vendedor)<1 then id_vendedor=request.form("nuser"         )
if len(id_vendedor)<1 then id_vendedor=request.querystring("nuser"  )
if len(id_vendedor)<1 then id_vendedor=request.cookies("fx_nusuario")
if len(id_porfolio)<1 then id_porfolio=request.form("id_porfolio"   )
if len(id_porfolio)<1 then id_porfolio=request.querystring("id_porfolio")
if len(empresa)<1     then empresa=request.querystring("empresa")
'if len(empresa)=0     then empresa="DESA"
'response.write id_porfolio
response.write("empresa : " & empresa)
if empresa="LACAV" then id_porfolio="porfolio2"

sql="select v.nombre from handheld.flexline.fx_vende_man as v where v.numero=" & id_vendedor & " "
if len(empresa)>0 then sql=sql & " and empresa='" & empresa & "'"
nb_vendedor=consultarapida(sql)
paso=request.form("paso")

':: Procesos ::
call encabezado("Pedidos")


sql="SELECT u.porfoliodesa1, u.porfoliodesa2 " & _
	"FROM Flexline.PDA_usuarios as u INNER JOIN " & _
	"Flexline.FX_vende_man as v  ON u.Num_vend = v.nu_2 " & _
	"WHERE     (v.nombre = '" & nb_vendedor & "')" 
	
	if len(empresa)>0 then sql=sql & " and u.empresa='" & empresa & "'"

'SELECT u.porfoliodesa1, u.porfoliodesa2 
'FROM Flexline.PDA_usuarios as u INNER JOIN Flexline.FX_vende_man as v ON u.Num_vend = v.nu_2 
'WHERE (v.nombre = 'pedro morales') and u.empresa='DESA'

set rs=oConn.execute(sql)
'response.write(sql)
if rs.eof then
	response.write(sql)
else
	if ucase(rs.fields(0))="S" and ucase(rs.fields(1))="S" then
		seltipoventa()
	else
		if ucase(rs.fields(0))="N" and ucase(rs.fields(1))="N" then
			selporfolio()
		else
			if ucase(rs.fields(0))="N" and ucase(rs.fields(1))="S" then id_porfolio="porfolio2"
			if ucase(rs.fields(0))="S" and ucase(rs.fields(1))="N" then id_porfolio="porfolio1"
			seltipoventa()
		end if
	end if
end if

call piedepagina()
'------------------------------------------------------------------------------------
function selporfolio()

%><FORM METHOD=POST ACTION="">
<TABLE>
<TR>
	<TD colspan="2">&nbsp;</TD>
</TR>
<TR>
	<TD colspan="2"><B><%=pc("Seleccion Porfolio",3)%></B></TD>
</TR>
<TR>
	<TD>
	<INPUT TYPE="submit" value="Porfolio1" name="id_porfolio">
	</TD>
	<TD>
	<INPUT TYPE="submit" value="Porfolio2" name="id_porfolio">
	</TD>
</TR>
</TABLE>
</FORM><%
end function 'selporfolio
'------------------------------------------------------------------------------------
%>

<!--#include Virtual="/includes/conexion.asp"-->
<!--#include Virtual="/includes/forms2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Edita Datos Bancarios </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
<%
	'dim cliente
	'cliente = "13714116-7 1"
	accion  = recuperavalor("guardar" )
	cliente = recuperavalor("cliente" )
	idbanco = recuperavalor("banco"   )
	ctabanco= recuperavalor("ctabanco")
	titular = recuperavalor("tit"     )
	ruttitul= recuperavalor("rutbco"  )
	sucursal= recuperavalor("sucbco"  )
	serieci = recuperavalor("cibanco" )
	'cliente = "13714116-7 1"
	
	if accion <> "Guardar" then
		call mostrardatos()
	else
		call guardardatos()
	end if
'response.write(cliente)
'--------------------------------------------------------------------------------------------
sub mostrardatos()
'response.write(cliente)
sql="select * from serverdesa.BDFlexline.flexline.ctacte " & _
	"where empresa='DESA' and tipoctacte='Cliente' and ctacte='" & cliente & "'" 
set rs=oConn.execute(sql)
if not rs.eof then
	idbanco =rs("analisisctacte7")
	nombre  =rs("razonsocial")  
	txtdata =rs("texto1")
	at=split(txtdata,"|")
	for x=0 to ubound(at)
		if x=0 then ctabanco=trim(at(x))
		if x=1 then titular =trim(at(x))
		if x=2 then ruttitul=trim(at(x))
		if x=3 then sucursal=trim(at(x))
		if x=4 then serieci =trim(at(x))
	next
end if

%>
<FORM METHOD=POST ACTION="">
<INPUT TYPE="hidden" name="mibanco" id="mibanco" value="<%=idbanco%>">
<INPUT TYPE="hidden" name="cliente" id="cliente" value="<%=cliente%>">
<center>
<table>
  <tr>
    <td>Cliente</td>
    <td><%=cliente%></td>
  </tr>
   <tr>
    <td>Nombre</td>
    <td><%=pc(nombre,0)%></td>
  </tr>
  <tr>
    <td></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Banco</td>
    <td><!-- <input name="BANCO" type="text" id="BANCO" value=""> -->
	<SELECT NAME="BANCO" id="BANCO">
		<OPTION value="ABN"       >ABN       </OPTION>
		<OPTION value="BBVA"      >BBVA      </OPTION>
		<OPTION value="BCI"       >BCI       </OPTION>
		<OPTION value="BICE"      >BICE      </OPTION>
		<OPTION value="BOSTON"    >BOSTON    </OPTION>
		<OPTION value="CHILE"     >CHILE     </OPTION>
		<OPTION value="CORPBANCA" >CORPBANCA </OPTION>
		<OPTION value="DESARROLLO">DESARROLLO</OPTION>
		<OPTION value="EDWARDS"   >EDWARDS   </OPTION>
		<OPTION value="ESTADO"    >ESTADO    </OPTION>
		<OPTION value="FALLABELLA">FALLABELLA</OPTION>
		<OPTION value="ITAU"      >ITAU      </OPTION>
		<OPTION value="SANTANDER" >SANTANDER </OPTION>
		<OPTION value="SANTIAGO"  >SANTIAGO  </OPTION>
		<OPTION value="SCOTIABANK">SCOTIABANK</OPTION>
		<OPTION value="SECURITY"  >SECURITY  </OPTION>
	</SELECT>
	<%=idbanco%>
	</td>
  </tr>
  <tr>
    <td>N&ordm; Cuenta&nbsp;&nbsp; </td>
    <td><input name="CTABANCO" type="text" id="CTABANCO" value="<%=ctabanco%>"></td>
  </tr>
  <tr>
    <td>Titular</td>
	
    <td><input name="TIT" type="text" id="TIT" value="<%=titular%>"></td>
  </tr>
  <tr>
    <td>Rut</td>
	
    <td><input name="RUTBCO" type="text" id="RUTBCO" value="<%=ruttitul%>"></td>
  </tr>
  <tr>
    <td>Sucursal</td>
    <td><input name="SUCBCO" type="text" id="SUCBCO" value="<%=sucursal%>"></td>
  </tr>
  <tr>
    <td>Serie CI </td>
    <td><input name="CIBANCO" type="text" id="CIBANCO" value="<%=serieci%>"></td>
  </tr>
</table>
<br>
<input type="button" value="<< Atras" onClick="history.back()">&nbsp;
<input type="submit" name="guardar" value="Guardar"> 
</center>
</form>
<%
End sub 'mostrardatos()
'--------------------------------------------------------------------------------------------
sub guardardatos()
	texto1=ctabanco & "|" & titular & "|" & ruttitul & "|" & sucursal & "|" & serieci  & "|" 
	e=split(cliente," ")
	sql="update handheld.flexline.ctacte " & _
	"set analisisctacte7='" & trim(idbanco) & "' , texto1='" & texto1 & "' " & _
	"where tipoctacte='cliente' and codlegal='" & e(0) & "'"
	oConn.execute(sql)
	sql=replace(sql,"handheld.","serverdesa.BDFlexline.")
	oConn.execute(sql)

	%><CENTER>
	Los Datos Se guardaron Correctamente
	<BR>
	<BR>
	<INPUT TYPE="button" NAME="atras" value="Aceptar" onClick="history.back();history.back()">
	</CENTER><%

end sub 'guardar()
'--------------------------------------------------------------------------------------------
%>
</BODY>
</HTML>

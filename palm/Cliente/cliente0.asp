<%
'response.write("<basefont size='2' color='#000000' face='verdana'>")
Dim RUT, RAZSOC, GIRO, CREDITO, DESCUENTO, CONDPAGO, BANCO, CTABANCO, TITULAR, RUTBCO, SUCBCO
Dim CIBANCO, SIGLA, CALLE, NCALLE, OFICINA, COMUNA, CIUDAD, REGION, ZONA, ENCARLOCAL, CARGOENC
Dim FONOENC, MOVILENC, MAILENC, DIAV, TIPOV, SEMANAV, DIAR, HAI, HAT, HFI, HFT, HSI, HST, CTACTE
Dim TIPO, oConn, oConn2, EJECUTIVO, DIRECCION1, DIRECCION2, TEXTO1, TEXTO2, TEXTO3, GIRO2, GRUPO, TIPOG
Dim VIGENCIA, VCREDITO, ESTADO, execlog, empresa
'public mmivalor(10)
response.Buffer=true
set oConn = server.CreateObject("ADODB.Connection")
oConn.Open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=flexline;Password=flexline"
set oConn2 = server.CreateObject("ADODB.Connection")
oConn2.Open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=pda;User Id=sa;Password=desakey"

vendedor = request.querystring("nuser")
if len(vendedor)<2 then vendedor=request.Cookies("pdausr")
empresa  = request.querystring("empresa")

usuario="ROOT"
if empresa="DESAZOFRI" then usuario="PDA"

step = request.querystring("step")
if len(step)=0 then step = request.form("step")

'Obtener ejecutivo
sqlEJ="select top 1 nombre from handheld.Flexline.PDA_usuarios where num_vend = '"& right("000" & vendedor,3) &"'"
Set rsEJ = oConn2.execute(sqlEJ)

EJECUTIVO = rsEJ(0)

call encabezado()
call DATOSA()
'llamada a secciones
if step=""   then call ingresorut()
if step="00" then call validarut()
if step="01" then call sugeridos()
if step="02" then call datosbanco()
if step="03" then call detcliente()
if step="04" then call dirDespacho3()'2
if step="14" then call dirDespacho2()
if step="05" then call datencargado()
if step="06" then call progvisita()
if step="07" then call progrecep()
if step="08" then call detalle()
if step="09" then call guardar()
if step="EDITAR" then call selecciona()
'fin llamada secciones
call piepag()
'-------------------------------------------------------------------------------------------
sub encabezado() %>
<html>
<head>
<title>Sistema Creaci&oacute;n de Clientes DESA</title>
<script language="JavaScript" src="validaciones.js"></script>
<style type="text/css">
	#noprint {
		display:none;
	}
	body{
		font: 12px Arial;
		background-color: #FFFFFF;
		color: #000033;
	}
	td{
		font: 12px verdana;

	}
	input {
		color: #000033;
		/*border-style: solid;
		border-width: 4px;*/
	}
</style>
</head>
<body style="margin:0">
<% 
end sub 'encabezado
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub ingresorut()
%>
<div style="background-color:#666633; color:#FFFFFF; text-align:center">Creaci&oacute;n Clientes (Vendedor: <%=EJECUTIVO%>)</div>
<form name="datcte" method="post" action="" onSubmit="return validar0(this)">
<center>
<br>
<table align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>*RUT</td>
    <td><input name="RUT" type="text" id="RUT" tabindex="0"></td>
  </tr>
</table>
    <FONT SIZE="" COLOR="#808080">Ejemplo:&nbsp;96568970-2</FONT>
    <br><br>
<input type="submit" value="Continuar &raquo;&raquo;" >
<input type="hidden" name="step" value="00">
</center>
</form>
<HR>
<CENTER>
<FORM METHOD=POST ACTION="listaporvendedor.asp?nuser=<%=vendedor%>">
	<!-- <BR>Clasificar clientes entre 
	<BR>los que tienen patente de licores y los que no -->
	<INPUT TYPE="submit" value="Clasificar">
</FORM>
</CENTER>
<%
end sub 'ingresorut()
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub validarut()
Sql="SELECT * FROM CtaCte WHERE (TipoCtaCte = 'cliente') "&_
"AND (CodLegal = '" & replace(replace(RUT,".",""),",","") & "') AND (Empresa = '"& empresa &"')"
set rs=oConn.execute(Sql)
if rs.eof and rs.bof then 
Call Razon_y_otros()
'call sugeridos()
else
call selecciona()
end if
rs.close
end sub 'validarut
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub Razon_y_otros()
'validar rut 2
	if instr(rut,"-")=0 then
		response.write("<CENTER><BR><B>Falta el simbolo (-) <BR> que separa al Digito verificador</B><BR><BR><input type='button' value='<< Menu' onClick='history.back()'></CENTER>")
		exit sub
	end if 
	if instr(rut," ")<>0 then
		response.write("<CENTER><BR><B>El rut se ingresa sin espacios</B><BR><BR><input type='button' value='<< Menu' onClick='history.back()'></CENTER>")
		exit sub
	end if 
	if len(rut)>10 then
		response.write("<CENTER><BR><B>Por favor revice el rut este numero no es valido : " & rut & "</B><BR><BR><input type='button' value='<< Menu' onClick='history.back()'></CENTER>")
		exit sub
	end if 

sql="SELECT * "&_
"FROM flexline.GEN_TABCOD "&_
"WHERE (EMPRESA = '"& empresa &"') AND (TIPO = 'analisisctacte1') " & _
"and left(codigo,2) not in ('00','22','23','05','06','07')"
'response.Write(sql)
set rs=oConn.execute(sql)

%>
<div style="text-align:center; background-color:#666633; color:#FFFFFF">
Datos Cliente
</div>
<br>
<form method="post" action="" onSubmit="return validar1(this)">
<% call reenviar %>
<table align="center">
  <tr>
    <td>*Raz�n Social&nbsp;</td>
	<td><input name="RAZON" type="text" id="RAZON"></td>
  </tr>
  <tr>
    <td>*Canal</td>
	<td>&nbsp;</td>
  </tr>
</table>
<center>    <select name="GIRO" id="GIRO">
      <option value="">Seleccione....</option>
<% 
do until rs.eof
%>
      <option value="<%= rs.fields("codigo") %>">
	  <%= left(rs.fields("descripcion"),30) %></option>
<%
rs.movenext
loop
%>
    </select>
<br><br>
<input type="submit" value="Siguiente ��">
<input name="step" type="hidden" value="01">
</center>
</form>
<%
End Sub 'Razon_y_otros()
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub sugeridos()
%><div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Datos Sugeridos Por Vendedor</b></div>
<form method="post" action="">
<% call reenviar %>
<center>
<br>
<table>
  <tr>
    <td>Monto Cr&eacute;dito </td>
	<td><input name="CREDITO" type="text" id="CREDITO" value="500.000"></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td><FONT SIZE="" COLOR="#FFFFFF"></FONT></td>
    <td><select name="DESCUENTO" id="DESCUENTO" style="visibility:hidden;display:none">
				
				<!-- <option value="5" selected   > 5,0 %  </option>
				<option value="6.9" > 6,9 %</option>
				<option value="7"   > 7,0 %</option>
				<option value="8"   > 8,0 %</option>
				<option value="8.5" > 8,5 %</option>
				<option value="9"   > 9,0 %</option>
				<option value="10"  >10,0 %</option>
				<option value="10.2">10,2 %</option>
				<option value="11"  >11,0 %</option>
				<option value="12"  >12,0 %</option>
				<option value="13"  >13,0 %</option>
				<option value="15"  >15,0 %</option>
				<option value="18"  >18,0 %</option>
				<option value="20"  >20,0 %</option>
				<option value="20.9">20,9 %</option>
				<option value="21"  >21,0 %</option>	 -->
				<% 
				sql="select texto4 from flexline.gen_tabcod " & _
					"where empresa='"& empresa &"' and " & _
					"tipo='ANALISISCTACTE1' and CODIGO='" & giro & "'"
				set rs=oConn.execute(Sql)
				if rs.eof then 
					midescuento=13
				else
					midescuento=cint(rs.fields(0))
				end if
				if midescuento<5 then midescuento=5
				%>
				<option selected value="<%=midescuento%>"  ><%=midescuento%> %</option>
      </select></td>
  </tr>
  <tr>
    <td><br></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2">Condici&oacute;n de Pago</td>
    </tr>
</table>
<% Sql="SELECT NEMOTECNICO, DESCRIPCION, EMPRESA, TIPO "&_
"FROM flexline.GEN_TABCOD "&_
"WHERE (EMPRESA = 'DESA') AND (TIPO = 'GEN_CPAGO_PALM') " & _
"and nemotecnico not in ('5','11','13','8','3')"
set rs=oConn.execute(Sql) %>
<select name="CONDPAGO" id="CONDPAGO">
<%
do until rs.eof
if rs.fields("descripcion") = "EFECTIVO CAMION" then
%>
					<option selected="selected" value="<%= rs.fields("descripcion") %>">
					<%= rs.fields("descripcion") %>
					</option>
<%
else
%>
					<option value="<%= rs.fields("descripcion") %>" >
					<%= left(rs.fields("descripcion"),20) %>
					</option>
<%
end if
rs.movenext
loop
%>

</select>
<br>
<br>
<input type="submit" value="Siguiente &raquo;&raquo;">
<input type="hidden" name="step" value="02">
</center>
</form>
<CENTER><BR>
<BR>Todos los clientes se crean autom�ticamente 
<BR>en condici�n de pago efectivo y $500.000 en cr�dito
<BR>
<BR>Los cr�ditos sugeridos ser�n efectivos
<BR>una vez que sean aprobados por Gerencia</CENTER>
<%
end sub'sugeridos
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub datosbanco()
'if len(CONDPAGO)=0 then CONDPAGO=request.form("CONDPAGO")
'response.write("condpago:" & CONDPAGO)
if not (CONDPAGO = "EFECTIVO CAMION" Or CONDPAGO = "EFECTIVO POR ANTICIP") then
	if tipo<>"EDITAR" then
	validar = " onSubmit='return validar2(this)'"
	end if
else
validar = ""
end if
%><div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Datos Bancarios</b></div> 
<form name="datbanco" method="post" action="" <%= validar %>>
<% call reenviar 
BANCO   =trim(BANCO   )
CTABANCO=trim(CTABANCO)
TITULAR =trim(TITULAR )
RUTBCO  =trim(RUTBCO  )
SUCBCO  =trim(SUCBCO  )
CIBANCO =trim(CIBANCO )

%>
<center>
<table>
  <tr>
    <td>Banco</td>
    <td><!-- <input name="BANCO" type="text" id="BANCO" value="<%=BANCO%>"> -->
	<SELECT NAME="BANCO" id="BANCO">
	<%	sql="select codigo from serverdesa.BDFlexline.flexline.gen_tabcod where empresa='DESA' and tipo = 'ANALISISCTACTE7'"
		set rs=oConn2.execute(sql)
		do until rs.eof 
		xbanco=rs(0)
		xsel=""
		if trim(xbanco)=trim(rs(0)) then xsel=" selected "
	%>
	<OPTION value="<%=xbanco%>" <%=xsel%>><%=xbanco%></OPTION>
	<% rs.Movenext
	   loop %>
	</SELECT>
	</td>
  </tr>
  <tr>
    <td>N&ordm; Cuenta&nbsp;&nbsp; </td>
    <td><input name="CTABANCO" type="text" id="CTABANCO" value="<%=CTABANCO%>"></td>
  </tr>
  <tr>
    <td>Titular</td>
	<% if len(TITULAR)=0 then TITULAR=RAZSOC%>
    <td><input name="TIT" type="text" id="TIT" value="<%=TITULAR%>"></td>
  </tr>
  <tr>
    <td>Rut</td>
	<% if len(RUTBCO)=0 then RUTBCO=RUT %>
    <td><input name="RUTBCO" type="text" id="RUTBCO" value="<%=RUTBCO%>"></td>
  </tr>
  <tr>
    <td>Sucursal</td>
    <td><input name="SUCBCO" type="text" id="SUCBCO" value="<%=SUCBCO%>"></td>
  </tr>
  <tr>
    <td>Serie CI </td>
    <td><input name="CIBANCO" type="text" id="CIBANCO" value="<%=CIBANCO%>"></td>
  </tr>
</table>
<br>
<input type="submit" value="Siguiente &raquo;&raquo;">
<input type="hidden" name="step" value="03">
<% if len(tipo)=0 then TIPO="NUEVO" 'else EDITAR%>
<input type="hidden" name="TIPO" value="<%=TIPO%>">
</center>
</form>
<%
end sub 'datosbanco()
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub detcliente()
%><div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Datos Cliente</b></div>
<center>
<form method="post" name="detalle1" action="" onSubmit="return validar3(this)">
<% call reenviar() 
if request.form("CTACTE")="" then
response.write "<input type='hidden' name='CTACTE' value='"& RUT & " 1'>"
end if
%>
<small>
<table>
 <tr>
   <td>Rut : </td>
   <% RUTPRINT=formatnumber(left(RUT, len(RUT)-2),0) & right(RUT,2) %>
   <td><%= RUTPRINT %>&nbsp;</td>
 </tr>
 <tr>
   <td>Raz&oacute;n Social : </td>
   <td><%= RAZSOC %>&nbsp;</td>
 </tr>
 <tr>
   <td>Giro : </td>
   <td><%= GIRO %>&nbsp;</td>
 </tr>
 <tr>
   <td>Cr&eacute;dito : </td>
   <td><%= formatnumber(CREDITO,0) %>&nbsp;</td>
 </tr>
 <tr>
   <td>Cond. Pago : </td>
   <td><%= CONDPAGO %>&nbsp;</td>
 </tr>
 <tr>
   <td>Descuento : </td>
   <td><%= DESCUENTO %>&nbsp;%</td>
 </tr>
<% if TIPO = "AGREGAR" then ver=" style='display:none'" %>
 <tr<%= ver %>>
   <td>Cta. Cte N&ordm;: </td>
   <td><%= CTABANCO %>&nbsp;</td>
 </tr>
 <tr<%= ver %>>
   <td>Banco : </td>
   <td><%= BANCO %>&nbsp;</td>
 </tr>
 <tr<%= ver %>>
   <td>Titular : </td>
   <td><%= TITULAR %>&nbsp;</td>
 </tr>
 <tr>
   <td>Sigla Local : </td>
   <td><input name="SIGLA" type="text" id="SIGLA" value="<%= SIGLA %>"></td>
 </tr>
</table>
<p>
<%  if len(trim(CONDPAGO))>5 then
	%><input type="submit" value="Siguiente &raquo;&raquo;"><%
	else
	%>Condicion de pago invalida<BR><input type="button" value="Volver" onClick="history.back()"><%
	end if %>

  <input type="hidden" name="step" value="04">
</p>
</small>
</form>
</center>
<%
end sub 'detcliente
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub dirDespacho3()
	if TIPO="EDITAR" then
		dirDespacho()	
		exit sub
	end if
	on error resume next
	letra=left(request.form("letra"),1)
	'response.write("TIPO : " & TIPO )
	if len(letra)=0 then letra=""
	'exit sub
	sql="select C.idcomuna, C.nombre " & _
		"from sqlserver.desaerp.dbo.DIM_COMUNAS as C, " & _
		"sqlserver.desaerp.dbo.DIM_CIUDADES as D " & _
		"where c.idciudad=d.idciudad and replace(c.nombre,'�','N') like '" & letra & "%' " & _
		"group by c.idcomuna, c.nombre " & _
		"order by c.nombre"
	'rs.close
	'set rs=nothing
	set rs=oConn2.execute(sql)
	%><SCRIPT LANGUAGE="JavaScript">
	<!--
	function selletra(vletra){
		document.getElementById('letra').value=vletra;
		document.getElementById('step').value='04';
		document.getElementById('mipaso').value=1;

		//alert(vletra);
		document.frmdir.submit();
	}
	//-->
	</SCRIPT>
	<CENTER><FORM METHOD=POST ACTION="" name="frmdir" id="frmdir"><BR><INPUT TYPE="hidden" name="letra" id="letra">
	<FONT SIZE="2" face="verdana" COLOR="#000066"><B>Busqueda Alfab�tica</B><BR>
	<% 

	for x=65 to 90
		%>[<A HREF="javascript:selletra('<%=chr(x)%>')"><%= chr(x) %></A>] <%
	next
	%>
	<BR><BR>
	<B>Comuna</B>
	<% call reenviar() 
	response.write(pc(mititulo,1))
	%><BR><BR>
	<SELECT name="idcomuna">
	<%do until rs.eof
		response.write(chr(10))
		%><option value="<%=rs.fields(0) %>"><%=rs.fields(1) %></option><%
		rs.movenext
	loop
	rs.close
	%>
	</SELECT>
	<BR><BR>
	<input type="hidden" name="step"     value="14">
	<!-- <INPUT TYPE="hidden" name="idregion" value="<%=idregion%>">
	<INPUT TYPE="hidden" name="idciudad" value="<%=idciudad%>">
	<INPUT TYPE="hidden" name="idcomuna" value="<%=idcomuna%>"> -->
	<INPUT TYPE="hidden" name="mipaso"   value="3">
	
	<INPUT TYPE="button" value="<< Atras " onClick="history.back()">&nbsp;
	<INPUT TYPE="submit" value="Siguiente >>">
	</FORM></CENTER><%
end sub 'dirDespacho3()
'-----------------------------------------------------------------------------
sub dirDespacho2()
on error resume next
mipaso=request.form("mipaso")
idmivalor=request.form("idmivalor")
idregion=request.form("idregion")
idciudad=request.form("idciudad")
idcomuna=request.form("idcomuna")

'response.write(idregion & "<BR>")
'response.write(idciudad & "<BR>")
'response.write(idcomuna & "<BR>")

mititulo="sin titulo"
sql="select 'sin datos' as valor0, 'sin datos' as valor1 "
if TIPO = "EDITAR" then mipaso=3
'mipaso=2'forzar paso 2 (comuna)
if len(mipaso)=0 then mipaso=0
	if mipaso=0  then 
		mititulo="Seleccione Region"
		sql="select idregion, nombre from sqlserver.desaerp.dbo.Dim_regiones order by idregion"
	elseif mipaso=1   then
		idregion=idmivalor
		mititulo="Seleccione Ciudad"
		sql="select d.idciudad, d.nombre from sqlserver.desaerp.dbo.DIM_COMUNAS as C, " & _
		"sqlserver.desaerp.dbo.DIM_CIUDADES as D " & _
		"where c.idciudad=d.idciudad and c.idregion=" & idregion & " " & _
		"group by d.idciudad, d.nombre " & _
		"order by d.nombre"
		elseif mipaso=2  then 
		idciudad=idmivalor
		mititulo="Seleccione Comuna"
		sql="select C.idcomuna, C.nombre " & _
		"from sqlserver.desaerp.dbo.DIM_COMUNAS as C, " & _
		"sqlserver.desaerp.dbo.DIM_CIUDADES as D " & _
		"where c.idciudad=d.idciudad and c.idregion=" & idregion & " and c.idciudad=" & idciudad & " " & _
		"group by c.idcomuna, c.nombre " & _
		"order by c.nombre"
		'sql="select C.idcomuna, C.nombre " & _
		'"from sqlserver.desaerp.dbo.DIM_COMUNAS as C, " & _
		'"sqlserver.desaerp.dbo.DIM_CIUDADES as D " & _
		'"where c.idciudad=d.idciudad " & _
		'"group by c.idcomuna, c.nombre " & _
		'"order by c.nombre"
		elseif mipaso=3 then
		idcomuna=idmivalor
	end if

if mipaso<3 then
	rs.close
	set rs=nothing
	set rs=oConn2.execute(sql)
	%><CENTER><FORM METHOD=POST ACTION=""><BR>
	<% call reenviar() 
	response.write(pc(mititulo,1))
	%><BR><BR>
	<SELECT NAME="idmivalor">
	<%do until rs.eof
		response.write(chr(10))
		%><option value="<%=rs.fields(0) %>"><%=rs.fields(1) %></option><%
		rs.movenext
	loop
	rs.close
	%>
	</SELECT>
	<BR>
	<input type="hidden" name="step"     value="14">
	<INPUT TYPE="hidden" name="idregion" value="<%=idregion%>">
	<INPUT TYPE="hidden" name="idciudad" value="<%=idciudad%>">
	<INPUT TYPE="hidden" name="idcomuna" value="<%=idcomuna%>">
	<INPUT TYPE="hidden" name="mipaso"   value="<%=mipaso+1%>">
	
	<INPUT TYPE="button" value="<< Atras " onClick="history.back()">&nbsp;
	<INPUT TYPE="submit" value="Siguiente >>">
	</FORM></CENTER><%
else
on error resume next
	sql="select r.nombre as region, D.nombre as Ciudad,C.nombre as Comuna " & _
"from sqlserver.desaerp.dbo.DIM_COMUNAS as C, sqlserver.desaerp.dbo.DIM_CIUDADES as D, sqlserver.desaerp.dbo.DIM_REGIONES as R " & _ 
"where c.idciudad=d.idciudad and c.idregion=r.idregion " & _
"and c.idregion=" & idregion & " and c.idciudad=" & idciudad & " and c.idcomuna=" & idcomuna
'response.write(sql)
set rs=oConn2.execute(sql)
	REGION=rs.fields(0)
	CIUDAD=rs.fields(1)
	COMUNA=rs.fields(2)
	call dirDespacho()
rs.close
end if
end sub 'dirDespacho2()
'--------------------------------------------------------------------------------------------------
function PC(texto, formato)
	texto=replace(texto," ","&nbsp;")
	texto="<FONT SIZE='2' face='arial' COLOR='#000066'>" & texto & "</FONT>"
	if formato=1 then texto="<B>" & texto & "</B>"
	pc=texto
end function 'PC()
'--------------------------------------------------------------------------------------------------
sub dirDespacho()
	idcomuna=request.form("idcomuna")
	'response.write(idcomuna)
	if len(idcomuna)>0 then
		if len(COMUNA)=0 then
			sql="select r.nombre as region, D.nombre as Ciudad,C.nombre as Comuna " & _
			"from sqlserver.desaerp.dbo.DIM_COMUNAS as C, sqlserver.desaerp.dbo.DIM_CIUDADES as D, sqlserver.desaerp.dbo.DIM_REGIONES as R " & _ 
			"where c.idciudad=d.idciudad and c.idregion=r.idregion " & _
			" and c.idcomuna=" & idcomuna
			'response.write(sql)
			set rs=oConn2.execute(sql)
			REGION=rs.fields(0)
			CIUDAD=rs.fields(1)
			COMUNA=rs.fields(2)
		end if
	End if
LCK = ""
if TIPO = "EDITAR" then LCK = " readonly"
%>
<div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Direcci&oacute;n Despacho</b></div>
<center>
<form method="post" id="despacho" name="despacho" action="" 
<%= replace(LCK, "readonly","><!--")%> ><%= replace(LCK, "readonly","-->")%>
<% call reenviar() %>
<table>
  <tr>
    <td>
		*<SELECT NAME="tipocalle">
			<option selected value='Calle'  >Calle  </option>
			<option          value='Pasaje' >Pasaje </option>
			<option          value='Avenida'>Avenida</option>
		</SELECT>
	</td>
    <td><input name="CALLE" type="text" id="CALLE" maxlength="70" value="<%= CALLE %>"<%= LCK %>></td>
  </tr>
  <tr>
    <td>*N&uacute;mero</td>
    <td><input name="NCALLE" type="text" id="NCALLE" maxlength="10" value="<%= NCALLE %>"<%= LCK %>></td>
  </tr>
  <tr>
    <td>Oficina/Local<br>/Comentario</td>
    <td><input name="OFICINA" type="text" id="OFICINA" maxlength="20" value="<%= OFICINA %>"<%= LCK %>></td>
  </tr>
  <tr>
    <td>*Comuna</td><INPUT TYPE="hidden">
    <td><input name="COMUNA" type="hidden" id="COMUNA" value="<%= COMUNA %>"<%= LCK %>><%=COMUNA %></td>
  </tr>
  <tr>
    <td>*Ciudad</td>
    <td><input name="CIUDAD" type="hidden" id="CIUDAD" value="<%= CIUDAD %>"<%= LCK %>><%=CIUDAD %></td>
  </tr>
  <tr<%= replace(LCK,"readonly","style='display:none'") %>>
    <td>*Regi&oacute;n<INPUT TYPE="hidden" name="REGION" id="REGION" value="<%=REGION%>"></td>
    <td><%=REGION%><select name="REGIONX" id="REGIONX" disabled = "false" style='display:none'>
<%
sql="select * from sqlserver.desaerp.dbo.DIM_REGIONES as R order by idregion "
set rs=oConn2.execute(sql)
do until rs.eof
	selected=""
	if REGION=rs.fields("NOMBRE") then selected="selected"
	response.write("<option " & selected & " value='" & rs.fields("nombre") & "'>" & rs.fields("nombre") & "</option>")
	rs.movenext
loop
%>
<!-- 			<option value="I REGION"
<% if REGION = "I REGION" then response.write(" selected")%>>I REGION</option>
			<option value="II REGION"  
<% if REGION = "II REGION" then response.write(" selected")%>>II REGION</option>
			<option value="III REGION" 
<% if REGION = "III REGION" then response.write(" selected")%>>III REGION</option>
			<option value="IV REGION"  
<% if REGION = "IV REGION" then response.write(" selected")%>>IV REGION</option>
			<option value="V REGION"
<% if REGION = "V REGION" then response.write(" selected")%>>V REGION</option>
			<option value="VI REGION"
<% if REGION = "VI REGION" then response.write(" selected")%>>VI REGION</option>
			<option value="VII REGION" 
<% if REGION = "VII REGION" then response.write(" selected")%>>VII REGION</option>
			<option value="VIII REGION"
<% if REGION = "VIII REGION" then response.write(" selected")%>>VIII REGION</option>
			<option value="IX REGION"
<% if REGION = "IX REGION" then response.write(" selected")%>>IX REGION</option>
			<option value="X REGION"
<% if REGION = "X REGION" then response.write(" selected")%>>X REGION</option>
			<option value="XI REGION"
<% if REGION = "XI REGION" then response.write(" selected")%>>XI REGION</option>
			<option value="XII REGION"
<% if REGION = "XII REGION" then response.write(" selected")%>>XII REGION</option>
			<option value="METROPOLITANA"
<% if REGION = "METROPOLITANA" Or not TIPO ="EDITAR" then response.write(" selected")%>
>R. METROP.</option> -->
    </select>    </td>
  </tr>
  <tr<%= replace(LCK,"readonly","style='display:none'") %>>
    <td></td>
    <td><select name="ZONA" id="ZONA" style='display:none'>
	  		<option<%if ZONA = "0-0" then response.write(" selected")%>>0-0</option>
<%
for x=1 to 5
  for y =1 to 5
ds = cstr(x&"-"&y)
selA=""
if ds = ZONA then selA=" selected"
      %><option<%=selA%>><%= x %>-<%= y %></option>
<%
  next
next
%>
    </select>    </td>
  </tr>
<% if TIPO = "NUEVO" Or TIPO = "EDITAR" then ver = " style='display:none'" %>
  <tr<%= ver %>>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr<%= ver %>>
    <td colspan="2" align="center"><input type="checkbox" name="DUPLICAR" value="SI">
      <font size="-7"><b>Duplicar como Direcci&oacute;n Comercial</b></font></td>
    </tr>
</table>
<br>
<input type="button" value="Siguiente &raquo;&raquo;" onClick="validadirec()">
<input type="hidden" name="step" value="05">
<INPUT TYPE="hidden" name="checktipe" id="checktipe" value="<%=TIPO%>">
<SCRIPT LANGUAGE="JavaScript">

//-------------------------------------------------------------------------------
function validadirec(){
	var calle =document.getElementById('CALLE').value;
	var ncalle=document.getElementById('NCALLE').value;
	calle = calle.replace(/^\s*|\s*$/g,"");
	ncalle=ncalle.replace(/^\s*|\s*$/g,"");
  if (document.getElementById('checktipe').value=='NUEVO'){
	if (calle==""){
		window.alert('Falta ingresar La Direccion');
		return 0;
	}
	if (calle.toLowerCase()=="gran avenida"){
		window.alert('GRAN AVENIDA, no existe la calle se llama \n GRAN AVENIDA JOSE MIGUEL CARRERA');
		return 0;
	}
	if (calle.toLowerCase()=="alameda" || calle.toLowerCase()=="la alameda"){
		window.alert('ALAMEDA, no existe la calle se escribe \n ALAMEDA LIBERTADOR BERNARDO O');
		return 0;
	}
	
	if (ncalle==""){
		window.alert('Falta ingresar Numeracion');
		return 0;
	}
	if ( isNaN(ncalle) ){
		alert('En el campo numero solo se pueden ingresar numeros, \n Y solo una unica numeracion, si el cliente tiene mas de un local continuo \n los locales adicionales se ingresan en el campo "Oficina/Local/Comentario"');
		return 0 ;
	}

	var palabrascalle=calle.split(" ")

	for (x=0;x<palabrascalle.length;x++){
		//alert('|' + palabrascalle[x] + '|');
		if( !isNaN(palabrascalle[x]) && palabrascalle[x]!='' ){
			alert('No Se permiten numeros en la descripcion de la calle \n Si corresponde escriba en palabras \n Ejemplo : once de septiembre ');
			return 0;
		}
		if(palabrascalle[x].toLowerCase()=='frente'){
			alert('En el campo calle solo se puede ingresar el nombre de la calle \n comentarios y referencias en el campo "Oficina/Local/Comentario"');
			return 0;
		}
		if(palabrascalle[x].toLowerCase()=='av' || palabrascalle[x].toLowerCase()=='av.' || palabrascalle[x].toLowerCase()=='avda' ){
			alert('Si la direccion corresponde a una avenida \n esto se selecciona en el combo \n no se escribe en el nombre de la calle');
			return 0;
		}
		if(palabrascalle[x].toLowerCase()=='panamericana'){
			alert('PANAMERICANA, no existe la calle se llama, \n AUTOPISTA CENTRAL o \n CARRETERA PRESIDENTE EDUARDO F \n entre otros mas dependiendo del tramo');
			return 0;
		}
	}
	
	}// en if tipo Nuevo
	document.despacho.submit();
}
//-------------------------------------------------------------------------------

</SCRIPT>
</form>
<%= replace(LCK,"readonly","<font size='-7'><li>Datos no Modificables</li></font>") %>
</center>
<%
end sub 'dirDespacho()
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub datencargado() 
%><div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Datos Encargado<br>Local</b></div>
<center>
<form name="datEncarg" method="post" action="" onSubmit="return validar5(this)">
<% call reenviar() %>
<table>
  <tr>
  <td>*Responsable Local : </td>
  <td><input name="ENCARLOCAL" type="text" id="ENCARLOCAL" value="<%= ENCARLOCAL %>"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Cargo : </td>
    <td><input name="CARGOENC" type="text" id="CARGOENC" value="<%= CARGOENC %>"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>*Tel&eacute;fono : </td>
    <td><input name="FONOENC" type="text" id="FONOENC" value="<%= FONOENC %>"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Tel. M&oacute;vil </td>
    <td><input name="MOVILENC" type="text" id="MOVILENC" value="<%= MOVILENC %>"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>E-Mail</td>
    <td><input name="MAILENC" type="text" id="MAILENC" value="<%= MAILENC %>"></td>
  </tr>
</table>
<br>
<input type="submit" value="Siguiente &raquo;&raquo;">
<input type="hidden" name="step" value="06">	
</form>
</center>
<%
end sub ' datencargado()
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub progvisita()
%><div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Programaci&oacute;n de visita</b></div>
<center><br>
<form method="post" name="provisit" action="" onSubmit="">
<% call reenviar() %>
<table cellspacing="0" cellpadding="0">
  <tr>
    <td><input name="DIAV" type="radio" value="1"<%
	if DIAV = cstr(1) or not TIPO = "EDITAR" then response.write(" checked")%>> 
      Lunes&nbsp;&nbsp; </td>
    <td><input name="DIAV" type="radio" value="2"<%
	if DIAV = cstr(2) then response.write(" checked")%>>
      Martes&nbsp;&nbsp;</td>
    <td><input name="DIAV" type="radio" value="3"<%
	if DIAV = cstr(3) then response.write(" checked")%>>
      Mi&eacute;rcoles&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3"><br></td>
    </tr>
  <tr>
    <td><input name="DIAV" type="radio" value="4"<%
	if DIAV = cstr(4) then response.write(" checked")%>>
      Jueves&nbsp;&nbsp;</td>
    <td><input name="DIAV" type="radio" value="5"<%
	if DIAV = cstr(5) then response.write(" checked")%>>
      Viernes&nbsp;&nbsp;</td>
    <td><input name="DIAV" type="radio" value="6"<%
	if DIAV = cstr(6) then response.write(" checked")%>>
      S&aacute;bado&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3"><br></td>
    </tr>
  <tr>
    <td colspan="3">
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td>Tipo Visita</td>
    <td><select name="TIPOV">
			<option value="A"<%
	if TIPOV = "A" then response.write(" selected")%>>A - Vend 8 Fono 0</option>
			<option value="B"<%
	if TIPOV = "B" then response.write(" selected")%>>B - Vend 4 Fono 4</option>
			<option value="C"<%
	if TIPOV = "C" then response.write(" selected")%>>C - Vend 4 Fono 0</option>
			<option value="D"<%
	if TIPOV = "D" then response.write(" selected")%>>D - Vend 2 Fono 2</option>
			<option value="E"<%
	if TIPOV = "E" then response.write(" selected")%>>E - Vend 2 Fono 0</option>
			<option value="F"<%
	if TIPOV = "F" then response.write(" selected")%>>F - Vend 1 Fono 3</option>
			<option value="G"<%
	if TIPOV = "G" then response.write(" selected")%>>G - Vend 0 Fono 1</option>
			<option value="H"<%
	if TIPOV = "H" then response.write(" selected")%>>H - Vend 1 Fono 0</option>
			<option value="" <%
	if TIPOV = "" and TIPO = "EDITAR" then response.write(" selected")%>>Cliente Sin Ruta </option>
    </select>
    </td>
  </tr>
  <tr>
    <td>Semana Visita</td>
    <td><select name="SEMANAV">
			<option value="1"<%
	if SEMANAV = "1" then response.write(" selected")%>>Semana 1</option>
			<option value="2"<%
	if SEMANAV = "2" then response.write(" selected")%>>Semana 2</option>
			<option value="3"<%
	if SEMANAV = "3" then response.write(" selected")%>>Semana 3</option>
			<option value="4"<%
	if SEMANAV = "4" then response.write(" selected")%>>Semana 4</option>
    </select>
    </td>
  </tr>
</table>	
	</td>
    </tr>
</table>
<br>
<input type="submit" value="Siguiente &raquo;&raquo;">
<input type="hidden" name="step" value="07">
</form>
</center>
<%
end sub'progvisita
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub progrecep()
%><div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Programaci&oacute;n Recepci&oacute;n</b></div>
<center>
<form method="post" action="">
<% call reenviar() %><br>
<%
'response.write("TIPO : " & TIPO)
'dias=split(DIAR)
'response.write( " ubound(dias) : " & ubound(dias) )
DIAR=replace(DIAR,"L",", L")
DIAR=replace(DIAR,"M",", M")
DIAR=replace(DIAR,"R",", R")
DIAR=replace(DIAR,"J",", J")
DIAR=replace(DIAR,"V",", V")
DIAR=replace(DIAR,"S",", D")
DIAR=replace(DIAR,"D",", D")
if left(DIAR,2)=", " then DIAR=right(DIAR,len(DIAR)-2)

dias=split(DIAR,", ")
for x = lbound(dias) to  ubound(dias)
select case dias(x)
  case "L"
     L=1
  case "M"
     M=1
  case "R"
     R=1
  case "J"
     J=1
  case "V"
     V=1
  case "S"
     S=1
end select
'response.write(dias(x))
next
%><table cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="2"><input name="DIAR" id="DIARL" type="checkbox" value="L"<%
'
if L = 1 or not TIPO = "EDITAR" then response.write(" checked") %>>
      <label for="DIARL">Lunes&nbsp; </label></td>
    <td colspan="2"><input name="DIAR" id="DIARM" type="checkbox" value="M"<%
if M = 1 or not TIPO = "EDITAR" then response.write(" checked") %>>
      <label for="DIARM">Martes&nbsp; </label></td>
    <td><input name="DIAR" id="DIARR" type="checkbox" value="R"<%
if R = 1 or not TIPO = "EDITAR" then response.write(" checked") %>>
      <label for="DIARR">Mi&eacute;rcoles&nbsp; </label></td>
  </tr>
  <tr>
    <td colspan="5"><br></td>
    </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="2"><input name="DIAR" id="DIARJ" type="checkbox" value="J"<%
if J = 1 or not TIPO = "EDITAR" then response.write(" checked") %>>
      <label for="DIARJ">Jueves&nbsp; </label></td>
    <td colspan="2"><input name="DIAR" id="DIARV" type="checkbox" value="V"<%
if V = 1 or not TIPO = "EDITAR" then response.write(" checked") %>>
      <label for="DIARV">Viernes&nbsp; </label></td>
    </tr>
  <tr>
    <td colspan="5"><br></td>
    </tr>
  <tr>
    <td colspan="5">
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td>Ma&ntilde;ana</td>
    <td>&nbsp;</td>
    <td>Tarde</td>
    </tr>
  <tr>
    <td>Desde</td>
    <td>
	<select name="HI" id="HI">
		<% For x = 800 To 1150 step 50
			hora=right("00" & left(x,len(x)-2),2) & ":" & replace(Right(x, 2),"50","30") 
			selected=""
			if HAI=hora then selected=" selected "%>
			<option <%=selected %> ><%=hora %></option>
		<% Next %>
			<option>11:59</option>
    </select>    
	</td>
    <td>&nbsp;</td>
    <td>
	<select name="HIT" id="HIT">
		<% For x = 1200 To 2350 step 50
			hora=right("00" & left(x,len(x)-2),2) & ":" & replace(Right(x, 2),"50","30") 
			selected=""
			if HAT=hora then selected=" selected "%>
			<option <%=selected %> ><%=hora %></option>
		<% Next %>
    </select>
	</td>
    </tr>
  <tr>
    <td>Hasta</td>
    <td>
	<select name="HF" id="HF">
<!-- 		<% For x = 8 To 11 %>
			<option><%= Right("00" & x, 2) & ":00" %></option>
			<option><%= Right("00" & x, 2) & ":30" %></option>
		<% Next %>
			<option selected="selected">11:59</option>
		<% For x = 12 To 17 %>
			<option><%= Right("00" & x, 2) & ":00" %></option>
			<option><%= Right("00" & x, 2) & ":30" %></option>
		<% Next %> -->
		
		<% For x = 800 To 1750 step 50
			hora=right("00" & left(x,len(x)-2),2) & ":" & replace(Right(x, 2),"50","30") 
			selected=""
			if HFI=hora then selected=" selected "%>
			<option <%=selected %> ><%=hora %></option>
		<% 
			if x=1150 then
				selected=""
				if HFI="11:59" then selected=" selected " %>
			<option <%=selected %>>11:59</option>
			<%
			end if
		   Next %>
    </select>
	</td>
    <td>&nbsp;</td>
    <td>
	<select name="HFT" id="HFT">
	<!-- 	<%For x = 12 To 23
			if x = 18 then 
				sel="selected"
			else
				sel = ""
			end if
		%>
			<option <%= sel %>><%= Right("00" & x, 2) & ":00" %></option>
			<option><%= Right("00" & x, 2) & ":30" %></option>
		<% Next %>		-->
		<% For x = 1200 To 2350 step 50
			hora=right("00" & left(x,len(x)-2),2) & ":" & replace(Right(x, 2),"50","30") 
			selected=""
			if HFT=hora then selected=" selected "%>
			<option <%=selected %> ><%=hora %></option>
		<% Next %>
	
			<option>00:00</option>
    </select>
	</td>
    </tr>
</table>	</td>
  </tr>
  <tr>
    <td colspan="5" align="center"><br></td>
  </tr>
  <tr>
    <td colspan="5" align="center">
	<input name="DIAR" type="checkbox" id="DIARS" value="S" <% if S = 1 or not TIPO = "EDITAR" then response.write(" checked") %>>
      <label for="DIARS">Cliente Recibe S&aacute;bado </label></td>
  </tr>
  <tr>
    <td colspan="5" align="center"><br></td>
  </tr>
  <tr>
    <td colspan="5" align="center">Desde 
      <select name="HSI" id="HSI">
<%
for x = 8 to 23
hr1=right("00" & x, 2)&":00"
hr2=right("00" & x, 2)&":30"
%>
				<option><%= hr1 %></option>
				<option><%= hr2 %></option>
<%
next
%>
				<option>00:00</option>
      </select>
      &nbsp;&nbsp; Hasta 
      <select name="HST" id="HST">
<%
for x = 8 to 23
hr1=right("00" & x, 2)&":00"
hr2=right("00" & x, 2)&":30"
%>
				<option <% if x = 18 then response.write("selected") %>><%= hr1 %></option>
				<option><%= hr2 %></option>
<%
next
%>
      </select>      </td>
  </tr>
</table>
<br>
<input type="submit" value="Siguiente &raquo;&raquo;">
<input type="hidden" name="step" value="08">
</form>
</center>
<%
end sub 'progrecep
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub detalle()
%><div style="background-color:#666633; color:#FFFFFF; text-align:center">
<b>Detalle Solicitud</b>   (id vend <%=vendedor%>)</div>
<center>
<form method="post" action="" id="frmclac" name="frmclac">
<% call reenviar() %>
<table cellpadding="0" cellspacing="0">
	<tr>
		<td>Sigla : </td>
		<td><%= SIGLA %>
			&nbsp;</td>
	</tr>
	<tr>
		<td>Raz&oacute;n Social : </td>
		<td><%= RAZSOC %>
			&nbsp;</td>
	</tr>
	<tr>
		<td>Rut.:</td>
		<% RUTPRINT=formatnumber(left(RUT, len(RUT)-2),0) & right(RUT,2) %>
		<td><%= RUTPRINT %>
			&nbsp;</td>
	</tr>
	<tr>
		<td>Giro:</td>
		<td><%= GIRO %>
			&nbsp;
		  </td>
	</tr>
<!--		<tr>
		<td>Vendedor:</td>
		<td><%= EJECUTIVO %>
			&nbsp;
		  </td>
	</tr>-->
	<tr>
		<td colspan="2"><hr></td>
	</tr>
<% if TIPO = "AGREGAR" then ver=" style='display:none'" %>
	<tr<%= ver %>>
		<td colspan="2">Datos Bancarios </td>
	</tr>
	<tr<%= ver %>>
		<td>Banco</td>
		<td><%= BANCO %>&nbsp;</td>
	</tr>
	<tr<%= ver %>>
		<td>Cta. Cte. </td>
		<td><%= CTABANCO %>&nbsp;</td>
	</tr>
	<tr<%= ver %>>
		<td>Tit.</td>
		<td><%= TITULAR %>&nbsp;</td>
	</tr>
	<tr<%= ver %>>
		<% on error resume next 
		RUTPRINT=formatnumber(left(RUTBCO, len(RUTBCO)-2),0) & right(RUTBCO,2) 
		%>
		<td>Rut.</td>
		<td><%= RUTPRINT %>&nbsp;</td>
	</tr>
	<tr<%= ver %>>
		<td>Suc.</td>
		<td><%= SUCBCO %>&nbsp;</td>
	</tr>
	<tr<%= ver %>>
		<td>Serie C.I. </td>
		<td><%= CIBANCO %>&nbsp;</td>
	</tr>
	<tr<%= ver %>>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td colspan="2">Direcci&oacute;n:</td>
	</tr>
	<tr>
		<td>Avenida:</td>
		<td><%= CALLE & " N&ordm; " & NCALLE & " - " & OFICINA %>
			&nbsp;</td>
	</tr>
	<tr>
		<td>Comuna:</td>
		<td><%= COMUNA %></td>
	</tr>
	<tr>
		<td>Ciudad:		</td>
		<td><%= CIUDAD %></td>
	</tr>
	<tr>
		<td>Regi&oacute;n:</td>
		<td><%= REGION %></td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td>Encargado Local: </td>
		<td><%= ENCARLOCAL %>
			&nbsp;</td>
	</tr>
	<tr>
		<td>Tel&eacute;fono:</td>
		<td><%= FONOENC %>
			&nbsp;</td>
	</tr>
	<tr>
		<td>Celular</td>
		<td><%= MOVILENC %>&nbsp;</td>
	</tr>
	<tr>
		<td>E-Mail:</td>
		<td><%= MAILENC %>
			&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td>Tipo Visita : </td>
		<td><%
with response
select case TIPOV 
  case "A"
    .write "A - Vend 8 Fono 0"
  case "B"
    .write "B - Vend 4 Fono 4"
  case "C"
    .write "C - Vend 4 Fono 0"
  case "D"
    .write "D - Vend 2 Fono 2"
  case "E"
    .write "E - Vend 2 Fono 0"
  case "F"
    .write "F - Vend 1 Fono 3"
  case "G"
    .write "G - Vend 0 Fono 1"
  case else
    .write "Sin Ruta"
end select
end with
%>
			&nbsp;</td>
	</tr>
	<tr>
		<td>D&iacute;a Visita : </td>
		<td><%
with response
select case DIAV 
  case "1"
    .write "Lunes"
  case "2"
    .write "Martes"
  case "3"
    .write "Mi&eacute;rcoles"
  case "4"
    .write "Jueves"
  case "5"
    .write "Viernes"
  case "6"
    .write "Sabado"
  case "7"
    .write "Domingo"
end select
end with
%>
			&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2"><b><i>Dia y horario Recepci&oacute;n:</i></b></td>
	</tr>
	<tr>
		<td colspan="2">D&iacute;a o d&iacute;as recepci&oacute;n : <br>
<%= replace(replace(replace(replace(replace(replace(DIAR,"L","LU"),"M"," MA"),"R"," MI"),"J"," JU"),"V"," VI"),"S"," SA") %></td>
	</tr>
	<tr>
		<td>Ma&ntilde;ana</td>
		<td>Tarde</td>
	</tr>
	<tr>
		<td>Desde las 
			&nbsp;
			<%= HAI %>
			&nbsp;
			<br>
		  hasta las 
			&nbsp;
			<%= HFI %>
&nbsp;
			Hrs.</td>
		<td>Desde las 
			&nbsp;
			<%= HAT %>
&nbsp;
			<br>
		  hasta las 
			&nbsp;
			<%= HFT %>
&nbsp;
		Hrs.</td>
	</tr>
<%
If Not InStr(DIAR,"S") = 0 Then
With Response
    .Write "<tr>" & Chr(13)
	.Write "<td colspan='2'><b><i>Horario Recep. S&aacute;bado:" & Chr(13)
	.Write "</i></b></td>" & Chr(13)
	.Write "</tr>" & Chr(13)
    .Write "<tr>" & Chr(13)
	.Write "<td>Desde Las:<br>" & Chr(13)
	.Write HSI
	.Write "Hrs.</td>" & Chr(13)
	.Write "<td>Hasta Las:<br>" & Chr(13)
	.Write HST
	.Write "Hrs.</td>" & Chr(13)
	.Write "</tr>" & Chr(13)
End With
end If
%>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td>Monto Cr&eacute;dito:</td>
		<td><%= formatnumber(CREDITO,0) %></td>
	</tr>
	<tr>
		<td>Condici&oacute;n Pago:</td>
		<td><%= CONDPAGO %></td>
	</tr>
	<tr>
		<td>Descuento</td>
		<td><%= DESCUENTO %>&nbsp;
			<strong>%</strong></td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td align="right">&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
</table>
<FONT SIZE="2" COLOR="#808080" face="arial">Observacion</FONT><BR>
<TEXTAREA NAME="OBS" ID="OBS" ROWS="3" COLS="50" onKeypress="limitalargo(this)"></TEXTAREA>
<input type="hidden" name="step" value="09">
<input type="hidden" name="IDVENDEDOR" value="<%=vendedor%>">
<CENTER>
<% 
if len(trim(CONDPAGO))>0 then
	%> <input type="submit" value="Almacenar datos"><%
else
	%>Condicion de pago invalida<BR><input type="button" value="Volver" onClick="history.back()"><%
end if
%><!--
<input type="button" value="beta" onClick="
	document.getElementById('frmclac').action='guarda.asp';
	document.getElementById('frmclac').submit(); 
">-->
<input type="hidden" name="empresa" value="<%=empresa%>">
</CENTER>
</form>
</center>
<SCRIPT LANGUAGE="JavaScript">
<!--
function limitalargo(objeto){
	if (objeto.value.length>200){
		alert('El Texto a superado el maximo permitido');
		event.returnValue = false;
	};
	if ((event.keyCode > 32 && event.keyCode < 48) 
	 || (event.keyCode > 57 && event.keyCode < 65) 
	 || (event.keyCode > 90 && event.keyCode < 97)){
		event.returnValue = false;
	}
}
//-->
</SCRIPT>
<%
end sub 'detalle
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub selecciona()

Sql="SELECT MIN(CtaCte) AS CtaCte, CodLegal, Sigla, RazonSocial, Giro, Vigencia, CondPago, LimiteCredito, Analisisctacte1 "&_
"FROM         flexline.CtaCte "&_
"WHERE     (TipoCtaCte = 'cliente') AND (Empresa = '"& empresa &"') "&_
"GROUP BY CodLegal, Sigla, RazonSocial, Giro, Vigencia, CondPago, LimiteCredito, Analisisctacte1 "&_
"HAVING      (CodLegal = '"& RUT &"') "&_
"ORDER BY MIN(CtaCte)"
'response.write(Sql)
set rs1=oConn.execute(Sql)

SIGLA	 = rs1.fields("Sigla")
RAZSOC	 = rs1.Fields("RazonSocial")
GIRO	 = rs1.Fields("Analisisctacte1")
ESTCTE   = rs1.fields("vigencia")
CONDPAGO = rs1.fields("CondPago")
CREDITO  = rs1.fields("LimiteCredito") 
if ESTCTE = "S" then stdo_cte = "<FONT color='#66CC00'><b>ACTIVO</b></FONT>"
if ESTCTE = "B" then stdo_cte = "<FONT color='#ff0000'><b>BLOQUEADO / ESPERANDO APROBACION</b></FONT>"
if ESTCTE = "N" then stdo_cte = "<FONT color='#0099ff'><b>INACTIVO / ESPERANDO APROBACION</b></FONT>"
%>
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
	<tr>
		<td colspan="3" align="center" nowrap bgcolor="#666633" style="color:#FFFFFF;"><b>EDITAR
			/ MODIFICAR LOCAL </b><%=vendedor%></td>
	</tr>
	<tr>
		<td colspan="3" nowrap><strong>Datos Cliente </strong></td>
	</tr>
	<tr>
		<td nowrap>Nombre Fantasia:		</td>
		<td colspan="2" nowrap><strong>estado</strong></td>
	</tr>
	<tr>
		<td align="center" nowrap><%= SIGLA %></td>
		<td colspan="2"><%= ESTCTE %></td>
	</tr>
	<tr>
		<td nowrap>R.U.T.: </td>
		<td colspan="2" nowrap>&nbsp;</td>
	</tr>
	<tr>
		<td align="center" nowrap><%= RUT %></td>
		<td colspan="2" nowrap>&nbsp;</td>
	</tr>
	<tr>
		<td nowrap>Raz&oacute;n Social </td>
		<td colspan="2" nowrap>&nbsp;</td>
	</tr>
	<tr>
		<td align="center" nowrap><%= RAZSOC %></td>
		<td colspan="2" nowrap>&nbsp;</td>
	</tr>
	<tr>
		<td nowrap>Giro</td>
		<td colspan="2" nowrap>&nbsp;</td>
	</tr>
	<tr>
		<td align="center" nowrap><%= GIRO %></td>
		<td colspan="2" nowrap>&nbsp;<INPUT TYPE="button" value="Ver Resumen cliente" onClick="window.open('/palm/Clientes.asp?mirut=<%=RUT %>')"></td>
	</tr>
	<tr bordercolor="#000000" bgcolor="#666633">
		<td nowrap style="color:#FFFFFF;">Locales</span></td>
		<td nowrap style="color:#FFFFFF;">NV</span></td>
		<td nowrap>&nbsp;</td>
	</tr>
	<%
Sql="SELECT	 flexline.CtaCte.CtaCte, flexline.CtaCte.CodLegal, "&_
			"flexline.CtaCte.RazonSocial, flexline.CtaCte.Giro, "&_
			"flexline.CtaCte.LimiteCredito, flexline.CtaCte.Texto2, "&_
			"flexline.CtaCte.CondPago, flexline.CtaCte.Analisisctacte7, "&_
			"flexline.CtaCte.Texto1, flexline.CtaCte.Sigla, "&_
			"flexline.CtaCteDirecciones.Direccion, flexline.CtaCteDirecciones.Comuna, "&_
			"flexline.CtaCteDirecciones.Ciudad, flexline.CtaCte.Zona, "&_
			"flexline.CtaCte.Analisisctacte5, flexline.CtaCte.Contacto, "&_
			"flexline.CtaCte.Telefono, flexline.CtaCte.eMail, "&_
			"flexline.CtaCte.Ejecutivo, flexline.CtaCte.Texto3, "&_
			"flexline.CtaCte.Fax, flexline.CtaCte.analisisctacte1, "&_
			"flexline.CtaCte.codpostal "&_
	"FROM	 flexline.CtaCte INNER JOIN flexline.CtaCteDirecciones "&_
			"ON flexline.CtaCte.Empresa = flexline.CtaCteDirecciones.Empresa "&_
			"AND flexline.CtaCte.CtaCte = flexline.CtaCteDirecciones.CtaCte "&_
			"AND flexline.CtaCte.TipoCtaCte = flexline.CtaCteDirecciones.TipoCtaCte "&_
	"WHERE	 (flexline.CtaCte.TipoCtaCte = 'cliente') AND (flexline.CtaCte.Empresa = '"& empresa &"') "&_
			"AND (flexline.CtaCteDirecciones.Principal <> 's') "&_
			"AND (flexline.CtaCte.CodLegal = '"& RUT &"') "&_
	"ORDER BY CAST(SUBSTRING(flexline.CtaCte.CtaCte, "&_
			"CHARINDEX('-', flexline.CtaCte.CtaCte) + 2, LEN(flexline.CtaCte.CtaCte) - "&_
			" CHARINDEX('-', flexline.CtaCte.CtaCte) + 2) AS numeric)"
'response.Write("<br>" & sql)
set rs=oConn.execute(Sql)
do until rs.eof
gn = InStr(1, rs.fields("ctacte"), "-")
call datos_edit(rs)
%>
	<tr>
			<td nowrap bordercolor="#000000"><b>DIRECCI&Oacute;N:</b>&nbsp;</td>
			<td nowrap bordercolor="#000000">&nbsp;</td>
			<td nowrap bordercolor="#000000">&nbsp;</td>
	</tr>
		<tr>
<%NVLOCAL = Mid(rs.fields("CTACTE"), gn + 2) %>
			<td bordercolor="#000000">
<%  DRR=split(rs.fields("direccion") & "|","|")
with response
  .write DRR(0)
  .Write " <b>Of/Local: </b>"
  .write DRR(1)			
  .write"<br><b>LOCAL N&ordm;&nbsp;</b>" & NVLOCAL 
end with  
  %></td>
			<td nowrap bordercolor="#000000"><%= rs.fields("ejecutivo") %></td>
			<td nowrap bordercolor="#000000"><form method="post" action="">
											 <input type="submit" value="EDITAR">
<input type="hidden" name="RUT"        value="<%= RUT        %>">
<input type="hidden" name="RAZON"      value="<%= RAZSOC     %>"> 
<input type="hidden" name="GIRO"       value="<%= GIRO       %>">
<input type="hidden" name="CREDITO"    value="<%= CREDITO    %>">
<input type="hidden" name="DESCUENTO"  value="<%= DESCUENTO  %>"> 
<input type="hidden" name="CONDPAGO"   value="<%= CONDPAGO   %>"> 
<input type="hidden" name="BANCO"      value="<%= BANCO      %>">
<input type="hidden" name="CTABANCO"   value="<%= CTABANCO   %>">
<input type="hidden" name="TIT"        value="<%= TITULAR    %>">
<input type="hidden" name="RUTBCO"     value="<%= RUTBCO     %>">
<input type="hidden" name="SUCBCO"     value="<%= SUCBCO     %>">
<input type="hidden" name="CIBANCO"    value="<%= CIBANCO    %>">
<input type="hidden" name="SIGLA"      value="<%= SIGLA      %>">
<input type="hidden" name="CALLE"      value="<%= CALLE      %>">
<input type="hidden" name="NCALLE"     value="<%= NCALLE     %>">
<input type="hidden" name="OFICINA"    value="<%= OFICINA    %>">
<input type="hidden" name="COMUNA"     value="<%= COMUNA     %>">
<input type="hidden" name="CIUDAD"     value="<%= CIUDAD     %>">
<input type="hidden" name="REGION"     value="<%= REGION     %>">
<input type="hidden" name="ZONA"       value="<%= ZONA       %>">
<input type="hidden" name="ENCARLOCAL" value="<%= ENCARLOCAL %>"> 
<input type="hidden" name="CARGOENC"   value="<%= CARGOENC   %>">
<input type="hidden" name="FONOENC"    value="<%= FONOENC    %>">
<input type="hidden" name="MOVILENC"   value="<%= MOVILENC   %>">
<input type="hidden" name="MAILENC"    value="<%= MAILENC    %>">
<input type="hidden" name="DIAV"       value="<%= DIAV       %>">
<input type="hidden" name="TIPOV"      value="<%= TIPOV      %>">
<input type="hidden" name="SEMANAV"    value="<%= SEMANAV    %>">
<input type="hidden" name="DIAR"       value="<%= DIAR       %>">
<input type="hidden" name="HI"         value="<%= HAI        %>">
<input type="hidden" name="HF"         value="<%= HAT        %>">
<input type="hidden" name="HIT"        value="<%= HFI        %>">
<input type="hidden" name="HFT"        value="<%= HFT        %>">
<input type="hidden" name="HSI"        value="<%= HSI        %>">
<input type="hidden" name="HST"        value="<%= HST        %>">
<% mistep="03" 
if len(trim(banco))=0 then mistep="02" %>
<input type="hidden" name="step" value="<%=mistep%>">
<input type="hidden" name="CTACTE" value="<%= rs.fields("ctacte") %>">
<input type="hidden" name="TIPO" value="EDITAR">
			</form></td>
		</tr>
	<tr>
		<td colspan="3" nowrap><hr></td>
	</tr>
		<%
rs.movenext
loop
%>
	
	<form method="post" action="">
		<input type="hidden" name="RUT"       value="<%= RUT %>"      >
		<input type="hidden" name="RAZON"     value="<%= RAZSOC %>"   >
		<input type="hidden" name="GIRO"      value="<%= GIRO %>"     >
		<input type="hidden" name="CREDITO"   value="<%= CREDITO %>"  >
		<input type="hidden" name="CONDPAGO"  value="<%= CONDPAGO %>" >
		<input type="hidden" name="DESCUENTO" value="<%= DESCUENTO %>">
		<input type="hidden" name="CTACTE"    value="<%= RUT & " " & cdbl(trim(NVLOCAL)) + 1 %>">
		<tr>
			<td nowrap>&nbsp;</td>
			<td nowrap>
				<input type="submit"             value="Crear Nuevo">
				<input type="hidden" name="step" value="03"         >
				<input type="hidden" name="TIPO" value="AGREGAR"    >
			</td>
			<td nowrap>&nbsp;</td>
		</tr>
	</form>
</table>
<BR>
<CENTER>
	<FORM METHOD=POST ACTION="com_clisolicita.asp">
	<TABLE bgcolor="#FFFFCC">
		<TR>
			<TH>Nuevo : Solicitud Cambio Monto credito y Tipo de Pago</TH>
		</TR>
		<TR>
			<TD>Esta Herramiente le permitira solicitar cambios de condiciondes de pagos y limite del monto de credito de sus clientes</TD>
		</TR>
		<TR>
			<TD align="center" >
				<INPUT TYPE="hidden" name="idvendedor" value="<%=request.querystring("nuser")%>">
				<INPUT TYPE="hidden" name="codlegal"   value="<%=RUT%>">
				<INPUT TYPE="submit" value="Ingresar Solicitud">
			</TD>
		</TR>
		<TR>
			<TD align="center" >&nbsp;</TD>
		</TR>
	</TABLE>
	</FORM>
</CENTER>
<%
end sub 'selecciona
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub guardar()
on error resume next
'obtengo Vendedor
SV="SELECT TOP 1 DESCRIPCION FROM GEN_TABCOD "&_
"WHERE (TIPO LIKE '%GEN_VENDEDORES_PALM%') AND (EMPRESA = '"& empresa &"') AND "&_
"CODIGO = '"& right("000" & request.QueryString("nuser"),3) &"'"

SV="select top 1 nombre from handheld.Flexline.PDA_usuarios where num_vend = '"& right("000" & vendedor,3) &"'"

'response.write(sv)
set rs = oConn2.execute(SV)
EJECUTIVO = rs(0)
%><CENTER>Vendedor:<%=EJECUTIVO%></CENTER><%

'fin vendedor
'secciones giro
SG = "SELECT * FROM Flexline.GEN_TABCOD "&_
"WHERE (TIPO = 'analisisctacte1') AND (CODIGO = '"& GIRO &"')"
set rs = oConn.execute(SG)
GIRO2 = rs.fields("texto1")
GRUPO = rs.fields("texto2")
TIPOG = rs.fields("texto3")
'datos banco
'::::::::::::::::
'Numero Cuenta | Nombre Titular | Rut Titular | Sucursal Banco | Serie CI |
 
TEXTO1 = Right("               " & CTABANCO, 15)& "|" &_
         Left(TITULAR & "            ", 12)& "|" &_ 
		 Right("          " & RUTBCO  , 10)& "|" &_
		 Left(SUCBCO & "         ", 9)& "|" &_
		 Right("         " & CIBANCO , 9)& "|"

'otros datos
TEXTO2 = DESCUENTO & "|CREADO:" & DATE & "|"
'programacion entrega
':::::::::::::::
'D�a o d�as Recepci�n | Horario inicio ma�ana | Horario final  ma�ana
'Horario inicio tarde | 'Horario final  tarde | Horario inicio sabado 
'Horario Final  sabado
 
DIAR=replace(DIAR,", ","")
TEXTO3 =DIAR & "|" & HAI & "|" & HFI  & "|" & HAT  & "|" &_ 
		HFT  & "|" & HSI  & "|" & HST  & "|"
'estado de cliente
'si es efectivo
if CONDPAGO = "EFECTIVO CAMION" then
	VIGENCIA = "S"
	VCREDITO = "31-12-"& year(date)
	ESTADO   = ""
	execlog  = 1
else
	VIGENCIA = "B"
	VCREDITO = "01-01-"& year(date)
	ESTADO   = "PENDIENTE"
	execlog  = 0
end if
	'Anula pegunta anterior
	VIGENCIA = "S"
	VCREDITO = "31-12-"& year(date)
	ESTADO   = ""

'almaceno por tipo
select case TIPO
  case "NUEVO"
	'direccion
	DIRECCION1=CALLE&" "&NCALLE&"|" & OFICINA
	DIRECCION2=CALLE&" "&NCALLE&"|" & OFICINA
	call ingnew()
	SQLcontab = "INSERT INTO Flexline.GEN_TABCOD "&_
	"(EMPRESA, TIPO, CODIGO, DESCRIPCION, NEMOTECNICO, VALOR1, VALOR2, VALOR3, VALOR4, VALOR5) "&_
	"VALUES "&_
	"('"& empresa &"', 'CON_CLIENT', '"& RUT &"', '"& RAZSOC &"', '"& RUT &"', '0', '0', '0', '0', '0')"
	oConn.execute(SQLcontab)

	SQLcontab = "INSERT INTO Flexline.GEN_TABCOD " & _
	"(EMPRESA, TIPO, CODIGO, DESCRIPCION, NEMOTECNICO, VALOR1, VALOR2, VALOR3, VALOR4, VALOR5) "&_
	"VALUES "&_
	"('LACAV', 'CON_CLIENT', '"& RUT &"', '"& RAZSOC &"', '"& RUT &"', '0', '0', '0', '0', '0')"
	oConn.execute(SQLcontab)
	' response.write SQLcontab & "<br><br>"
	 
	if execlog=1 then call guardalog()
	with response
	 .write("<p align='center'>debe actualizar el cliente, presione el boton para actualizar<br>")
	 .write("<input type='button' value='Actualizar' onclick="&_
	  chr(34)&"location='../../cliente/subeServer.asp?code="& CTACTE &"&paso=SendData&palm=1&empresa="& empresa &"'"& chr(34) &">")
	  .write "<p align='center'>Local Creado Con &Eacute;xito<br><br>"
	  .write "<input type='button' value='Salir' "
	  .write "onClick=""location='/palm/'""></p>"
	end with
  case "EDITAR"
	call editclie()
	with response
	 .write("<p align='center'>debe actualizar el cliente, presione el boton para actualizar<br>")
	 .write("<input type='button' value='Actualizar' onclick="&_
	  chr(34)&"location='../../cliente/subeServer.asp?code="& CTACTE &"&paso=SendData&palm=1&empresa="& empresa &"'"& chr(34) &">")
	  .write "<p align='center'>Local Creado Con &Eacute;xito<br><br>"
	  .write "<input type='button' value='Salir'"
	  .write "onClick=""location='/palm/'""></p>"
	end with
  case "AGREGAR"
	'direccion
	if request.form("DUPLICAR") = "SI" Then
		DIRECCION1=CALLE&" "&NCALLE'&"|"&OFICINA
	else
		sqlD= "SELECT direccion From CtaCte WHERE CtaCte = '"& RUT & " 1" &"' AND Tipoctacte = 'CLIENTE'"
		set rs = oConn.execute(sqlD)
		DIRECCION1=rs.fields("direccion")
	end if
	OFICINA=left(trim(OFICINA), 15)
	DIRECCION2=CALLE&" "&NCALLE&" | "&OFICINA
	if len(DIRECCION2)>=100 then
		DIRECCION2=left(trim(CALLE),50) & " " & NCALLE & " | " & trim(OFICINA)
	end if

	call ingnew()
	call enviocorreo()

oConn2.Execute(envio) 
	
	if execlog=1 then call guardalog()
	with response
	 .write("<p align='center'>debe actualizar el cliente, presione el boton para actualizar<br>")
	 .write("<input type='button' value='Actualizar' onclick="&_
	  chr(34)&"location='../../cliente/subeServer.asp?code="& CTACTE &"&paso=SendData&palm=1&empresa="& empresa &"'"& chr(34) &">")
	  .write "<p align='center'>Local Creado Con &Eacute;xito<br><br>"
	  .write "<input type='button' value='Salir'"
	  .write "onClick=""location='/palm/'""></p>"
	end with
end select

end sub 'guardar
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub ingnew()
'ingreso en ctacte
OBS=request.form("OBS")
if len(OBS)=0 then 
	OBS=""
else
	OBS="Vend:" & OBS
end if
CREDITO=replace(replace(trim(CREDITO),".",""),",","")
CONDPAGOX="EFECTIVO CAMION"
CREDITOX="500000"
ASEG=0
apruebacredito=0
local=cint(replace(CTACTE,RUT,""))

if local>1 then
	sql="select * from Flexline.CtaCte where empresa='"& empresa &"' and tipoctacte='cliente' and ctacte='" & RUT & " 1'"
	set rs=oConn.execute(sql)
	if not rs.eof then
		CREDITO  =rs("limitecredito"  )
		CONDPAGOX=rs("condpago"       )
		CREDITOX =rs("limitecredito"  )
		DESCUENTO=rs("analisisctacte4")
		ASEG=rs("analisisctacte10")
		apruebacredito=1
	end if
end if

fechacreado=year(date) & right("00" & month(date),2) & right("00" & day(date),2)
'response.write("OBS:" & OBS)
SQLctacte = "INSERT INTO Flexline.CtaCte (EMPRESA, TipoCtaCte, CtaCte, Codlegal, RazonSocial, "&_
"Sigla, Giro, Grupo, Tipo, Analisisctacte1, Analisisctacte4, Analisisctacte5, Analisisctacte7, Analisisctacte9, Texto1, "&_
"Texto2, Texto3, Ejecutivo, CondPago, CodPostal, "&_
"Listaprecio, Zona, Direccion, Ciudad, Comuna, Pais, Telefono, Fax, Email, Contacto, Direccionenvio, "&_
"Limitecredito, Retrasocredito, FechaModif, UsuarioModif, PorcDr1, PorcDr2, PorcDr3, PorcDr4, "&_
"Moneda, Vigencia, VigenciaCredito, Estado, Comentario1, ANALISISCTACTE12, ANALISISCTACTE25,ANALISISCTACTE10) "&_
" VALUES "&_
"('"& empresa &"', 'CLIENTE', '"& CTACTE &"', '"& RUT &"', '"& RAZSOC &"', '"& left(SIGLA,20) &"', "&_
"'" & GIRO2 &"', '"& GRUPO &"', '"& TIPOG &"', '"& GIRO &"', "&_
"'"& DESCUENTO &"', '"& TIPOV&DIAV&SEMANAV &"', '"& BANCO &"', '" & fechacreado & "', '" & TEXTO1 &"', "&_
"'"& TEXTO2 &"', '"& TEXTO3 &"', '"& EJECUTIVO &"', '"& CONDPAGOX &"', '"& ZONA &"', "&_
"'LISTA PRECIO SANTA I', '"& REGION &"', '"& DIRECCION1 &"', '"& CIUDAD &"', '"& COMUNA &"', "&_
"'CHILE', '"& FONOENC &"', '"& MOVILENC &"', '"& MAILENC &"', '"& ENCARLOCAL &"|"&_
 CARGOENC &"', '"& DIRECCION2 &"', '"& CREDITOX &"', '0', '"& date + time &"', '"& usuario &"', "&_
"'0', '0', '0', '0', 'PS', '"& VIGENCIA &"', '"& VCREDITO &"', '"& ESTADO &"','" & OBS & "','02','N','" & ASEG & "')"  
 'response.write SQLctacte & "<br><br>" 
oConn.execute(SQLctacte)
'SQLctacte=replace(SQLctacte,"'"& empresa &"'","'LACAV'")
'oConn.execute(SQLctacte)
	'-------------------------------------------------
'Ingeso Direccion Comercial
SQLctacteD1 = "INSERT INTO CtaCteDirecciones "&_
"(Empresa, CtaCte, Direccion, Comuna, Ciudad, Estado, Pais, Telefono, Fax, CodPostal, Email, ModoEnvio, Principal, TipoCtaCte, Referencia_desa) "&_
"VALUES "&_
"('"& empresa &"','"& CTACTE &"','"& DIRECCION1 &".','"& COMUNA &"','"& CIUDAD &"','','CHILE','"& FONOENC &"','"& MOVILENC &"','','"& MAILENC &"','','S','CLIENTE','" & OFICINA & "')"
 'response.write SQLctacteD1 & "<br><br>"
oConn.execute(SQLctacteD1)
'SQLctacteD1=replace(SQLctacteD1,"'"& empresa &"'","'LACAV'")
'oConn.execute(SQLctacteD1)

	'--------------------------------------------------
'Ingreso Direccion Despacho  
SQLctacteD2 = "INSERT INTO CtaCteDirecciones "&_
"(Empresa, CtaCte, Direccion, Comuna, Ciudad, Pais, Telefono, Principal, TipoCtaCte, CodPostal, ESTADO, FAX, EMAIL, MODOENVIO) "&_
"VALUES "&_
"('"& empresa &"','"& CTACTE &"','"& DIRECCION2 & "','"& COMUNA &"','"& CIUDAD &"','','"& FONOENC &"','','CLIENTE','','','"& MOVILENC &"','"& MAILENC &"','')"
' response.write SQLctacteD2 & "<br><br>"
oConn.execute(SQLctacteD2)
'SQLctacteD2=replace(SQLctacteD2,"'"& empresa &"'","'LACAV'")
'oConn.execute(SQLctacteD2)


'.::: Guarda datos DESAERP ::.
if empresa="DESAZOFRI" then
idempresa=3
else
idempresa=1
end if

idvendedor=request.querystring("nuser")
if len(idvendedor)=0 then idvendedor=34

idcondpago=14
sql="select NEMOTECNICO from serverdesa.BDFlexline.flexline.GEN_TABCOD " & _
		"WHERE (EMPRESA = '"& empresa &"') AND (TIPO = 'GEN_CPAGO_PALM') and (DESCRIPCION='" & CONDPAGO & "')"
set rs=oConn2.execute(sql)
if not rs.eof then idcondpago=rs(0)
	on error resume next
	sw5_aprobacionlc="P"
	sw6_aprobacioncp="P"
	sw5_nomusuario=""
	sw5_fecha=0
	sw6_nomusuario=""
	sw6_fecha=0
	limitecredito_tiene=CREDITOX
	limitecredito=CREDITO
	idcondpago_tiene=12
	condpago=idcondpago
	idsucursal=cint(replace(CTACTE,RUT,""))
	limitecredito_apr=0
	idcondpago_apr=0

	if limitecredito_tiene>=limitecredito or apruebacredito=1 then
		sw5_aprobacionlc="A"
		sw5_nomusuario="ROOT"
		sw5_fecha=year(date()) & right("00" & month(date()),2) & right("00" & day(date()),2)
		limitecredito_apr=limitecredito
	end if
	if cint(idcondpago_tiene)=cint(idcondpago) or apruebacredito=1 then 
		sw6_aprobacioncp="A"
		sw6_nomusuario="ROOT"
		sw6_fecha=year(date()) & right("00" & month(date()),2) & right("00" & day(date()),2)
		idcondpago_apr=idcondpago
	End if
'CONDPAGO
'CREDITO
	sql="select isnull(max(idsolicitud),0) as ult from sqlserver.desaerp.dbo.COM_CLISOLICITA"
	set rs=oConn2.execute(sql)
	idsolicitud=cdbl(rs("ult"))+1
	response.write("<CENTER><BR><FONT SIZE='3'><B>Solicitud Credito Nro. " & idsolicitud & "</B></FONT><BR></CENTER>")
	response.flush()
	sql="select * from sqlserver.desaerp.dbo.COM_CLISOLICITA where idsolicitud=" & idsolicitud 
	rs.close
	rs.Open SQL, oConn2, 1, 3
	rs.addnew
	rs.fields("idempresa"          )=idempresa           'ok
	rs.fields("idsolicitud"        )=idsolicitud         'ok
	rs.fields("idcliente"          )=RUT                 'ok
	rs.fields("idsucursal"         )=idsucursal          'ok
	rs.fields("fechasolicitud"     )=year(date()) & right("00" & month(date()),2) & right("00" & day(date()),2) 'ok
	rs.fields("idvendedor"         )=idvendedor          'ok
	rs.fields("limitecredito_tiene")=limitecredito_tiene 'ok
	rs.fields("idcondpago_tiene"   )=idcondpago_tiene    'ok
	rs.fields("limitecredito_sol"  )=limitecredito       'ok
	rs.fields("idcondpago_sol"     )=idcondpago          'ok
	rs.fields("limitecredito_apr"  )=limitecredito_apr   'ok
	rs.fields("idcondpago_apr"     )=idcondpago_apr      'ok
	rs.fields("sw1_informedicom"   )="N"
	rs.fields("sw1_fecha"          )=0
	rs.fields("sw1_nomusuario"     )=""
	rs.fields("sw1_observacion"    )=""
	rs.fields("sw2_autorizadicom"  )="N"
	rs.fields("sw2_fecha"          )=0
	rs.fields("sw2_nomusuario"     )=""
	rs.fields("sw2_observacion"    )=""
	rs.fields("sw3_ivas"           )="N"
	rs.fields("sw3_fecha"          )=0
	rs.fields("sw3_nomusuario"     )=""
	rs.fields("sw3_observacion"    )=""
	rs.fields("sw4_patente"        )="N"
	rs.fields("sw4_fecha"          )=0
	rs.fields("sw4_nomusuario"     )=""
	rs.fields("sw4_observacion"    )=""
	rs.fields("sw5_aprobacionlc"   )=sw5_aprobacionlc
	rs.fields("sw5_fecha"          )=sw5_fecha
	rs.fields("sw5_nomusuario"     )=sw5_nomusuario
	rs.fields("sw5_observacion"    )=""
	rs.fields("sw6_aprobacioncp"   )=sw6_aprobacioncp
	rs.fields("sw6_fecha"          )=sw6_fecha
	rs.fields("sw6_nomusuario"     )=sw6_nomusuario
	rs.fields("sw6_observacion"    )=""
	rs.fields("sw7_refbancaria"    )="N"
	rs.fields("sw7_fecha"          )=0
	rs.fields("sw7_nomusuario"     )=""
	rs.fields("sw7_observacion"    )=""
	rs.update
	rs.close

'COM_CLIDESPACHO
	
	sql="select * from sqlserver.desaerp.dbo.COM_CLIDESPACHO where codcliente='" & CTACTE & "'"  
	'rs.close
	rs.Open SQL, oConn2, 1, 3
	rs.addnew
	'idempresa, codcliente, idcliente, idsucursal, lun, mar, mie, jue, vie, sab, dom, 
	'horade, horaa, fechaact, responsableact, horamananade, horamananaa, horatardede, horatardea
	rs.fields("idempresa"     )=idempresa           'ok
	rs.fields("codcliente"    )=ctacte
	rs.fields("idcliente"     )=RUT
	rs.fields("idsucursal"    )=idsucursal 
	rs.fields("codzona"       )=0
	rs.fields("lun"           )=1
	rs.fields("mar"           )=1
	rs.fields("mie"           )=1
	rs.fields("jue"           )=1
	rs.fields("vie"           )=1
	rs.fields("sab"           )=0
	rs.fields("dom"           )=0
	rs.fields("horade"        )="08:00"
	rs.fields("horaa"         )="16:00"
	rs.fields("fechaact"      )=fechacreado
	rs.fields("responsableact")="ROOT"
	rs.fields("horamananade"  )="08:00"
	rs.fields("horamananaa"   )="11:59"
	rs.fields("horatardede"   )="13:00"
	rs.fields("horatardea"    )="16:59"
	'rs.fields(""    )=
	
	rs.update
	rs.close
	
	call enviocorreo()
end sub 'ingnew
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub editclie()
off=split(OFICINA,",")
if ubound(off)>1 then 
	'response.write ubound(off)=
	OFICINA=off(0)
end if
'response.write "<HR>"
DIRECCION1=left( CALLE & " " & NCALLE & " | " & OFICINA , 100)
SQLctacte = "UPDATE CtaCte SET Sigla = '"& SIGLA &"', DireccionEnvio = '"& DIRECCION1 &_
            "', Comuna = '"& COMUNA &"', Ciudad = '"& CIUDAD &"', Zona = '"& REGION &_
            "', CodPostal = '"& ZONA &"', Contacto = '"& ENCARLOCAL &"|"& CARGOENC &_
			"', Telefono = '"& FONOENC &"', Fax = '"& MOVILENC &"', eMail = '"& MAILENC &_
			"', Analisisctacte5 = '"& TIPOV&DIAV&SEMANAV &"', Analisisctacte7='" & BANCO & "',Texto1='" & TEXTO1 & "', Texto3 = '"& TEXTO3 &_
			"' WHERE (CtaCte = '"& CTACTE &"') AND (tipoctacte = 'CLIENTE')"
'response.write SQLctacte & "<br><br>"
oConn.execute(SQLctacte)

'direccion despacho
SQLctacteD1 = "UPDATE CtaCteDirecciones SET Direccion = '"& DIRECCION1 &_
			  "', Comuna = '"& COMUNA &"', Ciudad = '"& CIUDAD &"', CodPostal = '"& ZONA &_
			  "', Telefono = '"& FONOENC &"', Fax = '"& MOVILENC &"' WHERE (CtaCte = '"& CTACTE &"') AND "&_
			  "(TipoCtaCte = 'CLIENTE') AND (Principal <> 'S')"
'response.write SQLctacteD1 & "<br><br>"			  
oConn.execute(SQLctacteD1)

SQLctacteD2 = "UPDATE CtaCteDirecciones " & _
			"SET  Telefono = '"& FONOENC &"', Fax = '"& MOVILENC &"' " & _
			" WHERE (CtaCte = '"& CTACTE &"') AND (TipoCtaCte = 'CLIENTE') AND (Principal = 'S')"
'response.write SQLctacteD1 & "<br><br>"			  
oConn.execute(SQLctacteD2)

call enviocorreo()
end sub 'editclie()
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub DATOSA()
RUT			= ucase(trim(replace(replace(replace(request.form("RUT"),".",""),",",""),"  "," ")))
RAZSOC		= ucase(request.Form("RAZON"     ))
GIRO		= ucase(request.Form("GIRO"      ))
CREDITO		= ucase(request.Form("CREDITO"   ))
DESCUENTO	= ucase(request.Form("DESCUENTO" ))
CONDPAGO	= ucase(request.Form("CONDPAGO"  ))
BANCO		= ucase(request.Form("BANCO"     ))
CTABANCO	= ucase(request.Form("CTABANCO"  ))
TITULAR		= ucase(request.Form("TIT"       ))
RUTBCO		= ucase(request.Form("RUTBCO"    ))
SUCBCO		= ucase(request.Form("SUCBCO"    ))
CIBANCO		= ucase(request.Form("CIBANCO"   ))
SIGLA		= ucase(request.Form("SIGLA"     ))
CALLE		= ucase(request.Form("CALLE"     ))
NCALLE		= ucase(request.Form("NCALLE"    ))
OFICINA		= ucase(request.Form("OFICINA"   ))
COMUNA		= ucase(request.Form("COMUNA"    ))
CIUDAD		= ucase(request.Form("CIUDAD"    ))
REGION		= ucase(request.Form("REGION"    ))
ZONA		= ucase(request.Form("ZONA"      ))
ENCARLOCAL	= ucase(request.Form("ENCARLOCAL"))
CARGOENC	= ucase(request.Form("CARGOENC"  ))
FONOENC		= ucase(request.Form("FONOENC"   ))
MOVILENC	= ucase(request.Form("MOVILENC"  ))
MAILENC		= ucase(request.Form("MAILENC"   ))
DIAV		= ucase(request.Form("DIAV"      ))
TIPOV		= ucase(request.Form("TIPOV"     ))
SEMANAV		= ucase(request.Form("SEMANAV"   ))
DIAR		= ucase(request.Form("DIAR"      ))
HAI			= ucase(request.Form("HI"        ))
HAT			= ucase(request.Form("HIT"       ))
HFI			= ucase(request.Form("HF"        ))
HFT			= ucase(request.Form("HFT"       ))
HSI			= ucase(request.Form("HSI"       ))
HST			= ucase(request.Form("HST"       ))
CTACTE		= ucase(request.Form("CTACTE"    ))
TIPO		= ucase(request.Form("TIPO"      ))

HAI=elininadupli(HAI)
HAT=elininadupli(HAT)
HFI=elininadupli(HFI)
HFT=elininadupli(HFT)
HSI=elininadupli(HSI)
HST=elininadupli(HST)

end sub 'datosA
'--------------------------------
function elininadupli(valor)
	r=split(valor,",")
	if ubound(r)<0 then
		elininadupli=""
	else
		elininadupli=trim(r(ubound(r)) )
	end if
end function
'::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::
sub reenviar()
st = request.form("step")
'response.write(st)
tp = request.Form("TIPO")
for each campo in request.Form
'if campo="mipaso" then goto :fineach
if tp = "EDITAR" then
select case st
   case "02"
	  select case campo
	  case "TIPO"
		response.write ""
	  case "CTABANCO"
		response.write ""
	  case "TIT"
		response.write ""
	  case "RUTBANCO"
		response.write ""
	 case "SUCBCO"
		response.write ""
	  case "CIBANCO"
		response.write ""
	  case "BANCO"
		response.write ""
	  case else
		if not campo = "step" then
		with response
		.write "<input type='hidden' name='"
		.write campo 
		.write "' value='"
		.write request.form(campo) & "'>"&chr(13)
		end with
		end if		
	  end select

   case "03"
      select case campo
	      case "SIGLA"
        response.write ""
	      case else
		     	'response.write(campo & "<BR>")
				if not campo = "step" then
				with response
				.write "<input type='hidden' name='"
				.write campo 
				.write "' value='"
				.write request.form(campo) & "'>"&chr(13)
				end with
				end if		  
	   end select  
   case "04"
      select case campo
	      case "CALLE"
		            response.write ""
		  case "NCALLE" 
		            response.write ""
		  case "OCICINA"
		            response.write ""
		  case "COMUNA"
		          response.write ""   
		  case "CIUDAD"  
   		        response.write ""
		  case "REGION" 
		            response.write ""
		  case "ZONA"  
		          response.write ""
		  case "idciudad"
				response.write ""
		  case "idcomuna"
				response.write ""
		  case "mipaso"
				response.write ""
		  case else
				'response.write(campo & "<BR>")
		     	if not campo = "step" then
					if not campo="mipaso" then
						if not campo="idciudad" then
							if not campo="idcomuna" then
								if not campo="idregion" then
				with response
				.write "<input type='hidden' name='"
				.write campo 
				.write "' value='"
				.write request.form(campo) & "'>"&chr(13)
				end with
								end if 
							End if
						End if
					end if
				end if
	  end select
  case "05"
      select case campo
	      case "FONOENC"
		            response.write ""
		  case "ENCARLOCAL" 
		            response.write ""
		  case "MOVILENC"
		            response.write ""
		  case "MAILENC"
		          response.write ""   
		  case "CARGOENC"  
   		        response.write ""
		  case else
		     	if not campo = "step" then
				with response
				.write "<input type='hidden' name='"
				.write campo 
				.write "' value='"
				.write request.form(campo) & "'>"&chr(13)
				end with
				end if
	  end select    
  case "06"
      select case campo
	      case "TIPOV"
		            response.write ""
		  case "DIAV" 
		            response.write ""
		  case "SEMANAV"
		            response.write ""
		  case else
		     	if not campo = "step" then
				with response
				.write "<input type='hidden' name='"
				.write campo 
				.write "' value='"
				.write request.form(campo) & "'>"&chr(13)
				end with
				end if
	  end select    
  case "07"
      select case campo
	      case "DIAR" 'dias entrega
		            response.write ""
		  case "HI" 'ma�ana
		            response.write ""
		  case "HF" 'ma�ana final
		            response.write ""
		  case "HIT" 'tarde
		          response.write ""   
		  case "HIF"  'tarde final
   		        response.write ""
		  case "HSI" 'tarde
		          response.write ""   
		  case "HST"  'tarde final
   		        response.write ""
		  case else
		     	if not campo = "step" then
				with response
				.write "<input type='hidden' name='"
				.write campo 
				.write "' value='"
				.write request.form(campo) & "'>"&chr(13)
				end with
				end if
	  end select    
 case else
		     	if not campo = "step" then
				with response
				.write "<input type='hidden' name='"
				.write campo 
				.write "' value='"
				.write request.form(campo) & "'>"&chr(13)
				end with
				end if
end select
else
	if not campo = "step" then
		if not campo="mipaso" then
			if not campo="idciudad" then
				if not campo="idcomuna" then
					if not campo="idregion" then
						if not campo="idmivalor" then
	with response
	.write "<input type='hidden' name='"
	.write campo 
	.write "' value='"
	.write request.form(campo) & "'>"&chr(13)
	end with
						end if
					end if
				end if
			end if
		end if
	end if
end if
':fineach
next
end sub 'reenviar
'::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::
sub datos_edit(rs)

if isnull(rs.fields("contacto")) then 
	contacto=split("|","|")
else
	contacto=split(rs.fields("contacto") & "|","|")
end if

if isnull(rs.fields("direccion"))then 
	direccion=split("|","|")
else
	direccion=split(rs.fields("direccion") & "|","|")
end if

if isnull(rs.fields("texto2"))   then 
	texto2=split("|||||","|")
else
	texto2=split(rs.fields("texto2"),"|")
end if

if isnull(rs.fields("texto1"))   then 
	texto1=split("|||||","|")
else
	texto1=split(rs.fields("texto1"),"|")
end if

if isnull(rs.fields("texto3"))   then 
	texto3=split("||||||","|")
else
	texto3=split(rs.fields("texto3"),"|")
	'response.write("<HR>" & rs.fields("texto3") & "<HR>")
	if ubound(texto3)<6 then
		texto3=split(rs.fields("texto3") & "|||||||","|")
	end if
end if
if ubound(texto1)<4 then texto1=split(rs.fields("texto1") & "||||","|")
on error resume next
RUT			= rs.fields("codlegal")
RAZSOC		= rs.fields("razonsocial")
GIRO		= rs.fields("analisisctacte1")
CREDITO		= rs.fields("LimiteCredito")
DESCUENTO	= replace(texto2(0),"%", "")
CONDPAGO	= rs.fields("condpago")
BANCO		= rs.fields("analisisctacte7")
CTABANCO	= texto1(0)
TITULAR		= texto1(1)
RUTBCO		= texto1(2)
SUCBCO		= texto1(3)
CIBANCO		= texto1(4)
SIGLA		= rs.fields("sigla")
x = Len(direccion(0))
y = InStr(StrReverse(direccion(0)), " ")
CALLE		= Mid(direccion(0), 1, x - y)
NCALLE		= Right(direccion(0),y)
OFICINA		= direccion(1)
COMUNA 		= rs.fields("comuna")
CIUDAD 		= rs.fields("ciudad")
REGION		= rs.fields("zona")
ZONA 		= rs.fields("codpostal")
ENCARLOCAL	= contacto(0)
CARGOENC	= contacto(1)
FONOENC 	= rs.fields("telefono")
MOVILENC 	= rs.fields("fax")
MAILENC 	= rs.fields("email")
DIAV 		= mid(rs.fields("analisisctacte5"),2,1)
TIPOV		= left(rs.fields("analisisctacte5"),1)
SEMANAV 	= right(rs.fields("analisisctacte5"),1)
DIAR 		= replace(texto3(0),",",", ")
HAI 		= texto3(1)
HAT 		= texto3(2)
HFI 		= texto3(3)
HFT			= texto3(4)
HSI			= texto3(5)
HST			= texto3(6)

call enviocorreo()
end sub 'datos edit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub guardalog()
'archivo = "c:\inetpub\wwwroot\cliente\clienteauto.log"
'set mitexto = createObject("scripting.filesystemobject")
'set mi_rs = mitexto.OpenTextFile(archivo,8,true)
' mi_rs.WriteLine(CTACTE &"|"& RAZSOC &"|"& vendedor &"|"& date &"|")
' mi_rs.close()
' response.write("<p align='center'><br><br>debe actualizar el cliente, presione el boton para actualizar.<br>")
' response.write("<input type='button' value='Actualizar' onclick="&_
' 				 chr(34)&"location='/cliente/subeServer.asp?code="& CTACTE &"&paso=SendData&palm=1'"& chr(34) &">")
end sub 'guardalog()
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub piepag()
'response.write(replace(request.form, "&","<br>"))
%>
</body>
</html>
<%
end sub 'pie pagina
sub enviocorreo()
'ENVIA CORREO POR CREACION	
asunto = "'Nuevo Cliente Creado " & CTACTE & " Por vendedor "  & request.QueryString("nuser") & "-" &  EJECUTIVO &"'"
mensaje = "'El vendedor " & EJECUTIVO & " ha creado el cliente " & CTACTE & "'"
consulta = "'SELECT * FROM handheld.flexline.ctacte WHERE empresa = ''" & EMPRESA & "'' AND tipoctacte = ''CLIENTE'' AND ctacte = ''" & CTACTE  & "'''"


envio="exec msdb.dbo.sp_send_dbmail @profile_name = '"& empresa &"', @recipients = 'cpalma@desa.cl', " & _
"@body = " & asunto & _
", @query = " & consulta &_
", @subject = " & mensaje & _
", @attach_query_result_as_file = 1 ;"

'response.Write(envio)
end sub
%>
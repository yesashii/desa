<%'================================================================================================
'|  prepara | Ver 1.0 | Simon Hernandez
'==================================================================================================

':: CONSTANTES ::
BD_Source = "SQLSERVER"
BD_Base   = "handheld"
BD_User   = "sa"
BD_Passw  = "desakey"
N_Empresa = "Distrubuidora Errazuriz S. A."

':: CONEXION ::
dim oConn, rs, sql
Set oConn = server.createobject("ADODB.Connection")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0
oConn.open "Provider=SQLOLEDB;Data Source=" & BD_Source & ";Initial Catalog=" & BD_Base & ";User Id=" & BD_User & ";Password=" & BD_Passw & ";"
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")

':: Sub Formularios ::
'--------------------------------------------------------------------------------------------------
Sub encabezado(titulo)
	%><TITLE><%=titulo%></TITLE>
	<body topmargin="0"><CENTER>
	<TABLE bgcolor="#666633" width="100%" border="0">
	<TR>
		<TD align="left"   width="10%"><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><B><%=id_vendedor%></B></FONT></TD>
		<TD align="center" width="80%"><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><B><%=left(pc(nb_vendedor,0),30)%></B></FONT></TD>
		<TD align="right"  width="10%" ><FONT SIZE="2" face="Arial" COLOR="#FFFFFF" ><B>
			<A HREF="../Default.asp" style='text-decoration:none; color:#FFFFFF' >[Inicio]</A></B></FONT></TD>
	</TR>
	</TABLE>
	<%
End Sub 'encabezado()
'--------------------------------------------------------------------------------------------------
Sub piedepagina()
	%>
	<FONT SIZE="1" face="verdana" COLOR="#808080"><%=N_Empresa%></FONT>
	</CENTER>
	</body>
	<%
End Sub 'piedepagina()
'--------------------------------------------------------------------------------------------------
Sub seltipoventa() 
	%><TABLE>
	<TR>
		<TD align="center"><FONT SIZE="1">&nbsp;</FONT></TD>
	</TR>
	<TR>
		<TD align="center">

<SCRIPT LANGUAGE="JavaScript">
<!--
	function xseliva(miobjeto){
		//alert(miobjeto.value)
		if (miobjeto.value=='015'){
			document.getElementById('sel_iva').style.display='none';
			document.getElementById('ck_iva').checked=0 ;
		}else{
			document.getElementById('sel_iva').style.display='';
		}

	}
//-->
</SCRIPT>
		<%
		if ucase(empresa)="DESAZOFRI" then
		%>
		<TABLE bgcolor="#E3EAF2">
		<FORM METHOD=POST ACTION="Pedido00.asp?vend=<%=ucase(nb_vendedor)%>&tipPed=S&id_porfolio=<%=id_porfolio%>&empresa=<%=empresa%>">
			<TR>
				<TD><FONT SIZE="2" face="verdana" COLOR="#000000">Tipodocto</FONT>
				<SELECT NAME="idtipodocto" id="idtipodocto" onChange="xseliva(this)">
					<option value="008">008</option>
					<option value="015">015</option>
				</SELECT>
				</TD>
			</TR>
			<TR>
				<TD>
				<TABLE id="sel_iva" name="sel_iva">
				<TR>
					<TD><INPUT TYPE="checkbox" NAME="ck_iva" id="ck_iva"></TD>
					<TD><FONT SIZE="2" face="verdana" COLOR="#000000"><label for="ck_iva">Con I.V.A.</label></FONT></TD>
				</TR>
				</TABLE>
				</TD>
			</TR>
			<TR>
				<TD align="center"><INPUT TYPE="submit" Value="Hacer Pedido" style="width: 100"></TD>
			</TR>
		</FORM>
		</TABLE>
		<%
		else
		':::: bloqueo htc
			sw_htc="N"
			sqlH="select * FROM Handheld.dbo.Dim_vendedores where sw_htc='S' and idempresa=1 and nombre='" & nb_vendedor & "'"
			Set rs=oConn.execute(sqlH)
			if not rs.eof then
				sw_htc="S"'rs("sw_HTC")
			end if

			'if ucase(nb_vendedor)="ALDO LOPEZ" or ucase(nb_vendedor)="MARCELO RUBIO" or ucase(nb_vendedor)="MARCO CONTRERAS" or _
			 '  ucase(nb_vendedor)="PATRICIO UMAÑA" or ucase(nb_vendedor)="" then 
			if sw_htc="S" then
				%>
				<TABLE bgcolor="#E4E4E4">
				<TR>
					<TD>&nbsp;Sistema bloqueado, <BR>&nbsp;solo ingresar pedidos por HTC&nbsp;</TD>
				</TR>
				</TABLE>
				<%
			else
			%>
			Vendedores DESA (RUTA): el sistema de ventas sera bloqueado el 08/05/2011, para ingreso de pedidos via web, solo se permitiran ingresos via HTC
			<FORM METHOD=POST ACTION="Pedido00.asp?vend=<%=ucase(nb_vendedor)%>&tipPed=S&id_porfolio=<%=id_porfolio%>&empresa=<%=empresa%>">
				<INPUT TYPE="submit" Value="Hacer Pedido" style="width: 100">
			</FORM>
			<%
			end if
		end if
		%>
		</TD>
	</TR>
	<TR>
		<TD align="center"><FONT SIZE="1">&nbsp;</FONT></TD>
	</TR>
	<TR>
		<TD align="center">
		<FORM METHOD=POST ACTION="Pedido00.asp?vend=<%=ucase(nb_vendedor)%>&tipPed=N&id_porfolio=<%=id_porfolio%>&empresa=<%=empresa%>">
			<INPUT TYPE="submit" 
				Value=" No Venta " 
			style="width: 100">
		</FORM>
		</TD>
	</TR>
	<TR>
		<TD align="center"><FONT SIZE="1">&nbsp;</FONT></TD>
	</TR>
	<TR>
		<TD align="center">
		<FORM METHOD=POST ACTION="../Default.asp">
			<INPUT TYPE="submit" 
				Value=" [ << ] . Salir " 
			style="width: 100">
		</FORM>
		</TD>
	</TR>
	<TR>
		<TD align="center"><!-- <INPUT TYPE="text" NAME=""> --></TD>
	</TR>
	</TABLE>
	<%
End Sub 'seltipoventa() 
'--------------------------------------------------------------------------------------------------
Sub listaclientes()
	SQL="SELECT *, rutflex as 'CodLegal', getdate() as fecha " & _
	"FROM handheld.Flexline.PDA_RUTA_HOY " & _
	"WHERE (Vendedor = '" & nb_vendedor & "')"
	set rs=oConn.execute(sql)
	if not rs.eof then
		codsem=rs("codsem")
		coddia=rs("coddia")
	end if
	
	if coddia=1 then nomdia="Lunes"
	if coddia=2 then nomdia="Martes"
	if coddia=3 then nomdia="Miercol"
	if coddia=4 then nomdia="Jueves"
	if coddia=5 then nomdia="Viernes"
	if coddia=6 then nomdia="Sabado"
	if coddia=7 then nomdia="Domingo"
	titulo="Lista Clientes : " & nomdia & " Semana " & codsem
	%>
	<FORM METHOD=POST ACTION="">
	<TABLE>
	<TR>
		<TD><FONT SIZE="2" face="verdana" COLOR="#000040"><B>Busqueda</B></FONT></TD>
		<TD><INPUT TYPE="text" NAME="busca" size="20"></TD>
		<TD><INPUT TYPE="submit" value="Buscar"></TD>
	</TR>
	</TABLE>
	</FORM>
	<TABLE border="0" cellspacing='0' bgcolor="#BEC0D8" width="100%">
	<TR bgcolor="#000066">
		<TD colspan="2" align="center"><FONT SIZE="2" face="arial" COLOR="#FFFFFF"><B><%=titulo%></B></FONT></TD>
	</TR>
	<!-- Lineas Clientes -->
	<% x=1
	do until rs.eof 
	bgcolor=iif(x mod 2=0,"#C0C0C0","#BEC0D8") %>
	<TR bgcolor="<%=bgcolor%>">
		<TD width="8%" align="center">
			<INPUT TYPE="radio" NAME="1" onClick="selcliente(<%=x%>)">
		</TD>
		<TD width="92
		%">
			<INPUT TYPE="text" size="34" NAME="" value="<%=pc(rs(1),0)%>" disabled >
		</TD>
	</TR>
	<%
	rs.movenext
	x=x+1
	loop
	%></TABLE>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	//-------------------------------------------
	function selcliente(id_cliente){
		id_cliente--
		msg='Cliente : ' + Matriz2D[id_cliente][0] + '\n';
		msg=msg + Matriz2D[id_cliente][1] + '\n';
		msg=msg + 'Direccion :' + '\n';
		a=confirm(msg);
		if (a=true){
			alert(a);
		}
	}
	//--------------------------------------------
	//-->
	</SCRIPT>
	<%
	rs.movefirst
	filas=x
	columnas=rs.fields.count
	response.write("<SCRIPT LANGUAGE=" & chr(34) & "JavaScript" & chr(34) & " TYPE=" & chr(34) & "text/javascript" & chr(34) & "> " & chr(10) )

	response.write("var Matriz2D=new Array(" & i & ");" & chr(10) )
	response.write("for(i=0;i<=" & filas & ";i++){Matriz2D[i]=new Array(" & columnas & ");}" & chr(10) )

	a=0
	do until rs.eof
		for e=0 to (columnas-1)
			response.write("Matriz2D[" & a & "][" & e & "]='" & pc(rs.fields(e),0) & "';" & chr(10) )
		next
	a=a+1
	rs.movenext
	loop
	response.write("</SCRIPT>" & chr(10) )

End Sub 'listaclientes()
'--------------------------------------------------------------------------------------------------
Sub infocliente()  
End Sub 'infocliente()  
'--------------------------------------------------------------------------------------------------


':: Funciones ::
'--------------------------------------------------------------------------------------------------
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
PC=replace(replace(PC,"ñ","n"),"Ñ","N")
End Function 
'--------------------------------------------------------------------------------------------------
Function cisnull(valor, reemplazo)
	if isnull(valor) then
		cisnull=reemplazo
	else
		cisnull=valor
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function iif(consulta, valortrue, valorfalse)
	if consulta then
		iif=valortrue
	else
		iff=valorfalse
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function recuperavalor(objeto)
	if len(request.querystring(objeto))=0 then
		if len(request.form(objeto))>0 then
			recuperavalor=request.form(objeto)
		else
			recuperavalor=""
		end if
	else
		recuperavalor=request.querystring(objeto)
	end if
End Function
'--------------------------------------------------------------------------------------------------
function flabel(texto, Destino, Formato)
	'=====================================================================
	'| FLablel : Funcion constructora de Labels  | Ver 1.0 -  26/08/2008 |
	'=====================================================================
	'		Formato - Numerico = Tamaño de letra						|
	'               - Explicito : Font/style							|
	'		Destino = For												|
	'-------------------------------------------------------------------
	lformato=""
	if isnumeric(formato) then 
		ltamano=cint(formato)
		lformato=" face='verdana' SIZE='" & ltamano & "' COLOR='#000033' "
	else
		lformato=formato
	end if
	%><label id="lb_<%=Destino%>" name="lb_<%=Destino%>" for="<%=Destino%>">
	<FONT <%=lformato%>><%=pc(texto,0)%></FONT>
	</label><%
end function 'flabel
'--------------------------------------------------------------------------------------------------
function milistbox(sql, nombre, default)
	'=====================================================================
	'|  milistbox ver : 1.0 - 05 06 2008								|
	'=====================================================================
	' SQL de 1 Columna , Valor=Col(1); Texto=Col(1)
	' SQL de 2 Columnas, Valor=Col(1); Texto=Col(2)
nombre=replace(trim(nombre)," ","_")
i=0

set frs=oConn.execute(sql)
if frs.eof then
	%><SELECT NAME="<%=nombre%>" id="<%=nombre%>" onChange="textbox_onChange('<%=nombre%>')">
			<option >EOF: La Consulta SQL no Retorno ningun Resultado</option>
	</SELECT><%
'exit function 
else
if frs.fields.count >1 then i=1
	%><SELECT NAME="<%=nombre%>" id="<%=nombre%>" onChange="textbox_onChange('<%=nombre%>')">
		<%
		do until frs.eof 
			sel=""
			if lcase(frs.fields(0))=lcase(default) then sel=" selected "
			valor0=frs.fields(0)
			valor1=pc(frs.fields(i),0)
			if lcase(trim(valor0))="todos" then valor0=""
			%>
			<option <%=sel%>value="<%=valor0%>"><%=valor1%></option>
		<% frs.movenext 
		loop %>
	</SELECT><%
end if
End function 'milistbox
'--------------------------------------------------------------------------------------------
function consultarapida(sql)
	set frs=oConn.execute(sql)
	if not frs.eof then consultarapida=frs(0)
	consultarapida=cisnull(consultarapida,"Sin datos")
End function 'consultarapida
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
%>
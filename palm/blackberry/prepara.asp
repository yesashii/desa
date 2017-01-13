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
Sub encabezado()
	%><TITLE>Blackberry</TITLE>
	<body topmargin="0" bgcolor="#C0C0C0"><CENTER>
	<TABLE bgcolor="#000066" width="100%" border="0">
	<TR>
		<TD align="left"   width="10%"><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><B><%=id_vendedor%></B></FONT></TD>
		<TD align="center" width="80%"><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><B><%=left(pc(nb_vendedor,0),30)%></B></FONT></TD>
		<TD align="right"  width="10%"><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><B>H</B></FONT></TD>
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
	%><FORM METHOD=POST ACTION="pedido.asp"><INPUT TYPE="hidden" name="paso" value="1">
	<TABLE>
	<TR>
		<TD align="center"><FONT SIZE="1">&nbsp;</FONT></TD>
	</TR>
	<TR>
		<TD align="center">
		<INPUT TYPE="submit" value="Hacer Pedido" style="width: 100">
		</TD>
	</TR>
	<TR>
		<TD align="center"><FONT SIZE="1">&nbsp;</FONT></TD>
	</TR>
	<TR>
		<TD align="center">
		<INPUT TYPE="submit" value=" No Ventas . " style="width: 100">
		</TD>
	</TR>
	<TR>
		<TD align="center"><FONT SIZE="1">&nbsp;</FONT></TD>
	</TR>
	<TR>
		<TD align="center">
		<INPUT TYPE="button" VAlue=" [ << ] . Salir " onClick="history.back()" style="width: 100">
		</TD>
	</TR>
	<TR>
		<TD align="center"><!-- <INPUT TYPE="text" NAME=""> --></TD>
	</TR>
	</TABLE>
	
	</FORM><%
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
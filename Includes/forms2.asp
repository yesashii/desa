<%
':milistbox
':flabel
'----------------------------------------------------------------------------------------------------
function milistbox(sql, nombre, default)
	'=====================================================================
	'|  milistbox ver : 1.0 - 05 06 2008								|
	'=====================================================================
	' SQL de 1 Columna , Valor=Col(1); Texto=Col(1)
	' SQL de 2 Columnas, Valor=Col(1); Texto=Col(2)
nombre=replace(trim(nombre)," ","_")
i=0
set frs=oConn.execute(Sql)
'if rs.eof exit function 
if frs.fields.count >1 then i=1
	%><SELECT NAME="<%=nombre%>" id="<%=nombre%>" onChange="textbox_onChange('<%=nombre%>')" style="font-size:10px">
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
End function 'milistbox
'----------------------------------------------------------------------------------------------------
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
'----------------------------------------------------------------------------------------------------
%>
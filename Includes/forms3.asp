<%
'----------------------------------------------------------------------------------------------------
function flabel(texto, Destino, Formato)
'------------------------------------------------------
'| FLablel : Funcion constructora de Labels           |
'------------------------------------------------------
'		Formato - Numerico = Tama�o de letra          |
'               - Explicito : Font/style              |
'------------------------------------------------------
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
function pc(texto,n)
	pc=texto
end function

call flabel("Etiqueta prueba","demo","2")
%>
<INPUT TYPE="checkbox" NAME="demo" id="demo">
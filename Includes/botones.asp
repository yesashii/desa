<script type="text/javascript" src="/includes/glossy.js"></script>
<%
'-------------------------------------------------------------------------------------------
sub miboton(boton,formato)
	boton=lcase(trim(boton))
	if not isnumeric(formato) then formato=0

select Case (formato)
case 0:'botones
	select Case (boton)
	Case "print":
		%><input type="button" value="Imprimir" onClick="window.print()"><%
	Case "back":
		%><input type="button" value="<< Volver" onClick="history.back()"><%
	Case "3":

	End Select 'boton

case 1:'imagenes
	select Case (boton)
	Case "print":
		%><IMG SRC="/includes/print.GIF" WIDTH="48" HEIGHT="48" BORDER="0" ALT="Imprimir" class="glossy iradius50"  onClick="window.print()"><%
	Case "back":
		%><IMG SRC="/includes/back.GIF" WIDTH="48" HEIGHT="48" BORDER="0" ALT="Imprimir" class="glossy iradius50"  onClick="history.back()"><%
	Case "buscar":
		%><IMG SRC="/includes/buscar.GIF" WIDTH="48" HEIGHT="48" BORDER="0" ALT="Imprimir" class="glossy iradius50"  onClick="history.back()"><%
	Case "info":
		%><IMG SRC="/includes/info.GIF" WIDTH="48" HEIGHT="48" BORDER="0" ALT="Imprimir" class="glossy iradius50"  onClick="history.back()"><%
	Case "camion":
		%><IMG SRC="/includes/camion.GIF" WIDTH="48" HEIGHT="48" BORDER="0" ALT="Imprimir" class="glossy iradius50"  onClick="history.back()"><%
	End Select 'boton

End Select 'formato



end sub 'miboton()
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------

%>
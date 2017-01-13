<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Orden de Trabajo Unidades de Frio</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY topmargin="0" bgcolor="#C0C0C0" >
<CENTER>
	<FONT size="3" face="arial" COLOR="#FFFFFF">
		<B>ORDEN DE TRABAJO UNIDADES DE FRIO</B>
	&nbsp;&nbsp;&nbsp;&nbsp;Folio:
	</FONT>
<HR>
<FORM METHOD=POST ACTION="">
<TABLE border="0" cellspacing="0" cellpadding="0">
<TR bgcolor="#DFDFDF">
	<TD>Tipo&nbsp;de&nbsp;trabajo</TD>
	<TD>
		<INPUT TYPE="radio" NAME="">Instalacion<BR>
		<INPUT TYPE="radio" NAME="">Serv&nbsp;Tecnico<BR>
		<INPUT TYPE="radio" NAME="">Retiro
	</TD>
</TR>
<TR>
	<TD>Fecha&nbsp;Orden</TD>
	<TD><%=date()%></TD>
</TR>
<TR bgcolor="#DFDFDF">
	<TD>Tipo Cooler</TD>
	<TD>
		<SELECT NAME="tipocooler">
		</SELECT>
	</TD>
</TR>
<TR>
	<TD>Numero&nbsp;Serie</TD>
	<TD><INPUT TYPE="text" NAME="serie"></TD>
</TR>
<TR bgcolor="#DFDFDF">
	<TD>Marca&nbsp;Asociada</TD>
	<TD>
		<SELECT NAME="marca">
		</SELECT>
	</TD>
</TR>
<TR>
	<TD>Motivo<BR>visita</TD>
	<TD>
		<TEXTAREA NAME="" ROWS="" COLS="">
		</TEXTAREA>
	</TD>
</TR>
<TR bgcolor="#DFDFDF">
	<TD>Dia<BR>Preferente<BR>Visita</TD>
	<TD>
		<INPUT TYPE="checkbox" NAME="">Lunes<BR>
		<INPUT TYPE="checkbox" NAME="">Martes<BR>
		<INPUT TYPE="checkbox" NAME="">Miercoles<BR>
		<INPUT TYPE="checkbox" NAME="">Jueves<BR>
		<INPUT TYPE="checkbox" NAME="">Viernes
	</TD>
</TR>
<TR>
	<TD>Hora Visita</TD>
	<TD>
		<SELECT NAME="horavisita">
		</SELECT>
	</TD>
</TR>
<TR bgcolor="#DFDFDF">
	<TD>Observacion</TD>
	<TD>
		<TEXTAREA NAME="" ROWS="" COLS="">
		</TEXTAREA>
	</TD>
</TR>
<TR>
	<TD>Nombre<BR>contacto<BR>local</TD>
	<TD><INPUT TYPE="text" NAME="contacto"></TD>
</TR>
<TR bgcolor="#DFDFDF">
	<TD>fono<BR>contacto<BR>local</TD>
	<TD><INPUT TYPE="text" NAME="fonocontacto"></TD>
</TR>
</TABLE>
<INPUT TYPE="submit" value="Guardar">&nbsp;
<INPUT TYPE="reset" value="borrar">&nbsp;
<INPUT TYPE="button" value="<<Menu">
</FORM>
</CENTER>
</BODY>
</HTML>

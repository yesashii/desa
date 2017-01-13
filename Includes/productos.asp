<!--#include Virtual="/includes/conexion.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Transformadores</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">

 <style type="text/css">
  body {
    /*color: purple;
    background-color: #d8da3d */
	font-family: Arial;
	font-size: 10;
	}
	td{
	font-size: 12;
	color: #336699;
	}
	td.ck{
	font-size: 12;
	color: #575757;
	font-weight : bold;
	}
	td.nm{
	font-size: 12;
	color: #575757;
	/*font-weight : bold;*/
	}
	td.tl{
	font-size: 14;
	color: #003366;
	font-weight : bold;
	}
  </style>
</HEAD>
<%
'productosel="Poder JCM 50W"
productosel=recuperavalor("idproducto")
%>

<BODY topmargin="0" bgcolor="#336699">
<CENTER>
<TABLE border="0" height="100%"  bgcolor="#FFFFFF" style="border-collapse: collapse;">
<TR bgcolor="#336699" height="100px" >
	<TD background="img/borde1.PNG"><IMG SRC="img/borde1.PNG" WIDTH="13" HEIGHT="1" BORDER="0" ALT=""></TD>
	<TD colspan="2"><IMG SRC="img/logo2.JPG" WIDTH="800" HEIGHT="100" BORDER="0" ALT=""></TD>
	<TD background="img/borde2.PNG"><IMG SRC="img/borde2.PNG" WIDTH="11" HEIGHT="1" BORDER="0" ALT=""></TD>
</TR>
<TR>
	<TD background="IMG/borde1.PNG"></TD>
	<TD width="160px"
	valign="top"
	bordercolor="#808080"
	style="border-right: solid;
	border-right-width: 2px;
	" >
	<TABLE width="100%" border="0" style="border-collapse: collapse;">
	<%
	familiaant=""
	sql="SELECT * " & _
	"FROM old.dbo.web_productos " & _
	"ORDER BY familia, producto"
	set rs=oconn.execute(sql)
	Do until rs.eof
		if len(productosel)=0 then productosel=rs("producto")
		if familiaant<>rs("familia") then
		%>
			<TR>
				<TD>&nbsp;<%=rs("familia")%></TD>
			</TR>
		<%
		end if
	if productosel=rs("producto") then
		precio =rs("precio")
		milink =split(rs("link"),";")
		esquema=rs("esquema")
		descrip=rs("descripcion")
		imagens=rs("imagenes")
		imgs=split(imagens,";")
	%>
	<TR>
		<TD bgcolor="#D1D9E7">
			<TABLE border="0" style="border-collapse: collapse;">
			<TR>
				<TD><IMG SRC="img/<%=rs("icono")%>" WIDTH="40" HEIGHT="33" BORDER="0" ALT=""></TD>
				<TD class="ck"><%=rs("producto")%></TD>
			</TR>
			</TABLE>
		</TD>
	</TR>
	<%
	else
	%>
	<TR>
		<TD onmouseover="this.style.color='red'; 
				this.style.background='#EAECF4'; 
				this.style.cursor='hand'"
			onmouseout="this.style.color='black'; 
				this.style.background='transparent';"
			onclick="window.open('productos.asp?idproducto=<%=rs("producto")%>','_self')">
			<TABLE border="0" style="border-collapse: collapse;">
			<TR>
				<TD><IMG SRC="img/<%=rs("icono")%>" WIDTH="40" HEIGHT="33" BORDER="0" ALT=""></TD>
				<TD class="nm"><%=rs("producto")%></TD>
			</TR>
			</TABLE>
		</TD>
	</TR>
	<%
	end if

	familiaant=rs("familia")
	rs.movenext
	loop
	%>
	</TABLE>
	</TD>
	<TD width="640px" valign="top">
<!-- incio cuerpo -->
<TABLE width="500px">
<TR>
	<TD class="tl"><%=productosel%></TD>
</TR>
<TR>
	<TD><IMG id="img_1" name="img_1" SRC="img/<%=imgs(0)%>" WIDTH="323" HEIGHT="250" BORDER="1" ALT=""
	onclick="document.getElementById('img_2').src=this.src;
	         document.getElementById('div_img').style.display=''"
	></TD>
</TR>
<TR>
	<TD>
	<TABLE>
	<TR>
	<%
	for x=0 to ubound(imgs)
		%>
		<TD><IMG SRC="img/<%=imgs(x)%>" WIDTH="64" HEIGHT="48" BORDER="1" ALT=""
			style="opacity:0.4;filter:alpha(opacity=40)"
		onmouseover="this.style.opacity=1;this.filters.alpha.opacity=100;
		this.style.cursor='hand'"
		onmouseout="this.style.opacity=0.4;this.filters.alpha.opacity=40"
		onclick="document.getElementById('img_1').src=this.src"></TD>
		<%
	next
	%>
		</TR>
	</TABLE>
	</TD>
</TR>
<TR>
	<TD>Precio : $<B><%=formatnumber(precio,0)%></B><BR><BR></TD>
</TR>
<TR>
	<TD><%=milink(0)%>&nbsp;<B><A HREF="<%=milink(1)%>">Link</A></B></TD>
</TR>
<TR>
	<TD><IMG SRC="img/<%=esquema%>" WIDTH="299" HEIGHT="201" BORDER="0" ALT=""></TD>
</TR>
<TR>
	<TD><%=descrip%></TD>
</TR>
</TABLE>
<!-- fin cuerpo -->	
	</TD>
	<TD background="img/borde2.PNG"></TD>
</TR>
<TR height="100px">
	<TD background="img/borde1.PNG"></TD>
	<TD colspan="2">
<TABLE>
	<TR>
		<TD><IMG SRC="img/logo_phpBB.gif" WIDTH="360" HEIGHT="100" BORDER="0" ALT=""></TD>
		<TD><IMG SRC="img/kowka.jpg" WIDTH="400" HEIGHT="100" BORDER="0" ALT=""></TD>
	</TR>
</TABLE>
	</TD>
	<TD background="img/borde2.PNG"></TD>
</TR>
</TABLE>
</CENTER>

<div id="div_img" name="div_img"
style="position:absolute; text-align:center; left:0px; top:100px; display:none;">
<TABLE height="100%" width="100%" bgcolor="#336699">
<TR>
	<TD align="center" valign="center">
<BR>
<BR>
	<TABLE style="border-collapse: collapse;">
		<TR>
			<!-- <TD background="img/borde1.PNG"><IMG SRC="img/borde1.PNG" WIDTH="13" HEIGHT="1" BORDER="0" ALT=""></TD> -->
			<TD onclick="document.getElementById('div_img').style.display='none'" align="right"><B><FONT SIZE="arial" COLOR="#FFFFFF">[x] Cerrar</FONT></B></TD>
			<!-- <TD background="img/borde2.PNG"><IMG SRC="img/borde2.PNG" WIDTH="11" HEIGHT="1" BORDER="0" ALT=""></TD> -->
		</TR>
		<TR>
			<!-- <TD background="img/borde1.PNG"><IMG SRC="img/borde1.PNG" WIDTH="13" HEIGHT="1" BORDER="0" ALT=""></TD> -->
			<TD onclick="document.getElementById('div_img').style.display='none'"  align="center"><IMG id="img_2" name="img_2" SRC="" BORDER="2" ALT=""></TD>
			<!-- <TD background="img/borde2.PNG"><IMG SRC="img/borde2.PNG" WIDTH="11" HEIGHT="1" BORDER="0" ALT=""></TD> -->
		</TR>
	</TABLE>
<BR>
<BR>
<BR>
<BR>
<BR>
<BR>
<BR>
</TD>
</TR>
</TABLE>
</div>

</BODY>
</HTML>




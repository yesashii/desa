<!--#include Virtual="/includes/conexion.asp"-->
<%

empresa   =request.querystring("empresa")  'EFECTIVO CAMION 
CONDPAGO  =request.querystring("CONDPAGO")  'EFECTIVO CAMION 
CIUDAD    =request.querystring("CIUDAD")    'ALGARROBO 
letra     =request.querystring("letra")     ' 
CTABANCO  =request.querystring("CTABANCO")  ' 
MOVILENC  =request.querystring("MOVILENC")  ' 
MAILENC   =request.querystring("MAILENC")   ' 
HSI       =request.querystring("HSI")       '08:00 
HST       =request.querystring("HST")       '18:00 
RUT       =request.querystring("RUT")       '65465456-5 
DESCUENTO =request.querystring("DESCUENTO") '11 
RUTBCO    =request.querystring("RUTBCO")    '65465456-5 
ZONA      =request.querystring("ZONA")      '0-0 
FONOENC   =request.querystring("FONOENC")   'fgh 
RAZON     =request.querystring("RAZON")     'fghfghf 
CIBANCO   =request.querystring("CIBANCO")   ' 
checktipe =request.querystring("checktipe") 'NUEVO 
GIRO      =request.querystring("GIRO")      '02 FARMACIAS 
TIT       =request.querystring("TIT")       'FGHFGHF 
ENCARLOCAL=request.querystring("CONTACTO")'yfgh 
TIPOV     =request.querystring("TIPOV")     'A 
HI        =request.querystring("HI")        '08:00 
CREDITO   =request.querystring("CREDITO")   '500.000 
SUCBCO    =request.querystring("SUCBCO")    ' 
CALLE     =request.querystring("CALLE")     'fghf 
REGION    =request.querystring("REGION")    'QUINTA REGION 
BANCO     =request.querystring("BANCO")     'SECURITY 
OBS       =request.querystring("comentario1")       ' 
TIPO      =request.querystring("TIPO")      'NUEVO 
CTACTE    =request.querystring("CTACTE")    '65465456-5 1 
NCALLE    =request.querystring("NCALLE")    '4564 
OFICINA   =request.querystring("OFICINA")   ' 
CARGOENC  =request.querystring("CARGOENC")  ' 
DIAV      =request.querystring("DIAV")      '1 
HF        =request.querystring("HF")        '08:00 
SIGLA     =request.querystring("SIGLA")     'fghgfh 
tipocalle =request.querystring("tipocalle") 'Calle 
DIAR      =request.querystring("DIAR")      'L, M, R, J, V, S 
COMUNA    =request.querystring("COMUNA")    'ALGARROBO 
HIT       =request.querystring("HIT")       '12:00 
stepfrm   =request.querystring("step")      '09 
SEMANAV   =request.querystring("SEMANAV")   '1 
HFT       =request.querystring("HFT")       '12:00 
'EJECUTIVO =request.querystring("EJECUTIVO") '
IDVENDEDOR=request.querystring("IDVENDEDOR") 

direccion    = UCASE(CALLE & " " & NCALLE & " | " & OFICINA)
Limitecredito= replace(replace(cstr(CREDITO),".",""),",","")
RAZON        = ucase(RAZON)
SIGLA        = ucase(SIGLA)

sql="Flexline.GUARDA_CLIENTE @empresa='" & empresa & "', @ctacte='" & trim(CTACTE) & "', @codlegal='" & RUT & "', @razonsocial='" & RAZON & "', @Sigla='" & SIGLA & "',@IDVENDEDOR='" & IDVENDEDOR & "', @CondPago='" & condpago & "', @Zona='" & ZONA & "', @Direccion='" & direccion & "', @Ciudad='" & CIUDAD & "', @Comuna='" & COMUNA & "', @Telefono='" & FONOENC & "', @eMail='" & MAILENC & "', @Contacto='" & ENCARLOCAL & "', @DireccionEnvio='" & direccion & " . " & "', @LimiteCredito='" & Limitecredito & "', @Comentario1='" & OBS & "', @Texto1='', @Texto2='', @Texto3='', @Analisisctacte1='" & GIRO & "',  @Analisisctacte4='" & DESCUENTO & "', @Analisisctacte7='" & BANCO & "'"

'response.write("<BR>" & sql)

set rs=oConn.execute(sql)

%>
<BR>
<%=empresa%>
<BR>
<CENTER>
<FONT SIZE="2" face="verdana" COLOR="#000066"><B>Cliente Creado con Exito </B></FONT>
<%
sql="SELECT * FROM handheld.Flexline.PDA_usuarios WHERE (Num_vend = N'" & IDVENDEDOR & "')"
set rs=oConn.execute(sql)
%>
<FORM METHOD=POST ACTION="../Default.asp">
<INPUT TYPE="hidden" name="empresa"  value="<%=rs("empresa")%>">
<INPUT TYPE="hidden" name="usuario"  value="<%=rs("usuario")%>">
<INPUT TYPE="hidden" name="password" value="<%=rs("password")%>">
<INPUT TYPE="hidden" name="paso"     value="1">
<INPUT TYPE="submit" value="Aceptar">
</FORM>
</CENTER>
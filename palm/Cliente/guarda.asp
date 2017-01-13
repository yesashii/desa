<!--#include Virtual="/includes/conexion.asp"-->
<%

empresa   =request.form("empresa")  'EFECTIVO CAMION 
CONDPAGO  =request.form("CONDPAGO")  'EFECTIVO CAMION 
CIUDAD    =request.form("CIUDAD")    'ALGARROBO 
letra     =request.form("letra")     ' 
CTABANCO  =request.form("CTABANCO")  ' 
MOVILENC  =request.form("MOVILENC")  ' 
MAILENC   =request.form("MAILENC")   ' 
HSI       =request.form("HSI")       '08:00 
HST       =request.form("HST")       '18:00 
RUT       =request.form("RUT")       '65465456-5 
DESCUENTO =request.form("DESCUENTO") '11 
RUTBCO    =request.form("RUTBCO")    '65465456-5 
ZONA      =request.form("ZONA")      '0-0 
FONOENC   =request.form("FONOENC")   'fgh 
RAZON     =request.form("RAZON")     'fghfghf 
CIBANCO   =request.form("CIBANCO")   ' 
checktipe =request.form("checktipe") 'NUEVO 
GIRO      =request.form("GIRO")      '02 FARMACIAS 
TIT       =request.form("TIT")       'FGHFGHF 
ENCARLOCAL=request.form("ENCARLOCAL")'yfgh 
TIPOV     =request.form("TIPOV")     'A 
HI        =request.form("HI")        '08:00 
CREDITO   =request.form("CREDITO")   '500.000 
SUCBCO    =request.form("SUCBCO")    ' 
CALLE     =request.form("CALLE")     'fghf 
REGION    =request.form("REGION")    'QUINTA REGION 
BANCO     =request.form("BANCO")     'SECURITY 
OBS       =request.form("OBS")       ' 
TIPO      =request.form("TIPO")      'NUEVO 
CTACTE    =request.form("CTACTE")    '65465456-5 1 
NCALLE    =request.form("NCALLE")    '4564 
OFICINA   =request.form("OFICINA")   ' 
CARGOENC  =request.form("CARGOENC")  ' 
DIAV      =request.form("DIAV")      '1 
HF        =request.form("HF")        '08:00 
SIGLA     =request.form("SIGLA")     'fghgfh 
tipocalle =request.form("tipocalle") 'Calle 
DIAR      =request.form("DIAR")      'L, M, R, J, V, S 
COMUNA    =request.form("COMUNA")    'ALGARROBO 
HIT       =request.form("HIT")       '12:00 
stepfrm   =request.form("step")      '09 
SEMANAV   =request.form("SEMANAV")   '1 
HFT       =request.form("HFT")       '12:00 
'EJECUTIVO =request.form("EJECUTIVO") '
IDVENDEDOR=request.form("IDVENDEDOR") 

if len(request.form("IDVENDEDOR") < 1 then
IDVENDEDOR=request.querystring("nuser") 
end if

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
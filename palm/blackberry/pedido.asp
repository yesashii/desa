<%'=====================================================================
'|  Sistema preventa BlackBerry | Ver 1.0 | Simon Hernandez
'=======================================================================
':: includes ::
%><!--#include file="prepara.asp" //--><%
'on error resume next
':: Variables ::
public id_vendedor, id_producto, nro_producto, nb_vendedor
id_vendedor="28"
sql="select nombre from handheld.flexline.fx_vende_man where numero=" & id_vendedor
nb_vendedor=consultarapida(sql)
paso=request.form("paso")


call encabezado()

if paso=0 then call seltipoventa()   'prepara.asp
if paso=1 then call listaclientes()  'prepara.asp
if paso=2 then call infocliente()    'prepara.asp
'if paso=3 then call productos()      'producto.asp
'if paso=4 then call resumenpedido()  'producto.asp
'if paso=5 then call guardar()        '

call piedepagina()
%>

<% ':::: conexion :::::::::::::::::
public oConn, rs
Set oConn = server.createobject("ADODB.Connection")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0
oConn.open "DRIVER={HDBODBC};SERVERNODE=HANAB1:30015;DATABASE=HIS_DESA;UID=SYSTEM;PWD=Passw0rd"
'--------------------------------------------------------------------------------------------
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
End Function 
'--------------------------------------------------------------------------------------------
function recuperavalor(mi_elemento)
	if len(mi_valor)=0 then mi_valor=trim(request.querystring(mi_elemento))
	if len(mi_valor)=0 then mi_valor=trim( cstr( request.form(mi_elemento)))
	if len(mi_valor)=0 then mi_valor=""
	recuperavalor=mi_valor
end function
'--------------------------------------------------------------------------------------------
function consultarapida(mi_sql)
	'misql="select idvendedor from SQLSERVER.desaerp.dbo.DIM_VENDEDORES where nombre='" & vend & "'"
	set rs1=oConn.execute(mi_sql)
	if not rs1.eof then consultarapida  = rs1.fields(0)
end function
'--------------------------------------------------------------------------------------------
function combertfecha(xyfecha,xyformato)
	xyformato = ucase(xyformato)
	xyYYYY    = cstr(year(xyfecha))
	xyYY      = right(xyYYYY,2)
	xyMM      = right("00" & month(xyfecha),2)
	xyDD      = right("00" & day(xyfecha),2)
	combertfecha=replace(replace(replace(xyformato,"YYYY",xyYYYY),"MM",xyMM),"DD",xyDD)
End function
'--------------------------------------------------------------------------------------------
function Jmatrix(jnombre, jsql)
	set rs1=oConn.execute(jsql)
	i=0
	do until rs1.eof
		i=i+1
		rs1.movenext
	loop
	if i>0 then rs1.moveFirst

	response.write("<SCRIPT LANGUAGE=" & chr(34) & "JavaScript" & chr(34) & " TYPE=" & chr(34) & "text/javascript" & chr(34) & "> " & chr(10) )
	response.write("var " & jnombre & "=new Array(" & i & ");" & chr(10) )
	response.write("for(i=0;i<=" & i & ";i++){" & jnombre & "[i]=new Array(" & rs1.fields.count & ");}" & chr(10) )
	a=0
	do until rs1.eof
		for x1=0 to (rs1.fields.count-1)
		response.write(jnombre & "[" & a & "][" & x1 & "]='" & rs1.fields(x1) & "';" & chr(10) )
		'response.write(jnombre & "[" & a & "][1]='" & rs1.fields(0) & "';" & chr(10) )
		'response.write(jnombre & "[" & a & "][2]='" & rs1.fields(0) & "';" & chr(10) )
		'response.write(jnombre & "[" & a & "][3]='" & rs1.fields(0) & "';" & chr(10) )
		next
	a=a+1
	rs1.movenext
	loop
	response.write("</SCRIPT>" & chr(10) )
End function
'--------------------------------------------------------------------------------------------
%>
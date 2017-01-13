<%
'-------------------------------------------------------------------------------------------
sub encabezado(nuser, tipocab)
nuser=trim(nuser)
nvendedor=""
sql="select nombre from sqlserver.desaerp.dbo.seg_usuarios as U where idusuario='" & nuser & "'"
Set rs=oConn.execute(Sql)
if not rs.eof then nvendedor=rs.fields(0)
rs.close	
%><TABLE width="100%" border='0' cellpadding='0' cellspacing='0'>
<TR bgcolor="#000066">
	<TD align="center"><B><FONT SIZE="1" face="Arial" COLOR="#C0C0C0"><%=pc(nuser,0)    %></B></FONT></TD>
	<TD align="center"><B><FONT SIZE="1" face="Arial" COLOR="#FFFFFF"><%=pc(nvendedor,0)%></B></FONT></TD>
	<TD align="center"><B><FONT SIZE="1" face="Arial" COLOR="#C0C0C0"><%=pc(tipocab,0)  %></B></FONT></TD>
</TR>
</TABLE><%
End sub 'encabezado(tipocab)
'-------------------------------------------------------------------------------------------
Function mf(sString, mzise, mcolor) 
	fcolor="#000000" 'default Negro
	if mcolor=1 then fcolor="#FFFFFF" 'blanco
	if mcolor=2 then fcolor="#C0C0C0" 'gris claro
	if mcolor=3 then fcolor="#808080" 'plomo
	if mcolor=4 then fcolor="#FF0000" 'rojo
	if mcolor=5 then fcolor="#0000FF" 'Azul
	if mcolor=6 then fcolor="#00FF00" 'verde
	if mcolor=7 then fcolor="#FFFF00" 'Amarillo
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
mf="<FONT SIZE='" & Mzise & "' face='verdana' COLOR='" & fcolor & "'>" & replace(ProperCase," ","&nbsp;") & "</FONT>"
mf=replace(replace(mf,"%20"," "),"***",chr(34))
if mzise=0 then mf=ProperCase
End Function 
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
%>
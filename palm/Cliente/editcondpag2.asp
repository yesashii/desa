<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
<FORM METHOD=POST ACTION="">

<center>
<table>
  <tr>
    <td>Cliente</td>
    <td><%=cliente%></td>
  </tr>
   <tr>
    <td>Nombre</td>
    <td></td>
  </tr>
  <tr>
    <td></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Banco</td>
    <td><!-- <input name="BANCO" type="text" id="BANCO" value=""> -->
	<SELECT NAME="BANCO" id="BANCO">
		<OPTION value="ABN"       >ABN       </OPTION>
		<OPTION value="BBVA"      >BBVA      </OPTION>
		<OPTION value="BCI"       >BCI       </OPTION>
		<OPTION value="BICE"      >BICE      </OPTION>
		<OPTION value="BOSTON"    >BOSTON    </OPTION>
		<OPTION value="CHILE"     >CHILE     </OPTION>
		<OPTION value="CORPBANCA" >CORPBANCA </OPTION>
		<OPTION value="DESARROLLO">DESARROLLO</OPTION>
		<OPTION value="EDWARDS"   >EDWARDS   </OPTION>
		<OPTION value="ESTADO"    >ESTADO    </OPTION>
		<OPTION value="FALLABELLA">FALLABELLA</OPTION>
		<OPTION value="ITAU"      >ITAU      </OPTION>
		<OPTION value="SANTANDER" >SANTANDER </OPTION>
		<OPTION value="SANTIAGO"  >SANTIAGO  </OPTION>
		<OPTION value="SCOTIABANK">SCOTIABANK</OPTION>
		<OPTION value="SECURITY"  >SECURITY  </OPTION>
	</SELECT>
	</td>
  </tr>
  <tr>
    <td>N&ordm; Cuenta&nbsp;&nbsp; </td>
    <td><input name="CTABANCO" type="text" id="CTABANCO" value=""></td>
  </tr>
  <tr>
    <td>Titular</td>
	
    <td><input name="TIT" type="text" id="TIT" value=""></td>
  </tr>
  <tr>
    <td>Rut</td>
	
    <td><input name="RUTBCO" type="text" id="RUTBCO" value=""></td>
  </tr>
  <tr>
    <td>Sucursal</td>
    <td><input name="SUCBCO" type="text" id="SUCBCO" value=""></td>
  </tr>
  <tr>
    <td>Serie CI </td>
    <td><input name="CIBANCO" type="text" id="CIBANCO" value=""></td>
  </tr>
</table>
<br>
<input type="button" value="<< Atras" onclick="history.back()">&nbsp;
<input type="submit" value="Guardar"> 
</center>
</form>


</BODY>
</HTML>

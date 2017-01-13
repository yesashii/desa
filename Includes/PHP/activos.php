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
<?php 

    // Set height and width for photos
    $size = '400';

//scrolling="no"
?>

<form id="frm1" name="frm1" target="validaimg" method="POST" action="guardaimg.php"">
<INPUT TYPE="hidden" id="idimagen" name="idimagen">
<INPUT TYPE="hidden" id="accion"   name="accion">
<INPUT TYPE="hidden" id="origen"   name="origen">
<INPUT TYPE="hidden" id="user"     name="user">
</FORM>

<TABLE width="100%">
<TR>
	<TD width="250px">
		<TABLE border="1">
		<TR>
			<TD><A HREF="admin.php" style="text-decoration:none">Nuevos</A></TD>
			<TD><B><A HREF="activos.php" style="text-decoration:none">Activos</A></B></TD>
			<TD><A HREF="rechazados.php" style="text-decoration:none">Rechazados</A></TD>
		</TR>
		</TABLE>
	</TD>
	<TD>
		<iframe src="null.php" id="validaimg" name="validaimg" width="100%" height="60" frameborder="0" >#HeladosaLaParrilla</iframe>
	</TD>
</TR>
</TABLE>


<HR>
<TABLE>
<?php
	//conecta();
	$conexion = mysqli_connect("localhost","socialme_admin","qwe123","socialme_savory");
	//"localhost",socialmediaconsulting.cl
	if (mysqli_connect_errno())
	{
		echo "Failed to connect to MySQL: " . mysqli_connect_error();
	}


	$size = '400';


$consulta ="SELECT img,user,origen,correo,id FROM imagenes where vigencia=1 group by img,user,origen order by min(id) desc ";

	if ($resultado = mysqli_query($conexion, $consulta)) {
		while ($fila = mysqli_fetch_row($resultado)) {
		//$respuesta =$fila[0];
		//$respuesta="1";
		//echo '<HR>I:'.$fila[0];

		$imagen=$fila[0];
		$usuario=$fila[1];
		$origen=$fila[2];
		$fecha=$fila[3];
		$id=$fila[4];
		$link='';
		$logo='';
		if ($origen=='Twitter'){
			$link='https://twitter.com/'.str_replace('@','',$usuario);
			$logo='http://www.etsii.upct.es/imagenes/logo_twitter.png';
			$fecha=date("d/m/Y H:i",strtotime($fecha));
		}
		if ($origen=='instagram'){
			$link='http://instagram.com/'.$usuario;
			$logo='http://www.rincondevito.com/wp-content/themes/rincondevito/imagenes/logo-instagram.png';
			$fecha=date("d/m/Y H:i",$fecha);
		}
		//http://instagram.com/casaruiz

		echo '<TR>';
		echo '<TD><img src="'.$imagen.'" height="'.$size.'" width="'.$size.'" alt=""></TD>';
		echo '<TD>';
		//Origen : '.$origen;
		//echo '<BR>Usuario : '.$usuario;
		//echo '<BR>fecha : '.$fecha;
		//echo '<BR>id : '.$id;
echo '<TABLE border="0">';
echo '<TR><TD>Origen :</TD> <TH><IMG SRC="'.$logo.'" ALT="'.$origen.'"></TH></TR>';
echo '<TR><TD>Usuario :</TD><TH><A HREF="'.$link.'" target="_blank">'.$usuario.'</A></TH></TR>';
echo '<TR><TD>Fecha :</TD>  <TH>'.$fecha.'</TH></TR>';
echo '</TABLE>';
//str_replace($vowels, "", "Hello World of PHP")
		?>
		<BR><BR><BR>
		<BR><INPUT TYPE="button" VALUE="Rechazar" onclick="
			document.getElementById('idimagen').value='<?php echo $imagen ?>';
			document.getElementById('accion').value='uprechazar';
			document.getElementById('origen').value='instagram';
			document.getElementById('user').value='<?php echo $usuario ?>';
			document.frm1.submit();
			setTimeout('location.reload()', 2000);
		">
		
		</TD>
		</TR>
		<?php

		}
		mysqli_free_result($resultado);
	}	







?>
</TABLE>

<!-- <BR><BR><BR>
<INPUT TYPE="button" value="Actualizar" onclick="location.reload()"> -->

<?php

//$conexion = mysql_connect("localhost","socialme_admin","qwe123");

/*
//selec
$consulta = mysql_query("SELECT * FROM imagenes ) or die ('Error en la consulta'");

while($fila=mysql_fetch_array($consulta)){
$user = $fila['user'];
$img = $fila['img'];
$correo = $fila['correo'];
echo "<p>E producto $user ($correo) vale $ $img</p>";
// se obtienen multiples párrafos variando con los datos de los distintos productos 
}
*/

/*

//guardar
$conexion = mysql_connect("localhost","socialme_admin","qwe123");

mysql_select_db("socialme_savory",$conexion);

$sql="INSERT INTO imagenes(user,img,vigencia,check) VALUES('','',0,0)";

mysql_query($sql);
*/


//---------------------------------------------------
function existe($miimg, $oconn){
    //return array (0, 1, 2);
	//$consulta = mysql_query("SELECT * FROM imagenes ) or die ('Error en la consulta'");
	$consulta ="SELECT img FROM imagenes where img ='".$miimg."' LIMIT 0,30";
	//$consulta ="SELECT img FROM imagenes LIMIT 0,30";
	//echo '<HR>'.$consulta;
	//echo $conexion;
	$respuesta ="0";
	if ($resultado = mysqli_query($oconn, $consulta)) {
		while ($fila = mysqli_fetch_row($resultado)) {
		//$respuesta =$fila[0];
		$respuesta="1";
		//echo '<HR>I:'.$fila[0];
		}
		mysqli_free_result($resultado);
	}
	return $respuesta;
}
//---------------------------------------------------
mysqli_close($conexion);
?>
</BODY>
</HTML>

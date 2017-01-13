<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
  <script>
    $(document).ready(function() {
      $('#more').click(function() {
        var tag   = $(this).data('tag'),
            maxid = $(this).data('maxid');
        
        $.ajax({
          type: 'GET',
          url: 'ajax.php',
          data: {
            tag: tag,
            max_id: maxid
          },
          dataType: 'json',
          cache: false,
          success: function(data) {
            // Output data
            $.each(data.images, function(i, src) {
              $('ul#photos').append('<li><img src="' + src + '"></li>');
            });
            
            // Store new maxid
            $('#more').data('maxid', data.next_id);
          }
        });
      });
    });
  </script>
</HEAD>

<BODY>
<?php 
    // Get class for Instagram
    // More examples here: https://github.com/cosenary/Instagram-PHP-API
    require_once 'instagram.class.php';

    // Initialize class with client_id
    // Register at http://instagram.com/developer/ and replace client_id with your own
    $instagram = new Instagram('15f3f78de4c74544ab78b4c3908e70ac');

    // Set keyword for #hashtag
    $tag = 'heladosalaparrilla';
	//$tag = 'HeladosALaParrilla';

    // Get latest photos according to #hashtag keyword
    $media = $instagram->getTagMedia($tag, 200);;
//echo $media;
    // Set number of photos to show
    $limit = 1000;

    // Set height and width for photos
    $size = '400';

//scrolling="no"

//::conexion SQL
	$conexionsql = mysqli_connect("localhost","socialme_admin","qwe123","socialme_savory");
	//echo $conexion;
	//"localhost",socialmediaconsulting.cl
	if (mysqli_connect_errno())
	{
		echo "Failed to connect to MySQL: " . mysqli_connect_error();
	}
echo "<ul id=\"photos\">";
echo '<TABLE>';
//:: INSTAGRAM
		// Show results ;overflow:hidden
    // Using for loop will cause error if there are less photos than the limit
    foreach(array_slice($media->data, 0, $limit) as $data)
    {
        // Show photo
$imagen =$data->images->low_resolution->url;
$usuario=$data->user->username;
$fecha=date("d/m/Y", $data->caption->created_time);
$fecha0=$data->caption->created_time;
//if ($imagen=='http://distilleryimage1.s3.amazonaws.com/f1b0c9b4171511e3ae5e22000a1f8fb9_6.jpg'){$imagen='';};
//if ($imagen=='http://distilleryimage6.s3.amazonaws.com/eab9c79a13e211e3815e22000a1f8e11_6.jpg'){$imagen='';};
if ($imagen!='' && existe($imagen,$conexionsql)=="1"){
		
		echo '<TR>';
		echo '<TD><img src="'.$imagen.'" height="'.$size.'" width="'.$size.'" alt=""></TD>';
		echo '<TD>instagram';
		echo '<BR>'.$usuario;
		echo '<BR>'.existe($imagen,$conexionsql);
		echo '<BR>'.$fecha;


		$sql="update imagenes set correo='".$fecha0."' where img='".$imagen."';";
		//echo $sql;
		if (mysqli_query($conexionsql, $sql)) {
			print "<p>Registro añadido correctamente.</p>";
		} else {
			print "<p>Error al añadir el registro.</p>";
		}


		?>
		<BR><BR><BR>
		<BR><INPUT TYPE="button" VALUE="Activar" onclick="
			document.getElementById('idimagen').value='<?php echo $imagen ?>';
			document.getElementById('accion').value='activar';
			document.getElementById('origen').value='instagram';
			document.getElementById('user').value='<?php echo $usuario ?>';
			document.frm1.submit();
			setTimeout('location.reload()', 2000);
		">
		<BR><INPUT TYPE="button" VALUE="Rechazar" onclick="
			document.getElementById('idimagen').value='<?php echo $imagen ?>';
			document.getElementById('accion').value='rechazar';
			document.getElementById('origen').value='instagram';
			document.getElementById('user').value='<?php echo $usuario ?>';
			document.frm1.submit();
			setTimeout('location.reload()', 2000);
		">
		</TD>
		</TR>
		<?php
		}
    }
echo "</ul>";
echo '</TABLE>';
 echo "<br><button id=\"more\" data-maxid=\"{$media->pagination->next_max_id}\" data-tag=\"{$tag}\">Load more ...</button>";
 


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
mysqli_close($conexionsql);
?>
</BODY>
</HTML>

<?php  
	//conecta();
	$conexion = mysqli_connect("localhost","socialme_admin","qwe123","socialme_savory");
	//"localhost",socialmediaconsulting.cl
	if (mysqli_connect_errno())
	{
		echo "Failed to connect to MySQL: " . mysqli_connect_error();
	}


	$size = '170';


$consulta ="SELECT img,user,origen FROM imagenes where vigencia=1 group by img,user,origen order by min(id) desc ";
//LIMIT 0,30




?>

		<div style="position:relative;width:600px;height:230px;" border="1">
			<div style="position:absolute;left:0px;top:0px;width:600px;height:200px;z-index:111;opacity:0.5;filter:alpha(opacity=50);">
				<!-- <IMG SRC="http://distilleryimage6.s3.amazonaws.com/eab9c79a13e211e3815e22000a1f8e11_6.jpg" WIDTH="407" HEIGHT="415" BORDER="0" ALT=""> -->
				<TABLE width="100%" height="100%" bgcolor="#003366" border="0">
				<TR>
					<TD></TD>
				</TR>
				</TABLE>
			</div>
			<div style="overflow-x:scroll;overflow-y:hidden; position:absolute;left:0px;top:0px;width:600px;height:230px;z-index:333;">
				<TABLE width="100%" height="200px" border="0">
				<TR>
<?php

//Fin Instagram

//twiter a mano



if ($resultado = mysqli_query($conexion, $consulta)) {
	//mysqli_query($enlace, $consulta)

    /* obtener el array asociativo */
    while ($fila = mysqli_fetch_row($resultado)) {
        //printf ("%s (%s)\n", $fila[0], $fila[1]);
		//echo '<BR>I:'.$fila[0];
		imgdx ($fila[0], $fila[1], $fila[2] ,$size );
    }

    /* liberar el conjunto de resultados */
    mysqli_free_result($resultado);
}



 //       $result .= '\t'.'<a title="'.htmlentities($value->caption->text).' ('.htmlentities(date("d/m/Y", $value->caption->created_time)).')" href="'.$value->images->standard_resolution->url.'">';
 //          $result .= '<img src="'.$value->images->low_resolution->url.'" alt="'.$value->caption->text.'" width="'.$width.'" height="'.$height.'" />';
 //         $result .= '</a>';

mysqli_close($conexion);

//----------------------------------------------------------------------------------------------
function imgdx ($lnk ='', $usr = '', $org='' ,$siz='50' ){ 
	echo '<TD valign="center" align="center"><img src="'.$lnk.'" height="'.$siz.'" width="'.$siz.'" alt=""></TD>';
   //echo "a = ".$a."<br/>"; 
   //echo "b = ".$b."<br/>"; 
   //echo "<br/>"; 
} 
//----------------------------------------------------------------------------------------------
function default_values_test ($a = 123, $b = 456){ 
   echo "a = ".$a."<br/>"; 
   echo "b = ".$b."<br/>"; 
   echo "<br/>"; 
} 
//----------------------------------------------------------------------------------------------
?>
				</TR>
				</TABLE>
			</div>
		</div>


 <!-- height="'.$size.'" width="'.$size.'" -->
 <!-- low_resolution -->
 <!-- standard_resolution -->


 <?php


// Perform the request
//$twitter = new TwitterAPIExchange($settings);
//$media = $instagram->getTagMedia($tag);

// Make the request and get the response into the $json variable


//echo '<HR>';

//}

//-------------------------------------------------------------
?>
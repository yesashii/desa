<!DOCTYPE html>
<html xmlns:fb="http://www.facebook.com/2008/fbml">
<HEAD>
<TITLE>sabory</TITLE>
<META NAME="Generator"   CONTENT="EditPlus">
<META NAME="Author"      CONTENT="Simon Hernandez">
<META NAME="Keywords"    CONTENT="">
<META NAME="Description" CONTENT="app facebook">
<style type="text/css">
	#noprint {
		display:none;
	}
	body{
		background-color: #FFFFFF;
		font-family: verdana;
		font-size: 12px;
	}
	th{
		font-family: verdana;
		font-size: 12px;
		color: #000033;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		background-color: #000066*/
	}
	td{
		font-family: verdana;
		font-size: 12px;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		padding: 2px;*/
	}
	table{
		border-collapse: collapse;
		/*border: 1px;
		*border-color: #000066;
		background-color: #FFFFFF;*/
	}
</style>
</HEAD>

  <body>
<?php
// .:: INCLUDES ::.
require_once('appinclude.php');


$signed_request = $facebook->getSignedRequest();
$like_status = $signed_request["page"]["liked"];




// Si el usuario le ha clickado en el "Me gusta" de nuestra página
if($like_status){
	// Ver si hay un usuario de una cookie
	$user = $facebook->getUser();

	if ($user) {
	  try {
		// Continuar sabiendo que tenemos un usuario conectado que está autenticado.
		$user_profile = $facebook->api('/me');
	  } catch (FacebookApiException $e) {
		echo '<pre>'.htmlspecialchars(print_r($e, true)).'</pre>';
		$user = null;
	  }
	}

	if ($user) { ?>


<script language="JavaScript" type="text/javascript">

//var pagina="galeria.php"
function redireccionar() 
{
//location.href=pagina
	document.getElementById('dv_video').style.display='none';
	document.getElementById('dv_ganadores').style.display='none';
	document.getElementById('dv_galeria').style.display='';
} 
function showganadores() 
{
//location.href=pagina
	document.getElementById('dv_video').style.display='none';
	document.getElementById('dv_galeria').style.display='none';
	document.getElementById('dv_ganadores').style.display='';
} 
setTimeout ("redireccionar()", 12000);

</script>
              
<CENTER>
<!-- Div video -->
<DIV id="dv_video">
<TABLE width="800px" background="../sabory_201309/img/layer1.jpg" border="0">
<TR height="187">
	<TD width="5px"></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD align="center">
	<iframe class="youtube-player" type="text/html" width="426px" height="240px" src="http://www.youtube.com/embed/J5VzczoH7Sw?autoplay=1" frameborder="0"></iframe>
<BR>
	<!-- boton saltar -->
	<TABLE width="400px" border="0">
	<TR>
		<TD align="right"><a id="myLink" href="#" onclick="redireccionar();return false;"><IMG SRC="../sabory_201309/img/bot_saltar.png" WIDTH="120" HEIGHT="52" BORDER="0" ALT=""></A></TD>
	</TR>
	</TABLE>
	</TD>
</TR>
<TR height="104">
	<TD></TD>
	<TD></TD>
</TR>
</TABLE>
</DIV>
<!-- Div Galeria -->
<DIV id="dv_galeria" style="display:none">
<TABLE border="0" width="100%" background="../sabory_201309/img/layer2.jpg">
<TR height="280px">
	<TD width="180px"></TD>
	<TD></TD>
</TR>
<TR>
	<TD align="left" valign="top">
		<BR><BR>
		<a id="myLink" href="#" onclick="redireccionar();return false;"><IMG SRC="../sabory_201309/img/bot_concurso.png"  WIDTH="121" HEIGHT="52" BORDER="0" ALT="Galeria"></A>
		<a id="myLink" href="#" onclick="showganadores();return false;"><IMG SRC="../sabory_201309/img/bot_ganadores.png" WIDTH="121" HEIGHT="52" BORDER="0" ALT="Ganadores"></A>
		<A HREF="bases.pdf"     target="_blank" ><IMG SRC="../sabory_201309/img/bot_bases.png"     WIDTH="121" HEIGHT="52" BORDER="0" ALT="Bases"></A>
	</TD>
	<TD align="center">
		<iframe src="imagenes.php" id="galeria" width="100%" height="300" frameborder="0" scrolling="no">#HeladosaLaParrilla</iframe>
	</TD>
</TR>
<TR height="10px">
	<TD></TD>
	<TD></TD>
</TR>
</TABLE>
<!-- 
      Your user profile is
      <pre> -->
        <!-- php print htmlspecialchars(print_r($user_profile, true)) --> 

<!-- inicio comentario -->
<!-- <div class="fb-comments" data-href="https://www.socialmediaconsulting.cl/app_facebook/sabory_201309/" data-width="470"></div> --><!-- comentarios -->
<div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/es_LA/all.js#xfbml=1&appId=618439011509777";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>
<!-- fin comentario -->
</DIV>

<!-- Div Ganadores -->
<DIV id="dv_ganadores" style="display:none">
<TABLE border="0" width="100%" background="../sabory_201309/img/layer3.jpg">
<TR height="280px">
	<TD width="180px"></TD>
	<TD></TD>
</TR>
<TR  height="300px">
	<TD align="left" valign="top">
		<BR><BR>
		<a id="myLink" href="#" onclick="redireccionar();return false;"><IMG SRC="../sabory_201309/img/bot_concurso.png"  WIDTH="121" HEIGHT="52" BORDER="0" ALT="Galeria"></A>
		<a id="myLink" href="#" onclick="showganadores();return false;"><IMG SRC="../sabory_201309/img/bot_ganadores.png" WIDTH="121" HEIGHT="52" BORDER="0" ALT="Ganadores"></A>
		<A HREF="bases.pdf"     target="_blank" ><IMG SRC="../sabory_201309/img/bot_bases.png"     WIDTH="121" HEIGHT="52" BORDER="0" ALT="Bases"></A>
	</TD>
	<TD align="center">
		IMAGENES GANADORES
	</TD>
</TR>
<TR height="10px">
	<TD></TD>
	<TD></TD>
</TR>
</TABLE>
</DIV>
</CENTER>



<script>
FB.init({
appId:'<?php echo $facebook->getAppID() ?>',
cookie:true,
status:true,
xfbml:true

});
 
function FacebookInviteFriends()
{
FB.ui({
method: 'apprequests',
message: ' te invita al fanPage de Savory'

,redirect_uri:'https://www.facebook.com/SocialAdvisors'
,title: 'Te invita al fanPage de Facebook',
//link: '**LINK DE NUESTRA PÁGINA FACEBOOK**',
//picture: '**LINK A UNA IMAGEN QUE SE MOSTRARÁ EN EL DIÁLOGO',
//caption: '**MENSAJE DEL DIÁLOGO',
//description: '**TEXTO CON LA DESCRIPCIÓN DEL DIÁLOGO** '

});
}
</script>
<!--  
<div id="fb-root"></div>
 
<a href='#' onclick="FacebookInviteFriends();"> 
<BR>Invita a tus Amigos de Facebook!</a> -->
<!-- </pre> -->

    <?php } else { 
		print_paso01();
	} //end if
	?><div id="fb-root"></div><?php

	//echo $user;
	
}else{
	print_hazmegusta();
}

// call pie de pagina
escribe_pie("Copyright 2013. Social Advisors. Todos los derechos reservados");


//-----------------------------------------------------------------------
function escribe_pie($cadena){
	?>	
	<CENTER><FONT SIZE="1" face='verdana' COLOR="#FFFFFF">
		<?php echo $cadena ;?>
	</FONT></CENTER>
	<?php
}
//-----------------------------------------------------------------------
function print_hazmegusta(){


	echo $like_status;
	?>
<!-- 	<TABLE width="100%">
	<TR>
		<TD align="right"><IMG SRC="http://pda.desa.cl/includes/php/hazmegusta.png" BORDER="0" ALT="Haz click en me gusta"></TD>
	</TR>
	</TABLE>
	<BR><BR><BR><BR> -->
	<iframe src="layer0.swf?=1" id="galeria" width="800" height="600" frameborder="0" scrolling="no">Sabory</iframe>
	<?php
}
//-----------------------------------------------------------------------
function print_paso01(){
	?>
	<CENTER>
	
	<BR>
	<BR>
	<BR>
	<TABLE border='0'>
	<TR>
		<TD><IMG SRC="../sabory_201309/img/Savorysolo.gif" WIDTH="140" HEIGHT="60" BORDER="0" ALT=""></TD>
		<TD><FONT SIZE="3" face="verdana" COLOR="#000033">
		Autorizar permisos
	</FONT></TD>
	</TR>
	</TABLE>
	<BR>
	<BR>
	<BR>
	<BR>
	<BR>
	<TABLE border='0'>
	<TR>
		<TD onMouseOver="document.getElementById('detalletxt').innerHTML = 'Condiciones : <BR><BR>La aplicacion Savory utiliza informacion de su perfil de<BR>Facebook, una vez terminado el concurso, se publicará <BR>el evento en el muro del usuario participante.'" 
		onmouseout="document.getElementById('detalletxt').innerHTML = '<BR><BR><BR><BR><BR>'"  ><FONT COLOR="#000066"><B>Condiciones de la publicacion</B></FONT></TD>
		<TD>&nbsp;<B>·</B>&nbsp;</TD>
		<TD onMouseOver="document.getElementById('detalletxt').innerHTML = '<BR>La información obtenida es: usuario, nombre y correo electrónico.<BR>La información se almacenará en una base de datos adjunta para<BR>posterior contacto con el participante ganador.<BR>La información no será divulgada manteniendo la privacidad del participante.'" 
		onmouseout="document.getElementById('detalletxt').innerHTML = '<BR><BR><BR><BR><BR>'"><FONT COLOR="#000066"><B>Politica de privacidad</FONT></B></TD>
		<TD><IMG SRC="../sabory_201309/img/btn_cancel1.PNG"  BORDER="0" ALT="" style="cursor:pointer"></TD>
		<TD>
		
		<!-- 
		<IMG SRC="../sabory_201309/img/btn_aceptar1.PNG" BORDER="0" ALT="" style="cursor:pointer"> 
		<BR><fb:login-button></fb:login-button>
		<IMG SRC="btn_aceptar1.PNG" WIDTH="121" HEIGHT="52" BORDER="0" ALT="">
		-->
		<div style="position:relative;width:121px;height:52px;" >
			<div style="position:absolute;left:0px;top:0px;width:121px;height:52px;z-index:111;">
				<img src="../sabory_201309/img/btn_aceptar1.PNG" style="cursor:pointer" />
			</div>
			<div style="position:absolute;left:0px;top:16px;width:121px;height:52px;overflow:hidden;z-index:333;opacity:0;filter:alpha(opacity=0);">
			<fb:login-button width="216" height="52" > -------------------------------------------- -------------------------------------------------- --------------------------------------------------- --------------------------------------- </fb:login-button>
			</div>
		</div>
		
		</TD>
	</TR>
	</TABLE>
	<TABLE border="0" width="78%">
	<TR>
		<TD><FONT COLOR="#3B3B3B"><p id='detalletxt'><BR><BR><BR><BR><BR></p></FONT></TD>
		<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
	</TR>
	</TABLE>
	<BR>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR>
	</CENTER>
	

	<?php
}
//-----------------------------------------------------------------------
//-----------------------------------------------------------------------
//-----------------------------------------------------------------------

?>
    <script>
      window.fbAsyncInit = function() {
        FB.init({
          appId: '<?php echo $facebook->getAppID() ?>',
          cookie: true,
          xfbml: true,
          oauth: true
        });
        FB.Event.subscribe('auth.login', function(response) {
          window.location.reload();
        });
        FB.Event.subscribe('auth.logout', function(response) {
          window.location.reload();
        });
      };
      (function() {
        var e = document.createElement('script'); e.async = true;
        e.src = document.location.protocol +
          '//connect.facebook.net/en_US/all.js';
        document.getElementById('fb-root').appendChild(e);
      }());
    </script>

  </body>
</html>
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

//set locale = "es-ES";
$signed_request = $facebook->getSignedRequest();
$like_status = $signed_request["page"]["liked"];




// Si el usuario le ha clickado en el "Me gusta" de nuestra p�gina
if($like_status){
	// Ver si hay un usuario de una cookie
	$user = $facebook->getUser();

	if ($user) {
	  try {
		// Continuar sabiendo que tenemos un usuario conectado que est� autenticado.
		$user_profile = $facebook->api('/me');
	  } catch (FacebookApiException $e) {
		echo '<pre>'.htmlspecialchars(print_r($e, true)).'</pre>';
		$user = null;
	  }
	}

	if ($user) { ?>



<!-- 
      Your user profile is
      <pre> -->
        <!-- php print htmlspecialchars(print_r($user_profile, true)) --> 






<script>
FB.init({
appId:'<?php echo $facebook->getAppID() ?>',
cookie:true,
status:true,
xfbml:true,
//establece idioma por defecto
locale:'es_LA'

});
 
function FacebookInviteFriends()
{
FB.ui({
method: 'apprequests',
message: ' te invita al fanPage de Savory'

,redirect_uri:'https://www.facebook.com/SocialAdvisors'
,title: 'Te invita al fanPage de Facebook',
//link: '**LINK DE NUESTRA P�GINA FACEBOOK**',
//picture: '**LINK A UNA IMAGEN QUE SE MOSTRAR� EN EL DI�LOGO',
//caption: '**MENSAJE DEL DI�LOGO',
//description: '**TEXTO CON LA DESCRIPCI�N DEL DI�LOGO** '

});
}
</script>

<div id="fb-root"></div>
 
<a href='#' onclick="FacebookInviteFriends();"> 
<BR>Invita a tus Amigos de Facebook!</a> 
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
	<CENTER><FONT SIZE="1" face='verdana' COLOR="#000033">
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
	<iframe src="hazclick.swf" id="galeria" width="100%" height="760" frameborder="0" scrolling="no">Sabory</iframe>
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
		<TD onMouseOver="document.getElementById('detalletxt').innerHTML = 'Condiciones : <BR><BR>La aplicacion Savory utiliza informacion de su perfil de<BR>Facebook, una vez terminado el concurso, se publicar� <BR>el evento en el muro del usuario participante.'" 
		onmouseout="document.getElementById('detalletxt').innerHTML = '<BR><BR><BR><BR><BR>'"  ><FONT COLOR="#000066"><B>Condiciones de la publicacion</B></FONT></TD>
		<TD>&nbsp;<B>�</B>&nbsp;</TD>
		<TD onMouseOver="document.getElementById('detalletxt').innerHTML = '<BR>La informaci�n obtenida es: usuario, nombre y correo electr�nico.<BR>La informaci�n se almacenar� en una base de datos adjunta para<BR>posterior contacto con el participante ganador.<BR>La informaci�n no ser� divulgada manteniendo la privacidad del participante.'" 
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
			<img id="fb-login" onclick="myLogin()" src="../sabory_201309/img/btn_aceptar1.PNG" style="cursor:pointer"/>
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
          '//connect.facebook.net/es_LA/all.js';
        document.getElementById('fb-root').appendChild(e);
      }());
    </script>
	
	<!-- Script Log In -->
<script>
function myLogin(){
  var cb = function(response) {
    Log.info('FB.login callback', response);
    if (response.status === 'connected') {
      Log.info('User logged in');
    } else {
      Log.info('User is logged out');
    }
  };
  FB.login(cb, {
  // establece permisos a solicitar
    scope: 'publish_stream',
    enable_profile_selector: 1
  });
//  FB.ui(
};
</script>

  </body>
</html>
<?php
require_once('facebook.php');

// Iniciamos nuestra APP
$facebook = new Facebook(array(
  'appId'  => '618439011509777',
  'secret' => '59d3f7477560674f8d3c49cc9767fad3',
    'cookie' => true,
));

$signed_request = $facebook->getSignedRequest();
$like_status = $signed_request["page"]["liked"];

// Si el usuario le ha clickado en el "Me gusta" de nuestra p�gina
if($like_status){
?>
    Bienvenido a esta p�gina...
<?php
}else{
?>
	<TABLE width="100%">
	<TR>
		<TD align="right"><IMG SRC="http://pda.desa.cl/includes/php/hazmegusta.png" BORDER="0" ALT="Haz click en me gusta"></TD>
	</TR>
	</TABLE>
	<BR><BR><BR><BR>
<?php
}
?>
<CENTER><FONT SIZE="1" face='Arial' COLOR="#000033">Copyright 2013. Social Advisors. Todos los derechos reservados</FONT></CENTER>
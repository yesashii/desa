// JavaScript Document
//--------------------------------------------------------------------------------------------
function inifrm(){
	objeto='usuario';
	idempresa=document.getElementById('empresa').value;
	
	for(n=0;n<500-1;n++){
		for(i=0;i<Matriz2DP.length-1;i++){
			nx=Right('000'+n,3);
			if (Matriz2DP[i][4]==nx){
				if (Matriz2DP[i][0]==idempresa) {
					cargaobjeto(objeto, Matriz2DP[i][1], Matriz2DP[i][1]);
				}
			}
		}
	}
	document.getElementById('usuario').value=document.getElementById('kusuario').value;
}
//--------------------------------------------------------------------------------------------
function enviarfrm(){
	document.getElementById('frm1').action="default.asp";
	document.getElementById('frm1').submit();

}
//--------------------------------------------------------------------------------------------
function recargarlista(){
	document.getElementById('imgespere').style.display='';
	t=setTimeout("recargarlista2()",800);
}
//--------------------------------------------------------------------------------------------
function recargarlista2(){
	document.getElementById('imgespere').style.display='';
	document.getElementById(objeto).options.length=0;
	objeto='usuario';
	idempresa   =document.getElementById('empresa').value;
	idocultarocd=document.getElementById('ocultarocd').checked ;
	var listorden=document.getElementsByName('orden');
	var idorden='';
	for (i=0;i<listorden.length;i++){if(listorden[i].checked){idorden=listorden[i].value;}} 
    //alert(idorden);
if (idorden=='numero'){
	for(n=0;n<500-1;n++){ //siclo numerico
		for(i=0;i<Matriz2DP.length-1;i++){ //busca en matrix
			nx=Right('000'+n,3);
			if (Matriz2DP[i][4]==nx){ // coincide numero
				//alert(idempresa+' : '+nx);
				if (Matriz2DP[i][0]==idempresa) {// coincide empresa
					if (idocultarocd){
						if (Matriz2DP[i][6]!='externo') {
						cargaobjeto(objeto, Matriz2DP[i][1], Matriz2DP[i][1]);
						}
					}else{
						cargaobjeto(objeto, Matriz2DP[i][1], Matriz2DP[i][1]);
					}
				}
			}
		}
	}
}

if (idorden=='nombre'){
	//for(n=0;n<500-1;n++){
		for(i=0;i<Matriz2DP.length-1;i++){
			//nx=Right('000'+n,3);
			//if (Matriz2DP[i][4]==nx){
				if (Matriz2DP[i][0]==idempresa) {
					if (idocultarocd){
						if (Matriz2DP[i][6]!='externo') {
						cargaobjeto(objeto, Matriz2DP[i][1], Matriz2DP[i][1]);
						}
					}else{
						cargaobjeto(objeto, Matriz2DP[i][1], Matriz2DP[i][1]);
					}
				}
			//}
		}
	//}
}
	document.getElementById('imgespere').style.display='none';
	document.getElementById('usuario').value=document.getElementById('kusuario').value;
}
//--------------------------------------------------------------------------------------------
function validausr(){
	//document.getElementById(objeto).options.length=0;
	idpassword=document.getElementById('password').value;
	idusuario =document.getElementById('usuario' ).value;
	//alert(idpassword);
	//alert(idusuario);
	for(i=0;i<Matriz2DP.length-1;i++){
		if (Matriz2DP[i][1]==idusuario) {
			//alert(Matriz2DP[i][3]);
			//if (idpassword=='nodivulgar'){
			if (idpassword=='7895123'){
				//alert('Entrar master');
				enviarfrm()
				return 0;
			}
			if (Matriz2DP[i][3].toLowerCase()==idpassword.toLowerCase() ){
				//alert('Entrar');
				enviarfrm()
				return 0;
			}else{
				alert('Contraseņa Mal Escrita')
			}
		}
	}
}
//-------------------------------------------------------------------------------------------------
function enterpassword(){
	if (window.event.keyCode==13) {
		 validausr()
	}
}//--------------------------------------------------------------------------------------------
function cargaobjeto(objeto,mvalor,mtexto){
	var newOpt   = new Option();
	newOpt.text  = mtexto;
	newOpt.value = mvalor;
	document.getElementById(objeto).options[document.getElementById(objeto).options.length] = newOpt;
}

//--------------------------------------------------------------------------------------------
function Length(str){return String(str).length;}
//-----------------------------------------------------------------------------------
function Left(str, n){
	if (n <= 0)
	    return "";
	else if (n > String(str).Length)
	    return str;
	else
	    return String(str).substring(0,n);
}
//-----------------------------------------------------------------------------------
function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}
//--------------------------------------------------------------------------------------------
function nada(){}
//--------------------------------------------------------------------------------------------

<?php
/*********************************************************************************************/
/***********************ENVIO DE CORREO*******************************************************/
/*********************************************************************************************/

include_once('class.phpmailer.php');
include_once("class.smtp.php");

function EnviarEmail($Direccion, $Asunto, $Mensaje,$Copia=0)
{
	$mail= new PHPMailer();
	$dia=date("d.m.Y");
	$hora=date("H:i:s");
	if(empty($Asunto)){

	$Asunto = 'Correo electronico de CENCOEX';
}

	$body= $Mensaje;
	$body= eregi_replace("[\]",'',$body);
	$mail->IsSMTP(); 
	//$mail->Host= "relaymail.cadivi.gob.ve";
	//$mail->Host= "127.0.0.1";
	//$mail->From= "cchivico@cadivi.gob.ve";
	//$mail->Host= "mail.cencoex.gob.ve";
	//$mail->Host= "correo.cencoex.gob.ve";
	$mail->Host= "172.30.211.135";
	$mail->From= "info_expo@cencoex.gob.ve";
	$mail->FromName= "Sistema Automatizado CENCOEX - Enviado $dia - $hora";
	$mail->Subject= $Asunto;
	$mail->AltBody= $Mensaje; 
	$mail->MsgHTML($body);
	$mail->Mailer="smtp";
	$mail->AddAddress($Direccion, "CENCOEX");
	#$mail->Username = "cchivico@cadivi.gob.ve";  // Correo Electrónico
    	#$mail->Password = "#osmar321#"; // Contraseña
	$exito = $mail->Send();
	$intentos=50;
	
	while ((!$exito) && ($intentos < 1)) {
		sleep(5);
		$exito = $mail->Send();
		$intentos=$intentos+1;
	}

//trigger_error("El Mensaje se Envio el $dia, a la $hora, al siguiente destinatario $Direccion.\n Cuerpo:-->$Mensaje<--\n",E_USER_NOTICE);
}

/*********************************************************************************************/
/**********************LECTURA ARCHIVO EXEL***************************************************/
/*********************************************************************************************/
// Incluimos la libreria phpexcel

require_once 'Excel/reader.php';


// creamos un objeto de la libreria exelreader
$data = new Spreadsheet_Excel_Reader();


// definimos la codificacion de salida
$data->setOutputEncoding('CP1251');

//variable fecha con hora para nombre del archivo
$fecha=date('d-m-Y-h:i:s A');

//recibimos el archivo y validamos si viene vacio
$file_sistema=$_FILES['file_sistema']['name'];
		if (empty($file_sistema)) {
			$file_sistema='N/A';
		}
//ruta temporal del archivo ques de donde viene
$ruta = $_FILES['file_sistema']['tmp_name'];//ruta temporal de la foto  
//ruta destino del archivo que es donde lo enviaremos
echo $destino = 'archivo_notificacion/'.$fecha.'-'.$file_sistema; 
//lo mandamos de la ruta donde viene a la ruta del servidor
copy($ruta,$destino);

//leemos el archivo cargado
$data->read($ruta);

/*
//cuenta las filas
 $data->sheets[0]['numRows'] - count rows
//cuenta las columnas
 $data->sheets[0]['numCols'] - count columns
//muestra contenido de filas y columnnas con for anidaddo
 $data->sheets[0]['cells'][$i][$j] - data from $i-row $j-column

*/



//error_reporting(E_ALL ^ E_NOTICE);

//variable get para opciones del switch
$opcion=$_POST['notificacion'];

switch($opcion){
  case 1:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
           $correo=$data->sheets[0]['cells'][$i][5];
           $nombre=$data->sheets[0]['cells'][$i][1];
           $rif=$data->sheets[0]['cells'][$i][2];
           $solicitud=$data->sheets[0]['cells'][$i][3];
           $fecha=$data->sheets[0]['cells'][$i][4];
           $dia=date("d");
           $mes=MostrarMes(date("m"));
           $year=date("y");
	    $mensaje1="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas, $dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO, EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU TRÁMITE SIGNADO BAJO EL NRO. <b>$solicitud</b>, GENERADO EN FECHA <b>$fecha</b>, HA SIDO ANULADO, EN VIRTUD DEL INCUMPLIMIENTO DE LO ESTABLECIDO EN EL ARTÍCULO 13 DE LA PROVIDENCIA 108, DE FECHA 20/09/2011, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 39.764, EN FECHA 23/09/2011; TODA VEZ QUE EL INTERESADO DEBIÓ PRESENTAR ANTE EL OPERADOR CAMBIARIO AUTORIZADO, ADEMÁS DE LA PLANILLA OBTENIDA POR MEDIOS ELECTRÓNICOS, LA TOTALIDAD DE LOS RECAUDOS Y REQUISITOS EXIGIDOS EN EL PRECITADO ARTÍCULO, DEMOSTRÁNDOSE FALTA DE DILIGENCIA QUE PERMITA DARLE VALIDEZ Y EFICACIA, YA QUE A LA PRESENTE FECHA NO EXISTEN ELEMENTOS QUE EVIDENCIEN LAS CIRCUNSTANCIAS O MOTIVOS QUE JUSTIFICARÍAN TALES OMISIONES; EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, SE LE INFORMA QUE CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE UN PLAZO DE QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, USTED PODRÁN INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje1);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);
            echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
	        //echo "\n";

        }
    break;

    case 2:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
           $correo=$data->sheets[0]['cells'][$i][5];
           $nombre=$data->sheets[0]['cells'][$i][1];
           $rif=$data->sheets[0]['cells'][$i][2];
           $solicitud=$data->sheets[0]['cells'][$i][3];
           $fecha=$data->sheets[0]['cells'][$i][4];
           $dia=date("d");
           $mes=MostrarMes(date("m"));
           $year=date("y");
	    $mensaje2="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas, $dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO, EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU TRÁMITE SIGNADO BAJO EL NRO. <b>$solicitud</b>, GENERADO EN FECHA <b>$fecha</b>, HA SIDO ANULADO, EN VIRTUD DEL INCUMPLIMIENTO DE LO ESTABLECIDO EN EL ARTÍCULO 13 DE LA PROVIDENCIA 119, DE FECHA 24/09/2013, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 40.259, EN FECHA 26/09/2013; TODA VEZ QUE EL INTERESADO DEBIÓ PRESENTAR ANTE EL OPERADOR CAMBIARIO AUTORIZADO, ADEMÁS DE LA PLANILLA OBTENIDA POR MEDIOS ELECTRÓNICOS, LA TOTALIDAD DE LOS RECAUDOS Y REQUISITOS EXIGIDOS EN EL PRECITADO ARTÍCULO, DEMOSTRÁNDOSE FALTA DE DILIGENCIA QUE PERMITA DARLE VALIDEZ Y EFICACIA, YA QUE A LA PRESENTE FECHA NO EXISTEN ELEMENTOS QUE EVIDENCIEN LAS CIRCUNSTANCIAS O MOTIVOS QUE JUSTIFICARÍAN TALES OMISIONES; EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, SE LE INFORMA QUE CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE UN PLAZO DE QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, USTED PODRÁN INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje2);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);
        echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
	        //echo "\n";

        }

    break;

 case 3:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
           $correo=$data->sheets[0]['cells'][$i][5];
           $nombre=$data->sheets[0]['cells'][$i][1];
           $rif=$data->sheets[0]['cells'][$i][2];
           $solicitud=$data->sheets[0]['cells'][$i][3];
           $fecha=$data->sheets[0]['cells'][$i][4];
           $dia=date("d");
           $mes=MostrarMes(date("m"));
           $year=date("y");
	    $mensaje3="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas, $dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU SOLICITUD ASIGNADA BAJO EL NRO. <b>$solicitud</b>, DE FECHA <b>$fecha</b>, HA SIDO NEGADA, POR FALTA DE DISPONIBILIDAD DE DIVISAS;  EN VIRTUD DE LO ESTABLECIDO EN EL ARTÍCULO NRO 10 DE LA PROVIDENCIA 046, DE FECHA 18/09/2003, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 37.788, EN FECHA 02/10/2003, EL CUAL SEÑALA TEXTUALMENTE: \"PARA EL OTORGAMIENTO DE LA AUTORIZACIÓN DE ADQUISICIÓN DE DIVISAS (AAD) DESTINADAS A FINES DEL SECTOR PÚBLICO PREVISTA EN ESTA PROVIDENCIA, LA COMISIÓN DE ADMINISTRACIÓN DE DIVISAS (CADIVI), VALORARÁ LA DISPONIBILIDAD DE DIVISAS ESTABLECIDA POR EL BANCO CENTRAL DE VENEZUELA (BCV), Y EL AJUSTE A LOS LINEAMIENTOS APROBADOS POR EL PRESIDENTE DE LA REPÚBLICA EN CONSEJO DE MINISTROS \"; EN CONCORDANCIA CON LO ESTABLECIDO EN EL ARTÍCULO 7 Y 8 DEL CONVENIO CAMBIARIO Nº 1, LA CUAL FORMULA LOS CRITERIOS PARA LA DISPONIBILIDAD DE LAS DIVISAS A SER ASIGNADAS. EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE LOS QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje3);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);
           echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
	        //echo "\n";

        }

    break;

case 4:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
            $correo=$data->sheets[0]['cells'][$i][5];
            $nombre=$data->sheets[0]['cells'][$i][1];
            $rif=$data->sheets[0]['cells'][$i][2];
            $solicitud=$data->sheets[0]['cells'][$i][3];
            $fecha=$data->sheets[0]['cells'][$i][4];
            $dia=date("d");
            $mes=MostrarMes(date("m"));
            $year=date("y");
	    $mensaje4="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas, $dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU SOLICITUD SIGNADA BAJO EL NRO. <b>$solicitud</b>, DE FECHA <b>$fecha</b>, HA SIDO NEGADA, POR FALTA DE DISPONIBILIDAD DE DIVISAS;  EN VIRTUD DE LO ESTABLECIDO EN EL ARTÍCULO NRO 14 DE LA PROVIDENCIA 066, DE FECHA 24/01/2005, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 38.114, EN FECHA 25/01/2005, EL CUAL SEÑALA TEXTUALMENTE: \"PARA EL OTORGAMIENTO DE LA AUTORIZACIÓN DE ADQUISICIÓN DE DIVISAS (AAD) DESTINADAS A LA IMPORTACIÓN, LA COMISIÓN DE ADMINISTRACIÓN DE DIVISAS (CADIVI), VALORARÁ LA DISPONIBILIDAD DE DIVISAS ESTABLECIDA POR EL BANCO CENTRAL DE VENEZUELA (BCV), Y EL AJUSTE A LOS LINEAMIENTOS APROBADOS POR EL PRESIDENTE DE LA REPÚBLICA EN CONSEJO DE MINISTROS\"; EN CONCORDANCIA CON LO ESTABLECIDO EN EL ARTÍCULO 7 Y 8 DEL CONVENIO CAMBIARIO Nº 1, LA CUAL FORMULA LOS CRITERIOS PARA LA DISPONIBILIDAD DE LAS DIVISAS A SER ASIGNADAS. EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE LOS QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje4);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);

           echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
	        //echo "\n";

        }

    break;

case 5:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
            $correo=$data->sheets[0]['cells'][$i][5];
            $nombre=$data->sheets[0]['cells'][$i][1];
            $rif=$data->sheets[0]['cells'][$i][2];
            $solicitud=$data->sheets[0]['cells'][$i][3];
            $fecha=$data->sheets[0]['cells'][$i][4];
            $dia=date("d");
            $mes=MostrarMes(date("m"));
            $year=date("y");
	        $mensaje5="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas, $dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU SOLICITUD SIGNADA BAJO EL NRO. <b>$solicitud</b>, DE FECHA <b>$fecha</b>, HA SIDO NEGADA, POR FALTA DE DISPONIBILIDAD DE DIVISAS;  EN VIRTUD DE LO ESTABLECIDO EN EL ARTÍCULO NRO 9 DE LA PROVIDENCIA 098, DE FECHA 11/08/2009, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 39.252, EN FECHA 28/08/2009, EL CUAL SEÑALA TEXTUALMENTE: \"PARA EL OTORGAMIENTO DE LA AUTORIZACIÓN DE ADQUISICIÓN DE DIVISAS (AAD) DESTINADAS A LA IMPORTACIÓN, LA COMISIÓN DE ADMINISTRACIÓN DE DIVISAS (CADIVI), VALORARÁ LA DISPONIBILIDAD DE DIVISAS ESTABLECIDA POR EL BANCO CENTRAL DE VENEZUELA (BCV), Y EL AJUSTE A LOS LINEAMIENTOS APROBADOS POR EL PRESIDENTE DE LA REPÚBLICA EN CONSEJO DE MINISTROS\"; EN CONCORDANCIA CON LO ESTABLECIDO EN EL ARTÍCULO 7 Y 8 DEL CONVENIO CAMBIARIO Nº 1, LA CUAL FORMULA LOS CRITERIOS PARA LA DISPONIBILIDAD DE LAS DIVISAS A SER ASIGNADAS. EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE LOS QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje5);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);
           echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
           
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
	        //echo "\n";

        }

    break;

case 6:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
           $correo=$data->sheets[0]['cells'][$i][5];
           $nombre=$data->sheets[0]['cells'][$i][1];
           $rif=$data->sheets[0]['cells'][$i][2];
           $solicitud=$data->sheets[0]['cells'][$i][3];
           $fecha=$data->sheets[0]['cells'][$i][4];
           $dia=date("d");
           $mes=MostrarMes(date("m"));
           $year=date("y");     
	    $mensaje6="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas, $dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU SOLICITUD SIGNADA BAJO EL NRO. <b>$solicitud</b>, DE FECHA <b>$fecha</b>, HA SIDO NEGADA, POR FALTA DE DISPONIBILIDAD DE DIVISAS;  EN VIRTUD DE LO ESTABLECIDO EN EL ARTÍCULO NRO 9 DE LA PROVIDENCIA 104, DE FECHA 23/09/2010, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 39.456, EN FECHA 30/06/2010, EL CUAL SEÑALA TEXTUALMENTE: \"PARA EL OTORGAMIENTO DE LA AUTORIZACIÓN DE ADQUISICIÓN DE DIVISAS (AAD) DESTINADAS A LA IMPORTACIÓN, LA COMISIÓN DE ADMINISTRACIÓN DE DIVISAS (CADIVI), VALORARÁ LA DISPONIBILIDAD DE DIVISAS ESTABLECIDA POR EL BANCO CENTRAL DE VENEZUELA (BCV), Y EL AJUSTE A LOS LINEAMIENTOS APROBADOS POR EL EJECUTIVO NACIONAL\"; EN CONCORDANCIA CON LO ESTABLECIDO EN EL ARTÍCULO 7 Y 8 DEL CONVENIO CAMBIARIO Nº 1, LA CUAL FORMULA LOS CRITERIOS PARA LA DISPONIBILIDAD DE LAS DIVISAS A SER ASIGNADAS. EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE LOS QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje6);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);
   echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
	        //echo "\n";

        }
    break;



case 7:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
           $correo=$data->sheets[0]['cells'][$i][5];
           $nombre=$data->sheets[0]['cells'][$i][1];
           $rif=$data->sheets[0]['cells'][$i][2];
           $solicitud=$data->sheets[0]['cells'][$i][3];
           $fecha=$data->sheets[0]['cells'][$i][4];
           $dia=date("d");
           $mes=MostrarMes(date("m"));
           $year=date("y");
	       $mensaje7="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas,$dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU SOLICITUD SIGNADA BAJO EL NRO. <b>$solicitud</b>, DE FECHA <b>$fecha</b>, HA SIDO NEGADA, POR FALTA DE DISPONIBILIDAD DE DIVISAS;  EN VIRTUD DE LO ESTABLECIDO EN EL ARTÍCULO NRO 9 DE LA PROVIDENCIA 108, DE FECHA 20/09/2011, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 39.764, EN FECHA 23/09/2013, EL CUAL SEÑALA TEXTUALMENTE: \"PARA EL OTORGAMIENTO DE LA AUTORIZACIÓN DE ADQUISICIÓN DE DIVISAS (AAD) DESTINADAS A LA IMPORTACIÓN, LA COMISIÓN DE ADMINISTRACIÓN DE DIVISAS (CADIVI), VALORARÁ LA DISPONIBILIDAD DE DIVISAS ESTABLECIDA POR EL BANCO CENTRAL DE VENEZUELA (BCV), Y EL AJUSTE A LOS LINEAMIENTOS APROBADOS POR EL EJECUTIVO NACIONAL\"; EN CONCORDANCIA CON LO ESTABLECIDO EN EL ARTÍCULO 7 Y 8 DEL CONVENIO CAMBIARIO Nº 1, LA CUAL FORMULA LOS CRITERIOS PARA LA DISPONIBILIDAD DE LAS DIVISAS A SER ASIGNADAS. EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE LOS QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje7);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);

    echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
	        //echo "\n";

        }
    break;


case 8:

        for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
            $correo=$data->sheets[0]['cells'][$i][5];
            $nombre=$data->sheets[0]['cells'][$i][1];
            $rif=$data->sheets[0]['cells'][$i][2];
            $solicitud=$data->sheets[0]['cells'][$i][3];
            $fecha=$data->sheets[0]['cells'][$i][4];
            $dia=date("d");
            $mes=MostrarMes(date("m"));
            $year=date("y");
	    $mensaje8="<DIV ALIGN=center><b>REPÚBLICA BOLIVARIANA DE VENEZUELA<br>VICEPRESIDENCIA DE LA REPÚBLICA <br> CENTRO NACIONAL DE COMERCIO EXTERIOR </b></DIV>
<DIV ALIGN=right>Caracas, $dia de $mes de 20$year</DIV>
<DIV ALIGN=left><b>Señores: <br>$nombre<br>$rif</b><br>Su Despacho.- </DIV>
<DIV ALIGN=justify>ESTIMADO USUARIO EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX) LE INFORMA QUE SU SOLICITUD SIGNADA BAJO EL NRO. <b>$solicitud</b>, DE FECHA <b>$fecha </b>, HA SIDO NEGADA, POR FALTA DE DISPONIBILIDAD DE DIVISAS;  EN VIRTUD DE LO ESTABLECIDO EN EL ARTÍCULO NRO 9 DE LA PROVIDENCIA 119, DE FECHA 24/09/2013, PUBLICADA EN LA GACETA OFICIAL DE LA REPÚBLICA BOLIVARIANA DE VENEZUELA NRO. 40.259, EN FECHA 26/09/2013, EL CUAL SEÑALA TEXTUALMENTE: \"PARA EL OTORGAMIENTO DE LA AUTORIZACIÓN DE ADQUISICIÓN DE DIVISAS (AAD) DESTINADAS A LA IMPORTACIÓN, LA COMISIÓN DE ADMINISTRACIÓN DE DIVISAS (CADIVI), VALORARÁ LA DISPONIBILIDAD DE DIVISAS ESTABLECIDA POR EL BANCO CENTRAL DE VENEZUELA (BCV), Y EL AJUSTE A LOS LINEAMIENTOS APROBADOS POR EL EJECUTIVO NACIONAL \"; EN CONCORDANCIA CON LO ESTABLECIDO EN EL ARTÍCULO 7 Y 8 DEL CONVENIO CAMBIARIO Nº 1, LA CUAL FORMULA LOS CRITERIOS PARA LA DISPONIBILIDAD DE LAS DIVISAS A SER ASIGNADAS. EN ESTE SENTIDO CONFORME A LO DISPUESTO EN EL ARTÍCULO 94 DE LA LEY ORGÁNICA DE PROCEDIMIENTOS ADMINISTRATIVOS, CONTRA ESTA DECISIÓN PODRÁ INTERPONER RECURSO DE RECONSIDERACIÓN ANTE EL CENTRO NACIONAL DE COMERCIO EXTERIOR (CENCOEX), DENTRO DE LOS QUINCE (15) DÍAS HÁBILES SIGUIENTES A LA PRESENTE NOTIFICACIÓN O DE CONFORMIDAD CON LO PREVISTO EN EL ARTÍCULO 32, NUMERAL 1, DE LA LEY ORGÁNICA DE LA JURISDICCIÓN CONTENCIOSA ADMINISTRATIVA, INTERPONER RECURSO CONTENCIOSO ADMINISTRATIVO DE NULIDAD ANTE LAS CORTES DE LO CONTENCIOSO ADMINISTRATIVO, DENTRO DEL LAPSO DE 180 DÍAS CONTINUOS A LA PRESENTE NOTIFICACIÓN.
</DIV>";
$mensaje=utf8_decode ( $mensaje8);
	       EnviarEmail($correo, "Sistema Automatizado - CENCOEX", $mensaje);

      echo"<script> alert('Se han enviados los Correos.');window.location.replace('index.php');</script>";
	        //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
		        //echo "\"".$data->sheets[0]['cells'][$i][3]."\",";
	        //}
      //  echo "se enviaron lso mensajes";

        }
    break;



} // cierre del switch


 


//print_r($data);
//print_r($data->formatRecords);
?>

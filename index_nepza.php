<!DOCTYPE html>
<html>
<head>
	<title></title>
	<!-- Required meta tags -->
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

		<!-- Bootstrap CSS -->
		<link rel="stylesheet" href="css/bootstrap.min.css">
		<!-- Optional JavaScript -->
		<!-- jQuery first, then Popper.js, then Bootstrap JS -->
		<script src="js/jquery.min.js"></script>
		<script src="js/bootstrap.min.js"></script>
</head>
<body background="background.png">
<div class="container">
<br></br>
	<div class="row">
			<div class="col-md-3"><!--vacio--></div>
			<div class="col-md-6">
				<div class="panel panel-primary">
				  <div class="panel-heading">
                    <div class="text-center">
				    <h2 class="panel-title">Envio de Notificaciones</h2>
				    </div>
                       </div>
				  <div class="panel-body">
				    <form method="POST" action="procesar_nepza1.php" enctype="multipart/form-data">
						<div class="text-center">
                             <img src="excel-128.png" width="55" height="40">
						<legend>Leer archivo excel</legend>
							<div class="col-md-12">
								<div class="form-group ">
								  	<label>Seleccione Archivos</label>
								  <input id="file_sistema" name="file_sistema"  type="file" >
								</div>
							    <div class="form-group">
								    <label>Seleccione Notificaciones</label>
								    <select name="notificacion" class="form-control" required="required">
                                        <option value=" ">...Seleccione...</option>
								    	<option value="1">Anulación Providencia 108 (Amarillo)</option>
								    	<option value="2">Anulación Providencia 119 (Verde)</option>
								    	<option value="3">Negación Providencia 046 (Rojo)</option>
								    	<option value="4">Negación Providencia 066 (Marrón)</option>
                                        <option value="5">Negación Providencia 098 (Gris)</option>
                                        <option value="6">Negación Providencia 104 (Naranja)</option>
                                        <option value="7">Negación Providencia 108 (Morada)</option>
                                        <option value="8">Negación Providencia 119 (Azul)</option>
								    </select>
								</div>

								<div class="form-group">
								  
								     <button id="singlebutton"  name="singlebutton" class="btn btn-primary">ENVIAR</button>  
								</div>
							</div>
						</div>
					</form>
				  </div>
				</div>
			           </div>
		                 <div class="col-md-3"><!--vacio--></div>
                      </div>
          <div class="row">
<div class="text-center">
            <div class="col-md-12">
            </div>
<h4><b><span>El archivo Excel debe tener esta Estructura al momento de ser Cargado :</span></b></h4>
              <img src="captura.png">
        </div>
              </div>
                  </div>
	

</body>
</html>

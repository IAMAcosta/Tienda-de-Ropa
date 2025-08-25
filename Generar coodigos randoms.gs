function generarCodigoCompra() {
  var letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";  // Letras del alfabeto
  var numeros = "0123456789";                  // Números del 0 al 9
  var codigo = "";
  
  // Generar código de 3 letras aleatorias
  for (var i = 0; i < 3; i++) {
    codigo += letras.charAt(Math.floor(Math.random() * letras.length));
  }
  
  // Generar código de 3 números aleatorios
  for (var j = 0; j < 3; j++) {
    codigo += numeros.charAt(Math.floor(Math.random() * numeros.length));
  }
  
  return codigo;
}

function generarVariosCodigosCompra(cantidad) {
  var codigos = [];
  for (var i = 0; i < cantidad; i++) {
    codigos.push(generarCodigoCompra());
  }
  return codigos;
}

function llenarCodigosCompras() {
  var hojaCodigos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Compras_Validadas");
  
  var letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"; 
  var numeros = "0123456789"; 
  var cantidadCodigos = 10; 
  var codigosGenerados = [];

  // Generar códigos únicos
  while (codigosGenerados.length < cantidadCodigos) {
    var codigo = "";
    
    // Generar código de 3 letras aleatorias
    for (var i = 0; i < 3; i++) {
      codigo += letras.charAt(Math.floor(Math.random() * letras.length));
    }
    
    // Generar código de 3 números aleatorios
    for (var j = 0; j < 3; j++) {
      codigo += numeros.charAt(Math.floor(Math.random() * numeros.length));
    }

    // Comprobar si el código ya existe
    if (!codigosGenerados.includes(codigo)) {
      codigosGenerados.push(codigo);
    }
  }
  
  // Escribir los códigos en la hoja
  for (var k = 0; k < codigosGenerados.length; k++) {
    hojaCodigos.getRange(k + 2, 1).setValue(codigosGenerados[k]); 
  }
}





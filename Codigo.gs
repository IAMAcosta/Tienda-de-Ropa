function onFormSubmit(e) {
  var formData = e.values; // Datos del formulario enviado
  
  var fullName = formData[1];     // Nombre y Apellido
  var email = formData[2];        // Correo
  var dni = formData[3];          // DNI
  var purchaseCode = formData[4]; // Código de compra
  
  var validCodesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Compras_Validadas");
  var validCodesData = validCodesSheet.getDataRange().getValues();
  
  var codeFound = false;
  var codeUsed = false;

  // Buscar si el código está en la lista de códigos válidos
  for (var i = 1; i < validCodesData.length; i++) {
    if (validCodesData[i][0] === purchaseCode) {
      codeFound = true;
      if (validCodesData[i][1] === "Sí") {
        codeUsed = true;
      } else {
        // Marcar como usado
        validCodesSheet.getRange(i + 1, 2).setValue("Sí");
      }
      break;
    }
  }
  
  if (!codeFound) {
    MailApp.sendEmail({
      to: email,
      subject: "Error con tu código de compra",
      body: "Hola " + fullName + ", el código de compra que ingresaste no es válido. Por favor, verifica y vuelve a intentarlo."
    });
  } else if (codeUsed) {
    MailApp.sendEmail({
      to: email,
      subject: "Código de compra ya utilizado",
      body: "Hola " + fullName + ", el código de compra que ingresaste ya ha sido utilizado. Si crees que esto es un error, por favor contáctanos."
    });
  } else {
    var discounts = ["5%", "10%", "15%"];
    var randomDiscount = discounts[Math.floor(Math.random() * discounts.length)];
    
    MailApp.sendEmail({
      to: email,
      subject: "Tu descuento en la tienda de ropa",
      body: "Hola " + fullName + ", ¡gracias por tu compra! Aquí tienes tu descuento del " + randomDiscount + " en tu próxima compra."
    });
  }
}

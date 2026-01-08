// functions.js
// Archivo de inicialización para el complemento de Outlook

Office.initialize = function (reason) {
  console.log("Complemento Asistente Voz IA inicializado. Razón:", reason);
};

// Ejemplo de función que se puede invocar desde el manifiesto
function openTaskPane(event) {
  // Aquí podrías agregar lógica adicional antes de abrir el panel
  console.log("Botón 'Abrir IA' presionado.");
  
  // Siempre hay que completar el evento para que Outlook no bloquee la acción
  event.completed();
}

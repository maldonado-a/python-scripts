El script en Python automatiza la actualización de un formulario web y la simulación de envío de correos electrónicos 
basado en datos de un archivo Excel.
Lectura del Excel: Abre y lee un archivo Excel llamado "base_seguimiento.xlsx". 
Inicialización del navegador: Usa Selenium WebDriver para abrir Chrome. 
Procesamiento de filas: Si el estado es "Regularizado", completa y envía un formulario web. 
Si el estado es "Atrasado", simula el envío de un correo electrónico.(Cierre del achivo Excel)

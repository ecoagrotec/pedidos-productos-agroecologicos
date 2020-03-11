**Programa para procesar pedidos de productos agroecológicos en Google Forms/Sheets y generar planillas para los lugares de entrega**

El programa procesa un listado general de pedidos que generan los clientes a través de un formulario de Google (Google Forms) y genera listados individuales para cada uno de distintos puntos de entrega (que selecciona en cliente en el formulario de pedidos). También calcula los totales generales y por punto de entrega para los distintos productos que defina el usuario (por ejemplo bolsones de verdura, miel, huevos, etc.)

**Requisitos previos:**

1. Formulario de Google para tomar los pedidos
2. Planilla de cálculo de Google donde se almacenan los pedidos en formato de tabla
3. Agregar el código del programa (archivo _.gs_ en este repositorio) al _Editor de secuencias de comandos_ de Google Sheets
4. Agregar una nueva hoja con el nombre 'Configuración', y con el siguiente contenido (valores de la columnas B, C y D de ejemplo)

. | A | B | C | D
--- | --- | --- | --- | ---
1 | Planilla de pedidos (título completo)	| Hoja1	 
2 | Columna con lugares de entrega(título completo)	| Lugar de entrega 
3 | Columna por la cual ordenar (título completo)	| Apellido
4 | Título de las planillas de lugar de entrega (+número)	| #	
5 | Columnas para las cuales calcular totales (palabra clave, mayúsculas o minúsculas)	| Bolsones	| Miel	| Huevos

**Utilización:**

1. Copiar los datos de los pedidos a una hoja (si no están ya en la planilla)
2. Modificar los valores de la hoja 'Configuración' de acuerdo a los datos de la planilla de pedidos
3. Ir al menú 'Pedidos' y hacer clic en 'Procesar pedidos'
4. Si es la primera vez que se ejecuta, se deberá autorizar el permiso a la aplicación y volver a ejecutar 'Procesar pedidos'
5. Esperar :)

**Contacto:**

Gustavo Pereyra Irujo - pereyrairujo.gustavo@conicet.gov.ar

**Licencia:**

Este es un software de código abierto (Licencia GPL v3), por lo que cualquier persona, organización o compañía tiene la libertad de usarlo, estudiarlo, compartirlo, copiarlo y modificarlo

# pols-exceljs-plus

`pols-exceljs-plus` es una extensión de la biblioteca `exceljs` que proporciona utilidades avanzadas para agilizar la lectura, escritura y estructuración de datos en hojas de cálculo Excel.

## Características

* **PXls Workbook**: Extiende el `Workbook` nativo de exceljs inyectando métodos de ayuda en todas las hojas (`Worksheet`).
* **Lectura por Esquema Declarativo (`getValuesBySchema`)**: Extrae y valida filas o columnas a partir de un esquema en forma de objeto, con conversión automática de tipos, validación de obligatoriedad, valores por defecto y parseadores personalizados.
* **Lectura de Tablas Dinámicas (`getTableValues`)**: Extrae una tabla completa a partir de una fila de cabeceras, asociando dinámicamente columnas por nombre (soporta `RegExp`), acumulando múltiples coincidencias con `parse` y detectando el fin de la tabla automáticamente.
* **Escritura simplificada**: Métodos `setValues`, `setRowValues` y `setColumnValues` para escribir arreglos bidimensionales y unidimensionales aplicando estilos y fusiones de celdas fácilmente.
* **Lectura rápida de valores (`getValue`)**: Recuperación de valores de celdas con formateo de fechas, resolución de fórmulas y extracción de hipervínculos automáticos.

## Instalación

```bash
npm install pols-exceljs-plus
```

## Uso

### Inicializar PXls

Para utilizar las utilidades extendidas, utiliza la clase `PXls` en lugar del `Workbook` nativo de exceljs:

```typescript
import { PXls } from 'pols-exceljs-plus';

const workbook = new PXls();
await workbook.readFile('mi_archivo.xlsx');

// Obtener hoja (ya viene decorada con todos los métodos de utilidad)
const sheet = workbook.getWorksheet('Hoja1');
```

---

### Lectura por Esquema con `getValuesBySchema`

Este método permite extraer y formatear los valores de una fila o columna según un esquema definido por un objeto. El orden de las propiedades del objeto determina por defecto el orden secuencial de lectura.

#### Firma del método:
```typescript
sheet.getValuesBySchema(schema, readMode, row, column)
```

* **`schema`**: Objeto declarativo que define la forma de la respuesta esperada.
* **`readMode`**: `'row'` para leer horizontalmente (columnas consecutivas) o `'column'` para leer verticalmente (filas consecutivas).
* **`row`**: Fila de inicio (1-indexed).
* **`column`**: Columna de inicio (1-indexed).

#### Opciones del Esquema

Cada propiedad del esquema puede ser:
1. **Un Constructor / Tipo abreviado:** `String`, `Number`, `Boolean`, `Date`, `'string'`, `'number'`, `'boolean'`, `'date'`, o `'any'`.
2. **Un Objeto de Configuración Completo:**
   * `type`: Tipo de dato o constructor.
    * `cellIndex` (opcional): Desplazamiento relativo explícito desde la celda inicial (permite saltarse celdas o leer en desorden).
    * `parse` (opcional): Función callback para transformar/limpiar el valor obtenido: `(value: any) => any`.

#### Ejemplo de Lectura:

```typescript
// Supongamos que en la fila 5 de Excel tenemos:
// Col 1 (A): 456
// Col 2 (B): "María Pérez"
// Col 3 (C): 2026-06-24 (Fecha)
// Col 4 (D): (Celda vacía)
// Col 5 (E): "  texto con espacios  "

const esquemaCliente = {
  id: Number,                                                   // Convierte a número secuencialmente (Col A)
  nombre: String,                                               // Convierte a string (Col B)
  fechaRegistro: 'date',                                        // Convierte a objeto Date (Col C)
  estado: { type: String, cellIndex: 3 },                       // Lee Col D (vacía) y devuelve null
  observaciones: { 
    parse: (val) => typeof val === 'string' ? val.trim() : val,
    cellIndex: 4                                                // Lee Col E de forma explícita
  }
};

const cliente = sheet.getValuesBySchema(esquemaCliente, 'row', 5, 1);
console.log(cliente);
/*
Output:
{
  id: 456,
  nombre: "María Pérez",
  fechaRegistro: Date("2026-06-24..."),
  estado: null,
  observaciones: "texto con espacios"
}
*/
```


---

### Lectura de Tablas con `getTableValues`

Este método permite leer una tabla completa a partir de una fila de cabeceras. Busca dinámicamente las columnas basándose en sus nombres de cabecera definidos en el esquema (soporta coincidencia por texto exacto o expresión regular `RegExp`).

#### Firma del método:
```typescript
sheet.getTableValues(schema, row, column)
```

* **`schema`**: Objeto que define las propiedades y cómo encontrarlas/formatearlas. Cada propiedad debe ser un objeto con los siguientes atributos:
  * `headerName`: Cadena de texto (`string`) o expresión regular (`RegExp`) para identificar la cabecera de la columna correspondiente.
  * `type` (opcional): Constructor o tipo de conversión (`String`, `Number`, `Boolean`, `Date`, `'string'`, `'number'`, `'boolean'`, `'date'`, o `'any'`).
  * `parse` (opcional): Función callback para transformar el valor: `(value: any, prevValue?: any) => any`.
* **`row`**: Fila donde se encuentra la cabecera (1-indexed).
* **`column`**: Columna inicial desde donde se empezará a buscar las cabeceras de forma horizontal (1-indexed).

#### Características particulares:
1. **Detección de Cabeceras**: El método recorre horizontalmente la fila de partida (`row`) desde la columna inicial (`column`) hasta toparse con una celda vacía (lo cual da por terminado el escaneo de cabeceras).
2. **Detección de Fin de Tabla**: Lee hacia abajo fila por fila. Para optimizar el rendimiento, sólo lee las celdas cuyas cabeceras coinciden con algún `headerName`. Detiene su lectura cuando se topa con una fila completamente vacía en todas las columnas identificadas.
3. **Múltiples Cabeceras para una propiedad**:
   * Si el `headerName` coincide con más de una columna (por ejemplo `/(base imponible)|impuesto/`), por defecto se tomará el valor de la última columna procesada (de izquierda a derecha).
   * Si se define la función `parse`, esta se invocará secuencialmente recibiendo como primer parámetro el valor de la columna actual y como segundo parámetro el valor acumulado de las columnas coincidentes previas, permitiendo realizar agregaciones (por ejemplo, sumas o concatenaciones).

#### Ejemplo de Lectura de Tabla:

```typescript
// Supongamos que en la fila 2 de Excel tenemos las cabeceras:
// B2: "Nombres", C2: "Edades", D2: "Monto 1", E2: "Monto 2"
// Y las filas siguientes contienen la información

const schema = {
  nombre: {
    type: 'string',
    headerName: /Nombres/
  },
  montoTotal: {
    type: 'number',
    headerName: /Monto/,
    parse: (val, prevVal) => (prevVal || 0) + (val || 0) // Suma las columnas que coincidan con "Monto"
  }
};

const datos = sheet.getTableValues(schema, 2, 2);
console.log(datos);
/*
Output:
[
  { nombre: "Juan", montoTotal: 150 },
  { nombre: "Maria", montoTotal: 350 }
]
*/
```

---

### Escritura Simplificada

Escribe matrices o arreglos lineales de manera ágil usando los métodos de escritura extendida:

```typescript
// Escribir múltiples filas y columnas de una sola vez
sheet.setValues(1, 1, [
  [1, "Producto A", 19.99],
  [2, "Producto B", 25.50]
]);

// Escribir una sola fila con un estilo por defecto
sheet.setRowValues(3, 1, ["ID", "Descripción", "Precio"], {
  backgroundColor: "E0E0E0",
  color: "000000",
  span: 1
});

// Escribir una columna
sheet.setColumnValues(4, 1, [100, 200, 300]);
```

### Lectura Rápida de Celdas

```typescript
// Obtiene el valor parseado automáticamente resolviendo fechas, fórmulas e hipervínculos
const valor = sheet.getValue<string>(1, 2);
```

## Licencia

Este proyecto está bajo la licencia ISC.

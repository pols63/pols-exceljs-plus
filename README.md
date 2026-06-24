# pols-exceljs-plus

`pols-exceljs-plus` es una extensión de la biblioteca `exceljs` que proporciona utilidades avanzadas para agilizar la lectura, escritura y estructuración de datos en hojas de cálculo Excel.

## Características

* **PXls Workbook**: Extiende el `Workbook` nativo de exceljs inyectando métodos de ayuda en todas las hojas (`Worksheet`).
* **Lectura por Esquema Declarativo (`getValuesBySchema`)**: Extrae y valida filas o columnas a partir de un esquema en forma de objeto, con conversión automática de tipos, validación de obligatoriedad, valores por defecto y parseadores personalizados.
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

# walk-excel

Lector eficiente de archivos Excel (XLS/XLSX) para Node.js, con soporte para archivos grandes y streaming de datos.

## ✨ Características:

- ✅ Soporta archivos XLS/XLSX
- 📂 Procesamiento optimizado para archivos grandes (evita cargar todo en memoria).
- 🔄 API simple para iterar fila por fila o obtener datos en JSON.

## 🚀 Instalación

```bash
npm install walk-excel
```

## 📦 Dependencias

- `exceljs`: Para soporte de `.xlsx` con lectura de archivos grandes (streaming nativo).
- `xlsx`: Para soporte de `.xls` (hasta ahora sin streaming).

## 📖 Uso básico

```js
import { ExcelModule } from "walk-excel";

const excel = new ExcelModule();

excel.walkRows("./data.xlsx", (row, index) => {
  console.log("data:", row, "index:", index); // data: {...} index: 0
});
```

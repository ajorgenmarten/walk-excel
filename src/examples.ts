import { ExcelModule } from './utilities';

const excel = new ExcelModule();

excel.walkRows(
  './files/Listado-de-proveedores-y-contactos.xls',
  (row, index) => {
    console.log(`indice: ${index}`, row);
  }
);

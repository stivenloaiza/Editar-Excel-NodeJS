
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

let now= new Date();

//Abrimos el excel
workbook.xlsx.readFile('formato1.xlsx')
    .then(function() {
        //definimos la hoja de trabajo, si es un numero el numero, si es con nombre el nombre
        var worksheet = workbook.getWorksheet('Hoja1');
        // Asignamos un valor en una posicion especifica: getRow= fila, getCell= Columna
        var row = worksheet.getRow(9).getCell(7).value=0;
        row = worksheet.getRow(10).getCell(7).value="Inserte este texto en la fila 10 columna 7";
        row = worksheet.getRow(10).getCell(22).value="1234569";

        //Creamos un nuevo archivo con los cambios aplicados
        var fechaActual= now.getDate()+"-"+(now.getMonth()+1)+"-"+now.getFullYear()+"-"+now.getHours()+"-"+now.getMinutes();
        return workbook.xlsx.writeFile('formato-cambiado'+fechaActual+'.xlsx');
    }).catch(function (e){
        console.log("Error: "+e);
    });
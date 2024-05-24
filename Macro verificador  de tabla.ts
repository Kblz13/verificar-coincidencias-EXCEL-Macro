function verificar(columnay: string, i: number, workbook: ExcelScript.Workbook): boolean {
    // Obtener la hoja activa
    let sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();
    let range: ExcelScript.Range = sheet.getUsedRange();
    let cantidadfilas: number = range.getRowCount();
    let cambios: boolean = false;

    
    let maximo:number= 320//en esta | variable puedes poner el maximo filas que tiene la tabla id o puede usar cantidadfilas

    // Recorrer todas las filas en la columna AB para comparar con el valor de Y
    for (let j = 0; j < maximo; j++) {
        let valueAB: string = sheet.getCell(j, 27).getText(); // en getcell colocas tu fila/columna que necesites


        // Valida que valueAB no esté vacío
        if (valueAB === "") {
            // console.log(`Fila ${j}: Valor en AB está vacío, omitiendo...`);
            continue;
        }
     //   console.log(`FIla ${valueAB} subfuncion :comparando con fila ${columnay}`)
        // Comparar los valores
        if (valueAB === columnay) {
            console.log(`Fila ${i}: Encontrada coincidencia en fila ${j}, copiando valor AA`);
            let valueAA: string = sheet.getCell(j, 26).getText(); 
            sheet.getCell(i, 23).setValue(valueAA); 
            break; // Salir del bucle interior

        }
    }

}

function main(workbook: ExcelScript.Workbook): void {
    // Obtiene la hoja activa
    let sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();
    let range: ExcelScript.Range = sheet.getUsedRange();
    let lastRow: number = range.getRowCount();
    console.log(`Total de filas utilizadas: ${lastRow}`);

    // Recorre cada fila en la columna Y y ejecutar la función para hacer cambios si hay idénticos 
    //filamain es la fila de la tabla principal
    for (let FilaMAIN = 1; FilaMAIN < lastRow; FilaMAIN++) { 
        let valueY: string = sheet.getCell(FilaMAIN, 24).getText();
        verificar(valueY, FilaMAIN, workbook);
        console.log(`estamos funcion principal Fila ${FilaMAIN}:`);
        // Recorre las filas en incrementos de 317

    }
}

console.log("Proceso completado.");


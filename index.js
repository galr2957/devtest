const ExcelJS = require('exceljs');

const fillCell = (cell, color= "FFFFFF") => {
    cell.style = {
        ...cell.style,
        fill: {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb: color},
          },
          border: {
            left: { style: 'thin' },
            right: { style: 'thin' },
            top: { style: 'thin' },
            bottom: { style: 'thin' }
          },
      }
    return cell
}

const inserToMap = (map, department, value) => {
    if (map.has(department)) {
        map.set(department, map.get(department) + value)
    }
    else {
        map.set(department, value)
    }
}

const main = async() => {
    const Departments = new Map()
    const rentInfoArr = []
    const workbook = new ExcelJS.Workbook();  
    try {
    await workbook.xlsx.readFile("DevExercise.xlsx")
    
    workbook.eachSheet( (worksheet, sheetId) => {
        var shouldBeFixedCells = []
        const sheetSize = {columns: worksheet.columnCount, 
                           rows:  worksheet.actualRowCount}
        var maxRow = {sum: 0, rowId: null}
        worksheet.getRow(1).getCell(sheetSize.columns+1).value = "SUM"

        for (let rowIt =2; rowIt<= sheetSize.rows; rowIt++) {
            let row = worksheet.getRow(rowIt);
            let rowSum = 0;
            for (let colIt=3; colIt <= sheetSize.columns; colIt++) {
                let cell = row.getCell(colIt)
                if (typeof(cell.value) != "number") {
                    console.log(`cell in column ${colIt} and row ${rowIt} is not a number or empty`)
                    shouldBeFixedCells.push(cell)
                    continue;
                }
                rowSum += cell.value;
            }
            let departmentNameCell = row.getCell(2)
            if (typeof(departmentNameCell.value) === "string" && departmentNameCell.value.length > 0) {
                if (departmentNameCell.value.trim().toLowerCase() === "rent") {
                    var rentRowPointer = row
                    rentInfoArr.push({ 
                        location: row.getCell(1).value, 
                        value: rowSum }
                        )
                }
                    inserToMap( Departments, 
                                departmentNameCell.value.trim(),
                                rowSum)
            } else {
                console.log(`the value in row ${rowIt} and column 2 should be a department name(sttring)`)
                shouldBeFixedCells.push(departmentNameCell)
            }
            fillCell(row.getCell(sheetSize.columns +1)).value = rowSum;
            maxRow = rowSum > maxRow.sum ? {sum: rowSum, rowId: rowIt} : maxRow;
        }; 
        if(shouldBeFixedCells.length) {
            shouldBeFixedCells.forEach(cell => fillCell(cell, "F08080"))
            workbook.xlsx.writeFile("output_file_with_errors.xlsx"); 
            throw new Error(`in worksheet ${worksheet.name} typing errors has been found. please visit file 'output_file_with_errors.xlsx'`)
        }        
         // bgcolor to the row with the max sum 
         worksheet.getRow(maxRow.rowId).eachCell((cell, cellId) => {
            fillCell(cell, "A4BBD2")
        })
        //yellow bg for the Rent row
        if (rentRowPointer) {
            rentRowPointer.eachCell((cell, cellId) => {
                fillCell(cell, "EEE8AA")
            })
        }  
    });
    // writing results to new sheet "RentTotal"
    const rentTotalSheet = workbook.addWorksheet('RentTotal');
    const headerRow = rentTotalSheet.addRow(["REGION", "SUM"])
    rentInfoArr.forEach(({location, value}) => {
        rentTotalSheet.addRow([location, value])
    })
    // // writing results to new sheet "Departments_Total"
    const TotalSheet = workbook.addWorksheet('Departments_Total', {properties:{defaultColWidth: 12}});
    TotalSheet.addRow(["DEPARTMENT", "SUM"])
    Departments.forEach((value, key) => {
        TotalSheet.addRow([key, value])
    })
    workbook.xlsx.writeFile("output.xlsx")
    
} catch (error) {
        console.log(error.message)
    }
}
main();

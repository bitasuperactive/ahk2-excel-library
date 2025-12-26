#Requires AutoHotkey v2.0
#Include "WorkbookWrapper.ahk"
#Include "..\..\Util\Utils.ahk"

/************************************************************************
 * @class WriteWorkbookAdapter
 * @brief Adaptador dedicado a la escritura en libros de trabajo.
 * @author bitasuperactive
 * @date 25/12/2025
 * @version 0.9.1-Beta
 * @extends WorkbookWrapper
 * @warning Dependencias:
 * - WorkbookWrapper.ahk
 * - Utils.ahk
 * @see https://github.com/bitasuperactive/ahk2-excel-library/blob/master/ExcelLibrary/ExcelBridge/WriteWorkbookAdapter.ahk
 ***********************************************************************/
class WriteWorkbookAdapter extends WorkbookWrapper
{
    /**
     * @public
     * Crea un adaptor para la escritura en una de las hojas de cálculo
     * de un libro de trabajo específico.
     * 
     * - Envuelve los datos preexistentes en una tabla para facilitar su delimitación.
     * - Elimina las filas vacías para facilitar la escritura.
     * 
     * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo objetivo.
     * @param {Microsoft.Office.Interop.Excel.Worksheet} targetSheet (Opcional) Hoja de cálculo objetivo.
     * Por defecto, será la hoja de cálculo activa en el libro objetivo.
     * @throws {TargetError} (0x80010108) Si el libro de trabajo objetivo se encuentra cerrado.
     * @throws {Error} (0x80010001) Si Microsoft Excel rechaza la conexión a su interfaz.
     * @throws {ValueError} Si existe más de tabla definida en la hoja de cálculo objetivo.
     */
    __New(workbook, targetSheet?)
    {
        super.__New(workbook, targetSheet?)
        this._DeleteEmptyRows()
    }

    /**
     * @public
     * Crea o anexa una tabla, sean solo filas o solo columnas en la hoja de cálculo objetivo.
     * @warning Si utilizas la clase nativa Object para encapsular los datos a introducir, 
     * se impondrá un orden alfabético para la inserción de las columnas.
     * Se recomienda utilizar OrObject.
     * @param {Array<Object>} objArray Colección de objetos literales.
     * @param {Array<String>} expectedHeaders (Opcional) Colección de los nombres 
     * para las cabeceras requeridas.
     * @throws {UnsetError} Si la tabla no tiene alguna de las cabeceras esperadas.
     */
    AppendTable(objArray, expectedHeaders := [])
    {
        ;// Si no hay datos, crear tabla
        if (this.IsTargetSheetEmpty()) {
            this._CreateTable(objArray)
            return
        }

        ;// Validación (~0.25s)
        if (!IsObject(objArray))
            throw TypeError("Se esperaba un Object pero se ha recibido: " Type(objArray))
        if (Type(objArray) != "Array")
                objArray := [objArray]
        if (objArray.Length = 0)
            return
        if (!IsObject(objArray[1]))
            throw TypeError("Los elementos de la colección deben ser objetos, pero contiene: " Type(objArray[1]))
        if (!this.ValidateHeaders(expectedHeaders, &missingHeaders))
            throw UnsetError("La tabla es inválida, no dispone de las siguientes cabeceras requeridas: " Utils.ArrayToString(missingHeaders))
        
        ;// Normalizar
        ;// Crear colección de cabeceras con todas las propiedades únicas de la colección de objetos facilitada
        headerArr := []
        for obj in objArray {
            obj := WorkbookWrapper._NormalizeObjProps(obj)
            objArray[A_Index] := obj
            for prop in obj.OwnProps() {
                if (!Utils.ArrHasVal(headerArr, prop))
                    headerArr.Push(prop)
            }
        }

        sheet := this._targetSheet
        ;// Anexar las cabeceras que falten en la tabla
        firstUsedRowIndex := this._GetTargetRange().Row
        firstUsedColIndex := this._GetTargetRange().Column
        headerRow := this._GetRowSafeArray(1)
        for (header in headerArr) {
            Loop headerRow.MaxIndex(2) {
                if (headerRow[1, A_Index] = header) ; Los Arrays de Interop son 1-based
                    continue 2
            }
            nextFreeHeaderCell := sheet.Cells(
                firstUsedRowIndex, 
                firstUsedColIndex + this.GetColumnCount()
            )
            nextFreeHeaderCell.Value2 := header
        }
        headerRow := this._GetRowSafeArray(1) ; Actualizar valor

        ;// Calcular filas a añadir (los objetos sin ningún valor no cuentan)
        rowsToAdd := 0
        for obj in objArray {
            for _, val in obj.OwnProps() {
                if (val != "") {
                    rowsToAdd++
                    continue 2
                }
            }
        }
        if (rowsToAdd = 0)
            rowsToAdd := 1 ; Necesario para nulificar los -1

        ;// SafeArray a insertar
        newRowCount := this.GetRowCount() + objArray.Length - 1
        newColCount := this.GetColumnCount()
        safeArray := WorkbookWrapper._CreateInteropArray(newRowCount, newColCount) ;ComObjArray(VT_VARIANT:=12, newRowCount, newColCount) ; 0-based
        for obj in objArray {
            iRow := A_Index
            Loop headerRow.maxIndex(2) {
                iCol := A_Index
                header := headerRow[1, A_Index]
                value := (ObjHasOwnProp(obj, header)) ? obj.%header% : ""
                safeArray[iRow, iCol] := value
            }
        }

        ;// Inserción
        targetUpperLeftCell := sheet.Cells(
            firstUsedRowIndex + this.GetRowCount(),
            firstUsedColIndex
        )
        targetLowerRightCell := sheet.Cells(
            targetUpperLeftCell.Row + objArray.Length - 1,
            targetUpperLeftCell.Column + newColCount - 1
        )
        sheet.Range(targetUpperLeftCell, targetLowerRightCell).Value2 := safeArray
        
        this._WrapTargetRangeInTable(1)
    }

    /**
     * @public 
     * Rellena los espacios blancos de una fila.
     * @note Las propiedades del objeto a introducir serán utilizadas para validar 
     * las cabeceras de la tabla.
     * @param {Integer} row Índice de la fila objetivo.
     * @param {Object} obj Objeto fuente de los datos a utilizar.
     * @throws {TargetError} Si la tabla no tiene las mismas cabeceras que propiedades tiene el objeto.
     * @throws {UnsetError} Si la tabla no tiene alguna de las cabeceras esperadas.
     * @throws {ValueError} Si la fila objetivo está fuera del rango utilizado.
     */
    FillBlankFieldsOnRow(row, obj)
    {
        if (Type(row) != "Integer")
            throw TypeError("Se esperaba un Integer pero se ha recibido: " Type(row))
        if (!IsObject(obj))
            throw TypeError("Se esperaba un Object pero se ha recibido: " Type(obj))
        for (prop in obj.OwnProps())
            if (IsObject(prop))
                throw TypeError("Las propiedades del objeto deben ser valores primitivos pero contiene: " Type(prop))
        if (row < 1 || row > this.GetRowCount())
            throw ValueError('La fila {' row '} está fuera del rango utilizado.')

        obj := WorkbookWrapper._NormalizeObjProps(obj)
        objProps := []
        for prop in obj.OwnProps() 
            objProps.Push(prop)
        if (!this.ValidateHeaders(objProps, &missingHeaders))
            throw TargetError("La tabla no dispone de las cabeceras requeridas por el objeto facilitado. "
                                "Para añexar las cabeceras faltantes, utiliza la función AppendTable.")
        
        headerRow := this._GetRowSafeArray(1)
        targetRow := this._GetRowSafeArray(row)
        Loop headerRow.maxIndex(2) {
            cellValue := targetRow[1, A_Index]
            if (cellValue = "") {
                header := headerRow[1, A_Index]
                value := (ObjHasOwnProp(obj, header)) ? obj.%header% : ""
                this._GetTargetRange().Cells(row, A_Index).Value2 := value
            }
        }
    }

    ;// TODO - MAYBE: Guardar el objeto de la última fila eliminada y en caso de CTRL+Z, restaurarla
    /**
     * @public
     * Elimina la fila solicitada.
     * @param {Integer} row Índice de la fila objetivo.
     * @param {Object} expectedObj (Opcional) Objeto de validación para la fila objetivo.
     * Tanto las cabeceras de la tabla como su contenido debe coincidir con las 
     * propiedades y valores del objeto.
     * @param {Array<String>} expectedHeaders (Opcional) Colección de los nombres 
     * para las cabeceras requeridas. Útil si solo se requiere validar las cabeceras y no los datos.
     * @throws {TargetError} Si la fila a eliminar no coincide con el objeto de validación.
     * @throws {UnsetError} Si la tabla no tiene alguna de las cabeceras esperadas.
     */
    DeleteRow(row, expectedObj?, expectedHeaders := [])
    {
        if (Type(row) != "Integer")
            throw TypeError("Se esperaba un Integer pero se ha recibido: " Type(row))
        if (row < 1 || row > this.GetRowCount())
            throw ValueError('La fila {' row '} está fuera del rango utilizado.')
        if (row = 1)
            throw ValueError('No es posible eliminar la fila de cabeceras. Utiliza DeleteTable() para eliminar la tabla.')
        if (IsSet(expectedObj) && !IsObject(expectedObj))
            throw TypeError("Se esperaba un Object pero se ha recibido: " Type(expectedObj))
        if (!this.ValidateHeaders(expectedHeaders, &missingHeaders))
            throw UnsetError("La tabla es inválida, no dispone de las siguientes cabeceras requeridas: " Utils.ArrayToString(missingHeaders))
        
        ;// Confirmar la fila a eliminar
        if (IsSet(expectedObj)) {
            expectedObj := WorkbookWrapper._NormalizeObjProps(expectedObj)
            headerRow := this._GetRowSafeArray(1)
            targetRow := this._GetRowSafeArray(row)

            Loop headerRow.MaxIndex(2) {
                header := headerRow[1, A_Index]
                value := targetRow[1, A_Index]
                if (!ObjHasOwnProp(expectedObj, header) || expectedObj.%header% != value)
                    throw TargetError("La fila a eliminar no coincide con el objeto de validación.")
            }
        }

        row := this._GetTargetRange().Rows[row]
        row.EntireRow.Delete()
    }

    /**
     * @public
     * Elimina todo el rango objetivo.
     * @param {Array<String>} expectedHeaders (Opcional) Colección de los nombres 
     * para las cabeceras requeridas.
     * @throws {UnsetError} Si la tabla no tiene alguna de las cabeceras esperadas.
     */
    DeleteTable(expectedHeaders := [])
    {
        if (!this.ValidateHeaders(expectedHeaders, &missingHeaders))
            throw UnsetError("La tabla es inválida, no dispone de las siguientes cabeceras requeridas: " Utils.ArrayToString(missingHeaders))
        
        this._GetTargetRange().Delete()
    }

    /**
     * @private 
     * Crea una tabla a partir de la colección de objetos indicada.
     * @warning Si utilizas la clase nativa Object para encapsular los datos a introducir, 
     * se impondrá un orden alfabético para la inserción de las columnas.
     * Se recomienda utilizar OrObject.
     * @param {Array<Object>} objArray Colección de objetos literales.
     */
    _CreateTable(objArray)
    {
        ;// Validación
        if (!IsObject(objArray))
            throw TypeError("Se esperaba un Object pero se ha recibido: " Type(objArray))
        if (Type(objArray) != "Array")
                objArray := [objArray]
        if (objArray.Length = 0)
            return
        if (!IsObject(objArray[1]))
            throw TypeError("Los elementos de la colección deben ser objetos, pero contiene: " Type(objArray[1]))
        
        ;// Normalizar objetos
        ;// Crear objeto de cabeceras con todas las propiedades únicas de la colección de objetos
        headerArr := []
        for obj in objArray {
            obj := WorkbookWrapper._NormalizeObjProps(obj)
            objArray[A_Index] := obj
            for prop in obj.OwnProps() {
                if (!Utils.ArrHasVal(headerArr, prop))
                    headerArr.Push(prop)
            }
        }

        rows := objArray.Length + 1 ; Más la fila de cabeceras
        cols := headerArr.Length
        safeArray := WorkbookWrapper._CreateInteropArray(rows, cols)

        for header in headerArr {
            safeArray[1, A_Index] := header
        }
        for obj in objArray {
            iRow := A_Index + 1 ; Saltar encabezados
            for header in headerArr {
                iCol := A_Index
                value := (ObjHasOwnProp(obj, header)) ? obj.%header% : ""
                safeArray[iRow, iCol] := value
            }
        }

        sheet := this._targetSheet
        targetUpperLeftCell := this._GetTargetRange().Cells(1,1)
        targetLowerRightCell := sheet.Cells(
            (targetUpperLeftCell.Row - 1) + rows,
            (targetUpperLeftCell.Column - 1) + cols
        )
        targetRange := sheet.Range(targetUpperLeftCell, targetLowerRightCell)
        targetRange.Value2 := safeArray
        
        this._WrapTargetRangeInTable(1)
    }
}
#Requires AutoHotkey v2.0
#Include "WorkbookWrapper.ahk"
#Include "..\..\Util\OrObject.ahk"
#Include "..\..\Util\Utils.ahk"

/************************************************************************
 * @class ReadWorkbookAdapter
 * @brief Adaptador dedicado a la lectura de libros de trabajo.
 * 
 * - Conceptualizada para no alterar los datos del libro 
 * (excepto las cabeceras que se normalizan).
 * 
 * @author bitasuperactive
 * @date 25/12/2025
 * @version 0.9.1-Beta
 * @extends WorkbookWrapper
 * @warning Dependencias: 
 * - WorkbookWrapper.ahk
 * - OrObject.ahk
 * - Utils.ahk
 * @see https://github.com/bitasuperactive/ahk2-excel-library/blob/master/ExcelLibrary/ExcelBridge/ReadWorkbookAdapter.ahk
 ***********************************************************************/
class ReadWorkbookAdapter extends WorkbookWrapper
{
    /**
     * @public
     * Crea un adaptor para la lectura de una de las hojas de cálculo
     * de un libro de trabajo específico.
     * 
     * - Envuelve los datos preexistentes en una tabla para facilitar su delimitación.
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
    }

    /**
     * @public 
     * Lee la tabla de la hoja de cálculo objetivo.
     * @param {Array<String>} expectedHeaders (Opcional) Colección de los nombres de las cabeceras esperadas.
     * @returns {Array<Object>} Colección de objetos literales representativa de la
     * tabla objetivo. Sus atributos corresponden con los encabezados de la tabla.
     * @throws {UnsetError} Si la tabla no contiene alguna de las cabeceras esperadas.
     */
    ReadTable(expectedHeaders := [])
    {
        if (!this.ValidateHeaders(expectedHeaders, &missingHeaders))
            throw UnsetError("La tabla es inválida, no dispone de las siguientes cabeceras requeridas: " Utils.ArrayToString(missingHeaders))
        
        objArray := []
        range := this._GetTargetRange().Value2
        Loop range.MaxIndex(1) {
            obj := OrObject()
            rowIndex := A_Index
            Loop range.maxIndex(2) {
                colIndex := A_Index
                header := range[1, colIndex]
                value := range[rowIndex, colIndex]
                obj.%header% := value
            }
            objArray.Push(obj)
        }
        return objArray
    }

    /**
     * @public 
     * Lee la fila solicitada de la hoja de cálculo objetivo.
     * @param {Integer} row Índice de la fila objetivo.
     * @param {Array<String>} expectedHeaders (Opcional) Colección de los nombres de las cabeceras esperadas.
     * @returns {Object} Objeto literal representativo de la fila objetivo. Sus atributos corresponden con 
     * los encabezados de la tabla.
     * @throws {ValueError} Si la fila objetivo está fuera del rango utilizado.
     * @throws {UnsetError} Si la tabla no contiene alguna de las cabeceras esperadas.
     */
    ReadRow(row, expectedHeaders := [])
    {
        obj := OrObject()
        if (Type(row) != "Integer")
            throw TypeError("Se esperaba un Integer pero se ha recibido: " Type(row))
        if (row < 1 || row > this.GetRowCount())
            throw ValueError('La fila {' row '} está fuera del rango utilizado.')
        if (!this.ValidateHeaders(expectedHeaders, &missingHeaders))
            throw UnsetError("La tabla es inválida, no dispone de las siguientes cabeceras requeridas: " Utils.ArrayToString(missingHeaders))
        
        headerRow := this._GetRowSafeArray(1)
        targetRow := this._GetRowSafeArray(row)

        Loop headerRow.MaxIndex(2) {
            header := headerRow[1, A_Index]
            value := targetRow[1, A_Index]
            if (header = "")
                continue
            obj.%header% := value
        }
        return obj
    }
}
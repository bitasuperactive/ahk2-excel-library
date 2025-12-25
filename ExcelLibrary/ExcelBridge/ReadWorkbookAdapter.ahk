#Requires AutoHotkey v2.0
#Include "WorkbookWrapper.ahk"
#Include "..\..\Util\Utils.ahk"
#Include "..\..\Util\OrObject.ahk"

/************************************************************************
 * @class ReadWorkbookAdapter
 * @brief Funciones dedicadas a la lectura de libros de trabajo.
 * 
 * Conceptualizada para no alterar los datos del libro (excepto las cabeceras que se normalizan).
 * 
 * @author bitasuperactive
 * @date 17/12/2025
 * @version 0.9.0-Beta
 * @extends WorkbookWrapper
 * @see null
 * @note Dependencias: 
 * - WorkbookWrapper.ahk
 * - OrObject.ahk
 * - Utils.ahk
 ***********************************************************************/
class ReadWorkbookAdapter extends WorkbookWrapper
{
    /**
     * @public
     * Crea un adaptor para la lectura de una de las hojas de cálculo
     * de un libro de trabajo específico.
     * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo objetivo.
     * @param {Microsoft.Office.Interop.Excel.Worksheet} targetSheet (Opcional) Hoja de cálculo objetivo.
     * Por defecto, será la hoja de cálculo activa en el libro objetivo.
     * @throws {TargetError} Si el libro de trabajo objetivo se encuentra cerrado.
     * @throws {Error} Si Excel rechaza la conexión a su interfaz.
     */
    __New(workbook, targetSheet?)
    {
        super.__New(workbook, targetSheet?)
        ;this._DeleteEmptyRows()
    }

    /**
     * @public 
     * Lee la tabla de la hoja de cálculo objetivo.
     * @param {Array<String>} expectedHeaders (Opcional) Colección de los nombres de las cabeceras esperadas.
     * @returns {Array<Object>} Colección de objetos literales representativa de la
     * tabla objetivo, los encabezados son los atributos de los objetos.
     * 
     * @note Al utilizar objetos AHK para encapsular los datos, no se respeta el orden de la tabla.
     * Se recomienda utilizar OrObject.
     * 
     * @throws {TargetError} Si el libro de trabajo objetivo se encuentra cerrado.
     * @throws {Error} Si Excel rechaza la conexión al libro de trabajo objetivo.
     * @throws {UnsetError} Si el rango objetivo está vacío.
     * @throws {UnsetError} Si la tabla no contiene alguna de las cabeceras esperadas.
     */
    ReadTable(expectedHeaders := [])
    {
        if (!this.ValidateHeaders(expectedHeaders, &missingHeaders))
            throw UnsetError("La tabla es inválida, no dispone de las siguientes cabeceras requeridas: " Utils.ArrayToString(missingHeaders))

        this._NormalizeTableHeaders()
        
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
     * @returns {Object} Objeto literal representativo de la fila objetivo, 
     * los encabezados de la tabla serán los atributos del objeto.
     * 
     * (!) Al utilizar objetos para encapsular los datos, no se respeta el orden de la tabla.
     * 
     * @throws {TargetError} Si el libro de trabajo objetivo se encuentra cerrado.
     * @throws {Error} Si Excel rechaza la conexión al libro de trabajo objetivo.
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

        this._NormalizeTableHeaders()
        
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
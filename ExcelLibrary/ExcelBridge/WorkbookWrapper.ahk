#Requires AutoHotkey v2.0
#Include "..\..\Util\Utils.ahk"

/************************************************************************
 * @class WorkbookWrapper
 * @brief Funciones dedicadas a la administración general de
 * libros de trabajo y sus hojas de cálculo.
 * 
 * Es capaz de escapar la edición del usuario si impide la conexión con
 * la interfaz de Excel.
 * 
 * @author bitasuperactive
 * @date 21/12/2025
 * @version 0.9.0-Beta
 * @see null
 * @note Dependencias:
 * - Utils.ahk
 ***********************************************************************/
class WorkbookWrapper
{
    /** @protected */
    _name := unset ;
    /** @protected */
    _workbook := unset ;
    /** @protected */
    _targetSheet := unset ;
    /** @protected */
    _targetSheetName := unset ;
    /** @private */
    __lastHighlightedRow := 0 ;

    /**
     * @public
     * Nombre del libro de trabajo objetivo.
     * @type {String}
     */
    Name => this._name ;

    /**
     * @public
     * Nombre de la hoja de cálculo objetivo.
     * @type {String}
     */
    TargetSheetName => this._targetSheetName ;

    /**
     * @public
     * Crea un envoltorio para la administración de un libro de trabajo específico
     * y una de sus hojas de cálculo.
     * 
     * - Envuelve los datos preexistentes en una tabla para facilitar su manejo.
     * 
     * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo objetivo.
     * @param {Microsoft.Office.Interop.Excel.Worksheet} targetSheet (Opcional) Hoja de cálculo objetivo.
     * Por defecto, será la hoja de cálculo activa en el libro.
     * @throws {TargetError} Si el libro de trabajo objetivo se encuentra cerrado.
     * @throws {Error} Si Excel rechaza la conexión a su interfaz.
     */
    __New(workbook, targetSheet?)
    {
        if (!(workbook is ComObject) || Type(workbook) != "Workbook")
            throw TypeError('Se esperaba el tipo "ComObject.Workbook", pero se ha recibido: ' Type(workbook))
        if (IsSet(targetSheet) && (!(targetSheet is ComObject) || Type(targetSheet) != "Worksheet"))
            throw TypeError('Se esperaba el tipo "ComObject.Worksheet", pero se ha recibido: ' Type(targetSheet))

        Utils.ProxyObjFuncs(this, this.__InvokeExcelSafelyDelayed)
        
        this.__ThrowIfWorkbookIsInvalid(workbook)
        
        this._workbook := workbook
        this._name := workbook.Name
        this._targetSheet := (IsSet(targetSheet)) ? targetSheet : workbook.ActiveSheet
        this._targetSheetName := this._targetSheet.Name

        this._WrapTargetRangeInTable()
    }

    /**
     * @public
     * Comprueba si el libro de trabajo objetivo está abierto y accesible.
     * @returns {Boolean}
     */
    IsAvailable()
    {
        try {
            this.__ThrowIfWorkbookIsInvalid()
            return true
        } catch {
            return false
        }
    }

    /**
     * @public
     * Comprueba si el libro de cálculo objetivo está bloqueado.
     * @returns {Boolean}
     */
    IsWorkbookLocked() => this._workbook.ProtectStructure ;

    /**
     * @public
     * Comprueba si la hoja de cálculo objetivo está bloqueada.
     * @returns {Boolean}
     */
    IsSheetLocked() => this._targetSheet.ProtectScenarios ;

    /**
     * @public
     * Obtiene el número de filas utilizadas en el rango objetivo.
     * @returns {Integer}
     */
    GetRowCount() => this._GetTargetRange().Rows.Count ;

    /**
     * @public
     * Obtiene el número de columnas utilizadas en el rango objetivo.
     * @returns {Integer}
     */
    GetColumnCount() => this._GetTargetRange().Columns.Count ;

    /**
     * @public
     * Comprueba si no hay rango objetivo: No hay tablas definidas ni valor en la celda "A1".
     * @returns {Boolean}
     */
    IsTargetRangeEmpty() => this._GetTargetRange().Address = "$A$1" ;

    /**
     * @public
     * Comprueba si un libro de trabajo coincide con el libro objetivo.
     * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo a comparar.
     * @returns {Boolean} Verdadero si son equivalentes, Falso en su defecto.
     */
    IsTargetWorkbook(workbook)
    {
        if (!(workbook is ComObject) || Type(workbook) != "Workbook")
            throw TypeError('Se esperaba el tipo "ComObject.Workbook", pero se ha recibido: ' Type(workbook))
        
        return this._workbook = workbook
    }

    /**
     * @public 
     * Desactiva o reactiva las optimizaciones de rendimiento en Excel.
     * 
     * Principalmente, mejora el rendimiento de la escritura, llegando a
     * multiplificar por cinco su velocidad.
     * @param {Integer} i <1> para optimizar, <0> para restablecer.
     */
    SpeedupIO(i)
    {
        this._workbook.Application.EnableEvents := !i
        this._workbook.Application.ScreenUpdating := !i
        this._workbook.Application.Calculation := (i = 0) ? -4105 : -4135
    }

    /**
     * @public
     * Señala con un color amarillo la fila indicada y restablece la anterior.
     * 
     * Si se indica una fila ya señalada, restablece su color.
     * 
     * @param {Integer} row Fila a señalizar o restablecer.
     */
    HighlightRow(row)
    {
        if (row < 1)
            throw ValueError("La fila {" row "} no es válida.")

        BGR_HIGHTLIGHT := 0x00FFFF  ; Código hex BGR amarillo
        BGR_NONE := -4142
        reset := (row = this.__lastHighlightedRow)

        ;// Restablecer la fila anterior
        if (!reset && this.__lastHighlightedRow != 0) {
            this._targetSheet.Rows[this.__lastHighlightedRow].Interior.Color := BGR_NONE
        }
        this._targetSheet.Rows[row].Interior.Color := reset ? BGR_NONE : BGR_HIGHTLIGHT
        this.__lastHighlightedRow := reset ? 0 : row
    }

    /**
     * @public 
     * Valida las cabeceras de la tabla objetivo conforme a la colección facilitada.
     * @param {Array<String>} expectedHeaders Colección de los nombres de las cabeceras esperadas.
     * @param {VarRef<Array<String>>} missingHeaders (OUT Opcional) Collección de los nombres de las cabeceras faltantes.
     * @returns {Boolean} Verdadero si la tabla contiene todas las cabeceras, Falso en su defecto.
     */
    ValidateHeaders(expectedHeaders, &missingHeaders := unset)
    {
        missingHeaders := []
        
        ;// Validación
        if (Type(expectedHeaders) != "Array")
            expectedHeaders := [expectedHeaders]
        if (expectedHeaders.Length = 0)
            return true
        ; if (expectedHeaders.Length > 0 && this.IsUsedRangeEmpty())
        ;     throw UnsetError('El rango objetivo "' this._GetTargetRange().Address '" está vacío.')
        if (Type(expectedHeaders[1]) != "String")
            throw TypeError("Se esperaba una colección de String, pero se ha recibido: " Type(expectedHeaders[1]))
        
        ;// Normalizar
        for header in expectedHeaders
            expectedHeaders[A_Index] := WorkbookWrapper.__NormalizeHeader(header)
        this._NormalizeTableHeaders()

        ;// Obtener cabeceras de la tabla
        headerRow := this._GetRowSafeArray(1)
        headersMap := Map() ;// Se utiliza Map para facilitar la búsqueda O(n + m)
        Loop headerRow.MaxIndex(2) {
            header := headerRow[1, A_Index]
            headersMap[header] := true
        }

        ;// Validación de las cabeceras
        for (header in expectedHeaders) {
            if (!headersMap.Has(header))
                missingHeaders.Push(header)
        }
        return missingHeaders.Length = 0
    }

    /**
     * @private 
     * **EN DESUSO:**
     * Valida que el libro de trabajo objetivo se encuentre abierto y accesible.
     * @param {Microsoft.Office.Interop.Excel.Workbook} workbook (Opcional) Libro de trabajo a validar.
     * Por defecto es el libro de trabajo objetivo.
     * @throws {TargetError} Si el libro se encuentra cerrado.
     * @throws {Error} Si Excel rechaza la conexión a su interfaz.
     */
    __ThrowIfWorkbookIsInvalid(workbook := this._workbook)
    {
        if (workbook = 0)
            throw ValueError("No se ha establecido un libro de trabajo objetivo.")
        if !(workbook is ComObject) ; No es necesario comprobar el tipo aquí
            throw TypeError('Se esperaba el tipo ComObject, pero se ha recibido: ' Type(workbook))

        try {
            workbook.Name
        } catch Error as err {
            ;// El objeto ha sido desconectado de sus clientes.
            if (InStr(err.Message, "0x80010108"))
                throw TargetError("(0x80010108) El libro de trabajo solicitado ha sido cerrado.", -1, err)
            ;// Excel ha rechazado la conexión a sus objetos.
            if (InStr(err.Message, "0x80010001"))
                throw Error("(0x80010001) Excel ha rechazado la conexión a su interfaz.", -1, err)
            throw err
        }
    }

    /**
     * @protected
     * Bloquea el libro de trabajo objetivo impidiendo su cierre y la manipulación del número de hojas.
     * @param {Boolean} lock Si bloquear o desbloquear.
     */
    _LockWorkbook(lock)
    {
        if (this.IsWorkbookLocked() != lock) {
            this._workbook.Protect(Password := "") ; Esto es un interruptor
        }
    }

    /**
     * @protected 
     * Bloquea o desbloquea la edición y selección de las celdas de la
     * hoja de cálculo objetivo.
     * 
     * @param {Boolean} lock Veradero para bloquear, Falso para desbloquear.
     */
    _LockSheet(lock)
    {
        sheet := this._targetSheet
        if (lock && !this.IsSheetLocked()) {
            sheet.Protect(Password := "", DrawingObjects := true, Contents := true, Scenarios := true, UserInterfaceOnly := true)
            sheet.EnableSelection := -4142
        }
        else if (!lock && this.IsSheetLocked()) {
            sheet.Unprotect()
            sheet.EnableSelection := 0
        }
    }

    /**
     * @protected 
     * Elimina las filas vacías del rango objetivo (contempla fórmulas).
     * 
     * Para las tablas, hay que rellenar al menos una celda para
     * Auto-Expandir la tabla con nuevos datos. Por ello, si la tabla
     * carece de contenido, rellenará la primera fila de valores con
     * "null".
     */
    _DeleteEmptyRows()
    {
        range := this._GetTargetRange()
        rows := range.Rows
        rowCount := rows.Count
        maxRowCount := rowCount
        Loop rowCount {
            index := maxRowCount - A_Index + 1 ; Invertido
            row := rows[index]
            isThereAnyValue := row.Find("*",, xlFormulas:=-4123)
            if (!isThereAnyValue) {
                ;// No se puede borrar la última fila de valores de una tabla
                if (index = 2 && rowCount = 2) {
                    ;// 
                    ;// 
                    row.Value2 := "null"
                    break
                }
                try {
                    row.EntireRow.Delete()
                    rowCount--
                }
            }
        }
    }

    /**
     * @protected 
     * Obtiene el rango de la primera tabla si existiera,
     * o el rango **continuo** utilizado.
     * 
     * @returns {Microsoft.Office.Interop.Excel.Range} Rango objetivo.
     * @throws {ValueError} Si existe más de tabla definida en la hoja de cálculo objetivo.
     */
    _GetTargetRange()
    {
        sheet := this._targetSheet

        if (sheet.ListObjects.Count > 1)
            throw ValueError("Existe más de tabla definida en la hoja de cálculo objetivo. Utiliza otra hoja o borra las tablas sobrantes.")

        return (sheet.ListObjects.Count >= 1) ? sheet.ListObjects[1].Range : sheet.UsedRange.CurrentRegion
        
        ; if (sheet.ListObjects.Count >= 1) {
        ;     return sheet.ListObjects[1].Range ; Table #1
        ; }
        ; else {
        ;     if (sheet.Cells(1,1).Value2) { ; A1 value
        ;         return sheet.UsedRange.CurrentRegion
        ;     }
        ;     return sheet.Range("A1")
        ; }
    }

    /**
     * @protected
     * Envuelve el rango objetivo en una tabla si no existe ninguna.
     * @param {Integer} hasHeaders La cabecera tiene encabezados.
     * Debe ser un valor XlYesNoGuess (1,2,0). Por defecto es 0 (guess).
     */
    _WrapTargetRangeInTable(hasHeaders := 0)
    {
        if (Type(hasHeaders) != "Integer")
            throw ValueError("Se esperaba un Integer, pero se ha recibido: " Type(hasHeaders))
        if (hasHeaders < 0 || hasHeaders > 2)
            throw ValueError("Se esperaba un valor entre 0 y 2, pero se ha recibido: " hasHeaders)
        
        targetRange := this._GetTargetRange()
        if (targetRange.Value2 = "")
            return
        if (this._targetSheet.ListObjects.Count > 0)
            return
        
        ;// Se requiere desbloquear la hoja para crear una tabla
        sheetWasLocked := this.IsSheetLocked()
        this._LockSheet(false)
        this._targetSheet.ListObjects.Add(
            XlListObjectSourceType := 1, ; xlSrcRange
            targetRange,
            ,
            XlYesNoGuess := hasHeaders
        )
        this._LockSheet(sheetWasLocked)
    }

    /**
     * @protected 
     * Obtiene el contenido de una fila del rango objetivo como un SafeArray COM.
     *
     * Excel Interop presenta un comportamiento inconsistente al leer valores:
     * 
     *  - Si la fila contiene una única columna o está vacía, devuelve un valor.
     *  - Si contiene varias columnas, devuelve un SafeArray.
     *
     * Esta función normaliza dicho comportamiento garantizando que el valor
     * devuelto sea siempre un ComObjArray (SafeArray) con índice base 1,
     * incluso cuando la fila esté vacía o contenga una sola celda.
     *
     * @param {Integer} row Índice (1-based) de la fila dentro del rango objetivo.
     * @returns {ComObjArray} SafeArray bidimensional (1×N) que representa la fila solicitada.
     * @throws {ValueError} Si el índice de fila está fuera del rango utilizado.
     */
    _GetRowSafeArray(row)
    {
        range := this._GetTargetRange()
        if (Type(row) != "Integer")
            throw TypeError("Se esperaba un Integer pero se ha recibido: " Type(row))
        if (row < 1 || row > range.Rows.Count)
            throw ValueError('La fila {' row '} está fuera del rango utilizado.')
        
        val := range.Rows[row].Value2
        if (val = "") {
            safeArr := WorkbookWrapper._CreateInteropArray(1,0)
        }
        else if !(val is ComObjArray) {
            safeArr := WorkbookWrapper._CreateInteropArray(1,1)
            safeArr[1,1] := val
        } else {
            safeArr := val
        }
        
        return safeArr
    }

    /**
     * @protected
     * Normaliza las cabeceras de la tabla objetivo conforme a __NormalizeHeader.
     * @see WorkbookWrapper.__NormalizeHeader
     */
    _NormalizeTableHeaders()
    {
        headerRow := this._GetRowSafeArray(1)

        ;// Comprobar si es necesario normalizar
        Loop headerRow.MaxIndex(2) {
            header := headerRow[1, A_Index]
            normalizedHeader := WorkbookWrapper.__NormalizeHeader(header)
            if (header !== normalizedHeader) {
                this._GetTargetRange().Cells(1, A_Index).Value2 := normalizedHeader
            }
        }
    }

    /**
     * @protected 
     * Normaliza los nombres de las propiedades del objeto indicado conforme a __NormalizeHeader.
     * 
     * Se utiliza para mantener la coherencia entre las tablas y los objetos AHK.
     * 
     * @param {Object} obj Objeto a normalizar.
     * @returns {Object} Objeto normalizado.
     * @see WorkbookWrapper.__NormalizeHeader
     */
    static _NormalizeObjProps(obj)
    {
        if (!IsObject)
            throw TypeError('Se esperaba un Object, pero se ha recibido, pero se ha recibido: ' Type(obj))
        
        for prop in obj.OwnProps() {
            if (prop == this.__NormalizeHeader(prop))
                continue
            
            normalizedProp := this.__NormalizeHeader(prop)
            value := obj.%prop%
            obj.DeleteProp(prop)
            obj.%normalizedProp% := value
        }
        
        return obj
    }
    
    /**
     * @protected 
     * Crea un SafeArray bidimensional (VT_VARIANT) con índices de base 1 
     * como los que devuelve Interop, que se supone utiliza una versión 
     * descontinuada del SafeArray.
     * @param {Integer} size1 Tamaño para la primera dimensión.
     * @param {Integer} size2 Tamaño para la segunda dimensión.
     * @returns {ComObjArray} SafeArray bidimensional (VT_VARIANT) con índices de base 1.
     */
    static _CreateInteropArray(size1, size2)
    {
        if (size1 < 0 || size2 < 0)
            throw ValueError("Las capacidades del SafeArray no pueden ser negativas.")

        bounds := Buffer(16, 0)             ; 2 dims -> 2 * sizeof(SAFEARRAYBOUND) = 16 bytes
        NumPut("UInt", size1, bounds, 0)     ; cElements for dimension 1
        NumPut("UInt", 1, bounds, 4)        ; LBound for dimension 1
        NumPut("UInt", size2, bounds, 8)     ; cElements for dimension 2
        NumPut("UInt", 1, bounds, 12)       ; LBound for dimension 2

        safeArr := DllCall(
            "oleaut32.dll\SafeArrayCreate"
            , "ushort", VT_VARIANT:=12
            , "uint", 2
            , "ptr", bounds.Ptr
            , "ptr"                         ; return type
        )

        ;// Envolverlo en un VARIANT tipo VT_ARRAY|VT_VARIANT para pasar a COM
        return ComValue(VT_ARRAY:=0x2000 | VT_VARIANT:=12, safeArr)
    }

    /**
     * @private
     * Una cabecera normalizada debe estar en mayúsculas, sin tildes y con barras bajas
     * en vez de espacios.
     * @param {String} header Título de la cabecera a normalizar.
     * @returns {String} Cabecera normalizada.
     */
    static __NormalizeHeader(header)
    {
        if (Type(header) != "String")
            throw TypeError("Se esperaba un String, pero se ha recibido: " Type(header))
        
        return StrUpper(StrReplace(Utils.RemoveDiacritics(Trim(header)), ' ', '_'))
    }

    /**
     * @private
     * Ejecuta una función controlando la interacción del usuario con Excel
     * para evitar fallos de automatización durante operaciones críticas.
     *
     * Si Excel rechaza la llamada COM por estar ocupado (por ejemplo, debido
     * a edición activa de celdas o diálogos modales) tras 30 reintentos en 30 segundos, 
     * esta función envía {ESCAPE} para cancelar la edición en curso y 
     * reintenta la operación una única vez más.
     * 
     * Notifica al usuario mediante un TrayTip al escapar la edición.
     *
     * Solo intercepta errores COM conocidos relacionados con Excel ocupado
     * (HRESULT 0x80010001, 0x800AC472). Cualquier otro error se relanza.
     *
     * @param {Func} fun Función a ejecutar. Debe aceptar `this` como primer parámetro.
     * @param {Any} params Parámetros opcionales que se pasarán a la función.
     * @returns {Any} Valor devuelto por la función ejecutada.
     * @throws {Error} Relanza la excepción si el error no es recuperable 
     * o si los reintentos fallan.
     */
    __InvokeExcelSafelyDelayed(fun, params*)
    {
        Loop (retries := 30) + 1 {
            try {
                return fun(this, params*)
            }
            catch Error as err {
                ;// Aplicable solo cuando Excel rechace la conexión a su interfaz porque está ocupado
                if (!InStr(err.Message, "0x80010001") && !InStr(err.Message, "0x800AC472")) {
                    throw err
                }

                ;// Reintentar
                if (A_Index < retries) {
                    Sleep 1000
                    continue
                }
                ;// Permitir un solo reintento tras escapar la edición
                if (A_Index > retries) {
                    throw err
                }

                Utils.EscapeExcelEditMode()
                TrayTip("Ups, AutoHotKey ha tenido que cancelar su edición del libro porque está trabajando con Excel.", "AutoHotKey", 2)
            }
        }
    }
}
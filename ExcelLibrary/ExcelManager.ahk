#Requires AutoHotkey v2.0
#Include "ExcelEventController.ahk"
#Include "ExcelBridge\WorkbookWrapper.ahk"
#Include "ExcelBridge\ReadWorkbookAdapter.ahk"
#Include "ExcelBridge\WriteWorkbookAdapter.ahk"
#Include "..\Util\Utils.ahk"

/**
 * @internal
 * ROADMAP
 * - Logging system
 */

/************************************************************************
 * @brief
 * Administrador para el COM de Microsoft Excel.
 * 
 * - Conceptualizado para **NO TOCAR** los datos preexistentes en el libro de lectura 
 * y preservar así su integridad (excepto las cabeceras que se normalizan).
 * - Es capaz de escapar la edición del usuario si impide el acceso al COM.
 * - Es la ostia de rápido.
 * 
 * @note 
 * - Se impone el uso de tablas para definir los rangos utilizados.
 * - Relacionado con el anterior, las hojas de cálculo a trabajar 
 * solo pueden contener una tabla como máximo.
 * - La validación de tipos perjudica el rendimiento.
 * 
 * @author bitasuperactive
 * @date 25/12/2025
 * @version 0.9.1-Beta
 * @warning Dependencias:
 * - ExcelEventController.ahk
 * - WorkbookWrapper.ahk
 * - ReadWorkbookAdapter.ahk
 * - WriteWorkbookAdapter.ahk
 * - Utils.ahk
 * @see https://github.com/bitasuperactive/ahk2-excel-library/blob/master/ExcelLibrary/ExcelManager.ahk
 * @internal Documentación web: https://bitasuperactive.github.io/ahk2-excel-library
 ***********************************************************************/
class ExcelManager
{
    /** @private */
    _excelCOM := unset ;
    /** @private */
    _allowReadAndWrite := unset ;
    /** @private */
    _readWorkbookAdapter := 0 ;
    /** @private */
    _writeWorkbookAdapter := 0 ;

    /**
     * @public
     * @returns {ReadWorkbookAdapter} Adaptador para el libro de lectura conectado.
     */
    ReadWorkbookAdapter => this._readWorkbookAdapter ;

    /**
     * @public
     * @returns {WriteWorkbookAdapter} Adaptador para el libro de escritura conectado.
     */
    WriteWorkbookAdapter => this._writeWorkbookAdapter ;

    /**
     * @public
     * Crea un administrador para el COM de Microsoft Excel.
     * - Ejecuta Excel automáticamente.
     * - Establece la conexión con el COM de Excel.
     * 
     * @warning Al tratar con datos sensibles, es recomendable no permitir 
     * leer y escribir en la misma hoja de cálculo para garantizar la integridad de los datos.
     * 
     * @param {Boolean} allowReadAndWrite Si permitir que una misma hoja de cálculo se utilice tanto 
     * para lectura como para escritura.
     * @throws {TargetError} Si no ha sido posible iniciar Microsoft Excel automáticamente.
     * @throws {Error} (0x80004002) Si Microsoft Excel ha rechazado la conexión a su interfaz.
     */
    __New(allowReadAndWrite := false)
    {
        if (Type(allowReadAndWrite) != "Integer")
            throw TypeError("Se esperaba un Boolean, pero se ha recibido: " Type(allowReadAndWrite))

        Utils.ProxyObjFuncs(this, this.__InvokeExcelSafely)
        
        this._excelCOM := this.__InvokeExcelSafely(ExcelManager._GetExcelCOM)
        this._allowReadAndWrite := allowReadAndWrite

        ;// Establecer eventos
        ExcelEventController.SetupOnApplicationStateChangedEvent()
        ExcelEventController.OnEvent(
            ExcelEventController.ApplicationEventEnum.APPLICATON_TERMINATED, 
            (*) => this.Dispose()
        )
        ExcelEventController.OnEvent(
            ExcelEventController.WorkbookEventEnum.TARGET_WORKBOOK_BEFORE_CLOSE,
            (caller, cancel, workbook) => ( 
                this.__OnTargetWorkbookBeforeClose(caller, cancel, workbook)
            )
        )
        OnExit((*) => this.Dispose())   ; Se debe implementar así para no perder la referencia de la instancia
    }
    
    /**
     * @public 
     * Obtiene los nombres sin extensión de todos los libros de trabajo 
     * abiertos y compatibles (con extensión ".xlsx").
     * @returns {Array<String>}
     */
    GetAllOpenWorkbooksNames()
    {
        workbooks := this._excelCOM.Workbooks
        names := []
        for wb in workbooks {
            split := Utils.StrSplitExtension(wb.Name)
            name := split[1]
            ext := split[2]
            if (ext = "" || ext = ".xlsx")
                names.Push(name)
        }
        return names
    }

    /**
     * @public 
     * Conecta un libro de trabajo abierto mediante su nombre para el tipo de conexión
     * especificado.
     * - Toma la hoja de cálculo activa en el momento de la conexión como objetivo del uso.
     * - El libro es bloqueado para evitar su cierre y la manipulación del número de hojas.
     * 
     * @param {ExcelManager.ConnectionTypeEnum} connType Tipo de uso que se le dará al libro.
     * @param {String} name Nombre del libro de trabajo objetivo.
     * @param {Boolean} lockSheet (Opcional) Si bloquear la hoja de cálculo objetivo impidiendo 
     * la modificación y la selección de sus celdas. Por defecto es falso.
     * @throws {ValueError} Si no existe ningún libro de trabajo abierto y compatible 
     * con el nombre solicitado.
     * @throws {Error} Si no se ha permitido leer y escribir en la misma hoja de cálculo
     * pero se intenta establecer esa conexión.
     */
    ConnectWorkbookByName(connType, name, lockSheet := false)
    {
        ext := Utils.StrSplitExtension(name)[2]
        if (ext != "" && ext != ".xlsx")
            throw ValueError('El libro de trabajo "' name '" no es compatible. Solo se permiten extensiones ".xlsx".')

        workbooks := []
        workbooks := this._excelCOM.Workbooks
        for wb in workbooks {
            wbName := wb.Name
            ext := Utils.StrSplitExtension(wbName)[2]
            if ((ext = "" || ext = ".xlsx") && (wbName = name || wbName = name ".xlsx")) {
                this._ConnectWorkbook(connType, wb, lockSheet)
                return
            }
        }
        throw ValueError('No existe ningún libro de trabajo abierto y compatible llamado "' name '".')
    }

    /**
     * @public
     * Permite cerrar los libros de trabajo conectados. Por defecto es Falso.
     * @param {Boolean} allow Verdadero para permitir el cierre, Falso para impedirlo.
     */
    AllowWorkbookClosure(allow)
    {
        if (Type(allow) != "Integer")
            throw TypeError("Se esperaba un Boolean, pero se ha recibido: " Type(allow))
        
        ExcelEventController.WorkbookEventHandler.AllowClosure := allow
    }

    /**
     * @public
     * Desbloquea el libro de trabajo y hoja de cálculo objetivos para el tipo 
     * de conexión especificado.
     * 
     * Permisos administrados:
     * - Editar la hoja de cálculo.
     * - Cerrar el libro de trabajo.
     * - Manipular el número de hojas.
     * - Mostrar las alertas de Excel.
     * 
     * @warning No sé por qué querrías hacer esto, pero no te lo recomiendo.
     * 
     * @param {ExcelManager.ConnectionTypeEnum} connType Tipo del libro de trabajo
     * objetivo.
     * @param {Boolean} unlock Verdadero para desbloquear, Falso para bloquear.
     */
    UnlockWorkbook(connType, unlock)
    {
        if (!Utils.ValidateInheritanceClass(connType, ExcelManager.ConnectionTypeEnum))
            throw TypeError('Se esperaba el tipo "' ExcelManager.ConnectionTypeEnum.Prototype.__Class '", pero se ha recibido: ' Type(connType))
        if (Type(unlock) != "Integer")
            throw TypeError("Se esperaba un Boolean, pero se ha recibido: " Type(unlock))

        switch(connType) {
            case ExcelManager.ConnectionTypeEnum.READ:
                adapter := this._readWorkbookAdapter
            case ExcelManager.ConnectionTypeEnum.WRITE:
                adapter := this._writeWorkbookAdapter
            default:
                throw ValueError("El tipo de libro de trabajo solicitado no está definido.")
        }
        
        adapter._LockSheet(!unlock)
        this._LockWorkbook(adapter, !unlock)
    }

    /**
     * @public 
     * Desconecta el libro de trabajo conectado para el tipo de uso especificado, desbloqueándolo.
     * @param {ExcelManager.ConnectionTypeEnum} connType Tipo del libro de trabajo objetivo.
     */
    DisconnectWorkbook(connType)
    {
        if (!Utils.ValidateInheritanceClass(connType, ExcelManager.ConnectionTypeEnum))
            throw TypeError('Se esperaba el tipo "' ExcelManager.ConnectionTypeEnum.Prototype.__Class '", pero se ha recibido: ' Type(connType))
        
        switch(connType) {
            case ExcelManager.ConnectionTypeEnum.READ:
                this._DisconnectWorkbook(this._readWorkbookAdapter)
            case ExcelManager.ConnectionTypeEnum.WRITE:
                this._DisconnectWorkbook(this._writeWorkbookAdapter)
            default:
                throw ValueError("El tipo de libro de trabajo solicitado no está definido.")
        }
    }
    
    /**
     * @public 
     * Desecha la instancia desconectando los libros conectados 
     * y limpiando los manejadores de eventos configurados.
     */
    Dispose()
    {
        try this._DisconnectWorkbook(this._readWorkbookAdapter)
        try this._DisconnectWorkbook(this._writeWorkbookAdapter)
        try ComObjConnect(this._excelCOM)
        try this._excelCOM := unset
        ExcelEventController.DisposeEvents()
    }

    /**
     * @private 
     * Conecta el libro de trabajo abierto solicitado.
     * - Toma la hoja de cálculo activa en el momento de la conexión.
     * - El libro es bloqueado para evitar su cierre y la manipulación del número de hojas.
     * 
     * @param {ExcelManager.ConnectionTypeEnum} connType Tipo de uso que se le dará al libro.
     * @param {Microsoft.Office.Interop.Excel.Workbook} workbook (Opcional) Libro de trabajo 
     * objetivo. Por defecto es el libro activo.
     * @param {Boolean} lockSheet (Opcional) Si bloquear la hoja de cálculo objetivo impidiendo 
     * la modificación y la selección de sus celdas. Por defecto es falso.
     * @throws {Error} Si no se ha permitido leer y escribir en la misma hoja de cálculo
     * pero se intenta establecer esa conexión.
     */
    _ConnectWorkbook(connType, workbook?, lockSheet := false)
    {
        if (!Utils.ValidateInheritanceClass(connType, ExcelManager.ConnectionTypeEnum))
            throw TypeError('Se esperaba el tipo "' ExcelManager.ConnectionTypeEnum.Prototype.__Class '", pero se ha recibido: ' Type(connType))
        if (IsSet(workbook) && (!(workbook is ComObject) || Type(workbook) != "Workbook"))
            throw TypeError('Se esperaba el tipo "ComObject.Workbook", pero se ha recibido: ' Type(workbook))
        if (Type(lockSheet) != "Integer")
            throw TypeError("Se esperaba un Boolean, pero se ha recibido: " Type(lockSheet))

        workbook := IsSet(workbook) ? workbook : this._excelCOM.ActiveWorkbook
        
        switch(connType) {
            case ExcelManager.ConnectionTypeEnum.READ:
            {
                this._DisconnectWorkbook(this._readWorkbookAdapter)
                adapterClass := ReadWorkbookAdapter.Prototype.__Class
            }
            case ExcelManager.ConnectionTypeEnum.WRITE:
            {
                this._DisconnectWorkbook(this._writeWorkbookAdapter)
                adapterClass := WriteWorkbookAdapter.Prototype.__Class
            }
            default:
            {
                throw ValueError("El tipo de libro de trabajo solicitado no está definido.")
            }
        }

        adapter := %adapterClass%(workbook)
        this._SetWorkbookAdapter(adapter)
        this._LockWorkbook(adapter, true) ; Obligatorio
        if (!this._SameAdapterForReadAndWrite()) { ; Evita duplicar eventos
            ComObjConnect(adapter._workbook, ExcelEventController.WorkbookEventHandler)
            __ConnectSheet(lockSheet)
        }

        
        /**
         * Bloquea la hoja de cálculo objetivo y conecta sus eventos.
         */
        __ConnectSheet(lock)
        {
            ComObjConnect(adapter._targetSheet, ExcelEventController.WorksheetEventHandler)
            adapter._LockSheet(lock)
        }
    }

    /**
     * @private 
     * Desconecta un libro de trabajo conectado, desbloqueándolo.
     * @param {ReadWorkbookAdapter | WriteWorkbookAdapter} adapter Adaptador del libro de trabajo objetivo.
     */
    _DisconnectWorkbook(adapter)
    {
        if (adapter = 0) 
            return
        if (!Utils.ValidateInheritance(adapter, WorkbookWrapper))
            throw TypeError('Se esperaba una clase heredada de "' WorkbookWrapper.Prototype.__Class '", pero se ha recibido: ' Type(adapter))
        
        try {
            ;// Comprobar si Excel es accesible
            this._excelCOM.Application.Ready
        }
        catch Error as err {
            ;// Si Excel está ocupado, escapar la edición directamente
            if (InStr(err.Message, "0x80004002") || InStr(err.Message, "0x80010001") || InStr(err.Message, "0x800AC472"))
                Utils.EscapeExcelEditMode()
        }

        __DisconnectSheet()
        this._LockWorkbook(adapter, false)
        ComObjConnect(adapter._workbook)
        __NullifyAdapter(adapter)

        
        /**
         * Desbloquea la hoja de cálculo objetivo y desconecta sus eventos.
         */
        __DisconnectSheet()
        {
            ComObjConnect(adapter._targetSheet)
            adapter._LockSheet(0)
        }

        /**
         * Nulifica la propiedad de instancia que almacena el adaptador especificado.
         * @param {ReadWorkbookAdapter | WriteWorkbookAdapter} adapter Adaptador del libro de trabajo objetivo.
         */
        __NullifyAdapter(adapter)
        {
            switch(Type(adapter)) {
                case ReadWorkbookAdapter.Prototype.__Class:
                    this._readWorkbookAdapter := 0
                case WriteWorkbookAdapter.Prototype.__Class:
                    this._writeWorkbookAdapter := 0
                default:
                    throw ValueError("El tipo de libro de trabajo solicitado no está definido.")
            }
        }
    }
    
    /**
     * @private
     * Ejecuta y/u obtiene el COM del proceso activo de Microsoft Excel
     * y conecta el manejador de eventos para su aplicación.
     * 
     * @note Si se requiere ejecutar Microsoft Excel, esperará hasta `60` segundos tras lanzarlo
     * para que acepte la conexión y permita el acceso a su interfaz COM.
     * 
     * @returns {ComObject} Common Object Model para la instancia activa de Microsoft Excel.
     * @throws {TargetError} Si no ha sido posible iniciar Microsoft Excel automáticamente.
     * @throws {Error} (0x80004002) Si Microsoft Excel ha rechazado la conexión a su interfaz.
     */
    static _GetExcelCOM()
    {
        if (!ProcessExist("EXCEL.EXE")) {
            try Run("EXCEL")
            catch Error as err
                throw TargetError("No ha sido posible iniciar Microsoft Excel automáticamente.", -1, err)

            ;// Esperar hasta 60 segundos para completar la conexión
            Loop 60 {
                try {
                    if (WinGetCount("ahk_class XLMAIN") > 0) {  ; Ventana activa de Excel
                        excelCOM := ComObjActive("Excel.Application")

                        ;// Crear un libro en blanco si el usuario no lo ha creado ya
                        if (excelCOM.Workbooks.Count = 0) {
                            excelCOM.Workbooks.Add()
                        }

                        ComObjConnect(excelCOM, ExcelEventController.ApplicationEventHandler)
                        return excelCOM
                    }
                }
                Sleep 1000
            }
        }

        try {
            excelCOM := ComObjActive("Excel.Application")
            ComObjConnect(excelCOM, ExcelEventController.ApplicationEventHandler)

            ;// Crear un libro en blanco si el proceso quedó activo sin ventanas
            if (WinGetCount("ahk_class XLMAIN") = 0) {
                excelCOM.Workbooks.Add()
            }
            
            return excelCOM
        } 
        catch Error as err {
            if (InStr(err.Message, "0x800401E3")) ; Excel no está iniciado
                throw TargetError("(0x800401E3) No ha sido posible iniciar Microsoft Excel automáticamente.", -1, err)
            if (InStr(err.Message, "0x80004002")) ; Excel ha rechazado la conexión
                throw Error("(0x80004002) Excel ha rechazado la conexión a su interfaz.", -1, err)
            throw err
        }
    }

    /**
     * @private 
     * Establece el adaptador para el libro de trabajo objetivo según el tipo que le corresponda,
     * y comprueba si se está infringiendo la regla de lectura y escritura en la misma hoja de cálculo.
     * @param {ReadWorkbookAdapter | WriteWorkbookAdapter} adapter Adaptador para el libro de trabajo a establecer.
     * @throws {Error} Si no se ha permitido leer y escribir en la misma hoja de cálculo
     * pero se intenta establecer para ambos propósitos.
     */
    _SetWorkbookAdapter(adapter)
    {
        if (!Utils.ValidateInheritance(adapter, WorkbookWrapper))
            throw TypeError('Se esperaba un tipo heredado de "' WorkbookWrapper.Prototype.__Class '", pero se ha recibido: ' Type(adapter))

        switch(Type(adapter)) {
            case ReadWorkbookAdapter.Prototype.__Class:
            {
                if (!this._allowReadAndWrite && this._SameAdapterForReadAndWrite(adapter))
                    throw Error("No se ha permitido leer y escribir en la misma hoja de cálculo.")
                this._readWorkbookAdapter := adapter
            }
            case WriteWorkbookAdapter.Prototype.__Class:
            {
                if (!this._allowReadAndWrite && this._SameAdapterForReadAndWrite(,adapter))
                    throw Error("No se ha permitido leer y escribir en la misma hoja de cálculo.")
                this._writeWorkbookAdapter := adapter
            }
            default:
            {
                throw ValueError("El tipo de libro de trabajo solicitado no está definido.")
            }
        }
    }


    /**
     * Comprueba si los adaptadores de lectura y escritura son el mismo.
     * @param {ReadWorkbookAdapter} readAdapter (Opcional) Libro de lectura.
     * @param {WriteWorkbookAdapter} writeAdapter (Opcional) Libro de escritura.
     * @returns {Boolean} Verdadero si ambos adaptadores son el mismo, Falso en su defecto.
     */
    _SameAdapterForReadAndWrite(readAdapter := this._readWorkbookAdapter, writeAdapter := this._writeWorkbookAdapter)
    {
        if (readAdapter && !(readAdapter is ReadWorkbookAdapter))
            throw ValueError("Se esperaba el tipo '" ReadWorkbookAdapter.Prototype.__Class "', pero se ha recibido: " Type(readAdapter))
        if (writeAdapter && !(writeAdapter is WriteWorkbookAdapter))
            throw ValueError("Se esperaba el tipo '" WriteWorkbookAdapter.Prototype.__Class "', pero se ha recibido: " Type(writeAdapter))
        
        return (readAdapter
            && writeAdapter
            && readAdapter.Name = writeAdapter.Name
            && readAdapter.TargetSheetName = writeAdapter.TargetSheetName)
    }

    /**
     * @private
     * Bloquea el libro de trabajo especificado impidiendo su cierre y la manipulación del número de hojas.
     * También desactiva las alertas de Excel.
     * 
     * @note Es obligatorio bloquear el libro objetivo para evitar que el usuario lo cierre y manipule el número de hojas.
     * 
     * @param {WorkbookWrapper} adapter Adaptador objetivo.
     * @param {Boolean} lock Si bloquear o desbloquear.
     * @throws {Error} Si se intenta desbloquear el libro mientras la hoja de cálculo objetivo está bloqueada.
     */
    _LockWorkbook(adapter, lock)
    {
        if (!Utils.ValidateInheritance(adapter, WorkbookWrapper))
            throw TypeError('Se esperaba un tipo heredado de "' WorkbookWrapper.Prototype.__Class '", pero se ha recibido: ' Type(adapter))
        
        ;// Si la hoja está bloqueada, no se debe desbloquear el libro
        if (!lock && adapter.IsSheetLocked())
            throw Error("No es posible desbloquear el libro de trabajo mientras la hoja de cálculo objetivo esté bloqueada.")
        
        this._excelCOM.DisplayAlerts := !lock
        this.AllowWorkbookClosure(!lock)
        adapter._LockWorkbook(lock)
    }

    /**
     * @private 
     * Llamada ejecutada antes del cierre de alguno de los libros de trabajos conectados.
     * Desconecta el libro antes de completar el cierre.
     * 
     * @warning Si el usuario cancela el guardado, el libro queda desconectado.
     * 
     * @param {Object} caller Referencia al objeto llamante.
     * @param {Boolean} cancel Si se ha cancelado el cierre solicitado.
     * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo a cerrar.
     */
    __OnTargetWorkbookBeforeClose(caller, cancel, workbook)
    {
        if (this._readWorkbookAdapter && this._readWorkbookAdapter.IsTargetWorkbook(workbook))
            this._DisconnectWorkbook(this._readWorkbookAdapter)
        if (this._writeWorkbookAdapter && this._writeWorkbookAdapter.IsTargetWorkbook(workbook))
            this._DisconnectWorkbook(this._writeWorkbookAdapter)
    }

    /**
     * @private
     * Ejecuta una función controlando la interacción del usuario con Excel
     * para evitar fallos de automatización durante sus operaciones con el COM.
     *
     * Si Excel rechaza la llamada COM por estar ocupado (por ejemplo, debido
     * a edición activa de celdas o diálogos modales), esta función envía {ESCAPE} 
     * para cancelar la edición en curso y reintenta la operación una única vez.
     *
     * Solo intercepta errores COM conocidos relacionados con Excel ocupado
     * (HRESULT 0x80010001, 0x800AC472). Cualquier otro error se relanza.
     * 
     * @note No notifica al usuario al escapar la edición, ya que se asume que ha sido iniciada
     * manualmente.
     *
     * @param {Func} fun Función a ejecutar. Debe aceptar `this` como primer parámetro.
     * @param {Any} params Parámetros opcionales que se pasarán a la función.
     * @returns {Any} Valor devuelto por la función ejecutada.
     * @throws {Error} Relanza la excepción si el error no es recuperable 
     * o si el reintento falla.
     */
    __InvokeExcelSafely(fun, params*)
    {
        Loop 2 {
            try {
                return fun(this, params*)
            }
            catch Error as err {
                ;// Aplicable solo cuando Excel rechace la conexión a su interfaz porque está ocupado
                if (!InStr(err.Message, "0x80004002") && !InStr(err.Message, "0x80010001") && !InStr(err.Message, "0x800AC472")) {
                    throw err
                }
                
                ;// Permitir un solo reintento
                if (A_Index > 1) {
                    throw err
                }
                
                Utils.EscapeExcelEditMode()
            }
        }
    }
    
    /**
     * @public
     * Tipos de conexión o de uso admitidos para los libros de trabajo.
     */
    class ConnectionTypeEnum
    {
        /**
         * @public
         * Uso limitado a la lectura del libro de trabajo.
         */
        class READ {
        } ;
        
        /**
         * @public
         * Uso limitado a la escritura del libro de trabajo.
         */
        class WRITE {
        } ;
    }
}
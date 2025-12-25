#Requires AutoHotkey v2.0
#Include "..\Util\EventController.ahk"
#Include "..\Util\ProcessWMI.ahk"

/**
 * @internal
 * CAMBIOS PARA LA 0.9.1-Beta
 * - Refactorización del EventController para facilitar el acceso a la
 * definición de eventos. Corrigiendo la encapsulación.
 */

/************************************************************************
 * @brief 
 * Controlador para los eventos de Microsoft Excel.
 * 
 * Cada manejador de eventos se debe conectar con el ComObject correspondiente
 * de Microsoft Excel mediante ComObjConnect.
 * 
 * Los parámetros para los callbacks serán equivalentes a los EventHandlers de
 * Microsoft.Office.Interop.Excel, añadiendo el objeto llamante "this" como primer parámetro.
 * Y están documentados en los enumeradores de eventos.
 * 
 * @author bitasuperactive
 * @date 17/12/2025
 * @version 0.9.0-Beta
 * @see null
 * @requires EventController.ahk
 * @requires ProcessWMI.ahk
 * @note Faltan numerosos eventos por implementar, pero los más importantes ya están cubiertos.
 ***********************************************************************/
class ExcelEventController
{
    /** @private */
    static _controller := EventController() ;
    /** @private */
    static _applicationWatcher := 0 ;
    /** @private */
    static _applicationHandler := this._ApplicationEventHandler() ;
    /** @private */
    static _workbookHandler := this._WorkbookEventHandler() ;
    /** @private */
    static _worksheetHandler := this._WorksheetEventHandler() ;

    /**
     * @public
     * Manejador de eventos para la aplicación de Microsoft Excel.
     * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.application?view=excel-pia#events
     */
    static ApplicationEventHandler => this._applicationHandler ;

    /**
     * @public
     * Manejador de eventos para un libro de trabajo especifico.
     * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.workbook?view=excel-pia#events
     */
    static WorkbookEventHandler => this._workbookHandler ;

    /**
     * @public
     * Manejador de eventos para una hoja de cálculo especifica.
     * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.worksheet?view=excel-pia#events
     */
    static WorksheetEventHandler => this._worksheetHandler ;

    static __New()
    {
        OnExit((*) => this._Dispose()) ;// Se debe implementar así para no perder la referencia al objeto
    }

    /**
     * @public
     * Delega la llamada a EventController::OnEvent.
     * @copydoc EventController::OnEvent
     */
    static OnEvent(name, callback) => this._controller.OnEvent(name, callback) ;
    
    /**
     * @public
     * Delega la llamada a EventController::Trigger.
     * @copydoc EventController::Trigger
     */
    static Trigger(name, params*) => this._controller.Trigger(name, params*) ;

    /**
     * @protected
     * Crea un evento único para detectar la ejecución y el cierre del proceso Microsoft Excel.
     * 
     * (!) Si se establece antes de iniciar el proceso, el evento se dispará dando a entender 
     * erróneamente que ha finalizado.
     */
    static SetupOnApplicationStateChangedEvent()
    {
        if (!this._applicationWatcher) {
            try {
                this._applicationWatcher := ProcessWMIWatcher("EXCEL.EXE", ProcessWMIEventHandler(this._OnApplicationStateChanged))
            }
            catch Error as err {
                MsgBox("No ha sido posible establecer el escuchador del proceso de Microsoft Excel.`n`nError: " err.Message, "ERROR", 16)
            }
        }
    }

    /**
     * @private
     * Ocurre al iniciar o finalizar el proceso de Microsoft Excel.
     * 
     * - El proceso de Excel no finalizará tras cerrar todas sus ventanas si su COM no es liberado.
     * 
     * @param {Boolean} executed Verdadero si Excel ha sido ejecutado, Falso si ha sido finalizado.
     */
    static _OnApplicationStateChanged(executed)
    {
        event := (executed) ? ExcelEventController.ApplicationEventEnum.APPLICATON_EXECUTED : ExcelEventController.ApplicationEventEnum.APPLICATON_TERMINATED
        ExcelEventController.Trigger(event)
        OutputDebug("EVENTO (" event "): El estado del proceso de EXCEL.EXE ha cambiado.`n")
    }

    /**
     * Elimina todos los eventos y manejadores registrados.
     */
    static DisposeEvents()
    {
        this._controller.Dispose()
    }

    /**
     * Desecha al escuchador del proceso de Microsoft Excel.
     */
    static _Dispose()
    {
        try this._applicationWatcher.Dispose()
        this.DisposeEvents()
    }

    /**
     * @public
     * Enumerador de los eventos admitidos por la aplicación de Microsoft Excel.
     */
    class ApplicationEventEnum
    {
        /**
         * Nombre para el evento que ocurre cuando el proceso de Microsoft Excel es ejecutado.
         * 
         * El callback no debe implementar parámetros.
         */
        static APPLICATON_EXECUTED := "APPLICATON_EXECUTED" ;
        
        /**
         * Nombre para el evento que ocurre cuando el proceso de Microsoft Excel es finalizado.
         * 
         * El callback no debe implementar parámetros.
         */
        static APPLICATON_TERMINATED := "APPLICATON_TERMINATED" ;
        
        /**
         * @public
         * Nombre para el evento que ocurre al crear un nuevo libro de trabajo.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        static ANY_WORKBOOK_NEW := "ANY_WORKBOOK_NEW" ;
        
        /**
         * @public
         * Nombre para el evento que ocurre al abrir cualquier libro guardado.
         * 
         * Si previamente existe un único libro en blanco, Excel lo cerrará tras unos milisegundos.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        static ANY_WORKBOOK_OPEN := "ANY_WORKBOOK_OPEN" ;
        
        /**
         * @public
         * Nombre para el evento que ocurre cuando el usuario intenta cerrar cualquier libro de trabajo.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {VarRef<Boolean>} cancel Falso por defecto. Si se establece en Verdadero, 
         * no se permitirá el cierre del libro de trabajo.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        static ANY_WORKBOOK_BEFORE_CLOSE := "ANY_WORKBOOK_BEFORE_CLOSE" ;

        
        /**
         * @public
         * Nombre para el evento que ocurre tras finalizar el guardado de cualquier libro de trabajo.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Boolean} success Si el guardado ha sido exitoso.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        static ANY_WORKBOOK_AFTER_SAVE := "ANY_WORKBOOK_AFTER_SAVE" ;
    } ;

    /**
     * @public
     * Enumerador de los eventos admitidos por los libros de trabajo.
     */
    class WorkbookEventEnum
    {
        /**
         * @public
         * Nombre para el evento que ocurre cuando se permite el cierre del libro de trabajo.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Boolean} cancel Falso por defecto, si se establece en Verdadero, 
         * no se permitirá el cierre del libro de trabajo.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         */
        static TARGET_WORKBOOK_BEFORE_CLOSE := "TARGET_WORKBOOK_BEFORE_CLOSE" ;

        
        /**
         * @public
         * Nombre para el evento que ocurre cuando se deniega el cierre del libro de trabajo.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Boolean} cancel Falso por defecto, si se establece en Verdadero, 
         * no se permitirá el cierre del libro de trabajo.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         */
        static TARGET_WORKBOOK_CLOSE_DENIED := "TARGET_WORKBOOK_CLOSE_DENIED" ;

        
        /**
         * @public
         * Nombre para el evento que ocurre tras finalizar el guardado del libro de trabajo.
         * 
         * El libro continúa siendo accesible tras guardarlo en formato xlsx.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Boolean} success Si el guardado ha sido exitoso.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         */
        static TARGET_WORKBOOK_AFTER_SAVE := "TARGET_WORKBOOK_AFTER_SAVE" ;

        /**
         * @public
         * Nombre para el evento que ocurre al activar una de las hojas de cálculo del libro de trabajo.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Microsoft.Office.Interop.Excel.Worksheet} sheet Puede ser un Worksheet o un Chart.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         */
        static SHEET_ACTIVATE := "SHEET_ACTIVATE" ;
    } ;

    /**
     * @public
     * Enumerador de los eventos admitidos por las hojas de cálculo.
     */
    class WorksheetEventEnum
    {        
        /**
         * @public
         * Nombre para el evento que ocurre cuando un rango de la hoja sufre un cambio.
         * 
         * El callback debe implementar los siguientes parámetros:
         * @param {Object} caller Referencia al objeto llamante.
         * @param {Microsoft.Office.Interop.Excel.Range} target Rango modificado.
         * @param {Microsoft.Office.Interop.Excel.Worksheet} worksheet Hoja de cálculo emisora del evento.
         */
        static TARGET_SHEET_CHAGE := "TARGET_SHEET_CHAGE" ;
    } ;

    /**
     * @private
     * Manejador de eventos para la aplicación de Microsoft Excel.
     * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.application?view=excel-pia#events
     */
    class _ApplicationEventHandler
    {
        /**
         * @private
         * Ocurre al crear un nuevo libro.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        NewWorkbook(workbook, application)
        {
            OutputDebug('EVENTO (NewWorkbook): Se ha creado un nuevo libro de trabajo llamado "' workbook.Name '".`n')
            ExcelEventController.Trigger(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_NEW, workbook, application)
        }

        /**
         * @private
         * Ocurre al abrir cualquier libro guardado.
         * 
         * - Si previamente existe un único libro en blanco, Excel lo cerrará tras unos milisegundos.
         * 
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        WorkbookOpen(workbook, application)
        {
            OutputDebug('EVENTO (WorkbookOpen): Se ha abierto un libro de trabajo llamado "' workbook.Name '".`n')
            ExcelEventController.Trigger(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_OPEN, workbook, application)
        }

        /**
         * @private
         * Ocurre cuando el usuario intenta cerrar cualquier libro de trabajo.
         * @param {VarRef<Boolean>} cancel Falso por defecto. Si se establece en Verdadero, 
         * no se permitirá el cierre del libro de trabajo.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        WorkbookBeforeClose(&cancel, workbook, application)
        {
            OutputDebug('EVENTO (WorkbookBeforeClose): Se ha intentando cerrar el libro de trabajo llamado "' workbook.Name '" con el parámetro "cancel"="' cancel '".`n')
            ExcelEventController.Trigger(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_BEFORE_CLOSE, &cancel, workbook, application)
        }

        /**
         * @private
         * Ocurre tras finalizar el guardado de cualquier libro de trabajo.
         * @param {Boolean} success Si el guardado ha sido exitoso.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         * @param {Microsoft.Office.Interop.Excel.Application} application Microsoft Excel COM.
         */
        WorkbookAfterSave(success, workbook, application)
        {
            OutputDebug('EVENTO (WorkbookAfterSave): Se ha guardado el libro de trabajo llamado "' workbook.Name '" con el parámetro "success"="' success '".`n')
            ExcelEventController.Trigger(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_AFTER_SAVE, success, workbook, application)
        }
    } ;

    /**
     * @private
     * Manejador de eventos para un libro de trabajo especifico.
     * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.workbook?view=excel-pia#events
     */
    class _WorkbookEventHandler
    {
        /**
         * @public
         * Permitir el cierre del libro de trabajo objetivo.
         * @type {Boolean} Falso por defecto
         */
        AllowClosure := false ;

        /**
         * @private
         * Ocurre al activar una de las hojas de cálculo del libro de trabajo.
         * @param {Microsoft.Office.Interop.Excel.Worksheet} sheet Puede ser un Worksheet o un Chart.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         */
        SheetActivate(sheet, workbook)
        {
            OutputDebug('EVENTO (SheetActivate): Se ha activado la hoja de cálculo llamada "' sheet.Name '", correspondiente al libro de trabajo llamado "' workbook.Name '".`n')
            ExcelEventController.Trigger(ExcelEventController.WorkbookEventEnum.SHEET_ACTIVATE, sheet, workbook)
        }

        /**
         * @private
         * Ocurre cuando el usuario intenta cerrar el libro de trabajo.
         * 
         * Controla el cierre del libro de trabajo y lo guarda automáticamente para evitar interferencias de la interfaz,
         * garantizando su cierre. 
         * 
         * @param {Boolean} cancel Falso por defecto, si se establece en Verdadero, 
         * no se permitirá el cierre del libro de trabajo.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         */
        BeforeClose(&cancel, workbook)
        {
            OutputDebug('EVENTO (BeforeClose): Se ha intentando cerrar el libro de trabajo conectado "' workbook.Name '" con el parámetro "cancel"="' cancel '".`n')
            
            if (!this.AllowClosure) {
                ExcelEventController.Trigger(ExcelEventController.WorkbookEventEnum.TARGET_WORKBOOK_CLOSE_DENIED, cancel, workbook)
                return cancel := true
            }
            if (workbook.Saved = false) {
                workbook.Save()
            }

            ExcelEventController.Trigger(ExcelEventController.WorkbookEventEnum.TARGET_WORKBOOK_BEFORE_CLOSE, cancel, workbook)
        }

        /**
         * @private
         * Ocurre tras finalizar el guardado del libro de trabajo.
         * 
         * - El libro continúa siendo accesible tras guardarlo en formato xlsx.
         * 
         * @param {Boolean} success Si el guardado ha sido exitoso.
         * @param {Microsoft.Office.Interop.Excel.Workbook} workbook Libro de trabajo emisor del evento.
         */
        AfterSave(success, workbook)
        {
            OutputDebug('EVENTO (AfterSave): Se ha guardado el libro de trabajo conectado "' workbook.Name '" con el parámetro "success"="' success '".`n')
            ExcelEventController.Trigger(ExcelEventController.WorkbookEventEnum.TARGET_WORKBOOK_AFTER_SAVE, success, workbook)
        }
    } ;

    /**
     * @private
     * Manejador de eventos para una hoja de cálculo especifica.
     * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.worksheet?view=excel-pia#events
     */
    class _WorksheetEventHandler
    {
        /**
         * @private
         * Ocurre cuando un rango de la hoja sufre un cambio.
         * @param {Microsoft.Office.Interop.Excel.Range} target Rango modificado.
         * @param {Microsoft.Office.Interop.Excel.Worksheet} worksheet Hoja de cálculo emisora del evento.
         */
        Change(target, worksheet)
        {
            OutputDebug('EVENTO (Worksheet.Change): Se ha producido un cambio en el rango "' target.Address '" de la hoja de cálculo llamada "' worksheet.Name '".`n')
            ExcelEventController.Trigger(ExcelEventController.WorksheetEventEnum.TARGET_SHEET_CHAGE, target, worksheet)
        }
    } ;
}
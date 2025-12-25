#Requires AutoHotkey v2.0
#Include "Dependencies\ExcelManager.ahk"
#Include "Dependencies\ExcelEventHandlers.ahk"
#Include "Dependencies\ChromeManager.ahk"
#Include "Dependencies\IController.ahk"
#Include "Dependencies\Utils.ahk"

AyesaGUI()

;// Los manejadores de eventos no aceptan métodos de instancia,
;// se deben utilizar métodos estáticos o anónimos.
;OnExit((ExitReason, ExitCode) => _gui.ExitHandler(ExitReason, ExitCode))

class AyesaGUI
{
    ;// Constantes

    _maxGuiWidth := 300

    ;// Variables

    GUIComponents := unset
    _excelManager := unset
    _chromeManager := unset

    __New()
    {
        this.GUIComponents := this.InitGUI()
        this._excelManager := Utils.GetExcelManager(IController(this))
        this._chromeManager := Utils.GetChromeManager(IController(this))
    }

    /**
     * Devuelve todos los componentes de la interfaz gráfica de usuario.
     * @returns {Object} 
     */
    InitGUI() 
    {
        mainForm := Gui("+AlwaysOnTop +Border +Caption +Resize -MaximizeBox +MinSize" this._maxGuiWidth "x160")
        mainForm.OnEvent("Close", (*) => ExitApp())
        inputWbLabel := mainForm.Add("Text",, "Conecta un libro de lectura")
        mainForm.Add("Text", "x+m", "|") ;// Separador
        inputWsLabel := mainForm.Add("Text", "x+m", "Hoja activa: 0")
        setInputWbButton := mainForm.Add("Button", "xm", "Conectar libro de lectura activo")
        setInputWbButton.OnEvent("Click", (*) => this.SetInputWbButtonHandler())
        setInputWbButton.ToolTip := "Establece el libro de trabajo que servirá de entrada de datos para la consulta"

        outputWbLabel := mainForm.Add("Text",, "Conecta un libro de escritura")
        mainForm.Add("Text", "x+m", "|") ;// Separador
        outputWsLabel := mainForm.Add("Text", "x+m", "Hoja activa: 0")
        setOutputWbButton := mainForm.Add("Button", "xm", "Conectar libro de escritura activo")
        setOutputWbButton.OnEvent("Click", (*) => this.SetOutputWbButtonHandler())
        setOutputWbButton.ToolTip := "Establece el libro de trabajo que servirá de salida de datos para la consulta"

        mainForm.Add("Text",, "Funciones:")
        cumasButton := mainForm.Add("Button", , "Consulta CUMAS")
        cumasButton.OnEvent("Click", (*) => this.QueryCUMAS())
        cumasButton.ToolTip := "Realiza una consulta masiva en CUMAS de todos los CUPS de la columna CUPS en el libro de lectura y traspone los datos al libro de escritura"

        ;OnMessage(0x0200, On_WM_MOUSEMOVE) ;// Desencadenado cuando se hace focus en un control de la IGU

        mainForm.Show()

        return {
            MainForm: mainForm,
            ReadWbLabel: inputWbLabel,
            WriteWbLabel: outputWbLabel,
            ReadWsLabel: inputWsLabel,
            WriteWsLabel: outputWsLabel,
            SetReadWbButton: setInputWbButton,
            SetWriteWbButton: setOutputWbButton,
        }

        On_WM_MOUSEMOVE(wParam, lParam, msg, Hwnd) {
            static PrevHwnd := 0
            if (Hwnd != PrevHwnd) {
                Text := "", ToolTip() ; Turn off any previous tooltip.
                CurrControl := GuiCtrlFromHwnd(Hwnd)
                if CurrControl {
                    if !CurrControl.HasProp("ToolTip")
                        return ; No tooltip for this control.
                    Text := CurrControl.ToolTip
                    SetTimer () => ToolTip(Text), -1000
                    SetTimer () => ToolTip(), -7000 ; Remove the tooltip.
                }
                PrevHwnd := Hwnd
            }
        }
    }

    ;-----------------------------------------
    ; EVENTOS
    ;-----------------------------------------

    SetInputWbButtonHandler(*)
    {
        this._excelManager.ConnectActiveWorkBook()
    }

    SetOutputWbButtonHandler(*)
    {
        ;// TODO - SetOutputWbButtonHandler(*)
    }

    excelStateChanged(connected)
    {
        this.GUIComponents.ReadWbLabel.Text := connected ? "Libro conectado: " this._excelManager.Workbook.Name : "Conecta un libro de trabajo"
        this.GUIComponents.ReadWsLabel.Text := connected ? "Hoja activa: " this._excelManager.Worksheet.Name : "Hoja activa: 0"
        if (connected)
            WinActivate(this._excelManager.Workbook.Name)
    }

    workbookCloseDenied() {
        Utils.TopMostMsgBox(
            "El cierre del libro de trabajo seleccionado ha sido denegado por AutoHotKey.", 
            "Información"
        )
    }

    excelProcessClosed() {
        Utils.TopMostMsgBox("Excel ha sido cerrado abruptamente, la aplicación se cerrará.", "Error")
        ExitApp
    }

    chromeProcessClosed() {
        Utils.TopMostMsgBox("Chrome ha sido cerrado abruptamente, la aplicación se cerrará.", "Error")
        ExitApp
    }


    ;-----------------------------------------
    ; FUNCIONES
    ;-----------------------------------------

    ;// CUMAS CARGA 17 CONSULTAS POR MINUTO.
    ;// TODO - CONSULTAR TAMBIÉN SUMINISTROS DE AUTOCONSUMO
    QueryCUMAS() 
    {
        mainPage := this._chromeManager.GetPage("http://10.123.125.205/WebCUMAS/login")
        loginSucc := this._chromeManager.Login(mainPage, "#username", "#password", "[class='btn btn-lg btn-primary btn-block']", "#messageError", "EX151296", "LCgi6Iidx2s8*")["value"]
        if (!loginSucc) {
            throw Error("No ha sido posible iniciar sesión en CUMAS, revisa tus credenciales.", -1)
        }

        ;// TODO - Pinchar en el botón de LECTURAS NORMALES
        mainPage.Evaluate2("")
        ;// TODO - Abrir pestaña para AUTOCONSUMO
        autoconsumoPage := this._chromeManager.GetPage("")

        ;// TODO - Leer del libro de lectura: CUPS y AUTOCONSUMO
        cupsStr :=
            (
                ''
            )
        cupsArray := StrSplit(cupsStr, "`n")


        ;// Seleccionar AMM
        jsSelectAMM := 'document.querySelector("#idOrigenCurva").selectedIndex = 1; document.querySelector("#idOrigenCurva").dispatchEvent(new Event("change", { bubbles: true }));'
        mainPage.Evaluate(jsSelectAMM)
        autoconsumoPage.Evaluate(jsSelectAMM)

        fullQueryArray := []

        for cups in cupsArray {

            jsQueryCUPS := 
                (
                    'document.querySelector("#cups").value = "' cups '";
                    document.querySelector("#cups").dispatchEvent(new Event("change", { bubbles: true }));
                    document.querySelector("#button-upload").click();'
                )

            ;// Insertar CUPS y leer la tabla de la 1a página
            if (true) {
                mainPage.Evaluate2(jsQueryCUPS)
                objs := this._chromeManager.GetTableEntries(mainPage, '#lectura')
            } else {
                autoconsumoPage.Evaluate2(jsQueryCUPS)
                objs := this._chromeManager.GetTableEntries(autoconsumoPage, '#lectura')
            }

            ;// Si la consulta está vacía, registrar solo el CUPS
            if (objs.length = 0) {
                fullQueryArray.Push({CUPS: cups})
                continue
            }

            for row in objs
                fullQueryArray.Push(row)
        }

        this._excelManager.InsertObjs(fullQueryArray)
        
        MsgBox "La consulta masiva en CUMAS ha finalizado. Si los valores no se han volcado en el libro de escritura, pégalos manualmente.", "Información"
    }
}
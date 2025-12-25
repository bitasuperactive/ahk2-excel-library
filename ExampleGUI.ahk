#Requires AutoHotkey v2.0
#SingleInstance Force
#Include "Dependencies\ExcelManager.ahk"
#Include "Dependencies\ChromeManager.ahk"
#Include "Dependencies\Util\Utils.ahk"
#Include "Dependencies\Util\OrObject.ahk"

GUIComponents := InitGUI()
ExcelMan := unset
ChromeMan := 0 ;//Utils.GetChromeManager(IController(MyController()))

/**
 * Devuelve todos los componentes de la interfaz gráfica de usuario.
 * @returns {Object} 
 */
InitGUI()
{
    ;// Controles ; +AlwaysOnTop
    mainForm := Gui("+Border +Caption +Resize +MinSize600x380")
    connectExcelBtn := mainForm.AddButton("h30 w100", "Conectar Excel")
    disconnectExcelBtn := mainForm.AddButton("x+m h30 w100 Disabled", "Desconectar")

    excelGroup := mainForm.AddGroupBox("xm y+10 w580 h160 Section", "EXCEL")
    mainForm.AddText("xp+10 yp+20", "Libro de lectura")
    ReadWorkbookListBox := mainForm.AddListBox("xs+10 y+m w150 h120")
    openReadWorkbookBtn := mainForm.AddButton("x+-19 yp-22 w20 h20", "📁")
    mainForm.AddText("x+m ys+20", "Libro de escritura")
    WriteWorkbookListBox := mainForm.AddListBox("xp y+m w150 h120")
    openWriteWorkbookBtn := mainForm.AddButton("x+-19 yp-22 w20 h20", "📁")
    workbookListBoxesItems := Map()

    createTableBtn := mainForm.AddButton("x+20 ys+20 h20 w220 Disabled", "Crear/Anexar tabla de ejemplo")
    readTableBtn := mainForm.AddButton("xp y+m h20 w220 Disabled", "Leer tabla")
    validateTableBtn := mainForm.AddButton("xp y+m h20 w220 Disabled", "Validar tabla")
    mainForm.AddText("xp y+10", "Copiar columna:")
    columnDropDown := mainForm.AddDropDownList("x+m yp-3 w80")
    shortcutEdit := mainForm.AddEdit("x+m yp w30 Disabled", "F1")

    connectChromeBtn := mainForm.AddButton("xm ys+180 h30 w100", "Conectar Chrome")
    disconnectChromeBtn := mainForm.AddButton("x+m h30 w100 Disabled", "Desconectar")
    chromeGroup := mainForm.AddGroupBox("xm y+10 w580 h50 Section", "CHROME")
    copyTableFromWebBtn := mainForm.AddButton("xs+10 ys+20 h20 w220 Disabled", "Copiar primera tabla")

    excelAndChromeGroup := mainForm.AddGroupBox("xm ys+60 w580 h50 Section Disabled", "EXCEL + CHROME")
    tableFromWebToExcelBtn := mainForm.AddButton("xs+10 ys+20 h20 w220 Disabled", "Copiar tabla web en Excel")

    ;// Eventos
    mainForm.OnEvent("Close", (*) => ExitApp())
    connectExcelBtn.OnEvent("Click", (*) => ConnectExcel())
    disconnectExcelBtn.OnEvent("Click", (*) => DisconnectExcel())
    ReadWorkbookListBox.OnEvent("Change", readWorkbook_ListBox_Change)
    WriteWorkbookListBox.OnEvent("Change", writeWorkbook_ListBox_Change)
    openReadWorkbookBtn.OnEvent("Click", (*) => openWorkbookFile(ExcelManager.ConnectionTypeEnum.READ))
    openWriteWorkbookBtn.OnEvent("Click", (*) => openWorkbookFile(ExcelManager.ConnectionTypeEnum.WRITE))

    createTableBtn.OnEvent("Click", (*) => CreateExampleTable())
    readTableBtn.OnEvent("Click", (*) => ReadTable())
    validateTableBtn.OnEvent("Click", (*) => ValidateTableHeaders())

    mainForm.Show()

    return {
        MainForm: mainForm,
        ConnectExcelBtn: connectExcelBtn,
        DisconnectExcelBtn: disconnectExcelBtn,
        ExcelGroup: excelGroup,
        ReadWorkbookListBox: ReadWorkbookListBox,
        WriteWorkbookListBox: WriteWorkbookListBox,
        WorkbookListBoxesItems : workbookListBoxesItems,
        CreateTableBtn: createTableBtn,
        ReadTableBtn : readTableBtn,
        ValidateTableBtn : validateTableBtn,
        ColumnDropDown : columnDropDown
    }
}

ConnectExcel()
{
    GUIComponents.ConnectExcelBtn.Enabled := false
    GUIComponents.ConnectExcelBtn.Text := "Cargando ..."
    
    try global ExcelMan := ExcelManager(true)
    catch Error as err {
        Utils.TopMostMsgBox(err.Message, "Error")
        GUIComponents.ConnectExcelBtn.Text := "Conectar Excel"
        GUIComponents.ConnectExcelBtn.Enabled := true
        return
    }

    ExcelEventController.OnEvent(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_NEW, UpdateWorkbookLists)
    ExcelEventController.OnEvent(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_OPEN, (*) => SetTimer(UpdateWorkbookLists, -150))
    ExcelEventController.OnEvent(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_AFTER_SAVE, UpdateWorkbookLists)
    ExcelEventController.OnEvent(ExcelEventController.ApplicationEventEnum.ANY_WORKBOOK_BEFORE_CLOSE, (*) => SetTimer(UpdateWorkbookLists, -150))
    ExcelEventController.OnEvent(ExcelEventController.ApplicationEventEnum.APPLICATON_TERMINATED, excelProcessClosed)
    ExcelEventController.OnEvent(ExcelEventController.WorkbookEventEnum.TARGET_WORKBOOK_CLOSE_DENIED, targetWorkbookCloseDenied)
    UpdateWorkbookLists()

    GUIComponents.ExcelGroup.Text .= " ✅"
    GUIComponents.ConnectExcelBtn.Text := "Conectar Excel"
    GUIComponents.DisconnectExcelBtn.Enabled := true
}

DisconnectExcel()
{
    GUIComponents.DisconnectExcelBtn.Enabled := false
    GUIComponents.DisconnectExcelBtn.Text := "Cargando ..."

    ExcelMan.Dispose()
    GUIComponents.WorkbookListBoxesItems.Clear()

    GUIComponents.ReadWorkbookListBox.Delete()
    GUIComponents.WriteWorkbookListBox.Delete()
    GUIComponents.ExcelGroup.Text := "EXCEL"
    GUIComponents.CreateTableBtn.Enabled := false
    GUIComponents.ReadTableBtn.Enabled := false
    GUIComponents.ValidateTableBtn.Enabled := false
    GUIComponents.ConnectExcelBtn.Enabled := true
    GUIComponents.DisconnectExcelBtn.Text := "Desconectar"
}

UpdateWorkbookLists(*)
{
    actualWorkbookNames := ExcelMan.GetAllOpenWorkbooksNames()

    for wbName, index in GUIComponents.WorkbookListBoxesItems {
        if (!actualWorkbookNames.Has(index) || actualWorkbookNames[index] != wbName)
        {
            GUIComponents.ReadWorkbookListBox.Delete(index)
            GUIComponents.WriteWorkbookListBox.Delete(index)
            GUIComponents.WorkbookListBoxesItems.Delete(wbName)
        }
    }
    for wbName in actualWorkbookNames {
        if (GUIComponents.WorkbookListBoxesItems.Has(wbName))
            continue
        GUIComponents.ReadWorkbookListBox.Add([wbName])
        GUIComponents.WriteWorkbookListBox.Add([wbName])
        GUIComponents.WorkbookListBoxesItems.Set(wbName, GUIComponents.WorkbookListBoxesItems.Count + 1)
    }
}

;-----------------------------------------
; EVENTOS
;-----------------------------------------

openWorkbookFile(connType)
{
    path := FileSelect(1,, "", "Libro de excel (*.xlsx)")
    If (!path || !InStr(path, ".xlsx"))
        return
    Run(path)
    
    Sleep 1000

    if (!IsSet(ExcelMan))
        ConnectExcel()
    UpdateWorkbookLists()
    lastWorkbookIndex := GUIComponents.WorkbookListBoxesItems.Count

    switch(connType) {
        case ExcelManager.ConnectionTypeEnum.READ:
            ControlChooseIndex(lastWorkbookIndex, GUIComponents.ReadWorkbookListBox)
        case ExcelManager.ConnectionTypeEnum.WRITE:
            ControlChooseIndex(lastWorkbookIndex, GUIComponents.WriteWorkbookListBox)
        default:
            throw TypeError('Se esperaba el tipo "' ExcelManager.ConnectionTypeEnum.__Class '" pero se ha recibido: ' Type(connType))
    }
}

readWorkbook_ListBox_Change(*)
{
    selectedWbName := GUIComponents.ReadWorkbookListBox.Text
    if (selectedWbName = "")
        return

    GUIComponents.ReadWorkbookListBox.Enabled := false
    ExcelMan.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.READ, selectedWbName, false)

    GUIComponents.ReadTableBtn.Enabled := true
    GUIComponents.ValidateTableBtn.Enabled := true
    GUIComponents.ReadWorkbookListBox.Enabled := true

    headerObj := ExcelMan.ReadWorkbookAdapter.ReadRow(1)
    __ShowHeaderOnListBox(headerObj)

    __ShowHeaderOnListBox(headers)
    {
        GUIComponents.ColumnDropDown.Enabled := false
        GUIComponents.ColumnDropDown.Delete()
        arr := []
        for header in headers.OwnProps()
            arr.Push(header)
        GUIComponents.ColumnDropDown.Add(arr)
        GUIComponents.ColumnDropDown.Enabled := true
    }
}

writeWorkbook_ListBox_Change(*)
{
    selectedWbName := GUIComponents.WriteWorkbookListBox.Text
    if (selectedWbName = "")
        return
        
    GUIComponents.WriteWorkbookListBox.Enabled := false
    ExcelMan.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.WRITE, selectedWbName, false)
    
    GUIComponents.CreateTableBtn.Enabled := true
    GUIComponents.WriteWorkbookListBox.Enabled := true
}

targetWorkbookCloseDenied(caller, cancel, wb) {
    global ExcelMan
    global GUIComponents
    Utils.TopMostMsgBox(
        "El libro de trabajo `"" wb.Name "`""
        " está siendo utilizado por " GUIComponents.MainForm.title ", y no se puede cerrar.",
        "Información"
    )
}

excelProcessClosed(*) {
    Utils.TopMostMsgBox("Excel ha sido cerrado abruptamente.", "Error")
    DisconnectExcel()
}

chromeProcessClosed(*) {
    Utils.TopMostMsgBox("Chrome ha sido cerrado abruptamente, la aplicación se cerrará.", "Error")
    ExitApp
}


;-----------------------------------------
; FUNCIONES
;-----------------------------------------

CreateExampleTable()
{
    GUIComponents.CreateTableBtn.Enabled := false

    objArray := []
    objArray.Push(OrObject(
        "Cuenta", "Valor Cuenta 1",
        "Nombre", "Valor Nombre 1",
        "Apellido", "Valor Apellido 1",
        "Dirección", "Valor Dirección 1",
        "Teléfono", "Valor Teléfono 1",
        "NewProp1", "Valor NewProp 1"
    ))
    objArray.Push(OrObject(
        "Cuenta", "Valor Cuenta 2",
        "Nombre", "Valor Nombre 2",
        "Apellido", "Valor Apellido 2",
        "Dirección", "Valor Dirección 2",
        "Teléfono", "Valor Teléfono 2",
        "NewProp2", "Valor NewProp 2"
    ))

    ExcelMan.WriteWorkbookAdapter.AppendTable(objArray)

    __FillMissingValue()
    __AddColumns()
    
    GUIComponents.CreateTableBtn.Enabled := true

    __FillMissingValue()
    {
        ;// Corregir primer objeto
        corrObj := {
            NewProp1: ">Valor NewProp 1<"
        }
        ExcelMan.WriteWorkbookAdapter.FillBlankFieldsOnRow(2, corrObj)
    }

    __AddColumns()
    {
        obj := {
            Columna1:"",
            Columna2:"",
            Columna3:""
        }
        ExcelMan.WriteWorkbookAdapter.AppendTable(obj)
    }
}

ReadTable()
{
    GUIComponents.ReadTableBtn.Enabled := false

    ;objs := ExcelMan.ReadWorkbookAdapter.ReadTable()
    Loop ExcelMan.ReadWorkbookAdapter.GetRowCount() {
        str := ""
        ;obj := objs[A_Index] 
        obj := ExcelMan.ReadWorkbookAdapter.ReadRow(A_Index)
        for name, value in obj.OwnProps() {
            str := str name ": " value "`n"
        }
        ;// MsgBox añade +>1 segundo por cada ejecución
        Utils.TopMostMsgBox("[ OBJETO " A_Index " ]`n" str, "TEST")
    }
    
    GUIComponents.ReadTableBtn.Enabled := true
}


ValidateTableHeaders()
{
    GUIComponents.ValidateTableBtn.Enabled := false
    headers := ["APELLIDOS", "nombre", "Dirección", "Teléfono", "Cuenta", "NewProp"]
    validation := ExcelMan.ReadWorkbookAdapter.ValidateHeaders(headers, &missingHeaders := []) ; No hace falta validar manualmente
    Utils.TopMostMsgBox("Validación de cabeceras: " validation, "TEST")
    str := ""
    for header in missingHeaders
        str .= header ", "
    MsgBox "Faltan las siguientes cabeceras: " str
    GUIComponents.ValidateTableBtn.Enabled := true
}

















InsertTableFromWebIntoExcel()
{
    tableQuery := '#scrollTable' ;// [class="table table-light traversal-table"]
    page := ChromeMan.GetPage("https://wdi.worldbank.org/table/3.2")
    objs := ChromeMan.GetTableEntries(page, tableQuery)
    
    ExcelMan.CreateTable(objs)
}

LoginInWeb()
{
    page := ChromeMan.GetPage("https://www.saucedemo.com/")
    ;ChromeMan.SetValueWithNativeSetter(page, "#user-name", "standard_user")
    ;ChromeMan.SetValueWithNativeSetter(page, "#password", "secret_sauce")
    var := ChromeMan.Login(page, "#user-name", "#password", "#login-button", "[class='error-button']", "standard_user", "secret_sauce")
    Utils.TopMostMsgBox(var["value"], "TEST")
}
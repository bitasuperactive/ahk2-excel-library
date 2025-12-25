#Requires AutoHotkey v2.0
#Include "Dependencies\ChromeManager.ahk"


URL := "https://portal.trianglerrhh.es/ETT/Trabajadores/Candi/CapturaPresencia.aspx"
GUIComps := InitGUI()
ChromeMan := ChromeManager(URL)
cancel := false

/**
 * Devuelve todos los componentes de la interfaz gráfica de usuario.
 * @returns {Object} 
 */
InitGUI()
{
    mainForm := Gui("+AlwaysOnTop +Border +Caption -MaximizeBox")
    mainForm.Title := "AutoFichador"
    mainForm.OnEvent("Close", (*) => ExitApp()) ;// Falta chrome kill
    label1 := mainForm.Add("Text",, "Mes:")
    monthChoice := mainForm.Add("DropDownList", "vMonthChoice x+m", ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"])
    monthChoice.Value := FormatTime(A_Now, "M")
    runButton := mainForm.Add("Button", "x+m", "Ejecutar")
    runButton.OnEvent("Click", (*) => AutoFileAllWorkDaysAsCommonHours())
    cancelButton := mainForm.Add("Button", "x+m", "Cancelar")
    cancelButton.OnEvent("Click", (*) => CancelOperation())
    mainForm.Show()

    return {
        MainForm: mainForm,
        Label1: label1,
        MonthChoice: monthChoice,
        RunButton: runButton
    }
}


/**
 * Ficha de 8 a 16 todos los días laborales del mes seleccionado, en horario normal.
 */
AutoFileAllWorkDaysAsCommonHours()
{
    year := FormatTime(A_Now, "yyyy")
    month := GUIComps.MonthChoice.Value
    horaEntrada := 8
    horaSalida := 15

    if (!page := ChromeMan.IsPageOpened(URL)) {
        MsgBox('Primero debes iniciar sesión y acceder a "Grabar Partes > Nuevo Parte".', "Error")
        return
    }
    answer := MsgBox("Vas a registrar los turnos trabajados (L a V) para el mes " month " en horario de " horaEntrada " a " horaSalida ",`n¿Es correcto?", 
                "Atención", 
                'OKCancel Owner' GUIComps.MainForm.Hwnd)
    if (answer = "Cancel")
        return

    loop 31 {
        day := A_Index
        date := Format("{1:04}{2:02}{3:02}", year, month, day) ;// TODO - Utils?
        weekDay := FormatTime(date, "WDay")
        if (weekDay = 7 OR weekDay = 1) ;// Sábado o domingo
            continue

        datesMissed := []
        try {
            JSAutoFile(date, horaEntrada, horaSalida, false)
        }
        catch (Error as er) {
            MsgBox(er.Message, "Error")
            break
        }
        if (cancel)
            break
        if (A_Index = 31)
            MsgBox "Hecho.", "Info"
    }

    GUIComps.MainForm.Hide()
    ExitApp
}

CancelOperation()
{
    global cancel := true
}

/**
 * TriangleSolutionsRRHH auto-fichaje como temporal, porque su web es una mierda.
 * @param {String} date Fecha a fichar con formato "YYYYMMDD".
 * @param {Integer} checkInHour Hora 24,00 de entrada.
 * @param {Integer} checkOutHour Hora 24,00 de salida.
 * @param {Boolean} specialCheck (Opcional) Contabilizar horas como festivas.
 * @returns 1 para fichaje exitoso, 0 ante error.
 * @link https://portal.trianglerrhh.es/ETT/Trabajadores/Candi/CapturaPresencia.aspx
 */
JSAutoFile(date, checkInHour, checkOutHour, specialCheck := false)
{
    if (Type(date) != "String" || B := RegExMatch(date, "\d{6,12}") = 0 || C := !FormatTime(date, "dd"))
        throw TypeError("El valor `"" date "`" debe ser una fecha válida con formato `"YYYYMMDDHH24MISS`".", -1)
    if ((checkOutHour - checkInHour) <= 0)
        throw ValueError("La hora de entrada no puede ser igual o posterior a la hora de salida.", -1)
    if (!page := ChromeMan.IsPageOpened(URL))
        throw Error('Primero debes iniciar sesión y acceder a "Grabar Partes > Nuevo Parte".', -1)

    year := FormatTime(date, "yyyy")
    month := FormatTime(date, "MM")
    day := FormatTime(date, "dd")

    jsFillForm :=
    (
        '{
            let date = new Date(' year ',' month '-1,' day ') // Enero es el mes 0
            let dateControl = ASPxClientControl.GetControlCollection().GetByName("tbfeparte");
            let inHourControl = ASPxClientControl.GetControlCollection().GetByName("teHoraEntrada1");
            let outHourControl = ASPxClientControl.GetControlCollection().GetByName("teHoraSalida1");
            let conceptControl = ASPxClientControl.GetControlCollection().GetByName("txtConcepto0");
            let spConceptControl = ASPxClientControl.GetControlCollection().GetByName("txtConcepto8");
            dateControl.SetValue(new Date(date))
            inHourControl.SetValue(new Date(0,0,0,' checkInHour '))
            outHourControl.SetValue(new Date(0,0,0,' checkOutHour '))
            if (' specialCheck ') {
                spConceptControl.SetValue("' checkOutHour - checkInHour '")
                conceptControl.SetValue("0")
            }
            else {
                conceptControl.SetValue("' checkOutHour - checkInHour '")
                spConceptControl.SetValue("0")
            }
        }'
    )
    jsSubmit := 'document.querySelector("#botAceptarInsertar_CD").click()'
    jsErrorThrown := 'document.querySelector("#lblError") ? true : false'
    jsErrorMessage := 'document.querySelector("#lblError").innerText'
    jsContinue := 'document.querySelector("#btOK_I")?.click()'

    page.WaitForLoad()
    page.Evaluate(jsFillForm)
    page.Evaluate(jsSubmit)
    page.WaitForLoad()
    errorThrown := page.Evaluate(jsErrorThrown)["value"]
    if (errorThrown) {
        errorMsg := page.Evaluate(jsErrorMessage)["value"]
        if (!InStr(errorMsg, "Ya se ha introducido un parte"))
            return JSAutoFile(date, checkInHour, checkOutHour)
    }
    if (!errorThrown) {
        page.Evaluate(jsContinue)
        return 1
    }
    return 0
}
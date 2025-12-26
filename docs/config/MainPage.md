\mainpage Inicio

<!--
LIMITACIONES DE DOXYGEN C++
En los siguientes tipos no reconoce su delimitación y hay
que poner un ';' al final de la definición:
- '=>' externos, los internos de funciones no dan problemas
- ':='
- clases anidadas

Tipos incompatibles:
- (*)           Hay que darle nombre (p*)
- "extends"     Hay que quitarlo antes de parsear y documentarlo con @extends
- @type
-->

**Excel Library** puede definirse como un adaptador entre los libros de trabajo de Microsoft Excel y <a href="https://www.autohotkey.com/v2/" target="_blank">⇱AutoHotkey V2</a>.

**Ha sido diseñada con un propóstio muy concreto en mente:** Automatizar la lectura de datos locales y la escritura de datos externos mientras el usuario sigue trabajando manualmente en Excel. 

La idea surgió de mi experiencia como administrativo en el sector BPO. Observé que gran parte de mi trabajo consistía en contrastar datos procedentes de múltiples fuentes web, y pensé que si pudiera automatizar la recopilación de toda esa información diaria, conseguiría deshacerme de la parte más boluptuosa y cargante de mi trabajo, pertimiéndome centrar mi atención en el análisis de la información, que es lo interesante.
La primera parte de esta idea se consolida en **Excel Library**. La segunda consistirá en la integración de Google Chrome en una librería independiente.

Ya he conseguido implementar varias versiones de este proyecto en mis propios flujos de trabajo, y aunque queda camino por recorrer, el objetivo está cada vez más cerca 🏖.

## Características

- 💡 **Minimalista**
<br/>No pretende ser un wrapper completo de Microsoft Interop. Su funcionalidad está limitada a su propósito: leer y escribir datos. 
Aquí no encontrarás una integración completa.

- 👨‍💻 **Compatible con el uso paralelo del usuario**
<br/>Controla la interacción mediante eventos y dispone de una función controlador (de la que estoy muy orgulloso) capaz de interrumpir una edición manual para evitar así que el script se rompa.

- 🔐 **Protección de la información**
<br/>Pensada para entornos de negocio, separa explícitamente las funciones de lectura y escritura para preservar la integridad de los datos existentes.
<br/>Se recomienda utilizar una hoja de cálculo para leer y otra para escribir, y una vez procesada toda la información requerida de las fuentes externas, se portaría manualmente a la hoja de cálculo principal. 
Esta funcionalidad es opcional, pero añade una capa extra de seguridad.

## Ejemplo básico

Dependencias (OrObject es opcional):

@code
#Include "ExcelLibrary\ExcelManager.ahk"
#Include "Util\OrObject.ahk"
@endcode

Conectarse al COM de Excel es tan fácil como inicializar ExcelManager:

@code
ExcelMan := ExcelManager(true) ; 'true' permite leer y escribir en la misma hoja
@endcode

Lo único que necesitas para empezar a automatizar tus libros de trabajo,
es definir una hoja de escritura y otra (o la misma) de lectura:

@code
;// Obtener los nombres de todos los libros `.xlsx` abiertos
workbookNames := ExcelMan.GetAllOpenWorkbooksNames()

;// Conectarse al libro1-hoja1 (hoja activa) para escribir
;// Así habilitamos el adaptador de escritura WriteWorkbookAdapter
ExcelMan.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.WRITE, workbookNames[1])

;// Conectarse al libro1-hoja1 (hoja activa) para leer
;// Así habilitamos el adaptador de lectura ReadWorkbookAdapter
ExcelMan.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.READ, workbookNames[1])
@endcode

De esta manera habilitamos los adaptadores que nos permitirán meternos en materia:

@code
;// Escribir un objeto en la hoja conectada
;// Utilizamos OrObject para que los objetos se inserten en el orden de creación
;// y no por orden alfabético, pero puedes usar los objetos nativos si no te importa
;// el orden
;// OrObject funciona como un objeto normal, exceptuando la inicialización directa como 
;// en el siguiente caso
obj := OrObject(
    "Cuenta", "Valor Cuenta 1",
    "Nombre", "Valor Nombre 1",
    "Apellido", "Valor Apellido 1",
    "Dirección", "Valor Dirección 1",
    "Teléfono", 689068093
)
ExcelMan.WriteWorkbookAdapter.AppendTable(obj) ; Fíjate en que las cabeceras se normalizan
 

;// Leer la tabla que hemos creado
objs := ExcelMan.ReadWorkbookAdapter.ReadTable()

;// Mostrar objetos leídos
Loop ExcelMan.ReadWorkbookAdapter.GetRowCount() {
    str := ""
    for name, value in objs[A_Index].OwnProps() {
        str := str name ": " value "`n"
    }
    MsgBox("[ FILA " A_Index " ]`n" str)
}
@endcode

Una vez hemos terminado de trabajar con los libros, podemos desconectarlos explícitamente mediante [DisconnectWorkbook](#ExcelManager::DisconnectWorkbook) o simplemente conectar otros.

@note Las instancias se auto-desechan al cerrar el script.

#### 🚀 ¡Pruébalo en tu script!

Hala, y ahora arremete sin miedo contra la documentación de clases. Ha sido escrita con mimo y es muy sencillita, espero que te sirva 😉.

## Métodos y clases esenciales

#### [ExcelManager](#ExcelManager::__New)
> @copydoc ExcelManager::__New
> <br/><br/>

#### [GetAllOpenWorkbooksNames](#ExcelManager::GetAllOpenWorkbooksNames)
> @copydoc ExcelManager::GetAllOpenWorkbooksNames
> <br/><br/>

#### [ConnectionTypeEnum](#ExcelManager::ConnectionTypeEnum)
> @copydoc ExcelManager::ConnectionTypeEnum
> **Tipos**<br/>
> [READ](#ExcelManager::ConnectionTypeEnum::READ).- @copybrief ExcelManager::ConnectionTypeEnum::READ <br/>
> [WRITE](#ExcelManager::ConnectionTypeEnum::WRITE).- @copybrief ExcelManager::ConnectionTypeEnum::WRITE
> <br/><br/>

#### [ConnectWorkbookByName](#ExcelManager::ConnectWorkbookByName)
> @copydoc ExcelManager::ConnectWorkbookByName
> <br/><br/>

#### [WriteWorkbookAdapter](#WriteWorkbookAdapter)
> @copybrief WriteWorkbookAdapter
> <br/><br/>

#### [ReadWorkbookAdapter](#ReadWorkbookAdapter)
> @copybrief ReadWorkbookAdapter
> <br/><br/>

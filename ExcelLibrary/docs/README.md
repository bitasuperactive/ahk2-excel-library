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
-->

# Excel Library

Librería para manejar libros de Excel en AutoHotkey v2.

## Características
- Arquitectura basada en adaptadores
- Documentación compatible con Doxygen

## Ejemplo mínimo

Dependencias (OrObject es opcional):

@code
#Include "Dependencies\ExcelManager.ahk"
#Include "Dependencies\Util\OrObject.ahk"
@endcode

Conectarse al COM de Excel esta tan fácil como inicializar ExcelManager:

@code
;// Establecer conexión con Excel
ExcelMan := ExcelManager(true) ; 'true' permite leer y escribir en la misma hoja
@endcode

Lo único que necesitas para empezar a automatizar tus libros de trabajo,
es definir un libro de escritura y otro (o el mismo) de lectura.

@code
;// Obtener los nombres de todos los libros .xlsx abiertos
workbookNames := ExcelMan.GetAllOpenWorkbooksNames()

;// Conectarse al libro 1 para escribir en él
;// Así habilitamos las funciones del WriteWorkbookAdapter
ExcelMan.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.WRITE, workbookNames[1])

;// Conectarse al libro 1 para leerlo
;// Así habilitamos las funciones del ReadWorkbookAdapter
ExcelMan.ConnectWorkbookByName(ExcelManager.ConnectionTypeEnum.READ, workbookNames[1])
@endcode

De esta manera habilitamos los adaptadores que nos permitirán meternos en materia.

@code
;// Escribir un objeto en la hoja conectada
;// Utilizamos OrObject para que los objetos se inserten en el orden de creación
;// y no por orden alfabético, pero puedes usar objetos normales
;// Exceptuando la inicialización directa como en el siguiente caso, 
;// OrObject funciona como un objeto normal
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

¡Pruébalo en tu script!

@warning Si Excel no está iniciado puede tardar más de la cuenta en permitir el acceso a su COM y lanzar un Error, ¡Reinténtalo!

Hala, y ahora sin miedo métete en la documentación de clases, ha sido escrita con mimo y es muy sencillita. Espero que te sirva 😉.

## Métodos y clases esenciales

#### [ExcelManager](#ExcelManager::__New)
> @copydoc ExcelManager::__New

#### [GetAllOpenWorkbooksNames](#ExcelManager::GetAllOpenWorkbooksNames)
> @copydoc ExcelManager::GetAllOpenWorkbooksNames

#### [ConnectionTypeEnum](#ExcelManager::ConnectionTypeEnum)
> @copydoc ExcelManager::ConnectionTypeEnum <br/>
> <br/>Tipos:<br/>
> [READ](#ExcelManager::ConnectionTypeEnum::READ).- @copybrief ExcelManager::ConnectionTypeEnum::READ <br/>
> [WRITE](#ExcelManager::ConnectionTypeEnum::WRITE).- @copybrief ExcelManager::ConnectionTypeEnum::WRITE

#### [ConnectWorkbookByName](#ExcelManager::ConnectWorkbookByName)
> @copydoc ExcelManager::ConnectWorkbookByName

#### [WriteWorkbookAdapter](#WriteWorkbookAdapter)
> @copybrief WriteWorkbookAdapter

#### [ReadWorkbookAdapter](#ReadWorkbookAdapter)
> @copybrief ReadWorkbookAdapter

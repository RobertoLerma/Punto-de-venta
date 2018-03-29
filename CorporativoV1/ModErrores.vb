'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE ERRORES JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB

Public Module ModErrores
    Dim Errorsito As ADODB.Error

    'Dim cnn As Connection
    'Dim mdcon As ModConexion

    Public Sub Errores()
        Dim MensajeError As String

        'Try
        ' Obtiene el Mensaje de Error
        Select Case Err.Number
            Case 3 ' Return without GoSub
                MensajeError = "RETURN Sin GOSUB"
            Case 5 'Invalid Procedure call or argument
                MensajeError = "Llamada Inválida a un Procedimiento"
            Case 6 ' OverFlow
                MensajeError = "Desbordamiento de Tipo de Dato"
            Case 7 ' Out of Memory
                MensajeError = "Fuera de Memoria"
            Case 9 ' SubScript Out of Range
                MensajeError = "Subíndice Fuera de Rango"
            Case 10 'This Array is Fixed or Temporarily Locked
                MensajeError = "Arreglo Fijo o Bloqueado Temporalmente"
            Case 11 ' Division By Zero
                MensajeError = "No es Posible Dividir Entre Cero"
            Case 13 ' Type Mismatch
                MensajeError = "Tipo de Dato No Coincide"
            Case 14 ' Out of Space String
                MensajeError = "Out Of Space String"
            Case 16 ' Expression Too Complex
                MensajeError = "Expresión Muy Compleja"
            Case 17 ' Can't perform requested operation
                MensajeError = "No se Puede Ejecutar la Operación Requerida"
            Case 18 ' User interrupt occurred
                MensajeError = "Ocurrió una Interrupción"
            Case 20 ' Resume without error
                MensajeError = "RESUME Sin Error"
            Case 28 ' Out of stack space
                MensajeError = "PILA Sin Espacio"
            Case 35 ' Sub or Function not defined
                MensajeError = "SUB o FUNCTION no Definida"
            Case 47 ' Too many DLL application clients
                MensajeError = "Demasiados Clientes de Aplicaciones DLL"
            Case 48 ' Error in loading DLL
                MensajeError = "Error Cargando DLL"
            Case 49 ' Bad DLL calling convention
                MensajeError = "Llamada DLL Incorrecta"
            Case 51 ' Internal Error
                MensajeError = "Error Interno"
            Case 52 ' Bad file name or number
                MensajeError = "Número o Nombre de Archivo Incorrecto"
            Case 53 ' File not found
                MensajeError = "Archivo No Encontrado"
            Case 54 ' Bad file mode
                MensajeError = "Modo de Archivo Incorrecto"
            Case 55 ' File already open
                MensajeError = "Archivo ya Abierto"
            Case 57 ' Device I/O error
                MensajeError = "Error I/O de Dispositivo"
            Case 58 ' File already exists
                MensajeError = "Archivo ya Existe"
            Case 59 'Bad record length
                MensajeError = "Longitud de Registro Incorrecta"
            Case 61 ' Disk full
                MensajeError = "Disco Lleno"
            Case 62 ' Input past end of file
                MensajeError = "Fin de Archivo"
            Case 63 ' Bad record number
                MensajeError = "Número de Registro Incorrecto"
            Case 67 ' Too many files
                MensajeError = "Demasiados Archivos"
            Case 68 ' Device unavailable
                MensajeError = "La Unidad Seleccionada No esta Lista"
            Case 70 ' Permission denied
                MensajeError = "Permiso Denegado"
            Case 71 ' Disk not ready
                MensajeError = "El Disco No esta Listo"
            Case 74 ' Can't rename with difference drive
                MensajeError = "Can't rename with difference drive"
            Case 75 ' Path/File access error
                MensajeError = "Error al Accesar Ruta/Archivo"
            Case 76 ' Path not found
                MensajeError = "Ruta No Encontrada"
            Case 91 ' Object Variable or With block not set
                MensajeError = "VARIABLE o Bloque WITH no Definido"
            Case 92 ' For loop not initialized
                MensajeError = "FOR-LOOP No Iniciado"
            Case 93 ' Invalid pattern string
                MensajeError = "Modelo de Cadena Inválido"
            Case 94 ' Invalid use of null
                MensajeError = "Uso Inválido de NULL"
            Case 96 ' Unable to sink events of object because the objects is already fring events to the maximun number of event receivers that it supports
                MensajeError = "Unable to sink events of object because the objects is already fring events to the maximun number of event receivers that it supports"
            Case 97 ' Can not call friend function on object which is not an instance of defining class
                MensajeError = "No Puede Llamar una Funcion sobre un Objeto que no es Instancia de una Clase Definida"
            Case 98 ' A property or method call cannot include a reference to private object, either as an argument or as a return value
                MensajeError = "La Llamada a una Propiedad o Método no puede Incluir una Referencia para Privar un Objeto, o un Argumento o Valor Regresado"
            Case 325 ' Invalid format in resurce file
                MensajeError = "Formato Inválido en Archivo Fuente"
            Case 380 ' Invalid Property Value
                MensajeError = "Tipo de Letra Incorrecto  " & Chr(13) & "Seleccione Correctamente"
            Case 381 ' Invalid Property array index
                MensajeError = "Invalid Property array index"
            Case 382 ' Set not supported at runtime
                MensajeError = "Set not supported at runtime"
            Case 383 ' Set not supported (read-only property)
                MensajeError = "Set not supported (read-only property)"
            Case 385 ' Need property array index
                MensajeError = "Need property array index"
            Case 387 ' Set not permitted
                MensajeError = "Set not permitted"
            Case 422 ' Property not found
                MensajeError = "PROPIEDAD no Encontrada"
            Case 423 ' Property or Method not found
                MensajeError = "PROPIEDAD o METODO no Encontrado"
            Case 424 ' Object required
                MensajeError = "OBJETO Requerido"
            Case 429 ' ActiveX Component can't create object
                MensajeError = "El Componente ActiveX no puede Crear el OBJETO"
            Case 430 ' Class does not support Automation or does not support expected interface
                MensajeError = "Class does not support Automation or does not support expected interface"
            Case 432 ' File name or class name not found during Automation operation
                MensajeError = "File name or class name not found during Automation operation"
            Case 438 ' Object doesn't support this property or method
                MensajeError = "El OBJETO no Soporta esta PROPIEDAD o METODO"
            Case 440 ' Automation error
                MensajeError = "Automation error"
            Case 442 ' Connection to type library or object library for remote process has been lost. Press OK for dialog to remove reference
                MensajeError = "Se ha Perdido la Conexión a la BIBLIOTECA para el Proceso Remoto. Presione OK para Remover la Referencia"
            Case 445 ' Object doesn't support this action
                MensajeError = "El OBJETO no Soporta esta Acción"
            Case 446 ' Object doesn't support named arguments
                MensajeError = "Object doesn't support named arguments"
            Case 447 ' Object doesn't support current locale setting
                MensajeError = "Object doesn't support current locale setting"
            Case 448 ' Named argument not found
                MensajeError = "No se Encontró el ARGUMENTO Nombrado"
            Case 449 ' Argument not optional
                MensajeError = "ARGUMENTO no Opcional"
            Case 450 ' Wrong number of arguments or invalid property assignment
                MensajeError = "Número de ARGUMENTOS Incorrecto o PROPIEDAD Asignada Inválida"
            Case 451 ' Property let procedure not defined and property get procedure did not return and object
                MensajeError = "Property let procedure not defined and property get procedure did not return and object"
            Case 452 ' Invalid Ordinal
                MensajeError = "Invalid Ordinal"
            Case 453 ' Specified DLL function not found
                MensajeError = "No se Encontró la Función DLL Especificada"
            Case 454 ' Code resource not found
                MensajeError = "No se Encontró Código Fuente"
            Case 455 ' Code resource lock error
                MensajeError = "Error de Bloqueo en Código Fuente"
            Case 457 ' This key is already associated with an element of this collection
                MensajeError = "This key is already associated with an element of this collection"
            Case 458 ' Variable uses an Automation type not supported in Visual Basic
                MensajeError = "Variable uses an Automation type not supported in Visual Basic"
            Case 459 ' Object or Class does not support the set of events
                MensajeError = "Object or Class does not support the set of events"
            Case 460 ' Invalid clipboard format
                MensajeError = "Formato de Portapapeles Inválido"
            Case 461 ' Method or data member not found
                MensajeError = "Método o Miembro de Dato no Encontrado"
            Case 462 ' The remote server machine does not exist or is unavailable
                MensajeError = "El Servidor Remoto no Existe o no esta Disponible"
            Case 463 ' Class not registered on local machine
                MensajeError = "CLASE no Registrada en la Máquina Local"
            Case Else
                MensajeError = Err.Number & Chr(13) & Err.Description
        End Select

        '*****  COLECCION DE ERROES DE SQL *****'
        For Each Errorsito In cnn.Errors
                If Errorsito.NativeError Then
                    MensajeError = MensajeError & Chr(13) & "SQL Error No. " & Errorsito.NativeError & Chr(13) & Errorsito.Description & Chr(13) & Errorsito.SQLState & Chr(13) & Errorsito.Source
                End If
            Next Errorsito
            MsgBox(MensajeError, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Error... " & Str(Err.Number))
            Err.Clear()
            cnn.Errors.Clear()

    End Sub
End Module
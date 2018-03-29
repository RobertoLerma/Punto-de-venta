'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE ESTANDAR JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6

'Friend Class MSHierarchicalFlexGridLib
'End Class

Public Module ModEstandar

    Public Const CC_Numeros As String = "0123456789"
    Public Const CC_Letras As String = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ abcdefghijklmnñopqrstuvwxyz"

    Public Const BIF_BROWSEFORCOMPUTER As Integer = &H1000S
    Public Const BIF_BROWSEFORPRINTER As Integer = &H2000S
    Public Const BIF_BROWSEINCLUDEFILES As Integer = &H4000S
    Public Const BIF_BROWSEINCLUDEURLS As Integer = &H80S
    Public Const BIF_DONTGOBELOWDOMAIN As Integer = &H2S
    Public Const BIF_EDITBOX As Integer = &H10S
    Public Const BIF_NEWDIALOGSTYLE As Integer = &H40S
    Public Const BIF_RETURNFSANCESTORS As Integer = &H8S
    Public Const BIF_RETURNONLYFSDIRS As Integer = &H1S
    Public Const BIF_SHAREABLE As Integer = &H8000S
    Public Const BIF_STATUSTEXT As Integer = &H4S
    Public Const BIF_VALIDATE As Integer = &H20S

    Public Sub MoverFramesVirtuales(ByRef Forma As System.Windows.Forms.Form, ByRef CtrlIzq As System.Windows.Forms.Control, ByRef CtrlDer As System.Windows.Forms.Control, ByRef CtrlSeparador As System.Windows.Forms.Control, ByRef CtrlLinea As System.Windows.Forms.Control, Optional ByRef TamMinCtrlIzq As Single = 0, Optional ByRef TamMinCtrlDer As Single = 0, Optional ByRef MargenIzq As Single = 0, Optional ByRef MargenDer As Single = 0)

        On Error GoTo Merr

        Dim LimiteDerCtrlDer, sglPx, LimiteIzqCtrlIzq, MitadSep As Single
        Dim AreaDisponible As Single
        If MargenIzq < 30 Then MargenIzq = 30
        If MargenDer < 30 Then MargenDer = 30
        If TamMinCtrlIzq < 60 Then TamMinCtrlIzq = 60
        If TamMinCtrlDer < 60 Then TamMinCtrlDer = 60

        MitadSep = PixelsToTwipsX(CtrlSeparador.Width) / 2

        LimiteIzqCtrlIzq = IIf(PixelsToTwipsX(CtrlIzq.Left) < MargenIzq, MargenIzq, PixelsToTwipsX(CtrlIzq.Left))
        LimiteDerCtrlDer = IIf((PixelsToTwipsX(CtrlDer.Left) + PixelsToTwipsX(CtrlDer.Width)) > (PixelsToTwipsX(Forma.ClientRectangle.Width) - MargenDer), (PixelsToTwipsX(Forma.ClientRectangle.Width) - MargenDer), (PixelsToTwipsX(CtrlDer.Left) + PixelsToTwipsX(CtrlDer.Width)))

        AreaDisponible = LimiteDerCtrlDer - LimiteIzqCtrlIzq - PixelsToTwipsX(CtrlSeparador.Width) - 60

        If TamMinCtrlIzq > AreaDisponible Then TamMinCtrlIzq = AreaDisponible
        If TamMinCtrlDer > AreaDisponible Then TamMinCtrlDer = AreaDisponible

        sglPx = PixelsToTwipsX(CtrlLinea.Left) + (PixelsToTwipsX(CtrlLinea.Width) / 2)
        Select Case sglPx
            Case Is < (LimiteIzqCtrlIzq + TamMinCtrlIzq + MitadSep)
                sglPx = LimiteIzqCtrlIzq + TamMinCtrlIzq + MitadSep
            Case Is > (LimiteDerCtrlDer - TamMinCtrlDer - MitadSep)
                sglPx = LimiteDerCtrlDer - TamMinCtrlDer - MitadSep
        End Select

        CtrlLinea.Left = TwipsToPixelsX(sglPx)

        CtrlIzq.Left = TwipsToPixelsX(LimiteIzqCtrlIzq)
        CtrlIzq.Width = TwipsToPixelsX(sglPx - LimiteIzqCtrlIzq - MitadSep)
        CtrlSeparador.Left = TwipsToPixelsX(sglPx - MitadSep)
        CtrlDer.Left = TwipsToPixelsX(PixelsToTwipsX(CtrlSeparador.Left) + PixelsToTwipsX(CtrlSeparador.Width))
        CtrlDer.Width = TwipsToPixelsX(LimiteDerCtrlDer - PixelsToTwipsX(CtrlDer.Left))

Merr:

    End Sub



    Public Sub DesActivaMenus()
        'With MenuPrincipal
        '    .mnuArchivoOpc(0).Enabled = False
        '    .mnuArchivoOpc(1).Enabled = False
        '    .mnuArchivoOpc(2).Enabled = False
        '    .mnuArchivoOpc(3).Enabled = True
        '    .mnuEdicionOpc(0).Enabled = False
        '    .mnuEdicionOpc(1).Enabled = False
        '    .mnuEdicionOpc(2).Enabled = False
        '    .mnuEdicionOpc(3).Enabled = False

        '    .ToolbarStandar.Items.Item(1).Enabled = False
        '    .ToolbarStandar.Items.Item(2).Enabled = False
        '    .ToolbarStandar.Items.Item(4).Enabled = False
        '    .ToolbarStandar.Items.Item(5).Enabled = False
        '    .ToolbarStandar.Items.Item(6).Enabled = False
        '    .ToolbarStandar.Items.Item(7).Enabled = False
        'End With
    End Sub

    ''Configura el formulario y el grid de las
    ''consultas (ayudas) automaticamente
    'Public Sub ConfiguraConsultas(frm As FrmConsultas, nAncho As Integer, _
    ''                                Rs As ADODB.Recordset, cTag As String, _
    ''                                cCaption As String)
    '    Dim LnAltoGrid As Integer
    '    Dim nContador As Integer
    '
    '    Load frm
    '    ModEstandar.CentrarForma FrmConsultas
    '
    '    Rs.MoveLast
    '    Rs.MoveFirst
    '    'Determina el alto del formulario
    '    If Rs.RecordCount > 10 Then
    '        LnAltoGrid = 3000 + 400 ' tamaño maximo para 10 registros se obtubo de: (250*10)+300
    '        nAncho = nAncho + 350
    '    Else '            No de registros*Alto de la celda + alto del cabecero
    '        LnAltoGrid = (300 * RsGral.RecordCount) + 400
    '        nAncho = nAncho + 100
    '    End If
    '
    '    With frm
    '        .Caption = cCaption
    '        .Tag = cTag
    '        .Width = nAncho + 350
    '        .Height = LnAltoGrid + 580
    '        With .Flexdet
    '            .ClearStructure
    '            .Width = nAncho
    '            .Height = LnAltoGrid
    '            .AllowUserResizing = flexResizeBoth
    '            .RowSizingMode = flexRowSizeIndividual
    '            .SelectionMode = flexSelectionByRow
    '            .FocusRect = flexFocusNone
    '
    '            .WordWrap = True
    '            .FixedCols = 0
    '            .FixedRows = 1
    '            .ScrollBars = flexScrollBarBoth
    '            'Asigna la consulta al grid del formulario de busquedas
    '            Set .Recordset = Rs
    '            .RowHeight(0) = 350
    '            .Row = 0
    '            For nContador = 0 To (.Cols - 1) Step 1
    '                .Col = nContador
    '                .CellAlignment = flexAlignCenterCenter
    '                .ColAlignment = flexAlignGeneral
    '            Next nContador
    '            '.Col = 0
    '            .Row = 1
    '        End With
    '    End With
    '
    'End Sub

    Sub BorraCmd()
        Dim I As Integer
        If Cmd.Parameters.Count > 0 Then
            For I = Cmd.Parameters.Count - 1 To 0 Step -1
                Cmd.Parameters.Delete(I)
            Next
        End If
    End Sub

    Public Sub DesctivaMenus()
        'With MenuPrincipal
        '    ''MENU PRINCIPAL
        '    .mnuArchivoOpc(0).Enabled = False
        '    .mnuArchivoOpc(1).Enabled = False
        '    .mnuArchivoOpc(2).Enabled = False
        '    .mnuArchivoOpc(3).Enabled = True
        '    .mnuArchivoOpc(4).Enabled = True
        '    .mnuEdicionOpc(0).Enabled = False
        '    .mnuEdicionOpc(1).Enabled = False
        '    .mnuEdicionOpc(2).Enabled = False
        '    .mnuEdicionOpc(3).Enabled = False

        '    .ToolbarStandar.Items.Item(1).Enabled = False
        '    .ToolbarStandar.Items.Item(2).Enabled = False
        '    .ToolbarStandar.Items.Item(4).Enabled = False
        '    .ToolbarStandar.Items.Item(5).Enabled = False
        '    .ToolbarStandar.Items.Item(6).Enabled = False
        '    .ToolbarStandar.Items.Item(7).Enabled = False
        '    ''MENU DE CONTEXTO
        '    .menuContextualGenOpc(1).Enabled = False
        '    .menuContextualGenOpc(2).Enabled = False
        '    .menuContextualGenOpc(3).Enabled = True
        '    .menuContextualGenOpc(4).Enabled = True
        '    .menuContextualGenOpc(7).Enabled = False
        '    .menuContextualGenOpc(8).Enabled = False
        '    .menuContextualGenOpc(9).Enabled = False
        '    .menuContextualGenOpc(10).Enabled = False
        'End With
    End Sub

    ' Convierte todos los caracteres a Mayúsculas
    Public Function gp_CampoMayusculas(ByVal LS_keyascii As Integer) As Object
        gp_CampoMayusculas = Asc(UCase(Chr(LS_keyascii)))
    End Function

    ' Valida que solo sean aceptados NUMEROS y algunos caracteres
    ' especificados en la variable string lv_Caracteres_Adicionales
    Public Sub gp_CampoNumerico(ByRef Li_keyascii As Integer, Optional ByRef Lv_Caracteres_Adicionales As Object = Nothing)
        Dim Ls_Temp As String
        If IsNothing(Lv_Caracteres_Adicionales) Then
            Ls_Temp = ""
        Else
            Ls_Temp = Lv_Caracteres_Adicionales
        End If

        If InStr(1, CC_Numeros & Ls_Temp & Chr(8), Chr(Li_keyascii)) = 0 Then
            If Li_keyascii <> 13 Then
                Li_keyascii = 0
            End If
        End If
    End Sub

    ' Valida que solo sean aceptados LETRAS y algunos caracteres
    ' especificados en la variable string lv_Caracteres_Adicionales

    Public Sub gp_CampoLetras(ByRef Li_keyascii As Integer, Optional ByRef Lv_Caracteres_Adicionales As Object = Nothing)
        Dim Ls_Temp As String
        If IsNothing(Lv_Caracteres_Adicionales) Then
            Ls_Temp = ""
        Else
            Ls_Temp = Lv_Caracteres_Adicionales
        End If

        If InStr(1, CC_Letras & Ls_Temp & Chr(8), Chr(Li_keyascii)) = 0 Then
            If Li_keyascii <> 13 Then
                Li_keyascii = 0
            End If
        End If
    End Sub

    ' Valida que solo sean aceptados LETRAS Y NUMEROS
    ' y algunos caracteres especificados en la variable
    ' string lv_Caracteres_Adicionales

    Public Sub gp_CampoAlfanumerico(ByRef Li_keyascii As Integer, Optional ByRef Lv_Caracteres_Adicionales As Object = Nothing)
        Dim Ls_Temp As String
        If IsNothing(Lv_Caracteres_Adicionales) Then
            Ls_Temp = ""
        Else
            Ls_Temp = Lv_Caracteres_Adicionales
        End If

        If InStr(1, CC_Numeros & CC_Letras & Ls_Temp & Chr(8), Chr(Li_keyascii)) = 0 Then
            If Li_keyascii <> 13 Then
                Li_keyascii = 0
            End If
        End If
    End Sub
    'Función para rellenar de ceros una variable
    Public Function gf_cerosVar(ByRef lv_valor As Object, ByRef li_Ancho As Integer) As String
        gf_cerosVar = Right(New String("0", li_Ancho) & Trim(lv_valor), li_Ancho)
    End Function
    'Función para rellenar de ceros un objeto txt
    Public Function gf_cerosTxt(ByRef lv_valor As Object, ByRef obj_Texto As System.Windows.Forms.Control) As String
        gf_cerosTxt = Right(New String("0", Convert.ToInt32(obj_Texto.MaximumSize)) & Trim(lv_valor), Convert.ToInt32(obj_Texto.MaximumSize))
    End Function
    'Funcion para convertir el formato de la fecha
    'a mes/Día/Año
    'Regularmente se utiliza para grabar inf. en la bd.
    Function FechaBd(ByRef Fecha As Date) As String
        Dim lnDia As Integer
        Dim lnMes As Integer

        lnDia = VB.Day(Fecha)
        lnMes = Month(Fecha)
        FechaBd = lnMes & "/" & lnDia & "/" & Year(Fecha)
    End Function
    ' Esta función es para simular una máscara para valores numericos
    Function MskCantidad(ByRef Valor As String, ByRef tecla As Object, ByRef Enteros As Integer, ByRef Decimales As Integer, ByRef Cursor As Integer) As Object

        Dim I As Integer
        Dim Punto As Integer
        Dim Numeros As Integer

        If ((tecla < 48) Or (tecla > 58)) And (tecla <> 8) And (tecla <> Asc(".")) Then
            tecla = 0
        End If
        If tecla = Asc(".") And Valor = "" Then
            Valor = "0"
            Cursor = Cursor + 1
        End If

        'Por si quieren borrar el PUNTO con la tecla back-space
        If tecla = 8 Then
            If Cursor < 1 Then
                MskCantidad = 0
                Exit Function
            Else
                If Mid(Valor, Cursor, 1) = "." Then
                    If Cursor = 1 Or Cursor = Len(Valor) Then
                        MskCantidad = tecla
                        Exit Function
                    Else
                        MskCantidad = 0
                        Exit Function
                    End If
                Else
                    MskCantidad = tecla
                    Exit Function
                End If
            End If
        End If
        If tecla = Asc(".") And Valor = "" Then
            MskCantidad = 0
            Exit Function
        End If
        If Valor <> "" Then
            If tecla = Asc(".") And Len(Mid(Valor, Cursor + 1, Len(Valor) - Cursor + 1)) > Decimales Then
                MskCantidad = 0
                Exit Function
            End If
        End If
        If tecla = Asc(".") And Decimales = 0 Then
            MskCantidad = 0
            Exit Function
        End If
        Punto = 0
        For I = 1 To Len(Valor)
            If Mid(Valor, I, 1) = "." Then
                Punto = I
                I = Len(Valor)
                If tecla = Asc(".") Then
                    MskCantidad = 0
                    Exit Function
                End If
            End If
        Next I
        If Punto = 0 And tecla <> Asc(".") Then
            If Len(Valor) = Enteros Then
                MskCantidad = 0
                Exit Function
            End If
        ElseIf Punto <> 0 And tecla <> Asc(".") Then
            If Cursor < Punto Then
                Numeros = 0
                For I = 1 To Punto
                    If Mid(Valor, I, 1) <> "." Then
                        Numeros = Numeros + 1
                    End If
                    If Numeros >= Enteros Then
                        MskCantidad = 0
                        Exit Function
                    End If
                Next I
            Else
                Numeros = 0
                For I = Punto To Len(Valor)
                    If Mid(Valor, I, 1) <> "." Then
                        Numeros = Numeros + 1
                    End If
                    If Numeros >= Decimales Then
                        MskCantidad = 0
                        Exit Function
                    End If
                Next I
            End If
        End If
        MskCantidad = tecla
    End Function

    Sub MSHFlexGridEdit(ByRef MSHFlexGrid As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid, ByRef Edt As System.Windows.Forms.Control, ByRef KeyAscii As Integer)
        Dim cRenAnt As Integer
        ' Usar el carácter escrito.
        Select Case KeyAscii

            ' Un espacio significa modificar el texto actual.
            Case 0 To 32
                Edt.Text = MSHFlexGrid.Text
                cRenAnt = MSHFlexGrid.Row
                ' Otro carácter reemplaza el texto actual.
            Case Else
                Edt.Text = Chr(KeyAscii)
                '
                'Edt. = 1 
        End Select

        ' Mostrar Edt en la posición correcta.
        Edt.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(MSHFlexGrid.Left) - 25 + MSHFlexGrid.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(MSHFlexGrid.Top) - 25 + MSHFlexGrid.CellTop), VB6.TwipsToPixelsX(MSHFlexGrid.CellWidth + 15), 0, BoundsSpecified.X Or BoundsSpecified.Y Or BoundsSpecified.Width)

        Edt.Visible = True
        'Y hacer que funcione.
        Edt.Focus()
    End Sub

    ' Fucion Salir que recibe como parametro opcional el mensaje
    ' si no se le pone nada al mensaje entoces por default
    ' toma "¿Desea Abandonar captura?"
    Function Salir(Optional ByRef strTexto As String = "", Optional ByRef Aviso As String = "") As Boolean

        If Trim(strTexto) = "" Then strTexto = "¿Desea Abandonar captura?"

        If MsgBox(strTexto, MsgBoxStyle.YesNo + MsgBoxStyle.Question, Aviso) = MsgBoxResult.Yes Then
            Salir = True
        Else
            Salir = False
        End If
    End Function

    'Esta funcion muestra seleccionado un control Text
    'Parámetros
    ' 1.- Objeto txt
    ' 2.- Posición inicial de la selección  (opcional)
    ' 3.- Número de caracteres que desea que se seleccione  (opcional)
    Public Sub SelTextoTxt(ByRef txtCtrl As System.Windows.Forms.TextBox, Optional ByRef Inicio As Integer = 0, Optional ByRef Longitud As Integer = 0)
        If Inicio <= 1 Then
            Inicio = 0
        Else
            Inicio = Inicio - 1
        End If

        If Longitud <= 0 Then Longitud = Len(txtCtrl.Text)

        With txtCtrl
            .SelectionStart = Inicio '0
            .SelectionLength = Longitud 'Len(Trim(.Text))
        End With
    End Sub
    'Procedimiento para desplegar que un dato no existe
    'Parámetros
    '  1.- Dato que no existe
    '  2.- Caption de la ventana  (OPCIONAL - Si no se especifica pone el nombre de la aplicación)
    Public Sub MsjNoExiste(ByRef Dato As Object, Optional ByRef Aviso As Object = Nothing)
        If IsNothing(Aviso) = True Then
            Aviso = My.Application.Info.AssemblyName
        End If

        MsgBox(Dato & " no existe" & vbNewLine & "Verifique por favor...", MsgBoxStyle.Exclamation, Aviso)
    End Sub

    'Función para mostrar Mensaje de error
    'Parámetros
    '  1.- Mensaje de error que se generó(Opcional)
    '  2.- Icono que mostrará la ventana de error (Opcional)
    '  3.- Aviso : Caption de la ventana de error
    Public Sub MostrarError(Optional ByRef Mensaje As String = "", Optional ByRef Icono As MsgBoxStyle = MsgBoxStyle.Critical, Optional ByRef Conexion As ADODB.Connection = Nothing)
        Dim strMsg, strArchivo As String
        Dim SQLErr As ADODB.Error
        Dim blnErroresSQL As Boolean
        Dim fsoSA As New Scripting.FileSystemObject

        'Función que muestra el error que se generó ~Å~
        If Mensaje <> "" Then
            strMsg = Mensaje & vbNewLine & vbNewLine & "Error: " & Err.Number & vbNewLine & Err.Description
        Else
            strMsg = "Error: " & Err.Number & vbNewLine & Err.Description
        End If
        If Icono <> MsgBoxStyle.Critical And Icono <> MsgBoxStyle.Exclamation And Icono <> MsgBoxStyle.Information Then
            Icono = MsgBoxStyle.Critical
        End If

        On Error Resume Next
        If Conexion Is Nothing Then Conexion = ModVariables.Cnn
        For Each SQLErr In Conexion.Errors
            If Not blnErroresSQL Then
                strMsg = strMsg & vbNewLine & "Errores SQL: " & SQLErr.NativeError
                blnErroresSQL = True
            Else
                strMsg = strMsg & ", "
                strMsg = strMsg & SQLErr.NativeError
            End If
        Next SQLErr

        MsgBox(strMsg, Icono, ModVariables.gstrNombCortoEmpresa)

        strArchivo = fsoSA.BuildPath(My.Application.Info.DirectoryPath, My.Application.Info.AssemblyName & ".log")

        'LogMsj(System.Diagnostics.TraceEventType.Error, vbLogToFile, strMsg, strArchivo)

        'Limpia los errores
        Conexion.Errors.Clear()
        Err.Clear()
    End Sub


    'Función que determina si está la impresora conectada y funcionando
    'Parámetros
    '  1.- RutaNombreImpresora Nombre de la impresora a la que desean imprimir
    '  2.- AvisoError (caption del mensaje de error en caso de que no encuentre la impresora)
    Public Function BuscarImpresora(Optional ByRef RutaNombreImpresora As String = "", Optional ByRef AvisoError As String = "") As Boolean
        Dim intMA As Integer
        Dim Impresora As Printer
        Dim I As Byte
        Dim sglTiempoInicio As Single

        sglTiempoInicio = VB.Timer() - 3

        On Error GoTo Merr

        BuscarImpresora = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        'Si no se envia como parametro una impresora se toma la predeterminada del sistema
        If IsNothing(RutaNombreImpresora) Or Trim(RutaNombreImpresora) = "" Then
            If Printers.Count > 0 Then 'Si hay impresoras instaladas en el sistema
                RutaNombreImpresora = Impresora.DeviceName


            Else
                'Si es que no se encuentran impresoras instaladas en el sistema se pregunta si se desea dar entrada a una
                RutaNombreImpresora = InputBox("Proporcione una impresora:" & vbNewLine & "p.e. \\ServidorOMaquina\NombreDeLaImpresora", "No hay impresoras instaladas")
                If Trim(RutaNombreImpresora) = "" Then
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Function
                End If
            End If
        End If

        intMA = FreeFile()
        'abre el puerto de impresión. Si ocurre algún error lo manda a la etiqueta MErr
        FileOpen(intMA, RutaNombreImpresora, OpenMode.Output)
        FileClose(intMA)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        BuscarImpresora = True
        Exit Function

Merr:
        'Si no se encuentra activa manda el sig. mensaje
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        If MsgBox("No se puede accesar a la impresora:" & vbNewLine & RutaNombreImpresora & vbNewLine & vbNewLine & "Verifique que la impresora este conectada y funcionando.", MsgBoxStyle.RetryCancel + MsgBoxStyle.Critical) = MsgBoxResult.Retry Then

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        End If

        Err.Clear()
    End Function

    'Parámetros:
    '   1.- Aviso: (Opcional)Caption que aparecerá en la ventana de error
    Public Function CtrlErrImpresion(Optional ByRef Aviso As String = "") As MsgBoxResult
        'Los valores que regresa la función son: _
        'vbRetry  -> Cuando se quiere reintentar la operación de impresión _
        'vbCancel -> Cuando se canceló la operación por el usuario _
        'vbAbort  -> Cuando es un error irrecuperable de la impresora _
        'vbNo     -> Cuando NO es un error de la impresora...
        Dim bytR As Byte

        Select Case Err.Number
            Case 389 'Intento asignar el valor diferente de una propiedad dentro de una página
                MsgBox("A la propiedad le fué asignado un valor diferente dentro de la página" & vbNewLine & "Avise a su administrador de sistema.", MsgBoxStyle.Critical, "Error en la impresión")
                bytR = MsgBoxResult.Abort
            Case 482
                bytR = MsgBox("Ha ocurrido un error en la impresión" & vbNewLine & "Verifique que su impresora funcione correctamente", MsgBoxStyle.RetryCancel + MsgBoxStyle.Exclamation, Aviso)
            Case 483 'El driver de la impresora no soporta una propiedad
                MsgBox("El controlador de la impresora no soporta la propiedad" & vbNewLine & "Avise a su administrador de sistema.", MsgBoxStyle.Critical, "Error en la impresión")
                bytR = MsgBoxResult.Abort
            Case 484 'No hay información suficiente de la impresora en el WIN.INI
                MsgBox("No hay suficiente información sobre la impresora" & vbNewLine & "Talvez necesitará volver a instalar el controlador de la impresora.", MsgBoxStyle.Critical, "Error de la impresora")
                bytR = MsgBoxResult.Abort
            Case Else
                MostrarError("Ocurrió el siguiente error.")
                bytR = MsgBoxResult.No
        End Select
        'Limpia el Error.
        Err.Clear()

        CtrlErrImpresion = bytR
    End Function

    'Función para convertir un expresión string a valor
    'Considerando que puede tener un formato numérico
    Function Numerico(Numero As String) As String
        Numerico = "0.00"
        Numero = Trim(Numero)

        If Trim(Numero) = "" Then
            'Numerico = "0.00"
            Return Numerico
            Exit Function
        End If

        If IsNumeric(Numero) Then
            Numerico = (CDec(Numero.ToString()))
            Return Numerico
            Exit Function

        End If

        Numero = "0" & Numero

        If IsNumeric(Numero) Then
            Numerico = (CDec(Numero.ToString()))
            Return Numerico
            Exit Function
        End If

    End Function

    'Función para saber el número de DÍAS DE UN MES
    'Parámetros
    '   1.- fecha: Fecha de la que se calcularán los días
    Function DiasMes(ByRef Fecha As Date) As Byte

        Dim FecIni As String
        Dim FecFin As String

        'Toma el día inicial del mes
        FecIni = "01/" & Right("00" & CStr(Month(Fecha)), 2) & "/" & Right("0000" & CStr(Year(Fecha)), 4)

        'Si es el mes de diciembre
        'FechFin = al mes de enero
        If Month(Fecha) = 12 Then
            FecFin = "01/01" & "/" & Right("0000" & CStr(Year(Fecha) + 1), 4)
        Else
            FecFin = "01/" & Right("00" & CStr(Month(Fecha) + 1), 2) & "/" & Right("0000" & CStr(Year(Fecha)), 4)
        End If

        DiasMes = DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FecIni), CDate(FecFin))
    End Function

    'Función para encriptar un string
    ' Parámetros
    '     1.- Text : Dato a encriptar
    '     2.- Llave: Llave que se utiliza para encriptar el dato
    Public Function Encriptar(ByRef text As String, ByRef llave As Double) As String
        On Error GoTo Errores
        Dim I As Integer

        Encriptar = ""

        If Len(Trim(text)) > 0 Then
            'Ciclo para cada caracter del texto
            For I = 1 To Len(text)
                'Al asc del caracter se le suma la llave
                'Cada caracter se separa con ";"
                Encriptar = Encriptar & CStr((Asc(Mid(text, I, 1))) + llave) & ";"
            Next I

            'Quita ";"
            Encriptar = Left(Encriptar, Len(Encriptar) - 1)
        Else
            'Si es cadena vacía
            'toma el valor de la llave
            Encriptar = CStr(llave)
        End If
Errores:
        If Err.Number <> 0 Then
            MsgBox(Err.Description, MsgBoxStyle.Critical)
            Err.Clear()
        End If
    End Function
    'Función para DesEncriptar un string encriptado
    ' Parámetros
    '     1.- Text : Dato  encriptado
    '     2.- Llave: Llave que se utiliza para desencriptar el dato
    Public Function Desencriptar(ByRef text As String, ByRef llave As Double) As String
        On Error GoTo Errores
        Dim Desenc As String
        Dim I As Integer

        Desenc = ""
        Desencriptar = ""

        If text = Trim(Str(llave)) Then
            Desencriptar = ""
            Exit Function
        End If

        If Len(Trim(text)) > 0 Then
            For I = 1 To Len(text)
                If Mid(text, I, 1) <> ";" Then
                    Desenc = Desenc & CStr(Mid(text, I, 1))
                Else
                    Desencriptar = Desencriptar & Chr(CDbl(Desenc) - llave)
                    Desenc = ""
                End If
            Next I
            Desencriptar = Desencriptar & Chr(CDbl(Desenc) - llave)
        End If

Errores:
        If Err.Number <> 0 Then
            MsgBox(Err.Description, MsgBoxStyle.Critical)
            Err.Clear()
        End If
    End Function

    'Restaura la Forma MDI de su estado minimizado a su estado normal y al máximizado como opcional
    Public Sub RestaurarMDI(ByRef FormaMDI As System.Windows.Forms.Form, Optional ByRef Maximizar As Boolean = False)

        If Maximizar Then FormaMDI.WindowState = System.Windows.Forms.FormWindowState.Maximized

        If FormaMDI.WindowState = System.Windows.Forms.FormWindowState.Minimized Then FormaMDI.WindowState = System.Windows.Forms.FormWindowState.Normal

    End Sub

    'Restaura la forma de su estado minimizado o maximizado y la pone normal
    'Además de que tiene como opcional el centrarla en la pantalla o en la forma MDI
    Public Sub RestaurarForma(ByRef Forma As System.Windows.Forms.Form, Optional ByRef Centrar As Boolean = False, Optional ByRef FormaMDI As System.Windows.Forms.Form = Nothing, Optional ByRef MaximizarMDI As Boolean = False)

        If Not (FormaMDI Is Nothing) Then RestaurarMDI(FormaMDI, MaximizarMDI)
        Forma.BringToFront()
        Forma.WindowState = System.Windows.Forms.FormWindowState.Normal
        If Centrar Then
            If FormaMDI Is Nothing Then
                CentrarForma(Forma)
            Else
                CentrarForma(Forma, FormaMDI)
            End If
        End If
    End Sub

    'Procedimiento para Centrar una forma, sea hija o no de un MDI
    '   Parámetros:
    '   1.- Forma: Nombre la forma que quiere centrarse
    '   2.- FormaMDI: (Opcional) Nombre del mdi en caso de que sea hija

    Public Sub CentrarForma(ByRef Forma As System.Windows.Forms.Form, Optional ByRef FormaMDI As System.Windows.Forms.Form = Nothing)
        Dim objForma As System.Windows.Forms.Form
        Dim sglAncho, sglAlto As Double
        Dim sglx, sgly As Double

        If FormaMDI Is Nothing Then
            sglAlto = PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height)
            sglAncho = PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width)
        Else
            sglAlto = PixelsToTwipsY(FormaMDI.ClientRectangle.Height)
            sglAncho = PixelsToTwipsX(FormaMDI.ClientRectangle.Width)
        End If

        sglx = ((sglAncho - PixelsToTwipsX(Forma.Width)) / 2)
        sgly = ((sglAlto - PixelsToTwipsY(Forma.Height)) / IIf(PixelsToTwipsY(Forma.Height) < 2500, 3, 4))
        Forma.Left = TwipsToPixelsX(IIf(sglx < 0, 0, sglx))
        Forma.Top = TwipsToPixelsY(IIf(sgly < 0, 0, sgly))
    End Sub

    '''Procedimiento para estar alojado dentro del evento Resize de una FormaMDI
    '   Su funcionamiento consiste en mover la forma activa del MDI al centro si es que puede
    '   ser visible en su totalidad; sino es así pondrá su esquina superior izquierda visible
    Public Sub OrganizarFormasEnMDI(ByRef FormaMDI As System.Windows.Forms.Form)
        If FormaMDI.WindowState <> System.Windows.Forms.FormWindowState.Minimized And My.Application.OpenForms.Count > 1 Then
            If FormaMDI.ActiveMdiChild Is Nothing Then Exit Sub
            If Not FormaMDI.ActiveMdiChild.MdiParent Is Nothing = True Then
                If FormaMDI.ActiveMdiChild.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
                    CentrarForma((FormaMDI.ActiveMdiChild), FormaMDI)
                End If
                FormaMDI.LayoutMdi(System.Windows.Forms.MdiLayout.ArrangeIcons)
            End If
        End If
    End Sub
    Public Sub Atributos_RecordSet(ByRef Recor As ADODB.Recordset)
        On Error GoTo Errores
        With Recor
            .let_ActiveConnection(ModVariables.Cnn)
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            .LockType = ADODB.LockTypeEnum.adLockReadOnly
        End With
Errores:
        If Err.Number <> 0 Then ModErrores.Errores()
    End Sub

    'Función para regresa la fecha en un determinado formato
    'Parámetros
    '   1.- Fecha: Fecha que se desea formatear
    '   2.- Formato: Formato en que se presentará. Puede obtener los siguientes valores
    '       1 =  Jueves, 08 de Agosto de 2002.
    '       2 =  08 de Agosto de 2002
    '       3 =  08/Ago/02
    Public Function FormatoFecha(ByRef Fecha As Date, Optional ByRef Formato As Byte = 1) As String
        Dim strDia As String
        Dim strMes As String

        'Si el parámetros formato no está dentro del rango
        'Toma el valor por default de 1
        If Formato <> 1 And Formato <> 3 And Formato <> 2 Then
            Formato = 1
        End If

        If Formato = 1 Then
            FormatoFecha = DiaLetra(Fecha) & ", " & Format(VB.Day(Fecha), "00") & " de " & MesLetra(Fecha, False) & " de " & Year(Fecha)
        ElseIf Formato = 2 Then
            FormatoFecha = Format(VB.Day(Fecha), "00") & " de " & MesLetra(Fecha, False) & " de " & Year(Fecha)
        ElseIf Formato = 3 Then
            FormatoFecha = Format(VB.Day(Fecha), "00") & "/" & MesLetra(Fecha, True) & "/" & Year(Fecha)
        End If

    End Function
    'Función para desplegar el día con letra
    'Parámetros
    '   1.- Fecha: Fecha que desea transformar
    Public Function DiaLetra(ByRef Fecha As Date) As String
        Select Case Weekday(Fecha)
            Case FirstDayOfWeek.Monday
                DiaLetra = "Lunes"
            Case FirstDayOfWeek.Tuesday
                DiaLetra = "Martes"
            Case FirstDayOfWeek.Wednesday
                DiaLetra = "Miércoles"
            Case FirstDayOfWeek.Thursday
                DiaLetra = "Jueves"
            Case FirstDayOfWeek.Friday
                DiaLetra = "Viernes"
            Case FirstDayOfWeek.Saturday
                DiaLetra = "Sábado"
            Case FirstDayOfWeek.Sunday
                DiaLetra = "Domingo"
        End Select
    End Function
    'Función para desplegar el Mes con letra
    'Parámetros
    '   1.- Fecha: Fecha que desea transformar
    '   2.- FormatoCorto: Si desea o no un formato corto
    Public Function MesLetra(ByRef Fecha As Date, Optional ByRef FormatoCorto As Boolean = True) As String

        Select Case Month(Fecha)
            Case 1
                MesLetra = IIf(FormatoCorto, "Ene", "Enero")
            Case 2
                MesLetra = IIf(FormatoCorto, "Feb", "Febrero")
            Case 3
                MesLetra = IIf(FormatoCorto, "Mzo", "Marzo")
            Case 4
                MesLetra = IIf(FormatoCorto, "Abr", "Abril")
            Case 5
                MesLetra = IIf(FormatoCorto, "May", "Mayo")
            Case 6
                MesLetra = IIf(FormatoCorto, "Jun", "Junio")
            Case 7
                MesLetra = IIf(FormatoCorto, "Jul", "Julio")
            Case 8
                MesLetra = IIf(FormatoCorto, "Ago", "Agosto")
            Case 9
                MesLetra = IIf(FormatoCorto, "Sep", "Septiembre")
            Case 10
                MesLetra = IIf(FormatoCorto, "Oct", "Octubre")
            Case 11
                MesLetra = IIf(FormatoCorto, "Nov", "Noviembre")
            Case 12
                MesLetra = IIf(FormatoCorto, "Dic", "Diciembre")
        End Select

    End Function

    'Función para Quitar el Enter de un texto
    'Parámetros:
    '   1.- Texto : String al que se le quitarán los enters
    Function QuitaEnter(ByRef Txt As String) As String
        Dim cadena As Integer

        cadena = InStr(1, Txt, Chr(13) & Chr(10), CompareMethod.Text)
        If Len(Trim(Txt)) > 0 Then
            While cadena > 0
                Txt = Mid(Txt, 1, cadena - 1) & " " & Mid(Txt, cadena + 2, Len(Trim(Txt)))
                cadena = InStr(1, Txt, Chr(13) & Chr(10), CompareMethod.Text)
            End While
        End If

        QuitaEnter = Txt
    End Function

    'Función para desplegar un dato con formato RFC
    'Parámetros:
    '   1.- Texto : String que se formateará
    Function Mask_RFC(ByRef Texto As String) As String
        Select Case Len(Trim(Texto))
            Case 13 'AAAA######???
                Mask_RFC = Mid(Texto, 1, 4) & "-" & Mid(Texto, 5, 6) & "-" & Mid(Texto, 11, 3)
            Case 12 'AAA######???
                Mask_RFC = Mid(Texto, 1, 3) & "-" & Mid(Texto, 4, 6) & "-" & Mid(Texto, 10, 3)
            Case 10 'AAAA######
                Mask_RFC = Mid(Texto, 1, 4) & "-" & Mid(Texto, 5, 6)
            Case 9 'AAA######
                Mask_RFC = Mid(Texto, 1, 3) & "-" & Mid(Texto, 4, 6)
        End Select
        'Texto = Mask_RFC
    End Function

    'Función para quitar caracteres que no sean numeros
    'Parámetros
    '   1.- Texto: string a transformar
    Function Quita_Letra(ByRef Texto As String) As String
        Dim Lni As Integer

        Quita_Letra = ""

        For Lni = 1 To Len(Texto)
            If Not (Asc(Mid(Texto, Lni, 1)) < System.Windows.Forms.Keys.D0 Or Asc(Mid(Texto, Lni, 1)) > System.Windows.Forms.Keys.D9) Then
                Quita_Letra = Quita_Letra & Mid(Texto, Lni, 1)
            End If
        Next

    End Function
    'Función para quitar números en un string
    'Parámetros
    '   1.- Texto: string a transformar
    Function Quita_Num(ByRef Texto As String) As String
        Dim Lni As Integer

        Quita_Num = ""

        For Lni = 1 To Len(Texto)
            If Not (Asc(Mid(Texto, Lni, 1)) >= System.Windows.Forms.Keys.D0 And Asc(Mid(Texto, Lni, 1)) <= System.Windows.Forms.Keys.D9) Then
                Quita_Num = Quita_Num & Mid(Texto, Lni, 1)
            End If
        Next

    End Function
    'Función para quitar un o unos determinados caracteres
    'Parámetros
    '   1.- Texto: string a transformar
    '   2.- Caracteres:  caracteres que se eliminarán
    Function Quita_Caracteres(ByRef Texto As String, ByRef Caracteres As String) As String
        Dim Lni As Integer

        Quita_Caracteres = ""

        For Lni = 1 To Len(Caracteres)
            'Mid(Texto, Lni, 1)
            Texto = Replace(Texto, Mid(Caracteres, Lni, 1), "")
        Next
        Quita_Caracteres = Texto
    End Function

    'Función para validar el RFC Completo
    'valida RFC's con formato :AAA-######-???, AAAA-######-???, AAA-###### ó AAAA-######
    'Parámetro
    '   1.- Cad: Valor string a validar
    Function valida_RFCC(ByRef Cad As String) As Boolean
        valida_RFCC = False

        Dim LnJ As Byte
        Dim LnLong1 As Byte 'longitud de la primera parte del rfc : AAA ó AAAA
        Dim LnLong2 As Byte 'longitud de la segunda parte del rfc : ######
        Dim LnLong3 As Byte 'longitud de la tercera parte del rfc : ??? o nada

        For LnJ = 1 To Len(Cad)
            If LnJ > Len(Trim(Cad)) Then Exit For
            If Asc(Mid(Cad, LnJ, 1)) >= System.Windows.Forms.Keys.A And Asc(Mid(Cad, LnJ, 1)) <= System.Windows.Forms.Keys.Z Then
                LnLong1 = LnLong1 + 1
                If LnLong1 > 4 Then Exit Function
            Else
                Exit For
            End If
        Next

        If Mid(Cad, LnLong1 + 1, 1) <> "-" Or LnLong1 < 3 Then Exit Function

        For LnJ = LnLong1 + 2 To LnLong1 + 7
            If LnJ > Len(Trim(Cad)) Then Exit For
            If Asc(Mid(Cad, LnJ, 1)) >= System.Windows.Forms.Keys.D0 And Asc(Mid(Cad, LnJ, 1)) <= System.Windows.Forms.Keys.D9 Then
                LnLong2 = LnLong2 + 1
                If LnLong2 > 6 Then Exit Function
            Else
                Exit For
            End If
        Next

        Dim lcfec As String
        If LnLong2 < 6 Then
            Exit Function
        Else
            lcfec = Mid(Cad, LnLong1 + 2, 2) & "/" & Mid(Cad, LnLong1 + 4, 2) & "/" & Mid(Cad, LnLong1 + 6, 2)
            Dim fechaRFC As String = Format(lcfec, "dd-MM-yyyy")
            If Not IsDate(fechaRFC) Then
                Exit Function
            End If
        End If

        If Len(Cad) <= LnLong1 + LnLong2 + 2 Then
            valida_RFCC = True
            Return valida_RFCC
            Exit Function
        End If

        For LnJ = LnLong1 + LnLong2 + 3 To Len(Trim(Cad))
            If LnJ > Len(Trim(Cad)) Then Exit For

            If ((Asc(Mid(Cad, LnJ, 1)) >= System.Windows.Forms.Keys.D0 And Asc(Mid(Cad, LnJ, 1)) <= System.Windows.Forms.Keys.D9) Or (Asc(Mid(Cad, LnJ, 1)) >= System.Windows.Forms.Keys.A And Asc(Mid(Cad, LnJ, 1)) <= System.Windows.Forms.Keys.Z)) Then
                LnLong3 = LnLong3 + 1
                If LnLong3 > 3 Then Exit Function
            Else
                Exit For
            End If
        Next

        If LnLong3 = 0 Then
            valida_RFCC = True
            Return valida_RFCC
            Exit Function
        Else
            If LnLong3 <> 3 Then Exit Function
        End If

        valida_RFCC = True
        Return valida_RFCC
    End Function

    'Función para validar la entrada de un RFC
    'permite una mascara [AAA-######-???] ó [AAAA-######-???]
    'Parámetros
    '   1.- RFC     : string que almacena el RFC
    '   2.- Tecla   : Key que se va a validar
    '   3.- Pos     : Posición que va a validar
    Function Valida_RFC(ByRef Rfc As String, ByRef tecla As Integer, ByRef Pos As Byte) As Byte ', Band As Boolean) As Byte

        Dim LnLongi As Byte 'variable para determinar la longitud que debe tener el rfc en base a la posición del primer "-"
        Dim LnPosi As Byte 'variable para determinar la posición del primer "-"

        LnPosi = InStr(1, Rfc, "-")
        LnLongi = LnPosi + 10

        If Len(Rfc) > LnLongi And LnLongi <> 0 Then 'And Band Then
            Valida_RFC = 0
            Exit Function
        End If

        Valida_RFC = tecla
        'primeros 3 o 4 caracteres
        If Pos <= 4 Then
            If tecla < System.Windows.Forms.Keys.A Or tecla > System.Windows.Forms.Keys.Z Then
                If Pos = 4 And tecla = Asc("-") Then
                    'If Band Then
                    If Pos <> 1 Then
                        If Mid(Rfc, Pos - 1, 1) <> "-" And Mid(Rfc, Pos, 1) <> "-" Then
                            Exit Function
                        Else
                            Valida_RFC = 0
                            Exit Function
                        End If
                    End If
                    'Else
                    '    Exit Function
                    'End If
                Else
                    Valida_RFC = 0
                    Exit Function
                End If
            Else
                If Pos <> 1 Then
                    If Mid(Rfc, Pos - 1, 1) = "-" Then
                        Valida_RFC = 0
                        Exit Function
                    End If
                End If
            End If
        Else
            ''posición 5 en adelante(serie de 6 numeros)
            If Pos = 5 And Mid(Rfc, Pos - 1, 1) = "-" Then
                If tecla < System.Windows.Forms.Keys.D0 Or tecla > System.Windows.Forms.Keys.D9 Then
                    Valida_RFC = 0
                    Exit Function
                Else
                    'If Band Then
                    If LnLongi = Len(Rfc) Then
                        Valida_RFC = 0
                        Exit Function
                    Else
                        Exit Function
                    End If
                    'Else
                    '    Exit Function
                    'End If
                End If
            Else
                If Pos = 5 And Mid(Rfc, Pos - 1, 1) <> "-" Then
                    If tecla <> Asc("-") Then
                        Valida_RFC = 0
                        Exit Function
                    End If
                Else
                    If Pos > LnPosi And Pos <= LnPosi + 6 Then
                        If tecla < System.Windows.Forms.Keys.D0 Or tecla > System.Windows.Forms.Keys.D9 Then
                            Valida_RFC = 0
                            Exit Function
                        Else
                            If LnLongi = Len(Rfc) Then 'And Band Then
                                Valida_RFC = 0
                                Exit Function
                            Else
                                If Mid(Rfc, Pos - 1, 1) = "-" And tecla = Asc("-") Then 'And Band Then
                                    Valida_RFC = 0
                                    Exit Function
                                Else
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        ''me posiciono en el segundo "-"
                        If Pos - 1 = LnPosi + 6 Then
                            If tecla <> Asc("-") Then
                                Valida_RFC = 0
                                Exit Function
                            Else
                                If Mid(Rfc, Pos - 1, 1) <> "-" And Mid(Rfc, Pos + 1, 1) <> "-" Then
                                    Exit Function
                                Else
                                    Valida_RFC = 0
                                    Exit Function
                                End If
                            End If
                        Else
                            'valido despues del segundo "-"
                            If Pos > LnLongi Then
                                Valida_RFC = 0
                                Exit Function
                            Else
                                If LnLongi = Len(Rfc) Then 'And Band Then
                                    Valida_RFC = 0
                                    Exit Function
                                Else
                                    If (tecla < System.Windows.Forms.Keys.D0 Or tecla > System.Windows.Forms.Keys.D9) And (tecla < Asc("A") Or tecla > Asc("Z")) Then
                                        Valida_RFC = 0
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Function
    'Función para calcular el dígito verificador módulo 10
    'Parámetro
    '   1.- cClave: String al que se le quiere obtener el dígito Verificador
    Function Digito(ByRef cClave As String) As Object
        Dim LsCad1(10) As Integer

        'Si no tiene valor le asigna 0
        If Trim(cClave) = "" Then Exit Function

        LsCad1(0) = 6
        LsCad1(1) = 3
        LsCad1(2) = 1
        LsCad1(3) = 7
        LsCad1(4) = 4
        LsCad1(5) = 8
        LsCad1(6) = 2
        LsCad1(7) = 1
        LsCad1(8) = 5
        LsCad1(9) = 9

        Digito = (Val(cClave) * CDbl(cClave) / LsCad1(Val(cClave) Mod 10)) Mod 10
    End Function

    'Función para convertir un texto a formato teléfono
    'Parámetros
    '   1.- Tel     : String que se va a formatear
    '   2.- Caracter: (Opcional) Caracter separador del formato de teléfono (-)
    '   3.- DigTel  : Cantidad de dígitos del Teléfono  (7, 9,12,etc)

    Public Function ConvTel(ByVal TEL As String, Optional ByRef Caracter As String = "", Optional ByRef DigTel As Integer = 7) As String
        Dim Tmp As String

        Dim tmp1 As String
        Dim tmp2 As String
        Dim tmp3 As String

        Dim x As Integer
        Dim Y As Integer

        'Dim DigTel As Integer

        If Len(TEL) > DigTel Then
            DigTel = Len(TEL)
        End If

        tmp1 = ""
        tmp2 = ""
        tmp3 = ""
        ConvTel = ""

        If Trim(TEL) = "" Then
            Exit Function
        End If

        TEL = Trim(TEL)
        Y = Len(TEL)

        For x = 1 To DigTel
            Tmp = ""
            If x <= Len(TEL) Then
                Tmp = Mid(TEL, Y, 1)
            Else
                If Caracter <> "" Then
                    Tmp = "_"
                End If
            End If
            Select Case x
                Case Is <= 2
                    tmp3 = Tmp & tmp3
                Case Is <= 4
                    tmp2 = Tmp & tmp2
                Case Is > 4
                    tmp1 = Tmp & tmp1
            End Select
            Y = Y - 1
        Next x

        ConvTel = tmp1 & "-" & tmp2 & "-" & tmp3

    End Function

    'Función para asignar el ToolTipText de un control a un Statusbar
    'Parámetros:
    '   1.- Cont: nombre del control del cual se tomará el tooltiptext
    '   2.- ContStatus:
    Public Sub Tip(ByRef Cont As System.Windows.Forms.Control, ByRef ContStatus As System.Windows.Forms.StatusStrip)
        On Error Resume Next
        'ContStatus.Items.Item(2).Text = Cont.ToolTipText
    End Sub
    'Función para asignar a una forma el icono del MDI
    'Parámetros:
    '   1.- FormaDestino : Nombre de la forma a la que se asignará el icono
    '   2.- FormaOrigen  : Forma de la cual se tomará el ícono
    Function Icono(ByRef FormaDestino As System.Windows.Forms.Form, ByRef FormaOrigen As Object) As Object
        FormaDestino.Icon = FormaOrigen.Icon
    End Function

    'Función para cerrar todas las formas cargadas excepto MDI's
    'Puede especificarsele que una forma no sea cerrada
    'Parámetros:
    '   1.- IgnorarForma: Nombre de la forma que no se cerrará
    Public Function CerrarFormas(Optional ByRef IgnorarForma As System.Windows.Forms.Form = Nothing) As Boolean
        Dim Forma As System.Windows.Forms.Form
        Dim intNF As Integer

        On Error GoTo Merr

        If My.Application.OpenForms.Count > 1 Then
            For Each Forma In My.Application.OpenForms
                'If UCase(Left(Forma.LinkTopic, 3)) <> "MDI" Then
                '    If IgnorarForma Is Nothing Then
                '        intNF = My.Application.OpenForms.Count
                '        Forma.Close()
                '        If intNF = My.Application.OpenForms.Count Then Exit Function
                '    Else
                '        If Forma.Name <> IgnorarForma.Name Then
                '            intNF = My.Application.OpenForms.Count
                '            Forma.Close()
                '            If intNF = My.Application.OpenForms.Count Then Exit Function
                '        End If
                '    End If
                'End If
            Next Forma
        End If

        CerrarFormas = True
        Exit Function
Merr:
    End Function

    'Obtiene un objeto forma (generico) a partir de un nombre,
    'buscando entre las formas cargadas en memoria
    'Parámetros:
    ' · Nombre: Nombre de la forma a obtener
    'Valor que regresa:
    'Si encuentra la forma cargada en memoria regresa el objeto,
    'Sino regresa Nothing
    Public Function ObtenerForma(ByRef Nombre As String) As Object
        On Error Resume Next '~Å~
        Dim varForma As Object
        If Nombre = "" Then Exit Function
        Nombre = UCase(Trim(Nombre))
        For Each varForma In My.Application.OpenForms
            If UCase(varForma.Name) = Nombre Then
                ObtenerForma = varForma
                Exit Function
            End If
        Next varForma
    End Function

    'Configura el formulario y el grid de las
    'consultas (ayudas) automaticamente
    Public Sub ConfiguraConsultas(ByRef frm As FrmConsultas, ByRef nancho As Integer, ByRef Rs As ADODB.Recordset, ByRef cTag As String, ByRef cCaption As String)
        Dim LnAltoGrid As Integer
        Dim nContador As Integer
        frm.InitializeComponent()
        ModEstandar.CentrarForma(frm)
        Rs.MoveLast()
        Rs.MoveFirst()
        'Determina el alto del formulario
        If Rs.RecordCount > 10 Then
            LnAltoGrid = 3000 + 400 ' tamaño maximo para 10 registros se obtubo de: (250*10)+300
            nancho = nancho + 350
        Else '            No de registros*Alto de la celda + alto del cabecero
            LnAltoGrid = (300 * RsGral.RecordCount) + 600
            nancho = nancho + 100
        End If
        With frm
            .Text = cCaption
            .Tag = cTag
            .Width = VB6.TwipsToPixelsX(nancho + 350)
            .Height = VB6.TwipsToPixelsY(LnAltoGrid + 580)
            With .Flexdet
                .ClearStructure()
                .Width = VB6.TwipsToPixelsX(nancho)
                .Height = VB6.TwipsToPixelsY(LnAltoGrid)
                .AllowUserResizing = MSHierarchicalFlexGridLib.AllowUserResizeSettings.flexResizeBoth
                .RowSizingMode = MSHierarchicalFlexGridLib.RowSizingSettings.flexRowSizeIndividual
                .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionByRow
                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                .WordWrap = False
                .FixedCols = 0
                .FixedRows = 1
                .ScrollBars = MSHierarchicalFlexGridLib.ScrollBarsSettings.flexScrollBarBoth
                'Asigna la consulta al grid del formulario de busquedas
                .Recordset = Rs
                .set_RowHeight(0, 350)
                .Row = 0
                For nContador = 0 To (.get_Cols() - 1) Step 1
                    .Col = nContador
                    '''.CellAlignment = flexAlignCenterCenter
                    '''.CellAlignment = flexAlignLeftCenter
                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignGeneral)
                Next nContador
                .Col = 0
                .Row = 1
                '            FrmConsultas.PonerColor 1
            End With
        End With
    End Sub

    'FUNCION QUE ACTIVA LAS OPCIONES DEL MENU Y LOS BOTONES DEL TOOLBAR
    Public Sub ActivaEstandar()
        'MenuPrincipal.mnuArchivoOpc(0).Enabled = True
        'MenuPrincipal.mnuArchivoOpc(2).Enabled = True
        'MenuPrincipal.mnuArchivoOpc(3).Enabled = True
        'MenuPrincipal.mnuArchivoOpc(4).Enabled = True
        'MenuPrincipal.mnuArchivoOpc(6).Enabled = True
        'MenuPrincipal.mnuArchivoOpc(7).Enabled = True
        ''MDIPrincipal.mnuEdicionOpc(4).Enabled = True

        'MenuPrincipal.ToolbarStandar.Items.Item(1).Enabled = True
        'MenuPrincipal.ToolbarStandar.Items.Item(3).Enabled = True
        'MenuPrincipal.ToolbarStandar.Items.Item(4).Enabled = True
        'MenuPrincipal.ToolbarStandar.Items.Item(5).Enabled = True
        'MenuPrincipal.ToolbarStandar.Items.Item(6).Enabled = True
        'MenuPrincipal.ToolbarStandar.Items.Item(8).Enabled = True
        ''MDIPrincipal.ToolbarStandar.Buttons(8).Enabled = True
    End Sub
    Public Sub ActivaDesActivaGuardar(ByRef Valor As Boolean)
        'MenuPrincipal.mnuArchivoOpc(0).Enabled = Valor
        'MenuPrincipal.ToolbarStandar.Items.Item(1).Enabled = Valor
    End Sub

    Public Sub ActivaDesActivaCerrar(ByRef Valor As Boolean)
        'MenuPrincipal.mnuArchivoOpc(2).Enabled = Valor
    End Sub

    Public Sub ActivaDesActivaSalir(ByRef Valor As Boolean)
        'MenuPrincipal.mnuArchivoOpc(4).Enabled = Valor
    End Sub

    Public Sub ActivaDesActivaImprimir(ByRef Valor As Boolean)
        'MenuPrincipal.mnuArchivoOpc(1).Enabled = Valor
        'MenuPrincipal.ToolbarStandar.Items.Item(2).Enabled = Valor
    End Sub

    Public Sub ActivaDesActivaNuevo(ByRef Valor As Boolean)
        'MenuPrincipal.mnuEdicionOpc(0).Enabled = Valor
        'MenuPrincipal.ToolbarStandar.Items.Item(4).Enabled = Valor
    End Sub

    Public Sub ActivaDesActivaCancelar(ByRef Valor As Boolean)
        'MenuPrincipal.mnuEdicionOpc(1).Enabled = Valor
        'MenuPrincipal.ToolbarStandar.Items.Item(5).Enabled = Valor
    End Sub

    Public Sub ActivaDesActivaEliminar(ByRef Valor As Boolean)
        'MenuPrincipal.mnuEdicionOpc(2).Enabled = Valor
        'MenuPrincipal.ToolbarStandar.Items.Item(6).Enabled = Valor
    End Sub

    Public Sub ActivaDesActivaBuscar(ByRef Valor As Boolean)
        'MenuPrincipal.mnuEdicionOpc(3).Enabled = Valor
        'MenuPrincipal.ToolbarStandar.Items.Item(7).Enabled = Valor
    End Sub

    Public Sub DesactivaEstandar()
        'MenuPrincipal.mnuArchivoOpc(0).Enabled = False
        'MenuPrincipal.mnuArchivoOpc(2).Enabled = False
        'MenuPrincipal.mnuArchivoOpc(3).Enabled = False
        'MenuPrincipal.mnuArchivoOpc(4).Enabled = False
        'MenuPrincipal.mnuArchivoOpc(6).Enabled = False
        'MenuPrincipal.mnuArchivoOpc(7).Enabled = False
        ''MDIPrincipal.mnuEdicionOpc(4).Enabled = False

        'MenuPrincipal.ToolbarStandar.Items.Item(1).Enabled = False
        'MenuPrincipal.ToolbarStandar.Items.Item(3).Enabled = False
        'MenuPrincipal.ToolbarStandar.Items.Item(4).Enabled = False
        'MenuPrincipal.ToolbarStandar.Items.Item(5).Enabled = False
        'MenuPrincipal.ToolbarStandar.Items.Item(6).Enabled = False
        'MenuPrincipal.ToolbarStandar.Items.Item(8).Enabled = False
        ''MDIPrincipal.ToolbarStandar.Buttons(8).Enabled = False

    End Sub

    Public Function ActivaMenu(ByRef Nuevo As Integer, ByRef Guardar As Integer, ByRef Cancelar As Integer, ByRef Eliminar As Integer, ByRef Buscar As Integer, ByRef Imprimir As Integer, ByRef Cerrar As Integer) As Object

        Dim Edicion As Boolean

        '0 desactiva
        '1 activa
        '2 ó cualquier otro numero, no hace nada

        Edicion = False

        '''*-*-*-*-*-* NUEVO *-*-*-*-*-*-*-
        'If Nuevo = C_ACTIVADO Then
        '    'MDIMenuPrincipalCorpo.mnuEdicionOpc(0).Enabled = True
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(7).Enabled = True
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(4).Enabled = True
        '    Edicion = True
        'ElseIf Nuevo = C_DESACTIVADO Then
        '    'MDIMenuPrincipalCorpo.mnuEdicionOpc(0).Enabled = False
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(7).Enabled = False
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(4).Enabled = False
        'End If

        ''*-*-*-*-*-* GUARDAR *-*-*-*-*-*-*-
        'If Guardar = C_ACTIVADO Then
        '    MDIMenuPrincipalCorpo.mnuArchivoOpc(0).Enabled = True
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(1).Enabled = True
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(1).Enabled = True
        '    Edicion = True
        'ElseIf Guardar = C_DESACTIVADO Then
        '    MDIMenuPrincipalCorpo.mnuArchivoOpc(0).Enabled = False
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(1).Enabled = False
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(1).Enabled = False
        'End If

        '''*-*-*-*-*-* IMPRIMIR *-*-*-*-*-*-*-
        'If Imprimir = C_ACTIVADO Then
        '    'MDIMenuPrincipalCorpo.mnuArchivoOpc(1).Enabled = True
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(2).Enabled = True
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(2).Enabled = True
        '    Edicion = True
        'ElseIf Imprimir = C_DESACTIVADO Then
        '    'MDIMenuPrincipalCorpo.mnuArchivoOpc(1).Enabled = False
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(2).Enabled = False
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(2).Enabled = False
        'End If

        '''*-*-*-*-*-* CANCELAR *-*-*-*-*-*-*-
        'If Cancelar = C_ACTIVADO Then
        '    'MDIMenuPrincipalCorpo.mnuEdicionOpc(1).Enabled = True
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(8).Enabled = True
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(5).Enabled = True
        '    Edicion = True
        'ElseIf Cancelar = C_DESACTIVADO Then
        '    ' MDIMenuPrincipalCorpo.mnuEdicionOpc(1).Enabled = False
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(8).Enabled = False
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(5).Enabled = False
        'End If

        '''*-*-*-*-*-* ELIMINAR *-*-*-*-*-*-*-
        'If Eliminar = C_ACTIVADO Then
        '    ' MDIMenuPrincipalCorpo.mnuEdicionOpc(2).Enabled = True
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(9).Enabled = True
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(6).Enabled = True
        '    Edicion = True
        'ElseIf Eliminar = C_DESACTIVADO Then
        '    ' MDIMenuPrincipalCorpo.mnuEdicionOpc(2).Enabled = False
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(9).Enabled = False
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(6).Enabled = False
        'End If

        '''*-*-*-*-*-* BUSCAR *-*-*-*-*-*-*-
        'If Buscar = C_ACTIVADO Then
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(7).Enabled = True
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(10).Enabled = True
        '    'MDIMenuPrincipalCorpo.mnuEdicionOpc(4).Enabled = True
        '    Edicion = True
        'ElseIf Buscar = C_DESACTIVADO Then
        '    MDIMenuPrincipalCorpo.ToolbarStandar.Items.Item(7).Enabled = False
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(10).Enabled = False
        '    MDIMenuPrincipalCorpo.mnuEdicionOpc(4).Enabled = False
        'End If

        '''*-*-*-*-*-*CERRAR *-*-*-*-*-*-*-
        'If Cerrar = C_ACTIVADO Then
        '    ' MDIMenuPrincipalCorpo.mnuArchivoOpc(2).Enabled = True
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(3).Enabled = True
        '    Edicion = True
        'ElseIf Cerrar = C_DESACTIVADO Then
        '    ' MDIMenuPrincipalCorpo.mnuArchivoOpc(2).Enabled = False
        '    MDIMenuPrincipalCorpo.menuContextualGenOpc(3).Enabled = False
        'End If

        'MDIMenuPrincipalCorpo.mnuArchivoOpc(3).Enabled = True    'cerrar

        '*-*-*-*-*-* EDICION *-*-*-*-*-*-*-
        'If Edicion Then
        'MenuPrincipal.MnuEdicion(0).Enabled = True
        'End If

    End Function

    'Función para convertir un Número a texto
    'Parámetros:
    '   1.- Numero     : Valor que se va a convertir
    '   2.- Mayúsculas : Si el texto de salida será con mayúsuculas o minúsculas  (OPCIONAL)
    '   3.- Moneda     : Moneda en la que se mostrará el texto  (P = Pesos, D = Dolares)
    Public Function ConLetra(ByVal Numero As Double, Optional ByVal mayusculas As Boolean = True, Optional ByRef Tipo As String = "P") As String
        On Error GoTo Error_Renamed

        Dim NumTmp As String
        Dim c01 As Integer
        Dim c02 As Integer
        Dim Pos As Integer
        Dim dig As Integer
        Dim cen As Integer
        Dim dec As Integer
        Dim uni As Integer
        Dim letra1 As String
        Dim letra2 As String
        Dim letra3 As String
        Dim Leyenda As String
        Dim Leyenda1 As String
        Dim TFNumero As String


        If Numero < 0 Then Numero = System.Math.Abs(Numero)

        NumTmp = Format(Numero, "000000000000000.00") 'Le da un formato fijo
        c01 = 1
        Pos = 1
        TFNumero = ""

        'Para extraer tres digitos cada vez
        Do While c01 <= 5
            c02 = 1
            Do While c02 <= 3

                'Extrae un digito cada vez de izquierda a derecha
                dig = Val(Mid(NumTmp, Pos, 1))
                Select Case c02
                    Case 1 : cen = dig
                    Case 2 : dec = dig
                    Case 3 : uni = dig
                End Select
                c02 = c02 + 1
                Pos = Pos + 1
            Loop
            letra3 = Centena(uni, dec, cen)
            letra2 = Decena(uni, dec, cen)
            letra1 = Unidad(uni, dec, cen)

            Select Case c01
                Case 1
                    If cen + dec + uni = 1 Then
                        Leyenda = "Billon "
                    ElseIf cen + dec + uni > 1 Then
                        Leyenda = "Billones "
                    End If
                Case 2
                    If cen + dec + uni >= 1 And Val(Mid(NumTmp, 7, 3)) = 0 Then
                        Leyenda = "Mil Millones "
                    ElseIf cen + dec + uni >= 1 Then
                        Leyenda = "Mil "
                    End If
                Case 3
                    If cen + dec = 0 And uni = 1 Then
                        Leyenda = "Millon "
                    ElseIf cen > 0 Or dec > 0 Or uni > 1 Then
                        Leyenda = "Millones "
                    End If
                Case 4
                    If cen + dec + uni >= 1 Then
                        Leyenda = "Mil "
                    End If
                Case 5
                    If cen + dec + uni >= 1 Then
                        Leyenda = ""
                    End If
            End Select

            c01 = c01 + 1
            TFNumero = TFNumero & letra3 & letra2 & letra1 & Leyenda

            Leyenda = ""
            letra1 = ""
            letra2 = ""
            letra3 = ""

        Loop

        If Val(NumTmp) = 0 Or Val(NumTmp) < 1 Then
            If CDbl(Tipo) = 2 Then
                Leyenda1 = "Cero dólares "
            Else
                Leyenda1 = "Cero Pesos "
            End If

        ElseIf Val(NumTmp) = 1 Or Val(NumTmp) < 2 Then
            If CDbl(Tipo) = 2 Then
                Leyenda1 = "dolar "
            Else
                Leyenda1 = "Peso "
            End If
        ElseIf Val(Mid(NumTmp, 4, 12)) = 0 Or Val(Mid(NumTmp, 10, 6)) = 0 Then
            If CDbl(Tipo) = 2 Then
                Leyenda1 = "de dólares "
            Else
                Leyenda1 = "de Pesos "
            End If

        Else
            If CDbl(Tipo) = 2 Then
                Leyenda1 = "dólares "
            Else
                Leyenda1 = "Pesos "
            End If
        End If

        If CDbl(Tipo) = 2 Then
            TFNumero = "(" & TFNumero & Leyenda1 & Mid(NumTmp, 17) & "/100 U.S.)"
        Else
            TFNumero = "(" & TFNumero & Leyenda1 & Mid(NumTmp, 17) & "/100 M.N.)"
        End If

        If mayusculas = True Then
            TFNumero = UCase(TFNumero)
        Else
            TFNumero = LCase(TFNumero)
        End If

        ConLetra = TFNumero
Error_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError(Err.Description, MsgBoxStyle.Critical)
        End If

    End Function

    'Función Auxiliar de la función CONLETRA
    Private Function Centena(ByVal uni As Integer, ByVal dec As Integer, ByVal cen As Integer) As String
        Dim ctexto As Object

        On Error GoTo Error_Renamed
        Select Case cen
            Case 1
                If dec + uni = 0 Then
                    ctexto = "cien "
                Else
                    ctexto = "ciento "
                End If
            Case 2
                ctexto = "doscientos "
            Case 3
                ctexto = "trescientos "
            Case 4
                ctexto = "cuatrocientos "
            Case 5
                ctexto = "quinientos "
            Case 6
                ctexto = "seiscientos "
            Case 7
                ctexto = "setecientos "
            Case 8
                ctexto = "ochocientos "
            Case 9
                ctexto = "novecientos "
            Case Else
                ctexto = ""
        End Select
        Centena = ctexto
        ctexto = ""
Error_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError(Err.Description, MsgBoxStyle.Critical)

        End If

    End Function

    'Función Auxiliar de la función CONLETRA
    Private Function Decena(ByVal uni As Integer, ByVal dec As Integer, ByVal cen As Integer) As String
        Dim ctexto As Object
        On Error GoTo Error_Renamed
        Select Case dec
            Case 1
                Select Case uni
                    Case 0
                        ctexto = "diez "
                    Case 1
                        ctexto = "once "
                    Case 2
                        ctexto = "doce "
                    Case 3
                        ctexto = "trece "
                    Case 4
                        ctexto = "catorce "
                    Case 5
                        ctexto = "quince "
                    Case 6 To 9
                        ctexto = "dieci"
                End Select
            Case 2
                If uni = 0 Then
                    ctexto = "veinte "
                ElseIf uni > 0 Then
                    ctexto = "veinti"
                End If
            Case 3
                ctexto = "treinta "
            Case 4
                ctexto = "cuarenta "
            Case 5
                ctexto = "cincuenta "
            Case 6
                ctexto = "sesenta "
            Case 7
                ctexto = "setenta "
            Case 8
                ctexto = "ochenta "
            Case 9
                ctexto = "noventa "
            Case Else
                ctexto = ""
        End Select

        If uni > 0 And dec > 2 Then ctexto = ctexto + "y "

        Decena = ctexto
        ctexto = ""
Error_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError(Err.Description, MsgBoxStyle.Critical)
        End If

    End Function

    'Función Auxiliar de la función CONLETRA
    Private Function Unidad(ByVal uni As Integer, ByVal dec As Integer, ByVal cen As Integer) As String
        Dim ctexto As Object

        On Error GoTo Error_Renamed
        If dec <> 1 Then
            Select Case uni
                Case 1
                    ctexto = "un "
                Case 2
                    ctexto = "dos "
                Case 3
                    ctexto = "tres "
                Case 4
                    ctexto = "cuatro "
                Case 5
                    ctexto = "cinco "
            End Select
        End If
        Select Case uni
            Case 6
                ctexto = "seis "
            Case 7
                ctexto = "siete "
            Case 8
                ctexto = "ocho "
            Case 9
                ctexto = "nueve "
        End Select

        Unidad = ctexto
        ctexto = ""
Error_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError(Err.Description, MsgBoxStyle.Critical)
        End If

    End Function

    'Función para convertir un Número a texto
    'Parámetros:
    '   1.- Año        : Valor que se va a convertir
    Public Function BICIESTO(ByRef Año As Integer) As Boolean
        BICIESTO = False

        If (Año Mod 4 = 0 And Año Mod 100 <> 0) Or Año Mod 400 = 0 Then
            BICIESTO = True
            Exit Function
        End If

    End Function
    'Esta función es para simular una máscara para valores de Hora
    'Parámetros:
    '   1.- Valor      : Contenido del txt que quiere formatear
    '   2.- Tecla      : Nombre del la tecla que se quiere controlar
    '   3.- Cursor     : Posición que se va a validar
    Public Function mskHora(ByRef Valor As String, ByRef tecla As Object, ByRef Cursor As Integer) As Object
        Dim I As Byte
        Dim Punto As Byte
        Dim Numeros As Byte

        'Por si quieren borrar los PUNTOS con la tecla back-space
        If tecla = 8 Then
            If Cursor < 1 Then
                mskHora = 0
                Exit Function
            Else
                If Mid(Valor, Cursor, 1) = ":" Then
                    If Cursor = 1 Or Cursor = Len(Valor) Then
                        mskHora = tecla
                        Exit Function
                    Else
                        mskHora = 0
                        Exit Function
                    End If
                Else
                    mskHora = tecla
                    Exit Function
                End If
            End If
        End If
        If tecla = Asc(":") And Valor = "" Then
            mskHora = 0
            Exit Function
        End If
        If Valor <> "" Then
            If tecla = Asc(":") And Len(Mid(Valor, Cursor + 1, Len(Valor) - Cursor + 1)) > 2 Then
                mskHora = 0
                Exit Function
            End If
        End If
        If tecla = Asc(":") And 2 = 0 Then
            mskHora = 0
            Exit Function
        End If
        Punto = 0
        For I = 1 To Len(Valor)
            If Mid(Valor, I, 1) = ":" Then
                Punto = I
                I = Len(Valor)
                If tecla = Asc(":") Then
                    mskHora = 0
                    Exit Function
                End If
            End If
        Next I
        If Punto = 0 And tecla <> Asc(":") Then
            If Len(Valor) = 2 Then
                mskHora = 0
                Exit Function
            End If
        ElseIf Punto <> 0 And tecla <> Asc(":") Then
            If Cursor < Punto Then
                Numeros = 0
                For I = 1 To Punto
                    If Mid(Valor, I, 1) <> ":" Then
                        Numeros = Numeros + 1
                    End If
                    If Numeros >= 2 Then
                        mskHora = 0
                        Exit Function
                    End If
                Next I
            Else
                Numeros = 0
                For I = Punto To Len(Valor)
                    If Mid(Valor, I, 1) <> ":" Then
                        Numeros = Numeros + 1
                    End If
                    If Numeros >= 2 Then
                        mskHora = 0
                        Exit Function
                    End If
                Next I
            End If
        End If
        mskHora = tecla
    End Function


    'Procedimiento que hace una emulación de la instrucción 'SendKeys "{TAB}"'
    'Su comportamiento es el de cambiar el foco al siguiente control habilitado y visible
    '   que siga en el orden de la propiedad TabIndex del control hacia adelante.
    'El parametro que recibe es la forma activa (Objeto 'Me')
    Public Sub AvanzarTab(ByRef Forma As System.Windows.Forms.Form)
        On Error Resume Next
        Dim intI As Integer
        Dim intTabIndSig, intIndSig As Integer
        'Si no hay control activo entonces toma valor -1
        If Forma.ActiveControl Is Nothing Then
            intIndSig = -1
        Else
            intTabIndSig = Forma.ActiveControl.TabIndex
            For intI = 0 To Forma.Controls.Count() - 1 Step 1
                If CType(Forma.Controls(intI), Object).TabIndex = intTabIndSig Then
                    intIndSig = intI
                    Exit For
                End If
            Next intI
            'Toma el tabindex mas grande de entre los controles habilitados y visibles
            For intI = 0 To Forma.Controls.Count() - 1 Step 1
                If CType(Forma.Controls(intI), Object).TabIndex > intTabIndSig And CType(Forma.Controls(intI), Object).Enabled And CType(Forma.Controls(intI), Object).Visible Then
                    intTabIndSig = CType(Forma.Controls(intI), Object).TabIndex
                    intIndSig = intI
                End If
            Next intI

            Err.Clear()
            'Ahora busca el menor de los que son mayores que el activo
            For intI = 0 To Forma.Controls.Count() - 1 Step 1
                If CType(Forma.Controls(intI), Object).TabIndex > Forma.ActiveControl.TabIndex Then
                    If CType(Forma.Controls(intI), Object).TabIndex > Forma.ActiveControl.TabIndex And CType(Forma.Controls(intI), Object).TabIndex < intTabIndSig And CType(Forma.Controls(intI), Object).Enabled And CType(Forma.Controls(intI), Object).Visible And CType(Forma.Controls(intI), Object).TabStop Then
                        If Err.Number = 0 Then
                            intTabIndSig = CType(Forma.Controls(intI), Object).TabIndex
                            intIndSig = intI
                        Else
                            Err.Clear()
                        End If
                    End If
                End If
            Next intI
        End If

        If intIndSig > -1 Then
            CType(Forma.Controls(intIndSig), Object).Focus()
        End If
    End Sub

    'Procedimiento que hace una emulación de la instrucción 'SendKeys "+{TAB}"'
    'Su comportamiento es el de cambiar el foco al siguiente control habilitado y visible
    '   que siga en el orden de la propiedad TabIndex del control hacia atrás.
    'El parametro que recibe es la forma activa (Objeto 'Me')
    Public Sub RetrocederTab(ByRef Forma As System.Windows.Forms.Form)
        On Error Resume Next
        Dim intI As Integer
        Dim intTabIndSig, intIndSig As Integer
        'Si no hay control activo entonces toma valor -1
        If Forma.ActiveControl Is Nothing Then
            intIndSig = -1
        Else
            intTabIndSig = Forma.ActiveControl.TabIndex
            For intI = 0 To Forma.Controls.Count() - 1 Step 1
                If CType(Forma.Controls(intI), Object).TabIndex = intTabIndSig Then
                    intIndSig = intI
                    Exit For
                End If
            Next intI
            'Toma el tabindex mas pequeño de entre los controles habilitados y visibles
            For intI = 0 To Forma.Controls.Count() - 1 Step 1
                If CType(Forma.Controls(intI), Object).TabIndex < intTabIndSig And CType(Forma.Controls(intI), Object).Enabled And CType(Forma.Controls(intI), Object).Visible Then
                    intTabIndSig = CType(Forma.Controls(intI), Object).TabIndex
                    intIndSig = intI
                End If
            Next intI

            Err.Clear()
            'Ahora busca el mayor de los que son menores que el activo
            For intI = 0 To Forma.Controls.Count() - 1 Step 1
                If CType(Forma.Controls(intI), Object).TabIndex < Forma.ActiveControl.TabIndex Then
                    If CType(Forma.Controls(intI), Object).TabIndex < Forma.ActiveControl.TabIndex And CType(Forma.Controls(intI), Object).TabIndex > intTabIndSig And CType(Forma.Controls(intI), Object).Enabled And CType(Forma.Controls(intI), Object).Visible And CType(Forma.Controls(intI), Object).TabStop Then
                        If Err.Number = 0 Then
                            intTabIndSig = CType(Forma.Controls(intI), Object).TabIndex
                            intIndSig = intI
                        Else
                            Err.Clear()
                        End If
                    End If
                End If
            Next intI
        End If

        If intIndSig > -1 Then
            CType(Forma.Controls(intIndSig), Object).Focus()
        End If
    End Sub

    'Graba mensajes en un log
    'los parametros que recibe son:
    '   ·Tipo    - Tipo de mensaje se esta grabando
    '   ·Modo    - Comportamiento del log
    '   ·Info    - Mensaje que se envía al log
    '   ·Archivo - Ruta y nombre del archivo log (opcional)
    Public Function LogMsj(ByRef Tipo As System.Diagnostics.TraceEventType, ByRef Modo As Object, ByVal Info As String, Optional ByRef Archivo As String = "") As Boolean
        On Error GoTo Merr
        Dim strFrmAct As String


        'If (Trim(Archivo) = "" And Modo = vbLogToFile) Or Modo = vbLogOff Or Modo = vbLogOverwrite Then Modo = vbLogAuto

        If Not System.Windows.Forms.Form.ActiveForm Is Nothing Then
            strFrmAct = System.Windows.Forms.Form.ActiveForm.Name
        Else
            strFrmAct = ""
        End If

        'Agrega caracteres de fin de mensaje
        Info = Info & vbNewLine & "[^^========^^]"

        'Inicia, Escribe y Cierra el log de la aplicación ~Å~
        'App.StartLogging(Archivo, Modo)
        My.Application.Log.WriteEntry(vbNewLine & "[" & Format(Now, "dd/mm/yyyy-hh:mm:ss") & "] <" & strFrmAct & ">" & vbNewLine & Info, Tipo)
        'App.StartLogging(Archivo, vbLogOff)

        LogMsj = True
        Exit Function
Merr:
        MsgBox("Error al intentar grabar el log" & vbNewLine & vbNewLine & Err.Number & vbNewLine & Err.Description, MsgBoxStyle.Critical, "Inserción en el log")
        Err.Clear()
    End Function

    '/////////////////////////////////////////////////////////////////////////////////
    Public Sub Pon_Tool()
        'Tip(System.Windows.Forms.Form.ActiveForm.ActiveControl, (MDIMenuPrincipalCorpo.status))
    End Sub


    Public Sub LimpiaDescBarraEstado()
        'MenuPrincipal.status.Items.Item(2).Text = ""
    End Sub

    Public Sub SelTxt()
        On Error Resume Next
        SelTextoTxt(GetActiveControl())
        If Err.Number <> 0 Then Err.Clear()
    End Sub

    Public Function BuscarRutaCarpeta(ByRef Forma As Object, ByRef cTitulo As Object) As String
        'Paimi
        On Error GoTo Merr
        Dim shlBusca As New Shell32.Shell
        Dim fldRecurso As Shell32.Folder
        Dim lngOpciones As Integer
        lngOpciones = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE Or BIF_VALIDATE
        'fldRecurso = shlBusca.BrowseForFolder(Forma.hWnd, Trim(cTitulo), lngOpciones)
        fldRecurso = shlBusca.BrowseForFolder(0, Trim(cTitulo), lngOpciones)
        If Not fldRecurso Is Nothing Then
            If Trim(fldRecurso.Items.Item.Path) <> "" Then
                BuscarRutaCarpeta = fldRecurso.Items.Item.Path
            End If
        End If
        Exit Function
Merr:
        'MostrarError "Ocurrió un error al intentar abrir la busqueda de las carpetas"
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub LimpiaControles(ByRef Forma As System.Windows.Forms.Form, Optional ByRef Ctl1 As String = "", Optional ByRef Ctl2 As String = "", Optional ByRef Ctl3 As String = "", Optional ByRef Ctl4 As String = "", Optional ByRef ctl5 As String = "")
        Dim Ctl As System.Windows.Forms.Control
        For Each Ctl In Forma.Controls
            If Ctl.Name <> Ctl1 And Ctl.Name <> Ctl2 And Ctl.Name <> Ctl3 And Ctl.Name <> Ctl4 And Ctl.Name <> ctl5 Then
                If TypeOf Ctl Is System.Windows.Forms.TextBox Or TypeOf Ctl Is System.Windows.Forms.MaskedTextBox Then
                    Ctl.Text = ""
                    Ctl.Tag = ""
                ElseIf TypeOf Ctl Is AxMSHierarchicalFlexGridLib.AxMSHFlexGrid Then
                    'Ctl.Clear
                    Ctl.Visible = True
                ElseIf TypeOf Ctl Is AxMSDataListLib.AxDataCombo Then
                    'Ctl.RowSource = Nothing
                    Ctl.Visible = Nothing

                End If
            End If
        Next Ctl
    End Sub

    ''''Procedimiento para mostrar la imagen de un Articulo
    'Public Sub BuscaImagen(Imagen As String, Image1 As Control)
    'On Error GoTo Errores
    '    Dim RecImg As Recordset
    '    Dim Archivo
    '
    '    If rutadeimagenes <> "" Then
    '        If Right(rutadeimagenes, 1) <> "\" Then rutadeimagenes = rutadeimagenes & "\"
    '        SQL = "Select CodRelacionado From CodigosRel Where CodigoInt = " & Val(Imagen) & " And Renglon = 2"
    '        Cmd.CommandText = SQL
    '        Set RecImg = Cmd.Execute
    '        Imagen = Trim(IIf(RecImg.RecordCount > 0, RecImg!CodRelacionado, ""))
    '        Archivo = Dir(rutadeimagenes & Imagen & ".*", vbHidden)
    '        If Archivo <> "" Then
    '            Image1.Picture = LoadPicture(rutadeimagenes & Archivo)
    '        Else
    '            Image1.Picture = LoadPicture(rutadeimagenes & "Rodarte.bmp")
    '        End If
    '    End If
    '
    'Errores:
    '    If Err.Number <> 0 Then
    '        If Err.Number = 53 Then
    '            MsgBox "No se Encontró la Imágen del Producto", vbOKOnly + vbExclamation, "Aviso"
    '            Err.Clear
    '        Else
    '            ModErrores.Errores
    '        End If
    '    End If
    'End Sub
End Module
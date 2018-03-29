'**********************************************************************************************************************'
'*PROGRAMA: ABC RUBROS DE APLICACIÓN Y ORIGEN JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018 
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualStudio.Data

Public Class frmCorpoABCRubrosdeAplicacionyOrigen

    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             ABC A RUBROS DE APLICACION Y ORIGEN                                                          *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      LUNES 12 DE MAYO DE 2003                                                                     *'
    '*FECHA DE TERMINACION : MARTES 13 DE MAYO DE 2003                                                                    *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    'Variables
    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSalir As Boolean 'Para Salir de la Captura Sin Preguntar Por Cambios
    Dim intSubindice As Integer 'con esta variable se lleva el conteo de los codigos de actividades borrados
    Dim intRenActual As Integer 'con esta se controla el renglon actual en el grid
    Dim intI As Integer 'esta variable se usa para contador de ciclos
    Dim intConsecutivo As Integer 'Para Conocer el Maximo Consecutivo
    Dim mblnBusqueda As Boolean 'Para identificar cuando esta buscando y que no provoque el lostfocus del txtdescripcion

    'Constantes
    Const C_ColDESCRIPCION As Integer = 0
    Const C_COLCODIGO As Integer = 1
    Const C_COLSTATUS As Integer = 2
    Const C_COLTAG As Integer = 3
    Public strControlActual As String 'Nombre del control actual
    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
        mblnBusqueda = True

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTCODIGO"
                strCaptionForm = "Consulta de Rubros de Aplicación y Origen"
                gStrSql = "SELECT RIGHT('0000'+LTRIM(CodOrigenAplicR),4) AS CODIGO, DescOrigenAplicR AS DESCRIPCION FROM CatOrigenAplicRecursos ORDER BY CodOrigenAplicR"
            Case "TXTDESCRIPCION"
                strCaptionForm = "Consulta de Rubros de Aplicación y Origen"
                gStrSql = "SELECT DescOrigenAplicR AS DESCRIPCION, RIGHT('0000'+LTRIM(CodOrigenAplicR),4) AS CODIGO FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & Trim(txtDescripcion.Text) & "%' ORDER BY DescOrigenAplicR"
            Case Else
                'Sale de este sub para QUE no ejecute ninguna opcion
                Exit Sub
        End Select
        strSQL = gStrSql 'Se hace uso de una variable temporal para el query
        'Si hubo cambios y es una modificacion entonces preguntara que si desea gravar los cambios
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se carguela consulta
                Case MsgBoxResult.Cancel 'Cancela la consulta
                    Exit Sub
            End Select
        End If
        gStrSql = strSQL 'Se regresa el valor de la variavle temporal a la variable original
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
            Exit Sub
        End If

        'Carga el formulario de consulta
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODIGO"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTDESCRIPCION"
                    .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                    .set_ColWidth(1, 0, 900) 'Columna del Código
            End Select
        End With
        FrmConsultas.ShowDialog()
        mblnBusqueda = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Eliminar()
        Try
            'On Error GoTo Merr
            If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "flexRubros" Then Exit Sub
            Dim blnTransaccion As Boolean
            gStrSql = "SELECT * FROM CatRubrosOrigenAplicRecursos WHERE CodRubro=" & Val(flexRubros.get_TextMatrix(flexRubros.Row, 1))
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount = 0 Then
                'MsgBox "Proporcione un código valido para eliminar.", vbInformation + vbOKOnly, "Mensaje"
                flexRubros.Clear()
                Encabezado()
                LlenaGrid()
                Exit Sub
            End If
            'Preguntar si desea borrar el registro
            Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "")
                Case MsgBoxResult.No
                    Exit Sub
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
            Cnn.BeginTrans()
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            blnTransaccion = True
            ModStoredProcedures.PR_IMECatRubrosOrigenAplicRecursos(txtCodigo.Text, flexRubros.get_TextMatrix(flexRubros.Row, 1), flexRubros.get_TextMatrix(flexRubros.Row, 0), C_ELIMINACION, CStr(0))
            Cmd.Execute()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Cnn.CommitTrans()
            blnTransaccion = False
            'Nuevo
            flexRubros.Clear()
            Encabezado()
            LlenaGrid()

            If Err.Number <> 0 Then
                If blnTransaccion = True Then Cnn.RollbackTrans()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                ModEstandar.MostrarError()
            End If
            'Merr:
        Catch ex As Exception
        End Try
    End Sub

    Function Guardar() As Boolean
        'On Error GoTo Merr
        Try
            Dim blnTransaccion As Boolean

            txtFlex.Visible = False
            txtFlex.Text = ""
            txtFlex.Focus()

            Guardar = False
            mblnCambiosEnCodigo = True

            If Cambios() = False Then
                Limpiar()
                Exit Function
            End If

            'Valida si todos los datos han sido llenados para poder ser guardados
            If ValidaDatos() = False Then
                Exit Function
            End If
            With flexRubros
                Cnn.BeginTrans()
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                blnTransaccion = True
                For intI = 1 To .Rows - 1
                    If Trim(.get_TextMatrix(intI, 0)) <> "" Then
                        ModEstandar.BorraCmd()
                        If Trim(.get_TextMatrix(intI, 2)) = C_NUEVO Then
                            ModStoredProcedures.PR_IMECatRubrosOrigenAplicRecursos(txtCodigo.Text, .get_TextMatrix(intI, 1), .get_TextMatrix(intI, 0), C_INSERCION, CStr(0))
                            Cmd.Execute()
                        ElseIf Trim(.get_TextMatrix(intI, 2)) = C_MODIFICADO Then
                            ModStoredProcedures.PR_IMECatRubrosOrigenAplicRecursos(txtCodigo.Text, .get_TextMatrix(intI, 1), .get_TextMatrix(intI, 0), C_MODIFICACION, CStr(0))
                            Cmd.Execute()
                        End If
                    End If
                Next intI
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Cnn.CommitTrans()
                MsgBox("La información fue guardada con éxito", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                blnTransaccion = False
            End With

            If Err.Number <> 0 Then
                If blnTransaccion = True Then Cnn.RollbackTrans()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                ModEstandar.MostrarError()
            End If
            Guardar = True
            mblnNuevo = True
            mblnCambiosEnCodigo = True
            Limpiar()
            'Merr:
        Catch ex As Exception
        End Try
    End Function

    Sub LlenaDatos()
        On Error GoTo Merr
        If Val(txtCodigo.Text) = 0 Then
            Nuevo()
            ModEstandar.AvanzarTab(Me)
            Exit Sub
        End If
        gStrSql = "SELECT MAX(CodRubro)+1 AS Codigo FROM CatRubrosOrigenAplicRecursos"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If IsDBNull(RsGral.Fields("Codigo").Value) Then
            intConsecutivo = 1
        Else
            intConsecutivo = RsGral.Fields("Codigo").Value
        End If

        'txtCodigo.Text = Format(txtCodigo.Text, "0000")
        For i = 1 To 4 - (txtCodigo.TextLength)
            txtCodigo.Text = String.Concat("0" + txtCodigo.Text)
        Next i

        gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR=" & Val(txtCodigo.Text)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtDescripcion.Text = Trim(RsGral.Fields("DescOrigenAplicR").Value)
            txtDescripcion.Tag = Trim(RsGral.Fields("DescOrigenAplicR").Value)
            If RsGral.Fields("Aplicacion").Value = "E" Then
                Label2.Visible = True
                Label2.Text = "Tipo de Aplicación: Entrada"
            ElseIf RsGral.Fields("Aplicacion").Value = "S" Then
                Label2.Visible = True
                Label2.Text = "Tipo de Aplicación: Salida"
            End If
        Else
            MsjNoExiste("El Origen y Aplicación de Recursos", gstrNombCortoEmpresa)
            Limpiar()
            Exit Sub
        End If
        LlenaGrid()
        mblnCambiosEnCodigo = False
        mblnNuevo = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatos_Desc()
        On Error GoTo Merr

        If Trim(txtDescripcion.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        '''gStrSql = "SELECT A.CodOrigenAplicR, A.DescOrigenAplicR, A.Aplicacion, IsNull(B.CodRubro,0) as CodRubro, IsNull(B.DescRubro,'') as DescRubro FROM CatOrigenAplicRecursos A INNER JOIN " & _
        '"CatRubrosOrigenAplicRecursos B ON A.CodOrigenAplicR = B.CodOrigAplicR WHERE A.DescOrigenAplicR = '" & Trim(txtDescripcion.text) & "' "
        gStrSql = "SELECT CodOrigenAplicR, DescOrigenAplicR, Aplicacion FROM CatOrigenAplicRecursos " & "WHERE DescOrigenAplicR = '" & Trim(txtDescripcion.Text) & "' "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtCodigo.Text = Format(RsGral.Fields("CodOrigenAplicR").Value, "0000")
            txtDescripcion.Text = Trim(RsGral.Fields("DescOrigenAplicR").Value)
            txtDescripcion.Tag = Trim(RsGral.Fields("DescOrigenAplicR").Value)
            If RsGral.Fields("Aplicacion").Value = "E" Then
                Label2.Visible = True
                Label2.Text = "Tipo de Aplicación: Entrada"
            ElseIf RsGral.Fields("Aplicacion").Value = "S" Then
                Label2.Visible = True
                Label2.Text = "Tipo de Aplicación: Salida"
            End If
        Else
            MsjNoExiste("El Origen y Aplicación de Recursos", gstrNombCortoEmpresa)
            Limpiar()
            Exit Sub
        End If
        LlenaGrid()
        mblnCambiosEnCodigo = False
        mblnNuevo = False

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaGrid()
        gStrSql = "SELECT * FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR=" & Val(txtCodigo.Text)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With RsGral
                intI = 1
                Do While Not .EOF
                    If intI > (flexRubros.Rows - 1) Then
                        flexRubros.Rows = flexRubros.Rows + 1
                    End If
                    With flexRubros
                        .Row = intI
                        .Col = 0
                        .Text = Trim(RsGral.Fields("DescRubro").Value)
                        .Col = 1
                        .Text = Format(RsGral.Fields("CodRubro").Value, "0000")
                        .Col = 2
                        .Text = "A"
                        .Col = 3
                        .Text = Trim(RsGral.Fields("DescRubro").Value)
                    End With
                    intI = intI + 1
                    .MoveNext()
                Loop
            End With
        End If
        flexRubros.Row = 1
        flexRubros.Col = 0
        flexRubros.Rows = flexRubros.Rows + 1
    End Sub

    Sub Limpiar()
        On Error Resume Next
        'Valida si Hubo Cambios, Pregunta si Desea Guardar
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la accion de limpiar la pantalla
                    Exit Sub
            End Select
        End If
        'If mblnCambiosEnCodigo Then
        txtCodigo.Text = ""
        Nuevo()
        Label2.Visible = False
        InicializaVariables()
        txtCodigo.Focus()
        '    Else
        '        If txtCodigo <> "" And txtDescripcion <> "" Then
        '            For intI = 1 To flexRubros.Rows - 1
        '                If flexRubros.TextMatrix(intI, 0) = "" Then
        '                    Exit For
        '                End If
        '            Next
        '            If intI > 9 Then
        '                flexRubros.TopRow = intI
        '            End If
        '            flexRubros.Row = intI
        '            flexRubros.SetFocus
        '        End If
        'End If
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
        mblnBusqueda = False
    End Sub

    Sub Nuevo()
        'On Error GoTo Merr
        Try
            txtCodigo.Enabled = True
            txtCodigo.Text = ""
            txtDescripcion.Text = ""
            flexRubros.Clear()
            Encabezado()
            Label2.Visible = False
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub


    Function Cambios() As Boolean
        Cambios = True
        With flexRubros
            For intI = 1 To .Rows - 1
                If Trim(.get_TextMatrix(intI, C_ColDESCRIPCION)) <> "" Then
                    If Trim(.get_TextMatrix(intI, C_ColDESCRIPCION)) <> Trim(.get_TextMatrix(intI, C_COLTAG)) Then Exit Function
                Else
                    Exit For
                End If
            Next intI
        End With
        Cambios = False
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        With flexRubros
            For intI = 1 To flexRubros.Rows - 1
                If Trim(.get_TextMatrix(intI, 0)) = "" And Trim(.get_TextMatrix(intI, 1)) = "" And Trim(.get_TextMatrix(intI, 3)) = "" Then Exit For
                If Trim(.get_TextMatrix(intI, 0)) = "" Then
                    MsgBox("Proporcione Descripción Del Rubro" & vbNewLine & ", Verifique por favor")
                    .Col = 0
                    Exit Function
                End If
            Next intI
        End With
        ValidaDatos = True
    End Function

    Private Sub flexRubros_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexRubros.DblClick
        flexRubros_KeyPressEvent(flexRubros, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub flexRubros_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexRubros.Enter
        Pon_Tool()
    End Sub

    Private Sub flexRubros_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexRubros.KeyDownEvent
        With flexRubros
            '''esta parte se usa cuando se suprime un codigo de
            '''una actividad, cada codigo suprimido se va almacenando
            '''en el grid  FlexDetalleSuprimido
            '        intRenActual = .Row
            '        Select Case KeyCode
            '            Case vbKeyDelete
            '                flexRubros_KeyPress vbKeyDelete
            '        End Select
        End With
    End Sub

    Private Sub flexRubros_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexRubros.KeyPressEvent
        With flexRubros
            '''Para que en la columna de porcentage
            '''no deje capturar caracteres sino solo numeros
            If .Col = 0 Then
                ModEstandar.gp_CampoMayusculas(eventArgs.keyAscii)
            End If
            '''si ya se capturo algo entonces se edita el grid
            '''ya sea con numeros, letras o enter
            'keyascii = 13
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And CDbl(Numerico(txtCodigo.Text)) <> 0 And Trim(txtDescripcion.Text) <> "" Then
                If (.Row > 1) Then
                    '''de tal modo que si el renglon es mayor que 1
                    '''y si un renglon antes del renglon actual esta vacio,
                    '''el renlgon actual no se editará
                    If Trim(.get_TextMatrix(.Row - 1, 0)) = "" Then
                        .Focus()
                        Exit Sub
                    End If
                End If
                MSHFlexGridEdit(flexRubros, txtFlex, eventArgs.keyAscii)
                If Len(Trim(txtFlex.Text)) <> 1 Then
                    SelTxt()
                End If
                '        ElseIf KeyAscii = vbKeyDelete Then
                '            Eliminar
            End If
        End With
    End Sub

    Private Sub flexRubros_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexRubros.Leave
        flexRubros.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub FlexRubros_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseDownEvent) Handles flexRubros.MouseDownEvent
        If eventArgs.button = 2 Then
            'PopupMenu MenuPrincipal.mnuContextual(0)
        End If
    End Sub

    Private Sub frmCorpoABCRubrosdeAplicacionyOrigen_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoABCRubrosdeAplicacionyOrigen_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCRubrosdeAplicacionyOrigen_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtCodigo" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
            Case System.Windows.Forms.Keys.Delete
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "flexRubros" Then
                    Eliminar()
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCRubrosdeAplicacionyOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCRubrosdeAplicacionyOrigen_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InicializaVariables()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        Encabezado()
    End Sub

    Private Sub frmCorpoABCRubrosdeAplicacionyOrigen_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        '    If Cambios() = True And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes 'Guardar el registro
        '                If Guardar() = False Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSalir = False
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCRubrosdeAplicacionyOrigen_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me.Frame1 = Nothing
        'MDIMenuPrincipalCorpo.mnuCatalogosOpc(10).Enabled = True
    End Sub

    Sub Encabezado()
        With flexRubros
            .Rows = 11
            .Row = 0
            .Col = 0
            .set_ColWidth(0, 0, 5300)
            .CellAlignment = 5
            .Text = "Descripción"
            .Col = 1 'Columna Para el Codigo
            .set_ColWidth(1, 0, 0)
            .Col = 2 'Columna Para el Status
            .set_ColWidth(2, 0, 0)
            .Col = 3 'Columna Para el Tag de la Descripción
            .set_ColWidth(3, 0, 0)
            .Row = 1
            .Col = 0
        End With
    End Sub

    Private Sub txtCodigo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Enter
        strControlActual = UCase("txtCodigo")
        SelTextoTxt(txtCodigo)
        Pon_Tool()
    End Sub

    Private Sub txtCodigo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodigo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodigo.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000

        If KeyCode = Keys.Enter Then
            txtCodigo_Leave(New Object, New EventArgs)
        End If

        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            KeyCode = 0
        Else
            'Si la tecla presionada fue Delete y Hay cambios, pregunta si se desea guardar
            If Cambios() = True And KeyCode = System.Windows.Forms.Keys.Delete Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes 'Guardar el registro
                        If Guardar() = False Then
                            KeyCode = 0
                            Exit Sub
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                        Nuevo()
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        txtCodigo.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
        End If
    End Sub

    Private Sub txtCodigo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'Si la tecla presionada no es numero regresa un 0
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta solo si existieron cambios
            If Cambios() = True And mblnNuevo = False Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes 'Guardar el registro
                        If Guardar() = False Then
                            KeyAscii = 0
                            GoTo EventExitSub
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        txtCodigo.Focus()
                        KeyAscii = 0
                        GoTo EventExitSub
                End Select
            End If
        End If
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodigo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If mblnCambiosEnCodigo = True And txtCodigo.Text <> "" Then 'si hubo cambios en el codigo hace la consulta
            LlenaDatos()
        End If
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        strControlActual = UCase("txtDescripcion")
        SelTextoTxt(txtDescripcion)
        Pon_Tool()
    End Sub

    Private Sub txtDescripcion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Leave
        'If Not mblnBusqueda Then LlenaDatos_Desc()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        With flexRubros
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    'txtFlex.Visible = False
                    'txtFlex.Text = ""
                    'flexRubros.Focus()
                Case System.Windows.Forms.Keys.Return
                    If Trim(txtFlex.Text) = "" Then
                        txtFlex.Visible = False
                        Exit Sub
                    End If
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .set_TextMatrix(.Row, .Col, txtFlex.Text)
                    'txtFlex.Visible = False
                    'txtFlex.Text = ""
                    'flexRubros.Focus()
                    If Trim(.get_TextMatrix(.Row, C_COLSTATUS)) = "" Then
                        .set_TextMatrix(.Row, C_COLSTATUS, C_NUEVO)
                    ElseIf Trim(.get_TextMatrix(.Row, C_COLSTATUS)) <> "" And (Trim(.get_TextMatrix(.Row, 0)) <> Trim(.get_TextMatrix(.Row, 3))) And Trim(.get_TextMatrix(.Row, C_COLSTATUS)) <> C_NUEVO Then
                        .set_TextMatrix(.Row, C_COLSTATUS, C_MODIFICADO)
                    ElseIf Trim(.get_TextMatrix(.Row, C_COLSTATUS)) <> "" And (Trim(.get_TextMatrix(.Row, 0)) = Trim(.get_TextMatrix(.Row, 3))) And Trim(.get_TextMatrix(.Row, C_COLSTATUS)) <> C_NUEVO Then
                        .set_TextMatrix(.Row, C_COLSTATUS, C_ACTIVO)
                    End If
                    If .get_TextMatrix(.Row, C_COLSTATUS) = C_NUEVO And .get_TextMatrix(.Row, C_COLCODIGO) = "" Then
                        .set_TextMatrix(.Row, C_COLCODIGO, intConsecutivo)
                        intConsecutivo = intConsecutivo + 1
                    End If
                    If .get_TextMatrix(.Row, C_ColDESCRIPCION) <> "" Then
                        .Row = .Row + 1
                        If .Row > 9 Then
                            .TopRow = .Row
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class
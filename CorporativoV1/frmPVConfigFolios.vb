Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmPVConfigFolios
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CONFIGURACION DE FOLIOS                                                                      *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 16 DE MAYO DE 2003                                                                   *'
    '*FECHA DE TERMINACION : VIERNES 16 DE MAYO DE 2003                                                                   *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'quedo pendiente validar si los consecutivos de los folios ya existian
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtDetalle As System.Windows.Forms.TextBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents msgFolios As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblNota As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnExistenFolioSucursal As Boolean 'Para Saber si Existes folios REgistrado para la sucursal Dada. Si no existen folio, no se puede modificar el consecutivo, y se guarda como cero unicamente.
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSALIR As Boolean 'Para Salir de la Captura Sin Preguntar Por Cambios
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim I As Integer
    Dim RsAux As ADODB.Recordset
    ' Para Manejar el FlexGrid  de Detalle
    Const C_ColFOLIO As Integer = 1
    Const C_ColDESCRIPCION As Integer = 0
    Const C_ColPREFIJO As Integer = 2
    Const C_ColCONSECUTIVO As Integer = 3
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Const C_ColCONSECUTIVOTAG As Integer = 4

    Function Guardar() As Boolean
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        Guardar = False
        If Cambios() = False And mblnExistenFolioSucursal = True Then
            Limpiar()
            Exit Function
        End If
        'Valida si todos los datos han sido llenados para poder ser guardados
        If ValidaDatos() = False Then
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True

        With msgFolios
            For I = 1 To .Rows - 1
                If mblnExistenFolioSucursal = True Then
                    If Val(.get_TextMatrix(I, C_ColFOLIO)) = 0 Then Exit For
                    ModStoredProcedures.PR_IMECatFolios(.get_TextMatrix(I, C_ColFOLIO), CStr(intCodSucursal), .get_TextMatrix(I, C_ColDESCRIPCION), .get_TextMatrix(I, C_ColPREFIJO), .get_TextMatrix(I, C_ColCONSECUTIVO), C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                Else
                    If Val(.get_TextMatrix(I, C_ColFOLIO)) = 0 Then Exit For
                    ModStoredProcedures.PR_IMECatFolios(.get_TextMatrix(I, C_ColFOLIO), CStr(intCodSucursal), .get_TextMatrix(I, C_ColDESCRIPCION), .get_TextMatrix(I, C_ColPREFIJO), .get_TextMatrix(I, C_ColCONSECUTIVO), C_INSERCION, CStr(0))
                    Cmd.Execute()
                End If
            Next
        End With
        Cnn.CommitTrans()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        blnTransaccion = False
        MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrCorpoNOMBREEMPRESA)
        Nuevo()
        InicializaVariables()
        Guardar = True
        Limpiar()
MErr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub LlenaDatos()
        On Error GoTo MErr
        gStrSql = "If not exists (SELECT * FROM CATFOLIOS WHERE CODALMACEN= " & intCodSucursal & ")" & "Select * , 0 as Existe  From CatFolios where CodAlmacen in ( select Min(Distinct(CodAlmacen)) from CatFolios)  " & "Else " & "SELECT * , 1 as Existe FROM CATFOLIOS WHERE CODALMACEN= " & intCodSucursal
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            'Llenar el Grid
            With msgFolios
                If RsGral.RecordCount > 5 Then .Rows = RsGral.RecordCount + 2
                For I = 1 To RsGral.RecordCount
                    .set_TextMatrix(I, C_ColFOLIO, RsGral.Fields("CodFolio").Value)
                    .set_TextMatrix(I, C_ColDESCRIPCION, RsGral.Fields("DescFolio").Value)
                    .set_TextMatrix(I, C_ColPREFIJO, RsGral.Fields("Prefijo").Value)
                    If RsGral.Fields("Existe").Value = 1 Then
                        .set_TextMatrix(I, C_ColCONSECUTIVO, RsGral.Fields("Consecutivo").Value)
                        .set_TextMatrix(I, C_ColCONSECUTIVOTAG, RsGral.Fields("Consecutivo").Value)
                        lblNota.Text = "Sucursal: " & UCase(Trim(dbcSucursales.Text))
                        mblnExistenFolioSucursal = True
                    Else
                        .set_TextMatrix(I, C_ColCONSECUTIVO, 0)
                        .set_TextMatrix(I, C_ColCONSECUTIVOTAG, 0)
                        lblNota.Text = "Folios no registrados para esta sucursal"
                        mblnExistenFolioSucursal = False
                    End If
                    RsGral.MoveNext()
                Next
            End With
        Else

        End If
        mblnCambiosEnCodigo = False
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function Cambios() As Boolean
        Cambios = True
        With msgFolios
            For I = 1 To .Rows - 1
                If Val(.get_TextMatrix(I, C_ColFOLIO)) = 0 Then Exit For
                If .get_TextMatrix(I, C_ColCONSECUTIVO) <> .get_TextMatrix(I, C_ColCONSECUTIVOTAG) Then
                    Exit Function
                End If
            Next
        End With
        Cambios = False
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = True
    End Function

    Sub Limpiar()
        On Error Resume Next
        'Valida si Hubo Cambios, Pregunta si Desea Guardar
        If Cambios() = True Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la accion de limpiar la pantalla
                    Exit Sub
            End Select
        End If
        Nuevo()
        FueraChange = True
        dbcSucursales.Text = ""
        dbcSucursales.Focus()
        FueraChange = False
    End Sub

    Sub Nuevo()
        On Error GoTo MErr
        InicializaVariables()
        lblNota.Text = ""
        msgFolios.Clear()
        Encabezado()
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnExistenFolioSucursal = False
        mblnCambiosEnCodigo = False
        mblnSALIR = False
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If

        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0

        If Trim(dbcSucursales.Text) <> "" Then
            DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
            LlenaDatos()
            msgFolios.TopRow = 1
            msgFolios.Row = 1
            msgFolios.Col = 0
        End If

    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen where TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
            Me.Close()
        End If
    End Sub

    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyUp
        Dim Aux As String
        Aux = dbcSucursales.Text
        FueraChange = True
        If dbcSucursales.SelectedItem <> 0 Then
            gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
            DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        End If
        FueraChange = False
        dbcSucursales.Text = Aux
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        FueraChange = True
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        If Trim(dbcSucursales.Text) <> "" Then
            LlenaDatos()
            'msgFolios.Row = 1
            'msgFolios.Col = 0
        End If
        FueraChange = False
    End Sub

    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucursales.MouseUp
        Dim Aux As String
        Aux = dbcSucursales.Text
        FueraChange = True
        'If dbcSucursales.SelectedItem <> 0 Then
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        'End If
        FueraChange = False
        dbcSucursales.Text = Aux
    End Sub

    Private Sub frmPVConfigFolios_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmPVConfigFolios_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmPVConfigFolios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtDetalle" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    If UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) = "DBCSUCURSALES" Then
                        mblnSALIR = True
                        Me.Close()
                    End If
                End If
        End Select
    End Sub

    Private Sub frmPVConfigFolios_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmPVConfigFolios_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        InicializaVariables()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Encabezado()
    End Sub

    Private Sub frmPVConfigFolios_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If Not mblnSALIR Then
            If Cambios() = True Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
                    Case MsgBoxResult.Yes 'Guardar el registro
                        If Guardar() = False Then
                            Cancel = 1
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
                    Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
                        Cancel = 1
                End Select
            End If
        Else
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSALIR = False
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmPVConfigFolios_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
    End Sub


    Sub Encabezado()
        'Genera el encabezao del Grid, asigna el tamaño y número de columas y centra el texto dentro de ellas
        Dim LnContador As Integer

        With msgFolios
            '''.FocusRect = flexFocusHeavy 'flexFocusLight 'flexFocusNone
            .WordWrap = False
            .FixedRows = 1
            .FixedCols = 0
            .set_ColWidth(C_ColFOLIO, 0, 0)
            .set_ColWidth(C_ColDESCRIPCION, 0, 3000)
            .set_ColWidth(C_ColPREFIJO, 0, 900)
            .set_ColWidth(C_ColCONSECUTIVO, 0, 1300)
            .set_ColWidth(C_ColCONSECUTIVOTAG, 0, 0)

            .set_TextMatrix(0, C_ColFOLIO, "FOLIO")
            .set_TextMatrix(0, C_ColDESCRIPCION, "DESCRIPCION")
            .set_TextMatrix(0, C_ColPREFIJO, "PREFIJO")
            .set_TextMatrix(0, C_ColCONSECUTIVO, "CONSECUTIVO")
            .set_TextMatrix(0, C_ColCONSECUTIVOTAG, "CONSECUTIVOTAG")

            .Row = 0
            For LnContador = 0 To C_ColCONSECUTIVO
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .Col = C_ColDESCRIPCION
            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
        End With
    End Sub

    Private Sub msgFolios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgFolios.Enter
        msgFolios.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        Pon_Tool()
    End Sub

    Private Sub msgFolios_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgFolios.DblClick
        msgFolios_KeyPressEvent(msgFolios, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub msgFolios_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgFolios.KeyPressEvent
        With msgFolios
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
                Select Case .Col
                    Case C_ColCONSECUTIVO ''-------------- SE EDITA EL codigo ---------------------'''''
                        eventArgs.keyAscii = ModEstandar.MskCantidad(txtDetalle.Text, eventArgs.keyAscii, 6, 0, (txtDetalle.SelectionStart))
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        If CDbl(Numerico(Trim(.get_TextMatrix(.Row, C_ColFOLIO)))) = 0 Then
                            .Focus()
                            Exit Sub
                        End If
                        ModEstandar.MSHFlexGridEdit(msgFolios, txtDetalle, eventArgs.keyAscii)
                        txtDetalle.SelectionStart = Len(txtDetalle.Text)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then
                            ModEstandar.SelTextoTxt(txtDetalle)
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub msgFolios_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgFolios.Leave
        msgFolios.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub txtDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Enter
        txtDetalle.Text = Trim(txtDetalle.Text)
        Pon_Tool()
    End Sub

    Private Sub txtDetalle_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDetalle.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Aqui se muestran los datos del control editable, en el Grid
        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
        With msgFolios
            Select Case KeyCode

                Case System.Windows.Forms.Keys.Escape
                    .Focus()
                    txtDetalle.Visible = False
                    txtDetalle.Text = ""
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                    .Focus()
                Case System.Windows.Forms.Keys.Return
                    'Si la Columna en que se está escribiendo es Codigo o Cantidad, Formatear el Valor par que quede numérico
                    If .Col = C_ColCONSECUTIVO Then
                        .set_TextMatrix(.Row, .Col, Val(txtDetalle.Text))
                        FueraChange = True
                        txtDetalle.Text = ""
                        txtDetalle.Visible = False
                        .Row = .Row + 1
                        .Col = C_ColCONSECUTIVO
                        .Focus()
                    End If
            End Select
        End With
    End Sub

    Private Sub txtDetalle_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDetalle.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'En este Evento se validan los datos que se introduzcan al control txtDetalle,dependiendo de la columan en que se esté editando
        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
        With msgFolios
            If .Col = C_ColCONSECUTIVO Then
                'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                KeyAscii = ModEstandar.MskCantidad(txtDetalle.Text, KeyAscii, 10, 0, (txtDetalle.SelectionStart))
            End If
        End With
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Leave
        txtDetalle.Visible = False
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPVConfigFolios))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblNota = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtDetalle = New System.Windows.Forms.TextBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.msgFolios = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        CType(Me.msgFolios, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblNota
        '
        Me.lblNota.BackColor = System.Drawing.SystemColors.Info
        Me.lblNota.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNota.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNota.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblNota.Location = New System.Drawing.Point(12, 208)
        Me.lblNota.Name = "lblNota"
        Me.lblNota.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNota.Size = New System.Drawing.Size(376, 21)
        Me.lblNota.TabIndex = 5
        Me.lblNota.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.lblNota, "Descripción de Artículos")
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.Color.Black
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(168, 24)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(60, 17)
        Me._Label1_0.TabIndex = 1
        Me._Label1_0.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_0, "Nombre de la Farmacia Actual")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtDetalle)
        Me.Frame2.Controls.Add(Me.dbcSucursales)
        Me.Frame2.Controls.Add(Me.msgFolios)
        Me.Frame2.Controls.Add(Me.lblNota)
        Me.Frame2.Controls.Add(Me._Label1_0)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(8, 8)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(393, 241)
        Me.Frame2.TabIndex = 0
        Me.Frame2.TabStop = False
        '
        'txtDetalle
        '
        Me.txtDetalle.AcceptsReturn = True
        Me.txtDetalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtDetalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDetalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDetalle.Location = New System.Drawing.Point(83, 95)
        Me.txtDetalle.MaxLength = 0
        Me.txtDetalle.Name = "txtDetalle"
        Me.txtDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDetalle.Size = New System.Drawing.Size(65, 21)
        Me.txtDetalle.TabIndex = 3
        Me.txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDetalle.Visible = False
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(220, 21)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(155, 21)
        Me.dbcSucursales.TabIndex = 2
        '
        'msgFolios
        '
        Me.msgFolios.DataSource = Nothing
        Me.msgFolios.Location = New System.Drawing.Point(16, 72)
        Me.msgFolios.Name = "msgFolios"
        Me.msgFolios.OcxState = CType(resources.GetObject("msgFolios.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgFolios.Size = New System.Drawing.Size(369, 130)
        Me.msgFolios.TabIndex = 4
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(8, 255)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(393, 74)
        Me.Panel3.TabIndex = 70
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(208, 14)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(160, 14)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(50, 42)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = " "
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.grabar
        Me.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnGuardar.Location = New System.Drawing.Point(11, 14)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(110, 14)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(50, 42)
        Me.btnLimpiar.TabIndex = 66
        Me.btnLimpiar.Text = " "
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Eliminar
        Me.btnEliminar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEliminar.Location = New System.Drawing.Point(61, 14)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'frmPVConfigFolios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(406, 348)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Frame2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(277, 168)
        Me.MaximizeBox = False
        Me.Name = "frmPVConfigFolios"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Folios del Sistema"
        Me.Frame2.ResumeLayout(False)
        CType(Me.msgFolios, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub
End Class
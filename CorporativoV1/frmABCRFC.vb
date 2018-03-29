Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmABCRFC
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optVigente As System.Windows.Forms.RadioButton
    Public WithEvents optCancelado As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodigo As System.Windows.Forms.TextBox
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents txtDomicilio As System.Windows.Forms.TextBox
    Public WithEvents txtColonia As System.Windows.Forms.TextBox
    Public WithEvents txtRFC As System.Windows.Forms.TextBox
    Public WithEvents txtCiudad As System.Windows.Forms.TextBox
    Public WithEvents txtCURP As System.Windows.Forms.TextBox
    Public WithEvents txtCodPostal As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaAlta As System.Windows.Forms.DateTimePicker
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_13 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_10 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    'Variables
    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnVigente As Boolean 'Estatus Vigente
    Dim mblnCancelado As Boolean 'Estatus Cancelado
    Dim FueraChange As Boolean
    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim tecla As Short
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnSalir As Button
    Public WithEvents btnImpirmir As Button
    Public WithEvents btnBuscar As Button
    Dim intCodSucursal As Short

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual

        If (txtCodigo.Text = "") Then
            strControlActual = UCase(txtCodigo.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
            strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control
        End If

        If strControlActual <> "TXTCODIGO" And strControlActual <> "TXTRFC" Then Exit Sub

        '    If Trim(dbcSucursales) = "" Then
        '        MsgBox "Proporcione la sucursal donde se buscarán los clientes", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
        '        dbcSucursales.SetFocus
        '        Exit Sub
        '    End If

        Select Case strControlActual
            Case "TXTCODIGO", "TXTRFC"
                FueraChange = True
                FrmConsultasClientes.strFormaActual = Me.Name
                FrmConsultasClientes.strControlActual = strControlActual
                FrmConsultasClientes.ShowDialog()
                FueraChange = False
                '        Case "TXTCODIGO"
                '            strCaptionForm = "Consulta de Clientes"
                '            gStrSql = "SELECT RIGHT('00000'+LTRIM(CodRFC),5) AS CODIGO, Rfc as RFC, DescClienteRFC AS NOMBRE From CatRFC " & _
                ''                    "WHERE  CodAlmacen = " & intCodSucursal & " or Isnull(Codalmacen,0) = 0   ORDER BY CodRFC"
                '        Case "TXTRFC"
                '            strCaptionForm = "Consulta de Clientes"
                '            gStrSql = "SELECT Rfc as RFC,  DescClienteRFC AS NOMBRE, RIGHT('00000'+LTRIM(CodRFC),5) AS CODIGO From CatRFC " & _
                ''                    "WHERE  CodAlmacen = " & intCodSucursal & " or Isnull(Codalmacen,0) = 0 ORDER BY RFC "
            Case Else
                'Sale de este sub para QUE no ejecute ninguna opcion
                Exit Sub
        End Select
        '    strSQL = gStrSql 'Se hace uso de una variable temporal para el query
        'Si hubo cambios y es una modificacion entonces preguntara que si desea gravar los cambios
        '    If Cambios = True And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrCorpoNOMBREEMPRESA)
        '            Case vbYes: 'Guardar el registro
        '                If Guardar = False Then
        '                    Exit Sub
        '                End If
        '            Case vbNo: 'No hace nada y permite que se carguela consulta
        '            Case vbCancel: 'Cancela la consulta
        '                Exit Sub
        '        End Select
        '    End If
        '    gStrSql = strSQL 'Se regresa el valor de la variavle temporal a la variable original
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
        '    Set RsGral = Cmd.Execute
        '    'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        '    If RsGral.RecordCount = 0 Then
        '        MsjNoExiste C_msgSINDATOS, gstrCorpoNOMBREEMPRESA
        '        Exit Sub
        '    End If
        '    'Carga el formulario de consulta
        '    Load FrmConsultas
        '    Call ConfiguraConsultas(FrmConsultas, 7700, RsGral, strTag, strCaptionForm)
        '    With FrmConsultas.Flexdet
        '        Select Case strControlActual
        '            Case "TXTCODIGO"
        '                .ColWidth(0) = 900 'Columna del Código
        '                .ColWidth(1) = 2000 'Columna de la RFC
        '                .ColWidth(2) = 4800 'Columna del Nombre
        '                .ColAlignment(0) = flexAlignLeftCenter
        '                .ColAlignment(1) = flexAlignCenterCenter
        '                .ColAlignment(2) = flexAlignLeftCenter
        '            Case "TXTRFC"
        '                .ColWidth(0) = 2000 'Columna de la RFC
        '                .ColWidth(1) = 4800 'Columna del Nombre
        '                .ColWidth(2) = 900 'Columna del Código
        '                .ColAlignment(2) = flexAlignLeftCenter
        '                .ColAlignment(0) = flexAlignCenterCenter
        '                .ColAlignment(1) = flexAlignLeftCenter
        '        End Select
        '    End With
        '    FueraChange = True
        '
        '    FrmConsultas.Show vbModal
        '    FueraChange = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    Function Guardar() As Object
        '    On Local Error GoTo MErr
        '    Dim blnTransaccion As Boolean
        '    Dim strEstatus As String
        '    Guardar = False
        '    Do While (Timer - sglTiempoCambio) <= 2.1
        '    Loop
        '    DoEvents
        '    If Cambios = False Then
        '        Limpiar
        '        Exit Function
        '    End If
        '    'Valida si todos los datos han sido llenados para poder ser guardados
        '    If ValidaDatos = False Then
        '        Exit Function
        '    End If
        '    If Val(txtCodigo) = 0 Then
        '        mblnNuevo = True
        '    End If
        '    Cnn.BeginTrans
        '    Me.MousePointer = vbHourglass
        '    blnTransaccion = True
        '    If optVigente(0).Value Then
        '        strEstatus = "V"
        '    ElseIf optCancelado(1).Value Then
        '        strEstatus = "C"
        '    End If
        '    If mblnNuevo Then
        '        ModStoredProcedures.PR_IMECatRFC Trim(txtCodigo), txtNombre, txtRFC, Trim(txtCURP), txtDomicilio, txtColonia, txtCiudad, _
        ''        txtCodPostal, strEstatus, Format(dtpFechaAlta, C_FORMATFECHAGUARDAR), CStr(gintCodAlmacen), C_INSERCION, 0
        '        Cmd.Execute
        '        txtCodigo = Format(Cmd("ID"), "00000")
        '    Else
        '        ModStoredProcedures.PR_IMECatRFC Trim(txtCodigo), txtNombre, txtRFC, Trim(txtCURP), txtDomicilio, txtColonia, txtCiudad, _
        ''        txtCodPostal, strEstatus, Format(dtpFechaAlta, C_FORMATFECHAGUARDAR), CStr(gintCodAlmacen), C_MODIFICACION, 0
        '        Cmd.Execute
        '    End If
        '    Me.MousePointer = vbDefault
        '    Cnn.CommitTrans
        '    blnTransaccion = False
        '    If mblnNuevo Then
        '        MsgBox "Los Datos del RFC han sido Grabados Correctamente con el Código: " & _
        ''            txtCodigo.text, vbInformation + vbOKOnly, gstrCorpoNOMBREEMPRESA
        '    Else
        '        MsgBox C_msgACTUALIZADO, vbInformation + vbOKOnly, ModVariables.gstrCorpoNOMBREEMPRESA
        '    End If
        '    Nuevo
        '    InicializaVariables
        '    Guardar = True
        '    Limpiar
        'MErr:
        '    If Err.Number <> 0 Then
        '        If blnTransaccion = True Then Cnn.RollbackTrans
        '        Me.MousePointer = vbDefault
        '        ModEstandar.MostrarError
        '    End If
        ''    Resume
    End Function

    Sub Limpiar()
        On Error Resume Next
        'Valida si Hubo Cambios, Pregunta si Desea Guardar
        '    If Cambios = True And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrCorpoNOMBREEMPRESA)
        '            Case vbYes: 'Guardar el registro
        '                If Guardar = False Then
        '                    Exit Sub
        '                End If
        '            Case vbNo: 'No hace nada y permite que se limpie la pantalla
        '            Case vbCancel: 'Cancela la accion de limpiar la pantalla
        '                Exit Sub
        '        End Select
        '    End If
        txtCodigo.Text = ""
        dbcSucursales.Text = ""
        Nuevo()
        dbcSucursales.Focus()
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr
        If Val(txtCodigo.Text) = 0 Then
            Nuevo()
            Exit Sub
        End If
        'txtCodigo.Text = Format(txtCodigo.Text, "00000")

        For i = 0 To 4 - txtCodigo.TextLength
            txtCodigo.Text = String.Concat("0" + txtCodigo.Text)
        Next i

        gStrSql = "SELECT Rfc.*, Alm.DescAlmacen FROM CatRFC Rfc INNER JOIN CatAlmacen Alm ON Rfc.CodAlmacen = Alm.CodAlmacen WHERE Rfc.CodRFC=" & Numerico(txtCodigo.Text)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            dbcSucursales.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            txtNombre.Text = Trim(RsGral.Fields("DescClienteRFC").Value)
            txtNombre.Tag = Trim(RsGral.Fields("DescClienteRFC").Value)
            txtRFC.Text = Trim(RsGral.Fields("Rfc").Value)
            txtRFC.Tag = Trim(RsGral.Fields("Rfc").Value)
            txtCURP.Text = Trim(RsGral.Fields("Curp").Value)
            txtCURP.Tag = Trim(RsGral.Fields("Curp").Value)
            txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value)
            txtDomicilio.Tag = Trim(RsGral.Fields("Domicilio").Value)
            txtColonia.Text = Trim(RsGral.Fields("Colonia").Value)
            txtColonia.Tag = Trim(RsGral.Fields("Colonia").Value)
            txtCiudad.Text = Trim(RsGral.Fields("Ciudad").Value)
            txtCiudad.Tag = Trim(RsGral.Fields("Ciudad").Value)
            txtCodPostal.Text = Trim(RsGral.Fields("CP").Value)
            txtCodPostal.Tag = Trim(RsGral.Fields("CP").Value)
            If RsGral.Fields("Estatus").Value = "V" Then
                optVigente.Checked = True
                mblnVigente = True
                mblnCancelado = False
            ElseIf RsGral.Fields("Estatus").Value = "C" Then
                optCancelado.Checked = True
                mblnVigente = False
                mblnCancelado = True
            End If
            dtpFechaAlta.Value = Format(RsGral.Fields("FechaAlta").Value, C_FORMATFECHAMOSTRAR)
            DesHabilitaControles()
        Else
            MsjNoExiste("El Cliente", gstrCorpoNOMBREEMPRESA)
            Limpiar()
        End If
        mblnCambiosEnCodigo = False
        mblnNuevo = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Nuevo()
        On Error GoTo Merr
        txtCodigo.Text = ""
        txtNombre.Text = ""
        txtNombre.Tag = ""
        txtDomicilio.Text = ""
        txtDomicilio.Tag = ""
        txtColonia.Text = ""
        txtColonia.Tag = ""
        txtCiudad.Text = ""
        txtCiudad.Tag = ""
        txtRFC.Text = ""
        txtRFC.Tag = ""
        txtCURP.Text = ""
        txtCURP.Tag = ""
        txtCodPostal.Text = ""
        txtCodPostal.Tag = ""
        dtpFechaAlta.Value = Today
        optVigente.Checked = True
        dbcSucursales.Text = ""
        InicializaVariables()
        HabilitaControles()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function Cambios() As Boolean
        Cambios = True
        If Trim(txtNombre.Text) <> txtNombre.Tag Then Exit Function
        If Trim(txtDomicilio.Text) <> txtDomicilio.Tag Then Exit Function
        If Trim(txtColonia.Text) <> txtColonia.Tag Then Exit Function
        If Trim(txtCiudad.Text) <> txtCiudad.Tag Then Exit Function
        If Trim(txtRFC.Text) <> txtRFC.Tag Then Exit Function
        If Trim(txtCURP.Text) <> Trim(txtCURP.Tag) Then Exit Function
        If optVigente.Checked <> mblnVigente Then Exit Function
        If optCancelado.Checked <> mblnCancelado Then Exit Function
        Cambios = False
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Len(Trim(txtNombre.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Nombre del Cliente", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtNombre.Focus()
            Exit Function
        End If
        If Len(Trim(txtDomicilio.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "DOMICILIO", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtDomicilio.Focus()
            Exit Function
        End If
        If Len(Trim(txtColonia.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Colonia", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtColonia.Focus()
            Exit Function
        End If
        If Len(Trim(txtCiudad.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Ciudad", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtCiudad.Focus()
            Exit Function
        End If
        If Len(Trim(txtRFC.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "RFC", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtRFC.Focus()
            Exit Function
        End If
        If ModEstandar.valida_RFCC(Trim(txtRFC.Text)) = False Then
            '        MsgBox "RFC Incorrecto." + vbNewLine + "Verifique Porfavor", vbInformation, gstrCorpoNombreEmpresa
            '        txtRFC.SetFocus
            '        Exit Function
        End If
        If Len(Trim(txtCodPostal.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Codigo Postal", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtCodPostal.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnVigente = True
        mblnCancelado = False
        mblnSalir = False
        FueraChange = False
    End Sub


    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        If dbcSucursales.Name <> "dbcSucursales" Then
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacen,Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        Nuevo()
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen  FROM CatAlmacen where  TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
        End If
    End Sub

    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoSucursal()
            '        Buscar
            Exit Sub
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE  TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
    End Sub

    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucursales.MouseUp
        PonerCodigoSucursal()
        '    Buscar
    End Sub

    Private Sub frmABCRFC_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmABCRFC_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmABCRFC_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If ActiveControl.Name <> "txtCodigo" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmABCRFC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        If ActiveControl.Name <> "txtEmail" Then
            KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmABCRFC_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        Nuevo()
    End Sub

    Private Sub frmABCRFC_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If mblnSalir Then
        '    '        If Cambios = True And mblnNuevo = False Then
        '    '            Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrCorpoNOMBREEMPRESA)
        '    '                Case vbYes: 'Guardar el registro
        '    '                    If Guardar = False Then
        '    '                        Cancel = 1
        '    '                    End If
        '    '                Case vbNo: 'No hace nada y permite el cierre del formulario
        '    '                Case vbCancel: 'Cancela el cierre del formulario sin guardar
        '    '                    Cancel = 1
        '    '            End Select
        '    '        End If
        '    '    Else
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSalir = False
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmABCRFC_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        '    MenuPrincipal.mnuCatalogosOpc(0).Enabled = True
    End Sub

    Private Sub optCancelado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCancelado.Enter
        Pon_Tool()
    End Sub

    Private Sub optVigente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optVigente.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCiudad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCiudad.Enter
        SelTextoTxt(txtCiudad)
        Pon_Tool()
    End Sub

    Private Sub txtCiudad_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCiudad.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '    ModEstandar.gp_CampoAlfanumerico KeyAscii
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodigo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Enter
        SelTextoTxt(txtCodigo)
        Pon_Tool()
    End Sub

    Private Sub txtCodigo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.TextChanged
        If FueraChange = True Then Exit Sub
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodigo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta solo si existieron cambios
            If Cambios() = True And mblnNuevo = False Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
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
        'If txtCodigo.Text <> Me.Text Then
        '    Exit Sub
        'End If
        '    If Numerico(Trim(txtCodigo)) = 0 Then 'si hubo cambios en el codigo hace la consulta
        '        txtCodigo = "00000"
        '    End If
        If (txtCodigo.Text <> "") Then
            LlenaDatos()
        End If
    End Sub

    Private Sub txtCodPostal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodPostal.Enter
        SelTextoTxt(txtCodPostal)
        Pon_Tool()
    End Sub

    Private Sub txtCodPostal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodPostal.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtColonia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtColonia.Enter
        SelTextoTxt(txtColonia)
        Pon_Tool()
    End Sub

    Private Sub txtCURP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCURP.Enter
        SelTextoTxt(txtCURP)
    End Sub


    Private Sub txtDomicilio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Enter
        SelTextoTxt(txtDomicilio)
        Pon_Tool()
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        SelTextoTxt(txtNombre)
        Pon_Tool()
    End Sub

    Private Sub TxtNombre_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtNombre.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoLetras(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRFC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.Enter
        SelTextoTxt(txtRFC)
        Pon_Tool()
    End Sub

    Private Sub txtRFC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRFC.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Back Then GoTo EventExitSub
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        KeyAscii = ModEstandar.Valida_RFC(txtRFC.Text, KeyAscii, Len(txtRFC.Text) + 1)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub PonerCodigoSucursal()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount = 0 Then
            intCodSucursal = 0
        Else
            intCodSucursal = RsGral.Fields("CodAlmacen").Value
        End If

    End Sub

    Sub DesHabilitaControles()
        txtCURP.Enabled = False
        txtNombre.Enabled = False
        txtDomicilio.Enabled = False
        txtColonia.Enabled = False
        txtCiudad.Enabled = False
        txtCodPostal.Enabled = False
        optCancelado.Enabled = False
        optVigente.Enabled = False
    End Sub


    Sub HabilitaControles()
        txtCURP.Enabled = True
        txtNombre.Enabled = True
        txtDomicilio.Enabled = True
        txtColonia.Enabled = True
        txtCiudad.Enabled = True
        txtCodPostal.Enabled = True
        optCancelado.Enabled = True
        optVigente.Enabled = True
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.optVigente = New System.Windows.Forms.RadioButton()
        Me.optCancelado = New System.Windows.Forms.RadioButton()
        Me.txtCodigo = New System.Windows.Forms.TextBox()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.TextBox()
        Me.txtColonia = New System.Windows.Forms.TextBox()
        Me.txtRFC = New System.Windows.Forms.TextBox()
        Me.txtCiudad = New System.Windows.Forms.TextBox()
        Me.txtCURP = New System.Windows.Forms.TextBox()
        Me.txtCodPostal = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaAlta = New System.Windows.Forms.DateTimePicker()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me._Label1_13 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_10 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnImpirmir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'optVigente
        '
        Me.optVigente.BackColor = System.Drawing.SystemColors.Control
        Me.optVigente.Checked = True
        Me.optVigente.Cursor = System.Windows.Forms.Cursors.Default
        Me.optVigente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optVigente.Location = New System.Drawing.Point(31, 19)
        Me.optVigente.Name = "optVigente"
        Me.optVigente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optVigente.Size = New System.Drawing.Size(64, 21)
        Me.optVigente.TabIndex = 10
        Me.optVigente.TabStop = True
        Me.optVigente.Text = "Vigente"
        Me.ToolTip1.SetToolTip(Me.optVigente, "Estatus Vigente")
        Me.optVigente.UseVisualStyleBackColor = False
        '
        'optCancelado
        '
        Me.optCancelado.BackColor = System.Drawing.SystemColors.Control
        Me.optCancelado.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCancelado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCancelado.Location = New System.Drawing.Point(120, 19)
        Me.optCancelado.Name = "optCancelado"
        Me.optCancelado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCancelado.Size = New System.Drawing.Size(77, 21)
        Me.optCancelado.TabIndex = 11
        Me.optCancelado.TabStop = True
        Me.optCancelado.Text = "Cancelado"
        Me.ToolTip1.SetToolTip(Me.optCancelado, "Estatus Cancelado")
        Me.optCancelado.UseVisualStyleBackColor = False
        '
        'txtCodigo
        '
        Me.txtCodigo.AcceptsReturn = True
        Me.txtCodigo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodigo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigo.Location = New System.Drawing.Point(88, 40)
        Me.txtCodigo.MaxLength = 5
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigo.Size = New System.Drawing.Size(89, 20)
        Me.txtCodigo.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodigo, "Codigo del Cliente")
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(88, 120)
        Me.txtNombre.MaxLength = 40
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(401, 20)
        Me.txtNombre.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre")
        '
        'txtDomicilio
        '
        Me.txtDomicilio.AcceptsReturn = True
        Me.txtDomicilio.BackColor = System.Drawing.SystemColors.Window
        Me.txtDomicilio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilio.Location = New System.Drawing.Point(88, 144)
        Me.txtDomicilio.MaxLength = 65
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilio.Size = New System.Drawing.Size(401, 20)
        Me.txtDomicilio.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtDomicilio, "Domicilio")
        '
        'txtColonia
        '
        Me.txtColonia.AcceptsReturn = True
        Me.txtColonia.BackColor = System.Drawing.SystemColors.Window
        Me.txtColonia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtColonia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtColonia.Location = New System.Drawing.Point(88, 168)
        Me.txtColonia.MaxLength = 30
        Me.txtColonia.Name = "txtColonia"
        Me.txtColonia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColonia.Size = New System.Drawing.Size(401, 20)
        Me.txtColonia.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtColonia, "Colonia")
        '
        'txtRFC
        '
        Me.txtRFC.AcceptsReturn = True
        Me.txtRFC.BackColor = System.Drawing.SystemColors.Window
        Me.txtRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFC.Location = New System.Drawing.Point(88, 72)
        Me.txtRFC.MaxLength = 15
        Me.txtRFC.Name = "txtRFC"
        Me.txtRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFC.Size = New System.Drawing.Size(137, 20)
        Me.txtRFC.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtRFC, "RFC")
        '
        'txtCiudad
        '
        Me.txtCiudad.AcceptsReturn = True
        Me.txtCiudad.BackColor = System.Drawing.SystemColors.Window
        Me.txtCiudad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCiudad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCiudad.Location = New System.Drawing.Point(88, 192)
        Me.txtCiudad.MaxLength = 30
        Me.txtCiudad.Name = "txtCiudad"
        Me.txtCiudad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCiudad.Size = New System.Drawing.Size(401, 20)
        Me.txtCiudad.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtCiudad, "Ciudad")
        '
        'txtCURP
        '
        Me.txtCURP.AcceptsReturn = True
        Me.txtCURP.BackColor = System.Drawing.SystemColors.Window
        Me.txtCURP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCURP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCURP.Location = New System.Drawing.Point(88, 96)
        Me.txtCURP.MaxLength = 18
        Me.txtCURP.Name = "txtCURP"
        Me.txtCURP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCURP.Size = New System.Drawing.Size(137, 20)
        Me.txtCURP.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtCURP, "RFC")
        '
        'txtCodPostal
        '
        Me.txtCodPostal.AcceptsReturn = True
        Me.txtCodPostal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodPostal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodPostal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodPostal.Location = New System.Drawing.Point(88, 216)
        Me.txtCodPostal.MaxLength = 10
        Me.txtCodPostal.Name = "txtCodPostal"
        Me.txtCodPostal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodPostal.Size = New System.Drawing.Size(97, 20)
        Me.txtCodPostal.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtCodPostal, "Codigo Postal")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.Controls.Add(Me.txtCodigo)
        Me.Frame1.Controls.Add(Me.txtNombre)
        Me.Frame1.Controls.Add(Me.txtDomicilio)
        Me.Frame1.Controls.Add(Me.txtColonia)
        Me.Frame1.Controls.Add(Me.txtRFC)
        Me.Frame1.Controls.Add(Me.txtCiudad)
        Me.Frame1.Controls.Add(Me.txtCURP)
        Me.Frame1.Controls.Add(Me.txtCodPostal)
        Me.Frame1.Controls.Add(Me.dtpFechaAlta)
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me._Label1_13)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_3)
        Me.Frame1.Controls.Add(Me._Label1_4)
        Me.Frame1.Controls.Add(Me._Label1_5)
        Me.Frame1.Controls.Add(Me._Label1_6)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me._Label1_10)
        Me.Frame1.Controls.Add(Me._Label1_7)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(8, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(499, 273)
        Me.Frame1.TabIndex = 12
        Me.Frame1.TabStop = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optVigente)
        Me.Frame2.Controls.Add(Me.optCancelado)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(272, 216)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(213, 46)
        Me.Frame2.TabIndex = 9
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Estatus "
        '
        'dtpFechaAlta
        '
        Me.dtpFechaAlta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaAlta.Location = New System.Drawing.Point(398, 88)
        Me.dtpFechaAlta.Name = "dtpFechaAlta"
        Me.dtpFechaAlta.Size = New System.Drawing.Size(89, 20)
        Me.dtpFechaAlta.TabIndex = 13
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(263, 19)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(224, 21)
        Me.dbcSucursales.TabIndex = 0
        '
        '_Label1_13
        '
        Me._Label1_13.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_13, CType(13, Short))
        Me._Label1_13.Location = New System.Drawing.Point(208, 24)
        Me._Label1_13.Name = "_Label1_13"
        Me._Label1_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_13.Size = New System.Drawing.Size(49, 15)
        Me._Label1_13.TabIndex = 23
        Me._Label1_13.Text = "Sucursal"
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(16, 40)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(41, 21)
        Me._Label1_2.TabIndex = 22
        Me._Label1_2.Text = "Codigo"
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_3, CType(3, Short))
        Me._Label1_3.Location = New System.Drawing.Point(16, 120)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(57, 21)
        Me._Label1_3.TabIndex = 21
        Me._Label1_3.Text = "Nombre :"
        '
        '_Label1_4
        '
        Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_4, CType(4, Short))
        Me._Label1_4.Location = New System.Drawing.Point(16, 144)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(66, 21)
        Me._Label1_4.TabIndex = 20
        Me._Label1_4.Text = "Dirección :"
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_5, CType(5, Short))
        Me._Label1_5.Location = New System.Drawing.Point(16, 168)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(57, 21)
        Me._Label1_5.TabIndex = 19
        Me._Label1_5.Text = "Colonia :"
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_6, CType(6, Short))
        Me._Label1_6.Location = New System.Drawing.Point(16, 72)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(41, 21)
        Me._Label1_6.TabIndex = 18
        Me._Label1_6.Text = "RFC :"
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(16, 192)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(57, 21)
        Me._Label1_1.TabIndex = 17
        Me._Label1_1.Text = "Ciudad :"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(16, 96)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(57, 21)
        Me._Label1_0.TabIndex = 16
        Me._Label1_0.Text = "C.U.R.P. :"
        '
        '_Label1_10
        '
        Me._Label1_10.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_10, CType(10, Short))
        Me._Label1_10.Location = New System.Drawing.Point(16, 216)
        Me._Label1_10.Name = "_Label1_10"
        Me._Label1_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_10.Size = New System.Drawing.Size(41, 21)
        Me._Label1_10.TabIndex = 15
        Me._Label1_10.Text = "C.P."
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_7, CType(7, Short))
        Me._Label1_7.Location = New System.Drawing.Point(328, 93)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(65, 15)
        Me._Label1_7.TabIndex = 14
        Me._Label1_7.Text = "Fecha Alta :"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(255, 278)
        Me.btnLimpiar.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(104, 42)
        Me.btnLimpiar.TabIndex = 42
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnEliminar
        '
        Me.btnEliminar.BackColor = System.Drawing.SystemColors.Control
        Me.btnEliminar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEliminar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEliminar.Location = New System.Drawing.Point(147, 278)
        Me.btnEliminar.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEliminar.Size = New System.Drawing.Size(104, 42)
        Me.btnEliminar.TabIndex = 41
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(39, 278)
        Me.btnGuardar.Margin = New System.Windows.Forms.Padding(2)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(104, 42)
        Me.btnGuardar.TabIndex = 40
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnSalir
        '
        Me.btnSalir.BackColor = System.Drawing.SystemColors.Control
        Me.btnSalir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSalir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSalir.Location = New System.Drawing.Point(255, 324)
        Me.btnSalir.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSalir.Size = New System.Drawing.Size(104, 42)
        Me.btnSalir.TabIndex = 44
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = False
        '
        'btnImpirmir
        '
        Me.btnImpirmir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImpirmir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImpirmir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImpirmir.Location = New System.Drawing.Point(147, 324)
        Me.btnImpirmir.Margin = New System.Windows.Forms.Padding(2)
        Me.btnImpirmir.Name = "btnImpirmir"
        Me.btnImpirmir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImpirmir.Size = New System.Drawing.Size(104, 42)
        Me.btnImpirmir.TabIndex = 43
        Me.btnImpirmir.Text = "Imprimir"
        Me.btnImpirmir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(363, 278)
        Me.btnBuscar.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(104, 42)
        Me.btnBuscar.TabIndex = 45
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmABCRFC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(515, 374)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnImpirmir)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(88, 226)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmABCRFC"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Catalogo de RFC  de Clientes"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
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

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnImpirmir_Click(sender As Object, e As EventArgs) Handles btnImpirmir.Click

    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class
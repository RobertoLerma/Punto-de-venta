Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class FrmConsultasClientes

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents chkTodasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.Panel
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Public RenAnt As Integer
    Public I As Integer
    Public intCodSucursal As Integer
    Public FueraChange As Boolean
    Public Tecla As Integer
    Public strSQL As String
    Public strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
    Public strCaptionForm As String 'Titulo que mostrara el formulario de consultas
    Public strControlActual As String 'Nombre del control actual
    Public strFormaActual As String 'Nombre de la Forma actual
    Public Columna As Integer
    Public cWHERE As String
    Public Desc As String
    Public WithEvents Flexdet As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public bandera As Boolean = False


    Private Sub chkTodasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodasSucursales.CheckStateChanged
        If chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            dbcSucursales.Text = ""
            intCodSucursal = 0
            dbcSucursales.Enabled = False
        Else
            txtNombre.Text = ""
            dbcSucursales.Enabled = True
            '        dbcSucursales.text = Trim(gstrNombreAlm)
            PonerCodigoSucursal()
        End If
        Desc = txtNombre.Text
        Buscar()
    End Sub

    Private Sub chkTodasSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkTodasSucursales.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If dbcSucursales.Enabled = True Then
                dbcSucursales.Focus()
            Else
                txtNombre.Focus()
            End If
        End If
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        'If dbcSucursales.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCChange(gStrSql, Tecla)
        intCodSucursal = 0
        '    Buscar
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen  FROM CatAlmacen where  TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursales.KeyDown
        Tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodasSucursales.Focus()
        End If
    End Sub

    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursales.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoSucursal()
            Buscar()
            Exit Sub
        ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            txtNombre.Focus()
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE  TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        Buscar()
    End Sub

    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursales.MouseUp
        'PonerCodigoSucursal()
        'Buscar()
    End Sub

    Private Sub FlexDet_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdet.DblClick
        Aceptar()
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
            intCodSucursal = RsGral.Fields("CodAlmacen").Value.ToString()
        End If
    End Sub

    Private Sub Flexdet_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'If Flexdet.Rows > 1 Then
        '    'Flexdet.TopRow = 1
        '    Flexdet.Row = 1
        '    'iNDICA LA CELDA QUE APARECERA SELECCIONADA
        '    Flexdet.ColSel = 1
        '    Flexdet.ColSel = 3
        '    '''Flexdet.HighLight = flexHighlightAlways
        '    '''Flexdet.FocusRect = flexFocusNone
        'End If
    End Sub

    Private Sub FlexDet_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent)
        If eventArgs.keyAscii = 13 Then
            Aceptar()
        End If
    End Sub

    Public Sub Aceptar()
        'On Error GoTo Merr
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        With Flexdet
            Select Case strTag
                Case "FRMCORPOABCCLIENTES.TXTCODIGO", "FRMCORPOABCCLIENTES.TXTNOMBRE"
                    With frmCorpoABCClientes
                        .intCodSucursal = CInt(Flexdet.get_TextMatrix(Flexdet.Row, 3))
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                        Exit Sub
                    End With
                Case "FRMFACTDATOSFISCALES.TXTCLIENTE", "FRMFACTDATOSFISCALES.TXTRFC"
                    With frmFactDatosFiscales
                        .txtCodRFC.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        Me.Close()
                        .LlenaDatosRFC()
                        Exit Sub
                    End With
                Case "FRMABCRFC.TXTCODIGO", "FRMABCRFC.TXTRFC"
                    With frmABCRFC
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        Me.Close()
                        .LlenaDatos()
                        Exit Sub
                    End With
                Case Else
                    Me.Close()
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
            End Select
            'System.Windows.Forms.SendKeys.Send("{ENTER}")
        End With
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MostrarError("Ha ocurrido un error")
    End Sub

    Private Sub Flexdet_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Flexdet.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub FrmConsultasClientes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then Me.Close()
        If KeyCode = System.Windows.Forms.Keys.Return Then ModEstandar.AvanzarTab(Me)
    End Sub

    Private Sub FrmConsultasClientes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Public Sub FrmConsultasClientes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'InitializeComponent()
        KeyPreview = True
        ModEstandar.CentrarForma(FrmConsultas)
        'System.Windows.Forms.SendKeys.Send("{RIGHT}")
        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        'strFormaActual = UCase(System.Windows.Forms.Form.ActiveForm.Name)
        strTag = UCase(strFormaActual & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
        chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        PonerCodigoSucursal()
        Buscar()
    End Sub

    Private Sub FrmConsultasClientes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'FrmConsultas = Nothing
        IsNothing(FrmConsultas)
    End Sub

    Sub BuscarClientes()
        'On Error GoTo Merr
        cWHERE = ""
        strCaptionForm = "Consulta de Clientes"
        Desc = Trim(txtNombre.Text)

        If chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = " And (CodAlmacen = " & intCodSucursal & " Or CodAlmacen is Null)"
        End If

        strSQL = "SELECT '' as XX, CodCliente as Código , Ltrim(rtrim(DescCliente))  as Descripción , CodAlmacen as Sucursal  From CatClientes " & "WHERE DescCliente  LIKE '%" & Trim(Desc) & "%' " & cWHERE
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        Flexdet_Leave(Flexdet, New System.EventArgs())
        Flexdet.Recordset = RsGral
        Me.Text = strCaptionForm
        ''FrmConsultasClientes.Flexdet.FocusRect = flexFocusNone
        ''FrmConsultasClientes.Flexdet.HighLight = flexHighlightWithFocus
        With Me.Flexdet
            '''.Clear
            .set_ColWidth(0, 0, 0)
            .set_ColWidth(1, 0, 900)
            .set_ColWidth(2, 0, 4800)
            .set_ColWidth(3, 0, 900)
            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            For I = 1 To 3
                .set_ColAlignmentFixed(I, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            Next I
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscarRfc()
        'On Error GoTo Merr
        cWHERE = ""
        strCaptionForm = "Consulta de RFC de clientes"
        If chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = " And (CodAlmacen = " & intCodSucursal & " Or CodAlmacen is Null)"
        End If
        strSQL = "SELECT CodRfc as Código ,Ltrim(Rtrim(rfc)) as RFC,  Ltrim(rtrim(DescClienteRfc))  as Nombre , CodAlmacen as Sucursal  From CatRFC " & "WHERE DescClienteRfc  LIKE '%" & Trim(Desc) & "%' " & cWHERE
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        'Carga el formulario de consulta
        Flexdet_Leave(Flexdet, New System.EventArgs())
        Flexdet.Recordset = RsGral
        Me.Text = strCaptionForm
        With Me.Flexdet
            .set_ColWidth(0, 0, 0)
            .set_ColWidth(1, 0, 1600)
            .set_ColWidth(2, 0, 4100)
            .set_ColWidth(3, 0, 900)
            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            For I = 1 To 3
                .set_ColAlignmentFixed(I, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            Next
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Buscar()
        'On Error GoTo Merr
        Select Case strFormaActual
            Case "FRMFACTURACIONTICKETS", "FRMFACTDATOSFISCALES", "FRMABCRFC"
                BuscarRfc()
            Case Else
                BuscarClientes()
        End Select
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub txtNombre_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.TextChanged
        Desc = Trim(txtNombre.Text)
        Buscar()
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        SelTextoTxt(txtNombre)
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmConsultasClientes))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.chkTodasSucursales = New System.Windows.Forms.CheckBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Flexdet = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(57, 88)
        Me.txtNombre.MaxLength = 40
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(314, 20)
        Me.txtNombre.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre")
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.Color.Black
        Me._Label1_1.Location = New System.Drawing.Point(8, 35)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(60, 17)
        Me._Label1_1.TabIndex = 6
        Me._Label1_1.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_1, "Nombre de la Farmacia Actual")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtNombre)
        Me.Frame2.Controls.Add(Me.Frame1)
        Me.Frame2.Controls.Add(Me._Label1_3)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(8, 0)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(465, 121)
        Me.Frame2.TabIndex = 3
        Me.Frame2.TabStop = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chkTodasSucursales)
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(57, 9)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(329, 65)
        Me.Frame1.TabIndex = 4
        '
        'chkTodasSucursales
        '
        Me.chkTodasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodasSucursales.Location = New System.Drawing.Point(7, 11)
        Me.chkTodasSucursales.Name = "chkTodasSucursales"
        Me.chkTodasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodasSucursales.Size = New System.Drawing.Size(147, 18)
        Me.chkTodasSucursales.TabIndex = 5
        Me.chkTodasSucursales.Text = "Todas las Sucursales"
        Me.chkTodasSucursales.UseVisualStyleBackColor = False
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(70, 31)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(251, 21)
        Me.dbcSucursales.TabIndex = 7
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(13, 92)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(45, 17)
        Me._Label1_3.TabIndex = 0
        Me._Label1_3.Text = "Cliente"
        '
        'Flexdet
        '
        Me.Flexdet.DataSource = Nothing
        Me.Flexdet.Location = New System.Drawing.Point(8, 127)
        Me.Flexdet.Name = "Flexdet"
        Me.Flexdet.OcxState = CType(resources.GetObject("Flexdet.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Flexdet.Size = New System.Drawing.Size(465, 182)
        Me.Flexdet.TabIndex = 10
        '
        'FrmConsultasClientes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(480, 330)
        Me.Controls.Add(Me.Flexdet)
        Me.Controls.Add(Me.Frame2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(207, 184)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmConsultasClientes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = " "
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

End Class
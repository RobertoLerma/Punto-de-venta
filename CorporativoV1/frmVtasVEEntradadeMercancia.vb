Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6



Public Class frmVtasVEEntradadeMercancia
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtFolioSalida As System.Windows.Forms.TextBox
    Public WithEvents txtFolio As System.Windows.Forms.TextBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodSucMatriz As System.Windows.Forms.TextBox
    Public WithEvents txtDescSucMatriz As System.Windows.Forms.TextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtDescVendExterno As System.Windows.Forms.TextBox
    Public WithEvents txtCodSucVendExterno As System.Windows.Forms.TextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents txtEnvia As System.Windows.Forms.TextBox
    Public WithEvents txtEntrega As System.Windows.Forms.TextBox
    Public WithEvents txtRecibe As System.Windows.Forms.TextBox
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public strControlActual As String 'Nombre del control actual

    'Variables
    Dim mblnSalir As Boolean
    Dim mblnNuevo As Boolean
    Dim mblnCambios As Boolean
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnEliminar As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnLimpiar As Button
    Dim Fecha As Date
    Public bandera As Boolean = False

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTFOLIO"
                strCaptionForm = "Busqueda de Folios de Recepción de Mercancia a Vend. Externos"
                gStrSql = "SELECT FolioAlmacen AS FOLIO,FechaAlmacen AS FECHA, Concepto AS CONCEPTO FROM " & "MovtosAlmacenCab WHERE CodAlmacen = " & txtCodSucMatriz.Text & " AND CodMovtoAlm = " & C_EntradaPorDevoluciondeVendedoresExternos & " AND Estatus = 'V' ORDER BY FolioAlmacen Desc,FechaAlmacen Desc"
            Case "TXTFOLIOSALIDA"
                strCaptionForm = "Busqueda de Folios de Entrega de Mercancia a Vend. Externos"
                gStrSql = "SELECT FolioAlmacen AS FOLIO,FechaAlmacen AS FECHA, Concepto AS CONCEPTO FROM " & "MovtosAlmacenCab WHERE CodAlmacen = " & txtCodSucMatriz.Text & " AND CodMovtoAlm = " & C_SalidaAVendedoresExternos & " AND Estatus = 'V' ORDER BY FolioAlmacen Desc,FechaAlmacen Desc"
            Case Else
                Exit Sub
        End Select
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
        ConfiguraConsultas(FrmConsultas, 11000, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTFOLIO", "TXTFOLIOSALIDA"
                    .set_ColAlignment(0, 0)
                    .set_ColAlignment(1, 3)
                    .set_ColAlignment(2, 0)
                    .set_ColWidth(0, 0, 1600)
                    .set_ColWidth(1, 0, 1400)
                    .set_ColWidth(2, 0, 7500)
                    For I = 1 To .Rows - 1
                        .set_TextMatrix(I, 1, VB6.Format(.get_TextMatrix(I, 1), "dd/mmm/yyyy"))
                        .set_TextMatrix(I, 2, Trim(QuitaEnter(.get_TextMatrix(I, 2))))
                    Next
                    FrmConsultas.Left = VB6.TwipsToPixelsX(2000)
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasVEEntradadeMercancia))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtFolioSalida = New System.Windows.Forms.TextBox()
        Me.txtFolio = New System.Windows.Forms.TextBox()
        Me.txtCodSucMatriz = New System.Windows.Forms.TextBox()
        Me.txtDescSucMatriz = New System.Windows.Forms.TextBox()
        Me.txtDescVendExterno = New System.Windows.Forms.TextBox()
        Me.txtCodSucVendExterno = New System.Windows.Forms.TextBox()
        Me.txtEnvia = New System.Windows.Forms.TextBox()
        Me.txtEntrega = New System.Windows.Forms.TextBox()
        Me.txtRecibe = New System.Windows.Forms.TextBox()
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame5.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtFolioSalida
        '
        Me.txtFolioSalida.AcceptsReturn = True
        Me.txtFolioSalida.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioSalida.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioSalida.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioSalida.Location = New System.Drawing.Point(221, 20)
        Me.txtFolioSalida.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioSalida.MaxLength = 17
        Me.txtFolioSalida.Name = "txtFolioSalida"
        Me.txtFolioSalida.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioSalida.Size = New System.Drawing.Size(86, 20)
        Me.txtFolioSalida.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtFolioSalida, "Folio de Salida de Mercancia")
        '
        'txtFolio
        '
        Me.txtFolio.AcceptsReturn = True
        Me.txtFolio.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolio.Location = New System.Drawing.Point(48, 20)
        Me.txtFolio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolio.MaxLength = 17
        Me.txtFolio.Name = "txtFolio"
        Me.txtFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolio.Size = New System.Drawing.Size(86, 20)
        Me.txtFolio.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolio, "Folio de Entrada de Mercancia")
        '
        'txtCodSucMatriz
        '
        Me.txtCodSucMatriz.AcceptsReturn = True
        Me.txtCodSucMatriz.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucMatriz.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucMatriz.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucMatriz.Location = New System.Drawing.Point(64, 23)
        Me.txtCodSucMatriz.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucMatriz.MaxLength = 3
        Me.txtCodSucMatriz.Name = "txtCodSucMatriz"
        Me.txtCodSucMatriz.ReadOnly = True
        Me.txtCodSucMatriz.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucMatriz.Size = New System.Drawing.Size(26, 20)
        Me.txtCodSucMatriz.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtCodSucMatriz, "Codigo de la Sucursal Matriz")
        '
        'txtDescSucMatriz
        '
        Me.txtDescSucMatriz.AcceptsReturn = True
        Me.txtDescSucMatriz.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescSucMatriz.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescSucMatriz.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescSucMatriz.Location = New System.Drawing.Point(101, 24)
        Me.txtDescSucMatriz.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescSucMatriz.MaxLength = 0
        Me.txtDescSucMatriz.Name = "txtDescSucMatriz"
        Me.txtDescSucMatriz.ReadOnly = True
        Me.txtDescSucMatriz.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescSucMatriz.Size = New System.Drawing.Size(155, 20)
        Me.txtDescSucMatriz.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtDescSucMatriz, "Descripción de la Sucursal Matriz")
        '
        'txtDescVendExterno
        '
        Me.txtDescVendExterno.AcceptsReturn = True
        Me.txtDescVendExterno.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescVendExterno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescVendExterno.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescVendExterno.Location = New System.Drawing.Point(88, 24)
        Me.txtDescVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescVendExterno.MaxLength = 0
        Me.txtDescVendExterno.Name = "txtDescVendExterno"
        Me.txtDescVendExterno.ReadOnly = True
        Me.txtDescVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescVendExterno.Size = New System.Drawing.Size(171, 20)
        Me.txtDescVendExterno.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtDescVendExterno, "Descripción de la Sucursal Matriz")
        '
        'txtCodSucVendExterno
        '
        Me.txtCodSucVendExterno.AcceptsReturn = True
        Me.txtCodSucVendExterno.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucVendExterno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucVendExterno.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucVendExterno.Location = New System.Drawing.Point(59, 24)
        Me.txtCodSucVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucVendExterno.MaxLength = 3
        Me.txtCodSucVendExterno.Name = "txtCodSucVendExterno"
        Me.txtCodSucVendExterno.ReadOnly = True
        Me.txtCodSucVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucVendExterno.Size = New System.Drawing.Size(26, 20)
        Me.txtCodSucVendExterno.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodSucVendExterno, "Codigo del Vendedor Externo")
        '
        'txtEnvia
        '
        Me.txtEnvia.AcceptsReturn = True
        Me.txtEnvia.BackColor = System.Drawing.SystemColors.Window
        Me.txtEnvia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEnvia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEnvia.Location = New System.Drawing.Point(71, 98)
        Me.txtEnvia.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEnvia.MaxLength = 50
        Me.txtEnvia.Name = "txtEnvia"
        Me.txtEnvia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEnvia.Size = New System.Drawing.Size(200, 20)
        Me.txtEnvia.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtEnvia, "Persona que Envia la Mercancia")
        '
        'txtEntrega
        '
        Me.txtEntrega.AcceptsReturn = True
        Me.txtEntrega.BackColor = System.Drawing.SystemColors.Window
        Me.txtEntrega.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEntrega.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEntrega.Location = New System.Drawing.Point(71, 119)
        Me.txtEntrega.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEntrega.MaxLength = 50
        Me.txtEntrega.Name = "txtEntrega"
        Me.txtEntrega.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEntrega.Size = New System.Drawing.Size(200, 20)
        Me.txtEntrega.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtEntrega, "Persona que Entrega la Mercancia")
        '
        'txtRecibe
        '
        Me.txtRecibe.AcceptsReturn = True
        Me.txtRecibe.BackColor = System.Drawing.SystemColors.Window
        Me.txtRecibe.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRecibe.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRecibe.Location = New System.Drawing.Point(71, 139)
        Me.txtRecibe.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRecibe.MaxLength = 50
        Me.txtRecibe.Name = "txtRecibe"
        Me.txtRecibe.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRecibe.Size = New System.Drawing.Size(200, 20)
        Me.txtRecibe.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtRecibe, "Persona que Recibe la Mercancia")
        '
        'txtConcepto
        '
        Me.txtConcepto.AcceptsReturn = True
        Me.txtConcepto.BackColor = System.Drawing.SystemColors.Window
        Me.txtConcepto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConcepto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConcepto.Location = New System.Drawing.Point(12, 20)
        Me.txtConcepto.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConcepto.MaxLength = 150
        Me.txtConcepto.Multiline = True
        Me.txtConcepto.Name = "txtConcepto"
        Me.txtConcepto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConcepto.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtConcepto.Size = New System.Drawing.Size(206, 47)
        Me.txtConcepto.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Concepto por el que se Entrega la Mercancia")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtFolioSalida)
        Me.Frame1.Controls.Add(Me.txtFolio)
        Me.Frame1.Controls.Add(Me.dtpFecha)
        Me.Frame1.Controls.Add(Me.Label8)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(12, 13)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(523, 53)
        Me.Frame1.TabIndex = 21
        Me.Frame1.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(383, 20)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(98, 20)
        Me.dtpFecha.TabIndex = 22
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(138, 23)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(92, 17)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Folio de Salida :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(37, 17)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Folio :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(339, 23)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(46, 17)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Fecha :"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.Frame3)
        Me.Frame2.Controls.Add(Me.Frame4)
        Me.Frame2.Controls.Add(Me.txtEnvia)
        Me.Frame2.Controls.Add(Me.txtEntrega)
        Me.Frame2.Controls.Add(Me.txtRecibe)
        Me.Frame2.Controls.Add(Me.Frame5)
        Me.Frame2.Controls.Add(Me.txtFlex)
        Me.Frame2.Controls.Add(Me.flexDetalle)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(12, 72)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(594, 376)
        Me.Frame2.TabIndex = 11
        Me.Frame2.TabStop = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.txtCodSucMatriz)
        Me.Frame3.Controls.Add(Me.txtDescSucMatriz)
        Me.Frame3.Controls.Add(Me.Label3)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(297, 20)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(280, 59)
        Me.Frame3.TabIndex = 16
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Entrega a..."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(6, 26)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Sucursal :"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtDescVendExterno)
        Me.Frame4.Controls.Add(Me.txtCodSucVendExterno)
        Me.Frame4.Controls.Add(Me.Label4)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(12, 20)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(268, 59)
        Me.Frame4.TabIndex = 14
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Vendedor Externo"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(6, 26)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(54, 13)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Sucursal :"
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.txtConcepto)
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame5.Location = New System.Drawing.Point(297, 84)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(280, 79)
        Me.Frame5.TabIndex = 13
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Concepto"
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(15, 240)
        Me.txtFlex.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(68, 20)
        Me.txtFlex.TabIndex = 12
        Me.txtFlex.Visible = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(16, 216)
        Me.flexDetalle.Margin = New System.Windows.Forms.Padding(2)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(561, 139)
        Me.flexDetalle.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(12, 98)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(50, 17)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Envia :"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(12, 119)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(55, 17)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Entrega :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(12, 141)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(50, 14)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Recibe :"
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(82, 470)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 65
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(189, 470)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 66
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(296, 470)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(93, 35)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(395, 470)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 68
        Me.btnLimpiar.Text = "Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'frmVtasVEEntradadeMercancia
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(615, 512)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVEEntradadeMercancia"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Recepción de Mercancia del Vendedor Externo"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Sub Cancelar()
        On Error GoTo Merr
        Dim FolioSalida As String
        Dim blnTransaccion As Boolean
        Dim I As Integer
        If BuscaLiquidacion() Then
            MsgBox("Este folio ya tiene una liquidación" & vbNewLine & "No es posible cancelarlo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If mblnNuevo Then
            Exit Sub
        End If
        Select Case MsgBox("¿Desea cancelar este folio de recepción de mercancia?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
            Case MsgBoxResult.No
                Exit Sub
            Case MsgBoxResult.Cancel
                Exit Sub
        End Select
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        gStrSql = "Select FolioAlmacen FROM MovtosAlmacenCab WHERE ReferenciadeOrigen = '" & txtFolioSalida.Text & "' " & "AND Estatus = 'V' AND CodMovtoAlm = " & C_SalidadeAlmacendeVendedorExterno
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            FolioSalida = RsGral.Fields("FolioAlmacen").Value
        End If
        'Cancelar los Movimientos de Cabecero
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(txtFolio.Text, "01/01/1900", txtCodSucMatriz.Text, "0", "", "0", "", "0", "0", "", "", "", "", "", "C", "", VB6.Format(Today, C_FORMATFECHAGUARDAR), gStrNomUsuario, "", "01/01/1900", "0", "", "01/01/1900", "0", "", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(FolioSalida, "01/01/1900", txtCodSucVendExterno.Text, "0", "", "0", "", "0", "0", "", "", "", "", "", "C", "", VB6.Format(Today, C_FORMATFECHAGUARDAR), gStrNomUsuario, "", "01/01/1900", "0", "", "01/01/1900", "0", "", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        'Cancelar los Detalle
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" Then
                    'And Numerico(.TextMatrix(I, 4)) <> 0
                    'Cancelar el Detalle de Salida
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(txtFolio.Text, "0", "01/01/1900", .get_TextMatrix(I, 0), "0", "0", "0", "0", "0", "C", VB6.Format(Today, C_FORMATFECHAGUARDAR), "0", C_ELIMINACION, CStr(0))
                    Cmd.Execute()
                    'Cancelar el Detalle de Entrada
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(FolioSalida, "0", "01/01/1900", .get_TextMatrix(I, 0), "0", "0", "0", "0", "0", "C", VB6.Format(Today, C_FORMATFECHAGUARDAR), "0", C_ELIMINACION, CStr(0))
                    Cmd.Execute()

                    'Guardar el Detalle de Inventario de Entrada
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucMatriz.Text, "1", .get_TextMatrix(I, 0), txtCodSucMatriz.Text, "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 6))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 6), "0", CStr(Numerico(.get_TextMatrix(I, 3))), "0", CStr(C_EntradaPorDevoluciondeVendedoresExternos), VB6.Format(Today, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Guardar el Detalle de Inventario de Salida
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucVendExterno.Text, "0", .get_TextMatrix(I, 0), txtCodSucMatriz.Text, "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 6))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 6), CStr(Numerico(.get_TextMatrix(I, 3))), "0", "0", CStr(C_SalidadeAlmacendeVendedorExterno), VB6.Format(Today, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Han sido cancelados con éxito el folío de recepción de mercancía de vend. externo " & txtFolio.Text & Chr(13) & "Y el folío de entrega de mercancia de vendedor externo " & FolioSalida, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Limpiar()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function BuscaLiquidacion() As Boolean
        On Error GoTo Merr
        gStrSql = "SELECT * FROM MovtosAlmacenCab WHERE ReferenciadeOrigen = '" & txtFolioSalida.Text & "' AND CodMovtoAlm = " & C_SalidaPorVentadeVendedoresExternos & " " & "AND Estatus = 'V'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            BuscaLiquidacion = True
        Else
            BuscaLiquidacion = False
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function Guardar() As Boolean
        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        Dim Consecutivo As Integer
        Dim FolioEntrada As String
        Dim I As Integer
        Dim NumPartida As Integer
        If Not mblnNuevo Then
            Exit Function
        End If
        If ValidaDatos() = False Then
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Generar el Folio de Entrada por Devolución de Vendedores Externos
        ModStoredProcedures.PR_I_FoliosAlmacen(txtCodSucMatriz.Text, CStr(Consecutivo), "", CStr(0))
        Cmd.Execute()
        Consecutivo = Cmd.Parameters("Consecutivo").Value
        txtFolio.Text = C_PrefijoFoliosAlmacen & VB6.Format(txtCodSucMatriz.Text, "00") & Year(dtpFecha.Value) & VB6.Format(Month(dtpFecha.Value), "00") & VB6.Format((dtpFecha.Value), "00") & VB6.Format(Consecutivo, "000000")
        'Generar el Folio de Salida de Almacen de Vendedor Externo
        ModStoredProcedures.PR_I_FoliosAlmacen(txtCodSucVendExterno.Text, CStr(Consecutivo), " ", CStr(0))
        Cmd.Execute()
        Consecutivo = Cmd.Parameters("Consecutivo").Value
        FolioEntrada = C_PrefijoFoliosAlmacen & VB6.Format(txtCodSucVendExterno.Text, "00") & Year(dtpFecha.Value) & VB6.Format(Month(dtpFecha.Value), "00") & VB6.Format((dtpFecha.Value), "00") & VB6.Format(Consecutivo, "000000")
        'Guardar el Movimiento de Cabecero de Entrada por Devolución de Vendedores Externos
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(txtFolio.Text, VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), txtCodSucMatriz.Text, "0", "", "0", "", txtCodSucVendExterno.Text, CStr(C_EntradaPorDevoluciondeVendedoresExternos), C_ENTRADA, txtEnvia.Text, txtEntrega.Text, txtRecibe.Text, QuitaEnter(txtConcepto.Text), "V", gStrNomUsuario, "01/01/1900", "", txtFolioSalida.Text, VB6.Format(Fecha, C_FORMATFECHAGUARDAR), "0", "", "01/01/1900", CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar el Movimiento de Cabecero de Salida de Almacen de Vendedor Externo
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(FolioEntrada, VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), txtCodSucVendExterno.Text, "0", "", "0", "", txtCodSucMatriz.Text, CStr(C_SalidadeAlmacendeVendedorExterno), C_SALIDA, txtEnvia.Text, txtEntrega.Text, txtRecibe.Text, "SALIDA DE ALMACEN DE VENDEDOR EXTERNO " & txtDescVendExterno.Text, "V", gStrNomUsuario, "01/01/1900", "", txtFolioSalida.Text, VB6.Format(Fecha, C_FORMATFECHAGUARDAR), "0", "", "01/01/1900", CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar los Detalles de Entrada y Salida
        NumPartida = 1
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And CDbl(Numerico(.get_TextMatrix(I, 4))) <> 0 Then
                    '''                'Guarda el Detalle de Entrada
                    '''                ModStoredProcedures.PR_IE_MovtosAlmacenDet txtFolio, CStr(NumPartida), Format(dtpFecha, C_FORMATFECHAGUARDAR), _
                    ''''                .TextMatrix(I, 0), "0", .TextMatrix(I, 4), .TextMatrix(I, 6), CStr(CCur(Numerico(.TextMatrix(I, 5))) / (1 + Round(gcurCorpoTASAIVA / 100, 2))), _
                    ''''                "0", "V", "01/01/1900", "0", C_INSERCION, 0
                    '''                Cmd.Execute
                    '''                'Guarda el Detalle de Salida
                    '''                ModStoredProcedures.PR_IE_MovtosAlmacenDet FolioEntrada, CStr(NumPartida), Format(dtpFecha, C_FORMATFECHAGUARDAR), _
                    ''''                .TextMatrix(I, 0), "0", .TextMatrix(I, 4), .TextMatrix(I, 6), CStr(CCur(Numerico(.TextMatrix(I, 5))) / (1 + Round(gcurCorpoTASAIVA / 100, 2))), _
                    ''''                "0", "V", "01/01/1900", "0", C_INSERCION, 0
                    '''                Cmd.Execute

                    '''Se elimino el precio de venta - elimino division del preciopub/tasaiva
                    'Guarda el Detalle de Entrada
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(txtFolio.Text, CStr(NumPartida), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), .get_TextMatrix(I, 4), .get_TextMatrix(I, 6), CStr(CDec(Numerico(.get_TextMatrix(I, 5)))), "0", "V", "01/01/1900", "0", C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Guarda el Detalle de Salida
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(FolioEntrada, CStr(NumPartida), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), .get_TextMatrix(I, 4), .get_TextMatrix(I, 6), CStr(CDec(Numerico(.get_TextMatrix(I, 5)))), "0", "V", "01/01/1900", "0", C_INSERCION, CStr(0))
                    Cmd.Execute()

                    '''Determinar el almacen origen del articulo - es el dato correcto
                    '''tenia: txtCodSucMatriz - esto se refiere al almacen que lo genera
                    '''20OCT2004

                    'Guardar el Detalle de Inventario de Entrada
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucMatriz.Text, "1", .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 6))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 6), .get_TextMatrix(I, 4), "0", "0", CStr(C_EntradaPorDevoluciondeVendedoresExternos), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Guardar el Detalle de Inventario de Salida
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucVendExterno.Text, "0", .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 6))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 6), "0", .get_TextMatrix(I, 4), "0", CStr(C_SalidadeAlmacendeVendedorExterno), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Los datos se han guardado con éxito" & vbNewLine & "Se han generado los siguientes folios  : " & vbNewLine & vbNewLine & "Folio de recepción por devolución             : " & txtFolio.Text & vbNewLine & "Folio de entrega de mcia. de vendedor externo : " & FolioEntrada, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        If MsgBox("¿Desea imprimir el comprobante de recepción de mercancia?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, gstrNombCortoEmpresa) = MsgBoxResult.Yes Then
            ImprimirComprobante()
        End If
        Limpiar()

Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            Select Case flexDetalle.Col
                Case 4 'Cantidad de Articulos a Devolver
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 5
            End Select
        End With
    End Sub

    Function ChecaGrid() As Boolean
        Dim I As Integer
        Dim TotArt As Integer

        TotArt = 0
        With flexDetalle
            For I = 1 To .Rows - 1
                '''se acabaron las partidas
                If Trim(.get_TextMatrix(I, 0)) = "" And Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "" And Trim(.get_TextMatrix(I, 3)) = "" Then
                    ChecaGrid = False
                    Exit For
                End If
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 7)) = "" Then
                    ChecaGrid = False
                    MsgBox("El artículo: " & .get_TextMatrix(I, 0) & " no tiene" & vbNewLine & "código del almacen origen" & vbNewLine & "No se puede guardar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Function
                End If
                TotArt = TotArt + CInt(Numerico(.get_TextMatrix(I, 4)))
            Next
        End With

        If TotArt = 0 Then
            ChecaGrid = False
        Else
            ChecaGrid = True
        End If
    End Function

    Sub Encabezado()
        With flexDetalle
            .set_Cols(0, 8)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .set_ColWidth(0, 0, 1500)
            .CellFontBold = True
            .Text = "Código"
            .Col = 1
            .CellAlignment = 5
            .set_ColWidth(1, 0, 4800)
            .CellFontBold = True
            .Text = "Descripción"
            .Col = 2
            .CellAlignment = 5
            .set_ColWidth(2, 0, 700)
            .CellFontBold = True
            .Text = "Unidad"
            .Col = 3
            .CellAlignment = 5
            .set_ColWidth(3, 0, 900)
            .CellFontBold = True
            .Text = "Cantidad"
            .Col = 4
            .CellAlignment = 5
            .set_ColWidth(4, 0, 1100)
            .CellFontBold = True
            .Text = "Devolución"
            .Col = 5
            .set_ColWidth(5, 0, 0)
            .Col = 6
            .set_ColWidth(6, 0, 0)
            .Col = 7
            .set_ColWidth(7, 0, 0)
            .set_ColAlignment(0, 7)
            .Rows = 11
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub EncabezadoComprobante(ByRef Empresa As String, ByRef Pagina As Integer)
        On Error GoTo Merr
        With Printer
            .ScaleMode = vbMillimeters
            .FontName = "Courier New"
            .Orientation = 1
            .Height = 140 'Cambia el tamaño de la hoja en la impresora
            .FontSize = 16
        End With
        Printer.CurrentX = 65
        Printer.CurrentY = 10
        Printer.Print(Empresa)
        Printer.FontSize = 10
        Printer.CurrentX = 65
        Printer.CurrentY = 15
        Printer.Print("Recepción de Mercancia a Vend. Externo")
        Printer.CurrentX = 5

        Printer.CurrentY = 20


        Printer.Print("Folio de Recepción : " & txtFolio.Text)

        Printer.CurrentX = 160

        Printer.CurrentY = 20


        Printer.Print("Pagina : " & (Space(5) & Pagina))

        Printer.CurrentX = 5

        Printer.CurrentY = 25


        Printer.Print("Vendedor :      " & txtDescVendExterno.Text)

        Printer.CurrentX = 160

        Printer.CurrentY = 25


        Printer.Print("Fecha : " & Format(dtpFecha.Value, "dd/mmm/yyyy"))

        Printer.CurrentX = 5

        Printer.CurrentY = 30


        Printer.Print("============================================================================================")

        Printer.CurrentX = 7

        Printer.CurrentY = 35


        Printer.Print("Código     Descripción                                                             Cantidad ")

        Printer.CurrentX = 5

        Printer.CurrentY = 40


        Printer.Print("============================================================================================")
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Imprime()
        ImprimirComprobante()
    End Sub

    Sub ImprimirComprobante()
        On Error GoTo Merr
        Dim strImpresora As String
        Dim strNomEmpresa As String
        Dim I As Integer
        Dim Salto As Boolean
        Dim NumPagina As Integer
        Dim Linea As Integer
        Dim TotalArt As Integer
        Dim NumPartida As Integer
        Salto = False
        NumPagina = 1
        TotalArt = 0
        ModCorporativo.CargarRutaImpresoras()
        strImpresora = gstrRutaImpresora
        If Not ModEstandar.BuscarImpresora(strImpresora) Then
            MsgBox("Impresora Incorrecta")
        End If
        strNomEmpresa = UCase(gstrNombCortoEmpresa)
        For I = 1 To Len(strNomEmpresa)
            Select Case Mid(strNomEmpresa, I, 1)
                Case "Á"
                    strNomEmpresa = Replace(strNomEmpresa, "Á", "A")
                Case "É"
                    strNomEmpresa = Replace(strNomEmpresa, "É", "E")
                Case "Í"
                    strNomEmpresa = Replace(strNomEmpresa, "Í", "I")
                Case "Ó"
                    strNomEmpresa = Replace(strNomEmpresa, "Ó", "O")
                Case "Ú"
                    strNomEmpresa = Replace(strNomEmpresa, "Ú", "U")
            End Select
        Next
        EncabezadoComprobante(strNomEmpresa, NumPagina)
        Linea = 44
        NumPartida = 1
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "" Then
                    If Salto Then

                        Printer.NewPage()
                        EncabezadoComprobante(strNomEmpresa, NumPagina)
                        Linea = 44
                        Salto = False
                    End If

                    Printer.CurrentX = 5

                    Printer.CurrentY = Linea


                    Printer.Print((Space(9) & .get_TextMatrix(I, 0)))

                    Printer.CurrentX = 30

                    Printer.CurrentY = Linea


                    Printer.Print(.get_TextMatrix(I, 1))

                    Printer.CurrentX = 185

                    Printer.CurrentY = Linea
                    If mblnNuevo Then


                        Printer.Print((Space(6) & .get_TextMatrix(I, 4)))
                        TotalArt = TotalArt + CShort(Numerico(.get_TextMatrix(I, 4)))
                    Else


                        Printer.Print((Space(6) & .get_TextMatrix(I, 3)))
                        TotalArt = TotalArt + CShort(Numerico(.get_TextMatrix(I, 3)))
                    End If
                    Linea = Linea + 4
                    If Linea > 250 Then
                        Salto = True
                        NumPagina = NumPagina + 1
                    End If
                    NumPartida = NumPartida + 1
                End If
            Next
            If Linea >= 250 Then
                Linea = 44
                NumPagina = NumPagina + 1

                Printer.NewPage()
                EncabezadoComprobante(strNomEmpresa, NumPagina)
            End If

            Printer.CurrentX = 175

            Printer.CurrentY = Linea


            Printer.Print("------------")
            Linea = Linea + 4

            Printer.CurrentX = 145

            Printer.CurrentY = Linea


            Printer.Print("Total de Artículos")

            Printer.CurrentX = 185

            Printer.CurrentY = Linea


            Printer.Print((Space(6) & TotalArt))
            Linea = Linea + 10
            '        If Linea >= 250 Then
            '            Linea = 44
            '            NumPagina = NumPagina + 1
            '            Printer.NewPage
            '            EncabezadoComprobante strNomEmpresa, NumPagina
            '        End If
            If NumPartida <= 15 Then

                Printer.CurrentX = 15

                Printer.CurrentY = 115


                Printer.Print("Entregó")

                Printer.CurrentX = 150

                Printer.CurrentY = 115


                Printer.Print("Recibi Artículos")

                Printer.CurrentX = 15

                Printer.CurrentY = 120


                Printer.Print(txtDescVendExterno.Text)

                Printer.CurrentX = 150

                Printer.CurrentY = 120


                Printer.Print(txtRecibe.Text)
            Else

                Printer.CurrentX = 15

                Printer.CurrentY = 250


                Printer.Print("Entregó")

                Printer.CurrentX = 150

                Printer.CurrentY = 250


                Printer.Print("Recibi Artículos")
                Printer.CurrentX = 15

                Printer.CurrentY = 255


                Printer.Print(txtDescVendExterno.Text)

                Printer.CurrentX = 150

                Printer.CurrentY = 255


                Printer.Print(txtRecibe.Text)
            End If
        End With
        Printer.EndDoc()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
        mblnNuevo = True
        mblnCambios = False
    End Sub

    Sub Limpiar()
        Nuevo()
        txtFolio.Text = ""
        txtFolio.Focus()
    End Sub

    Sub LlenaDatos()

        If (bandera = True) Then
            Exit Sub
        End If

        On Error GoTo Merr
        Dim I As Integer
        If Trim(txtFolio.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT CAB.FOLIOALMACEN,CAB.FECHAALMACEN,ALM.DESCALMACEN,CAB.CODALMACENREF,CAB.ENVIA,CAB.ENTREGA,CAB.RECIBE,CAB.CONCEPTO,CAB.REFERENCIADEORIGEN," & "DET.CodArticulo , ART.descArticulo,ART.CostoReal,ART.PrecioPubDolar,uni.DESCUNIDAD,DET.Cantidad " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "INNER JOIN CATALMACEN ALM ON CAB.CODALMACENREF = ALM.CODALMACEN " & "INNER JOIN CATARTICULOS ART ON DET.CODARTICULO = ART.CODARTICULO " & "INNER JOIN CATUNIDADES UNI ON ART.CODUNIDAD = UNI.CODUNIDAD " & "WHERE CAB.CodAlmacen = " & txtCodSucMatriz.Text & " AND CAB.FolioAlmacen = '" & txtFolio.Text & "' AND CAB.CodMovtoAlm = " & C_EntradaPorDevoluciondeVendedoresExternos & " " & "AND CAB.Estatus = 'V'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFolioSalida.Text = RsGral.Fields("ReferenciaDeOrigen").Value
            Fecha = CDate(Format(RsGral.Fields("FechaAlmacen").Value, C_FORMATFECHAMOSTRAR))
            txtCodSucVendExterno.Text = RsGral.Fields("CodALmacenREf").Value
            txtDescVendExterno.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            txtEnvia.Text = RsGral.Fields("Envia").Value
            txtRecibe.Text = RsGral.Fields("Recibe").Value
            txtEntrega.Text = RsGral.Fields("Entrega").Value
            txtConcepto.Text = RsGral.Fields("Concepto").Value
            I = 1
            Do While Not RsGral.EOF
                With flexDetalle
                    .set_TextMatrix(I, 0, RsGral.Fields("CodArticulo").Value)
                    .set_TextMatrix(I, 1, Trim(RsGral.Fields("DescArticulo").Value))
                    .set_TextMatrix(I, 2, Trim(RsGral.Fields("DescUnidad").Value))
                    .set_TextMatrix(I, 3, RsGral.Fields("Cantidad").Value)
                    .set_TextMatrix(I, 4, 0)
                    .set_TextMatrix(I, 5, RsGral.Fields("PrecioPubDolar").Value)
                    .set_TextMatrix(I, 6, RsGral.Fields("CostoReal").Value)
                    RsGral.MoveNext()
                    If Not RsGral.EOF Then
                        If .Rows - 1 = I Then
                            .Rows = .Rows + 1
                        End If
                        I = I + 1
                    End If
                End With
            Loop
            flexDetalle.set_ColWidth(2, 0, 1200)
            flexDetalle.set_ColWidth(3, 0, 1500)
            flexDetalle.set_ColWidth(4, 0, 0)
            mblnNuevo = False
            txtFolioSalida.Enabled = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            txtEnvia.Enabled = False
            txtEntrega.Enabled = False
            txtRecibe.Enabled = False
            ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Else
            MsgBox("Folio de almacén de recepción de mercancia de vendedor externo no existe" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtFolioSalida.Enabled = True
            txtFolio.Text = ""
            txtFolio.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatosFolioSalida()
        On Error GoTo Merr
        Dim I As Integer
        Dim RsExis As ADODB.Recordset
        Dim NumPart As Integer
        Dim Existencia As Integer
        Dim ExistenciaDisponible As Integer
        Dim CodSucursal As Integer
        NumPart = 0
        If Trim(txtFolioSalida.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacenRef FROM MOVTOSALMACENCAB WHERE FOLIOALMACEN = '" & txtFolioSalida.Text & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            CodSucursal = RsGral.Fields("CodALmacenREf").Value
        End If
        '    gStrSql = "SELECT ReferenciaDeOrigen FROM MovtosAlmacenCab WHERE ReferenciadeOrigen = '" & txtFolioSalida & "' AND " & _
        ''    "CodMovtoAlm = " & C_EntradaPorDevoluciondeVendedoresExternos & " AND Estatus = 'V'"
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        '    Set RsGral = Cmd.Execute
        '    If RsGral.RecordCount > 0 Then
        '        MsgBox "Ya se capturó una entrada por devolución de mercancía para este folio" & vbNewLine & "Favor de verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '        txtFolioSalida = ""
        '        txtFolioSalida.SetFocus
        '        Exit Sub
        '    End If
        gStrSql = "SELECT CAB.FOLIOALMACEN,CAB.FECHAALMACEN,ALM.DESCALMACEN,CAB.CODALMACENREF," & "DET.CodArticulo , ART.descArticulo,ART.CostoReal,ART.PrecioPubDolar,uni.DESCUNIDAD,DET.Cantidad,ISNULL(ENTRADAS.CANTIDAD,0) AS CANTIDADDEVOL,ISNULL(VENTAS.CANTIDAD,0) AS CANTIDADVEN,INV.EXISTENCIA, ART.CodAlmacenOrigen " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "LEFT OUTER JOIN " & "(SELECT CAB.REFERENCIADEORIGEN,DET.CODARTICULO,SUM(DET.CANTIDAD) AS CANTIDAD " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "WHERE CAB.CODMOVTOALM = " & C_SalidadeAlmacendeVendedorExterno & " AND CAB.ESTATUS <> 'C' GROUP BY CAB.REFERENCIADEORIGEN,DET.CODARTICULO) ENTRADAS " & "ON CAB.FOLIOALMACEN = ENTRADAS.REFERENCIADEORIGEN AND DET.CODARTICULO = ENTRADAS.CODARTICULO " & "LEFT OUTER JOIN " & "(SELECT CAB.REFERENCIADEORIGEN,DET.CODARTICULO,SUM(DET.CANTIDAD) AS CANTIDAD " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "WHERE CAB.CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " AND CAB.ESTATUS <> 'C' GROUP BY CAB.REFERENCIADEORIGEN,DET.CODARTICULO) VENTAS " & "ON CAB.FOLIOALMACEN = VENTAS.REFERENCIADEORIGEN AND DET.CODARTICULO = VENTAS.CODARTICULO " & "LEFT OUTER JOIN " & "(SELECT CODARTICULO,ISNULL(SUM((EXISTENCIAINICIAL + ENTRADAS) - (SALIDAS + APARTADOS)),0) AS EXISTENCIA " & "FROM INVENTARIO WHERE CODALMACEN = " & CodSucursal & " GROUP BY CODARTICULO) INV ON DET.CODARTICULO = INV.CODARTICULO " & "INNER JOIN CATALMACEN ALM ON CAB.CODALMACENREF = ALM.CODALMACEN " & "INNER JOIN CATARTICULOS ART ON DET.CODARTICULO = ART.CODARTICULO " & "INNER JOIN CATUNIDADES UNI ON ART.CODUNIDAD = UNI.CODUNIDAD " & "WHERE CAB.CodAlmacen = " & txtCodSucMatriz.Text & " AND CAB.FolioAlmacen = '" & txtFolioSalida.Text & "' AND CAB.CodMovtoAlm = " & C_SalidaAVendedoresExternos & " AND CAB.Estatus <> 'C'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Fecha = CDate(Format(RsGral.Fields("FechaAlmacen").Value, C_FORMATFECHAMOSTRAR))
            txtCodSucVendExterno.Text = RsGral.Fields("CodALmacenREf").Value
            txtDescVendExterno.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            I = 1
            Do While Not RsGral.EOF
                If RsGral.Fields("Existencia").Value > 0 Then
                    Existencia = RsGral.Fields("Cantidad").Value - (RsGral.Fields("CantidadDevol").Value + RsGral.Fields("CantidadVen").Value)
                    ExistenciaDisponible = RsGral.Fields("Existencia").Value
                    '                If RsGral!CantidadDevol = 0 Then
                    '                    If ExistenciaDisponible > RsGral!Cantidad Then
                    '                        Existencia = RsGral!Cantidad
                    '                    ElseIf ExistenciaDisponible <= RsGral!Cantidad Then
                    '                        Existencia = ExistenciaDisponible
                    '                    End If
                    '                End If
                    If Existencia > 0 Then
                        With flexDetalle
                            .set_TextMatrix(I, 0, RsGral.Fields("CodArticulo").Value)
                            .set_TextMatrix(I, 1, Trim(RsGral.Fields("DescArticulo").Value))
                            .set_TextMatrix(I, 2, Trim(RsGral.Fields("DescUnidad").Value))
                            .set_TextMatrix(I, 3, Existencia)
                            .set_TextMatrix(I, 4, 0)
                            .set_TextMatrix(I, 5, RsGral.Fields("PrecioPubDolar").Value)
                            .set_TextMatrix(I, 6, RsGral.Fields("CostoReal").Value)
                            .set_TextMatrix(I, 7, RsGral.Fields("CodAlmacenOrigen").Value)
                            NumPart = NumPart + 1
                            RsGral.MoveNext()
                            If Not RsGral.EOF Then
                                If .Rows - 1 = I Then
                                    .Rows = .Rows + 1
                                End If
                                I = I + 1
                            End If
                        End With
                    Else
                        RsGral.MoveNext()
                    End If
                Else
                    RsGral.MoveNext()
                End If

            Loop
            If NumPart = 0 Then
                MsgBox("Los articulos de este folio de salida, ya no existen en el inventario de este vendedor externo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Limpiar()
                Exit Sub
            End If
            txtFolioSalida.Enabled = False
        Else
            MsgBox("El folio de almacén no existe" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtFolioSalida.Text = ""
            txtFolioSalida.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Nuevo()
        txtFolioSalida.Text = ""
        dtpFecha.Value = Today
        txtCodSucVendExterno.Text = ""
        txtDescVendExterno.Text = ""
        txtEnvia.Text = ""
        txtEntrega.Text = ""
        txtRecibe.Text = ""
        txtConcepto.Text = "RECEPCION DE MERCANCIA DE VENDEDOR EXTERNO " & txtCodSucVendExterno.Text & " " & txtDescVendExterno.Text
        flexDetalle.Clear()
        Encabezado()
        txtFolioSalida.Enabled = True
        Frame4.Enabled = True
        Frame5.Enabled = True
        txtEnvia.Enabled = True
        txtEntrega.Enabled = True
        txtRecibe.Enabled = True
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        InicializaVariables()
    End Sub

    Function ValidarDevolucion() As Boolean
        With flexDetalle
            If (CInt(flexDetalle.get_TextMatrix(flexDetalle.Row, 3)) - CInt(txtFlex.Text)) < 0 Then
                MsgBox("No es posible regresar mas de lo entregado, Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidarDevolucion = False
            Else
                ValidarDevolucion = True
            End If
        End With
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Trim(txtFolioSalida.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Folio de salida ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtFolioSalida.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtCodSucVendExterno.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Código de vendedor externo ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodSucVendExterno.Focus()
            Exit Function
        End If
        If Trim(txtDescVendExterno.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Nombre de la sucursal ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtDescVendExterno.Focus()
            Exit Function
        End If
        If Trim(txtEnvia.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Quien envía la mercancía ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtEnvia.Focus()
            Exit Function
        End If
        If Trim(txtEntrega.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Quien entrega la mercancía ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtEntrega.Focus()
            Exit Function
        End If
        If Trim(txtRecibe.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Quien recibe la mercancía ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtRecibe.Focus()
            Exit Function
        End If
        If Not ChecaGrid() Then
            MsgBox("No hay artículos por devolver" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            flexDetalle.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub flexDetalle_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.ClickEvent
        txtFlex.Visible = False
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        FlexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub FlexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalle.KeyPressEvent
        Dim lonR, lonI As Integer
        Dim EsEnter As Boolean
        EsEnter = False
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And mblnNuevo Then
            If eventArgs.keyAscii = System.Windows.Forms.Keys.Return Then EsEnter = True
            'Verifica si se puede capturar la fila
            '        If flexDetalle.Row > 1 Then
            '            If flexDetalle.TextMatrix(flexDetalle.Row - 1, 0) <> "" Then
            '                For lonR = 1 To flexDetalle.Row - 1 Step 1
            '                    For lonI = 0 To 4 Step 1
            '                        If flexDetalle.TextMatrix(lonR, lonI) = "" Then
            '                            'MsgBox "Hace falta información en la captura", vbExclamation, cNomEmp
            '                            flexDetalle.Row = lonR
            '                            flexDetalle.Col = lonI
            '                            CambiarFormatoTxtenCaptura
            '                            MSHFlexGridEdit flexDetalle, txtFlex, KeyAscii
            '                            Exit Sub
            '                        End If
            '                    Next lonI
            '                Next lonR
            '            Else
            '                'flexDetalle.SetFocus
            '                Exit Sub
            '            End If
            '        End If
            'Edita el campo sólo si es Editable
            If flexDetalle.Col = 4 Then
                If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" And Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 1)) = "" And Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2)) = "" And Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 3)) = "" Then
                    Exit Sub
                Else
                    gStrSql = "SELECT SUM((EXISTENCIAINICIAL + ENTRADAS) - (SALIDAS + APARTADOS)) AS EXISTENCIA " & "FROM INVENTARIO WHERE CODARTICULO = " & Numerico(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) & " AND CODALMACEN = " & txtCodSucVendExterno.Text
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        If RsGral.Fields("Existencia").Value = 0 Then
                            MsgBox("Este articulo no tiene existencias para este almacen de vend. externo, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            flexDetalle.Focus()
                            Exit Sub
                        End If
                    End If
                    eventArgs.keyAscii = ModEstandar.MskCantidad(txtFlex.Text, eventArgs.keyAscii, 5, 0, (txtFlex.SelectionStart))
                    CambiarFormatoTxtenCaptura()
                    MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                    'If Len(Trim(txtFlex)) = 1 Then
                    If Not EsEnter Then
                        System.Windows.Forms.SendKeys.Send("{right}")
                    End If
                    'End If
                End If
            End If
        End If
    End Sub

    Private Sub frmVtasVEEntradadeMercancia_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVEEntradadeMercancia_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVEEntradadeMercancia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "txtFolio" And Trim(txtFolio.Text) <> "" And (txtFolio.Text) <> "000000" Then
                    txtFolioSalida.Enabled = False
                End If
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolio" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasVEEntradadeMercancia_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVEEntradadeMercancia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        bandera = True
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        Nuevo()
        ObtenerDatosSucursalMatriz(txtCodSucMatriz, txtDescSucMatriz)
    End Sub

    Private Sub frmVtasVEEntradadeMercancia_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        '    'If Cambios = True And mblnNuevo = False Then
        '    'Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '    'Case vbYes: 'Guardar el registro
        '    'If Guardar = False Then
        '    'Cancel = 1
        '    'End If
        '    'Case vbNo: 'No hace nada y permite el cierre del formulario
        '    'Case vbCancel: 'Cancela el cierre del formulario sin guardar
        '    'Cancel = 1
        '    'End Select
        '    'End If
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

    Private Sub frmVtasVEEntradadeMercancia_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtCodSucMatriz_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucMatriz.Enter
        Pon_Tool()
        SelTextoTxt(txtCodSucMatriz)
    End Sub

    Private Sub txtCodSucVendExterno_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucVendExterno.Enter
        Pon_Tool()
        SelTextoTxt(txtCodSucVendExterno)
    End Sub

    Private Sub txtConcepto_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.TextChanged
        If Trim(txtConcepto.Text) = "" Then
            txtConcepto.Text = "RECEPCION DE MERCANCIA DE VENDEDOR EXTERNO " & txtCodSucVendExterno.Text & " " & txtDescVendExterno.Text
        End If
        txtConcepto.Text = Trim(QuitaEnter(txtConcepto.Text))
    End Sub

    Private Sub txtConcepto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Enter
        Pon_Tool()
        txtConcepto.Text = "RECEPCION DE MERCANCIA DE VENDEDOR EXTERNO " & txtCodSucVendExterno.Text & " " & txtDescVendExterno.Text
        SelTextoTxt(txtConcepto)
    End Sub

    Private Sub txtConcepto_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Leave
        QuitaEnter(txtConcepto.Text)
    End Sub

    Private Sub txtDescSucMatriz_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescSucMatriz.Enter
        Pon_Tool()
        SelTextoTxt(txtDescSucMatriz)
    End Sub

    Private Sub txtDescVendExterno_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescVendExterno.Enter
        Pon_Tool()
    End Sub

    Private Sub txtEntrega_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEntrega.Enter
        Pon_Tool()
        SelTextoTxt(txtEntrega)
    End Sub

    Private Sub txtEnvia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEnvia.Enter
        Pon_Tool()
        SelTextoTxt(txtEnvia)
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        Pon_Tool()
        SelTextoTxt(txtFlex)
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        With flexDetalle
            If KeyCode = System.Windows.Forms.Keys.Return Then
                Select Case .Col
                    Case 4
                        If .Col = 4 And Trim(txtFlex.Text) <> "" Then
                            .Text = Trim(Numerico(txtFlex.Text))
                            If ValidarDevolucion() Then
                                txtFlex.Visible = False
                                Exit Sub
                            Else
                                .Text = ""
                                txtFlex.Text = ""
                                txtFlex.Focus()
                                Exit Sub
                            End If
                        ElseIf .Col = 4 And Trim(txtFlex.Text) = "" Then
                            .Text = "0"
                            txtFlex.Visible = False
                        End If
                        txtFlex.Visible = False
                End Select
            ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
                .Focus()
                txtFlex.Visible = False
            End If
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case Else
                Select Case flexDetalle.Col
                    Case 4
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                End Select
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub txtFolio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambios = True
    End Sub

    Private Sub txtFolio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Enter
        strControlActual = UCase("txtFolio")
        Pon_Tool()
        SelTextoTxt(txtFolio)
    End Sub

    Private Sub txtFolio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_PrefijoFoliosAlmacen)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Trim(txtFolio.Text) = "" Then
            'txtFolio.Text = C_PrefijoFoliosAlmacen & Format(txtCodSucMatriz.Text, "00") & Year(dtpFecha.Value) & Format(Month(dtpFecha.Value), "00") & Format((dtpFecha.Value), "00") & "000000"
            txtFolio.Text = C_PrefijoFoliosAlmacen & String.Concat("0" + txtCodSucMatriz.Text) & Year(dtpFecha.Value) & Format(Month(dtpFecha.Value), "00") & Format(dtpFecha.Value.Day, "00") & "000000"

        End If
        If mblnCambios = True And txtFolio.Text <> "" And (txtFolio.Text) <> "000000" Then
            'flexDetalle.Clear()
            Encabezado()
            LlenaDatos()
        End If
    End Sub

    Private Sub txtFolioSalida_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioSalida.Enter
        strControlActual = UCase("txtFolioSalida")
        Pon_Tool()
    End Sub

    Private Sub txtFolioSalida_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioSalida.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_PrefijoFoliosAlmacen)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioSalida_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioSalida.Leave
        If Trim(txtFolioSalida.Text) = "" Then
            Exit Sub
        Else
            flexDetalle.Clear()
            Encabezado()
            LlenaDatosFolioSalida()
        End If
    End Sub

    Private Sub txtRecibe_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRecibe.Enter
        Pon_Tool()
        SelTextoTxt(txtRecibe)
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

    End Sub
End Class
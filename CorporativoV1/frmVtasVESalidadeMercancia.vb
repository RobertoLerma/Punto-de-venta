Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6


Public Class frmVtasVESalidadeMercancia
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             SALIDA DE MERCANCIA                                                                          *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MARTES 19 DE AGOSTO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents dbcDescripcion As System.Windows.Forms.ComboBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents txtRecibe As System.Windows.Forms.TextBox
    Public WithEvents txtEntrega As System.Windows.Forms.TextBox
    Public WithEvents txtEnvia As System.Windows.Forms.TextBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents txtCodSucVendExterno As System.Windows.Forms.TextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents txtDescSucMatriz As System.Windows.Forms.TextBox
    Public WithEvents txtCodSucMatriz As System.Windows.Forms.TextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents txtFolio As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label8 As System.Windows.Forms.Label


    'Variables
    Dim mblnSalir As Boolean
    Dim mblnNuevo As Boolean
    Dim mblnCambios As Boolean
    Dim FueraChange As Boolean
    Dim intCodSucursal As Integer
    Dim intCodArticulo As Integer
    Dim tecla As Integer
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnBuscar As Button
    Dim blnBuscar As Boolean
    Dim strControlActual As String 'Nombre del control actual
    Public bandera As Boolean = False
    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        If strControlActual = "FLEXDETALLE" And mblnNuevo Then
            If flexDetalle.Col = 1 Or flexDetalle.Col = 2 Or flexDetalle.Col = 3 Then
                Exit Sub
            Else
                If flexDetalle.get_TextMatrix(flexDetalle.Row - 1, 3) = "" Then
                    FlexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
                    Exit Sub
                End If
            End If
        End If

        Select Case strControlActual
            Case "TXTFLEX", "FLEXDETALLE"
                If Not mblnNuevo Then
                    Exit Sub
                End If
                strCaptionForm = "Busqueda de Articulos"
                gStrSql = "SELECT RIGHT('       '+LTRIM(CodArticulo),7) AS CODIGO,RTRIM(DescArticulo) AS DESCRIPCION " & "FROM CatArticulos ORDER BY CodArticulo"
            Case "TXTCODSUCVENDEXTERNO"
                strCaptionForm = "Busqueda de Vendedores Externos"
                gStrSql = "SELECT RIGHT('000'+LTRIM(CodAlmacen),3) AS CODIGO,DescAlmacen AS DESCRIPCION " & "From CatAlmacen WHERE TipoAlmacen = 'V' ORDER BY CodAlmacen"
            Case "TXTFOLIO"
                strCaptionForm = "Busqueda de Folios de Entrega de Mercancia"
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
                Case "TXTCODSUCVENDEXTERNO"
                    'ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTFLEX", "FLEXDETALLE"
                    blnBuscar = False
                    'ConfiguraConsultas(FrmConsultas, 8350, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 1000)
                    .set_ColWidth(1, 0, 7000)
                    FrmConsultas.Left = VB6.TwipsToPixelsX(3200)
                Case "TXTFOLIO"
                    'ConfiguraConsultas(FrmConsultas, 11000, RsGral, strTag, strCaptionForm)
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasVESalidadeMercancia))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.txtRecibe = New System.Windows.Forms.TextBox()
        Me.txtEntrega = New System.Windows.Forms.TextBox()
        Me.txtEnvia = New System.Windows.Forms.TextBox()
        Me.txtCodSucVendExterno = New System.Windows.Forms.TextBox()
        Me.txtDescSucMatriz = New System.Windows.Forms.TextBox()
        Me.txtCodSucMatriz = New System.Windows.Forms.TextBox()
        Me.txtFolio = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dbcDescripcion = New System.Windows.Forms.ComboBox()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
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
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Concepto por el que Sale la Mercancia")
        '
        'txtRecibe
        '
        Me.txtRecibe.AcceptsReturn = True
        Me.txtRecibe.BackColor = System.Drawing.SystemColors.Window
        Me.txtRecibe.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRecibe.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRecibe.Location = New System.Drawing.Point(66, 141)
        Me.txtRecibe.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRecibe.MaxLength = 50
        Me.txtRecibe.Name = "txtRecibe"
        Me.txtRecibe.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRecibe.Size = New System.Drawing.Size(200, 20)
        Me.txtRecibe.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtRecibe, "Persona que Recibe la Mercancia")
        '
        'txtEntrega
        '
        Me.txtEntrega.AcceptsReturn = True
        Me.txtEntrega.BackColor = System.Drawing.SystemColors.Window
        Me.txtEntrega.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEntrega.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEntrega.Location = New System.Drawing.Point(66, 118)
        Me.txtEntrega.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEntrega.MaxLength = 50
        Me.txtEntrega.Name = "txtEntrega"
        Me.txtEntrega.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEntrega.Size = New System.Drawing.Size(200, 20)
        Me.txtEntrega.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtEntrega, "Persona que Entrega la Mercancia")
        '
        'txtEnvia
        '
        Me.txtEnvia.AcceptsReturn = True
        Me.txtEnvia.BackColor = System.Drawing.SystemColors.Window
        Me.txtEnvia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEnvia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEnvia.Location = New System.Drawing.Point(66, 95)
        Me.txtEnvia.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEnvia.MaxLength = 50
        Me.txtEnvia.Name = "txtEnvia"
        Me.txtEnvia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEnvia.Size = New System.Drawing.Size(200, 20)
        Me.txtEnvia.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtEnvia, "Persona que Envia la Mercancia")
        '
        'txtCodSucVendExterno
        '
        Me.txtCodSucVendExterno.AcceptsReturn = True
        Me.txtCodSucVendExterno.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucVendExterno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucVendExterno.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucVendExterno.Location = New System.Drawing.Point(57, 25)
        Me.txtCodSucVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucVendExterno.MaxLength = 3
        Me.txtCodSucVendExterno.Name = "txtCodSucVendExterno"
        Me.txtCodSucVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucVendExterno.Size = New System.Drawing.Size(26, 20)
        Me.txtCodSucVendExterno.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtCodSucVendExterno, "Codigo del Vendedor Externo")
        '
        'txtDescSucMatriz
        '
        Me.txtDescSucMatriz.AcceptsReturn = True
        Me.txtDescSucMatriz.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescSucMatriz.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescSucMatriz.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescSucMatriz.Location = New System.Drawing.Point(90, 24)
        Me.txtDescSucMatriz.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescSucMatriz.MaxLength = 0
        Me.txtDescSucMatriz.Name = "txtDescSucMatriz"
        Me.txtDescSucMatriz.ReadOnly = True
        Me.txtDescSucMatriz.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescSucMatriz.Size = New System.Drawing.Size(174, 20)
        Me.txtDescSucMatriz.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtDescSucMatriz, "Descripción de la Sucursal Matriz")
        '
        'txtCodSucMatriz
        '
        Me.txtCodSucMatriz.AcceptsReturn = True
        Me.txtCodSucMatriz.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucMatriz.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucMatriz.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucMatriz.Location = New System.Drawing.Point(60, 23)
        Me.txtCodSucMatriz.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucMatriz.MaxLength = 3
        Me.txtCodSucMatriz.Name = "txtCodSucMatriz"
        Me.txtCodSucMatriz.ReadOnly = True
        Me.txtCodSucMatriz.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucMatriz.Size = New System.Drawing.Size(26, 20)
        Me.txtCodSucMatriz.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodSucMatriz, "Codigo de la Sucursal Matriz")
        '
        'txtFolio
        '
        Me.txtFolio.AcceptsReturn = True
        Me.txtFolio.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolio.Location = New System.Drawing.Point(48, 18)
        Me.txtFolio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolio.MaxLength = 17
        Me.txtFolio.Name = "txtFolio"
        Me.txtFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolio.Size = New System.Drawing.Size(132, 20)
        Me.txtFolio.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolio, "Folio de Salida de Mercancia")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.dbcDescripcion)
        Me.Frame2.Controls.Add(Me.txtFlex)
        Me.Frame2.Controls.Add(Me.flexDetalle)
        Me.Frame2.Controls.Add(Me.Frame5)
        Me.Frame2.Controls.Add(Me.txtRecibe)
        Me.Frame2.Controls.Add(Me.txtEntrega)
        Me.Frame2.Controls.Add(Me.txtEnvia)
        Me.Frame2.Controls.Add(Me.Frame4)
        Me.Frame2.Controls.Add(Me.Frame3)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(6, 52)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(614, 420)
        Me.Frame2.TabIndex = 15
        Me.Frame2.TabStop = False
        '
        'dbcDescripcion
        '
        Me.dbcDescripcion.Location = New System.Drawing.Point(66, 195)
        Me.dbcDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcDescripcion.Name = "dbcDescripcion"
        Me.dbcDescripcion.Size = New System.Drawing.Size(44, 21)
        Me.dbcDescripcion.TabIndex = 20
        Me.dbcDescripcion.Visible = False
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(13, 195)
        Me.txtFlex.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(50, 20)
        Me.txtFlex.TabIndex = 10
        Me.txtFlex.Visible = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(16, 216)
        Me.flexDetalle.Margin = New System.Windows.Forms.Padding(2)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(586, 192)
        Me.flexDetalle.TabIndex = 11
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.txtConcepto)
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame5.Location = New System.Drawing.Point(374, 92)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(229, 79)
        Me.Frame5.TabIndex = 24
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Concepto"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.dbcSucursal)
        Me.Frame4.Controls.Add(Me.txtCodSucVendExterno)
        Me.Frame4.Controls.Add(Me.Label4)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(317, 17)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(285, 59)
        Me.Frame4.TabIndex = 17
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Vendedor Externo"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(87, 24)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(188, 21)
        Me.dbcSucursal.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(4, 27)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(54, 13)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Sucursal :"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.txtDescSucMatriz)
        Me.Frame3.Controls.Add(Me.txtCodSucMatriz)
        Me.Frame3.Controls.Add(Me.Label3)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(24, 17)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(281, 59)
        Me.Frame3.TabIndex = 16
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Origen ...."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(4, 26)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Sucursal :"
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
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Recibe :"
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
        Me.Label6.Size = New System.Drawing.Size(57, 17)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Entrega :"
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
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Envia :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFecha)
        Me.Frame1.Controls.Add(Me.txtFolio)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(6, 0)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(520, 46)
        Me.Frame1.TabIndex = 12
        Me.Frame1.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(352, 18)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(99, 20)
        Me.dtpFecha.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(306, 19)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(41, 17)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Fecha :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 20)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(37, 17)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Folio :"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label8.Location = New System.Drawing.Point(51, 550)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(498, 18)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Supr para Eliminar una Partida"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(243, 493)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 69
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(136, 493)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 68
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(441, 493)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 71
        Me.btnLimpiar.Text = "Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(342, 493)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(93, 35)
        Me.btnBuscar.TabIndex = 70
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'frmVtasVESalidadeMercancia
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(630, 588)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label8)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(194, 145)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVESalidadeMercancia"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Entrega de Mercancia al Vendedor Externo"
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Function BuscaMovimientos() As Boolean
        On Error GoTo Merr
        Dim RsAux As ADODB.Recordset
        gStrSql = "SELECT FOLIOALMACEN,FECHAALMACEN,CODMOVTOALM,REFERENCIADEORIGEN FROM MOVTOSALMACENCAB " & "Where CODALMACEN = " & txtCodSucVendExterno.Text & " AND (CodMovtoAlm = " & C_EntradaaAlmacendeVendedorExterno & " Or CodMovtoAlm = " & C_SalidadeAlmacendeVendedorExterno & " Or " & "CodMovtoAlm = " & C_SalidaPorVentadeVendedoresExternos & ") " & "ORDER BY FOLIOALMACEN DESC,FECHAALMACEN DESC"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            RsAux.MoveFirst()
            If RsAux.Fields("CodMovtoAlm").Value = C_SalidaPorVentadeVendedoresExternos Then
                BuscaMovimientos = True
            Else
                MsgBox("Este vendedor externo, no ha liquidado su última entrega de mercancía. " & Chr(13) & "No es posible registrar este movimiento...  ", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                BuscaMovimientos = False
            End If
        Else
            BuscaMovimientos = True
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function BuscaEntrada() As Boolean
        On Error GoTo Merr
        gStrSql = "SELECT * FROM MovtosAlmacenCab WHERE ReferenciadeOrigen = '" & txtFolio.Text & "' AND CodMovtoAlm = " & C_SalidadeAlmacendeVendedorExterno & " " & "AND Estatus = 'V'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            BuscaEntrada = True
        Else
            BuscaEntrada = False
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function BuscaLiquidacion() As Boolean
        On Error GoTo Merr
        gStrSql = "SELECT * FROM MovtosAlmacenCab WHERE ReferenciadeOrigen = '" & txtFolio.Text & "' AND CodMovtoAlm = " & C_SalidaPorVentadeVendedoresExternos & " " & "AND Estatus = 'V'"
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

    Sub Cancelar()
        On Error GoTo Merr
        Dim FolioEntrada As String
        Dim blnTransaccion As Boolean
        Dim I As Integer
        If BuscaLiquidacion() Then
            MsgBox("Este folío ya tiene una liquidación, no se puede cancelar.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If BuscaEntrada() Then
            MsgBox("No se puede cancelar este folío de almacén, ya que tiene un folío de recepción registrado" & Chr(13) & "Primero debe cancelar el folío de recepción", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If mblnNuevo Then
            Exit Sub
        End If
        Select Case MsgBox("¿Desea cancelar este folio de entrega de mercancia?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
            Case MsgBoxResult.No
                Exit Sub
            Case MsgBoxResult.Cancel
                Exit Sub
        End Select
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        gStrSql = "Select FolioAlmacen FROM MovtosAlmacenCab WHERE ReferenciadeOrigen = '" & txtFolio.Text & "' " & "AND Estatus = 'V' AND CodMovtoAlm = " & C_EntradaaAlmacendeVendedorExterno
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            FolioEntrada = RsGral.Fields("FolioAlmacen").Value
        End If
        'Cancelar los Movimientos de Cabecero
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(txtFolio.Text, "01/01/1900", txtCodSucMatriz.Text, "0", "", "0", "", "0", "0", "", "", "", "", "", "C", "", VB6.Format(Today, C_FORMATFECHAGUARDAR), gStrNomUsuario, "", "01/01/1900", "0", "", "01/01/1900", "0", "", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(FolioEntrada, "01/01/1900", txtCodSucVendExterno.Text, "0", "", "0", "", "0", "0", "", "", "", "", "", "C", "", VB6.Format(Today, C_FORMATFECHAGUARDAR), gStrNomUsuario, "", "01/01/1900", "0", "", "01/01/1900", "0", "", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        'Cancelar los Detalle
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" Then
                    'Cancelar el Detalle de Salida
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(txtFolio.Text, "0", "01/01/1900", .get_TextMatrix(I, 0), "0", "0", "0", "0", "0", "C", VB6.Format(Today, C_FORMATFECHAGUARDAR), "0", C_ELIMINACION, CStr(0))
                    Cmd.Execute()
                    'Cancelar el Detalle de Entrada
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(FolioEntrada, "0", "01/01/1900", .get_TextMatrix(I, 0), "0", "0", "0", "0", "0", "C", VB6.Format(Today, C_FORMATFECHAGUARDAR), "0", C_ELIMINACION, CStr(0))
                    Cmd.Execute()

                    'Guardar el Detalle de Inventario de Entrada
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucMatriz.Text, "1", .get_TextMatrix(I, 0), txtCodSucMatriz.Text, "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 5))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 5), CStr(Numerico(.get_TextMatrix(I, 3))), "0", "0", CStr(C_SalidaAVendedoresExternos), VB6.Format(Today, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Guardar el Detalle de Inventario de Salida
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucVendExterno.Text, "0", .get_TextMatrix(I, 0), txtCodSucMatriz.Text, "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 5))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 5), "0", CStr(Numerico(.get_TextMatrix(I, 3))), "0", CStr(C_EntradaaAlmacendeVendedorExterno), VB6.Format(Today, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Han sido cancelados con éxito el folio de entrega de mercancía " & txtFolio.Text & Chr(13) & "                       y el folío de recepción de mercancia " & FolioEntrada, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Limpiar()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

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
        'Generar el Folio de Salida a Vendedor Externo
        '    gStrSql = "SELECT * FROM FoliosAlmacen WHERE CodAlmacen = " & txtCodSucMatriz
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
        '    Set RsGral = Cmd.Execute
        '    If RsGral.RecordCount > 0 Then
        '        Consecutivo = RsGral!ConsecutivoMovtoAlm + 1
        '    Else
        '        ModStoredProcedures.PR_I_FoliosAlmacen txtCodSucMatriz, CStr(Consecutivo), " ", 0
        '        Cmd.Execute
        '    End If
        ModStoredProcedures.PR_I_FoliosAlmacen(txtCodSucMatriz.Text, CStr(Consecutivo), "", CStr(0))
        Cmd.Execute()
        Consecutivo = Cmd.Parameters("Consecutivo").Value
        txtFolio.Text = C_PrefijoFoliosAlmacen & Format(txtCodSucMatriz.Text, "00") & Year(dtpFecha.Value) & Format(Month(dtpFecha.Value), "00") & Format((dtpFecha.Value), "00") & Format(Consecutivo, "000000")
        'Generar el Folio de Entrada a Vendedor Externo
        '    gStrSql = "SELECT * FROM FoliosAlmacen WHERE CodAlmacen = " & txtCodSucVendExterno
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
        '    Set RsGral = Cmd.Execute
        '    Consecutivo = 0
        '    If RsGral.RecordCount > 0 Then
        '        Consecutivo = RsGral!ConsecutivoMovtoAlm + 1
        '    Else
        '        Consecutivo = Consecutivo + 1
        '        ModStoredProcedures.PR_I_FoliosAlmacen txtCodSucVendExterno, CStr(Consecutivo), " ", 0
        '        Cmd.Execute
        '    End If
        ModStoredProcedures.PR_I_FoliosAlmacen(txtCodSucVendExterno.Text, CStr(Consecutivo), " ", CStr(0))
        Cmd.Execute()
        Consecutivo = Cmd.Parameters("Consecutivo").Value
        FolioEntrada = C_PrefijoFoliosAlmacen & Format(txtCodSucVendExterno.Text, "00") & Year(dtpFecha.Value) & Format(Month(dtpFecha.Value), "00") & Format((dtpFecha.Value), "00") & Format(Consecutivo, "000000")
        'Guardar el Movimiento de Cabecero de Salida de Almacen
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(txtFolio.Text, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), txtCodSucMatriz.Text, "0", "", "0", "", txtCodSucVendExterno.Text, CStr(C_SalidaAVendedoresExternos), C_SALIDA, txtEnvia.Text, txtEntrega.Text, txtRecibe.Text, QuitaEnter(txtConcepto.Text), "V", gStrNomUsuario, "01/01/1900", "", "", "01/01/1900", "0", "", "01/01/1900", CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar el Movimiento de Cabecero de Entrada a Almacen de Vendedor Externo
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(FolioEntrada, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), txtCodSucVendExterno.Text, "0", "", "0", "", txtCodSucMatriz.Text, CStr(C_EntradaaAlmacendeVendedorExterno), C_ENTRADA, txtEnvia.Text, txtEntrega.Text, txtRecibe.Text, "ENTRADA A ALMACEN DE VENDEDOR EXTERNO " & dbcSucursal.Text, "V", gStrNomUsuario, "01/01/1900", "", txtFolio.Text, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), "0", "", "01/01/1900", CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar los Detalles de Entrada y Salida
        NumPartida = 1
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" Then
                    '''                'Guarda el Detalle de Salida
                    '''                ModStoredProcedures.PR_IE_MovtosAlmacenDet txtFolio, CStr(NumPartida), Format(dtpFecha, C_FORMATFECHAGUARDAR), _
                    ''''                .TextMatrix(I, 0), "0", .TextMatrix(I, 3), .TextMatrix(I, 5), CStr(CCur(Numerico(.TextMatrix(I, 4))) / (1 + Round(gcurCorpoTASAIVA / 100, 2))), _
                    ''''                "0", "V", "01/01/1900", "0", C_INSERCION, 0
                    '''                Cmd.Execute
                    '''                'Guarda el Detalle de Entrada
                    '''                ModStoredProcedures.PR_IE_MovtosAlmacenDet FolioEntrada, CStr(NumPartida), Format(dtpFecha, C_FORMATFECHAGUARDAR), _
                    ''''                .TextMatrix(I, 0), "0", .TextMatrix(I, 3), .TextMatrix(I, 5), CStr(CCur(Numerico(.TextMatrix(I, 4))) / (1 + Round(gcurCorpoTASAIVA / 100, 2))), _
                    ''''                "0", "V", "01/01/1900", "0", C_INSERCION, 0
                    '''                Cmd.Execute

                    '''se modifico el precio de venta por precio publico - se elimino dividir el preciopub/tasaiva
                    'Guarda el Detalle de Salida
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(txtFolio.Text, CStr(NumPartida), Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), .get_TextMatrix(I, 3), .get_TextMatrix(I, 5), CStr(CDec(Numerico(.get_TextMatrix(I, 4)))), "0", "V", "01/01/1900", "0", C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Guarda el Detalle de Entrada
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(FolioEntrada, CStr(NumPartida), Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), .get_TextMatrix(I, 3), .get_TextMatrix(I, 5), CStr(CDec(Numerico(.get_TextMatrix(I, 4)))), "0", "V", "01/01/1900", "0", C_INSERCION, CStr(0))
                    Cmd.Execute()

                    '''Determinar el almacen origen del articulo
                    '''tenia: txtCodSucMatriz - esto se refiere al almacen que lo genera
                    '''20OCT2004

                    'Guardar el Detalle de Inventario de Salida
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucMatriz.Text, "1", .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 5))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 5), "0", .get_TextMatrix(I, 3), "0", CStr(C_SalidaAVendedoresExternos), Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Guardar el Detalle de Inventario de Entrada
                    ModStoredProcedures.PR_IE_Inventario(txtCodSucVendExterno.Text, "0", .get_TextMatrix(I, 0), Trim(.get_TextMatrix(I, 7)), "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, 5))) * gcurCorpoTIPOCAMBIODOLAR), .get_TextMatrix(I, 5), .get_TextMatrix(I, 3), "0", "0", CStr(C_EntradaaAlmacendeVendedorExterno), Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()

                    NumPartida = NumPartida + 1
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Los datos se han guardado con éxito" & vbNewLine & "Se han generado los siguientes folios  : " & vbNewLine & vbNewLine & "Folio de entrega    : " & txtFolio.Text & vbNewLine & "Folio de entrada : " & FolioEntrada, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        If MsgBox("¿Desea imprimir el comprobante de entrega de mercancia?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, gstrNombCortoEmpresa) = MsgBoxResult.Yes Then
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

    Sub BuscarArticulos(ByRef CodArticulo As String)
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        'Dim strControlActual As String 'Nombre del control actual
        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        If Not mblnNuevo Then
            Exit Sub
        End If

        strCaptionForm = "Busqueda de Articulos"


        gStrSql = "SELECT RIGHT('       '+LTRIM(CodArticulo),7) AS CODIGO,RTRIM(DescArticulo) AS DESCRIPCION , CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]  " & "FROM CatArticulos  WHERE (CodArticulo = " & CDbl(CodArticulo) & ") " & "OR   (OrigenAnt = " & (CodArticulo) & ") AND (CodigoAnt = " & (CodArticulo) & ") " & "ORDER BY CodArticulo"
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
        ConfiguraConsultas(FrmConsultas, 7600, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            '        Select Case strControlActual
            '            Case "TXTCODSUCVENDEXTERNO"
            .set_ColWidth(0, 0, 900) 'Columna del Código
            .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
            .set_ColWidth(2, 0, 1900)
            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightBottom)
            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftBottom)
            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightBottom)

            FrmConsultas.Left = VB6.TwipsToPixelsX(2000)
            '        End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaVendedorExterno()
        On Error GoTo Merr
        gStrSql = "SELECT DescAlmacen,TipoAlmacen FROM CatAlmacen WHERE CodAlmacen = " & txtCodSucVendExterno.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("TipoAlmacen").Value = "P" Then
                MsgBox("Este código no es de un vendedor externo" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodSucVendExterno.Text = ""
                txtCodSucVendExterno.Focus()
                Exit Sub
            Else
                'If BuscaMovimientos Then
                txtCodSucVendExterno.Text = txtCodSucVendExterno.Text
                FueraChange = True
                dbcSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
                FueraChange = False
                'Else
                '   txtCodSucVendExterno = ""
                '    dbcSucursal.text = ""
                '    txtCodSucVendExterno.SetFocus
                'End If
            End If
        Else
            MsgBox("Código de almacén no existe" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodSucVendExterno.Text = ""
            txtCodSucVendExterno.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function BuscarCodigo(ByRef Codigo As Integer, ByRef RengNotBusca As Integer) As Boolean
        Dim I As Integer
        BuscarCodigo = False
        With flexDetalle
            For I = 1 To .Rows - 1
                If I <> RengNotBusca Then
                    If CDbl(Numerico(.get_TextMatrix(I, 0))) = Codigo Then
                        BuscarCodigo = True
                        Exit Function
                    End If
                End If
            Next
        End With
    End Function

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            Select Case flexDetalle.Col
                Case 0 'Codigo del Articulo
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 8
                Case 3 'Cantidad de Articulos
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 5
            End Select
        End With
    End Sub

    Function ChecaGrid() As Boolean
        Dim I As Integer

        ChecaGrid = False
        With flexDetalle
            For I = 1 To .Rows - 1
                If I = 1 Then
                    If Trim(.get_TextMatrix(I, 0)) = "" And Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "" And Trim(.get_TextMatrix(I, 3)) = "" Then
                        MsgBox("No ha capturado ninguna partida" & vbNewLine & "Favor de capturar al menos una partida", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                        ChecaGrid = False
                        .Row = 1
                        .Col = 0
                        .Focus()
                        Exit Function
                    End If
                    If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" Then
                        If CDbl(Numerico(.get_TextMatrix(I, 4))) = 0 And CDbl(Numerico(.get_TextMatrix(I, 5))) = 0 Then
                            MsgBox("El artículo: " & .get_TextMatrix(I, 0) & " no tiene" & vbNewLine & "Precio público, ni costo" & vbNewLine & "No se puede guardar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            .Row = I
                            ChecaGrid = False
                            Exit Function
                        End If
                        ChecaGrid = True
                    End If
                    If Trim(.get_TextMatrix(I, 7)) = "" And Trim(.get_TextMatrix(I, 0)) <> "" Then '''el articulo no tiene origen
                        MsgBox("El artículo: " & .get_TextMatrix(I, 0) & " no tiene" & vbNewLine & "código del almacen origen" & vbNewLine & "No se puede guardar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        ChecaGrid = False
                        Exit Function
                    End If
                    '''                    MsgBox "No ha capturado toda la información" & vbNewLine & "de la primera partida" & vbNewLine & "Favor de verificar..", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                    '''                    ChecaGrid = False
                    '''                    Exit Function
                    '''                End If
                Else
                    '''se acabaron las partidas
                    If Trim(.get_TextMatrix(I, 0)) = "" And Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "" And Trim(.get_TextMatrix(I, 3)) = "" Then
                        ChecaGrid = True
                        Exit Function
                    End If
                    If Trim(.get_TextMatrix(I, 7)) = "" And Trim(.get_TextMatrix(I, 0)) <> "" Then '''el articulo no tiene origen
                        MsgBox("El artículo: " & .get_TextMatrix(I, 0) & " no tiene" & vbNewLine & "código del almacen origen" & vbNewLine & "No se puede guardar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        ChecaGrid = False
                        Exit Function
                    End If
                    If Trim(.get_TextMatrix(I, 0)) = "" And Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "" And Trim(.get_TextMatrix(I, 3)) = "" Then
                        ChecaGrid = True
                    End If
                    If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" Then
                        If CDbl(Numerico(.get_TextMatrix(I, 4))) = 0 And CDbl(Numerico(.get_TextMatrix(I, 5))) = 0 Then
                            MsgBox("El artículo: " & .get_TextMatrix(I, 0) & " no tiene" & vbNewLine & "Precio público, ni costo" & vbNewLine & "No se puede guardar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            .Row = I
                            ChecaGrid = False
                            Exit Function
                        End If
                        ChecaGrid = True
                        If I = .Rows - 1 Then Exit Function
                    End If
                    '''                Else
                    '''                    MsgBox "No ha capturado toda la información" & vbNewLine & "de la última partida" & vbNewLine & "Favor de verificar..", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                    '''                    ChecaGrid = False
                    '''                    Exit Function
                    '''                End If
                End If
            Next
        End With
    End Function

    Sub EliminarLinea()
        Dim Ren As Integer
        Ren = flexDetalle.Rows
        flexDetalle.RemoveItem(flexDetalle.Row)
        flexDetalle.Rows = Ren
    End Sub

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
            .set_ColWidth(1, 0, 4500)
            .CellFontBold = True
            .Text = "Descripción"
            .Col = 2
            .CellAlignment = 5
            .set_ColWidth(2, 0, 1500)
            .CellFontBold = True
            .Text = "Unidad"
            .Col = 3
            .CellAlignment = 5
            .set_ColWidth(3, 0, 1500)
            .CellFontBold = True
            .Text = "Cantidad"
            .Col = 4
            .set_ColWidth(4, 0, 0)
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
        Printer.CurrentX = 80
        Printer.CurrentY = 15
        Printer.Print("Entrega de Mercancia")
        Printer.CurrentX = 5
        Printer.CurrentY = 20
        Printer.Print("Folio Entrega : " & txtFolio.Text)
        Printer.CurrentX = 160
        Printer.CurrentY = 20
        Printer.Print("Pagina : " & (Space(5) & Pagina))
        Printer.CurrentX = 5
        Printer.CurrentY = 25
        Printer.Print("Vendedor :     " & dbcSucursal.Text)
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
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" Then
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
                    Printer.Print((Space(6) & .get_TextMatrix(I, 3)))
                    TotalArt = TotalArt + CShort(Numerico(.get_TextMatrix(I, 3)))
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
                Printer.Print(txtEntrega.Text)
                Printer.CurrentX = 150
                Printer.CurrentY = 120
                Printer.Print(dbcSucursal.Text)
            Else
                Printer.CurrentX = 15
                Printer.CurrentY = 250
                Printer.Print("Entregó")
                Printer.CurrentX = 150
                Printer.CurrentY = 250
                Printer.Print("Recibi Artículos")
                Printer.CurrentX = 15
                Printer.CurrentY = 255
                Printer.Print(txtEntrega.Text)
                Printer.CurrentX = 150
                Printer.CurrentY = 255
                Printer.Print(dbcSucursal.Text)
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
        intCodSucursal = 0
        intCodArticulo = 0
        tecla = 0
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
        gStrSql = "SELECT CAB.FOLIOALMACEN,CAB.FECHAALMACEN,ALM.DESCALMACEN,CAB.CODALMACENREF,CAB.ENVIA,CAB.ENTREGA,CAB.RECIBE,CAB.CONCEPTO," & "DET.CodArticulo , ART.descArticulo,ART.CostoReal,ART.PrecioPubDolar,uni.DESCUNIDAD,DET.Cantidad " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "INNER JOIN CATALMACEN ALM ON CAB.CODALMACENREF = ALM.CODALMACEN " & "INNER JOIN CATARTICULOS ART ON DET.CODARTICULO = ART.CODARTICULO " & "INNER JOIN CATUNIDADES UNI ON ART.CODUNIDAD = UNI.CODUNIDAD " & "WHERE CAB.CodAlmacen = " & txtCodSucMatriz.Text & " AND CAB.FolioAlmacen = '" & txtFolio.Text & "' AND CAB.CodMovtoAlm = " & C_SalidaAVendedoresExternos & " " & "AND CAB.ESTATUS = 'V'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            dtpFecha.Value = Format(RsGral.Fields("FechaAlmacen").Value, C_FORMATFECHAMOSTRAR)
            txtCodSucVendExterno.Text = RsGral.Fields("CodALmacenREf").Value
            dbcSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            txtEnvia.Text = Trim(RsGral.Fields("Envia").Value)
            txtEntrega.Text = Trim(RsGral.Fields("Entrega").Value)
            txtRecibe.Text = Trim(RsGral.Fields("Recibe").Value)
            txtConcepto.Text = QuitaEnter(Trim(RsGral.Fields("Concepto").Value))
            I = 1
            Do While Not RsGral.EOF
                With flexDetalle
                    .set_TextMatrix(I, 0, RsGral.Fields("CodArticulo").Value)
                    .set_TextMatrix(I, 1, Trim(RsGral.Fields("DescArticulo").Value))
                    .set_TextMatrix(I, 2, Trim(RsGral.Fields("DescUnidad").Value))
                    .set_TextMatrix(I, 3, RsGral.Fields("Cantidad").Value)
                    .set_TextMatrix(I, 4, RsGral.Fields("PrecioPubDolar").Value)
                    .set_TextMatrix(I, 5, RsGral.Fields("CostoReal").Value)
                    RsGral.MoveNext()
                    If Not RsGral.EOF Then
                        If .Rows - 1 = I Then
                            .Rows = .Rows + 1
                        End If
                        I = I + 1
                    End If
                End With
            Loop
            mblnNuevo = False
            Frame4.Enabled = False
            Frame5.Enabled = False
            txtEnvia.Enabled = False
            txtEntrega.Enabled = False
            txtRecibe.Enabled = False
            ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Else
            MsgBox("El folio de entrega de mercancia del vendedor externo no existe" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtFolio.Text = ""
            If txtFolio.Enabled Then txtFolio.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function LlenaDatosArticulos() As Boolean
        On Error GoTo Merr
        Dim Existe As Boolean
        If Me.ActiveControl.Name = "dbcDescripcion" Then
            gStrSql = "SELECT CodArticulo,DescArticulo FROM CatArticulos WHERE DescArticulo LIKE '" & Trim(dbcDescripcion.Text) & "%' ORDER BY DescArticulo"
            DCLostFocus(dbcDescripcion, gStrSql, intCodArticulo)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                If BuscarCodigo(CInt(RsGral.Fields("CodArticulo").Value), (flexDetalle.Row)) Then
                    MsgBox("Artículo repetido" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Existe = False
                    Exit Function
                End If
                Existe = True
            Else
                Existe = False
            End If

            '''        gStrSql = "SELECT CodArticulo,DescArticulo FROM CatArticulos WHERE DescArticulo LIKE '" _
            ''''        & Trim(dbcDescripcion) & "%' ORDER BY DescArticulo"
            '''        DCLostFocus dbcDescripcion, gStrSql, intCodArticulo
            '''        gStrSql = "SELECT * FROM CatArticulos WHERE CodArticulo = " & intCodArticulo
            '''        ModEstandar.BorraCmd
            '''        Cmd.CommandText = "dbo.Up_Select_Datos"
            '''        Cmd.CommandType = adCmdStoredProc
            '''        Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
            '''        Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
            '''        Set RsGral = Cmd.Execute
            '''        If RsGral.RecordCount > 0 Then
            '''            If BuscarCodigo(CLng(intCodArticulo), flexDetalle.Row) Then
            '''                MsgBox "Artículo repetido" & vbNewLine & "Favor de verificar ...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
            '''                Existe = False
            '''                Exit Function
            '''            End If
            '''            Existe = True
            '''        Else
            '''            Existe = False
            '''        End If
            '''
            '''gStrSql = "SELECT CA.CODARTICULO,CA.DESCARTICULO,CU.DESCUNIDAD, SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) AS EXISTENCIA," & _
            '"CA.PrecioPubDolar,CA.CostoReal " & _
            '"FROM INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO " & _
            '"INNER JOIN CATUNIDADES CU ON CA.CODUNIDAD = CU.CODUNIDAD WHERE I.CODARTICULO = " & intCodArticulo & " AND I.CodAlmacen = " & txtCodSucMatriz & _
            '"GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CU.DESCUNIDAD,CA.PRECIOPUBDOLAR,CA.COSTOREAL"

            gStrSql = "SELECT   CA.CODARTICULO, CA.DESCARTICULO, CU.DESCUNIDAD, SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) AS EXISTENCIA, CA.PrecioPubDolar , CA.CostoReal, CA.CodAlmacenOrigen " & "FROM     INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO INNER JOIN CATUNIDADES CU ON CA.CODUNIDAD = CU.CODUNIDAD " & "Where    I.CodArticulo = " & RsGral.Fields("CodArticulo").Value & " And I.CodAlmacen = " & txtCodSucMatriz.Text & "GROUP    BY CA.CODARTICULO, CA.DESCARTICULO, CU.DESCUNIDAD, CA.PRECIOPUBDOLAR, CA.COSTOREAL, CA.CodAlmacenOrigen "
        Else
            gStrSql = "SELECT * FROM CatArticulos WHERE CodArticulo = " & CInt(Numerico(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)))
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                If BuscarCodigo(CInt(Numerico(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))), (flexDetalle.Row)) Then
                    MsgBox("Articulo repetido" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Existe = False
                    Exit Function
                End If
                Existe = True
            Else
                Existe = False
            End If

            gStrSql = "SELECT   CA.CODARTICULO, CA.DESCARTICULO, CU.DESCUNIDAD, SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) AS EXISTENCIA, CA.PrecioPubDolar , CA.CostoReal, CA.CodAlmacenOrigen " & "FROM     INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO INNER JOIN CATUNIDADES CU ON CA.CODUNIDAD = CU.CODUNIDAD " & "Where    I.CodArticulo = " & CInt(Numerico(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))) & " And I.CodAlmacen = " & txtCodSucMatriz.Text & "GROUP    BY CA.CODARTICULO, CA.DESCARTICULO, CU.DESCUNIDAD, CA.PRECIOPUBDOLAR, CA.COSTOREAL, CA.CodAlmacenOrigen"

            '''gStrSql = "SELECT CA.CODARTICULO,CA.DESCARTICULO,CU.DESCUNIDAD, SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) AS EXISTENCIA," & _
            '"CA.PrecioPubDolar,CA.CostoReal " & _
            '"FROM INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO " & _
            '"INNER JOIN CATUNIDADES CU ON CA.CODUNIDAD = CU.CODUNIDAD WHERE I.CODARTICULO = " & CLng(Numerico(flexDetalle.TextMatrix(flexDetalle.Row, 0))) & " AND I.CodAlmacen = " & txtCodSucMatriz & _
            '"GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CU.DESCUNIDAD,CA.PRECIOPUBDOLAR,CA.COSTOREAL"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Existencia").Value > 0 Then
                With flexDetalle
                    .set_TextMatrix(.Row, 0, RsGral.Fields("CodArticulo").Value)
                    .set_TextMatrix(.Row, 1, Trim(RsGral.Fields("DescArticulo").Value))
                    .set_TextMatrix(.Row, 2, Trim(RsGral.Fields("DescUnidad").Value))
                    .set_TextMatrix(.Row, 3, "")
                    .set_TextMatrix(.Row, 4, RsGral.Fields("PrecioPubDolar").Value)
                    .set_TextMatrix(.Row, 5, RsGral.Fields("CostoReal").Value)
                    .set_TextMatrix(.Row, 6, RsGral.Fields("Existencia").Value)
                    .set_TextMatrix(.Row, 7, RsGral.Fields("CodAlmacenOrigen").Value)
                End With
                LlenaDatosArticulos = True
                txtFlex.Text = ""
                txtFlex.Visible = False
                blnBuscar = False
                Exit Function
            End If
        ElseIf Existe = False Then
            MsgBox("Artículo no existe" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            flexDetalle.set_TextMatrix(flexDetalle.Row, 0, "")
            flexDetalle.set_TextMatrix(flexDetalle.Row, 1, "")
            flexDetalle.set_TextMatrix(flexDetalle.Row, 2, "")
            flexDetalle.set_TextMatrix(flexDetalle.Row, 3, "")
            flexDetalle.set_TextMatrix(flexDetalle.Row, 4, "")
            flexDetalle.set_TextMatrix(flexDetalle.Row, 5, "")
            flexDetalle.set_TextMatrix(flexDetalle.Row, 6, "")
            flexDetalle.set_TextMatrix(flexDetalle.Row, 7, "")
            If txtFlex.Visible = True Then
                txtFlex.Text = ""
                flexDetalle.set_TextMatrix(flexDetalle.Row, 0, "")
                txtFlex.Visible = True
                txtFlex.Focus()
            Else
                flexDetalle.set_TextMatrix(flexDetalle.Row, 0, "")
                '''flexDetalle.SetFocus
            End If
            LlenaDatosArticulos = False
            Exit Function
        End If
        MsgBox("Existencia insuficiente" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        flexDetalle.set_TextMatrix(flexDetalle.Row, 0, "")
        flexDetalle.set_TextMatrix(flexDetalle.Row, 1, "")
        flexDetalle.set_TextMatrix(flexDetalle.Row, 2, "")
        flexDetalle.set_TextMatrix(flexDetalle.Row, 3, "")
        flexDetalle.set_TextMatrix(flexDetalle.Row, 4, "")
        flexDetalle.set_TextMatrix(flexDetalle.Row, 5, "")
        flexDetalle.set_TextMatrix(flexDetalle.Row, 6, "")
        flexDetalle.set_TextMatrix(flexDetalle.Row, 7, "")
        If txtFlex.Visible = True Then
            txtFlex.Text = ""
            flexDetalle.set_TextMatrix(flexDetalle.Row, 0, "")
            txtFlex.Focus()
        Else
            flexDetalle.set_TextMatrix(flexDetalle.Row, 0, "")
            flexDetalle.Focus()
        End If
        LlenaDatosArticulos = False

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Nuevo()
        dtpFecha.Value = Today
        txtCodSucVendExterno.Text = ""
        dbcSucursal.Text = ""
        txtEnvia.Text = ""
        txtEntrega.Text = ""
        txtRecibe.Text = ""
        txtConcepto.Text = "ENTREGA DE MERCANCIA AL VENDEDOR EXTERNO " & txtCodSucVendExterno.Text & " " & dbcSucursal.Text
        flexDetalle.Clear()
        Encabezado()
        Frame4.Enabled = True
        Frame5.Enabled = True
        txtEnvia.Enabled = True
        txtEntrega.Enabled = True
        txtRecibe.Enabled = True
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        InicializaVariables()
    End Sub

    Function ValidaCantidad() As Boolean
        If (CInt(flexDetalle.get_TextMatrix(flexDetalle.Row, 6)) - CInt(txtFlex.Text)) < 0 Then
            MsgBox("Existencia insuficiente" & vbNewLine & "No es suficiente" & vbNewLine & "Existencia: " & flexDetalle.get_TextMatrix(flexDetalle.Row, 6) & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            ValidaCantidad = False
        Else
            ValidaCantidad = True
        End If
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If CInt(Numerico(txtCodSucVendExterno.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Código de vendedor externo ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodSucVendExterno.Focus()
            Exit Function
        End If
        If Trim(dbcSucursal.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Nombre de la sucursal ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcSucursal.Focus()
            Exit Function
        End If
        If Trim(txtEnvia.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Quien envía la mercancía ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtEnvia.Focus()
            Exit Function
        End If
        If Trim(txtEntrega.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Quien entrega la mercancía  ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtEntrega.Focus()
            Exit Function
        End If
        If Trim(txtRecibe.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Quien recibe la mercancía  ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtRecibe.Focus()
            Exit Function
        End If
        If Not ChecaGrid() Then
            flexDetalle.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub dbcDescripcion_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescripcion.CursorChanged
        If FueraChange = True Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcDescripcion" Then
            Exit Sub
        End If
        If Trim(dbcDescripcion.Text) = "" Then
            gStrSql = "SELECT CodArticulo,RTRIM(DescArticulo) AS DescArticulo FROM CatArticulos ORDER BY DescArticulo"
            DCGotFocus(gStrSql)
            Exit Sub
        End If
        gStrSql = "SELECT CodArticulo,DescArticulo FROM CatArticulos WHERE DescArticulo LIKE '" & Trim(dbcDescripcion.Text) & "%' ORDER BY DescArticulo"
        DCChange(gStrSql, tecla)
        intCodArticulo = 0
    End Sub

    Private Sub dbcDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescripcion.Enter
        gStrSql = "SELECT CodArticulo,RTRIM(DescArticulo) AS DescArticulo FROM CatArticulos ORDER BY DescArticulo"
        DCGotFocus(gStrSql)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcDescripcion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcDescripcion.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            If Trim(dbcDescripcion.Text) <> "" Then
                flexDetalle.Text = dbcDescripcion.Text
                blnBuscar = False
                If LlenaDatosArticulos() Then
                    flexDetalle.Col = 3
                    FueraChange = True
                    dbcDescripcion.Text = ""
                    FueraChange = False
                    dbcDescripcion.Visible = False
                    blnBuscar = False
                    flexDetalle.Focus()
                    Exit Sub
                Else
                    FueraChange = True
                    dbcDescripcion.Text = ""
                    FueraChange = False
                    If dbcDescripcion.Enabled Then
                        dbcDescripcion.Focus()
                    End If
                    Exit Sub
                End If
            Else
                flexDetalle.Text = dbcDescripcion.Text
                dbcDescripcion.Visible = False
                blnBuscar = False
                flexDetalle.Focus()
                Exit Sub
            End If
        End If
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            FueraChange = True
            dbcDescripcion.Text = ""
            FueraChange = False
            flexDetalle.Text = ""
            dbcDescripcion.Visible = False
            flexDetalle.Focus()
            blnBuscar = True
        End If
    End Sub

    Private Sub dbcDescripcion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcDescripcion.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcDescripcion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescripcion.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        If blnBuscar Then
            dbcDescripcion_KeyDown(dbcDescripcion, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape))
        End If
        blnBuscar = True
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
            Exit Sub
        End If
        If Trim(dbcSucursal.Text) = "" Then
            txtCodSucVendExterno.Text = ""
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
            Exit Sub
        End If
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursal" Then Exit Sub
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventSender.keyCode
        If eventSender.keyCode = System.Windows.Forms.Keys.Escape Then
            txtCodSucVendExterno.Focus()
        End If
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        FueraChange = True
        If dbcSucursal.SelectedItem <> 0 Then
            gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
            DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
            txtCodSucVendExterno.Text = IIf(intCodSucursal <> 0, intCodSucursal, "")
        End If
        FueraChange = False
        dbcSucursal.Text = Aux
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        FueraChange = True
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        txtCodSucVendExterno.Text = IIf(intCodSucursal <> 0, intCodSucursal, "")
        If CDbl(Numerico(txtCodSucVendExterno.Text)) <> 0 Then
            '''Debe validar que el almacen not tenga  existencia para poder registrar otra salida de mcia al VExt
            '''De lo contrario no es posible registrarla.
            '        If ExistenciaAlmVExt(intCodSucursal) > 0 Then
            '            MsgBox "Este vendedor externo, no ha liquidado su última salida de mercancía. " & Chr(13) & _
            ''                   "No es posible registrar este movimiento...  ", vbOKOnly + vbInformation, gstrNombCortoEmpresa
            '            txtCodSucVendExterno = ""
            '            dbcSucursal.text = ""
            '            ModEstandar.RetrocederTab Me
            '        End If
        End If
        FueraChange = False

    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        FueraChange = True
        If dbcSucursal.SelectedItem <> 0 Then
            gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
            DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
            txtCodSucVendExterno.Text = IIf(intCodSucursal <> 0, intCodSucursal, "")
        End If
        FueraChange = False
        dbcSucursal.Text = Aux
    End Sub

    Private Sub flexDetalle_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.ClickEvent
        txtFlex.Visible = False
        dbcDescripcion.Visible = False
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
        txtConcepto.Text = QuitaEnter(txtConcepto.Text)
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        FlexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub FlexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete And mblnNuevo And Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) <> "" Then
            Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    EliminarLinea()
            End Select
            flexDetalle.Focus()
        End If
    End Sub

    Private Sub FlexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalle.KeyPressEvent
        Dim lonR, lonI As Integer
        Dim EsEnter As Boolean
        EsEnter = False
        blnBuscar = True
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And mblnNuevo Then
            If eventArgs.keyAscii = System.Windows.Forms.Keys.Return Then EsEnter = True
            'Verifica si se puede capturar la fila
            If flexDetalle.Row > 1 Then
                If flexDetalle.get_TextMatrix(flexDetalle.Row - 1, 0) <> "" Then
                    For lonR = 1 To flexDetalle.Row - 1 Step 1
                        For lonI = 0 To 4 Step 1
                            If flexDetalle.get_TextMatrix(lonR, lonI) = "" Then
                                'MsgBox "Hace falta información en la captura", vbExclamation, cNomEmp
                                flexDetalle.Row = lonR
                                flexDetalle.Col = lonI
                                If flexDetalle.Col = 0 Or flexDetalle.Col = 3 Then
                                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                                End If
                                CambiarFormatoTxtenCaptura()
                                MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                                If Len(Trim(txtFlex.Text)) = 1 Then
                                    System.Windows.Forms.SendKeys.Send("{right}")
                                End If
                                Exit Sub
                            End If
                        Next lonI
                    Next lonR
                Else
                    'flexDetalle.SetFocus
                    Exit Sub
                End If
            End If
            'Edita el campo sólo si es Editable
            If flexDetalle.Row >= 1 And flexDetalle.Col < 4 Then
                If flexDetalle.Col = 3 Then
                    If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" Or Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 1)) = "" Or Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2)) = "" Then
                        MsgBox("Debe capturar primero" & vbNewLine & "el código o la descripción del Artículo ..", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        flexDetalle.Col = 0
                        blnBuscar = False
                        Exit Sub
                    End If
                End If
                If (flexDetalle.Col = 0 Or flexDetalle.Col = 3) And blnBuscar Then
                    CambiarFormatoTxtenCaptura()
                    eventArgs.keyAscii = ModEstandar.MskCantidad(txtFlex.Text, eventArgs.keyAscii, 5, 0, (txtFlex.SelectionStart))
                    MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                    If Not EsEnter Then
                        System.Windows.Forms.SendKeys.Send("{right}")
                    End If
                ElseIf flexDetalle.Col = 1 And blnBuscar Then
                    MSHFlexGridEdit(flexDetalle, dbcDescripcion, eventArgs.keyAscii)
                    If Not EsEnter Then
                        System.Windows.Forms.SendKeys.Send("{right}")
                    End If
                End If
                blnBuscar = True
            End If
        End If
    End Sub

    Private Sub frmVtasVESalidadeMercancia_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVESalidadeMercancia_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVESalidadeMercancia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "flexDetalle" Then
                    FlexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
                    Exit Sub
                End If
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "dbcDescripcion" Then
                    Exit Sub
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

    Private Sub frmVtasVESalidadeMercancia_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVESalidadeMercancia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        bandera = True
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        Nuevo()
        ObtenerDatosSucursalMatriz(txtCodSucMatriz, txtDescSucMatriz)
    End Sub

    Private Sub frmVtasVESalidadeMercancia_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If Not mblnSalir Then
            'If Cambios = True And mblnNuevo = False Then
            'Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
            'Case vbYes: 'Guardar el registro
            'If Guardar = False Then
            'Cancel = 1
            'End If
            'Case vbNo: 'No hace nada y permite el cierre del formulario
            'Case vbCancel: 'Cancela el cierre del formulario sin guardar
            'Cancel = 1
            'End Select
            'End If
        Else
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSalir = False
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasVESalidadeMercancia_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtCodSucMatriz_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucMatriz.Enter
        Pon_Tool()
        SelTextoTxt(txtCodSucMatriz)
    End Sub

    Private Sub txtCodSucVendExterno_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucVendExterno.TextChanged
        If FueraChange Then Exit Sub
        dbcSucursal.Text = ""
    End Sub

    Private Sub txtCodSucVendExterno_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucVendExterno.Enter
        strControlActual = UCase("txtCodSucVendExterno")
        Pon_Tool()
        SelTextoTxt(txtCodSucVendExterno)
    End Sub

    Private Sub txtCodSucVendExterno_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucVendExterno.Leave
        If CDbl(Numerico(txtCodSucVendExterno.Text)) = 0 Then
            txtCodSucVendExterno.Text = ""
        Else
            intCodSucursal = CShort(Numerico(txtCodSucVendExterno.Text))
            '        If ExistenciaAlmVExt(intCodSucursal) > 0 Then
            '            MsgBox "Este vendedor externo, no ha liquidado su última salida de mercancía. " & Chr(13) & _
            ''                   "No es posible registrar este movimiento...  ", vbOKOnly + vbInformation, gstrNombCortoEmpresa
            '            txtCodSucVendExterno = ""
            '            dbcSucursal.text = ""
            '            ModEstandar.RetrocederTab Me
            '        Else
            BuscaVendedorExterno()
            '        End If
        End If
    End Sub

    Private Sub txtConcepto_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.TextChanged
        If Trim(txtConcepto.Text) = "" Then
            txtConcepto.Text = "ENTREGA DE MERCANCIA AL VENDEDOR EXTERNO " & txtCodSucVendExterno.Text & " " & dbcSucursal.Text
        End If
        Trim(QuitaEnter(txtConcepto.Text))
    End Sub

    Private Sub txtConcepto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Enter
        Pon_Tool()
        txtConcepto.Text = "ENTREGA DE MERCANCIA AL VENDEDOR EXTERNO " & txtCodSucVendExterno.Text & " " & dbcSucursal.Text
        SelTextoTxt(txtConcepto)
    End Sub

    Private Sub txtConcepto_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Leave
        QuitaEnter(txtConcepto.Text)
    End Sub

    Private Sub txtDescSucMatriz_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescSucMatriz.Enter
        Pon_Tool()
        SelTextoTxt(txtDescSucMatriz)
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
        SelTextoTxt(txtFlex)
        Pon_Tool()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Renamed
        Dim ResBusquedaArt As Integer

        With flexDetalle
            If KeyCode = System.Windows.Forms.Keys.Return Then
                Select Case .Col
                    Case 0, 3
                        If .Col = 0 And Trim(txtFlex.Text) <> "" Then
                            .set_TextMatrix(.Row, .Col, Trim(txtFlex.Text))
                            ResBusquedaArt = BuscarCodigoArticulo(Trim(.get_TextMatrix(.Row, 0)))
                            If (ResBusquedaArt > 0 Or ResBusquedaArt = -1) Then
                                intCodArticulo = ResBusquedaArt
                                '''19OCT2004 - dato incorrecto en el flex - por eso no lo encontraba
                                '''txtFlex = IIf((ResBusquedaArt > 0), .TextMatrix(.Row, .Col), "")
                                '''flexDetalle.TextMatrix(flexDetalle.Row, 0) = IIf((ResBusquedaArt > 0), .TextMatrix(.Row, .Col), "")
                                txtFlex.Text = IIf((ResBusquedaArt > 0), ResBusquedaArt, "")
                                flexDetalle.set_TextMatrix(flexDetalle.Row, 0, IIf((ResBusquedaArt > 0), ResBusquedaArt, ""))

                                If LlenaDatosArticulos() Then
                                    .Col = 3
                                    txtFlex.Visible = False
                                    Exit Sub
                                Else
                                    .Text = ""
                                    txtFlex.Text = ""
                                    txtFlex.Visible = True
                                    txtFlex.Focus()
                                    Exit Sub
                                End If
                            ElseIf ResBusquedaArt = -2 Then
                                BuscarArticulos((New String("0", 6) & Trim(.get_TextMatrix(.Row, 0))))
                            End If
                            Exit Sub
                        ElseIf .Col = 0 And Trim(txtFlex.Text) = "" Then
                            .Text = Trim(txtFlex.Text)
                            '''txtFlex.Visible = False
                        ElseIf .Col = 3 Then
                            If CInt(Numerico((txtFlex.Text))) = 0 Then
                                MsgBox("Teclee una cantidad mayor que cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                                txtFlex.Text = ""
                                txtFlex.Visible = True
                                txtFlex.Focus()
                                Exit Sub
                            End If
                            If ValidaCantidad() Then
                                .Text = Trim(txtFlex.Text)
                            Else
                                txtFlex.Visible = True
                                txtFlex.Focus()
                                Exit Sub
                            End If
                            If .Row = .Rows - 1 Then
                                .Rows = .Rows + 1
                                .Row = .Row + 1
                                .TopRow = .Row
                            Else
                                .Row = .Row + 1
                            End If
                            .Col = 0
                        End If
                        txtFlex.Visible = False
                End Select
            ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
                If flexDetalle.Col = 3 And CDbl(Numerico(txtFlex.Text)) = 0 Then
                    MsgBox("Teclee una cantidad mayor que cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    txtFlex.Text = ""
                    txtFlex.Focus()
                    Exit Sub
                End If
                .Focus()
                txtFlex.Visible = False
            ElseIf KeyCode = System.Windows.Forms.Keys.F3 Then
                Buscar()
            End If
        End With

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case Else
                Select Case flexDetalle.Col
                    Case 0
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 3
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                End Select
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        '''txtFlex.Visible = False
        ''' If Screen.ActiveForm.Name <> Me.Name Then
        '''    Exit Sub
        ''' End If
        ''' If blnBuscar Then
        '''     txtFlex_KeyDown vbKeyEscape, 0
        ''' End If
        ''' blnBuscar = True
    End Sub

    Private Sub txtFlex_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFlex.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If flexDetalle.Col = 3 And CDbl(Numerico(txtFlex.Text)) = 0 Then
            MsgBox("Teclee una cantidad mayor que cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
            Cancel = True
        Else
            Cancel = False
        End If
        eventArgs.Cancel = Cancel
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
            LlenaDatos()
        End If
    End Sub

    Private Sub txtRecibe_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRecibe.Enter
        Pon_Tool()
        SelTextoTxt(txtRecibe)
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub
End Class
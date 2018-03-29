Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioConsultaIngresosXSucursal
    Inherits System.Windows.Forms.Form

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents Label1 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents lblTotalTarjetas As System.Windows.Forms.Label
    Public WithEvents lblTotalDolares As System.Windows.Forms.Label
    Public WithEvents Panel1 As Panel
    Public WithEvents Panel2 As Panel
    Public WithEvents lblTotalPesos As System.Windows.Forms.Label

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoDiarioConsultaIngresosXSucursal))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblTotalTarjetas = New System.Windows.Forms.Label()
        Me.lblTotalDolares = New System.Windows.Forms.Label()
        Me.lblTotalPesos = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdAceptar
        '
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Location = New System.Drawing.Point(470, 275)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.Size = New System.Drawing.Size(114, 36)
        Me.cmdAceptar.TabIndex = 5
        Me.cmdAceptar.Text = "&Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(12, 13)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(538, 126)
        Me.flexDetalle.TabIndex = 4
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(440, 7)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(110, 20)
        Me.dtpFecha.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(394, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Fecha :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(95, 245)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(124, 21)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Efectivo Disponible :"
        '
        'lblTotalTarjetas
        '
        Me.lblTotalTarjetas.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalTarjetas.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalTarjetas.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalTarjetas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalTarjetas.Location = New System.Drawing.Point(445, 240)
        Me.lblTotalTarjetas.Name = "lblTotalTarjetas"
        Me.lblTotalTarjetas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalTarjetas.Size = New System.Drawing.Size(106, 21)
        Me.lblTotalTarjetas.TabIndex = 8
        Me.lblTotalTarjetas.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalDolares
        '
        Me.lblTotalDolares.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalDolares.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalDolares.Location = New System.Drawing.Point(335, 240)
        Me.lblTotalDolares.Name = "lblTotalDolares"
        Me.lblTotalDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalDolares.Size = New System.Drawing.Size(109, 21)
        Me.lblTotalDolares.TabIndex = 7
        Me.lblTotalDolares.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPesos
        '
        Me.lblTotalPesos.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalPesos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalPesos.Location = New System.Drawing.Point(223, 240)
        Me.lblTotalPesos.Name = "lblTotalPesos"
        Me.lblTotalPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalPesos.Size = New System.Drawing.Size(111, 21)
        Me.lblTotalPesos.TabIndex = 6
        Me.lblTotalPesos.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.dtpFecha)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(565, 37)
        Me.Panel1.TabIndex = 10
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.flexDetalle)
        Me.Panel2.Location = New System.Drawing.Point(12, 68)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(565, 155)
        Me.Panel2.TabIndex = 11
        '
        'frmBancosProcesoDiarioConsultaIngresosXSucursal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(589, 320)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdAceptar)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblTotalTarjetas)
        Me.Controls.Add(Me.lblTotalDolares)
        Me.Controls.Add(Me.lblTotalPesos)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(266, 214)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBancosProcesoDiarioConsultaIngresosXSucursal"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Consulta de Ingresos por Sucursal"
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Sub Encabezado()
        On Error GoTo Err_Renamed
        With flexDetalle
            .Col = 0
            .Row = 0
            .set_TextMatrix(.Row, .Col, "Sucursal")
            .CellAlignment = 5
            .CellFontBold = True
            .Col = 1
            .set_TextMatrix(.Row, .Col, "Pesos")
            .CellAlignment = 5
            .CellFontBold = True
            .Col = 2
            .set_TextMatrix(.Row, .Col, "Dólares")
            .CellAlignment = 5
            .CellFontBold = True
            .Col = 3
            .set_TextMatrix(.Row, .Col, "Tarj. No Acred.")
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(0, 0, 3065)
            .set_ColWidth(1, 0, 1635)
            .set_ColWidth(2, 0, 1635)
            .set_ColWidth(3, 0, 1635)
            .Col = 0
            .Row = 1
        End With
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub ObtenerImportes()
        On Error GoTo Err_Renamed
        '    gStrSql = "SELECT C.CodAlmacen,C.DescAlmacen,ISNULL(SUM(ImpDolIng) - SUM(ImpDolDev) - SUM(ImpDolRet),0) AS EfectivoDolares," & _
        ''    "ISNULL(SUM(ImpPesIng) - SUM(ImpPesDev) - SUM(ImpPesRet),0) AS EfectivoPesos," & _
        ''    "ISNULL(Sum(ImpPesTar), 0) As ImporteTarjetas " & _
        ''    "FROM " & _
        ''    "(SELECT I.CodSucursal,ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'I' THEN ImporteDolares END)),0) ImpDolIng," & _
        ''    "ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'I' THEN ImportePesos END)),0) ImpPesIng," & _
        ''    "ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'D' THEN ImporteDolares END)),0) ImpDolDev," & _
        ''    "ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'D' THEN ImportePesos END)),0) ImpPesDev," & _
        ''    "ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'R' THEN ImporteDolares END)),0) ImpDolRet," & _
        ''    "ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'R' THEN ImportePesos END)),0) ImpPesRet," & _
        ''    "ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'T' THEN ImporteDolares END)),0) ImpDolTar," & _
        ''    "ISNULL(ABS(SUM(CASE WHEN I.Tipo = 'T' THEN ImportePesos END)),0) ImpPesTar " & _
        ''    "FROM DBO.vw_ObtenerIngresos I " & _
        ''    "WHERE FechaIngreso <= '" & Format(dtpFecha, "MM/DD/YYYY") & "' " & _
        ''    "GROUP BY I.CodSucursal,I.Tipo) R RIGHT OUTER JOIN CatAlmacen C ON R.CodSucursal = C.CodAlmacen " & _
        ''    "WHERE C.TipoAlmacen = 'P' GROUP BY C.CodAlmacen,C.DescAlmacen ORDER BY C.DescAlmacen"
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        '    Set RsGral = Cmd.Execute
        '    If RsGral.RecordCount > 0 Then
        '        With flexDetalle
        '            .Row = 1
        '            Do While Not RsGral.EOF
        '                .TextMatrix(.Row, 0) = Trim(RsGral!DescAlmacen)
        '                .TextMatrix(.Row, 1) = Format(RsGral!EfectivoPesos, "###,##0.00")
        '                .TextMatrix(.Row, 2) = Format(RsGral!EfectivoDolares, "###,##0.00")
        '                .TextMatrix(.Row, 3) = Format(RsGral!ImporteTarjetas, "###,##0.00")
        '                RsGral.MoveNext
        '                If Not RsGral.EOF Then
        '                    If .Row = .Rows - 1 Then
        '                        .Rows = .Rows + 1
        '                    End If
        '                    .Row = .Row + 1
        '                End If
        '            Loop
        '            .Col = 0
        '            .Row = 1
        '        End With
        '    End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Suma()
        On Error GoTo Err_Renamed
        Dim I As Integer
        lblTotalPesos.Text = CStr(0)
        lblTotalDolares.Text = CStr(0)
        lblTotalTarjetas.Text = CStr(0)
        With flexDetalle
            For I = 1 To .Rows - 1
                lblTotalPesos.Text = CStr(CDbl(Numerico(lblTotalPesos.Text)) + CDbl(Numerico(.get_TextMatrix(I, 1))))
                lblTotalDolares.Text = CStr(CDbl(Numerico(lblTotalDolares.Text)) + CDbl(Numerico(.get_TextMatrix(I, 2))))
                lblTotalTarjetas.Text = CStr(CDbl(Numerico(lblTotalTarjetas.Text)) + CDbl(Numerico(.get_TextMatrix(I, 3))))
            Next
        End With
        lblTotalPesos.Text = VB6.Format(lblTotalPesos.Text, "###,##0.00")
        lblTotalDolares.Text = VB6.Format(lblTotalDolares.Text, "###,##0.00")
        lblTotalTarjetas.Text = VB6.Format(lblTotalTarjetas.Text, "###,##0.00")
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        Me.Close()
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        'System.Windows.Forms.SendKeys.Send("{right}")
    End Sub

    Private Sub frmBancosProcesoDiarioConsultaIngresosXSucursal_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Suma()
    End Sub

    Private Sub frmBancosProcesoDiarioConsultaIngresosXSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            ModEstandar.AvanzarTab(Me)
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioConsultaIngresosXSucursal_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'InitializeComponent()
        dtpFecha.Value = Today
        Encabezado()
        Suma()
    End Sub

    Private Sub frmBancosProcesoDiarioConsultaIngresosXSucursal_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Me = Nothing
        'Me.Close()
    End Sub
End Class
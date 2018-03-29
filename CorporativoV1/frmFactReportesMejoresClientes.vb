Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmFactReportesMejoresClientes

    Inherits System.Windows.Forms.Form

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE LOS MEJORES CLIENTES(FACTURADOS)                                                  *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :                                                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Dim mblnSalir As Boolean
    Dim intCodSucursal As Integer
    Dim rsReporte As ADODB.Recordset
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo

    '''Agregar Facturación Especial

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkIncluirIva As System.Windows.Forms.CheckBox
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtTextoAdicional As System.Windows.Forms.TextBox
    Public WithEvents txtNumMaxClientes As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtTextoAdicional = New System.Windows.Forms.TextBox()
        Me.txtNumMaxClientes = New System.Windows.Forms.TextBox()
        Me.chkIncluirIva = New System.Windows.Forms.CheckBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtTextoAdicional
        '
        Me.txtTextoAdicional.AcceptsReturn = True
        Me.txtTextoAdicional.BackColor = System.Drawing.SystemColors.Window
        Me.txtTextoAdicional.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTextoAdicional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTextoAdicional.Location = New System.Drawing.Point(9, 240)
        Me.txtTextoAdicional.MaxLength = 120
        Me.txtTextoAdicional.Multiline = True
        Me.txtTextoAdicional.Name = "txtTextoAdicional"
        Me.txtTextoAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTextoAdicional.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTextoAdicional.Size = New System.Drawing.Size(373, 67)
        Me.txtTextoAdicional.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtTextoAdicional, "Texto Adicional.")
        '
        'txtNumMaxClientes
        '
        Me.txtNumMaxClientes.AcceptsReturn = True
        Me.txtNumMaxClientes.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumMaxClientes.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumMaxClientes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumMaxClientes.Location = New System.Drawing.Point(206, 20)
        Me.txtNumMaxClientes.MaxLength = 5
        Me.txtNumMaxClientes.Name = "txtNumMaxClientes"
        Me.txtNumMaxClientes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumMaxClientes.Size = New System.Drawing.Size(52, 20)
        Me.txtNumMaxClientes.TabIndex = 2
        Me.txtNumMaxClientes.Text = "0"
        Me.txtNumMaxClientes.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtNumMaxClientes, "Numero Maximo de Clientes.")
        '
        'chkIncluirIva
        '
        Me.chkIncluirIva.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirIva.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncluirIva.Location = New System.Drawing.Point(303, 203)
        Me.chkIncluirIva.Name = "chkIncluirIva"
        Me.chkIncluirIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirIva.Size = New System.Drawing.Size(79, 17)
        Me.chkIncluirIva.TabIndex = 4
        Me.chkIncluirIva.Text = "Incluir Iva"
        Me.chkIncluirIva.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optDolares)
        Me.Frame3.Controls.Add(Me.optPesos)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(9, 130)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(373, 57)
        Me.Frame3.TabIndex = 13
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Moneda"
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Checked = True
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(65, 24)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(76, 17)
        Me.optDolares.TabIndex = 5
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dolares"
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(205, 24)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(89, 17)
        Me.optPesos.TabIndex = 3
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtNumMaxClientes)
        Me.Frame2.Controls.Add(Me.Label1)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(10, 69)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(372, 57)
        Me.Frame2.TabIndex = 10
        Me.Frame2.TabStop = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(57, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(151, 14)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Numero Maximo de Clientes :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me._Label2_1)
        Me.Frame1.Controls.Add(Me._Label2_0)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(10, 8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(372, 57)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Considerar Facturas"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(77, 23)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(101, 20)
        Me.dtpFechaInicial.TabIndex = 0
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(249, 23)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(101, 20)
        Me.dtpFechaFinal.TabIndex = 1
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_1.Location = New System.Drawing.Point(190, 27)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(53, 16)
        Me._Label2_1.TabIndex = 9
        Me._Label2_1.Text = "Hasta el :"
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_0.Location = New System.Drawing.Point(14, 27)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(57, 16)
        Me._Label2_0.TabIndex = 8
        Me._Label2_0.Text = "Desde el :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(9, 221)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(79, 17)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Texto Adicional"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(126, 323)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 119
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(11, 323)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 118
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmFactReportesMejoresClientes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(396, 371)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.chkIncluirIva)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.txtTextoAdicional)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(275, 244)
        Me.MaximizeBox = False
        Me.Name = "frmFactReportesMejoresClientes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Los Mejores Clientes"
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Sub Imprime()
        Dim RptFactFacturacionImpresionDeLosMejoresClientes As New RptFactFacturacionImpresionDeLosMejoresClientes

        Dim sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim TextoAdicional As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim Servidor As String
        Dim BasedeDatos As String
        Dim I As Object
        Dim Pos As Integer
        On Error GoTo ImprimeErr

        'Do While (sglTiempoCambio) <= 2.1
        'Loop
        'System.Windows.Forms.Application.DoEvents()

        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Sub
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaFinal.Focus()
            Exit Sub
        End If

        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Los Mejores Clientes")
        Dim fechaInicial1 As String = AgregarHoraAFecha(dtpFechaInicial.Value)
        Dim fechaFinal2 As String = AgregarHoraAFecha(dtpFechaFinal.Value)
        PeriodoReporte = "Del " & fechaInicial1 & " al " & fechaFinal2
        'PeriodoReporte = "Del " & Format(dtpFechaInicial.Value, "dd/MMM/yyyy") & " al " & Format(dtpFechaFinal.Value, "dd/MMM/yyyy")
        txtTextoAdicional.Text = ModEstandar.QuitaEnter(txtTextoAdicional.Text)
        TextoAdicional = txtTextoAdicional.Text
        'FechaInicial = Format(Month(dtpFechaInicial.Value), "00") & "/" & Format(VB.Day(dtpFechaInicial.Value), "00") & "/" & Format(Year(dtpFechaInicial.Value), "0000")
        'FechaFinal = Format(Month(dtpFechaFinal.Value), "00") & "/" & Format(VB.Day(dtpFechaFinal.Value), "00") & "/" & Format(Year(dtpFechaFinal.Value), "0000")
        FechaInicial = AgregarHoraAFecha(dtpFechaInicial.Value)
        FechaFinal = AgregarHoraAFecha(dtpFechaFinal.Value)


        If optDolares.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) <> 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Checked Then
            '''sql = "SELECT TOP " & Numerico(txtNumMaxClientes) & " Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & _
            '"SUM(Fac.Total + Fac.Redondeo) As Importe " & _
            '"FROM Facturas Fac " & _
            '"WHERE Fac.Estatus <> 'C' AND Fac.TipoFactura  <> 'E' AND Fac.CodCliente <> 1 " & _
            '"AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' " & _
            '"GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & _
            '"ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"

            sql = "SELECT TOP " & Numerico(txtNumMaxClientes.Text) & " Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM(Fac.Total + Fac.Redondeo) As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"
        ElseIf optDolares.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) <> 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            sql = "SELECT TOP " & Numerico(txtNumMaxClientes.Text) & " Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM((Fac.subTotal - fac.descuento) + Fac.Redondeo) As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM((Fac.subTotal - fac.descuento)) Desc"
        ElseIf optDolares.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) = 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Checked Then
            '''sql = "SELECT Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & _
            '"SUM(Fac.Total + Fac.Redondeo) As Importe " & _
            '"FROM Facturas Fac " & _
            '"WHERE Fac.Estatus <> 'C' AND Fac.TipoFactura  <> 'E' AND Fac.CodCliente <> 1 " & _
            '"AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' " & _
            '"GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & _
            '"ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"

            sql = "SELECT Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM(Fac.Total + Fac.Redondeo) As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"
        ElseIf optDolares.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) = 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            sql = "SELECT Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM((Fac.subTotal - fac.descuento) + Fac.Redondeo) As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM((Fac.subTotal - fac.descuento) + Fac.Redondeo) Desc"
        ElseIf optPesos.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) <> 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Checked Then
            '''sql = "SELECT TOP " & Numerico(txtNumMaxClientes) & " Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & _
            '"SUM(Fac.Total + Fac.Redondeo) * " & gcurCorpoTIPOCAMBIODOLAR & " As Importe " & _
            '"FROM Facturas Fac " & _
            '"WHERE Fac.Estatus <> 'C' AND Fac.TipoFactura  <> 'E' AND Fac.CodCliente <> 1 " & _
            '"AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' " & _
            '"GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & _
            '"ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"

            sql = "SELECT TOP " & Numerico(txtNumMaxClientes.Text) & " Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM(Fac.Total + Fac.Redondeo) * " & gcurCorpoTIPOCAMBIODOLAR & " As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"
        ElseIf optPesos.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) <> 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            sql = "SELECT TOP " & Numerico(txtNumMaxClientes.Text) & " Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM((Fac.subTotal - fac.descuento) + Fac.Redondeo) * " & gcurCorpoTIPOCAMBIODOLAR & " As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM((Fac.subTotal - fac.descuento) + Fac.Redondeo) Desc"
        ElseIf optPesos.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) = 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Checked Then
            '''sql = "SELECT Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & _
            '"SUM(Fac.Total + Fac.Redondeo) * " & gcurCorpoTIPOCAMBIODOLAR & " As Importe " & _
            '"FROM Facturas Fac " & _
            '"WHERE Fac.Estatus <> 'C' AND Fac.TipoFactura  <> 'E' AND Fac.CodCliente <> 1 " & _
            '"AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' " & _
            '"GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & _
            '"ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"

            sql = "SELECT Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM(Fac.Total + Fac.Redondeo) * " & gcurCorpoTIPOCAMBIODOLAR & " As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM(Fac.Total + Fac.Redondeo) Desc"
        ElseIf optPesos.Checked And CDbl(Numerico(txtNumMaxClientes.Text)) = 0 And chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            sql = "SELECT Fac.CodCliente,Fac.Nombre,Fac.Rfc,MAX(Fac.FechaFactura) AS FechaUltimaCompra,Count(Fac.FolioFactura) AS Operaciones," & "SUM((Fac.subTotal - fac.descuento) + Fac.Redondeo) * " & gcurCorpoTIPOCAMBIODOLAR & " As Importe " & "FROM Facturas Fac " & "WHERE Fac.Estatus <> 'C' AND Fac.CodCliente <> 1 " & "AND Fac.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Fac.DesgloseIva = 1 " & "GROUP BY Fac.CodCliente,Fac.Nombre,Fac.Rfc " & "ORDER BY SUM((Fac.subTotal - fac.descuento) + Fac.Redondeo) Desc"
        End If

        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No Existe Información En Este Rango de Fechas...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            'frmReportes.Report = RptFactFacturacionImpresionDeLosMejoresClientes
            RptFactFacturacionImpresionDeLosMejoresClientes.SetDataSource(frmReportes.rsReport)
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte", "TextoAdicional"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, TextoAdicional}
        frmReportes.Text = "los Mejores Clientes"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = RptFactFacturacionImpresionDeLosMejoresClientes
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
    End Sub

    Sub Limpiar()
        dtpFechaInicial.Value = Now
        dtpFechaFinal.Value = Now
        txtNumMaxClientes.Text = CStr(0)
        optDolares.Checked = True
        optPesos.Checked = False
        txtTextoAdicional.Text = ""
        dtpFechaInicial.Focus()
        chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Unchecked
    End Sub

    Private Sub dtpFechaFinal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaFinal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaFinal.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmFactReportesMejoresClientes_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmFactReportesMejoresClientes_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmFactReportesMejoresClientes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dtpFechaInicial" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmFactReportesMejoresClientes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmFactReportesMejoresClientes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        dtpFechaInicial.MinDate = C_FECHAINICIAL
        dtpFechaInicial.MaxDate = C_FECHAFINAL
        dtpFechaFinal.MinDate = C_FECHAINICIAL
        dtpFechaFinal.MaxDate = C_FECHAFINAL
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        chkIncluirIva.CheckState = System.Windows.Forms.CheckState.Unchecked
    End Sub

    Private Sub frmFactReportesMejoresClientes_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
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

    Private Sub frmFactReportesMejoresClientes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'InitializeComponent()
        'ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        'ModEstandar.LimpiaDescBarraEstado()
        ''Me = Nothing
        'IsNothing(Me)
        'MDIMenuPrincipalCorpo.mnuFacturacionRptFactOpc(3).Enabled = True
    End Sub

    Private Sub optDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub optPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.Enter
        Pon_Tool()
    End Sub
    Private Sub txtNumMaxClientes_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumMaxClientes.TextChanged
        If Trim(txtNumMaxClientes.Text) = "" Then
            txtNumMaxClientes.Text = CStr(0)
        End If
    End Sub

    Private Sub txtNumMaxClientes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumMaxClientes.Enter
        Pon_Tool()
        SelTextoTxt(txtNumMaxClientes)
    End Sub

    Private Sub txtNumMaxClientes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNumMaxClientes.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTextoAdicional_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTextoAdicional.Enter
        Pon_Tool()
        SelTextoTxt(txtTextoAdicional)
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class
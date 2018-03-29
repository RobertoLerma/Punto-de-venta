Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasRelacionGastos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkFueraEnter As System.Windows.Forms.CheckBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents flexGastos As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblDesc As System.Windows.Forms.Label
    'Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents Line2 As System.Windows.Forms.Label
    Public WithEvents Line1 As System.Windows.Forms.Label

    Dim mblnSalir As Boolean
    Dim FechaInicial As String
    Dim FechaFinal As String
    Dim Fecha As String
    Dim Moneda As String
    Dim Sucursales() As Integer
    Friend WithEvents GroupBox1 As GroupBox
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim NumSucursales As Integer


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasRelacionGastos))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.chkFueraEnter = New System.Windows.Forms.CheckBox()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.flexGastos = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame5.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.flexGastos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(78, 15)
        Me.cmbMes.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(130, 21)
        Me.cmbMes.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(264, 15)
        Me.cmbAño.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(80, 21)
        Me.cmbAño.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Checked = True
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(12, 14)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(67, 20)
        Me.optPesos.TabIndex = 2
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optPesos, "Muestra los Importes en Pesos")
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(12, 32)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(67, 20)
        Me.optDolares.TabIndex = 3
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dolares"
        Me.ToolTip1.SetToolTip(Me.optDolares, "Muestra los Importes en Dolares")
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'chkFueraEnter
        '
        Me.chkFueraEnter.BackColor = System.Drawing.SystemColors.Control
        Me.chkFueraEnter.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFueraEnter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFueraEnter.Location = New System.Drawing.Point(383, 361)
        Me.chkFueraEnter.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFueraEnter.Name = "chkFueraEnter"
        Me.chkFueraEnter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFueraEnter.Size = New System.Drawing.Size(62, 17)
        Me.chkFueraEnter.TabIndex = 13
        Me.chkFueraEnter.Text = "Check1"
        Me.chkFueraEnter.UseVisualStyleBackColor = False
        Me.chkFueraEnter.Visible = False
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.cmbMes)
        Me.Frame5.Controls.Add(Me.cmbAño)
        Me.Frame5.Controls.Add(Me.Label1)
        Me.Frame5.Controls.Add(Me.Label2)
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame5.Location = New System.Drawing.Point(11, 20)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(391, 44)
        Me.Frame5.TabIndex = 10
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Periodo"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(40, 17)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(34, 17)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Mes :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(228, 16)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(31, 17)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Año :"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optPesos)
        Me.Frame2.Controls.Add(Me.optDolares)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(308, 86)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(94, 56)
        Me.Frame2.TabIndex = 9
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Moneda"
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Location = New System.Drawing.Point(21, 147)
        Me.Line2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(388, 1)
        Me.Line2.TabIndex = 11
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(14, 83)
        Me.Line1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(388, 1)
        Me.Line1.TabIndex = 12
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(12, 44)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(65, 20)
        Me.txtFlex.TabIndex = 5
        Me.txtFlex.Visible = False
        '
        'flexGastos
        '
        Me.flexGastos.DataSource = Nothing
        Me.flexGastos.Location = New System.Drawing.Point(12, 21)
        Me.flexGastos.Margin = New System.Windows.Forms.Padding(2)
        Me.flexGastos.Name = "flexGastos"
        Me.flexGastos.OcxState = CType(resources.GetObject("flexGastos.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexGastos.Size = New System.Drawing.Size(400, 156)
        Me.flexGastos.TabIndex = 4
        '
        'lblDesc
        '
        Me.lblDesc.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDesc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblDesc.Location = New System.Drawing.Point(12, 338)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesc.Size = New System.Drawing.Size(431, 21)
        Me.lblDesc.TabIndex = 6
        Me.lblDesc.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.flexGastos)
        Me.GroupBox1.Controls.Add(Me.txtFlex)
        Me.GroupBox1.Location = New System.Drawing.Point(17, 151)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(428, 182)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Registro de las Cuentas de Gastos"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 393)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 109
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(12, 393)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 108
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(242, 394)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 107
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasRelacionGastos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(465, 441)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Line2)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.chkFueraEnter)
        Me.Controls.Add(Me.lblDesc)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(323, 127)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRelacionGastos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Relación de Gastos por Periodo"
        Me.Frame5.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        CType(Me.flexGastos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub


    Sub Buscar()
        'On Error GoTo Merr
        Try
            Dim strSQL As String
            Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
            Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
            Dim strControlActual As String 'Nombre del control actual
            Dim strDesc As String
            Dim I As Object
            Dim J As Integer
            If flexGastos.Row > 1 Then
                With flexGastos
                    For I = 1 To .Row - 1
                        If Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "" Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    Next
                    If (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                        If .Col = 3 Or .Col = 4 Then Exit Sub
                    ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                        If .Col = 1 Or .Col = 2 Then Exit Sub
                    End If
                End With
            End If
            If flexGastos.Col > 0 And Trim(flexGastos.get_TextMatrix(flexGastos.Row, 0)) = "" Then Exit Sub
            Me.Tag = "RELGASTOS"
            With flexGastos
                If .Col = 0 Then
                    strControlActual = "DESCRIPCION SUCURSAL"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 1 Then
                    strControlActual = "CODIGO AGRUPADOR"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 2 Then
                    strControlActual = "DESCRIPCION AGRUPADOR"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 3 Then
                    strControlActual = "CODIGO RUBRO"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 4 Then
                    strControlActual = "DESCRIPCION RUBRO"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                End If
                If Me.ActiveControl.Name = "txtFlex" Then
                    strDesc = Trim(txtFlex.Text)
                Else
                    Exit Sub
                End If
                Select Case strControlActual
                    Case "DESCRIPCION SUCURSAL"
                        strCaptionForm = "Consulta de Sucursales"
                        If Me.ActiveControl.Name = "txtFlex" Then
                            gStrSql = "SELECT Descalmacen AS DESCRIPCION,RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO " & "From Catalmacen WHERE tipoalmacen = 'P' ORDER BY DescAlmacen"
                        ElseIf Me.ActiveControl.Name = "flexGastos" Then
                            gStrSql = "SELECT Descalmacen AS DESCRIPCION,RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO " & "From Catalmacen WHERE tipoalmacen = 'P' ORDER BY DescAlmacen"
                        End If
                    Case "CODIGO AGRUPADOR"
                        strCaptionForm = "Consulta de Agrupadores de Origen y Aplicación"
                        gStrSql = "SELECT RIGHT('0000' + LTRIM(CodOrigenAplicR),4) AS AGRUPADOR, DescOrigenAplicR AS DESCRIPCION " & "FROM CatOrigenAplicRecursos ORDER BY CodOrigenAplicR"
                    Case "CODIGO RUBRO"
                        strCaptionForm = "Consulta de Rubros de Origen y Aplicación"
                        If Trim(.get_TextMatrix(.Row, 1)) = "" And Len(Trim(.get_TextMatrix(.Row, 1))) < 4 Then
                            gStrSql = "SELECT RIGHT('000000' + LTRIM(CodRubro),6) AS RUBRO, DescRubro AS DESCRIPCION " & "FROM CatRubrosOrigenAplicRecursos ORDER BY CodRubro"
                        ElseIf Trim(.get_TextMatrix(.Row, 1)) <> "" And Len(Trim(.get_TextMatrix(.Row, 1))) = 4 Then
                            gStrSql = "SELECT RIGHT('000000' + LTRIM(R.CodRubro),6) AS RUBRO, R.DescRubro AS DESCRIPCION " & "FROM CatRubrosOrigenAplicRecursos R,CatOrigenAplicRecursos A WHERE A.CodOrigenAplicR = " & Numerico(.get_TextMatrix(.Row, 1)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR ORDER BY R.CodRubro"
                        End If
                    Case "DESCRIPCION AGRUPADOR"
                        strCaptionForm = "Consulta de Agrupadores de Origen y Aplicación"
                        If Trim(.get_TextMatrix(.Row, 3)) = "" And Len(Trim(.get_TextMatrix(.Row, 3))) < 6 Then
                            gStrSql = "SELECT DescOrigenAplicR AS DESCRIPCION, RIGHT('0000' + LTRIM(CodOrigenAplicR),4) AS AGRUPADOR " & "FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & strDesc & "%' ORDER BY DescOrigenAplicR"
                        ElseIf Trim(.get_TextMatrix(.Row, 3)) <> "" And Len(Trim(.get_TextMatrix(.Row, 3))) = 6 Then
                            gStrSql = "SELECT A.DescOrigenAplicR AS DESCRIPCION, RIGHT('0000' + LTRIM(R.CodOrigAplicR),4) AS AGRUPADOR " & "FROM CatOrigenAplicRecursos A ,CatRubrosOrigenAplicRecursos R WHERE A.DescOrigenAplicR LIKE '" & strDesc & "%' AND R.CodRubro = " & Numerico(.get_TextMatrix(.Row, 3)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR GROUP BY R.CodOrigAplicR,A.DescOrigenAplicR ORDER BY A.DescOrigenAplicR"
                        End If
                    Case "DESCRIPCION RUBRO"
                        strCaptionForm = "Consulta de Rubros de Origen y Aplicación"
                        If Trim(.get_TextMatrix(.Row, 1)) = "" And Len(Trim(.get_TextMatrix(.Row, 1))) < 4 Then
                            gStrSql = "SELECT DescRubro AS DESCRIPCION, RIGHT('000000' + LTRIM(CodRubro),6) AS RUBRO " & "FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos WHERE DescRubro LIKE '" & strDesc & "%' AND CodOrigenAplicR = CodOrigAplicR ORDER BY DescRubro"
                        ElseIf Trim(.get_TextMatrix(.Row, 1)) <> "" And Len(Trim(.get_TextMatrix(.Row, 1))) = 4 Then
                            gStrSql = "SELECT R.DescRubro AS DESCRIPCION, RIGHT('000000' + LTRIM(R.CodRubro),6) AS RUBRO " & "FROM CatRubrosOrigenAplicRecursos R,CatOrigenAplicRecursos A WHERE R.DescRubro LIKE '" & strDesc & "%' AND A.CodOrigenAplicR = " & Numerico(.get_TextMatrix(.Row, 1)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR ORDER BY DescRubro"
                        End If
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
                '    Load(FrmConsultas)
                '    With FrmConsultas.Flexdet
                '        Select Case strControlActual
                '            Case "DESCRIPCION SUCURSAL"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0,  , 4800) 'Columna de la Descripción
                '                .set_ColWidth(1,  , 900) 'Columna del Código
                '            Case "CODIGO AGRUPADOR"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0,  , 1300) 'Columna del Código Agrupador
                '                .set_ColWidth(1,  , 4500) 'Columna de la Descripción del Agrupador
                '            Case "CODIGO RUBRO"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0,  , 1300) 'Columna del Codigo del Rubro
                '                .set_ColWidth(1,  , 4500) 'Columna de la Descripción del Rubro
                '            Case "DESCRIPCION AGRUPADOR"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0,  , 4500) 'Columna de la Descripción del Agrupador
                '                .set_ColWidth(1,  , 1300) 'Columna del Codigo del Agrupador
                '            Case "DESCRIPCION RUBRO"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0,  , 4500) 'Columna de la Descripción del Rubro
                '                .set_ColWidth(1,  , 1300) 'Columna del Codigo del Rubro
                '        End Select
                '    End With
            End With
            'chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Checked
            'FrmConsultas.ShowDialog()
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            Select Case flexGastos.Col
                Case 0 'Descripcion de la Sucursal
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
                Case 1 'Codigo del Agrupador
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 4
                Case 2 'Descripción del Agrupador
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
                Case 3 'Codigo del Rubro
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 6
                Case 4 'Descripción del Rubro
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
            End Select
        End With
    End Sub

    Function DescripcionAgrupador() As Boolean
        On Error GoTo Err_Renamed
        DescripcionAgrupador = False
        'If Trim(flexgastos.TextMatrix(flexgastos.Row, 3)) = "" Then
        gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR like '" & Trim(txtFlex.Text) & "%'"
        'End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = Trim(RsGral.Fields("DescOrigenAplicR").Value)
            flexGastos.set_TextMatrix(flexGastos.Row, 1, ("0000" & CStr(RsGral.Fields("CodOrigenAplicR").Value)))
            If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 2)) <> Trim(txtFlex.Text) Then
                flexGastos.set_TextMatrix(flexGastos.Row, 3, "")
                flexGastos.set_TextMatrix(flexGastos.Row, 4, "")
            End If
            DescripcionAgrupador = True
            txtFlex_Leave(txtFlex, New System.EventArgs())
        Else
            MsgBox("Descripción Inexistente Favor de Verificar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexGastos.Col = 2
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DescripcionRubro() As Boolean
        On Error GoTo Err_Renamed
        DescripcionRubro = False
        If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) = "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE DescRubro like '" & Trim(txtFlex.Text) & "%' GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        ElseIf Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) <> "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE DescRubro like '" & Trim(txtFlex.Text) & "%' and codorigaplicr = " & Numerico(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1))) & " GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = Trim(RsGral.Fields("DescRubro").Value)
            flexGastos.set_TextMatrix(flexGastos.Row, 3, ("000000" & CStr(RsGral.Fields("CodRubro").Value)))
            txtFlex_Leave(txtFlex, New System.EventArgs())
            DescripcionRubro = True
        Else
            MsgBox("Descripción Inexistente Favor de Investigar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DevuelveQuery() As String
        Dim Sql As String
        Dim subsql As String
        Dim I As Object
        Dim J As Integer
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim blnExiste As Boolean
        Dim FormaQuery As Boolean

        NumSucursales = 0
        ModCorporativo.ObtenerLimitedeFechas(CShort((Trim(cmbMes.Text))), CShort(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
        With flexGastos
            If Vacio() Then
                DevuelveQuery = ""
                Exit Function
            End If
            Sql = ""
            subsql = ""
            gStrSql = "Select count(*) NumSuc from catalmacen where tipoalmacen = 'P'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                NumSucursales = RsGral.Fields("numsuc").Value
                ReDim Sucursales(NumSucursales)
            End If
            gStrSql = "Select codalmacen from catalmacen where tipoalmacen = 'P'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            I = 1
            If RsGral.RecordCount > 0 Then
                Do While Not RsGral.EOF
                    Sucursales(I) = RsGral.Fields("CodAlmacen").Value
                    RsGral.MoveNext()
                    I = I + 1
                Loop
            End If
            Sql = "select a.codsucursal,b.descalmacen,a.codorigenaplicr,a.descorigenaplicr,a.codrubro,a.descrubro,sum(" & IIf(optPesos.Checked = True, "a.importepesosconimpuesto", "a.impdolaresconimpuesto") & ") as importe from ("
            For J = 1 To NumSucursales
                blnExiste = False
                FormaQuery = False
                For I = 1 To .Rows - 1
                    If Trim(.get_TextMatrix(I, 0)) <> "" Or Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "" Or Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "" Then
                        If Sucursales(J) = CDbl(Numerico(.get_TextMatrix(I, 5))) Then
                            blnExiste = True
                            If Trim(subsql) = "" And Not FormaQuery Then
                                subsql = "(select " & Numerico(.get_TextMatrix(I, 5)) & " as codsucursal,a.codorigenaplicr,c.descorigenaplicr,a.codrubro,d.descrubro,'H' as tipomovto,'Gastos' as descripcion," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaresconimpuesto," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaressinimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesosconimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesossinimpuesto," & "b.fechamovto as fecha " & "from movimientosorigenaplic a inner join movimientosbancarios b " & "on a.foliomovto = b.foliomovto inner join catorigenaplicrecursos c on a.codorigenaplicr = c.codorigenaplicr " & "inner join catrubrosorigenaplicrecursos d on a.codorigenaplicr = d.codorigaplicr and a.codrubro = d.codrubro " & "where a.estatus <> 'C' and ( (" & IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & " and a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")", IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) = "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & ") ", "a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")"))
                                FormaQuery = True
                            ElseIf Trim(subsql) <> "" And Not FormaQuery Then
                                subsql = subsql & " union " & "(select " & Numerico(.get_TextMatrix(I, 5)) & " as codsucursal,a.codorigenaplicr,c.descorigenaplicr,a.codrubro,d.descrubro,'H' as tipomovto,'Gastos' as descripcion," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaresconimpuesto," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaressinimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesosconimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesossinimpuesto," & "b.fechamovto as fecha " & "from movimientosorigenaplic a inner join movimientosbancarios b " & "on a.foliomovto = b.foliomovto inner join catorigenaplicrecursos c on a.codorigenaplicr = c.codorigenaplicr " & "inner join catrubrosorigenaplicrecursos d on a.codorigenaplicr = d.codorigaplicr and a.codrubro = d.codrubro " & "where a.estatus <> 'C' and ( (" & IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & " and a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")", IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) = "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & ") ", "a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")"))
                                FormaQuery = True
                            ElseIf FormaQuery Then
                                subsql = subsql & " or (" & IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & " and a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")", IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) = "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & ") ", "a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")"))
                            End If
                        End If
                    Else
                        If blnExiste Then
                            subsql = subsql & " ) group by a.codorigenaplicr,c.descorigenaplicr,a.codrubro,d.descrubro,b.fechamovto) "
                        End If
                        Exit For
                    End If
                Next
            Next
        End With
        Sql = Sql & subsql & ") a inner join catalmacen b on a.codsucursal = b.codalmacen " & "where a.fecha between '" & FechaInicial & "' and '" & FechaFinal & "' " & "group by a.codsucursal,b.descalmacen,a.codorigenaplicr,a.descorigenaplicr,a.codrubro,a.descrubro "
        DevuelveQuery = Sql
    End Function

    Sub EliminarLinea()
        Dim Ren As Integer
        With flexGastos
            If Trim(.get_TextMatrix(.Row, 0)) = "" And Trim(.get_TextMatrix(.Row, 1)) = "" And Trim(.get_TextMatrix(.Row, 2)) = "" And Trim(.get_TextMatrix(.Row, 3)) = "" And Trim(.get_TextMatrix(.Row, 4)) = "" Then Exit Sub
        End With
        Select Case MsgBox("¿Desea Eliminar Esta Informacion?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                Ren = flexGastos.Rows
                flexGastos.RemoveItem(flexGastos.Row)
                flexGastos.Rows = Ren
                flexGastos.Focus()
            Case MsgBoxResult.No, MsgBoxResult.Cancel
                flexGastos.Focus()
                Exit Sub
        End Select
    End Sub

    Sub EncabezadoFlex()
        With flexGastos
            .Clear()
            .Rows = 11
            .set_Cols(0, 6)
            .set_ColWidth(0, 0, 1200)
            .set_ColWidth(1, 0, 1000)
            .set_ColWidth(2, 0, 2020)
            .set_ColWidth(3, 0, 1000)
            .set_ColWidth(4, 0, 2020)
            .set_ColWidth(5, 0, 0)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Sucursal"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Agrup."
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Descripción"
            .Col = 3
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Rubro"
            .Col = 4
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Descripción"
            .Row = 1
            .Col = 0
        End With
    End Sub

    Sub Imprime()

        Dim rptRelaciondeGastos As New rptRelaciondeGastos

        Dim Query, sql1 As Object
        Dim sql2 As String
        Dim Moneda As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim Periodo As String
        Dim MonedaExp As String

        'On Error GoTo ImprimeErr

        If Not ValidaGrid() Then
            MsgBox("No se ha capturado de forma adecuada la información de las cuentas de gastos, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            flexGastos.Focus()
            Exit Sub
        End If
        If optPesos.Checked = True Then
            Moneda = "P"
        ElseIf optDolares.Checked = True Then
            Moneda = "D"
        End If

        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = "RELACIÓN DE GASTOS POR PERIODO"
        Periodo = "CORRESPONDIENTE AL MES DE " & UCase(Mid(cmbMes.Text, 6, 12))

        If Moneda = "P" Then
            MonedaExp = "** Los importes estan expresados en pesos"
        ElseIf Moneda = "D" Then
            MonedaExp = "** Los importes estan expresados en dólares"
        End If
        Cmd.CommandTimeout = 300
        Query = DevuelveQuery()
        If Len(Query) > 8000 Then
            sql1 = (Query)
            'sql2 = (Query, Len(Query))
        Else
            sql1 = Query
            sql2 = ""
        End If

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_DatosSql"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, sql1))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, sql2))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe información para mostrar en este periodo, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Limpiar()
            Exit Sub
        Else
            'frmReportes.Report = rptRelaciondeGastos
            rptRelaciondeGastos.SetDataSource(frmReportes.rsReport)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = RsGral
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte", "Moneda"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Periodo, MonedaExp}
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.Text = "Reporte de Relación de Gastos"
        frmReportes.reporteActual = rptRelaciondeGastos
        frmReportes.ShowDialog()
        Cursor = System.Windows.Forms.Cursors.Default
        Cmd.CommandTimeout = 90
        Exit Sub

ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")

    End Sub

    Sub Limpiar()
        Nuevo()
        cmbMes.Focus()
    End Sub

    Sub LlenaAños()
        Dim I As Integer
        cmbAño.Items.Clear()
        For I = Year(Today) To 2001 Step -1
            cmbAño.Items.Add(CStr(I))
        Next
    End Sub

    Function LlenaDatosAgrupador() As Boolean
        On Error GoTo Err_Renamed
        LlenaDatosAgrupador = False
        If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3)) = "" And Len(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3))) < 6 Then
            gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR = " & Numerico(txtFlex.Text)
        ElseIf Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3)) <> "" And Len(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3))) = 6 Then
            'gStrSql = "SELECT * " & _
            ''"FROM CatOrigenAplicRecursos A, CatRubrosOrigenAplicRecursos R WHERE R.CodRubro = " & Numerico(flexgastos.TextMatrix(flexgastos.Row, 2)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR AND A.Aplicacion = '" & gstrMovimiento & "'"
            gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR = " & Numerico(txtFlex.Text)
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = ("0000" & CStr(RsGral.Fields("CodOrigenAplicR").Value))
            flexGastos.set_TextMatrix(flexGastos.Row, 2, Trim(RsGral.Fields("DescOrigenAplicR").Value))
            If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) <> Trim(txtFlex.Text) Then
                flexGastos.set_TextMatrix(flexGastos.Row, 3, "")
                flexGastos.set_TextMatrix(flexGastos.Row, 4, "")
            End If
            LlenaDatosAgrupador = True
            txtFlex_Leave(txtFlex, New System.EventArgs())
        Else
            MsgBox("Codigo Inexistente Favor de Verificar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexGastos.Col = 1
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function LlenaDatosRubro() As Boolean
        On Error GoTo Err_Renamed
        LlenaDatosRubro = False
        If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) = "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE CodRubro = " & Numerico(txtFlex.Text) & " GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        ElseIf Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) <> "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE CodRubro = " & Numerico(txtFlex.Text) & " and codorigaplicr = " & Numerico(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1))) & " GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = ("000000" & CStr(RsGral.Fields("CodRubro").Value))
            flexGastos.set_TextMatrix(flexGastos.Row, 4, Trim(RsGral.Fields("DescRubro").Value))
            txtFlex_Leave(txtFlex, New System.EventArgs())
            LlenaDatosRubro = True
        Else
            MsgBox("Codigo Inexistente Favor de Investigar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function LlenaSucursal() As Boolean
        On Error GoTo Err_Renamed
        LlenaSucursal = False
        gStrSql = "SELECT * FROM CatAlmacen WHERE DescAlmacen like '" & Trim(txtFlex.Text) & "%' and TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            flexGastos.set_TextMatrix(flexGastos.Row, 5, RsGral.Fields("CodAlmacen").Value)
            LlenaSucursal = True
            txtFlex_Leave(txtFlex, New System.EventArgs())
        Else
            MsgBox("Descripción inexistente, Favor de verificar..", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexGastos.Col = 0
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Nuevo()
        LlenaAños()
        cmbMes.SelectedIndex = Month(Today) - 1
        cmbAño.SelectedIndex = 0
        cmbMes.Enabled = True
        optPesos.Checked = True
        EncabezadoFlex()
        chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Unchecked
        NumSucursales = 0
    End Sub

    Function Vacio() As Boolean
        Dim I As Integer
        Vacio = False
        With flexGastos
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Or Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "" Or Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "" Then
                    Exit Function
                End If
            Next
        End With
        Vacio = True
    End Function

    Function ValidaCodigos(ByRef RenActual As Integer) As Boolean
        Dim I As Integer
        ValidaCodigos = False
        With flexGastos
            For I = 1 To .Rows - 1
                If I <> RenActual Then
                    If RenActual = 1 Then
                        If Trim(flexGastos.get_TextMatrix(.Row, 1)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 2)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 3)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 4)) = "" Then Exit Function
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "" Then
                            If Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(1, 1)) And Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(1, 3)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And (Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "") And (Trim(.get_TextMatrix(I, 3)) = "" And Trim(.get_TextMatrix(I, 4)) = "") Then
                            If Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(1, 1)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And (Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "") And (Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "") Then
                            If Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(1, 3)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        End If
                    Else
                        If Trim(flexGastos.get_TextMatrix(.Row, 1)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 2)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 3)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 4)) = "" Then Exit Function
                        If Trim(flexGastos.get_TextMatrix(1, 0)) <> "" And Trim(flexGastos.get_TextMatrix(1, 1)) <> "" And Trim(flexGastos.get_TextMatrix(1, 2)) <> "" And Trim(flexGastos.get_TextMatrix(1, 3)) <> "" And Trim(flexGastos.get_TextMatrix(1, 4)) <> "" Then
                            If Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(.Row, 3)) And Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(.Row, 1)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf (Trim(flexGastos.get_TextMatrix(1, 0)) <> "" And Trim(flexGastos.get_TextMatrix(1, 1)) <> "" And Trim(flexGastos.get_TextMatrix(1, 2)) <> "") And (Trim(flexGastos.get_TextMatrix(1, 3)) = "" And Trim(flexGastos.get_TextMatrix(1, 4)) = "") Then
                            If Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(.Row, 1)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf (Trim(flexGastos.get_TextMatrix(1, 1)) = "" And Trim(flexGastos.get_TextMatrix(1, 2)) = "") And (Trim(flexGastos.get_TextMatrix(1, 0)) <> "" And Trim(flexGastos.get_TextMatrix(1, 3)) <> "" And Trim(flexGastos.get_TextMatrix(1, 4)) <> "") Then
                            If Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(.Row, 3)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        End With
    End Function

    Function ValidaGrid() As Boolean
        Dim I As Integer
        ValidaGrid = False
        If Vacio() Then
            'ValidaGrid = True
            Exit Function
        End If
        With flexGastos
            If Trim(.get_TextMatrix(1, 0)) = "" And (Trim(.get_TextMatrix(1, 0)) <> "" Or (Trim(.get_TextMatrix(1, 1)) = "" Or Trim(.get_TextMatrix(1, 2)) = "" Or Trim(.get_TextMatrix(1, 3)) = "" Or Trim(.get_TextMatrix(1, 4)) = "")) Then
                Exit Function
                'ElseIf Trim(.TextMatrix(1, 0)) <> "" And (Trim(.TextMatrix(1, 0)) = "" Or (Trim(.TextMatrix(1, 1)) = "" Or Trim(.TextMatrix(1, 2)) = "" Or Trim(.TextMatrix(1, 3)) = "" Or Trim(.TextMatrix(1, 4)) = "")) Then

                Exit Function
            End If
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Or Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "" Or Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "" Then
                    If Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "" Then
                        If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 2)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Or Trim(.get_TextMatrix(I, 4)) = "" Then Exit Function
                    ElseIf (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                        If (Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 2)) = "") And (Trim(.get_TextMatrix(I, 3)) = "" And Trim(.get_TextMatrix(I, 4)) = "") Then Exit Function
                        If (Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "") Then Exit Function
                    ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                        If (Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "") And (Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Or Trim(.get_TextMatrix(I, 4)) = "") Then Exit Function
                        If (Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "") Then Exit Function
                    End If
                Else
                    ValidaGrid = True
                    Exit Function
                End If
            Next
        End With
    End Function

    Private Sub cmbAño_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmbAño.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            cmbMes.Focus()
        End If
    End Sub

    Private Sub flexGastos_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.ClickEvent
        txtFlex.Visible = False
    End Sub

    Private Sub flexGastos_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.DblClick
        flexGastos_KeyPressEvent(flexGastos, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub flexGastos_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.EnterCell
        With flexGastos
            If .Col = 0 Then
                lblDesc.Text = Trim(.get_TextMatrix(.Row, 0))
            ElseIf .Col = 1 Or .Col = 2 Then
                lblDesc.Text = Trim(.get_TextMatrix(.Row, 2))
            ElseIf .Col = 3 Or .Col = 4 Then
                lblDesc.Text = Trim(.get_TextMatrix(.Row, 4))
            End If
        End With
    End Sub

    Private Sub flexGastos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.Enter
        txtFlex.Visible = False
        Pon_Tool()
        flexGastos_EnterCell(flexGastos, New System.EventArgs())
    End Sub

    Private Sub flexGastos_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexGastos.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete Then
            EliminarLinea()
        End If
    End Sub

    Private Sub flexGastos_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexGastos.KeyPressEvent
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Unchecked
            Exit Sub
        End If
        Dim lonR, lonI As Integer
        Dim I As Integer
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then
            'Verifica si se puede capturar la fila
            If flexGastos.Row > 1 Then
                With flexGastos
                    For I = 1 To .Row - 1
                        If Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "" Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    Next
                    If (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                        If .Col = 3 Or .Col = 4 Then Exit Sub
                    ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                        If .Col = 1 Or .Col = 2 Then Exit Sub
                    End If
                End With
            End If
            'Edita el campo sólo si es Editable
            If flexGastos.Col > 0 And Trim(flexGastos.get_TextMatrix(flexGastos.Row, 0)) = "" Then Exit Sub
            If (flexGastos.Col = 3 Or flexGastos.Col = 4) And (Trim(flexGastos.get_TextMatrix(flexGastos.Row, 0)) = "" Or Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) = "" Or Trim(flexGastos.get_TextMatrix(flexGastos.Row, 2)) = "") Then Exit Sub
            If flexGastos.Row >= 1 And flexGastos.Col < 5 Then
                If flexGastos.Col = 0 And flexGastos.Col = 1 Or flexGastos.Col = 3 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                End If
                CambiarFormatoTxtenCaptura()
                MSHFlexGridEdit(flexGastos, txtFlex, eventArgs.keyAscii)
                If Len(Trim(txtFlex.Text)) = 1 Then
                    'System.Windows.Forms.SendKeys.Send("{right}")
                End If
            End If
        ElseIf eventArgs.keyAscii = System.Windows.Forms.Keys.Escape Then
            Exit Sub
        Else
            Exit Sub
        End If
    End Sub

    Private Sub flexGastos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.Leave
        lblDesc.Text = ""
    End Sub

    Private Sub frmVtasRelacionGastos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRelacionGastos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRelacionGastos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "txtFlex" Then Exit Sub
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "cmbMes" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasRelacionGastos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRelacionGastos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Nuevo()
    End Sub

    Private Sub frmVtasRelacionGastos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If mblnSalir Then
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSalir = False
                    Cancel = 1
                    cmbMes.Focus()
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRelacionGastos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        With flexGastos
            If KeyCode = System.Windows.Forms.Keys.Return Then
                Select Case .Col
                    Case 0, 1, 2, 3, 4
                        If .Col = 0 And Trim(txtFlex.Text) <> "" Then
                            If LlenaSucursal() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 1
                                    'txtFlex.Visible = False
                                    Exit Sub
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 1 And Trim(txtFlex.Text) <> "" Then
                            If LlenaDatosAgrupador() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 3
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 2 And Trim(txtFlex.Text) <> "" Then
                            If DescripcionAgrupador() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 3
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 3 And Trim(txtFlex.Text) <> "" Then
                            If LlenaDatosRubro() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 0
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                                txtFlex.Visible = False
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 4 And Trim(txtFlex.Text) <> "" Then
                            If DescripcionRubro() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 0
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                                txtFlex.Visible = False
                            Else
                                Exit Sub
                            End If
                        ElseIf Trim(txtFlex.Text) = "" Then
                            If .Col = 0 Then
                                'txtFlex.Visible = False
                                'flexGastos.Focus()
                                Exit Sub
                            ElseIf .Col = 1 Or .Col = 2 Then
                                .set_TextMatrix(.Row, 1, "")
                                .set_TextMatrix(.Row, 2, "")
                                .set_TextMatrix(.Row, 3, "")
                                .set_TextMatrix(.Row, 4, "")
                                txtFlex.Visible = False
                                If ValidaCodigos(.Row) Then
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "")
                                    .Col = 3
                                End If
                                flexGastos.Focus()
                                Exit Sub
                            ElseIf .Col = 3 Or .Col = 4 Then
                                .set_TextMatrix(.Row, 3, "")
                                .set_TextMatrix(.Row, 4, "")
                                txtFlex.Visible = False
                                If ValidaCodigos(.Row) Then
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    .Col = 1
                                End If
                                flexGastos.Focus()
                                Exit Sub
                            End If
                        End If
                End Select
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    If .Col = 1 Or .Col = 3 Then Exit Sub
                    .Row = .Row + 1
                    .TopRow = .Row
                Else
                    If .Col = 1 Or .Col = 3 Then Exit Sub
                    .Row = .Row + 1
                    If .Row > 6 Then
                        .TopRow = .Row
                    End If
                End If
                txtFlex.Visible = False
            ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
                txtFlex.Visible = False
                .Focus()
            ElseIf KeyCode = System.Windows.Forms.Keys.F3 Then
                Buscar()
            End If
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case Else
                Select Case flexGastos.Col
                    Case 0
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                    Case 1
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 2
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                    Case 3
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 4
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                End Select
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class
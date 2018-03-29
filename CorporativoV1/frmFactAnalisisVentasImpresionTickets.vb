Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmFactAnalisisVentasImpresionTickets
    Inherits System.Windows.Forms.Form

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdImprimir As System.Windows.Forms.Button
    Public WithEvents flexTickets As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents chkTodosLosTickets As System.Windows.Forms.CheckBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents lblSeleccionada As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label

    Public I As Integer
    Public FueraClick As Boolean

    Public Sub Encabezado()
        With flexTickets
            .set_ColWidth(0, 2000)
            .set_ColWidth(1, 2000)
            .set_ColWidth(2, 0)
            .set_ColWidth(3, 0)
            .set_ColWidth(4, 0)
            .set_ColWidth(5, 0)
            .set_ColWidth(6, 0)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Folio de Venta"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe"
            .Col = 0
            .Row = 1
        End With
    End Sub

    Public Function EstanTodosSeleccionados() As Boolean
        With flexTickets
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Then
                    .Row = I
                    If .CellBackColor.Equals(.BackColor) Then
                        EstanTodosSeleccionados = False
                        Exit Function
                    End If
                End If
            Next
        End With
        EstanTodosSeleccionados = True
    End Function

    Public Sub PonerColor()
        Dim Ren As Integer
        If Trim(flexTickets.get_TextMatrix(flexTickets.Row, 0)) = "" Then Exit Sub
        flexTickets.Col = 0
        If flexTickets.CellBackColor.Equals(flexTickets.BackColor) Then
            flexTickets.CellBackColor = lblSeleccionada.BackColor
            flexTickets.Col = 1
            flexTickets.CellBackColor = lblSeleccionada.BackColor
        ElseIf System.Drawing.ColorTranslator.ToOle(flexTickets.CellBackColor) = System.Drawing.ColorTranslator.ToOle(Me.lblSeleccionada.BackColor) Then
            flexTickets.CellBackColor = flexTickets.BackColor
            flexTickets.Col = 1
            flexTickets.CellBackColor = flexTickets.BackColor
        End If
        Ren = flexTickets.Row
        If EstanTodosSeleccionados() Then
            chkTodosLosTickets.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            FueraClick = True
            chkTodosLosTickets.CheckState = System.Windows.Forms.CheckState.Unchecked
            FueraClick = False
        End If
        flexTickets.Row = Ren
        flexTickets.Col = 0
    End Sub

    Public Sub SeleccionarTodosLosTickets()
        With flexTickets
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Then
                    If chkTodosLosTickets.CheckState = 1 Then
                        .Row = I
                        .Col = 0
                        .CellBackColor = lblSeleccionada.BackColor
                        .Col = 1
                        .CellBackColor = lblSeleccionada.BackColor
                    ElseIf chkTodosLosTickets.CheckState = 0 Then
                        .Row = I
                        .Col = 0
                        .CellBackColor = .BackColor
                        .Col = 1
                        .CellBackColor = .BackColor
                    End If
                End If
            Next
            .Col = 0
            .Row = 1
        End With
    End Sub

    Public Function Selecciono() As Boolean
        Selecciono = False
        With flexTickets
            For I = 1 To .Rows - 1
                .Col = 0
                .Row = I
                If System.Drawing.ColorTranslator.ToOle(.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblSeleccionada.BackColor) Then
                    Selecciono = True
                    Exit Function
                End If
            Next
        End With
        Return Selecciono
    End Function

    Public Sub chkTodosLosTickets_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosLosTickets.CheckStateChanged
        If FueraClick Then Exit Sub
        SeleccionarTodosLosTickets()
    End Sub

    Public Sub cmdImprimir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImprimir.Click
        If Not Selecciono() Then
            MsgBox("No selecciono ningun ticket, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Else
            With flexTickets
                For I = 1 To .Rows - 1
                    If Trim(.get_TextMatrix(I, 0)) <> "" Then
                        .Col = 0
                        .Row = I
                        If System.Drawing.ColorTranslator.ToOle(.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblSeleccionada.BackColor) Then
                            If .get_TextMatrix(I, 5) = "D" Then
                                ModCorporativo.TicketVentaReducidoDolares(Trim(.get_TextMatrix(I, 0)), CShort(.get_TextMatrix(I, 2)), IIf(Trim(.get_TextMatrix(I, 6)) = "V", Trim(.get_TextMatrix(I, 4)), "CO"), CShort(.get_TextMatrix(I, 3)), Trim(.get_TextMatrix(I, 6)))
                            ElseIf .get_TextMatrix(I, 5) = "P" Then
                                ModCorporativo.TicketVentaReducidoPesos(Trim(.get_TextMatrix(I, 0)), CShort(.get_TextMatrix(I, 2)), IIf(Trim(.get_TextMatrix(I, 6)) = "V", Trim(.get_TextMatrix(I, 4)), "CO"), CShort(.get_TextMatrix(I, 3)), Trim(.get_TextMatrix(I, 6)))
                            End If
                        End If
                    End If
                Next
            End With
        End If
    End Sub

    Public Sub flexTickets_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexTickets.DblClick
        PonerColor()
    End Sub

    Public Sub flexTickets_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexTickets.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
            flexTickets_DblClick(flexTickets, New System.EventArgs())
        End If
    End Sub

    Public Sub frmFactAnalisisVentasImpresionTickets_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            ModEstandar.AvanzarTab(Me)
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Public Sub frmFactAnalisisVentasImpresionTickets_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        Encabezado()
        Me.Activate()
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFactAnalisisVentasImpresionTickets))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdImprimir = New System.Windows.Forms.Button()
        Me.flexTickets = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.chkTodosLosTickets = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblSeleccionada = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.flexTickets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdImprimir
        '
        Me.cmdImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.cmdImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdImprimir.Location = New System.Drawing.Point(199, 236)
        Me.cmdImprimir.Name = "cmdImprimir"
        Me.cmdImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdImprimir.Size = New System.Drawing.Size(114, 37)
        Me.cmdImprimir.TabIndex = 5
        Me.cmdImprimir.Text = "&Imprimir"
        Me.cmdImprimir.UseVisualStyleBackColor = False
        '
        'flexTickets
        '
        Me.flexTickets.DataSource = Nothing
        Me.flexTickets.Location = New System.Drawing.Point(16, 48)
        Me.flexTickets.Name = "flexTickets"
        Me.flexTickets.OcxState = CType(resources.GetObject("flexTickets.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexTickets.Size = New System.Drawing.Size(297, 130)
        Me.flexTickets.TabIndex = 1
        '
        'chkTodosLosTickets
        '
        Me.chkTodosLosTickets.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodosLosTickets.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodosLosTickets.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodosLosTickets.Location = New System.Drawing.Point(16, 16)
        Me.chkTodosLosTickets.Name = "chkTodosLosTickets"
        Me.chkTodosLosTickets.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodosLosTickets.Size = New System.Drawing.Size(125, 21)
        Me.chkTodosLosTickets.TabIndex = 0
        Me.chkTodosLosTickets.Text = "Todos Los Tickets"
        Me.chkTodosLosTickets.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label2.Location = New System.Drawing.Point(48, 226)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(145, 21)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Tickets Seleccionados"
        '
        'lblSeleccionada
        '
        Me.lblSeleccionada.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSeleccionada.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSeleccionada.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSeleccionada.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSeleccionada.Location = New System.Drawing.Point(16, 224)
        Me.lblSeleccionada.Name = "lblSeleccionada"
        Me.lblSeleccionada.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSeleccionada.Size = New System.Drawing.Size(21, 21)
        Me.lblSeleccionada.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label1.Location = New System.Drawing.Point(16, 184)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(297, 33)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Haga Doble Click o Presione la Barra Espaciadora Para Seleccionar un Ticket"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmFactAnalisisVentasImpresionTickets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(332, 286)
        Me.Controls.Add(Me.cmdImprimir)
        Me.Controls.Add(Me.flexTickets)
        Me.Controls.Add(Me.chkTodosLosTickets)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblSeleccionada)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFactAnalisisVentasImpresionTickets"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Impresión de Tickets"
        CType(Me.flexTickets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub



End Class
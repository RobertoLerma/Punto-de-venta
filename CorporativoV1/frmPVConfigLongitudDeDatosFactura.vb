Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmPVConfigLongitudDeDatosFactura
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnAceptar As System.Windows.Forms.Button
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents FlexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Marco As System.Windows.Forms.GroupBox

    Sub Encabezado()
        With flexDetalle
            .Row = 0
            .Col = 0
            .set_ColWidth(0, 0, 1800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "DATOS"
            .Col = 1
            .set_ColWidth(1, 0, 1400)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "LONGITUD"
            .Col = 0
            .Row = 1
            .CellFontBold = True
            .Text = "CLIENTE"
            .Row = 2
            .Col = 0
            .CellFontBold = True
            .Text = "DIRECCION"
            .Row = 3
            .Col = 0
            .CellFontBold = True
            .Text = "COLONIA"
            .Row = 4
            .Col = 0
            .CellFontBold = True
            .Text = "CIUDAD"
            .Row = 5
            .Col = 0
            .CellFontBold = True
            .Text = "ESTADO"
            .Row = 6
            .Col = 0
            .CellFontBold = True
            .Text = "LEYENDA"
            .Row = 7
            .Col = 0
            .CellFontBold = True
            .Text = "DESC. PRODUCTO"
            .Row = 1
            .Col = 1
        End With
    End Sub

    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        Me.Hide()
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDetalle.DblClick
        flexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub flexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
    End Sub

    Private Sub flexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalle.KeyPressEvent
        With FlexDetalle
            If .Col = 1 Then
                ModEstandar.gp_CampoNumerico(eventArgs.keyAscii)
            End If
            If eventArgs.keyAscii = 13 And .Col = 1 Then
                If (.Row > 1) Then
                    If Trim(.get_TextMatrix(.Row - 1, 0)) = "" Then
                        .Focus()
                        Exit Sub
                    End If
                End If
                MSHFlexGridEdit(FlexDetalle, txtFlex, eventArgs.keyAscii)
                If Len(Trim(txtFlex.Text)) <> 1 Then
                    SelTxt()
                End If
            End If
        End With
    End Sub

    Private Sub flexDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Leave
        flexDetalle.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub frmPVConfigLongitudDeDatosFactura_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Encabezado()
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        With FlexDetalle
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    'txtFlex.Visible = False
                    'txtFlex.Text = ""
                    'FlexDetalle.Focus()
                Case System.Windows.Forms.Keys.Return
                    If Val(txtFlex.Text) <= 255 Then
                        .set_TextMatrix(.Row, .Col, txtFlex.Text)
                        txtFlex.Visible = False
                        txtFlex.Text = ""
                        FlexDetalle.Focus()
                        If .Row < .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            System.Windows.Forms.SendKeys.Send("{TAB}")
                        End If
                    Else
                        MsgBox("La Longitud Maxima es de 255.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        txtFlex.Text = ""
                    End If
            End Select
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPVConfigLongitudDeDatosFactura))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.Marco = New System.Windows.Forms.GroupBox()
        Me.FlexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Marco.SuspendLayout()
        CType(Me.FlexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.Color.White
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtFlex.Location = New System.Drawing.Point(14, 42)
        Me.txtFlex.MaxLength = 3
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(63, 20)
        Me.txtFlex.TabIndex = 2
        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtFlex, "Longitud")
        Me.txtFlex.Visible = False
        '
        'btnAceptar
        '
        Me.btnAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.btnAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAceptar.Location = New System.Drawing.Point(152, 200)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAceptar.Size = New System.Drawing.Size(97, 35)
        Me.btnAceptar.TabIndex = 3
        Me.btnAceptar.Text = "Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = False
        '
        'Marco
        '
        Me.Marco.BackColor = System.Drawing.SystemColors.Control
        Me.Marco.Controls.Add(Me.txtFlex)
        Me.Marco.Controls.Add(Me.FlexDetalle)
        Me.Marco.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Marco.Location = New System.Drawing.Point(9, 1)
        Me.Marco.Name = "Marco"
        Me.Marco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Marco.Size = New System.Drawing.Size(244, 187)
        Me.Marco.TabIndex = 0
        Me.Marco.TabStop = False
        '
        'FlexDetalle
        '
        Me.FlexDetalle.DataSource = Nothing
        Me.FlexDetalle.Location = New System.Drawing.Point(12, 20)
        Me.FlexDetalle.Name = "FlexDetalle"
        Me.FlexDetalle.OcxState = CType(resources.GetObject("FlexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FlexDetalle.Size = New System.Drawing.Size(217, 156)
        Me.FlexDetalle.TabIndex = 1
        '
        'frmPVConfigLongitudDeDatosFactura
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(262, 246)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.Marco)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(258, 165)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPVConfigLongitudDeDatosFactura"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Longitudes de Datos"
        Me.Marco.ResumeLayout(False)
        CType(Me.FlexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

End Class
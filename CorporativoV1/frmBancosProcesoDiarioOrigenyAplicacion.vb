Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioOrigenyAplicacion
    Inherits System.Windows.Forms.Form

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             ORIGEN Y APLICACION                                                                          *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MARTES 15 DE JULIO DE 2003                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblTotal As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents lblFolio As System.Windows.Forms.Label
    Public WithEvents lblFechaMovimiento As System.Windows.Forms.Label
    Public WithEvents lblImporte As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents Label3 As System.Windows.Forms.Label

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoDiarioOrigenyAplicacion))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblFolio = New System.Windows.Forms.Label()
        Me.lblFechaMovimiento = New System.Windows.Forms.Label()
        Me.lblImporte = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdAceptar
        '
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Location = New System.Drawing.Point(495, 272)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.Size = New System.Drawing.Size(79, 25)
        Me.cmdAceptar.TabIndex = 2
        Me.cmdAceptar.Text = "&Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtFlex)
        Me.Frame1.Controls.Add(Me.flexDetalle)
        Me.Frame1.Controls.Add(Me.lblTotal)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(10, 34)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(586, 225)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(16, 47)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(64, 20)
        Me.txtFlex.TabIndex = 7
        Me.txtFlex.Visible = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(13, 24)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(561, 151)
        Me.flexDetalle.TabIndex = 1
        '
        'lblTotal
        '
        Me.lblTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblTotal.Location = New System.Drawing.Point(455, 181)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotal.Size = New System.Drawing.Size(100, 21)
        Me.lblTotal.TabIndex = 9
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(414, 183)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(49, 21)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Total :"
        '
        'lblFolio
        '
        Me.lblFolio.BackColor = System.Drawing.SystemColors.Control
        Me.lblFolio.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFolio.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFolio.Location = New System.Drawing.Point(584, 273)
        Me.lblFolio.Name = "lblFolio"
        Me.lblFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFolio.Size = New System.Drawing.Size(22, 22)
        Me.lblFolio.TabIndex = 12
        Me.lblFolio.Visible = False
        '
        'lblFechaMovimiento
        '
        Me.lblFechaMovimiento.BackColor = System.Drawing.SystemColors.Window
        Me.lblFechaMovimiento.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFechaMovimiento.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFechaMovimiento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFechaMovimiento.Location = New System.Drawing.Point(65, 8)
        Me.lblFechaMovimiento.Name = "lblFechaMovimiento"
        Me.lblFechaMovimiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFechaMovimiento.Size = New System.Drawing.Size(93, 21)
        Me.lblFechaMovimiento.TabIndex = 11
        Me.lblFechaMovimiento.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblImporte
        '
        Me.lblImporte.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporte.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporte.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblImporte.Location = New System.Drawing.Point(461, 9)
        Me.lblImporte.Name = "lblImporte"
        Me.lblImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporte.Size = New System.Drawing.Size(100, 21)
        Me.lblImporte.TabIndex = 10
        Me.lblImporte.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(16, 270)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(350, 30)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Insert = Insertar Renglon                                                        " &
    "                 Supr  = Eliminar Renglon"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(42, 14)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Fecha :"
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(216, 8)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(161, 18)
        Me.lblMoneda.TabIndex = 0
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(409, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 15)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Importe :"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(134, 319)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(109, 36)
        Me.btnLimpiar.TabIndex = 42
        Me.btnLimpiar.Text = "&Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(19, 319)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 41
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoDiarioOrigenyAplicacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(605, 367)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.cmdAceptar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.lblFolio)
        Me.Controls.Add(Me.lblFechaMovimiento)
        Me.Controls.Add(Me.lblImporte)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblMoneda)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(174, 204)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioOrigenyAplicacion"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Origen y Aplicación de Recursos"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Public blnBuscar As Boolean
    Public Nuevo As Boolean
    Dim Carga As Boolean

    Function GuardarMovimientosOrigenAplicacion(ByRef TipoMovimiento As String) As Boolean
        Dim I As Integer
        On Error GoTo Err_Renamed
        GuardarMovimientosOrigenAplicacion = True
        With flexDetalle
            Select Case TipoMovimiento
                Case "REGISTRO DE PAGOS"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                            ModStoredProcedures.PR_IMEMovimientosOrigenAplic((frmBancosProcesoDiarioRegistrodePagos.txtFolioEgreso).Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", "S", .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next

                Case "REGISTRO DE DEPOSITOS"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                            'ModStoredProcedures.PR_IMEMovimientosOrigenAplic((frmBancosProcesoDiarioRegistrodeDepositos.txtFolioIngreso).Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", "E", .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
                Case "REGISTRO DE DEPOSITOS PES"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                            'ModStoredProcedures.PR_IMEMovimientosOrigenAplic((frmBancosProcesoDiarioRegistrodeDepositos.strFolioPesos), .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", "E", .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
                Case "REGISTRO DE DEPOSITOS DOL"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                            'ModStoredProcedures.PR_IMEMovimientosOrigenAplic((frmBancosProcesoDiarioRegistrodeDepositos.strFolioDolares), .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", "E", .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next

                Case "REGISTRO DE CARGOS"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                            'si jala ModStoredProcedures.PR_IMEMovimientosOrigenAplic((frmBancosProcesoDiarioCargosDiversos.txtFolioEgreso).Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", "S", .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
                Case "REGISTRO DE ANTICIPOS"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                            'si jala ModStoredProcedures.PR_IMEMovimientosOrigenAplic((frmBancosProcesoDiarioAnticipoProveedoresAcreed.txtFolioEgreso).Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", "S", .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
                Case "REGISTRO DE OTROS INGRESOS"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                            'si jala ModStoredProcedures.PR_IMEMovimientosOrigenAplic((frmBancosProcesoDiarioRegistrodeOtrosIngresos.txtFolioIngreso).Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", "E", .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
            End Select
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            GuardarMovimientosOrigenAplicacion = False
        End If
    End Function

    Function Cambios() As Boolean
        On Error GoTo Err_Renamed
        Dim I As Integer
        Cambios = False
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> VB.Left(Trim(.get_TextMatrix(I, 7)), 4) Or Trim(.get_TextMatrix(I, 2)) <> VB.Right(Trim(.get_TextMatrix(I, 7)), 6) Or Trim(.get_TextMatrix(I, 4)) <> Trim(.get_TextMatrix(I, 6)) Then
                    Cambios = True
                    Exit Function
                End If
            Next
        End With
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Guardar()
        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        Dim I As Integer
        If Cambios() Then
            If MsgBox("¿Desea Guardar los Cambios?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) = MsgBoxResult.No Then
                Me.Close()
            End If
        Else
            Me.Close()
        End If
        If CDbl(Numerico(lblTotal.Text)) <> CDbl(Numerico(lblImporte.Text)) Then
            MsgBox("El Total debe ser igual al importe, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        With flexDetalle
            'Borramos todos los Movimientos de Origen y Aplicación que correspondan al folio que estamos modificando
            ModStoredProcedures.PR_IMEMovimientosOrigenAplic(Trim(lblFolio.Text), "0", "0", "0", "0", "", "0", "", "01/01/1900", C_ELIMINACION, CStr(0))
            Cmd.Execute()
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                    ModStoredProcedures.PR_IMEMovimientosOrigenAplic(Trim(lblFolio.Text), .get_TextMatrix(I, 0), .get_TextMatrix(I, 2), "0", "0", gstrMovimiento, .get_TextMatrix(I, 4), "V", "01/01/1900", C_INSERCION, CStr(0))
                    Cmd.Execute()
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        'si jala frmBancosProcesoMensualConsultaOrigenAplicRec.VerMovimientos()
        Me.Close()
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub EliminarLinea()
        Dim Ren As Integer
        Ren = flexDetalle.Rows
        flexDetalle.RemoveItem(flexDetalle.Row)
        flexDetalle.Rows = Ren
        flexDetalle.set_TextMatrix(flexDetalle.Rows - 1, 4, "0.00")
        CalculoImporte()
    End Sub

    Sub InsertarLinea()
        flexDetalle.AddItem("", flexDetalle.Row)
        flexDetalle.set_TextMatrix(flexDetalle.Row, 4, "0.00")
    End Sub

    Function ChecarPartidas() As Boolean
        Dim I As Integer
        Dim J As Integer
        ChecarPartidas = True
        With flexDetalle
            For I = 1 To .Rows - 1
                If I = 1 Then
                    If Trim(.get_TextMatrix(I, 0)) = "" And Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "" And Trim(.get_TextMatrix(I, 3)) = "" And CDbl(Numerico(.get_TextMatrix(I, 4))) = 0 Then
                        MsgBox("No ha Capturado Ninguna Partida, Favor de Capturar al Menos una Partida", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                        ChecarPartidas = False
                        .Row = 1
                        .Col = 0
                        .Focus()
                        Exit Function
                    ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And CDbl(Numerico(.get_TextMatrix(I, 4))) > 0 Then
                        ChecarPartidas = True
                    Else
                        MsgBox("No ha Capturado Toda la Información de la Ultima Patida, Favor de Verificar..", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        ChecarPartidas = False
                        Exit Function
                    End If
                Else
                    If Trim(.get_TextMatrix(I, 0)) = "" And Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "" And Trim(.get_TextMatrix(I, 3)) = "" And CDbl(Numerico(.get_TextMatrix(I, 4))) = 0 Then
                        ChecarPartidas = True
                    ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And CDbl(Numerico(.get_TextMatrix(I, 4))) > 0 Then
                        ChecarPartidas = True
                        If I = .Rows - 1 Then
                            Exit Function
                        End If
                    Else
                        MsgBox("No ha Capturado Toda la Información de la Ultima Patida, Favor de Verificar..", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        ChecarPartidas = False
                        Exit Function
                    End If
                End If
            Next
        End With
    End Function

    Sub CalculoImporte()
        Dim I As Integer
        lblTotal.Text = ""
        With flexDetalle
            For I = 1 To .Rows - 1
                If CDbl(Numerico(.get_TextMatrix(I, 4))) <> 0 Then
                    lblTotal.Text = CStr(CDbl(VB6.Format(Numerico(lblTotal.Text), "#####0.00")) + CDbl(VB6.Format(Numerico(.get_TextMatrix(I, 4)), "#####0.00")))
                End If
            Next
        End With
        If Trim(lblTotal.Text) = "" Then
            lblTotal.Text = "0.00"
        Else
            lblTotal.Text = VB6.Format(lblTotal.Text, "###,##0.00")
        End If
        If CDbl(Numerico(lblTotal.Text)) <> CDbl(Numerico(lblImporte.Text)) Then
            lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0)
        ElseIf CDbl(Numerico(lblTotal.Text)) = CDbl(Numerico(lblImporte.Text)) Then
            lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
        End If
    End Sub

    Sub GuardaLlave()
        flexDetalle.set_TextMatrix(flexDetalle.Row, 5, Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) & Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2)))
    End Sub

    Function BuscarLlave(ByRef llave As String, ByRef LlaveNotBusca As Integer) As Boolean
        Dim I As Integer
        BuscarLlave = False
        With flexDetalle
            For I = 1 To .Rows - 1
                If I <> LlaveNotBusca Then
                    If Trim(.get_TextMatrix(I, 5)) = Trim(llave) Then
                        BuscarLlave = True
                        Exit Function
                    End If
                End If
            Next
        End With
    End Function

    Sub ValidaLlave()
        With flexDetalle
            GuardaLlave()
            If Len(Trim(.get_TextMatrix(.Row, 0))) = 4 And Len(Trim(.get_TextMatrix(.Row, 2))) = 6 Then
                If BuscarLlave(Trim(.get_TextMatrix(.Row, 0)) & Trim(.get_TextMatrix(.Row, 2)), .Row) Then
                    MsgBox("El Agrupador y el Rubro ya Existen, No se Pueden Repetir..", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    .set_TextMatrix(.Row, 2, "")
                    .set_TextMatrix(.Row, 3, "")
                    .set_TextMatrix(.Row, 5, "")
                    .Col = 2
                    .Focus()
                    Exit Sub
                    'Else
                    '    GuardaLlave
                End If
            End If
        End With
    End Sub

    Function LlenaDatosAgrupador() As Boolean
        On Error GoTo Err_Renamed
        LlenaDatosAgrupador = False
        If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2)) = "" And Len(Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2))) < 6 Then
            gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR = " & Numerico(txtFlex.Text)
        ElseIf Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2)) <> "" And Len(Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2))) = 6 Then
            'gStrSql = "SELECT * " & _
            ''"FROM CatOrigenAplicRecursos A, CatRubrosOrigenAplicRecursos R WHERE R.CodRubro = " & Numerico(flexDetalle.TextMatrix(flexDetalle.Row, 2)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR AND A.Aplicacion = '" & gstrMovimiento & "'"
            gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR = " & Numerico(txtFlex.Text)
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Aplicacion").Value <> Trim(gstrMovimiento) Then
                MsgBox("El Tipo de Aplicación de este Agrupador no Coincide con el Tipo de Movimiento..", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtFlex.Text = ""
            ElseIf RsGral.Fields("Aplicacion").Value = Trim(gstrMovimiento) Then
                txtFlex.Text = VB.Right("0000" & CStr(RsGral.Fields("CodOrigenAplicR").Value), 4)
                flexDetalle.set_TextMatrix(flexDetalle.Row, 1, Trim(RsGral.Fields("DescOrigenAplicR").Value))
                If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) <> Trim(txtFlex.Text) Then
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 2, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 3, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 4, "0.00")
                    CalculoImporte()
                End If
                LlenaDatosAgrupador = True
                txtFlex_Leave(txtFlex, New System.EventArgs())
            End If
        Else
            MsgBox("Codigo Inexistente Favor de Verificar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexDetalle.Col = 0
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DescripcionAgrupador() As Boolean
        On Error GoTo Err_Renamed
        DescripcionAgrupador = False
        'If Trim(flexDetalle.TextMatrix(flexDetalle.Row, 3)) = "" Then
        gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR = '" & Trim(txtFlex.Text) & "'"
        'End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Aplicacion").Value <> Trim(gstrMovimiento) Then
                MsgBox("El Tipo de Aplicación de este Agrupador no Coincide con el Tipo de Movimiento..", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtFlex.Text = ""
            ElseIf RsGral.Fields("Aplicacion").Value = Trim(gstrMovimiento) Then
                txtFlex.Text = Trim(RsGral.Fields("DescOrigenAplicR").Value)
                flexDetalle.set_TextMatrix(flexDetalle.Row, 0, VB.Right("0000" & CStr(RsGral.Fields("CodOrigenAplicR").Value), 4))
                If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 1)) <> Trim(txtFlex.Text) Then
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 2, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 3, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 4, "0.00")
                    CalculoImporte()
                End If
                DescripcionAgrupador = True
                txtFlex_Leave(txtFlex, New System.EventArgs())
                frmBancosProcesoDiarioAnticipoProveedoresAcreed.ConsultaAnticipos = True
            End If
        Else
            MsgBox("Descripción Inexistente Favor de Verificar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexDetalle.Col = 1
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function LlenaDatosRubro() As Boolean
        On Error GoTo Err_Renamed
        LlenaDatosRubro = False
        If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" And Len(Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))) < 4 Then
            MsgBox("No ha Capturado ningun Agupador para esta Partida" & Chr(13) & "        Favor de Capturar un Agrupador", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Function
            'gStrSql = "SELECT Aplicacion,CodRubro,DescRubro FROM CatOrigenAplicRecursos,CatRubrosOrigenAplicRecursos WHERE CodRubro = " & Numerico(txtFlex) & " AND CodOrigenAplicR = CodOrigAplicR GROUP BY Aplicacion,CodRubro,DescRubro"
        ElseIf Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) <> "" And Len(Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))) = 4 Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE CodOrigenAplicR = " & Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) & " AND CodOrigenAplicR = CodOrigAplicR AND CodRubro = " & Numerico(txtFlex.Text) & " GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Aplicacion").Value <> Trim(gstrMovimiento) Then
                MsgBox("Este Rubro Depende de Un Agrupador cuyo Tipo de Aplicación no Coincide con el Tipo de Movimiento..", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtFlex.Text = ""
                Exit Function
                '        ElseIf (Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0)) <> "" And Len(Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0))) = 4) Then
                '            If RsGral!CodOrigAplicR <> Numerico(flexDetalle.TextMatrix(flexDetalle.Row, 0)) Or (Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0)) <> "" And Len(Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0))) = 4) Then
                '                flexDetalle.TextMatrix(flexDetalle.Row, 0) = Format(RsGral!CodOrigenAplicR, "0000")
                '                flexDetalle.TextMatrix(flexDetalle.Row, 1) = Trim(RsGral!DescOrigenAplicR)
                '                txtFlex = Right("000000" & CStr(RsGral!CodRubro), 6)
                '                flexDetalle.TextMatrix(flexDetalle.Row, 3) = Trim(RsGral!DescRubro)
                '                txtFlex_LostFocus
                '             End If
            ElseIf RsGral.Fields("Aplicacion").Value = Trim(gstrMovimiento) Then
                txtFlex.Text = VB.Right("000000" & CStr(RsGral.Fields("CodRubro").Value), 6)
                flexDetalle.set_TextMatrix(flexDetalle.Row, 3, Trim(RsGral.Fields("DescRubro").Value))
                txtFlex_Leave(txtFlex, New System.EventArgs())
                LlenaDatosRubro = True
                frmBancosProcesoDiarioAnticipoProveedoresAcreed.ConsultaAnticipos = True
            End If
        Else
            MsgBox("Codigo Inexistente Favor de Investigar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DescripcionRubro() As Boolean
        On Error GoTo Err_Renamed
        DescripcionRubro = False
        If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" And Len(Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))) < 4 Then
            MsgBox("No ha Capturado ningun Agupador para esta Partida" & Chr(13) & "        Favor de Capturar un Agrupador", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Function
            'gStrSql = "SELECT Aplicacion,CodRubro,DescRubro FROM CatOrigenAplicRecursos,CatRubrosOrigenAplicRecursos WHERE CodRubro = " & Numerico(txtFlex) & " AND CodOrigenAplicR = CodOrigAplicR GROUP BY Aplicacion,CodRubro,DescRubro"
        ElseIf Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) <> "" And Len(Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))) = 4 Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE CodOrigenAplicR = " & Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) & " AND CodOrigenAplicR = CodOrigAplicR AND DescRubro = '" & Trim(txtFlex.Text) & "' GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Aplicacion").Value <> Trim(gstrMovimiento) Then
                MsgBox("Este Rubro Depende de Un Agrupador cuyo Tipo de Aplicación no Coincide con el Tipo de Movimiento..", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtFlex.Text = ""
                Exit Function
                '        ElseIf (Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0)) <> "" And Len(Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0))) = 4) Then
                '            If RsGral!CodOrigAplicR <> Numerico(flexDetalle.TextMatrix(flexDetalle.Row, 0)) Or (Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0)) <> "" And Len(Trim(flexDetalle.TextMatrix(flexDetalle.Row, 0))) = 4) Then
                '                flexDetalle.TextMatrix(flexDetalle.Row, 0) = Format(RsGral!CodOrigenAplicR, "0000")
                '                flexDetalle.TextMatrix(flexDetalle.Row, 1) = Trim(RsGral!DescOrigenAplicR)
                '                txtFlex = Right("000000" & CStr(RsGral!CodRubro), 6)
                '                flexDetalle.TextMatrix(flexDetalle.Row, 3) = Trim(RsGral!DescRubro)
                '                txtFlex_LostFocus
                '             End If
            ElseIf RsGral.Fields("Aplicacion").Value = Trim(gstrMovimiento) Then
                txtFlex.Text = Trim(RsGral.Fields("DescRubro").Value)
                flexDetalle.set_TextMatrix(flexDetalle.Row, 2, VB.Right("000000" & CStr(RsGral.Fields("CodRubro").Value), 6))
                txtFlex_Leave(txtFlex, New System.EventArgs())
                DescripcionRubro = True
            End If
        Else
            MsgBox("Codigo Inexistente Favor de Investigar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim strDesc As String
        Dim I As Object
        Dim J As Integer
        If Nuevo Then
            Exit Sub
        End If
        If flexDetalle.Row > 1 Then
            With flexDetalle
                For I = .Row - 1 To 1 Step -1
                    'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 2)) = "" Then Exit Sub
                Next
            End With
        End If
        If flexDetalle.Col = 2 Or flexDetalle.Col = 3 Then
            If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" And Len(Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))) < 4 Then
                MsgBox("No ha Capturado ningun Agupador para esta Partida" & Chr(13) & "        Favor de Capturar un Agrupador", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" And Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 1)) = "" Then
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 0, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 1, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 2, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 3, "")
                    flexDetalle.set_TextMatrix(flexDetalle.Row, 4, "0.00")
                    flexDetalle.Col = 0
                    CalculoImporte()
                    txtFlex.Visible = False
                    flexDetalle.Focus()
                End If
                Exit Sub
            End If
        End If
        With flexDetalle
            If .Col = 0 Then
                strControlActual = "CODIGO AGRUPADOR"
                strTag = UCase(Me.Tag) & "." & strControlActual
            ElseIf .Col = 2 Then
                strControlActual = "CODIGO RUBRO"
                strTag = UCase(Me.Tag) & "." & strControlActual
            ElseIf .Col = 1 Then
                strControlActual = "DESCRIPCION AGRUPADOR"
                strTag = UCase(Me.Tag) & "." & strControlActual
            ElseIf .Col = 3 Then
                strControlActual = "DESCRIPCION RUBRO"
                strTag = UCase(Me.Tag) & "." & strControlActual
            ElseIf .Col = 4 Then
                Exit Sub
            End If

            If Me.ActiveControl.Name = "txtFlex" And (.Col = 1 Or .Col = 3) Then
                strDesc = Trim(txtFlex.Text)

            ElseIf Me.ActiveControl.Name = "flexDetalle" And (.Col = 1 Or .Col = 3) Then
                strDesc = .Text
            End If
            Select Case strControlActual
                Case "CODIGO AGRUPADOR"
                    strCaptionForm = "Consulta de Agrupadores de Origen y Aplicación"
                    If Trim(.get_TextMatrix(.Row, 2)) = "" And Len(Trim(.get_TextMatrix(.Row, 2))) < 6 Then
                        gStrSql = "SELECT RIGHT('0000' + LTRIM(CodOrigenAplicR),4) AS AGRUPADOR, DescOrigenAplicR AS DESCRIPCION " & "FROM CatOrigenAplicRecursos WHERE Aplicacion = '" & gstrMovimiento & "' ORDER BY CodOrigenAplicR"
                    ElseIf Trim(.get_TextMatrix(.Row, 2)) <> "" And Len(Trim(.get_TextMatrix(.Row, 2))) = 6 Then
                        gStrSql = "SELECT RIGHT('0000' + LTRIM(CodOrigenAplicR),4) AS AGRUPADOR, DescOrigenAplicR AS DESCRIPCION " & "FROM CatOrigenAplicRecursos WHERE Aplicacion = '" & gstrMovimiento & "' ORDER BY CodOrigenAplicR"
                        '                    gStrSql = "SELECT RIGHT('0000' + LTRIM(R.CodOrigAplicR),4) AS AGRUPADOR, A.DescOrigenAplicR AS DESCRIPCION " & _
                        ''                    "FROM CatOrigenAplicRecursos A, CatRubrosOrigenAplicRecursos R WHERE R.CodRubro = " & Numerico(.TextMatrix(.Row, 2)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR AND A.Aplicacion = '" & gstrMovimiento & "' GROUP BY R.CodOrigAplicR,A.DescOrigenAplicR ORDER BY R.CodOrigAplicR"
                    End If
                Case "CODIGO RUBRO"
                    strCaptionForm = "Consulta de Rubros de Origen y Aplicación"
                    If Trim(.get_TextMatrix(.Row, 0)) = "" And Len(Trim(.get_TextMatrix(.Row, 0))) < 4 Then
                        gStrSql = "SELECT RIGHT('000000' + LTRIM(CodRubro),6) AS RUBRO, DescRubro AS DESCRIPCION " & "FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos WHERE Aplicacion = '" & gstrMovimiento & "' AND CodOrigenAplicR = CodOrigAplicR ORDER BY CodRubro"
                    ElseIf Trim(.get_TextMatrix(.Row, 0)) <> "" And Len(Trim(.get_TextMatrix(.Row, 0))) = 4 Then
                        gStrSql = "SELECT RIGHT('000000' + LTRIM(R.CodRubro),6) AS RUBRO, R.DescRubro AS DESCRIPCION " & "FROM CatRubrosOrigenAplicRecursos R,CatOrigenAplicRecursos A WHERE A.CodOrigenAplicR = " & Numerico(.get_TextMatrix(.Row, 0)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR AND A.Aplicacion = '" & gstrMovimiento & "' ORDER BY R.CodRubro"
                    End If
                Case "DESCRIPCION AGRUPADOR"
                    strCaptionForm = "Consulta de Agrupadores de Origen y Aplicación"
                    If Trim(.get_TextMatrix(.Row, 2)) = "" And Len(Trim(.get_TextMatrix(.Row, 2))) < 6 Then
                        gStrSql = "SELECT DescOrigenAplicR AS DESCRIPCION, RIGHT('0000' + LTRIM(CodOrigenAplicR),4) AS AGRUPADOR " & "FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & strDesc & "%' AND Aplicacion = '" & gstrMovimiento & "' ORDER BY DescOrigenAplicR"
                    ElseIf Trim(.get_TextMatrix(.Row, 2)) <> "" And Len(Trim(.get_TextMatrix(.Row, 2))) = 6 Then
                        gStrSql = "SELECT A.DescOrigenAplicR AS DESCRIPCION, RIGHT('0000' + LTRIM(R.CodOrigAplicR),4) AS AGRUPADOR " & "FROM CatOrigenAplicRecursos A ,CatRubrosOrigenAplicRecursos R WHERE A.DescOrigenAplicR LIKE '" & strDesc & "%' AND R.CodRubro = " & Numerico(.get_TextMatrix(.Row, 2)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR AND A.Aplicacion = '" & gstrMovimiento & "' GROUP BY R.CodOrigAplicR,A.DescOrigenAplicR ORDER BY A.DescOrigenAplicR"
                    End If
                Case "DESCRIPCION RUBRO"
                    strCaptionForm = "Consulta de Rubros de Origen y Aplicación"
                    If Trim(.get_TextMatrix(.Row, 0)) = "" And Len(Trim(.get_TextMatrix(.Row, 0))) < 4 Then
                        gStrSql = "SELECT DescRubro AS DESCRIPCION, RIGHT('000000' + LTRIM(CodRubro),6) AS RUBRO " & "FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos WHERE DescRubro LIKE '" & strDesc & "%' AND Aplicacion = '" & gstrMovimiento & "' AND CodOrigenAplicR = CodOrigAplicR ORDER BY DescRubro"
                    ElseIf Trim(.get_TextMatrix(.Row, 0)) <> "" And Len(Trim(.get_TextMatrix(.Row, 0))) = 4 Then
                        gStrSql = "SELECT R.DescRubro AS DESCRIPCION, RIGHT('000000' + LTRIM(R.CodRubro),6) AS RUBRO " & "FROM CatRubrosOrigenAplicRecursos R,CatOrigenAplicRecursos A WHERE R.DescRubro LIKE '" & strDesc & "%' AND A.CodOrigenAplicR = " & Numerico(.get_TextMatrix(.Row, 0)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR AND A.Aplicacion = '" & gstrMovimiento & "' ORDER BY DescRubro"
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
            'si jala    Load(FrmConsultas)
            '    With FrmConsultas.Flexdet
            '        Select Case strControlActual
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



        If ActiveControl.Name = "flexDetalle" Then
            blnBuscar = True
        End If
        'si jala FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Encabezado()
        Dim I As Integer
        With flexDetalle
            .set_Cols(0, 8)
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 800)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Agrup."
            .Col = 1
            .set_ColWidth(1, 0, 2500)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Descripción"
            .Col = 2
            .set_ColWidth(2, 0, 800)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Rubro"
            .Col = 3
            .set_ColWidth(3, 0, 2500)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Descripción"
            .Col = 4
            .set_ColWidth(4, 0, 1500)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Importe"
            .Col = 5
            .set_ColWidth(5, 0, 0)
            .set_ColWidth(6, 0, 0)
            .set_ColWidth(7, 0, 0)
            .Col = 0
            .Row = 1
            For I = 1 To .Rows - 1
                .set_TextMatrix(I, 4, "0.00")
                .set_TextMatrix(I, 6, "0.00")
            Next
        End With
    End Sub

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            Select Case flexDetalle.Col
                Case 0 'Codigo del Agrupador
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 4
                Case 1 'Descripción del Agrupador
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
                Case 2 'Codigo del Rubro
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 6
                Case 3 'Descripción del Rubro
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
                Case 4 'Importe
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 18
            End Select
        End With
    End Sub

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        txtFlex.Visible = False
        Select Case Me.Tag
            Case "frmPagos2"
                If frmBancosProcesoDiarioRegistrodePagos.ConsultaPagos Then
                    frmPagos.Hide()
                    Exit Sub
                End If
            Case "frmDepositos"
                If frmBancosProcesoDiarioRegistrodeDepositos.ConsultaDepositos Then
                    frmDepositos.Hide()
                    Exit Sub
                End If
            Case "frmDepositosIntPes"
                If frmBancosProcesoDiarioRegistrodeDepositos.ConsultaDepositos Then
                    frmDepositosIntPes.Hide()
                    Exit Sub
                End If
            Case "frmDepositosIntDol"
                If frmBancosProcesoDiarioRegistrodeDepositos.ConsultaDepositos Then
                    frmDepositosIntDol.Hide()
                    Exit Sub
                End If
            Case "frmCargos"
                If frmBancosProcesoDiarioCargosDiversos.ConsultaCargos Then
                    frmCargos.Hide()
                    Exit Sub
                End If
            Case "frmAnticipos2"
                If frmBancosProcesoDiarioAnticipoProveedoresAcreed.ConsultaAnticipos Then
                    frmBancosProcesoDiarioAnticipoProveedoresAcreed.frmAnticipos2.Hide()
                    Exit Sub
                End If
            Case "frmOtrosIngresos"
                If frmBancosProcesoDiarioRegistrodeOtrosIngresos.ConsultaOtrosIngresos Then
                    frmOtrosIngresos.Hide()
                    Exit Sub
                End If
        End Select
        If ChecarPartidas() Then
            Select Case Me.Tag
                Case "frmPagos"
                    frmPagos.Hide()
                Case "frmDepositos"
                    frmDepositos.Hide()
                Case "frmDepositosIntPes"
                    frmDepositosIntPes.Hide()
                Case "frmDepositosIntDol"
                    frmDepositosIntDol.Hide()
                Case "frmCargos"
                    frmCargos.Hide()
                Case "frmAnticipos"
                    frmAnticipos.Hide()
                Case "frmOtrosIngresos"
                    frmOtrosIngresos.Hide()
                Case "frmConsultaOrigenAplicacion"
                    'Guardar()
                    frmConsultaOrigenAplicacion.Hide()
            End Select
        End If
    End Sub

    Private Sub cmdAceptar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Enter
        txtFlex.Visible = False
    End Sub

    Private Sub cmdAceptar_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdAceptar.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            flexDetalle.Focus()
        End If
    End Sub

    Private Sub flexDetalle_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.ClickEvent
        txtFlex.Visible = False
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        FlexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        txtFlex.Visible = False
        Pon_Tool()
    End Sub

    Private Sub FlexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete And Not Nuevo Then
            EliminarLinea()
        ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.Insert And Not Nuevo Then
            InsertarLinea()
        ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.F3 Then
            Buscar()
            '    ElseIf KeyCode = vbKeyEscape Then
            '        cmdAceptar_Click
        End If
    End Sub

    Private Sub FlexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalle.KeyPressEvent
        Dim lonR, lonI As Integer
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And Not blnBuscar And Not Nuevo Then
            'Verifica si se puede capturar la fila
            If flexDetalle.Row > 1 Then
                If flexDetalle.get_TextMatrix(flexDetalle.Row - 1, 0) <> "" Then
                    For lonR = 1 To flexDetalle.Row - 1 Step 1
                        For lonI = 0 To 4 Step 1
                            If flexDetalle.get_TextMatrix(lonR, lonI) = "" Then
                                'MsgBox "Hace falta información en la captura", vbExclamation, cNomEmp
                                flexDetalle.Row = lonR
                                flexDetalle.Col = lonI
                                If flexDetalle.Col = 0 Or flexDetalle.Col = 2 Or flexDetalle.Col = 4 Then
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
            If flexDetalle.Row >= 1 And flexDetalle.Col < 5 Then
                If flexDetalle.Col = 4 Then
                    If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" Or Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 1)) = "" Or Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2)) = "" Or Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 3)) = "" Then
                        MsgBox("Debe Capturar Primero la Información de los Agrupadores y los Rubros ..", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        flexDetalle.Col = 0
                        Exit Sub
                    End If
                End If
                If flexDetalle.Col = 0 Or flexDetalle.Col = 2 Or flexDetalle.Col = 4 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                End If
                CambiarFormatoTxtenCaptura()
                MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                If Len(Trim(txtFlex.Text)) = 1 Then
                    System.Windows.Forms.SendKeys.Send("{right}")
                End If
                '        ElseIf flexDetalle.Col = 4 Then
                '            flexDetalle.SetFocus
                '            If flexDetalle.Row < flexDetalle.Rows - 1 Then
                '                flexDetalle.Row = flexDetalle.Row + 1
                '                flexDetalle.Col = 0
                '            Else
                '                flexDetalle.Rows = flexDetalle.Rows + 1
                '                flexDetalle.Row = flexDetalle.Row + 1
                '                flexDetalle.TopRow = flexDetalle.Row
                '                flexDetalle.Col = 0
                '            End If
            End If
        ElseIf eventArgs.keyAscii = System.Windows.Forms.Keys.Escape Then
            Exit Sub
        Else
            If Nuevo Then
                System.Windows.Forms.SendKeys.SendWait("{tab}")
                Exit Sub
            End If
            blnBuscar = False
            If flexDetalle.Col = 0 Or flexDetalle.Col = 1 Then
                flexDetalle.Col = 2
            ElseIf flexDetalle.Col = 2 Or flexDetalle.Col = 3 Then
                If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 2)) <> "" Then
                    flexDetalle.Col = 4
                End If
            End If
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioOrigenyAplicacion_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
        gblnSalir = False
        'If Nuevo Then Limpiar
        Select Case Me.Tag
            Case "frmPagos2"
                gstrMovimiento = "S"
            Case "frmDepositos", "frmDepositosIntPes", "frmDepositosIntDol"
                gstrMovimiento = "E"
            Case "frmCargos"
                gstrMovimiento = "S"
            Case "frmAnticipos2"
                gstrMovimiento = "S"
            Case "frmOtrosIngresos"
                gstrMovimiento = "E"
            Case "frmConsultaOrigenAplicacion"
                If Not Carga Then
                    If Trim(VB.Left(lblFolio.Text, 1)) = "E" Then
                        gstrMovimiento = "S"
                    ElseIf Trim(VB.Left(lblFolio.Text, 1)) = "I" Then
                        gstrMovimiento = "E"
                    End If
                    gStrSql = "SELECT M.FolioMovto,RIGHT('0000' + LTRIM(M.CodOrigenAplicR),4) AS CodOrigenAplicR,O.DescOrigenAplicR,RIGHT('000000' + LTRIM(M.CodRubro),6) AS CodRubro,R.DescRubro,M.Importe,MB.FechaMovto," & "RTRIM(LTRIM(CONVERT(NVARCHAR(10),M.CodOrigenAplicR))) + RTRIM(LTRIM(CONVERT(NVARCHAR(10),M.CodRubro))) AS Llave " & "FROM MovimientosOrigenAplic M INNER JOIN MovimientosBancarios MB ON M.FolioMovto = MB.FolioMovto INNER JOIN CatOrigenAplicRecursos O ON M.CodOrigenAplicR = O.CodOrigenAplicR " & "INNER JOIN CatRubrosOrigenAplicRecursos R ON M.CodOrigenAplicR = R.CodOrigAplicR AND M.CodRubro = R.CodRubro " & "WHERE M.FolioMovto = '" & Trim(lblFolio.Text) & "' " & "ORDER BY M.CodOrigenAplicR,M.CodRubro"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        With flexDetalle
                            .Row = 1
                            lblFechaMovimiento.Text = VB6.Format(RsGral.Fields("FechaMovto").Value, "DD/MMM/YYYY")
                            Do While Not RsGral.EOF
                                .set_TextMatrix(.Row, 0, RsGral.Fields("CodOrigenAplicR").Value)
                                .set_TextMatrix(.Row, 1, Trim(RsGral.Fields("DescOrigenAplicR").Value))
                                .set_TextMatrix(.Row, 2, RsGral.Fields("CodRubro").Value)
                                .set_TextMatrix(.Row, 3, Trim(RsGral.Fields("DescRubro").Value))
                                .set_TextMatrix(.Row, 4, Format(RsGral.Fields("importe").Value, "###,##0.00"))
                                .set_TextMatrix(.Row, 5, Trim(.get_TextMatrix(.Row, 0)) & .get_TextMatrix(.Row, 2))
                                .set_TextMatrix(.Row, 7, Trim(.get_TextMatrix(.Row, 0)) & .get_TextMatrix(.Row, 2))
                                .set_TextMatrix(.Row, 6, Format(RsGral.Fields("importe").Value, "###,##0.00"))
                                RsGral.MoveNext()
                                If Not RsGral.EOF Then
                                    If .Row = .Rows - 1 Then
                                        .Rows = .Rows + 1
                                    End If
                                    .Row = .Row + 1
                                End If
                            Loop
                            CalculoImporte()
                            lblImporte.Text = lblTotal.Text
                            lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                            .Row = 1
                            .Col = 0
                        End With
                    End If
                    Carga = True
                End If
                CalculoImporte()
        End Select
        cmdAceptar.Enabled = True
        flexDetalle.Enabled = True
    End Sub

    Private Sub frmBancosProcesoDiarioOrigenyAplicacion_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'ModEstandar.ActivaMenu C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO
    End Sub

    Private Sub frmBancosProcesoDiarioOrigenyAplicacion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If Not blnBuscar Then

                    If UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name)) <> "FLEXDETALLE" Then
                        ModEstandar.AvanzarTab(Me)
                    End If
                Else
                    flexDetalle.Focus()
                End If
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioOrigenyAplicacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Public Sub frmBancosProcesoDiarioOrigenyAplicacion_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Encabezado()
        CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        lblTotal.Text = "0.00"
        gblnSalir = False
        Nuevo = False
        Carga = False
    End Sub

    Private Sub frmBancosProcesoDiarioOrigenyAplicacion_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'If gblnSalir Then
        'ModEstandar.ActivaMenu C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO
        ModEstandar.LimpiaDescBarraEstado()
        'Set frmBancosProcesoDiarioRegistrodePagos = Nothing
        'Else
        '   Cancel = 1
        'End If
        If Me.Tag = "frmConsultaOrigenAplicacion" Then
            frmConsultaOrigenAplicacion = Nothing
        End If
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
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
                    Case 0, 1, 2, 3, 4
                        If .Col = 0 And Trim(txtFlex.Text) <> "" Then
                            If LlenaDatosAgrupador() Then
                                .Text = Trim(txtFlex.Text)
                                ValidaLlave()
                                .Col = 2
                            End If
                        ElseIf .Col = 1 And Trim(txtFlex.Text) <> "" Then
                            If DescripcionAgrupador() Then
                                .Text = Trim(txtFlex.Text)
                                ValidaLlave()
                                .Col = 2
                            End If
                        ElseIf .Col = 2 And Trim(txtFlex.Text) <> "" Then
                            If LlenaDatosRubro() Then
                                .Text = Trim(txtFlex.Text)
                                .Col = 4
                                ValidaLlave()
                                txtFlex.Visible = False
                                Exit Sub
                            Else
                                If Trim(.get_TextMatrix(.Row, 0)) = "" And Trim(.get_TextMatrix(.Row, 1)) = "" Then
                                    .set_TextMatrix(.Row, 0, "")
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "0.00")
                                    txtFlex.Visible = False
                                    .Col = 0
                                    CalculoImporte()
                                    Exit Sub
                                End If
                            End If
                        ElseIf .Col = 3 And Trim(txtFlex.Text) <> "" Then
                            If DescripcionRubro() Then
                                .Text = Trim(txtFlex.Text)
                                .Col = 4
                                ValidaLlave()
                                txtFlex.Visible = False
                                Exit Sub
                            Else
                                If Trim(.get_TextMatrix(.Row, 0)) = "" And Trim(.get_TextMatrix(.Row, 1)) = "" Then
                                    .set_TextMatrix(.Row, 0, "")
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "0.00")
                                    txtFlex.Visible = False
                                    .Col = 0
                                    CalculoImporte()
                                    Exit Sub
                                End If
                            End If
                        End If
                        If .Col = 4 Then
                            If CDbl(Numerico(txtFlex.Text)) = 0 Then
                                MsgBox("Debe Teclear una Cantidad Mayor que Cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                                txtFlex.Text = ""
                                txtFlex.Focus()
                                Exit Sub
                            End If
                            .Text = Trim(txtFlex.Text)
                            .set_TextMatrix(.Row, 4, VB6.Format(Numerico(.get_TextMatrix(.Row, 4)), "###,##0.00"))
                            CalculoImporte()
                            If .Row = .Rows - 1 Then
                                .Rows = .Rows + 1
                                .Row = .Row + 1
                                .TopRow = .Row
                            Else
                                .Row = .Row + 1
                            End If
                            .Col = 0
                        Else
                            '.Col = .Col + 1
                        End If
                        txtFlex.Visible = False
                End Select
            ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
                'If ActiveControl.Name = "txtFlex" Then Exit Sub
                If flexDetalle.Col = 4 And CDbl(Numerico(txtFlex.Text)) = 0 Then
                    MsgBox("Debe Teclear una Cantidad Mayor que Cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    txtFlex.Text = ""
                    Exit Sub
                End If
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
                Select Case flexDetalle.Col
                    Case 0
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 1
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                    Case 2
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 3
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                    Case 4
                        ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 15, 2, (txtFlex.SelectionStart))
                End Select
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> txtFlex.Name Then
        '    Exit Sub
        'End If
        If Not blnBuscar And Nuevo Then
            txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
        End If
    End Sub

    'Private Sub txtFlex_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFlex.Validating
    '    Dim Cancel As Boolean = eventArgs.Cancel
    '    If flexDetalle.Col = 4 And CDbl(Numerico(txtFlex.Text)) = 0 Then
    '        MsgBox("Debe Teclear una Cantidad Mayor que Cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
    '        txtFlex.Text = ""
    '        Cancel = True
    '    Else
    '        Cancel = False
    '    End If
    '    eventArgs.Cancel = Cancel
    'End Sub

    Private Sub Limpiar()
        Encabezado()
        Icono(Me, MDIMenuPrincipalCorpo)
        lblTotal.Text = "0.00"
        gblnSalir = False
        Nuevo = False
        flexDetalle.Clear()
        Encabezado()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        'Nuevo()
    End Sub
End Class
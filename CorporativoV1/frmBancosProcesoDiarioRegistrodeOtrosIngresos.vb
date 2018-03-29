Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioRegistrodeOtrosIngresos
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REGISTRO DE OTROS INGRESOS                                                                   *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      LUNES 22 DE SEPTIEMBRE DE 2003                                                               *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdOrigenyAplicacion As System.Windows.Forms.Button
    Public WithEvents cmdDesglose As System.Windows.Forms.Button
    Public WithEvents txtImporte As System.Windows.Forms.TextBox
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents dbcCuentaBancaria As System.Windows.Forms.ComboBox
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents lblCancelada As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtFolioIngreso As System.Windows.Forms.TextBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Friend WithEvents btnBuscar As Button
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdOrigenyAplicacion = New System.Windows.Forms.Button()
        Me.cmdDesglose = New System.Windows.Forms.Button()
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.txtFolioIngreso = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.dbcCuentaBancaria = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.lblCancelada = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOrigenyAplicacion
        '
        Me.cmdOrigenyAplicacion.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrigenyAplicacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOrigenyAplicacion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrigenyAplicacion.Location = New System.Drawing.Point(496, 209)
        Me.cmdOrigenyAplicacion.Name = "cmdOrigenyAplicacion"
        Me.cmdOrigenyAplicacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOrigenyAplicacion.Size = New System.Drawing.Size(108, 23)
        Me.cmdOrigenyAplicacion.TabIndex = 17
        Me.cmdOrigenyAplicacion.Text = "&Origen"
        Me.ToolTip1.SetToolTip(Me.cmdOrigenyAplicacion, "Muestra la Ventana de Captura de Origen.")
        Me.cmdOrigenyAplicacion.UseVisualStyleBackColor = False
        '
        'cmdDesglose
        '
        Me.cmdDesglose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDesglose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDesglose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDesglose.Location = New System.Drawing.Point(496, 184)
        Me.cmdDesglose.Name = "cmdDesglose"
        Me.cmdDesglose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDesglose.Size = New System.Drawing.Size(108, 23)
        Me.cmdDesglose.TabIndex = 16
        Me.cmdDesglose.Text = "&Desglose"
        Me.ToolTip1.SetToolTip(Me.cmdDesglose, "Muestra la Ventana de Captura de Desglose del Ingreso.")
        Me.cmdDesglose.UseVisualStyleBackColor = False
        '
        'txtImporte
        '
        Me.txtImporte.AcceptsReturn = True
        Me.txtImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporte.Location = New System.Drawing.Point(112, 105)
        Me.txtImporte.MaxLength = 18
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporte.Size = New System.Drawing.Size(121, 20)
        Me.txtImporte.TabIndex = 5
        Me.txtImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporte, "Importe del Ingreso.")
        '
        'txtConcepto
        '
        Me.txtConcepto.AcceptsReturn = True
        Me.txtConcepto.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtConcepto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConcepto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConcepto.Location = New System.Drawing.Point(112, 78)
        Me.txtConcepto.MaxLength = 100
        Me.txtConcepto.Name = "txtConcepto"
        Me.txtConcepto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConcepto.Size = New System.Drawing.Size(473, 20)
        Me.txtConcepto.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Concepto del Ingreso")
        '
        'txtFolioIngreso
        '
        Me.txtFolioIngreso.AcceptsReturn = True
        Me.txtFolioIngreso.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioIngreso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioIngreso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioIngreso.Location = New System.Drawing.Point(111, 16)
        Me.txtFolioIngreso.MaxLength = 13
        Me.txtFolioIngreso.Name = "txtFolioIngreso"
        Me.txtFolioIngreso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioIngreso.Size = New System.Drawing.Size(138, 20)
        Me.txtFolioIngreso.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolioIngreso, "Folio del Ingreso.")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtImporte)
        Me.Frame1.Controls.Add(Me.txtConcepto)
        Me.Frame1.Controls.Add(Me.dbcBanco)
        Me.Frame1.Controls.Add(Me.dbcCuentaBancaria)
        Me.Frame1.Controls.Add(Me.Label11)
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.lblMoneda)
        Me.Frame1.Controls.Add(Me.lblCancelada)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 68)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(601, 169)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(112, 24)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(192, 21)
        Me.dbcBanco.TabIndex = 2
        '
        'dbcCuentaBancaria
        '
        Me.dbcCuentaBancaria.Location = New System.Drawing.Point(112, 51)
        Me.dbcCuentaBancaria.Name = "dbcCuentaBancaria"
        Me.dbcCuentaBancaria.Size = New System.Drawing.Size(192, 21)
        Me.dbcCuentaBancaria.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(16, 106)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(49, 21)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "Importe :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(89, 21)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Concepto :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(16, 53)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(89, 21)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Cuenta Bancaria :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 26)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(65, 21)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Banco :"
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(312, 52)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(90, 18)
        Me.lblMoneda.TabIndex = 11
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCancelada
        '
        Me.lblCancelada.BackColor = System.Drawing.SystemColors.Control
        Me.lblCancelada.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCancelada.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblCancelada.Location = New System.Drawing.Point(16, 140)
        Me.lblCancelada.Name = "lblCancelada"
        Me.lblCancelada.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCancelada.Size = New System.Drawing.Size(393, 25)
        Me.lblCancelada.TabIndex = 10
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtFolioIngreso)
        Me.Frame4.Controls.Add(Me.dtpFecha)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(16, 16)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(601, 49)
        Me.Frame4.TabIndex = 6
        Me.Frame4.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(496, 16)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(89, 20)
        Me.dtpFecha.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(89, 21)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Folio de Ingreso :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(448, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 21)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Fecha :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(128, 255)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 100
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(13, 255)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 99
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(243, 256)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 98
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoDiarioRegistrodeOtrosIngresos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(638, 303)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.cmdOrigenyAplicacion)
        Me.Controls.Add(Me.cmdDesglose)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioRegistrodeOtrosIngresos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Deposito de Otros Ingresos"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub


    'Variables
    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim FueraChange As Boolean
    Dim intCodBanco As Integer
    Dim intUltFormaPago As Integer
    Dim tecla As Integer
    Dim strCuentaPesos As String
    Dim strCuentaDolares As String
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim Naturaleza As String
    Public bandera As Boolean = False
    Public strControlActual As String 'Nombre del control actual
    Public ConsultaOtrosIngresos As Boolean

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTFOLIOINGRESO"
                strCaptionForm = "Consulta de Registro de Otros Ingresos"
                gStrSql = "SELECT FolioMovto AS FOLIO,Concepto AS CONCEPTO, FechaMovto AS FECHA,Importe AS IMPORTE FROM MovimientosBancarios " & "WHERE FolioMovto LIKE '" & txtFolioIngreso.Text & "%' AND Movimiento = '" & C_OTROSINGRESOS & "' AND TipoMovto = '" & C_TIPOMOVINGRESO & "' ORDER BY FechaMovto DESC ,FolioMovto DESC"
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
        ConfiguraConsultas(FrmConsultas, 9000, RsGral, strTag, strCaptionForm)


        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTFOLIOINGRESO"
                    ConfiguraConsultas(FrmConsultas, 9000, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 1400) 'Columna del Folio
                    .set_ColWidth(1, 0, 4150) 'Columna del Concepto del Movimiento
                    .set_ColWidth(2, 0, 1200) 'Columna de la Fecha del Movimiento
                    .set_ColWidth(3, 0, 1800) 'Columna del Importe del Movimiento
                    .set_ColAlignment(2, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 2, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 2), "dd/MMM/yyyy"))
                        FrmConsultas.Flexdet.set_TextMatrix(I, 3, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "###,##0.00"))
                    Next I
                    FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
                    FrmConsultas.Left = VB6.TwipsToPixelsX(2970)
                Case "TXTFOLIORETIRO"
                    ConfiguraConsultas(FrmConsultas, 8000, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 2500)
                    .set_ColWidth(1, 0, 2500)
                    .set_ColWidth(2, 0, 2500)
                    .set_ColAlignment(2, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 2, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 2), "dd/MMM/yyyy"))
                    Next I
                    FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
                    FrmConsultas.Left = VB6.TwipsToPixelsX(3500)
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function Guardar() As Boolean
        Dim blnTransaccion As Boolean
        Dim strFolioPesos As String
        Dim strFolioDolares As String
        Dim Ejercicio As Integer
        Dim Periodo As String
        Dim I As Integer
        On Error GoTo Err_Renamed
        System.Windows.Forms.Application.DoEvents()
        If Not mblnNuevo Then
            Exit Function
        End If
        If ValidaDatos() = False Then
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Generar Folio del Movimiento
        Ejercicio = CInt(Format(Year(CDate(dtpFecha.Value)), "0000"))
        Periodo = Format(Month(CDate(dtpFecha.Value)), "00")
        BuscaEjercicio(dtpFecha.Value)
        gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVINGRESO & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFolioIngreso.Text = C_TIPOMOVINGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
            ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVINGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        'Buscar la Naturaleza del Movimiento
        gStrSql = "SELECT * FROM CatBancos WHERE CodBanco = " & intCodBanco
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("ControlInterno").Value = False Then
                Naturaleza = C_NATURALEZACOMERCIAL
            ElseIf RsGral.Fields("ControlInterno").Value = True Then
                Naturaleza = C_NATURALEZAINTERNA
            End If
        End If
        'Guardar el Movimiento Bancario
        ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioIngreso.Text, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_OTROSINGRESOS, C_TIPOMOVINGRESO, Naturaleza, IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_TIPOPAGOJOYERIA, CStr(intCodBanco), dbcCuentaBancaria.Text, "", txtConcepto.Text, "0", "", "0", "01/01/1900", "", txtImporte.Text, "V", "01/01/1900", "", CStr(0), "01/01/1900", C_MODULOBANCOS, "", "", C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar los Movimientos de Origen y Aplicación
        If Not frmOtrosIngresos.GuardarMovimientosOrigenAplicacion("REGISTRO DE OTROS INGRESOS") Then
            Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Function
        End If
        'Guardar el Desglose del Deposite
        If Not frmDesgloseOtrosIngresos.GuardarMovimientosDepositos Then
            Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Function
        End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se ha Generado el Folio de Ingreso " & txtFolioIngreso.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Limpiar()
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Function ChecaGridDepositos() As Boolean
        Dim I As Integer
        ChecaGridDepositos = False
        With frmDesgloseOtrosIngresos.flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    ChecaGridDepositos = True
                End If
            Next
        End With
    End Function

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        Dim Total As Decimal
        Dim Moneda As String
        Dim FolioRetiro As String
        Dim RsAux As New ADODB.Recordset
        If Trim(txtFolioIngreso.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT * FROM MovimientosBancarios MB,CatBancos CB WHERE MB.FolioMovto = '" & txtFolioIngreso.Text & "' AND MB.Movimiento = '" & C_OTROSINGRESOS & "' AND " & "MB.TipoMovto = '" & C_TIPOMOVINGRESO & "' AND CB.CodBanco = MB.CodBanco"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            gStrSql = "SELECT FolioMovto FROM MovimientosBancarios WHERE Referencia = '" & txtFolioIngreso.Text & "' AND Movimiento = '" & C_MOVCANCELACION & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                lblCancelada.Text = "Movimiento de Cancelación : " & RsAux.Fields("FolioMovto").Value
            End If
            txtFolioIngreso.Text = Trim(RsGral.Fields("FolioMovto").Value)
            dtpFecha.Value = Format(RsGral.Fields("FechaMovto").Value, C_FORMATFECHAMOSTRAR)
            FueraChange = True
            dbcBanco.Text = Trim(RsGral.Fields("DescBanco").Value)
            dbcCuentaBancaria.Text = Trim(RsGral.Fields("CtaBancaria").Value)
            FueraChange = False
            txtConcepto.Text = Trim(RsGral.Fields("Concepto").Value)
            txtImporte.Text = VB6.Format(RsGral.Fields("importe").Value, "###,##0.00")
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Text = C_DESCPESOS
                Moneda = C_PESO
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Text = C_DESCDOLARES
                Moneda = C_DOLAR
            End If
            FolioRetiro = RsGral.Fields("FolioRetiro").Value
            gStrSql = "SELECT * FROM MovimientosOrigenAplic MO,CatOrigenAplicRecursos CO,CatRubrosOrigenAplicRecursos CR " & "WHERE FolioMovto = '" & Trim(txtFolioIngreso.Text) & "' AND CO.CodOrigenAplicR = MO.CodOrigenAplicR AND CR.CodRubro = MO.CodRubro AND CO.CodOrigenAplicR = CR.CodOrigAplicR"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With frmOtrosIngresos.flexDetalle
                    I = 1
                    .Row = 1
                    frmOtrosIngresos.lblTotal.Text = "0.00"
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, Format(RsGral.Fields("CodOrigenAplicR").Value, "0000"))
                        .set_TextMatrix(.Row, 1, Trim(RsGral.Fields("DescOrigenAplicR").Value))
                        .set_TextMatrix(.Row, 2, Format(RsGral.Fields("CodRubro").Value, "000000"))
                        .set_TextMatrix(.Row, 3, Trim(RsGral.Fields("DescRubro").Value))
                        .set_TextMatrix(.Row, 4, Format(RsGral.Fields("importe").Value, "###,##0.00"))
                        With frmOtrosIngresos
                            .lblTotal.Text = CStr(CDec(Numerico(Format(.lblTotal.Text, "#####0.00"))) + CDbl(Format(RsGral.Fields("importe").Value, "###,##0.00")))
                        End With
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        .Row = .Row + 1
                        I = I + 1
                        RsGral.MoveNext()
                    Loop
                    frmOtrosIngresos.lblTotal.Text = Format(frmOtrosIngresos.lblTotal.Text, "###,##0.00")
                    frmOtrosIngresos.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    frmOtrosIngresos.Nuevo = True
                End With
            End If
            gStrSql = "SELECT * FROM MovimientosReferencias Where FolioMovto = '" & txtFolioIngreso.Text & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With frmDesgloseOtrosIngresos.flexDetalle
                    I = 1
                    .Row = 1
                    frmDesgloseOtrosIngresos.lblTotal.Text = Format(RsGral.Fields("ImporteDeposito").Value, "###,##0.00")
                    frmDesgloseOtrosIngresos.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("ReferenciaBanco").Value))
                        .set_TextMatrix(.Row, 1, Format(RsGral.Fields("ImporteRef").Value, "###,##0.00"))
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        .Row = .Row + 1
                        I = I + 1
                        RsGral.MoveNext()
                    Loop
                End With
            End If
            mblnNuevo = False
            dtpFecha.Enabled = False
            ConsultaOtrosIngresos = True
        Else
            MsgBox("Folio de Movimiento de Ingreso no Existe ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Frame1.Enabled = True
            txtFolioIngreso.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        InicializaVariables()
        txtFolioIngreso.Text = ""
        txtFolioIngreso.Focus()
    End Sub

    Sub Nuevo()
        lblMoneda.Text = ""
        dtpFecha.Value = DateTime.Now
        lblCancelada.Text = ""
        dbcBanco.Text = ""
        'dbcBanco.RowSource = Nothing
        dbcCuentaBancaria.Text = ""
        'dbcCuentaBancaria.RowSource = Nothing
        txtConcepto.Text = ""
        txtImporte.Text = "0.00"
        Frame1.Enabled = True
        InicializaVariables()
        gblnSalir = True
        frmOtrosIngresos.Close()
        gblnSalir = False
        frmDesgloseOtrosIngresos.Close()
        cmdDesglose.Enabled = True
        cmdOrigenyAplicacion.Enabled = True
        ConsultaOtrosIngresos = False
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
        FueraChange = False
        intCodBanco = 0
        Naturaleza = ""
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Not BuscaUltimoCierre(dtpFecha.Value) Then
            Exit Function
        End If
        If Len(Trim(dbcBanco.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Nombre del Banco", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dbcBanco.Focus()
            Exit Function
        End If
        If Len(Trim(dbcCuentaBancaria.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Cuenta Bancaria", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
            Exit Function
        End If
        If Len(Trim(txtConcepto.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Concepto", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtConcepto.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtImporte.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Importe del Pago", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtImporte.Focus()
            Exit Function
        End If
        If Not ChecaGrid(frmOtrosIngresos) Then
            MsgBox("No se Han Capturado los Movimientos de Origen ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
            Exit Function
        End If
        If ChecaGridDepositos() Then
            If Numerico((frmDesgloseOtrosIngresos.lblImporte).Text) <> Numerico((frmDesgloseOtrosIngresos.lblTotal).Text) Then
                MsgBox("El Total del Desglose de Depositos no es Igual al Importe del Deposito ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                cmdDesglose_Click(cmdDesglose, New System.EventArgs())
                Exit Function
            End If
        End If
        If Numerico((frmOtrosIngresos.lblImporte).Text) <> Numerico((frmOtrosIngresos.lblTotal).Text) Then
            MsgBox("El Total de los Movimientos de Origen no es Igual al Importe del Pago...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub cmdDesglose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDesglose.Click
        frmDesgloseOtrosIngresos.InitializeComponent()
        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            If CDbl(Numerico(txtImporte.Text)) > 0 Then
                If Not mblnNuevo Then
                    frmDesgloseOtrosIngresos.cmdAceptar.TabIndex = 0
                    frmDesgloseOtrosIngresos.flexDetalle.TabIndex = 1
                    frmDesgloseOtrosIngresos.flexDetalle.Enabled = False
                Else
                    frmDesgloseOtrosIngresos.cmdAceptar.TabIndex = 1
                    frmDesgloseOtrosIngresos.flexDetalle.TabIndex = 0
                    frmDesgloseOtrosIngresos.cmdAceptar.Enabled = False
                End If
                frmDesgloseOtrosIngresos.Text = "Desglose de Depósitos Bancarios"
                frmDesgloseOtrosIngresos.Label1.Text = "Importe del Deposito : "
                frmDesgloseOtrosIngresos.Panel1.Text = "Desglose del Depósito"
                frmDesgloseOtrosIngresos.lblMoneda.Text = lblMoneda.Text
                frmDesgloseOtrosIngresos.lblImporte.Text = txtImporte.Text
                frmDesgloseOtrosIngresos.flexDetalle.Col = 0
                frmDesgloseOtrosIngresos.flexDetalle.Row = 1
                frmDesgloseOtrosIngresos.Tag = "frmDesgloseOtrosIngresos"
                frmDesgloseOtrosIngresos.ShowDialog()
            Else
                MsgBox("El Importe del Depósito debe ser Mayor que Cero, Favor de Teclear un Importe ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtImporte.Focus()
            End If
        Else
            MsgBox("Favor de Seleccionar Una Cuenta Bancaria Valida ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
        End If
    End Sub

    Private Sub cmdDesglose_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDesglose.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdOrigenyAplicacion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigenyAplicacion.Click
        frmOtrosIngresos.InitializeComponent()
        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            If CDbl(Numerico(txtImporte.Text)) > 0 Then
                If frmOtrosIngresos.Nuevo Then
                    frmOtrosIngresos.cmdAceptar.TabIndex = 0
                    frmOtrosIngresos.flexDetalle.TabIndex = 1
                    frmOtrosIngresos.flexDetalle.Enabled = False
                Else
                    frmOtrosIngresos.flexDetalle.TabIndex = 0
                    frmOtrosIngresos.cmdAceptar.TabIndex = 1
                    frmOtrosIngresos.cmdAceptar.Enabled = False
                End If
                frmOtrosIngresos.Tag = "frmOtrosIngresos"
                frmOtrosIngresos.Text = "Origen de Recursos (Registro de Otros Ingresos)"
                frmOtrosIngresos.lblMoneda.Text = lblMoneda.Text
                frmOtrosIngresos.lblFechaMovimiento.Text = dtpFecha.Value
                frmOtrosIngresos.lblImporte.Text = txtImporte.Text
                frmOtrosIngresos.flexDetalle.Col = 0
                frmOtrosIngresos.flexDetalle.Row = 1
                frmOtrosIngresos.ShowDialog()
            Else
                MsgBox("El Importe del Depósito debe ser Mayor que Cero, Favor de Teclear un Importe ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtImporte.Focus()
            End If
        Else
            MsgBox("Favor de Seleccionar Una Cuenta Bancaria Valida ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
        End If
    End Sub

    Private Sub cmdOrigenyAplicacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigenyAplicacion.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcBanco_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcBanco" Then
        '    Exit Sub
        'End If
        dbcCuentaBancaria.Text = ""
        lblMoneda.Text = ""
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' /*AND ControlInterno = 0*/ ORDER BY DescBanco"
        DCChange(gStrSql, tecla)
        intCodBanco = 0
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos /*WHERE ControlInterno = 0*/ ORDER BY DescBanco"
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.KeyDown
        'tecla = eventSender.keyCode
        'If eventSender.keyCode = System.Windows.Forms.Keys.Escape Then
        '    txtFolioIngreso.Focus()
        'End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' /*AND ControlInterno = 0*/ ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
    End Sub

    Private Sub dbcCuentaBancaria_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcCuentaBancaria" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCChange(gStrSql, tecla)
        If Trim(dbcCuentaBancaria.Text) = "" Then
            lblMoneda.Text = ""
        End If
        'intCodBanco = 0
    End Sub

    Private Sub dbcCuentaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Enter
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCGotFocus(gStrSql, dbcCuentaBancaria)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcCuentaBancaria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.KeyDown
        'tecla = eventSender.keyCode
        'If eventSender.keyCode = System.Windows.Forms.Keys.Escape Then
        '    dbcBanco.Focus()
        'End If
    End Sub

    Private Sub dbcCuentaBancaria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcCuentaBancaria_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.KeyUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        '    dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dbcCuentaBancaria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Leave
        On Error GoTo Err_Renamed
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCLostFocus(dbcCuentaBancaria, gStrSql, intCodBanco)
        gStrSql = "SELECT Moneda FROM CatCuentasBancarias WHERE CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCDOLARES
            End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcCuentaBancaria_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.MouseUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        '    dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeOtrosIngresos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeOtrosIngresos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeOtrosIngresos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return

                If Me.ActiveControl.Name = "txtFolioIngreso" Then
                    If Len(Trim(txtFolioIngreso.Text)) = 13 And VB.Right(txtFolioIngreso.Text, 4) <> "0000" Then
                        Frame1.Enabled = False
                    End If
                End If
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape

                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioIngreso" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeOtrosIngresos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeOtrosIngresos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        Nuevo()
        BuscaEjercicio(dtpFecha.Value)
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeOtrosIngresos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmBancosProcesoDiarioRegistrodeOtrosIngresos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
        Me.Hide()
        gblnSalir = True
        frmOtrosIngresos.Close()
        frmOtrosIngresos = Nothing
        frmDesgloseOtrosIngresos.Close()
        frmDesgloseOtrosIngresos = Nothing
    End Sub

    Private Sub txtConcepto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Enter
        SelTextoTxt(txtConcepto)
        Pon_Tool()
    End Sub

    Private Sub txtConcepto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConcepto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¡¿*,;.:<>@+-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioIngreso_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioIngreso.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioIngreso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioIngreso.Enter
        strControlActual = UCase("txtFolioIngreso")
        SelTextoTxt(txtFolioIngreso)
        Pon_Tool()
    End Sub

    Private Sub txtFolioIngreso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioIngreso.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_TIPOMOVINGRESO)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioIngreso_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioIngreso.Leave

        If Me.ActiveControl.Name = "btnBuscar" Then
            Exit Sub
        End If

        If Trim(txtFolioIngreso.Text) = "" Then
            txtFolioIngreso.Text = C_TIPOMOVINGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & "0000"
            Exit Sub
        End If
        If mblnCambiosEnCodigo = True And txtFolioIngreso.Text <> "" And VB.Right(txtFolioIngreso.Text, 4) <> "0000" Then
            LlenaDatos()
            frmOtrosIngresos.Hide()
        End If
    End Sub

    Private Sub txtImporte_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.TextChanged
        If txtImporte.Text = "" Then
            txtImporte.Text = "0.00"
        End If
    End Sub

    Private Sub txtImporte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Enter
        SelTextoTxt(txtImporte)
        Pon_Tool()
    End Sub

    Private Sub txtImporte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImporte.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtImporte.Text, KeyAscii, 15, 2, (txtImporte.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImporte_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Leave
        txtImporte.Text = VB6.Format(txtImporte.Text, "###,##0.00")
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class
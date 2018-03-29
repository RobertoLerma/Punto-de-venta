Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports VB = Microsoft.VisualBasic

Public Class frmBancosProcesoDiarioReferenciaVouchers
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REFERENCIA DE VOUCHERS                                                                       *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      JUEVES 16 DE OCTUBRE DE 2003                                                                 *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lblImportenoAcreditado As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    'Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents cmdImportarVouchers As System.Windows.Forms.Button
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Label2 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents lblDeposito As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label4 As Label
    Public Nuevo As Boolean

    Sub Encabezado()
        With flexDetalle
            .Row = 0
            .Col = 0
            .set_ColWidth(0, 0, 2700)
            .CellAlignment = 5
            .CellFontSize = 8
            .CellFontBold = True
            .Text = "VOUCHER"
            .Col = 1
            .set_ColWidth(1, 0, 1100)
            .CellAlignment = 5
            .CellFontSize = 8
            .CellFontBold = True
            .Text = "FECHA"
            .Col = 2
            .set_ColWidth(2, 0, 3400)
            .CellAlignment = 5
            .CellFontSize = 8
            .CellFontBold = True
            .Text = "FORMA DE PAGO"
            .Col = 3
            .set_ColWidth(3, 0, 1450)
            .CellAlignment = 5
            .CellFontSize = 8
            .CellFontBold = True
            .Text = "IMPORTE REAL"
            .Col = 4
            .set_ColWidth(4, 0, 1450)
            .CellAlignment = 5
            .CellFontSize = 8
            .CellFontBold = True
            .Text = "IMPTE. S/COM"
            .Col = 5
            .set_ColWidth(5, 0, 0)
            .Col = 6
            .set_ColWidth(6, 0, 0)
            .Rows = 11
            .Col = 0
            .Row = 1
        End With
    End Sub

    Function EstaVacia() As Boolean
        Dim I As Integer
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "" Then
                    EstaVacia = False
                    Exit Function
                End If
            Next
            EstaVacia = True
        End With
    End Function

    Sub EliminaRenglon()
        flexDetalle.RemoveItem((flexDetalle.Row))
        flexDetalle.Rows = flexDetalle.Rows + 1
    End Sub

    Function Guardar() As Boolean
        Dim NumPartida As Integer
        Dim I As Integer
        On Error GoTo Err_Renamed
        Guardar = True
        With flexDetalle
            NumPartida = 1
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "" Then
                    ModStoredProcedures.PR_IMEMovimientosReferencias((frmBancosProcesoDiarioRegistrodeDepositos.txtFolioIngreso).Text, CStr(NumPartida), VB6.Format(lblDeposito.Text, "#####0.00"), .get_TextMatrix(I, 0), VB6.Format(.get_TextMatrix(I, 4), "#####0.00"), "V", "V", C_INSERCION, CStr(0))
                    Cmd.Execute()
                    ModStoredProcedures.PR_IEIngresosFormasdePago(.get_TextMatrix(I, 5), .get_TextMatrix(I, 6), "01/01/1900", "", "0", "0", "0", "0", "", "", "", "", "0", "0", "0", "", "01/01/1900", "1", VB6.Format(Today, C_FORMATFECHAGUARDAR), "0", C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                End If
            Next
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            Guardar = False
        End If
    End Function

    Sub ObtenerImporteTarjetasnoAcreditadas()
        On Error GoTo Err_Renamed
        Dim Fecha As String
        Dim Moneda As Byte
        Fecha = VB6.Format(System.DateTime.FromOADate(Today.ToOADate - 1), "dd/mmm/yyyy")
        If VB.Left(frmBancosProcesoDiarioRegistrodeDepositos.lblMoneda.Text, 1) = "P" Then
            Moneda = 0
        ElseIf VB.Left(frmBancosProcesoDiarioRegistrodeDepositos.lblMoneda.Text, 1) = "D" Then
            Moneda = 1
        End If
        gStrSql = "SELECT ISNULL(SUM(INGFP.ImpSinCom),0) AS ImporteSinComision " & "FROM (SELECT FolioIngreso,FechaIngreso,CodSucursal FROM INGRESOS WHERE Estatus <> 'C') ING " & "INNER JOIN (SELECT ING.FolioIngreso,ING.FechaIngreso,ING.NumPartida,ROUND(ING.Importe,2) AS ImporteReal," & "ROUND(ING.Importe - (ING.ComisionBancaria + ING.InteresesPromocion),2) AS ImpSinCom,ING.CodPlan , ING.CodBanco," & "ING.Autorizacion , CFP.DescFormaPago, CFP.EsDolar " & "FROM INGRESOSFORMADEPAGO ING INNER JOIN CATFORMASPAGO CFP ON ING.CODFORMAPAGO = CFP.CODFORMAPAGO " & "WHERE CFP.EsTarjeta = 1 AND ING.Estatus <> 'C' AND ING.PasoBancos = 0) INGFP ON ING.FOLIOINGRESO = INGFP.FOLIOINGRESO " & "LEFT OUTER JOIN CatPlanesXBanco CPB ON INGFP.CodPlan = CPB.CodPlan AND INGFP.CodBanco = CPB.CodBanco " & "WHERE /*INGFP.CodBanco = " & frmBancosProcesoDiarioRegistrodeDepositos.lblCodBanco.Text & " AND*/ INGFP.FechaIngreso <= '" & VB6.Format(Fecha, "mm/dd/yyyy") & "' /*AND INGFP.EsDolar = */" '& Moneda
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.Fields("ImporteSinComision").Value <> 0 Then
            lblImportenoAcreditado.Text = VB6.Format(RsGral.Fields("ImporteSinComision").Value, "###,##0.00")
        Else
            lblImportenoAcreditado.Text = "0.00"
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        If Nuevo Then
            If Not EstaVacia() Then
                frmBancosProcesoDiarioRegistrodeDepositos.cmdDesglose.Enabled = False
                frmBancosProcesoDiarioRegistrodeDepositos.txtImporte.Text = lblDeposito.Text
            Else
                frmBancosProcesoDiarioRegistrodeDepositos.cmdDesglose.Enabled = True
                frmBancosProcesoDiarioRegistrodeDepositos.cmdReferencias.Enabled = True
                frmBancosProcesoDiarioRegistrodeDepositos.txtImporte.Text = lblDeposito.Text
            End If
        End If
        Me.Hide()
    End Sub

    Private Sub cmdImportarVouchers_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImportarVouchers.Click
        frmBancosProcesoDiarioImportacionVouchers.Tag = "IMPORTACION"
        frmBancosProcesoDiarioImportacionVouchers.Text = "Importación de Vouchers"
        frmBancosProcesoDiarioImportacionVouchers.ShowDialog()
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        'System.Windows.Forms.SendKeys.SendWait("{right}")
    End Sub
    Private Sub frmBancosProcesoDiarioReferenciaVouchers_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ObtenerImporteTarjetasnoAcreditadas()
        If Nuevo Then
            flexDetalle.Focus()
        Else
            cmdAceptar.Focus()
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioReferenciaVouchers_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            ModEstandar.AvanzarTab(Me)
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        ElseIf KeyCode = System.Windows.Forms.Keys.Delete And Nuevo Then
            EliminaRenglon()
            frmBancosProcesoDiarioImportacionVouchers.ObtenerTotal()
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioReferenciaVouchers_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Encabezado()
        lblDeposito.Text = "0.00"
        ObtenerImporteTarjetasnoAcreditadas()
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoDiarioReferenciaVouchers))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblImportenoAcreditado = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdImportarVouchers = New System.Windows.Forms.Button()
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.lblDeposito = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblImportenoAcreditado
        '
        Me.lblImportenoAcreditado.BackColor = System.Drawing.SystemColors.Window
        Me.lblImportenoAcreditado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImportenoAcreditado.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImportenoAcreditado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImportenoAcreditado.Location = New System.Drawing.Point(589, 183)
        Me.lblImportenoAcreditado.Name = "lblImportenoAcreditado"
        Me.lblImportenoAcreditado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImportenoAcreditado.Size = New System.Drawing.Size(121, 21)
        Me.lblImportenoAcreditado.TabIndex = 10
        Me.lblImportenoAcreditado.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(480, 183)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(103, 30)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Importe de Tarjetas  no Acreditadas"
        '
        'cmdImportarVouchers
        '
        Me.cmdImportarVouchers.BackColor = System.Drawing.SystemColors.Control
        Me.cmdImportarVouchers.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdImportarVouchers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdImportarVouchers.Location = New System.Drawing.Point(15, 168)
        Me.cmdImportarVouchers.Name = "cmdImportarVouchers"
        Me.cmdImportarVouchers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdImportarVouchers.Size = New System.Drawing.Size(105, 45)
        Me.cmdImportarVouchers.TabIndex = 4
        Me.cmdImportarVouchers.Text = "&Importar Vouchers"
        Me.cmdImportarVouchers.UseVisualStyleBackColor = False
        '
        'cmdAceptar
        '
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Location = New System.Drawing.Point(126, 168)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.Size = New System.Drawing.Size(105, 45)
        Me.cmdAceptar.TabIndex = 5
        Me.cmdAceptar.Text = "&Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(10, 13)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(700, 130)
        Me.flexDetalle.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(289, 191)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(136, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Supr-Eliminar Renglón"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(16, 7)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(105, 21)
        Me.lblMoneda.TabIndex = 7
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblDeposito
        '
        Me.lblDeposito.BackColor = System.Drawing.SystemColors.Window
        Me.lblDeposito.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDeposito.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDeposito.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDeposito.Location = New System.Drawing.Point(582, 12)
        Me.lblDeposito.Name = "lblDeposito"
        Me.lblDeposito.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDeposito.Size = New System.Drawing.Size(121, 21)
        Me.lblDeposito.TabIndex = 2
        Me.lblDeposito.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(494, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(89, 21)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Importe Depósito"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.flexDetalle)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblImportenoAcreditado)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cmdImportarVouchers)
        Me.Panel1.Controls.Add(Me.cmdAceptar)
        Me.Panel1.Location = New System.Drawing.Point(12, 54)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(721, 231)
        Me.Panel1.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 38)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(115, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Movimientos bancarios"
        '
        'frmBancosProcesoDiarioReferenciaVouchers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(745, 297)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblMoneda)
        Me.Controls.Add(Me.lblDeposito)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(164, 187)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBancosProcesoDiarioReferenciaVouchers"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Referencia Vouchers"
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class
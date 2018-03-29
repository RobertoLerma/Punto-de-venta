Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports VB = Microsoft.VisualBasic
Public Class frmBancosProcesoDiarioImportacionVouchers
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents lblSeleccion As System.Windows.Forms.Label
    Public WithEvents lblFecha As System.Windows.Forms.Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Public WithEvents Label1 As System.Windows.Forms.Label

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoDiarioImportacionVouchers))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblSeleccion = New System.Windows.Forms.Label()
        Me.lblFecha = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdAceptar
        '
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Location = New System.Drawing.Point(192, 265)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.Size = New System.Drawing.Size(105, 37)
        Me.cmdAceptar.TabIndex = 2
        Me.cmdAceptar.Text = "&Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(13, 12)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(251, 149)
        Me.flexDetalle.TabIndex = 1
        '
        'chkTodaslasSucursales
        '
        Me.chkTodaslasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodaslasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodaslasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodaslasSucursales.Location = New System.Drawing.Point(16, 16)
        Me.chkTodaslasSucursales.Name = "chkTodaslasSucursales"
        Me.chkTodaslasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(138, 21)
        Me.chkTodaslasSucursales.TabIndex = 0
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 312)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(281, 29)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Presione la Barra Espaciadora o Haga Doble Click en el Grid para Seleccionar una " &
    "Sucursal"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(48, 270)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(106, 29)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Sucursal(s) Seleccionada(s)"
        '
        'lblSeleccion
        '
        Me.lblSeleccion.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSeleccion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSeleccion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSeleccion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSeleccion.Location = New System.Drawing.Point(16, 272)
        Me.lblSeleccion.Name = "lblSeleccion"
        Me.lblSeleccion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSeleccion.Size = New System.Drawing.Size(21, 21)
        Me.lblSeleccion.TabIndex = 6
        '
        'lblFecha
        '
        Me.lblFecha.BackColor = System.Drawing.SystemColors.Window
        Me.lblFecha.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFecha.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFecha.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFecha.Location = New System.Drawing.Point(192, 232)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFecha.Size = New System.Drawing.Size(105, 21)
        Me.lblFecha.TabIndex = 5
        Me.lblFecha.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(136, 234)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(57, 21)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Hasta el :"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.flexDetalle)
        Me.Panel1.Location = New System.Drawing.Point(16, 52)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(281, 173)
        Me.Panel1.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(29, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Sucursales"
        '
        'frmBancosProcesoDiarioImportacionVouchers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(314, 351)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdAceptar)
        Me.Controls.Add(Me.chkTodaslasSucursales)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblSeleccion)
        Me.Controls.Add(Me.lblFecha)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(413, 135)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBancosProcesoDiarioImportacionVouchers"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Importación de Vouchers"
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Dim I As Integer
    Dim FueraClick As Boolean
    Dim ListaSucursales As String

    Sub Encabezado()
        With flexDetalle
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 3500)
            .CellAlignment = 5
            .CellFontSize = 10
            .CellFontBold = True
            .Text = "Sucursal"
            .Col = 1
            .set_ColWidth(1, 0, 0)
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub SeleccionarTodasLasSucursales()
        With flexDetalle
            For I = 1 To .Rows - 1
                If chkTodaslasSucursales.CheckState = 1 Then
                    .Row = I
                    .CellBackColor = lblSeleccion.BackColor
                ElseIf chkTodaslasSucursales.CheckState = 0 Then
                    .Row = I
                    .CellBackColor = .BackColor
                End If
            Next
        End With
    End Sub

    Function EstanTodosSeleccionados() As Boolean
        With flexDetalle
            For I = 1 To .Rows - 1
                .Row = I
                If .CellBackColor.Equals(.BackColor) Then
                    EstanTodosSeleccionados = False
                    Exit Function
                End If
            Next
        End With
        EstanTodosSeleccionados = True
    End Function

    Sub MostrarSucursales()
        On Error GoTo Err_Renamed
        gStrSql = "SELECT * FROM CatAlmacen WHERE TipoAlmacen = 'P' ORDER BY DescAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With flexDetalle
                .Row = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescAlmacen").Value))
                    .set_TextMatrix(.Row, 1, Trim(RsGral.Fields("CodAlmacen").Value))
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Row = .Row + 1
                    RsGral.MoveNext()
                Loop
                .Rows = RsGral.RecordCount + 1
                .Row = 1
            End With
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub ObtenerSucursales()
        Dim NumSucursales As Integer
        ListaSucursales = ""
        With flexDetalle
            NumSucursales = 1
            For I = 1 To .Rows - 1
                .Row = I
                If System.Drawing.ColorTranslator.ToOle(.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblSeleccion.BackColor) Then
                    If NumSucursales = 1 Then
                        ListaSucursales = .get_TextMatrix(.Row, 1)
                    ElseIf NumSucursales > 1 Then
                        ListaSucursales = ListaSucursales & "," & .get_TextMatrix(.Row, 1)
                    End If
                    NumSucursales = NumSucursales + 1
                End If
            Next
        End With
    End Sub

    Sub ObtenerTotal()
        frmBancosProcesoDiarioReferenciaVouchers.lblDeposito.Text = CStr(0)
        With frmBancosProcesoDiarioReferenciaVouchers.flexDetalle
            For I = 1 To .Rows - 1
                frmBancosProcesoDiarioReferenciaVouchers.lblDeposito.Text = CStr(CDbl(Numerico((frmBancosProcesoDiarioReferenciaVouchers.lblDeposito).Text)) + CDbl(Numerico(.get_TextMatrix(I, 4))))
            Next
        End With
        frmBancosProcesoDiarioReferenciaVouchers.lblDeposito.Text = Format(frmBancosProcesoDiarioReferenciaVouchers.lblDeposito.Text, "###,##0.00")
    End Sub

    Sub PonerColor()
        Dim Ren As Integer
        flexDetalle.Col = 0
        If flexDetalle.CellBackColor.Equals(flexDetalle.BackColor) Then
            flexDetalle.CellBackColor = lblSeleccion.BackColor
        ElseIf System.Drawing.ColorTranslator.ToOle(flexDetalle.CellBackColor) = System.Drawing.ColorTranslator.ToOle(Me.lblSeleccion.BackColor) Then
            flexDetalle.CellBackColor = flexDetalle.BackColor
        End If
        Ren = flexDetalle.Row
        If EstanTodosSeleccionados() Then
            chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            FueraClick = True
            chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
            FueraClick = False
        End If
        flexDetalle.Row = Ren
    End Sub
    Private Sub chkTodaslasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.CheckStateChanged
        If FueraClick Then Exit Sub
        SeleccionarTodasLasSucursales()
    End Sub

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        Dim frmBancosProcesoDiarioConsultaIngresosXSucursal As New frmBancosProcesoDiarioConsultaIngresosXSucursal()
        frmBancosProcesoDiarioConsultaIngresosXSucursal.InitializeComponent()
        Dim Moneda As Integer
        On Error GoTo Err_Renamed
        Dim strWhere As String
        Dim Sucursales As String
        If Me.Tag = "IMPORTACION" Then
            If VB.Left(frmBancosProcesoDiarioRegistrodeDepositos.lblMoneda.Text, 1) = "P" Then
                Moneda = 0
            ElseIf VB.Left(frmBancosProcesoDiarioRegistrodeDepositos.lblMoneda.Text, 1) = "D" Then
                Moneda = 1
            End If
            strWhere = "WHERE INGFP.CodBanco = " & frmBancosProcesoDiarioRegistrodeDepositos.lblCodBanco.Text & " AND INGFP.FechaIngreso <= '" & Format(lblFecha.Text, "MM/dd/yyyy") & "' AND INGFP.EsDolar = " & Moneda & " "
            'strWhere = "WHERE INGFP.CodBanco = " & frmBancosProcesoDiarioRegistrodeDepositos.lblCodBanco & " AND INGFP.FechaIngreso <= '" & Format(lblFecha, "mm/dd/yyyy") & "' "
            If chkTodaslasSucursales.CheckState = 0 Then
                ObtenerSucursales()
                If Trim(ListaSucursales) = "" Then
                    Me.Close()
                    Exit Sub
                End If
                strWhere = "WHERE INGFP.CodBanco = " & frmBancosProcesoDiarioRegistrodeDepositos.lblCodBanco.Text & " AND INGFP.FechaIngreso <= '" & Format(lblFecha.Text, "MM/dd/yyyy") & "' AND INGFP.EsDolar = " & Moneda & " AND ING.CodSucursal IN(" & ListaSucursales & ") "
                'strWhere = "WHERE INGFP.CodBanco = " & frmBancosProcesoDiarioRegistrodeDepositos.lblCodBanco & " AND INGFP.FechaIngreso <= '" & Format(lblFecha, "mm/dd/yyyy") & "' AND ING.CodSucursal IN(" & ListaSucursales & ") "
            End If
            gStrSql = "SELECT ING.FolioIngreso,INGFP.FechaIngreso,ING.CodSucursal,INGFP.ImporteReal,INGFP.ImpSinCom," & "INGFP.CodPlan,INGFP.CodBanco,INGFP.NumPartida," & "'Suc ' + RIGHT('00' + RTRIM(CAST(ING.CodSucursal AS Char(2))),2) + ' Aut-' + INGFP.Autorizacion AS Voucher," & "RTRIM(INGFP.DescFormaPago) + ' ' + ISNULL(CPB.DescPlan,'') AS FormaPago " & "FROM " & "(SELECT FolioIngreso,FechaIngreso,CodSucursal FROM INGRESOS WHERE Estatus <> 'C') ING " & "INNER JOIN " & "(SELECT ING.FolioIngreso,ING.FechaIngreso,ING.NumPartida," & "ROUND(ING.Importe,2) AS ImporteReal," & "ROUND(ING.Importe - (ING.ComisionBancaria + ING.InteresesPromocion),2) AS ImpSinCom," & "ING.CodPlan , ING.CodBanco, ING.Autorizacion, CFP.DescFormaPago,CFP.EsDolar " & "FROM INGRESOSFORMADEPAGO ING INNER JOIN CATFORMASPAGO CFP " & "ON ING.CODFORMAPAGO = CFP.CODFORMAPAGO " & "WHERE CFP.EsTarjeta = 1 AND ING.Estatus <> 'C' AND ING.PasoBancos = 0) " & "INGFP " & "ON ING.FOLIOINGRESO = INGFP.FOLIOINGRESO " & "LEFT OUTER JOIN CatPlanesXBanco CPB " & "ON INGFP.CodPlan = CPB.CodPlan AND INGFP.CodBanco = CPB.CodBanco " & strWhere & "ORDER BY INGFP.FechaIngreso,ING.FolioIngreso,INGFP.NumPartida"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                frmBancosProcesoDiarioReferenciaVouchers.flexDetalle.Clear()
                frmBancosProcesoDiarioReferenciaVouchers.Encabezado()
                With frmBancosProcesoDiarioReferenciaVouchers.flexDetalle
                    .Row = 1
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, VB.Left(Trim(RsGral.Fields("Voucher").Value) & Space(30), 23))
                        .set_TextMatrix(.Row, 0, .get_TextMatrix(.Row, 0) & VB.Left(Format(RsGral.Fields("FechaIngreso").Value, "dd/mmm/yyyy"), 7))
                        .set_TextMatrix(.Row, 1, Format(RsGral.Fields("FechaIngreso").Value, "dd/mmm/yyyy"))
                        .set_TextMatrix(.Row, 2, Trim(RsGral.Fields("FormaPago").Value))
                        .set_TextMatrix(.Row, 3, Format(RsGral.Fields("ImporteReal").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, 4, Format(RsGral.Fields("ImpSinCom").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, 5, RsGral.Fields("FolioIngreso").Value)
                        .set_TextMatrix(.Row, 6, RsGral.Fields("NumPartida").Value)
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        .Row = .Row + 1
                        RsGral.MoveNext()
                    Loop
                    If RsGral.RecordCount > 10 Then
                        .Rows = RsGral.RecordCount + 1
                    Else
                        .Rows = 11
                    End If
                    .Row = 1
                    ObtenerTotal()
                End With
            Else
                MsgBox("No existen movimientos bancarios en esta sucursal(s)...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                frmBancosProcesoDiarioReferenciaVouchers.flexDetalle.Clear()
                frmBancosProcesoDiarioReferenciaVouchers.Encabezado()
                ObtenerTotal()
            End If
            Me.Close()
        ElseIf Me.Tag = "CONSULTASALDOS" Then
            frmBancosProcesoDiarioConsultaIngresosXSucursal.dtpFecha.Value = Now
            If chkTodaslasSucursales.CheckState = 0 Then
                ObtenerSucursales()
                If Trim(ListaSucursales) = "" Then
                    Me.Close()
                    Exit Sub
                End If
                Sucursales = " AND C.CodAlmacen IN(" & ListaSucursales & ") "
            Else
                Sucursales = ""
            End If
            gStrSql = "SELECT C.CodAlmacen,C.DescAlmacen,SUM(ISNULL(ImpDolIng,0)) - (SUM(ABS(ISNULL(ImpDolDev,0))) + SUM(ISNULL(ImpDolRet,0))) AS EfectivoDolares," & "SUM(ISNULL(ImpPesIng,0)) - (SUM(ABS(ISNULL(ImpPesDev,0))) + SUM(ISNULL(ImpPesRet,0))) AS EfectivoPesos," & "SUM(ISNULL(ImpPesTar, 0)) As ImporteTarjetas " & "FROM " & "(SELECT I.CodSucursal,SUM(ISNULL(CASE WHEN I.Tipo = 'I' THEN ImporteDolares END,0)) ImpDolIng," & "SUM(ISNULL(CASE WHEN I.Tipo = 'I' THEN ImportePesos END,0)) ImpPesIng," & "SUM(ISNULL(CASE WHEN I.Tipo = 'D' THEN ImporteDolares END,0)) ImpDolDev," & "SUM(ISNULL(CASE WHEN I.Tipo = 'D' THEN ImportePesos END,0)) ImpPesDev," & "SUM(ISNULL(CASE WHEN I.Tipo = 'R' THEN ImporteDolares END,0)) ImpDolRet," & "SUM(ISNULL(CASE WHEN I.Tipo = 'R' THEN ImportePesos END,0)) ImpPesRet," & "SUM(ISNULL(CASE WHEN I.Tipo = 'T' THEN ImporteDolares END,0)) ImpDolTar," & "SUM(ISNULL(CASE WHEN I.Tipo = 'T' THEN ImportePesos END,0)) ImpPesTar " & "FROM DBO.vw_Ingresos I " & "WHERE FechaIngreso <= '" & Format(frmBancosProcesoDiarioConsultaIngresosXSucursal.dtpFecha.Value, "MM/dd/yyyy") & "' " & "GROUP BY I.CodSucursal,I.Tipo) R RIGHT OUTER JOIN CatAlmacen C ON R.CodSucursal = C.CodAlmacen " & "WHERE C.TipoAlmacen = 'P' " & Sucursales & " GROUP BY C.CodAlmacen,C.DescAlmacen ORDER BY C.DescAlmacen"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With frmBancosProcesoDiarioConsultaIngresosXSucursal.flexDetalle
                    .Row = 1
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescAlmacen").Value))
                        .set_TextMatrix(.Row, 1, Format(RsGral.Fields("EfectivoPesos").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, 2, Format(RsGral.Fields("EfectivoDolares").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, 3, Format(RsGral.Fields("ImporteTarjetas").Value, "###,##0.00"))
                        RsGral.MoveNext()
                        If Not RsGral.EOF Then
                            If .Row = .Rows - 1 Then
                                .Rows = .Rows + 1
                            End If
                            .Row = .Row + 1
                        End If
                    Loop
                    .Col = 0
                    .Row = 1
                End With
            End If
            Me.Close()
            frmBancosProcesoDiarioConsultaIngresosXSucursal.ShowDialog()
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        PonerColor()
    End Sub

    Private Sub FlexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
            FlexDetalle_DblClick(flexDetalle, New System.EventArgs())
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioImportacionVouchers_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            ModEstandar.AvanzarTab(Me)
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioImportacionVouchers_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        Encabezado()
        MostrarSucursales()
        lblFecha.Text = Format(System.DateTime.FromOADate(Today.ToOADate - 1), "dd/MMM/yyyy")
        FueraClick = False
    End Sub
End Class
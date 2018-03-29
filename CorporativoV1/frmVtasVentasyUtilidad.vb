'**********************************************************************************************************************'
'*PROGRAMA: FORMULARIO DE VENTAS Y UTILIDAD JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports Microsoft.Office.Interop
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmVtasVentasyUtilidad
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE VENTAS Y UTILIDAD POR GRUPO                                                       *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 21 DE MAYO DE 2004                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents flexVentas As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents flexUtilidad As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox

    Dim mblnSalir As Boolean
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim RsAux As ADODB.Recordset
    Dim rsVentas As ADODB.Recordset
    Dim rsUtilidad As ADODB.Recordset
    Dim ObjExcel As Object
    Dim objLibro As Excel.Workbook
    Dim objHoja As Excel.Worksheet
    Dim Columna As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim Renglon As Integer
    'Dim cmd As ADODB.Command

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame4.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(6, 13)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(311, 63)
        Me.txtMensaje.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(168, 18)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(73, 17)
        Me.optDolares.TabIndex = 3
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.optDolares, "Muestra los Importes en Dólares")
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(42, 18)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(73, 17)
        Me.optPesos.TabIndex = 2
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optPesos, "Muestra los Importes en Pesos")
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtMensaje)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(6, 141)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(338, 90)
        Me.Frame4.TabIndex = 9
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Texto Adicional"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optDolares)
        Me.Frame1.Controls.Add(Me.optPesos)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(6, 81)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(265, 46)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Moneda"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.dtpFechaInicial)
        Me.Frame2.Controls.Add(Me.dtpFechaFinal)
        Me.Frame2.Controls.Add(Me.Label3)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(6, 6)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(338, 62)
        Me.Frame2.TabIndex = 5
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(64, 28)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaInicial.TabIndex = 0
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(222, 28)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaFinal.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(180, 33)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(44, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Hasta"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(18, 34)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Desde"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(121, 251)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 79
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(6, 251)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 78
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(236, 252)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 77
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasVentasyUtilidad
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(355, 310)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(341, 192)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVentasyUtilidad"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Ventas y Utilidad por Grupo"
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub



    Sub CalculaTotalesVentas()
        Dim I As Integer
        Dim TotalVentas As Decimal
        Dim TotalPiezas As Decimal
        If rsVentas.RecordCount > 0 Then
            rsVentas.MoveFirst()
        End If
        flexVentas.Clear()
        flexVentas.Rows = 2
        flexVentas.set_Cols(0, rsVentas.Fields.Count - 1)
        flexVentas.Col = 0
        flexVentas.Row = 1
        TotalVentas = 0
        TotalPiezas = 0
        Do While Not rsVentas.EOF
            For I = 1 To rsVentas.Fields.Count - 1 Step 2
                flexVentas.set_TextMatrix(flexVentas.Row, I - 1, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, I - 1))) + System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO))
                flexVentas.set_TextMatrix(flexVentas.Row, I, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, I))) + rsVentas.Fields(I + 1).Value)
                TotalVentas = TotalVentas + System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                TotalPiezas = TotalPiezas + rsVentas.Fields(I + 1).Value
                '           If flexVentas.Col + 1 < flexVentas.COLS - 1 Then
                '               flexVentas.Col = flexVentas.Col + 2
                '           End If
            Next
            rsVentas.MoveNext()
            If Not rsVentas.EOF Then
                flexVentas.Col = 0
            End If
        Loop
        flexVentas.set_Cols(0, flexVentas.get_Cols() + 2)
        flexVentas.Col = rsVentas.Fields.Count - 1
        flexVentas.set_TextMatrix(flexVentas.Row, flexVentas.Col, TotalVentas)
        flexVentas.Col = flexVentas.Col + 1
        flexVentas.set_TextMatrix(flexVentas.Row, flexVentas.Col, TotalPiezas)
        If rsVentas.RecordCount > 0 Then
            rsVentas.MoveFirst()
        End If
        flexVentas.Row = 1
        flexVentas.Col = 0
    End Sub

    Sub CalculaTotalesUtilidad()
        Dim I As Integer
        Dim TotalUtilidad As Decimal
        If rsUtilidad.RecordCount > 0 Then
            rsUtilidad.MoveFirst()
        End If
        flexUtilidad.Clear()
        flexUtilidad.Rows = 2
        flexUtilidad.set_Cols(0, rsUtilidad.Fields.Count - 1)
        flexUtilidad.Col = 0
        flexUtilidad.Row = 1
        Do While Not rsUtilidad.EOF
            For I = 1 To rsUtilidad.Fields.Count - 1
                flexUtilidad.set_TextMatrix(flexUtilidad.Row, I - 1, CDec(Numerico(flexUtilidad.get_TextMatrix(flexUtilidad.Row, I - 1))) + System.Math.Round(rsUtilidad.Fields(I).Value, C_REDONDEO))
                TotalUtilidad = TotalUtilidad + System.Math.Round(rsUtilidad.Fields(I).Value, C_REDONDEO)
            Next
            rsUtilidad.MoveNext()
            If Not rsUtilidad.EOF Then
                flexUtilidad.Col = 0
            End If
        Loop
        flexUtilidad.set_Cols(0, flexUtilidad.get_Cols() + 1)
        flexUtilidad.Col = rsUtilidad.Fields.Count - 1
        flexUtilidad.set_TextMatrix(flexUtilidad.Row, flexUtilidad.Col, TotalUtilidad)
        If rsUtilidad.RecordCount > 0 Then
            rsUtilidad.MoveFirst()
        End If
        flexUtilidad.Row = 1
        flexUtilidad.Col = 0
    End Sub

    Sub CierraInstanciasdeExcel(ByRef Tipo As Integer)
        If Tipo = 1 Then objLibro.Close()
        If ObjExcel Is Nothing Then ObjExcel = Nothing
        If objLibro Is Nothing Then objLibro = Nothing
        If objHoja Is Nothing Then objHoja = Nothing
    End Sub

    Function DevuelveQueryUtilidad() As String
        Dim Sql As String
        Sql = "Select Case When Tipo = 'R' Then 4 Else CodGrupo End as Grupo"
        gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                If optPesos.Checked = True Then
                    Sql = Sql & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then ((Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End - (CostoVenta*(Cantidad-CantidadDev)))) * TipoCambio Else 0 End),1) as Utilidad" & RsAux.Fields("CodAlmacen").Value & ""
                ElseIf optDolares.Checked = True Then
                    Sql = Sql & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then (Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End - (CostoVenta*(Cantidad-CantidadDev))) Else 0 End),2) as Utilidad" & RsAux.Fields("CodAlmacen").Value & ""
                End If
                RsAux.MoveNext()
            Loop
        End If
        Sql = Sql & " FROM    VENTAS_SALIDAMCIA('" & Format(dtpFechaInicial.Value, C_FORMATFECHAGUARDAR) & "', '" & Format(dtpFechaFinal.Value, C_FORMATFECHAGUARDAR) & "') Vta " & "Group   By Case When Tipo = 'R' Then 4 Else CodGrupo End " & "Order   By Case When Tipo = 'R' Then 4 Else CodGrupo End"
        DevuelveQueryUtilidad = Sql
    End Function

    Function DevuelveQueryVentas() As String
        Dim Sql As String
        Sql = "Select Case When Tipo = 'R' Then 4 Else CodGrupo End as Grupo"
        gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                If optPesos.Checked = True Then
                    Sql = Sql & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then (Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End) * TipoCambio Else 0 End),1) as Vta" & RsAux.Fields("CodAlmacen").Value & ""
                ElseIf optDolares.Checked = True Then
                    Sql = Sql & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End Else 0 End),2) as Vta" & RsAux.Fields("CodAlmacen").Value & ""
                End If
                Sql = Sql & ",Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then 1 Else (Cantidad-CantidadDev) End Else 0 End) as Piezas" & RsAux.Fields("CodAlmacen").Value & ""
                RsAux.MoveNext()
            Loop
        End If
        Sql = Sql & " FROM VENTAS_SALIDAMCIA('" & Format(dtpFechaInicial.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpFechaFinal.Value, C_FORMATFECHAGUARDAR) & "') Vta " & "Group By Case When Tipo = 'R' Then 4 Else CodGrupo End " & "Order By Case When Tipo = 'R' Then 4 Else CodGrupo End"
        DevuelveQueryVentas = Sql
    End Function

    Sub Encabezado()
        On Error GoTo Err_Renamed
        With objHoja
            .Range("C1").FormulaR1C1 = Trim(gstrCorpoNOMBREEMPRESA)
            .Range("C1:G1").Select()
            .Range("C1:G1").MergeCells = True
            .Range("C1:G1").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C1:G1").Font
                .Bold = True
                .Size = 12
                .Name = "Arial"
            End With
            .Range("C2").FormulaR1C1 = "Ventas y Utilidad por Grupo"
            .Range("C2:G2").Select()
            .Range("C2:G2").MergeCells = True
            .Range("C2:G2").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C2:G2").Font
                .Bold = False
                .Size = 11
                .Name = "Arial"
            End With
            .Range("C3").FormulaR1C1 = "Desde el " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & " Hasta el " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
            .Range("C3:G3").Select()
            .Range("C3:G3").MergeCells = True
            .Range("C3:G3").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C3:G3").Font
                .Bold = False
                .Size = 10
                .Name = "Arial"
            End With
            .Range("A4").FormulaR1C1 = "Fecha: " & Format(Today, "dd/mmm/yyyy")
            .Range("A4:B4").Select()
            .Range("A4:B4").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A4:B4").Font
                .Bold = False
                .Size = 9
                .Name = "Arial"
            End With
            .Range("A5").FormulaR1C1 = "Mensaje: "
            .Range("A5").Select()
            .Range("A5").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A5").Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            If Trim(txtMensaje.Text) <> "" Then
                .Range("B5").FormulaR1C1 = Trim(QuitaEnter(txtMensaje.Text))
                .Range("B5:J5").Select()
                .Range("B5:J5").MergeCells = True
                .Range("B5:J5").HorizontalAlignment = Excel.Constants.xlLeft
                With .Range("B5:J5").Font
                    .Bold = False
                    .Size = 9
                    .Name = "Arial"
                End With
            End If
            .Range("A6:C6").Select()
            .Range("A6:C6").MergeCells = True
            If optPesos.Checked = True Then
                .Range("A6:C6").FormulaR1C1 = "**Los importes estan expresados en pesos"
            ElseIf optDolares.Checked = True Then
                .Range("A6:C6").FormulaR1C1 = "**Los importes estan expresados en dólares"
            End If
            .Range("A6:C6").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A6:C6").Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range("A8")._Default = "VENTAS"
            .Range("A8").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A8").Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With

        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub EncabezadoVentas()
        On Error GoTo Err_Renamed
        With objHoja
            Columna = 1
            Renglon = 10
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Grupos"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Columna = Columna + 1
            gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                Do While Not RsAux.EOF
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = UCase((RsAux.Fields("DescAlmacen").Value)) & LCase(Mid(RsAux.Fields("DescAlmacen").Value, 2, 39))
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).WrapText = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlCenter
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).VerticalAlignment = Excel.Constants.xlTop
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        .Weight = Excel.XlBorderWeight.xlMedium
                        .ColorIndex = Excel.Constants.xlAutomatic
                    End With
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                        .Bold = True
                        .Size = 9
                        .Name = "Arial"
                    End With
                    Columna = Columna + 2
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Piezas"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        .Weight = Excel.XlBorderWeight.xlMedium
                        .ColorIndex = Excel.Constants.xlAutomatic
                    End With
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Bold = True
                        .Size = 9
                        .Name = "Arial"
                    End With
                    RsAux.MoveNext()
                    Columna = Columna + 1
                Loop
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = "Global"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).VerticalAlignment = Excel.Constants.xlTop
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                Columna = Columna + 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Piezas"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "%"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
            End If
            Columna = 1
            Renglon = Renglon + 1
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub EncabezadoUtilidad()
        On Error GoTo Err_Renamed
        With objHoja
            Columna = 1
            Renglon = Renglon + 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "UTILIDAD"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Renglon = Renglon + 2
            Columna = 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Grupos"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Columna = Columna + 1
            gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                Do While Not RsAux.EOF
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = UCase((RsAux.Fields("DescAlmacen").Value)) & LCase(Mid(RsAux.Fields("DescAlmacen").Value, 2, 39))
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).WrapText = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlCenter
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).VerticalAlignment = Excel.Constants.xlTop
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        If Columna > 2 Then
                            .Weight = Excel.XlBorderWeight.xlMedium
                            .ColorIndex = Excel.Constants.xlAutomatic
                        End If
                    End With
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        .Weight = Excel.XlBorderWeight.xlMedium
                        .ColorIndex = Excel.Constants.xlAutomatic
                    End With
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        .Weight = Excel.XlBorderWeight.xlMedium
                        .ColorIndex = Excel.Constants.xlAutomatic
                    End With
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                        .Bold = True
                        .Size = 9
                        .Name = "Arial"
                    End With
                    Columna = Columna + 2
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = "Piezas"
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).WrapText = True
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlCenter
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).VerticalAlignment = xlTop
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Interior.ColorIndex = 15
                    '                With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                    '                    .Bold = True
                    '                    .Size = 9
                    '                    .Name = "Arial"
                    '                End With
                    RsAux.MoveNext()
                    Columna = Columna + 1
                Loop
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = "Global"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).VerticalAlignment = Excel.Constants.xlTop
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                '            Columna = Columna + 2
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = "Piezas"
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).WrapText = True
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlCenter
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).VerticalAlignment = xlTop
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeTop).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Interior.ColorIndex = 15
                '            With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                '                .Bold = True
                '                .Size = 9
                '                .Name = "Arial"
                '            End With
                '            Columna = Columna + 1
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = "%"
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).WrapText = True
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlCenter
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).VerticalAlignment = xlTop
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeTop).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Interior.ColorIndex = 15
                '            With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                '                .Bold = True
                '                .Size = 9
                '                .Name = "Arial"
                '            End With
            End If
            Columna = 1
            Renglon = Renglon + 1
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub EnviaExcel()
        Dim Archivo As String
        On Error GoTo Err_Renamed
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Archivo = "VU" & CStr(Format(Month(Today), "00")) & CStr(Format((Today) + "00")) & (CStr(Format(Year(Today), "00"))) & ".xls"
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
            MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
        End If
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, FileAttribute.Archive) <> "" Then
            Kill(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo)
        End If
        ObjExcel = CreateObject("Excel.Application")
        objLibro = ObjExcel.Workbooks.Add
        objHoja = objLibro.ActiveSheet
        ObjExcel.Visible = False
        objLibro.Sheets(1).Select()
        objHoja = objLibro.ActiveSheet
        objLibro.ActiveSheet.Name = "Vtas. y Util. por Grupo"
        Encabezado()
        EncabezadoVentas()
        LlenaDatosVentas()
        EncabezadoUtilidad()
        LlenaDatosUtilidad()
        'LlenaDatos
        objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        System.Windows.Forms.Application.DoEvents()
        Select Case MsgBox("Se ha creado el archivo " & Archivo & " ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                ObjExcel.Visible = True
                ObjExcel = Nothing
                objLibro = Nothing
                objHoja = Nothing
            Case MsgBoxResult.No Or MsgBoxResult.Cancel
                CierraInstanciasdeExcel(1)
        End Select

Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar un nuevo archivo hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            CierraInstanciasdeExcel(2)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Imprime()
        On Error GoTo Err_Renamed
        Dim sqlVentas As Object
        Dim sqlUtilidad As String

        If Not ValidaDatos() Then Exit Sub
        sqlVentas = DevuelveQueryVentas()
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, sqlVentas))
        rsVentas = Cmd.Execute
        sqlUtilidad = DevuelveQueryUtilidad()
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, sqlUtilidad))
        rsUtilidad = Cmd.Execute
        If rsVentas.RecordCount = 0 And rsUtilidad.RecordCount = 0 Then
            MsgBox("No existen datos a mostrar en este periodo, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Else
            CalculaTotalesVentas()
            CalculaTotalesUtilidad()
            EnviaExcel()
        End If
        Cmd.CommandTimeout = 90

Err_Renamed:
        If Err.Number <> 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub Limpiar()
        Nuevo()
        dtpFechaInicial.Focus()
    End Sub

    Sub LlenaDatosVentas()
        Dim Total As Decimal
        Dim TotPiezas As Integer
        Dim Porcentaje As Decimal
        Dim Grupo As String
        Dim I As Integer
        flexVentas.Col = 0
        flexVentas.Row = 1
        On Error GoTo Err_Renamed
        With objHoja
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Do While Not rsVentas.EOF
                Columna = 1
                If rsVentas.Fields("Grupo").Value = 1 Then
                    Grupo = "Joyeria"
                ElseIf rsVentas.Fields("Grupo").Value = 2 Then
                    Grupo = "Relojeria"
                ElseIf rsVentas.Fields("Grupo").Value = 3 Then
                    Grupo = "Varios"
                ElseIf rsVentas.Fields("Grupo").Value = 4 Then
                    Grupo = "Reparaciones"
                End If
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Grupo
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 12
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                Total = 0
                TotPiezas = 0
                For I = 1 To rsVentas.Fields.Count - 1 Step 2
                    Total = Total + System.Math.Round(rsVentas.Fields(I).Value, 0)
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 13
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    If CDec(Numerico(flexVentas.get_TextMatrix(1, I - 1))) <> 0 Then
                        Porcentaje = System.Math.Round((System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO) / CDec(Numerico(flexVentas.get_TextMatrix(1, I - 1)))) * 100, 2)
                    Else
                        Porcentaje = 0
                    End If
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).FormulaR1C1 = VB6.Format(Porcentaje, "###,##0.00") & "%"
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).ColumnWidth = 5.86
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = Columna + 2
                    TotPiezas = TotPiezas + rsVentas.Fields(I + 1).Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = rsVentas.Fields(I + 1).Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 6.43
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = Columna + 1
                Next
                I = rsVentas.Fields.Count - 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Total
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 13
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                If CDec(Numerico(flexVentas.get_TextMatrix(1, I))) <> 0 Then
                    Porcentaje = System.Math.Round((Total / CDec(Numerico(flexVentas.get_TextMatrix(1, I)))) * 100, 2)
                Else
                    Porcentaje = 0
                End If
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).FormulaR1C1 = VB6.Format(Porcentaje, "###,##0.00") & "%"
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).ColumnWidth = 5.86
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = TotPiezas
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 6.43
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                If CDec(Numerico(flexVentas.get_TextMatrix(1, I + 1))) <> 0 Then
                    Porcentaje = System.Math.Round((TotPiezas / CDec(Numerico(flexVentas.get_TextMatrix(1, I + 1)))) * 100, 2)
                Else
                    Porcentaje = 0
                End If
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).FormulaR1C1 = VB6.Format(Porcentaje, "###,##0.00") & "%"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 5.86
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                rsVentas.MoveNext()
                Renglon = Renglon + 1
            Loop
            Columna = 1
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            I = 0
            Do While I < flexVentas.get_Cols() - 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = flexVentas.get_TextMatrix(1, I)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = flexVentas.get_TextMatrix(1, I + 1)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                I = I + 2
            Loop
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Size = 8
                .Name = "Arial"
            End With
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub LlenaDatosUtilidad()
        Dim Total As Decimal
        Dim Porcentaje As Decimal
        Dim Grupo As String
        Dim I As Integer
        flexUtilidad.Col = 0
        flexUtilidad.Row = 1
        On Error GoTo Err_Renamed
        With objHoja
            If rsUtilidad.RecordCount > 0 Then
                rsUtilidad.MoveFirst()
            End If
            Do While Not rsUtilidad.EOF
                Columna = 1
                If rsUtilidad.Fields("Grupo").Value = 1 Then
                    Grupo = "Joyeria"
                ElseIf rsUtilidad.Fields("Grupo").Value = 2 Then
                    Grupo = "Relojeria"
                ElseIf rsUtilidad.Fields("Grupo").Value = 3 Then
                    Grupo = "Varios"
                ElseIf rsUtilidad.Fields("Grupo").Value = 4 Then
                    Grupo = "Reparaciones"
                End If
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Grupo
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 12
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                Total = 0
                For I = 1 To rsUtilidad.Fields.Count - 1
                    Total = Total + rsUtilidad.Fields(I).Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = System.Math.Round(rsUtilidad.Fields(I).Value, C_REDONDEO)
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 13
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        If Columna > 2 Then
                            .Weight = Excel.XlBorderWeight.xlMedium
                            .ColorIndex = Excel.Constants.xlAutomatic
                        End If
                    End With
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    If CDec(Numerico(flexUtilidad.get_TextMatrix(1, I - 1))) <> 0 Then
                        Porcentaje = System.Math.Round((System.Math.Round(rsUtilidad.Fields(I).Value, C_REDONDEO) / CDec(Numerico(flexUtilidad.get_TextMatrix(1, I - 1)))) * 100, 2)
                    Else
                        Porcentaje = 0
                    End If
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).FormulaR1C1 = VB6.Format(Porcentaje, "###,##0.00") & "%"
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).ColumnWidth = 5.86
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        .Weight = Excel.XlBorderWeight.xlMedium
                        .ColorIndex = Excel.Constants.xlAutomatic
                    End With
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = Columna + 2
                    '                TotPiezas = TotPiezas + rsVentas(i + 1)
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = rsVentas.Fields(i + 1)
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).ColumnWidth = 9
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                    '                .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    '                With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                    '                    .Size = 8
                    '                    .Name = "Arial"
                    '                End With
                    Columna = Columna + 1
                Next
                I = rsUtilidad.Fields.Count - 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Total
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 13
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                If CDec(Numerico(flexUtilidad.get_TextMatrix(1, I))) <> 0 Then
                    Porcentaje = System.Math.Round((Total / CDec(Numerico(flexUtilidad.get_TextMatrix(1, I)))) * 100, 2)
                Else
                    Porcentaje = 0
                End If
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).FormulaR1C1 = VB6.Format(Porcentaje, "###,##0.00") & "%"
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).ColumnWidth = 5.86
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 2
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = TotPiezas
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).ColumnWidth = 9
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                '            With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                '                .Size = 8
                '                .Name = "Arial"
                '            End With
                Columna = Columna + 1
                '            If CCur(Numerico(flexUtilidad.TextMatrix(1, i + 1))) <> 0 Then
                '                Porcentaje = Round((TotPiezas / CCur(Numerico(flexUtilidad.TextMatrix(1, i + 1)))) * 100, 2)
                '            Else
                '                Porcentaje = 0
                '            End If
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).FormulaR1C1 = Format(Porcentaje, "###,##0.00") & "%"
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).ColumnWidth = 8
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                '            With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                '                .Size = 8
                '                .Name = "Arial"
                '            End With
                rsUtilidad.MoveNext()
                Renglon = Renglon + 1
            Loop
            Columna = 1
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            I = 0
            Do While I <= flexUtilidad.get_Cols() - 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = flexUtilidad.get_TextMatrix(1, I)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    If Columna > 2 Then
                        .Weight = Excel.XlBorderWeight.xlMedium
                        .ColorIndex = Excel.Constants.xlAutomatic
                    End If
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = flexVentas.TextMatrix(1, i + 1)
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                '            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                '            With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                '                .Bold = True
                '                .Size = 8
                '                .Name = "Arial"
                '            End With
                Columna = Columna + 1
                I = I + 1
            Loop
            Renglon = Renglon + 2
            Columna = 2
            For I = 0 To flexUtilidad.get_Cols() - 2
                If I = 0 Then
                    If CDec(Numerico(flexVentas.get_TextMatrix(1, I))) = 0 Then
                        Porcentaje = 0
                    Else
                        Porcentaje = System.Math.Round((CDec(Numerico(flexUtilidad.get_TextMatrix(1, I))) / CDec(Numerico(flexVentas.get_TextMatrix(1, I)))) * 100, 2)
                    End If
                Else
                    If CDec(Numerico(flexVentas.get_TextMatrix(1, I * 2))) = 0 Then
                        Porcentaje = 0
                    Else
                        Porcentaje = System.Math.Round((CDec(Numerico(flexUtilidad.get_TextMatrix(1, I))) / CDec(Numerico(flexVentas.get_TextMatrix(1, I * 2)))) * 100, 2)
                    End If
                End If
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = VB6.Format(Porcentaje, "###,##0.00") & "%"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlMedium
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 3
            Next
            Renglon = Renglon + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "MARGEN BRUTO"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Renglon = Renglon + 1
            Porcentaje = System.Math.Round((CDec(Numerico(flexUtilidad.get_TextMatrix(1, flexUtilidad.get_Cols() - 1))) / CDec(Numerico(flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 2)))) * 100, 2)
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = VB6.Format(Porcentaje, "###,##0.00") & "%"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlMedium
                .ColorIndex = Excel.Constants.xlAutomatic
            End With
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With

            'Seguimos con las graficas
            'Grafica de Ventas Global por Grupo
            .Range("A1001").Select()
            .Range("A1001").FormulaR1C1 = "Joyeria"
            .Range("A1002").Select()
            .Range("A1002").FormulaR1C1 = "Relojeria"
            .Range("A1003").Select()
            .Range("A1003").FormulaR1C1 = "Varios"
            .Range("A1004").Select()
            .Range("A1004").FormulaR1C1 = "Reparaciones"
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Do While Not rsVentas.EOF
                Total = 0
                '''17MAY2005
                '''toma Suc 1-3-5-7  -->  deben ser todas
                '''For I = 1 To rsVentas.Fields.Count - 2 Step 2
                For I = 1 To rsVentas.Fields.Count - 2 Step 1
                    Total = Total + System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                Next
                .Range("B100" & CStr(rsVentas.Fields("Grupo").Value) & "").Select()
                .Range("B100" & CStr(rsVentas.Fields("Grupo").Value) & "").FormulaR1C1 = Total
                rsVentas.MoveNext()
            Loop
            .Range("B1005").Select()
            .Range("B1005").FormulaR1C1 = "SUM(R[-4]C:R[-1]C)"
            .Range("D30").Select()
            .Application.ActiveWindow.SmallScroll(Down:=8)
            '.Range("K30").Select
            '.Application.ActiveWindow.SmallScroll TOLEFT:=8
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("A1001:B1004"), PlotBy:=Excel.XlRowCol.xlColumns)
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            With .Application.ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Ventas Global por Grupo"
            End With
            .Application.ActiveSheet.Shapes("Chart 1").ScaleHeight(1.1, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 1").ScaleWidth(1.26, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            .Application.ActiveWindow.Visible = False
            .Range("A1001:B1005").Font.ColorIndex = 2

            'Grafica de Utilidad Global por Grupo
            .Range("C1001").Select()
            .Range("C1001").FormulaR1C1 = "Joyeria"
            .Range("C1002").Select()
            .Range("C1002").FormulaR1C1 = "Relojeria"
            .Range("C1003").Select()
            .Range("C1003").FormulaR1C1 = "Varios"
            .Range("C1004").Select()
            .Range("C1004").FormulaR1C1 = "Reparaciones"
            If rsUtilidad.RecordCount > 0 Then
                rsUtilidad.MoveFirst()
            End If
            Do While Not rsUtilidad.EOF
                Total = 0
                '''17MAY2005
                '''toma Suc 1-3-5-7  -->  deben ser todas
                '''For I = 1 To rsUtilidad.Fields.Count - 2 Step 2
                For I = 1 To rsUtilidad.Fields.Count - 2 Step 1
                    Total = Total + System.Math.Round(rsUtilidad.Fields(I).Value, C_REDONDEO)
                Next
                .Range("D100" & CStr(rsUtilidad.Fields("Grupo").Value) & "").Select()
                .Range("D100" & CStr(rsUtilidad.Fields("Grupo").Value) & "").FormulaR1C1 = Total
                rsUtilidad.MoveNext()
            Loop
            .Range("D1005").Select()
            .Range("D1005").FormulaR1C1 = "SUM(R[-4]C:R[-1]C)"
            .Range("P30").Select()
            .Application.ActiveWindow.SmallScroll(Down:=8)
            .Range("V30").Select()
            .Application.ActiveWindow.SmallScroll(ToRight:=2)
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("C1001:D1004"), PlotBy:=Excel.XlRowCol.xlColumns)
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            With .Application.ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Utilidad Global por Grupo"
            End With
            .Application.ActiveSheet.Shapes("Chart 2").ScaleHeight(1.1, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 2").ScaleWidth(1.26, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            .Application.ActiveWindow.Visible = False
            .Range("C1001:D1005").Font.ColorIndex = 2

            'Grafica de Venta Global por Articulos (Joyeria,Relojeria, Varios y Reparaciones)
            .Range("E1001").Select()
            .Range("E1001").FormulaR1C1 = "Joyeria"
            .Range("E1002").Select()
            .Range("E1002").FormulaR1C1 = "Relojeria"
            .Range("E1003").Select()
            .Range("E1003").FormulaR1C1 = "Varios"
            .Range("E1004").Select()
            .Range("E1004").FormulaR1C1 = "Reparaciones"
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Do While Not rsVentas.EOF
                Total = 0
                For I = 2 To rsVentas.Fields.Count - 1 Step 2
                    Total = Total + System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                Next
                .Range("F100" & CStr(rsVentas.Fields("Grupo").Value) & "").Select()
                .Range("F100" & CStr(rsVentas.Fields("Grupo").Value) & "").FormulaR1C1 = Total
                rsVentas.MoveNext()
            Loop
            .Range("F1005").Select()
            .Range("F1005").FormulaR1C1 = "SUM(R[-4]C:R[-1]C)"
            .Range("D55").Select()
            .Application.ActiveWindow.SmallScroll(Down:=8)
            '.Range("K55").Select
            '.Application.ActiveWindow.SmallScroll TOLEFT:=8
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("E1001:F1004"), PlotBy:=Excel.XlRowCol.xlColumns)
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            With .Application.ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Venta Global por Articulos (Joyeria,Relojeria,Varios y Reparaciones)"
            End With

            .Application.ActiveSheet.Shapes("Chart 3").ScaleHeight(1.1, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 3").ScaleWidth(1.26, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            .Application.ActiveWindow.Visible = False
            .Range("E1001:F1005").Font.ColorIndex = 2

            'Grafica de Ventas por Joyeria
            gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                Total = 0
                Do While Not RsAux.EOF
                    Total = Total + 1
                    .Range("I" & 1000 + Total).Select()
                    .Range("I" & 1000 + Total).FormulaR1C1 = RsAux.Fields("DescAlmacen")
                    RsAux.MoveNext()
                Loop
            End If
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Do While Not rsVentas.EOF
                If rsVentas.Fields("Grupo").Value = 1 Then
                    Total = 0
                    For I = 1 To rsVentas.Fields.Count - 2 Step 2
                        Total = Total + 1
                        .Range("J" & CStr(1000 + Total) & "").Select()
                        .Range("J" & CStr(1000 + Total) & "").FormulaR1C1 = System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                    Next
                    Exit Do
                End If
                rsVentas.MoveNext()
            Loop
            .Range("J" & 1000 + (Total + 1)).Select()
            .Range("J" & 1000 + (Total + 1)).FormulaR1C1 = "SUM(R[-" & Total & "]C:R[-1]C)"
            .Range("D80").Select()
            .Application.ActiveWindow.SmallScroll(Down:=8)
            '.Range("I80").Select
            '.Application.ActiveWindow.SmallScroll TOLEFT:=8
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("I1001:J" & (1000 + Total)), PlotBy:=Excel.XlRowCol.xlColumns)
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            With .Application.ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Venta de Joyeria por Sucursal"
            End With
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.ChartTitle.Select()
            .Application.Selection.AutoScaleFont = True
            With .Application.Selection.Font
                .Name = "Arial"
                .Size = 10
                .Bold = True
            End With
            .Application.Selection.Left = 50
            With .Application.ActiveChart.Legend.Font
                .Size = 10
                .Name = "Arial"
            End With
            .Application.ActiveChart.Legend.Height = 1000
            .Application.ActiveChart.Legend.Top = 1
            .Application.ActiveSheet.Shapes("Chart 4").ScaleHeight(1.35, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 4").ScaleWidth(1.76, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
            With .Application.ActiveChart.SeriesCollection(1).DataLabels.Font
                .Size = 8
                .Name = "Arial"
            End With
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            With .Application.Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .Underline = False
            End With
            .Application.Selection.Top = 1
            '.Application.Selection.Left = 1800
            .Application.Selection.Height = 300
            .Application.Selection.Width = 200
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            .Application.ActiveWindow.Visible = False
            .Range("I1001:J" & (1000 + (Total + 1))).Font.ColorIndex = 2

            'Grafica de Ventas por Relojeria
            Total = 0
            If RsAux.RecordCount > 0 Then
                RsAux.MoveFirst()
            End If
            Do While Not RsAux.EOF
                Total = Total + 1
                .Range("K" & 1000 + Total).Select()
                .Range("K" & 1000 + Total).FormulaR1C1 = RsAux.Fields("DescAlmacen")
                RsAux.MoveNext()
            Loop
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Do While Not rsVentas.EOF
                If rsVentas.Fields("Grupo").Value = 2 Then
                    Total = 0
                    For I = 1 To rsVentas.Fields.Count - 2 Step 2
                        Total = Total + 1
                        .Range("L" & CStr(1000 + Total) & "").Select()
                        .Range("L" & CStr(1000 + Total) & "").FormulaR1C1 = System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                    Next
                    Exit Do
                End If
                rsVentas.MoveNext()
            Loop
            .Range("L" & (1000 + (Total + 1))).Select()
            .Range("L" & (1000 + (Total + 1))).FormulaR1C1 = "SUM(R[-" & Total & "]C:R[-1]C)"
            .Range("D110").Select()
            .Application.ActiveWindow.SmallScroll(Down:=8)
            '.Range("I105").Select
            '.Application.ActiveWindow.SmallScroll TOLEFT:=8
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("K1001:L" & (1000 + Total)), PlotBy:=Excel.XlRowCol.xlColumns)
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            With .Application.ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Venta de Relojeria por Sucursal"
            End With
            .Application.ActiveSheet.ChartObjects("Chart 5").Activate()
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.AutoScaleFont = True
            With .Application.Selection.Font
                .Name = "Arial"
                .FontStyle = "Regular"
                .Size = 10
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ColorIndex = Excel.Constants.xlAutomatic
                .Background = Excel.Constants.xlAutomatic
            End With
            '.Application.Selection.Left = 650
            .Application.Selection.Top = 1
            .Application.Selection.Height = 350
            '.Application.Selection.Width = 120
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.ChartTitle.Select()
            .Application.Selection.AutoScaleFont = True
            With .Application.Selection.Font
                .Name = "Arial"
                .Size = 10
                .Bold = True
            End With

            .Application.Selection.Left = 50
            With .Application.ActiveChart.Legend.Font
                .Size = 8
                .Name = "Arial"
            End With
            .Application.ActiveSheet.ChartObjects("Chart 5").Activate()
            .Application.ActiveChart.ChartArea.Select()

            .Application.ActiveSheet.Shapes("Chart 5").ScaleHeight(1.35, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 5").ScaleWidth(1.76, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
            With .Application.ActiveChart.SeriesCollection(1).DataLabels.Font
                .Size = 8
                .Name = "Arial"
            End With
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            With .Application.Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .Underline = False
            End With
            .Application.Selection.Top = 1
            '.Application.Selection.Left = 1300
            .Application.Selection.Height = 300
            .Application.Selection.Width = 200
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            .Application.ActiveWindow.Visible = False
            .Range("K1001:L" & (1000 + (Total + 1))).Font.ColorIndex = 2

            'Grafica de Ventas por Varios
            Total = 0
            If RsAux.RecordCount > 0 Then
                RsAux.MoveFirst()
            End If
            Do While Not RsAux.EOF
                Total = Total + 1
                .Range("M" & 1000 + Total).Select()
                .Range("M" & 1000 + Total).FormulaR1C1 = RsAux.Fields("DescAlmacen")
                RsAux.MoveNext()
            Loop
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Do While Not rsVentas.EOF
                If rsVentas.Fields("Grupo").Value = 3 Then
                    Total = 0
                    For I = 1 To rsVentas.Fields.Count - 2 Step 2
                        Total = Total + 1
                        .Range("N" & CStr(1000 + Total) & "").Select()
                        .Range("N" & CStr(1000 + Total) & "").FormulaR1C1 = System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                    Next
                    Exit Do
                End If
                rsVentas.MoveNext()
            Loop
            .Range("N" & (1000 + (Total + 1))).Select()
            .Range("N" & (1000 + (Total + 1))).FormulaR1C1 = "SUM(R[-" & Total & "]C:R[-1]C)"
            .Range("D140").Select()
            .Application.ActiveWindow.SmallScroll(Down:=8)
            '.Range("I130").Select
            '.Application.ActiveWindow.SmallScroll TOLEFT:=8
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("M1001:N" & (1000 + Total)), PlotBy:=Excel.XlRowCol.xlColumns)
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            With .Application.ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Venta de Varios por Sucursal"
            End With
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.ActiveSheet.ChartObjects. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.ActiveSheet.ChartObjects("Chart 6").Activate()
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.Legend.Select()
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.AutoScaleFont. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.Selection.AutoScaleFont = True
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            With .Application.Selection.Font
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Name = "Arial"
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .FontStyle = "Regular"
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Size = 10
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Strikethrough = False
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Superscript = False
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Subscript = False
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .OutlineFont = False
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Shadow = False
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .ColorIndex = Excel.Constants.xlAutomatic
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Background = Excel.Constants.xlAutomatic
            End With
            '.Application.Selection.Left = 650
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Top. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.Selection.Top = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Height. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.Selection.Height = 350
            '.Application.Selection.Width = 120
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.ChartTitle.Select()
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.AutoScaleFont. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.Selection.AutoScaleFont = True
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            With .Application.Selection.Font
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Name = "Arial"
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Size = 10
                'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Bold = True
            End With
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.Selection.Left. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.Selection.Left = 50
            With .Application.ActiveChart.Legend.Font
                .Size = 8
                .Name = "Arial"
            End With
            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.ActiveSheet.ChartObjects. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.ActiveSheet.ChartObjects("Chart 6").Activate()
            .Application.ActiveChart.ChartArea.Select()

            'UPGRADE_WARNING: Couldn't resolve default property of object objHoja.Application.ActiveSheet.Shapes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Application.ActiveSheet.Shapes("Chart 6").ScaleHeight(1.35, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 6").ScaleWidth(1.76, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
            With .Application.ActiveChart.SeriesCollection(1).DataLabels.Font
                .Size = 8
                .Name = "Arial"
            End With
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            With .Application.Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .Underline = False
            End With
            .Application.Selection.Top = 1
            '.Application.Selection.Left = 1300
            .Application.Selection.Height = 300
            .Application.Selection.Width = 200
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            .Application.ActiveWindow.Visible = False
            .Range("M1001:N" & (1000 + (Total + 1))).Font.ColorIndex = 2

            'Grafica de Ventas por Reparaciones
            Total = 0
            If RsAux.RecordCount > 0 Then
                RsAux.MoveFirst()
            End If
            Do While Not RsAux.EOF
                Total = Total + 1
                .Range("O" & 1000 + Total).Select()
                .Range("O" & 1000 + Total).FormulaR1C1 = RsAux.Fields("DescAlmacen")
                RsAux.MoveNext()
            Loop
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Do While Not rsVentas.EOF
                If rsVentas.Fields("Grupo").Value = 4 Then
                    Total = 0
                    For I = 1 To rsVentas.Fields.Count - 2 Step 2
                        Total = Total + 1
                        .Range("P" & CStr(1000 + Total) & "").Select()
                        .Range("P" & CStr(1000 + Total) & "").FormulaR1C1 = System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                    Next
                    Exit Do
                End If
                rsVentas.MoveNext()
            Loop
            .Range("P" & (1000 + (Total + 1))).Select()
            .Range("P" & (1000 + (Total + 1))).FormulaR1C1 = "SUM(R[-" & Total & "]C:R[-1]C)"
            .Range("D170").Select()
            .Application.ActiveWindow.SmallScroll(Down:=8)
            '.Range("I155").Select
            '.Application.ActiveWindow.SmallScroll TOLEFT:=8
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("O1001:P" & (1000 + Total)), PlotBy:=Excel.XlRowCol.xlColumns)
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            With .Application.ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = "Venta de Reparaciones por Sucursal"
            End With
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.ChartTitle.Select()
            .Application.Selection.AutoScaleFont = True
            With .Application.Selection.Font
                .Name = "Arial"
                .Size = 10
                .Bold = True
            End With
            .Application.Selection.Left = 50
            .Application.ActiveSheet.ChartObjects("Chart 7").Activate()
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveSheet.Shapes("Chart 7").ScaleHeight(1.35, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 7").ScaleWidth(1.76, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
            With .Application.ActiveChart.SeriesCollection(1).DataLabels.Font
                .Size = 8
                .Name = "Arial"
            End With
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            With .Application.Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .Underline = False
            End With
            .Application.Selection.Top = 1
            '.Application.Selection.Left = 1300
            .Application.Selection.Height = 300
            .Application.Selection.Width = 200
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Vtas. y Util. por Grupo")
            .Application.ActiveWindow.Visible = False
            .Range("O1001:P" & (1000 + (Total + 1))).Font.ColorIndex = 2
            .Application.ActiveWindow.Zoom = 85
            .Range("A1").Select()
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Nuevo()
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        optPesos.Checked = True
        optDolares.Checked = False
        txtMensaje.Text = ""
        mblnSalir = False
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        Do While (sglTiempoCambio) <= 2.1
        Loop
        System.Windows.Forms.Application.DoEvents()
        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Function
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Function
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaFinal.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

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

    Private Sub frmVtasVentasyUtilidad_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVentasyUtilidad_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVentasyUtilidad_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmVtasVentasyUtilidad_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVentasyUtilidad_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Nuevo()
    End Sub

    Private Sub frmVtasVentasyUtilidad_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        If mblnSalir Then
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

    Private Sub frmVtasVentasyUtilidad_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub optPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class
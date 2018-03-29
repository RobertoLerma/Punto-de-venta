Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports System.Data.SqlClient
Public Class frmBancosProcesoMensualConsultaOrigenAplicRec
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CONSULTA DE ORIGEN Y APLICACIÓN DE RECURSOS                                                  *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 08 DE AGOSTO DE 2003                                                                 *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtTotalDolares As System.Windows.Forms.TextBox
    Public WithEvents txtTotalPesos As System.Windows.Forms.TextBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents txtRubro As System.Windows.Forms.TextBox
    Public WithEvents chkTodoslosRubros As System.Windows.Forms.CheckBox
    Public WithEvents dbcRubro As System.Windows.Forms.ComboBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcAgrupador As System.Windows.Forms.ComboBox
    Public WithEvents txtAgrupador As System.Windows.Forms.TextBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents lblSeleccionado As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents lblModificados As System.Windows.Forms.Label

    'Variables
    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim FueraChange As Boolean
    Dim PierdeFoco As Boolean
    Dim intCodAgrupador As Integer
    Dim intCodRubro As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnBuscar As Button
    Public tecla As Integer
    Public strControlActual As String 'Nombre del control actual

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTAGRUPADOR"
                strCaptionForm = "Consulta de Agrupadores de Origen y Aplicación de Recursos"
                gStrSql = "SELECT RIGHT('0000'+LTRIM(CodOrigenAplicR),4) AS CODIGO, DescOrigenAplicR AS DESCRIPCION FROM CatOrigenAplicRecursos ORDER BY CodOrigenAplicR"
            Case "TXTRUBRO"
                If CDbl(Numerico(txtAgrupador.Text)) = 0 Then
                    MsgBox("Proporcione un Agrupador, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                End If
                strCaptionForm = "Consulta de Rubros de Origen y Aplicación de Recursos"
                gStrSql = "SELECT RIGHT('000000'+LTRIM(CodRubro),6) AS CODIGO, DescRubro AS DESCRIPCION FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR = " & txtAgrupador.Text & " ORDER BY CodRubro"
            Case Else
                strControlActual = ""
        End Select

        'If strControlActual = "" Then Exit Sub

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
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTAGRUPADOR"
                    'ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTRUBRO"
                    'ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaAgrupador()
        On Error GoTo Merr
        gStrSql = "Select CodOrigenAplicR,DescOrigenAplicR FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR = " & txtAgrupador.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtAgrupador.Text = VB6.Format(txtAgrupador.Text, "0000")
            dbcAgrupador.Text = Trim(RsGral.Fields("DescOrigenAplicR").Value)
        Else
            MsgBox("Codigo de Agrupador no Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtAgrupador.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaRubro()
        On Error GoTo Merr
        gStrSql = "Select CodRubro,DescRubro FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR = " & txtAgrupador.Text & " AND CodRubro = " & txtRubro.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtRubro.Text = VB6.Format(txtRubro.Text, "000000")
            dbcRubro.Text = Trim(RsGral.Fields("DescRubro").Value)
        Else
            MsgBox("Codigo de Rubro no Existe para este Agrupador, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtRubro.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub ConfiguraGrid()
        Dim I As Integer
        With flexDetalle
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1200)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Fecha"
            .Col = 1
            .set_ColWidth(1, 0, 1500)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Folio"
            .Col = 2
            .set_ColWidth(2, 0, 2760)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Concepto"
            .Col = 3
            .set_ColWidth(3, 0, 1700)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Pesos"
            .Col = 4
            .set_ColWidth(4, 0, 1700)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Dolares"
            .set_Cols(0, 11)
            .set_ColWidth(5, 0, 0)
            .set_ColWidth(6, 0, 0)
            .set_ColWidth(7, 0, 0)
            .set_ColWidth(8, 0, 0)
            .set_ColWidth(9, 0, 0)
            .set_ColWidth(10, 0, 0)
            For I = .FixedRows To .Rows - 1
                .set_TextMatrix(I, 0, VB6.Format(.get_TextMatrix(I, 0), "dd/mmm/yyyy"))
                .set_TextMatrix(I, 3, VB6.Format(.get_TextMatrix(I, 3), "###,##0.00"))
                .set_TextMatrix(I, 4, VB6.Format(.get_TextMatrix(I, 4), "###,##0.00"))
            Next
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub Encabezado()
        With flexDetalle
            .set_Cols(0, 11)
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1200)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Fecha"
            .Col = 1
            .set_ColWidth(1, 0, 1500)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Folio"
            .Col = 2
            .set_ColWidth(2, 0, 2760)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Concepto"
            .Col = 3
            .set_ColWidth(3, 0, 1700)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Pesos"
            .Col = 4
            .set_ColWidth(4, 0, 1700)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Dolares"
            .set_ColWidth(5, 0, 0)
            .set_ColWidth(6, 0, 0)
            .set_ColWidth(7, 0, 0)
            .set_ColWidth(8, 0, 0)
            .set_ColWidth(9, 0, 0)
            .set_ColWidth(10, 0, 0)
            .Rows = 11
            .Col = 0
            .Row = 1
            txtTotalPesos.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Left) + .get_ColPos(3) + 30)
            txtTotalPesos.Width = VB6.TwipsToPixelsX(.get_ColWidth(3) - 15)
            txtTotalDolares.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Left) + .get_ColPos(4) + 30)
            txtTotalDolares.Width = VB6.TwipsToPixelsX(.get_ColWidth(4) - 15)
        End With
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
        FueraChange = False
        PierdeFoco = False
        intCodAgrupador = 0
        intCodRubro = 0
        tecla = 0
    End Sub

    Sub Limpiar()
        Nuevo()
        txtAgrupador.Focus()
    End Sub

    Sub Nuevo()
        txtAgrupador.Text = "0000"
        dbcAgrupador.Text = ""
        txtRubro.Text = "000000"
        dbcRubro.Text = ""
        txtTotalPesos.Text = ""
        txtTotalDolares.Text = ""
        'dbcRubro.RowSource = Nothing
        chkTodoslosRubros.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtRubro.Enabled = True
        dbcRubro.Enabled = True
        cmbMes.SelectedIndex = 0
        cmbAño.SelectedIndex = 0
        flexDetalle.Clear()
        Encabezado()
        InicializaVariables()
    End Sub

    Sub ObtenerEjercicios()
        On Error GoTo Merr
        gStrSql = "SELECT DISTINCT Ejercicio FROM EjercicioPeriodo"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Do While Not RsGral.EOF
                cmbAño.Items.Add(RsGral.Fields("Ejercicio").Value)
                RsGral.MoveNext()
            Loop
        Else
            cmbAño.Items.Add("")
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub VerMovimientos()
        On Error GoTo Merr
        Dim FechaInicial As String
        Dim FechaFinal As String
        If CDbl(Numerico(txtAgrupador.Text)) = 0 Then
            flexDetalle.Clear()
            Encabezado()
            Exit Sub
        End If
        If Trim(dbcAgrupador.Text) = "" Then
            flexDetalle.Clear()
            Encabezado()
            Exit Sub
        End If
        If chkTodoslosRubros.CheckState = 0 Then
            If CDbl(Numerico(txtRubro.Text)) = 0 Then
                flexDetalle.Clear()
                Encabezado()
                Exit Sub
            End If
            If Trim(dbcRubro.Text) = "" Then
                flexDetalle.Clear()
                Encabezado()
                Exit Sub
            End If
        End If
        If Trim(cmbAño.Text) = "" Then
            flexDetalle.Clear()
            Encabezado()
            Exit Sub
        End If
        ObtenerLimitedeFechas(CInt(VB.Left(Trim(cmbMes.Text), 2)), CInt(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
        If chkTodoslosRubros.CheckState = 1 Then
            '''gStrSql = "SELECT MB.FechaMovto, MB.FolioMovto, CR.DescRubro, CASE WHEN MB.Moneda = 'P' THEN MOA.Importe ELSE 0 END AS ImpPesos, CASE WHEN MB.Moneda = 'D' THEN MOA.Importe ELSE 0 END AS ImpDolares, MOA.CodOrigenAplicR, MOA.CodRubro " & _
            '"FROM MovimientosBancarios MB INNER JOIN MovimientosOrigenAplic MOA " & _
            '"ON MB.FolioMovto = MOA.FolioMovto INNER JOIN CatRubrosOrigenAplicRecursos CR ON MOA.CodRubro = CR.CodRubro " & _
            '"WHERE MOA.Estatus <> 'C' AND MB.FechaMovto BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND MOA.CodOrigenAplicR = " & txtAgrupador & " ORDER BY MB.FechaMovto"
            gStrSql = "SELECT MB.FechaMovto, MB.FolioMovto, COA.DescRubro, CASE WHEN MB.Moneda = 'P' THEN MOA.Importe ELSE 0 END AS ImpPesos, CASE WHEN MB.Moneda = 'D' THEN MOA.Importe ELSE 0 END AS ImpDolares, MOA.CodOrigenAplicR, MOA.CodRubro " & "FROM   MovimientosBancarios MB INNER JOIN MovimientosOrigenAplic MOA ON MB.FolioMovto = MOA.FolioMovto " & "Inner  Join ( Select   COA.CodOrigenAplicR, COA.DescOrigenAplicR, COA.Aplicacion, CR.CodRubro, CR.DescRubro " & "             From  CatOrigenAplicRecursos COA INNER JOIN CatRubrosOrigenAplicRecursos CR ON COA.CodOrigenAplicR = CR.CodOrigAplicR ) COA On MOA.CodOrigenAplicR = COA.CodOrigenAplicR And MOA.CodRubro = COA.CodRubro " & "WHERE  MB.FechaMovto BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' " & "AND    MOA.CodOrigenAplicR = " & CInt(Trim(txtAgrupador.Text)) & " " & "AND    MOA.Estatus <> 'C' " & "Order  By MOA.CodOrigenAplicR, MOA.CodRubro "
        ElseIf chkTodoslosRubros.CheckState = 0 Then
            '''duplica movtos
            '''gStrSql = "SELECT MB.FechaMovto, MB.FolioMovto, CR.DescRubro, CASE WHEN MB.Moneda = 'P' THEN MOA.Importe ELSE 0 END AS ImpPesos, CASE WHEN MB.Moneda = 'D' THEN MOA.Importe ELSE 0 END AS ImpDolares, MOA.CodOrigenAplicR, MOA.CodRubro " & _
            '"FROM MovimientosBancarios MB INNER JOIN MovimientosOrigenAplic MOA " & _
            '"ON MB.FolioMovto = MOA.FolioMovto INNER JOIN CatRubrosOrigenAplicRecursos CR ON MOA.CodRubro = CR.CodRubro " & _
            '"WHERE MOA.Estatus <> 'C' AND MB.FechaMovto BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND MOA.CodOrigenAplicR = " & txtAgrupador & " AND MOA.CodRubro = " & txtRubro & " ORDER BY MB.FechaMovto"
            '''Sep 09 2004 - modific para que no duplique folios
            gStrSql = "SELECT MB.FechaMovto, MB.FolioMovto, COA.DescRubro, CASE WHEN MB.Moneda = 'P' THEN MOA.Importe ELSE 0 END AS ImpPesos, CASE WHEN MB.Moneda = 'D' THEN MOA.Importe ELSE 0 END AS ImpDolares, MOA.CodOrigenAplicR, MOA.CodRubro " & "FROM   MovimientosBancarios MB INNER JOIN MovimientosOrigenAplic MOA ON MB.FolioMovto = MOA.FolioMovto " & "Inner  Join ( Select   COA.CodOrigenAplicR, COA.DescOrigenAplicR, COA.Aplicacion, CR.CodRubro, CR.DescRubro " & "             From  CatOrigenAplicRecursos COA INNER JOIN CatRubrosOrigenAplicRecursos CR ON COA.CodOrigenAplicR = CR.CodOrigAplicR ) COA On MOA.CodOrigenAplicR = COA.CodOrigenAplicR And MOA.CodRubro = COA.CodRubro " & "WHERE  MB.FechaMovto BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' " & "AND    MOA.CodOrigenAplicR = " & CInt(Numerico(Trim(txtAgrupador.Text))) & " " & "AND    MOA.CodRubro = " & CInt(Numerico(Trim(txtRubro.Text))) & " " & "AND    MOA.Estatus <> 'C' " & "Order  By MOA.CodOrigenAplicR, MOA.CodRubro "
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            flexDetalle.Recordset = RsGral
            ConfiguraGrid()
            If flexDetalle.Rows - 1 <= 10 Then
                flexDetalle.Rows = 11
            End If
        Else
            flexDetalle.Clear()
            Encabezado()
        End If
        'Mostrar los totales
        ActualizarTotales()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub chkTodoslosRubros_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodoslosRubros.CheckStateChanged
        If chkTodoslosRubros.CheckState = 1 Then
            txtRubro.Text = "000000"
            txtRubro.Enabled = False
            dbcRubro.Text = ""
            dbcRubro.Enabled = False
        ElseIf chkTodoslosRubros.CheckState = 0 Then
            txtRubro.Enabled = True
            dbcRubro.Enabled = True
        End If
        VerMovimientos()
    End Sub

    Private Sub chkTodoslosRubros_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodoslosRubros.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbAño_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.SelectedIndexChanged
        VerMovimientos()
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbMes_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.SelectedIndexChanged
        VerMovimientos()
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcAgrupador_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAgrupador.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcAgrupador.Name Then
        '    Exit Sub
        'End If
        dbcRubro.Text = ""
        txtRubro.Text = "000000"
        txtTotalDolares.Text = ""
        txtTotalPesos.Text = ""
        gStrSql = "SELECT CodOrigenAplicR, RTRIM(LTRIM(DescOrigenAplicR)) AS DescOrigenAplicR FROM CatOrigenAplicRecursos (Nolock) WHERE DescOrigenAplicR LIKE '" & Trim(dbcAgrupador.Text) & "%' ORDER BY CodOrigenAplicR "
        DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcAgrupador_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAgrupador.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcAgrupador.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodOrigenAplicR, RTRIM(LTRIM(DescOrigenAplicR)) AS DescOrigenAplicR FROM CatOrigenAplicRecursos (Nolock) ORDER BY CodOrigenAplicR "
        DCGotFocus(gStrSql, dbcAgrupador)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcAgrupador_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcAgrupador.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then txtAgrupador.Focus()
    End Sub

    Private Sub dbcAgrupador_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcAgrupador.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcAgrupador_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcAgrupador.KeyUp
        Dim Aux As String
        Aux = dbcAgrupador.Text
        'If dbcAgrupador.SelectedItem <> 0 Then
        dbcAgrupador_Leave(dbcAgrupador, New System.EventArgs())
        'End If
        dbcAgrupador.Text = Aux
        VerMovimientos()
    End Sub

    Private Sub dbcAgrupador_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAgrupador.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodOrigenAplicR, RTRIM(LTRIM(DescOrigenAplicR)) AS DescOrigenAplicR FROM CatOrigenAplicRecursos (Nolock) WHERE DescOrigenAplicR = '" & Trim(dbcAgrupador.Text) & "' ORDER BY CodOrigenAplicR "
        FueraChange = True
        DCLostFocus(dbcAgrupador, gStrSql, intCodAgrupador)
        txtAgrupador.Text = VB6.Format(intCodAgrupador, "0000")
        FueraChange = False
        VerMovimientos()
    End Sub

    Private Sub dbcAgrupador_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcAgrupador.MouseUp
        Dim Aux As String
        Aux = dbcAgrupador.Text
        'If dbcAgrupador.SelectedItem <> 0 Then 
        'dbcAgrupador_Leave(dbcAgrupador, New System.EventArgs())
        'End if
        dbcAgrupador.Text = Aux
        VerMovimientos()
    End Sub

    Private Sub dbcRubro_cursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRubro.CursorChanged
        If FueraChange = True Then Exit Sub
        If Trim(dbcRubro.Text) = "" Then txtRubro.Text = ""
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRubro.Name Then Exit Sub
        gStrSql = "SELECT CodRubro, RTRIM(LTRIM(DescRubro)) AS DescRubro FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR = " & CInt(Numerico(Trim(txtAgrupador.Text))) & " AND DescRubro LIKE '" & Trim(dbcRubro.Text) & "%' ORDER BY DescRubro "
        DCChange(gStrSql, tecla)
        VerMovimientos()

    End Sub

    Private Sub dbcRubro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRubro.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRubro.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodRubro, RTRIM(LTRIM(DescRubro)) AS DescRubro FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR = " & CInt(Numerico(Trim(txtAgrupador.Text))) & " ORDER BY DescRubro "
        DCGotFocus(gStrSql, dbcRubro)
        Pon_Tool()
        FueraChange = False

    End Sub

    Private Sub dbcRubro_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRubro.KeyDown
        'tecla = eventSender.KeyCode
        'If eventSender.KeyCode = System.Windows.Forms.Keys.Escape Then txtRubro.Focus()
    End Sub

    Private Sub dbcRubro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcRubro.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcRubro_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRubro.KeyUp
        Dim Aux As String
        Aux = dbcRubro.Text
        'If dbcRubro.SelectedItem <> 0 Then
        dbcRubro_Leave(dbcRubro, New System.EventArgs())
        'End If
        dbcRubro.Text = Aux
        VerMovimientos()
    End Sub

    Private Sub dbcRubro_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRubro.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodRubro, RTRIM(LTRIM(DescRubro)) AS DescRubro FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR = " & CInt(Numerico(Trim(txtAgrupador.Text))) & " AND DescRubro = '" & Trim(dbcRubro.Text) & "' ORDER BY CodRubro "
        FueraChange = True
        DCLostFocus(dbcRubro, gStrSql, intCodRubro)
        txtRubro.Text = VB6.Format(intCodRubro, "000000")
        FueraChange = False
        VerMovimientos()

    End Sub

    Private Sub dbcRubro_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcRubro.MouseUp
        Dim Aux As String
        Aux = dbcRubro.Text
        'If dbcRubro.SelectedItem <> 0 Then dbcRubro_Leave(dbcRubro, New System.EventArgs())
        dbcRubro.Text = Aux
        VerMovimientos()
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        Dim I As Integer
        With flexDetalle
            If Trim(.get_TextMatrix(.Row, 0)) <> "" And Trim(.get_TextMatrix(.Row, 1)) <> "" And Trim(.get_TextMatrix(.Row, 2)) <> "" And Trim(.get_TextMatrix(.Row, 3)) <> "" Then
                For I = 0 To 4
                    .Col = I
                    .CellBackColor = lblSeleccionado.BackColor
                Next

                'frmBancosProcesoDiarioOrigenyAplicacion.frmBancosProcesoDiarioOrigenyAplicacion_Load(New Object, New EventArgs)
                frmConsultaOrigenAplicacion.Tag = "frmConsultaOrigenAplicacion"
                frmConsultaOrigenAplicacion.lblFolio.Text = .get_TextMatrix(.Row, 1)
                frmConsultaOrigenAplicacion.Text = frmConsultaOrigenAplicacion.Text & " (Reclasificación de Origen y Aplicacion)"
                frmConsultaOrigenAplicacion.ShowDialog()
            End If
        End With
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
    End Sub

    Private Sub FlexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
            FlexDetalle_DblClick(flexDetalle, New System.EventArgs())
        End If
    End Sub

    Private Sub frmBancosProcesoMensualConsultaOrigenAplicRec_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoMensualConsultaOrigenAplicRec_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualConsultaOrigenAplicRec_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtAgrupador" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoMensualConsultaOrigenAplicRec_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoMensualConsultaOrigenAplicRec_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        frmConsultaOrigenAplicacion.InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ObtenerEjercicios()
        Nuevo()
    End Sub

    Private Sub frmBancosProcesoMensualConsultaOrigenAplicRec_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        '    '        If ChecaCambios Then
        '    '            Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '    '                Case vbYes
        '    '                    If Guardar = False Then
        '    '                        Cancel = 1
        '    '                    End If
        '    '                Case vbNo
        '    '                Case vbCancel
        '    '                    Cancel = 1
        '    '            End Select
        '    '        End If
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

    Private Sub frmBancosProcesoMensualConsultaOrigenAplicRec_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec.Close()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtAgrupador_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAgrupador.TextChanged
        If FueraChange Then Exit Sub
        dbcAgrupador.Text = ""
        txtRubro.Text = "000000"
        dbcRubro.Text = ""
        txtTotalDolares.Text = ""
        txtTotalPesos.Text = ""
        VerMovimientos()
    End Sub

    Private Sub txtAgrupador_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAgrupador.Enter
        strControlActual = UCase("txtAgrupador")
        SelTextoTxt(txtAgrupador)
        Pon_Tool()
    End Sub

    Private Sub txtAgrupador_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAgrupador.KeyPress
        'Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'ModEstandar.gp_CampoNumerico(KeyAscii)
        'eventArgs.KeyChar = Chr(KeyAscii)
        'If KeyAscii = 0 Then
        '    eventArgs.Handled = True
        'End If
    End Sub

    Private Sub txtAgrupador_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAgrupador.Leave
        If CDbl(Numerico(txtAgrupador.Text)) = 0 Then
            txtAgrupador.Text = "0000"
        Else
            BuscaAgrupador()
        End If
    End Sub

    Private Sub txtRubro_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRubro.TextChanged
        If FueraChange Then Exit Sub
        dbcRubro.Text = ""
        VerMovimientos()
    End Sub

    Private Sub txtRubro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRubro.Enter
        strControlActual = UCase("txtRubro")
        SelTextoTxt(txtRubro)
        Pon_Tool()
    End Sub

    Private Sub txtRubro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRubro.KeyPress
        'Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'ModEstandar.gp_CampoNumerico(KeyAscii)
        'eventArgs.KeyChar = Chr(KeyAscii)
        'If KeyAscii = 0 Then
        '    eventArgs.Handled = True
        'End If
    End Sub

    Private Sub txtRubro_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRubro.Leave
        If CDbl(Numerico(txtAgrupador.Text)) = 0 Then
            MsgBox("Proporcione un Agrupador, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If CDbl(Numerico(txtRubro.Text)) = 0 Then
            txtRubro.Text = "000000"
        Else
            BuscaRubro()
        End If
    End Sub

    Private Sub ActualizarTotales()
        Dim lngI As Integer
        Dim varSumaD, varSumaP As Object
        varSumaD = CDec(0)
        varSumaP = CDec(0)
        With flexDetalle
            For lngI = .FixedRows To .Rows - 1 Step 1
                If Trim(.get_TextMatrix(lngI, 3)) <> "" Then
                    varSumaP = varSumaP + CDec(.get_TextMatrix(lngI, 3))
                End If
                If Trim(.get_TextMatrix(lngI, 4)) <> "" Then
                    varSumaD = varSumaD + CDec(.get_TextMatrix(lngI, 4))
                End If
            Next lngI
        End With
        txtTotalPesos.Text = VB6.Format(varSumaP, "#,##0.00")
        txtTotalDolares.Text = VB6.Format(varSumaD, "#,##0.00")
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoMensualConsultaOrigenAplicRec))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.txtRubro = New System.Windows.Forms.TextBox()
        Me.chkTodoslosRubros = New System.Windows.Forms.CheckBox()
        Me.txtAgrupador = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtTotalDolares = New System.Windows.Forms.TextBox()
        Me.txtTotalPesos = New System.Windows.Forms.TextBox()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dbcRubro = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dbcAgrupador = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblSeleccionado = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblModificados = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(322, 130)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(94, 21)
        Me.cmbAño.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(121, 130)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(145, 21)
        Me.cmbMes.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'txtRubro
        '
        Me.txtRubro.AcceptsReturn = True
        Me.txtRubro.BackColor = System.Drawing.SystemColors.Window
        Me.txtRubro.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRubro.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRubro.Location = New System.Drawing.Point(69, 38)
        Me.txtRubro.MaxLength = 6
        Me.txtRubro.Name = "txtRubro"
        Me.txtRubro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRubro.Size = New System.Drawing.Size(57, 20)
        Me.txtRubro.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtRubro, "Codigo del Rubro.")
        '
        'chkTodoslosRubros
        '
        Me.chkTodoslosRubros.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodoslosRubros.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodoslosRubros.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodoslosRubros.Location = New System.Drawing.Point(15, 13)
        Me.chkTodoslosRubros.Name = "chkTodoslosRubros"
        Me.chkTodoslosRubros.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodoslosRubros.Size = New System.Drawing.Size(121, 17)
        Me.chkTodoslosRubros.TabIndex = 2
        Me.chkTodoslosRubros.Text = "Todos los Rubros"
        Me.ToolTip1.SetToolTip(Me.chkTodoslosRubros, "Selecciona Todos los Rubros.")
        Me.chkTodoslosRubros.UseVisualStyleBackColor = False
        '
        'txtAgrupador
        '
        Me.txtAgrupador.AcceptsReturn = True
        Me.txtAgrupador.BackColor = System.Drawing.SystemColors.Window
        Me.txtAgrupador.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAgrupador.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAgrupador.Location = New System.Drawing.Point(77, 22)
        Me.txtAgrupador.MaxLength = 4
        Me.txtAgrupador.Name = "txtAgrupador"
        Me.txtAgrupador.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAgrupador.Size = New System.Drawing.Size(57, 20)
        Me.txtAgrupador.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtAgrupador, "Codigo del Agrupador.")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtTotalDolares)
        Me.Frame1.Controls.Add(Me.txtTotalPesos)
        Me.Frame1.Controls.Add(Me.flexDetalle)
        Me.Frame1.Controls.Add(Me.cmbAño)
        Me.Frame1.Controls.Add(Me.cmbMes)
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.Controls.Add(Me.dbcAgrupador)
        Me.Frame1.Controls.Add(Me.txtAgrupador)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(7, 2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(629, 359)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        '
        'txtTotalDolares
        '
        Me.txtTotalDolares.AcceptsReturn = True
        Me.txtTotalDolares.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotalDolares.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalDolares.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotalDolares.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTotalDolares.Location = New System.Drawing.Point(408, 336)
        Me.txtTotalDolares.MaxLength = 0
        Me.txtTotalDolares.Name = "txtTotalDolares"
        Me.txtTotalDolares.ReadOnly = True
        Me.txtTotalDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotalDolares.Size = New System.Drawing.Size(77, 20)
        Me.txtTotalDolares.TabIndex = 20
        Me.txtTotalDolares.TabStop = False
        Me.txtTotalDolares.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTotalPesos
        '
        Me.txtTotalPesos.AcceptsReturn = True
        Me.txtTotalPesos.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotalPesos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalPesos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotalPesos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTotalPesos.Location = New System.Drawing.Point(324, 336)
        Me.txtTotalPesos.MaxLength = 0
        Me.txtTotalPesos.Name = "txtTotalPesos"
        Me.txtTotalPesos.ReadOnly = True
        Me.txtTotalPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotalPesos.Size = New System.Drawing.Size(73, 20)
        Me.txtTotalPesos.TabIndex = 19
        Me.txtTotalPesos.TabStop = False
        Me.txtTotalPesos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(8, 162)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(613, 172)
        Me.flexDetalle.TabIndex = 7
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtRubro)
        Me.Frame2.Controls.Add(Me.chkTodoslosRubros)
        Me.Frame2.Controls.Add(Me.dbcRubro)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(8, 51)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(613, 70)
        Me.Frame2.TabIndex = 10
        Me.Frame2.TabStop = False
        '
        'dbcRubro
        '
        Me.dbcRubro.Location = New System.Drawing.Point(130, 38)
        Me.dbcRubro.Name = "dbcRubro"
        Me.dbcRubro.Size = New System.Drawing.Size(470, 21)
        Me.dbcRubro.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(15, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Rubro :"
        '
        'dbcAgrupador
        '
        Me.dbcAgrupador.Location = New System.Drawing.Point(138, 22)
        Me.dbcAgrupador.Name = "dbcAgrupador"
        Me.dbcAgrupador.Size = New System.Drawing.Size(471, 21)
        Me.dbcAgrupador.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(283, 134)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(30, 15)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Año :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(79, 134)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(33, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Mes :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(13, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(65, 21)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Agrupador :"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label8.Location = New System.Drawing.Point(54, 388)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(153, 21)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Folio Seleccionado"
        '
        'lblSeleccionado
        '
        Me.lblSeleccionado.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblSeleccionado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeleccionado.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSeleccionado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSeleccionado.Location = New System.Drawing.Point(22, 388)
        Me.lblSeleccionado.Name = "lblSeleccionado"
        Me.lblSeleccionado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSeleccionado.Size = New System.Drawing.Size(16, 16)
        Me.lblSeleccionado.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(256, 384)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(276, 53)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Presione la Barra Espaciadora o Haga Doble Click Sobre la Cuadricula Para Selecci" &
    "onar un Folio"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(54, 423)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(153, 16)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Movimientos Modificados"
        Me.Label6.Visible = False
        '
        'lblModificados
        '
        Me.lblModificados.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblModificados.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblModificados.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblModificados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblModificados.Location = New System.Drawing.Point(22, 421)
        Me.lblModificados.Name = "lblModificados"
        Me.lblModificados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblModificados.Size = New System.Drawing.Size(16, 16)
        Me.lblModificados.TabIndex = 14
        Me.lblModificados.Visible = False
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(128, 463)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 73
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(13, 463)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 72
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoMensualConsultaOrigenAplicRec
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(643, 511)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblSeleccionado)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblModificados)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(260, 106)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoMensualConsultaOrigenAplicRec"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Consulta de Origen y Aplicación de Recursos"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub
End Class
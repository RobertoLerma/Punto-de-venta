Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmRptKardexArticulo
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtCodArticulo As System.Windows.Forms.TextBox
    Public WithEvents txtDescArticulo As System.Windows.Forms.TextBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcJFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcJLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcJSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_4 As System.Windows.Forms.Label
    Public WithEvents fraJoyeria As System.Windows.Forms.Panel
    Public WithEvents dbcRModelo As System.Windows.Forms.ComboBox
    Public WithEvents dbcRMarca As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_6 As System.Windows.Forms.Label
    Public WithEvents fraRelojeria As System.Windows.Forms.Panel
    Public WithEvents optVarios As System.Windows.Forms.RadioButton
    Public WithEvents optRelojeria As System.Windows.Forms.RadioButton
    Public WithEvents optJoyeria As System.Windows.Forms.RadioButton
    Public WithEvents dbcVFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcVLinea As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_7 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_8 As System.Windows.Forms.Label
    Public WithEvents fraVarios As System.Windows.Forms.Panel
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicio As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFin As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents fraPeriodo As System.Windows.Forms.GroupBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Frame1_0 As System.Windows.Forms.GroupBox
    Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Public ResBusquedaArt As Integer
    Dim mblnSalir As Boolean
    Dim mintJFamilia As Integer
    Dim mintJLinea As Integer
    Dim mintJSubLinea As Integer
    Dim mintVFamilia As Integer
    Dim mintVLinea As Integer
    Dim mintRMarca As Integer
    Dim mintRModelo As Integer
    Dim intCodSucursal As Integer
    Dim intCodOrigen As Integer
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Dim intCodArticulo As Integer


    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Public strControlActual As String 'Nombre del control actual

    Sub Imprime()

        Dim rptKardexArticulo As New rptKardexArticulo
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr

        Dim aParam(3) As Object
        Dim aValues(3) As Object
        Dim FechaInicio As Date
        Dim FechaFin As Date
        Dim TextoAdicional As String
        Dim Encabezado As String
        Dim ConsultaGuardar As String
        Dim ConsultaReporte As String
        Dim mblnTRansaccion As Boolean
        If ValidaDatos() = False Then Exit Sub
        'FechaInicio = dtpFechaInicio
        'FechaFin = dtpFechaFin
        'TextoAdicional = Trim(ModEstandar.QuitaEnter(txtTextoAdicional))
        Encabezado = "Reporte de Kardex por Artículo"

        gStrSql = "SELECT K.*, RTRIM(LTRIM(G.NombreEmp)) AS NombreEmp, A.DescArticulo AS DescArticulo ,  A.OrigenAnt, " & "Case A.CodigoAnt When 0 Then '' Else  cast(A.OrigenAnt as nvarchar) + '-' + right('00000'+  Cast(A.CodigoAnt as varchar),5) End  as CodigoAnterior , A.CodigoAnt " & "FROM KardexArticulo('" & VB6.Format(dtpFechaInicio.Value, C_FORMATFECHAGUARDAR) & "', '" & VB6.Format(dtpFechaFin.Value, C_FORMATFECHAGUARDAR) & "', " & Trim(txtCodSucursal.Text) & " , " & Numerico(txtCodArticulo.Text) & ") K INNER JOIN " & "dbo.CatArticulos A ON K.CODARTICULO = A.CodArticulo CROSS JOIN " & "dbo.ConfiguracionGeneral G"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe que reportar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            Exit Sub
        Else
            rptKardexArticulo.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "EncabezadoReporte"
        'aValues(1) = Encabezado
        'aParam(2) = "FechaInicio"
        'aValues(2) = dtpFechaInicio.Value
        'aParam(3) = "FechaFin"
        'aValues(3) = dtpFechaFin.Value

        If (Encabezado <> Nothing) Then
            pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
            rptKardexArticulo.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
        End If

        If (dtpFechaInicio.Value <> Nothing) Then
            pdvNum.Value = dtpFechaInicio.Value : pvNum.Add(pdvNum)
            rptKardexArticulo.DataDefinition.ParameterFields("FechaInicio").ApplyCurrentValues(pvNum)
        End If

        If (dtpFechaFin.Value <> Nothing) Then
            pdvNum.Value = dtpFechaFin.Value : pvNum.Add(pdvNum)
            rptKardexArticulo.DataDefinition.ParameterFields("FechaFin").ApplyCurrentValues(pvNum)
        End If

        'frmReportes.Report = rptKardexArticulo 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Me.Text, aParam, aValues)
        frmReportes.reporteActual = rptKardexArticulo
        frmReportes.Show()
        '    Limpiar
        'Nuevo
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    '
    Function ValidaDatos() As Boolean
        If optJoyeria.Checked = False And optRelojeria.Checked = False And optVarios.Checked = False Then
            MsgBox("Debe elegir, por lo menos, un grupo con el cual generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            Exit Function
        End If
        If Trim(txtDescArticulo.Text) = "" Then
            MsgBox("Debe Proporcionar el Artículo, para generar el Reporte.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtDescArticulo.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoSucursal()
            Exit Sub
        End If
    End Sub

    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucursales.MouseUp
        PonerCodigoSucursal()
    End Sub


    'Private Sub dbcdescArticulo_Change()
    '    If mblnFueraChange = True Then Exit Sub
    '    gStrSql = "sELECT CodArticulo as CODIGO, lTRIM(RTRIM(DescArticulo)) AS  DESCRIPCION From CatArticulos A " + DevuelveQuery + "  And DescArticulo Like '" & Trim(dbcDescArticulo) & "' "
    '    txtCodArticulo = ""
    'End Sub
    '
    'Private Sub dbcdescArticulo_GotFocus()
    '    gStrSql = "sELECT CodArticulo as CODIGO, lTRIM(RTRIM(DescArticulo)) AS  DESCRIPCION From CatArticulos A " + DevuelveQuery
    '    ModDCombo.DCGotFocus gStrSql, dbcDescArticulo
    'End Sub
    '
    'Private Sub dbcdescArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    '    Select Case KeyCode
    '        Case vbKeyEscape
    '            txtCodArticulo.SetFocus
    '        Case vbKeyReturn
    '            dbcdescArticulo_LostFocus
    '    End Select
    'End Sub
    '
    'Private Sub dbcdescArticulo_LostFocus()
    '    gStrSql = "sELECT CodArticulo as CODIGO, lTRIM(RTRIM(DescArticulo)) AS  DESCRIPCION From CatArticulos A " + DevuelveQuery + "  And DescArticulo Like '" & Trim(dbcDescArticulo) & "' "
    '    ModDCombo.DCLostFocus dbcDescArticulo, gStrSql, intCodArticulo
    '    mblnFueraChange = True
    '    If intCodArticulo > 0 Then
    '        txtCodArticulo = intCodArticulo
    '    Else
    '        txtCodArticulo = ""
    '    End If
    '    mblnFueraChange = False
    'End Sub

    Private Sub dtpFechaFin_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFin.Leave
        'If CDate(dtpFechaFin) > Date Then
        '    MsgBox "La Fecha Final debe ser menor a la de Hoy." + vbNewLine + "Verifique por favor.", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
        '        dtpFechaFin.Value = Date
        '    dtpFechaFin.SetFocus
        '    Exit Sub
        'End If
        If CDate(dtpFechaFin.Value) < CDate(dtpFechaInicio.Value) Then
            MsgBox("La Fecha Final debe ser menor a la de Inicio." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            '        dtpFechaFin.Value = DateAdd("d", 1, dtpFechaInicio.Value)
            dtpFechaFin.Focus()
            Exit Sub
        End If
    End Sub


    Private Sub dtpFechaInicio_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dtpFechaInicio.KeyDown
        'If KeyCode = vbKeyEscape Then
        '    mblnSalir = True
        '    Unload Me
        '        KeyCode = 0
        'End If
    End Sub


    Private Sub dtpFechaInicio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicio.Leave
        If CDate(dtpFechaInicio.Value) > Today Then
            MsgBox("La Fecha de Inicio debe ser menor o igual a la de Hoy." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dtpFechaInicio.Value = Today
            dtpFechaInicio.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub optjoyeria_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optJoyeria.CheckedChanged
        If eventSender.Checked Then
            Select Case Me.optJoyeria.Checked
                Case True
                    fraJoyeria.Visible = True
                    fraRelojeria.Visible = False
                    fraVarios.Visible = False
                    mblnFueraChange = True
                    mintJFamilia = 0
                    Me.dbcJFamilia.Text = C_TODAS
                    Me.dbcJFamilia.Enabled = True
                    mintJLinea = 0
                    Me.dbcJLinea.Text = C_TODAS
                    Me.dbcJLinea.Enabled = False
                    mintJSubLinea = 0
                    Me.dbcJSubLinea.Text = C_TODAS
                    Me.dbcJSubLinea.Enabled = False
                    mblnFueraChange = False

                    mblnFueraChange = True
                    txtCodArticulo.Text = ""
                    txtDescArticulo.Text = ""
                    mblnFueraChange = False
            End Select
        End If
    End Sub

    Private Sub optrelojeria_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRelojeria.CheckedChanged
        If eventSender.Checked Then
            Select Case Me.optRelojeria.Checked
                Case System.Windows.Forms.CheckState.Checked
                    fraJoyeria.Visible = False
                    fraRelojeria.Visible = True
                    fraVarios.Visible = False
                    mblnFueraChange = True
                    mblnFueraChange = True
                    mintRMarca = 0
                    Me.dbcRMarca.Text = C_TODAS
                    Me.dbcRMarca.Enabled = True
                    mintRModelo = 0
                    Me.dbcRModelo.Text = C_TODOS
                    Me.dbcRModelo.Enabled = False
                    mblnFueraChange = False

                    mblnFueraChange = True
                    txtCodArticulo.Text = ""
                    txtDescArticulo.Text = ""
                    mblnFueraChange = False
            End Select
        End If
    End Sub
    Private Sub optVarios_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optVarios.CheckedChanged
        If eventSender.Checked Then
            Select Case Me.optVarios.Checked
                Case System.Windows.Forms.CheckState.Checked
                    fraJoyeria.Visible = False
                    fraRelojeria.Visible = False
                    fraVarios.Visible = True
                    mblnFueraChange = True
                    mblnFueraChange = True
                    mintVFamilia = 0
                    Me.dbcVFamilia.Text = C_TODAS
                    Me.dbcVFamilia.Enabled = True
                    mintVLinea = 0
                    Me.dbcVLinea.Text = C_TODAS
                    Me.dbcVLinea.Enabled = False
                    mblnFueraChange = False

                    mblnFueraChange = True
                    txtCodArticulo.Text = ""
                    txtDescArticulo.Text = ""
                    mblnFueraChange = False
            End Select
        End If
    End Sub

    Private Sub dbcJFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Leave
        'gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia) & "%' ORDER BY DescFamilia"
        'ModDCombo.DCLostFocus dbcJFamilia, gStrSql, mintJFamilia

        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
        Aux = mintJFamilia
        mintJFamilia = 0
        If Trim(Me.dbcJFamilia.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJFamilia.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJFamilia), gStrSql, mintJFamilia)
        End If
        If Aux <> mintJFamilia Then
            If mintJFamilia = 0 Then
                mblnFueraChange = True
                Me.dbcJFamilia.Text = C_TODAS
                Me.dbcJFamilia.Enabled = True
                mintJLinea = 0
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = False
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJLinea = 0
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = True
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
                Me.dbcJLinea.Focus()
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcJFamilia.Text) = "" Then Me.dbcJFamilia.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcJLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
        Aux = mintJLinea
        mintJLinea = 0
        If Trim(Me.dbcJLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJLinea), gStrSql, mintJLinea)
        End If
        If Aux <> mintJLinea Then
            If mintJLinea = 0 Then
                mblnFueraChange = True
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = True
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = True
                mblnFueraChange = False
                Me.dbcJSubLinea.Focus()
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcJLinea.Text) = "" Then Me.dbcJLinea.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcJSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Leave

        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  And (CodLinea = " & mintJLinea & ") and DescSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%' ORDER BY DescSubLinea"
        Aux = mintJSubLinea
        mintJSubLinea = 0
        If Trim(Me.dbcJSubLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJSubLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJSubLinea), gStrSql, mintJSubLinea)
        End If
        If Aux <> mintJSubLinea Then
            If mintJSubLinea = 0 Then
                mblnFueraChange = True
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = True
                mblnFueraChange = False
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcJSubLinea.Text) = "" Then Me.dbcJSubLinea.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    'Private Sub dbcOrigen_Change()
    '    On Error GoTo MError
    '    Dim lStrSql As String
    '
    '    If mblnFueraChange Then Exit Sub
    '
    '    lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me.dbcOrigen.text) & "%'"
    '    ModDCombo.DCChange lStrSql, tecla, Me.dbcOrigen
    '    IntCodOrigen = -1
    '    mblnFueraChange = True
    '    txtCodOrigen = ""
    '    mblnFueraChange = False
    'MError:
    '    If Err.Number <> 0 Then
    '        ModEstandar.MostrarError
    '    End If
    'End Sub
    '
    'Private Sub dbcOrigen_GotFocus()
    '    Pon_Tool
    '    gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen"
    '    ModDCombo.DCGotFocus gStrSql, Me.dbcOrigen
    'End Sub
    '
    'Private Sub dbcOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    '    If KeyCode = vbKeyEscape Then
    '        Me.txtCodOrigen.SetFocus
    '        KeyCode = 0
    '    End If
    '    tecla = KeyCode
    'End Sub

    'Private Sub dbcOrigen_LostFocus()
    '    Dim I As Integer
    '    If Screen.ActiveForm.Name <> Me.Name Then
    '        Exit Sub
    '    End If
    '    gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me.dbcOrigen.text) & "%'"
    '    IntCodOrigen = -1
    '    ModDCombo.DCLostFocus Me.dbcOrigen, gStrSql, IntCodOrigen
    '    mblnFueraChange = True
    '    If IntCodOrigen = -1 Or Trim(dbcOrigen) = "" Then
    '        txtCodOrigen = ""
    '    Else
    '        txtCodOrigen = IntCodOrigen
    '    End If
    '    mblnFueraChange = False
    'End Sub
    '

    Private Sub dbcRMarca_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcRMarca.Name Then Exit Sub
        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " and DescMarca LIKE '" & Trim(dbcRMarca.Text) & "%' ORDER BY DescMarca"
        ModDCombo.DCChange(gStrSql, tecla)
        '    LimpiaDatosMarca
    End Sub

    Private Sub dbcRMarca_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Enter
        Pon_Tool()
        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " ORDER BY DescMarca"
        ModDCombo.DCGotFocus(gStrSql, dbcRMarca)
    End Sub

    Private Sub dbcRMarca_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcRMarca.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.optRelojeria.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcRMarca_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " and DescMarca LIKE '" & Trim(dbcRMarca.Text) & "%' ORDER BY DescMarca"
        Aux = mintRMarca
        mintRMarca = 0
        If Trim(Me.dbcRMarca.Text) <> Trim(C_TODAS) Or Trim(Me.dbcRMarca.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcRMarca), gStrSql, mintRMarca)
        End If

        If Aux <> mintRMarca Then
            If mintRMarca = 0 Then
                mblnFueraChange = True
                Me.dbcRMarca.Text = C_TODAS
                Me.dbcRMarca.Enabled = True
                mintRModelo = 0
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintRModelo = 0
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = True
                mblnFueraChange = False
                Me.dbcRModelo.Focus()
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcRMarca.Text) = "" Then Me.dbcRMarca.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    '''Relojeria --Modelos
    Private Sub dbcRmodelo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcRModelo.Name Then Exit Sub
        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & mintRMarca & " and Descmodelo LIKE '" & Trim(dbcRModelo.Text) & "%' ORDER BY Descmodelo"
        ModDCombo.DCChange(gStrSql, tecla)
        '    LimpiaDatosPrecioYDescuento
    End Sub

    Private Sub dbcRmodelo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.Enter
        Pon_Tool()
        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & mintRMarca & " ORDER BY Descmodelo"
        ModDCombo.DCGotFocus(gStrSql, dbcRModelo)
    End Sub

    Private Sub dbcRmodelo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcRModelo.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcRMarca.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcRModelo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & mintRMarca & " and Descmodelo LIKE '" & Trim(dbcRModelo.Text) & "%' ORDER BY Descmodelo"
        Aux = mintRModelo
        mintRModelo = 0
        If Trim(Me.dbcRModelo.Text) <> Trim(C_TODOS) Or Trim(Me.dbcRModelo.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcRModelo), gStrSql, mintRModelo)
        End If
        If Aux <> mintRModelo Then
            If mintRModelo = 0 Then
                mblnFueraChange = True
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = True
                mblnFueraChange = False
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcRModelo.Text) = "" Then Me.dbcRModelo.Text = C_TODOS
        mblnFueraChange = False
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged

        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        mblnFueraChange = True
        txtCodSucursal.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen where TipoAlmacen ='P'ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        mblnFueraChange = True
        If intCodSucursal = 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = CStr(intCodSucursal)
        End If
        mblnFueraChange = False
    End Sub

    Private Sub frmRptKardexArticulo_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmRptKardexArticulo_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub Form_Initialize_Renamed()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmRptKardexArticulo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmRptKardexArticulo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()

        Me.optJoyeria.Checked = True
        optjoyeria_CheckedChanged(optJoyeria, New System.EventArgs())

        Me.optRelojeria.Checked = False
        optrelojeria_CheckedChanged(optRelojeria, New System.EventArgs())

        Me.optVarios.Checked = False
        optVarios_CheckedChanged(optVarios, New System.EventArgs())

        txtCodSucursal.Text = ""
        intCodSucursal = 0
        mintJFamilia = 0
        mintJLinea = 0
        mintJSubLinea = 0
        mintRMarca = 0
        mintRModelo = 0
        mintVFamilia = 0
        mintVLinea = 0
        mblnFueraChange = True
        dtpFechaInicio.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        dtpFechaFin.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        txtCodArticulo.Text = ""
        txtDescArticulo.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub frmRptKardexArticulo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        dtpFechaInicio.MinDate = C_FECHAINICIAL
        dtpFechaInicio.MaxDate = C_FECHAFINAL
        dtpFechaFin.MinDate = C_FECHAINICIAL
        dtpFechaFin.MaxDate = C_FECHAFINAL
        Nuevo()
        '    txtCodSucursal = gintCodAlmacen
        '    txtCodsucursal_LostFocus
    End Sub

    Private Sub frmRptKardexArticulo_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    Cancel = 0 'Para que no salga del Formulario hasta que guarde los datos, si no tiene premiso de hacerlo
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes 'Sale del Formulario, pero antes preguntar si desea grabar los datos registrados, solo cuando es nuevo
        '            Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmRptKardexArticulo_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
        dtpFechaInicio.Focus()
    End Sub

    Private Sub dbcJFAmilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcJFamilia.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.optJoyeria.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJFAmilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcJFamilia.Name Then Exit Sub
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcjFAmilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Enter
        Pon_Tool()
        '    If sstGrupos.Tab = 0 Then
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " ORDER BY DescFamilia"
        '    Else
        '        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " ORDER BY DescFamilia"
        '    End If
        ModDCombo.DCGotFocus(gStrSql, dbcJFamilia)
    End Sub

    Private Sub dbcJLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcJLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcJFamilia.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode

    End Sub

    Private Sub dbcJLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcJLinea.Name Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcJLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Enter
        If mblnFueraChange = True Then Exit Sub

        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  ORDER BY DescLinea"
        '    Else
        '        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & Numerico(GridActivo.TextMatrix(GridActivo.Row, C_ColJCODFAMILIA)) & ")  ORDER BY DescLinea"
        '    End If
        ModDCombo.DCGotFocus(gStrSql, dbcJLinea)
    End Sub

    Private Sub dbcJSubLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcJSubLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcJLinea.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJSubLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcJSubLinea.Name Then Exit Sub
        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  And (CodLinea = " & mintJLinea & ") and DescSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%' ORDER BY DescSubLinea"
        ModDCombo.DCChange(gStrSql, tecla)
        'LimpiaDatosPrecioYDescuento
    End Sub

    Private Sub dbcJSubLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Enter
        Pon_Tool()
        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  And (CodLinea = " & mintJLinea & ") ORDER BY DescSubLinea"
        ModDCombo.DCGotFocus(gStrSql, dbcJSubLinea)
    End Sub

    Private Sub dbcVFamilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcVFamilia.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.optVarios.Focus()
            eventSender.KeyCode = 0
        ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            '        AvanzarTab Me
            dbcVFamilia_Leave(dbcVFamilia, New System.EventArgs())
            '        KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVFamilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcVFamilia.Name Then Exit Sub
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " and DescFamilia LIKE '" & Trim(dbcVFamilia.Text) & "%' ORDER BY DescFamilia"
        ModDCombo.DCChange(gStrSql, tecla)
        If Trim(Me.dbcVFamilia.Text) = "" Then
            dbcVFamilia_Leave(dbcVFamilia, New System.EventArgs())
        End If
    End Sub

    Private Sub dbcVFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " ORDER BY DescFamilia"
        ModDCombo.DCGotFocus(gStrSql, dbcVFamilia)
    End Sub

    Private Sub dbcVFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " and DescFamilia LIKE '" & Trim(dbcVFamilia.Text) & "%' ORDER BY DescFamilia"
        Aux = mintVFamilia
        mintVFamilia = 0
        If Trim(Me.dbcVFamilia.Text) <> Trim(C_TODAS) Or Trim(Me.dbcVFamilia.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcVFamilia), gStrSql, mintVFamilia)
        End If

        If Aux <> mintVFamilia Then
            If mintVFamilia = 0 Then
                mblnFueraChange = True
                Me.dbcVFamilia.Text = C_TODAS
                Me.dbcVFamilia.Enabled = True
                mintVLinea = 0
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintVLinea = 0
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = True
                mblnFueraChange = False
                '''If dbcVLinea.Enabled Then Me.dbcVLinea.SetFocus
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcVFamilia.Text) = "" Then Me.dbcVFamilia.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcVLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcVLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcVFamilia.Focus()
            eventSender.KeyCode = 0
        ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            '        txtCodOrigen.SetFocus
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcVLinea.Name Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ") and DescLinea LIKE '" & Trim(dbcVLinea.Text) & "%' ORDER BY DescLinea"
        ModDCombo.DCChange(gStrSql, tecla)
        If Trim(Me.dbcVLinea.Text) = "" Then
            dbcVLinea_Leave(dbcVLinea, New System.EventArgs())
        End If
    End Sub

    Private Sub dbcVLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Enter
        If mblnFueraChange = True Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ")  ORDER BY DescLinea"
        ModDCombo.DCGotFocus(gStrSql, dbcVLinea)
    End Sub

    Private Sub dbcVLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        '    mblnFueraChange = True
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ") and DescLinea LIKE '" & Trim(dbcVLinea.Text) & "%' ORDER BY DescLinea"
        Aux = mintVLinea
        mintVLinea = 0
        If Trim(Me.dbcVLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcVLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcVLinea), gStrSql, mintVLinea)
        End If
        If Aux <> mintVLinea Then
            If mintVLinea = 0 Then
                mblnFueraChange = True
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = True
                mblnFueraChange = False
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcVLinea.Text) = "" Then Me.dbcVLinea.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub txtCodOrigen_Change()
        If mblnFueraChange = True Then Exit Sub
        mblnFueraChange = True
        '    dbcOrigen = ""
        mblnFueraChange = False
    End Sub

    Private Sub txtCodArticulo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.TextChanged
        If mblnFueraChange = True Then Exit Sub
        mblnFueraChange = True
        txtDescArticulo.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub txtCodArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.Enter
        strControlActual = UCase("txtCodArticulo")
        SelTextoTxt(txtCodArticulo)
    End Sub


    Private Sub txtCodArticulo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodArticulo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtCodArticulo.Text, KeyAscii, 8, 0, (txtCodArticulo.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.Leave
        ''''    LlenaDatosArticulo CLng((Val(txtCodArticulo)))
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        Dim CodAux As Integer
        If Trim(txtCodArticulo.Text) <> "" Then
            ResBusquedaArt = BuscarCodigoArticulo(Trim(txtCodArticulo.Text))
            If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then
                LlenaDatosArticulo(CDbl(ResBusquedaArt))
            ElseIf ResBusquedaArt = -2 Then
                CodAux = CInt(txtCodArticulo.Text)
                txtCodArticulo.Text = ""
                BuscarArticulos(True, VB.Right(New String("0", 6) & Trim(CStr(CodAux)), 6))
            End If
        End If
    End Sub

    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        If mblnFueraChange = True Then Exit Sub
        mblnFueraChange = True
        dbcSucursales.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        SelTextoTxt(txtCodSucursal)
    End Sub

    Private Sub txtCodSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodSucursal.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            KeyCode = 0
        End If
    End Sub

    Private Sub txtCodSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Leave
        LlenaDatosSucursal()
    End Sub

    Sub LlenaDatosSucursal()
        If CDbl(Numerico(Trim(txtCodSucursal.Text))) = 0 Then Exit Sub
        gStrSql = "SELECT      Ltrim(Rtrim(DescAlmacen)) as DescAlmacen From dbo.CatAlmacen Where CodAlmacen =" & Numerico(txtCodSucursal.Text) & "  And TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mblnFueraChange = True
            dbcSucursales.Text = RsGral.Fields("DescAlmacen").Value
            mblnFueraChange = False
        Else
            MsgBox("El código de sucursal no existe." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            txtCodSucursal.Text = ""
            txtCodSucursal.Focus()
            Exit Sub
        End If
    End Sub

    Sub LlenaDatosArticulo(ByRef CodArticulo As Integer)
        If CDbl(Numerico(Trim(CStr(CodArticulo)))) = 0 Then Exit Sub
        gStrSql = "SELECT      Ltrim(Rtrim(DescArticulo)) as DEscArticulo From CatArticulos " & DevuelveQuery() & " and CodArticulo = " & CodArticulo & ""

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mblnFueraChange = True
            txtCodArticulo.Text = CStr(CodArticulo)
            txtDescArticulo.Text = RsGral.Fields("DescArticulo").Value
            mblnFueraChange = False
        Else
            MsgBox("El código de artículo no existe." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            txtCodArticulo.Text = ""
            txtCodArticulo.Focus()
            Exit Sub
        End If
    End Sub

    Public Function DevuelveQuery() As String
        On Error GoTo Merr
        Dim I As Integer
        Dim cSELECT As String
        Dim cFROM As String
        Dim cWHERE As String
        Dim cGROUPBY As String
        Dim cORDERBY As String
        Dim rsLocal As ADODB.Recordset
        Dim cMSG As String

        Dim nJOYERIA As Integer
        Dim nRELOJERIA As Integer
        Dim nVARIOS As Integer

        'Obtener los códigos que va a tomar en cuenta en la consulta; estos códigos se enviarán como parámetros al
        'procedimiento almacenado que recopilará los datos

        nJOYERIA = IIf(Me.optJoyeria.Checked = True, 1, 0)
        nRELOJERIA = IIf(Me.optRelojeria.Checked = True, 1, 0)
        nVARIOS = IIf(Me.optVarios.Checked = True, 1, 0)

        If nJOYERIA = 0 And nRELOJERIA = 0 And nVARIOS = 0 Then
            MsgBox("Debe elegir, por lo menos, un grupo con el cual generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            Exit Function
        End If

        cWHERE = " Where  "

        Select Case True
            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS > 0
                'Todos los grupos
                cWHERE = cWHERE & " CodGrupo In (" & gCODJOYERIA & ", " & gCODRELOJERIA & ", " & gCODVARIOS & ") "
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        ' Todos
                        cWHERE = cWHERE & " and ((CodFamilia <> " & 0 & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea = " & mintJSubLinea & ")"
                End Select
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        cWHERE = cWHERE & " or (CodMarca <> " & 0 & ")"
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " or (CodMarca = " & mintRMarca & ")"
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " or (CodMarca = " & mintRMarca & " and CodModelo = " & mintRModelo & ")"
                End Select
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " or (CodFamilia <> 0 and CodSubLinea is NULL))"
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " or (CodFamilia = " & mintVFamilia & " and CodSubLinea is NULL))"
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " or (CodFamilia = " & mintVFamilia & " and CodLinea = " & mintVLinea & " and CodSubLinea is NULL))"
                End Select
            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS <= 0
                'Joyeria-Relojeria
                cWHERE = cWHERE & " CodGrupo <> " & gCODVARIOS
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        ' Todos
                        cWHERE = cWHERE & " and ((CodFamilia <> " & 0 & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea = " & mintJSubLinea & ")"
                End Select
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        cWHERE = cWHERE & " or (CodMarca <> " & 0 & "))"
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " or (CodMarca = " & mintRMarca & "))"
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " or (CodMarca = " & mintRMarca & " and CodModelo = " & mintRModelo & "))"
                End Select
            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS > 0
                'Joyeria-Varios
                cWHERE = cWHERE & " CodGrupo <> " & gCODRELOJERIA
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        ' Todos
                        cWHERE = cWHERE & " and ((CodFamilia <> " & 0 & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea is NOT NULL)"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and ((CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea = " & mintJSubLinea & ")"
                End Select
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " or (CodFamilia <> 0) and CodSubLinea is NULL)"
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " or (CodFamilia = " & mintVFamilia & " and CodSubLinea is NULL))"
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " or (CodFamilia = " & mintVFamilia & " and CodLinea = " & mintVLinea & " and CodSubLinea is NULL))"
                End Select
            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS <= 0
                'Joyeria
                cWHERE = cWHERE & " CodGrupo = " & gCODJOYERIA
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        ' Todos
                        '''cWHERE = cWHERE & " and CodFamilia <> " & 0 & " and CodSubLinea is NOT NULL "
                        cWHERE = cWHERE & " and CodFamilia <> 0 "
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        '''cWHERE = cWHERE & " and CodFamilia = " & mintJFamilia & " and CodSubLinea is NOT NULL "
                        cWHERE = cWHERE & " and CodFamilia = " & mintJFamilia & " "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        '''cWHERE = cWHERE & " and CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea is NOT NULL"
                        cWHERE = cWHERE & " and CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and CodSubLinea = " & mintJSubLinea
                End Select
            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS > 0
                'Relojeria-Varios
                cWHERE = cWHERE & " CodGrupo <> " & gCODJOYERIA
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        cWHERE = cWHERE & " and ((CodMarca <> " & 0 & ")"
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " and ((CodMarca = " & mintRMarca & ")"
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " and ((CodMarca = " & mintRMarca & " and CodModelo = " & mintRModelo & ")"
                End Select
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " or (CodFamilia <> 0) and CodSubLinea is NULL)"
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " or (CodFamilia = " & mintVFamilia & " and CodSubLinea is NULL))"
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " or (CodFamilia = " & mintVFamilia & " and CodLinea = " & mintVLinea & " and CodSubLinea is NULL))"
                End Select
            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS <= 0
                'Relojeria
                cWHERE = cWHERE & " CodGrupo = " & gCODRELOJERIA
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        '''cWHERE = cWHERE & " and CodMarca <> " & 0
                        cWHERE = cWHERE & " and CodMarca <> 0 "
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " and CodMarca = " & mintRMarca
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " and CodMarca = " & mintRMarca & " and CodModelo = " & mintRModelo
                End Select
            Case nJOYERIA <= 0 And nRELOJERIA <= 0 And nVARIOS > 0
                'Varios
                cWHERE = cWHERE & " CodGrupo = " & gCODVARIOS
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " and CodFamilia <> 0 and CodSubLinea is NULL "
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " and CodFamilia = " & mintVFamilia & " and CodSubLinea is NULL "
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " and CodFamilia = " & mintVFamilia & " and CodLinea = " & mintVLinea & " and CodSubLinea is NULL "
                End Select
        End Select


        DevuelveQuery = cWHERE ' & " and I.CodAlmacen = " & txtCodSucursal & " order by I.codarticulo "

        Exit Function

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    'Sub Buscar() 'Articulo()
    '    On Local Error GoTo Merr:
    '    Dim strSQL As String
    '    Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
    '    Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
    '    Dim strControlActual As String 'Nombre del control actual
    '    Dim Columna As Integer
    '
    '    strControlActual = UCase(Screen.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
    '    strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
    '
    '    Select Case strControlActual
    '        Case "TXTCODARTICULO"
    '            strCaptionForm = "Consulta de Articulos"
    '            strSQL = "sELECT CodArticulo as CODIGO, lTRIM(RTRIM(DescArticulo)) AS  DESCRIPCION From CatArticulos A " + DevuelveQuery
    '        Case Else
    '            'Sale de este sub para ke no ejecute ninguna opcion
    '            Exit Sub
    '    End Select
    '
    '    ModEstandar.BorraCmd
    '    Cmd.CommandText = "dbo.Up_Select_Datos"
    '    Cmd.CommandType = adCmdStoredProc
    '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, strSQL)
    '    Set RsGral = Cmd.Execute
    '
    '    'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
    '    If RsGral.RecordCount = 0 Then
    '        MsgBox C_msgSINDATOS & vbNewLine & "Verifique por favor....", vbExclamation, gstrCorpoNombreEmpresa
    '        RsGral.Close
    '        Exit Sub
    '    End If
    '
    '    'Carga el formulario de consulta
    '    Load FrmConsultas
    '    Call ConfiguraConsultas(FrmConsultas, 6400, RsGral, strTag, strCaptionForm)
    '
    '    With FrmConsultas.Flexdet
    '        Select Case strControlActual
    '            Case "TXTCODARTICULO"
    '                .ColWidth(0) = 900
    '                .ColWidth(1) = 5500
    '                .ColAlignment(0) = flexAlignCenterCenter
    '                .ColAlignment(1) = flexAlignLeftCenter
    '            Case Else
    '                Exit Sub
    '        End Select
    '    End With
    '
    '    FrmConsultas.Show vbModal
    '
    'Merr:
    '    If Err.Number <> 0 Then ModEstandar.MostrarError
    ''    Resume
    'End Sub

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        'Dim strControlActual As String 'Nombre del control actual
        Dim Columna As Integer


        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control


        Select Case strControlActual
            Case "TXTCODARTICULO"
                strCaptionForm = "Consulta de Artículos"
                strSQL = "SELECT     A.CodArticulo AS CODIGO, LTRIM(RTRIM(A.DescArticulo)) AS DESCRIPCION, M.DescTipoMaterial AS MATERIAL, A.CodigoArticuloProv AS [ARTICULO PROV] " & "FROM dbo.CatArticulos A lEFT OUTER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  " & DevuelveQuery()

            Case "TXTDESCARTICULO"
                strCaptionForm = "Consulta de Artículos"
                strSQL = "SELECT      LTRIM(RTRIM(A.DescArticulo))  AS DESCRIPCION,A.CodArticulo AS CODIGO, M.DescTipoMaterial AS MATERIAL, A.CodigoArticuloProv AS [ARTICULO PROV] " & "FROM dbo.CatArticulos A LEFT OUTER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  " & DevuelveQuery() & " AND  DescArticulo Like '" & Trim(txtDescArticulo.Text) & "%'"
            Case Else
                'Sale de este sub para ke no ejecute ninguna opcion
                Exit Sub
        End Select

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta 
        ConfiguraConsultas(FrmConsultas, 9200, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODARTICULO"
                    .set_ColWidth(0, 0, 900)
                    .set_ColWidth(1, 0, 4800)
                    .set_ColWidth(2, 0, 1900)
                    .set_ColWidth(3, 0, 1620)

                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)

                Case "TXTDESCARTICULO"
                    .set_ColWidth(0, 0, 4800)
                    .set_ColWidth(1, 0, 900)
                    .set_ColWidth(2, 0, 1900)
                    .set_ColWidth(3, 0, 1620)
                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            End Select
            .WordWrap = False
        End With
        mblnFueraChange = True
        CentrarForma(FrmConsultas)
        FrmConsultas.ShowDialog()
        mblnFueraChange = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub txtDescArticulo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescArticulo.TextChanged
        If mblnFueraChange = True Then Exit Sub
        mblnFueraChange = True
        txtCodArticulo.Text = ""
        mblnFueraChange = False
    End Sub


    Private Sub txtDescArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescArticulo.Enter
        strControlActual = UCase("txtDescArticulo")
        SelTextoTxt(txtDescArticulo)
    End Sub

    Sub PonerCodigoSucursal()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'DCLostFocus dbcAlmacenSalida, gStrSql, intCodSucursal
        mblnFueraChange = True
        If RsGral.RecordCount <= 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = Numerico(RsGral.Fields("CodAlmacen").Value)
        End If
        mblnFueraChange = False
    End Sub

    Sub BuscarArticulos(ByRef BusquedaEspecial As Boolean, ByRef CodArticulo As String)
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim Columna As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        strControlActual = IIf((BusquedaEspecial), "TXTCODARTICULO", strControlActual)
        Select Case strControlActual
            Case "TXTCODARTICULO"
                strCaptionForm = "Consulta de Artículos"
                If BusquedaEspecial Then
                    strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR], " & "dbo.FormatCantidad(A.PrecioPubDolar)  AS [PRECIO PÚBLICO] , " & "case PesosFijos WHEN 0 THEN 'DÓLARES' WHEN 1 THEN 'PESOS' END AS [MONEDA] " & "From CatArticulos A cross Join Configuraciongeneral c WHERE (CodArticulo = " & CInt(CodArticulo) & ") " & "OR   (OrigenAnt = " & CInt(VB.Left(CodArticulo, 1)) & ") AND (CodigoAnt = " & CInt(VB.Right(CodArticulo, 5)) & ")"
                Else
                    strSQL = "SELECT     A.CodArticulo AS CODIGO, LTRIM(RTRIM(A.DescArticulo)) AS DESCRIPCION, M.DescTipoMaterial AS MATERIAL, LTrim(Rtrim(A.CodigoArticuloProv)) AS [ARTICULO PROV],  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+ '-'+ RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]   " & "FROM dbo.CatArticulos A INNER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  " & DevuelveQuery()
                End If
            Case "TXTDESCARTICULO"
                strCaptionForm = "Consulta de Artículos"
                strSQL = "SELECT      LTRIM(RTRIM(A.DescArticulo))  AS DESCRIPCION,A.CodArticulo AS CODIGO, M.DescTipoMaterial AS MATERIAL, LTrim(Rtrim(A.CodigoArticuloProv)) AS [ARTICULO PROV],  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+ '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]   " & "FROM dbo.CatArticulos A INNER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  " & DevuelveQuery() & " AND  DescArticulo Like '" & Trim(txtDescArticulo.Text) & "%'"
            Case Else
                'Sale de este sub para ke no ejecute ninguna opcion
                Exit Sub
        End Select

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            If BusquedaEspecial = True Then
                MsgBox("El Artículo no existe." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                RsGral.Close()
                Exit Sub
            Else
                MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                RsGral.Close()
                Exit Sub
            End If
        End If

        'Carga el formulario de consulta
        ConfiguraConsultas(FrmConsultas, 11050, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODARTICULO"
                    .set_ColWidth(0, 0, 900)
                    .set_ColWidth(1, 0, 4800)
                    .set_ColWidth(2, 0, 1900)
                    .set_ColWidth(3, 0, 1620)
                    .set_ColWidth(4, 0, 1800)

                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

                Case "TXTDESCARTICULO"
                    .set_ColWidth(0, 0, 4800)
                    .set_ColWidth(1, 0, 900)
                    .set_ColWidth(2, 0, 1900)
                    .set_ColWidth(3, 0, 1620)
                    .set_ColWidth(4, 0, 1800)
                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            End Select
            .WordWrap = False
        End With
        mblnFueraChange = True
        CentrarForma(FrmConsultas)
        FrmConsultas.ShowDialog()
        mblnFueraChange = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodArticulo = New System.Windows.Forms.TextBox()
        Me._Frame1_0 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtDescArticulo = New System.Windows.Forms.TextBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.fraJoyeria = New System.Windows.Forms.Panel()
        Me.dbcJFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcJLinea = New System.Windows.Forms.ComboBox()
        Me.dbcJSubLinea = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me._lblVentas_4 = New System.Windows.Forms.Label()
        Me.fraRelojeria = New System.Windows.Forms.Panel()
        Me.dbcRModelo = New System.Windows.Forms.ComboBox()
        Me.dbcRMarca = New System.Windows.Forms.ComboBox()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._lblVentas_6 = New System.Windows.Forms.Label()
        Me.optVarios = New System.Windows.Forms.RadioButton()
        Me.optRelojeria = New System.Windows.Forms.RadioButton()
        Me.optJoyeria = New System.Windows.Forms.RadioButton()
        Me.fraVarios = New System.Windows.Forms.Panel()
        Me.dbcVFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcVLinea = New System.Windows.Forms.ComboBox()
        Me._lblVentas_7 = New System.Windows.Forms.Label()
        Me._lblVentas_8 = New System.Windows.Forms.Label()
        Me.fraPeriodo = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicio = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFin = New System.Windows.Forms.DateTimePicker()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Frame1 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me._Frame1_0.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.fraJoyeria.SuspendLayout()
        Me.fraRelojeria.SuspendLayout()
        Me.fraVarios.SuspendLayout()
        Me.fraPeriodo.SuspendLayout()
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCodArticulo
        '
        Me.txtCodArticulo.AcceptsReturn = True
        Me.txtCodArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodArticulo.Location = New System.Drawing.Point(8, 24)
        Me.txtCodArticulo.MaxLength = 8
        Me.txtCodArticulo.Name = "txtCodArticulo"
        Me.txtCodArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodArticulo.Size = New System.Drawing.Size(81, 20)
        Me.txtCodArticulo.TabIndex = 23
        Me.txtCodArticulo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCodArticulo, "Código del Artículo")
        '
        '_Frame1_0
        '
        Me._Frame1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_0.Controls.Add(Me.Frame2)
        Me._Frame1_0.Controls.Add(Me.Frame3)
        Me._Frame1_0.Controls.Add(Me.fraPeriodo)
        Me._Frame1_0.Controls.Add(Me.txtCodSucursal)
        Me._Frame1_0.Controls.Add(Me.dbcSucursales)
        Me._Frame1_0.Controls.Add(Me._Label1_1)
        Me._Frame1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.SetIndex(Me._Frame1_0, CType(0, Short))
        Me._Frame1_0.Location = New System.Drawing.Point(8, 0)
        Me._Frame1_0.Name = "_Frame1_0"
        Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_0.Size = New System.Drawing.Size(449, 297)
        Me._Frame1_0.TabIndex = 24
        Me._Frame1_0.TabStop = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtCodArticulo)
        Me.Frame2.Controls.Add(Me.txtDescArticulo)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(8, 232)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(433, 57)
        Me.Frame2.TabIndex = 22
        Me.Frame2.TabStop = False
        Me.Frame2.Text = " Artículo "
        '
        'txtDescArticulo
        '
        Me.txtDescArticulo.AcceptsReturn = True
        Me.txtDescArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescArticulo.Location = New System.Drawing.Point(104, 24)
        Me.txtDescArticulo.MaxLength = 0
        Me.txtDescArticulo.Name = "txtDescArticulo"
        Me.txtDescArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescArticulo.Size = New System.Drawing.Size(323, 20)
        Me.txtDescArticulo.TabIndex = 32
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.fraJoyeria)
        Me.Frame3.Controls.Add(Me.fraRelojeria)
        Me.Frame3.Controls.Add(Me.optVarios)
        Me.Frame3.Controls.Add(Me.optRelojeria)
        Me.Frame3.Controls.Add(Me.optJoyeria)
        Me.Frame3.Controls.Add(Me.fraVarios)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(8, 120)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(433, 105)
        Me.Frame3.TabIndex = 25
        Me.Frame3.TabStop = False
        '
        'fraJoyeria
        '
        Me.fraJoyeria.BackColor = System.Drawing.SystemColors.Control
        Me.fraJoyeria.Controls.Add(Me.dbcJFamilia)
        Me.fraJoyeria.Controls.Add(Me.dbcJLinea)
        Me.fraJoyeria.Controls.Add(Me.dbcJSubLinea)
        Me.fraJoyeria.Controls.Add(Me._lblVentas_0)
        Me.fraJoyeria.Controls.Add(Me._lblVentas_3)
        Me.fraJoyeria.Controls.Add(Me._lblVentas_4)
        Me.fraJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraJoyeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraJoyeria.Location = New System.Drawing.Point(88, 8)
        Me.fraJoyeria.Name = "fraJoyeria"
        Me.fraJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraJoyeria.Size = New System.Drawing.Size(337, 89)
        Me.fraJoyeria.TabIndex = 8
        '
        'dbcJFamilia
        '
        Me.dbcJFamilia.Location = New System.Drawing.Point(78, 12)
        Me.dbcJFamilia.Name = "dbcJFamilia"
        Me.dbcJFamilia.Size = New System.Drawing.Size(253, 21)
        Me.dbcJFamilia.TabIndex = 13
        '
        'dbcJLinea
        '
        Me.dbcJLinea.Location = New System.Drawing.Point(78, 36)
        Me.dbcJLinea.Name = "dbcJLinea"
        Me.dbcJLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcJLinea.TabIndex = 15
        '
        'dbcJSubLinea
        '
        Me.dbcJSubLinea.Location = New System.Drawing.Point(78, 60)
        Me.dbcJSubLinea.Name = "dbcJSubLinea"
        Me.dbcJSubLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcJSubLinea.TabIndex = 17
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_0, CType(0, Short))
        Me._lblVentas_0.Location = New System.Drawing.Point(16, 16)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_0.TabIndex = 12
        Me._lblVentas_0.Text = "Familia"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_3, CType(3, Short))
        Me._lblVentas_3.Location = New System.Drawing.Point(16, 41)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_3.TabIndex = 14
        Me._lblVentas_3.Text = "Línea"
        '
        '_lblVentas_4
        '
        Me._lblVentas_4.AutoSize = True
        Me._lblVentas_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_4, CType(4, Short))
        Me._lblVentas_4.Location = New System.Drawing.Point(16, 66)
        Me._lblVentas_4.Name = "_lblVentas_4"
        Me._lblVentas_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_4.Size = New System.Drawing.Size(54, 13)
        Me._lblVentas_4.TabIndex = 16
        Me._lblVentas_4.Text = "SubLínea"
        '
        'fraRelojeria
        '
        Me.fraRelojeria.BackColor = System.Drawing.SystemColors.Control
        Me.fraRelojeria.Controls.Add(Me.dbcRModelo)
        Me.fraRelojeria.Controls.Add(Me.dbcRMarca)
        Me.fraRelojeria.Controls.Add(Me._lblVentas_5)
        Me.fraRelojeria.Controls.Add(Me._lblVentas_6)
        Me.fraRelojeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraRelojeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraRelojeria.Location = New System.Drawing.Point(88, 8)
        Me.fraRelojeria.Name = "fraRelojeria"
        Me.fraRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRelojeria.Size = New System.Drawing.Size(337, 89)
        Me.fraRelojeria.TabIndex = 29
        '
        'dbcRModelo
        '
        Me.dbcRModelo.Location = New System.Drawing.Point(78, 56)
        Me.dbcRModelo.Name = "dbcRModelo"
        Me.dbcRModelo.Size = New System.Drawing.Size(253, 21)
        Me.dbcRModelo.TabIndex = 19
        '
        'dbcRMarca
        '
        Me.dbcRMarca.Location = New System.Drawing.Point(78, 24)
        Me.dbcRMarca.Name = "dbcRMarca"
        Me.dbcRMarca.Size = New System.Drawing.Size(253, 21)
        Me.dbcRMarca.TabIndex = 18
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_5, CType(5, Short))
        Me._lblVentas_5.Location = New System.Drawing.Point(16, 24)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(37, 13)
        Me._lblVentas_5.TabIndex = 31
        Me._lblVentas_5.Text = "Marca"
        '
        '_lblVentas_6
        '
        Me._lblVentas_6.AutoSize = True
        Me._lblVentas_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_6, CType(6, Short))
        Me._lblVentas_6.Location = New System.Drawing.Point(16, 48)
        Me._lblVentas_6.Name = "_lblVentas_6"
        Me._lblVentas_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_6.Size = New System.Drawing.Size(42, 13)
        Me._lblVentas_6.TabIndex = 30
        Me._lblVentas_6.Text = "Modelo"
        '
        'optVarios
        '
        Me.optVarios.BackColor = System.Drawing.SystemColors.Control
        Me.optVarios.Cursor = System.Windows.Forms.Cursors.Default
        Me.optVarios.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optVarios.Location = New System.Drawing.Point(8, 72)
        Me.optVarios.Name = "optVarios"
        Me.optVarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optVarios.Size = New System.Drawing.Size(89, 17)
        Me.optVarios.TabIndex = 11
        Me.optVarios.TabStop = True
        Me.optVarios.Text = "Varios"
        Me.optVarios.UseVisualStyleBackColor = False
        '
        'optRelojeria
        '
        Me.optRelojeria.BackColor = System.Drawing.SystemColors.Control
        Me.optRelojeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRelojeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRelojeria.Location = New System.Drawing.Point(8, 48)
        Me.optRelojeria.Name = "optRelojeria"
        Me.optRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRelojeria.Size = New System.Drawing.Size(81, 17)
        Me.optRelojeria.TabIndex = 10
        Me.optRelojeria.TabStop = True
        Me.optRelojeria.Text = "Relojería"
        Me.optRelojeria.UseVisualStyleBackColor = False
        '
        'optJoyeria
        '
        Me.optJoyeria.BackColor = System.Drawing.SystemColors.Control
        Me.optJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.optJoyeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optJoyeria.Location = New System.Drawing.Point(8, 24)
        Me.optJoyeria.Name = "optJoyeria"
        Me.optJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optJoyeria.Size = New System.Drawing.Size(89, 17)
        Me.optJoyeria.TabIndex = 9
        Me.optJoyeria.TabStop = True
        Me.optJoyeria.Text = "Joyería"
        Me.optJoyeria.UseVisualStyleBackColor = False
        '
        'fraVarios
        '
        Me.fraVarios.BackColor = System.Drawing.SystemColors.Control
        Me.fraVarios.Controls.Add(Me.dbcVFamilia)
        Me.fraVarios.Controls.Add(Me.dbcVLinea)
        Me.fraVarios.Controls.Add(Me._lblVentas_7)
        Me.fraVarios.Controls.Add(Me._lblVentas_8)
        Me.fraVarios.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraVarios.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraVarios.Location = New System.Drawing.Point(88, 8)
        Me.fraVarios.Name = "fraVarios"
        Me.fraVarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraVarios.Size = New System.Drawing.Size(337, 89)
        Me.fraVarios.TabIndex = 26
        '
        'dbcVFamilia
        '
        Me.dbcVFamilia.Location = New System.Drawing.Point(78, 24)
        Me.dbcVFamilia.Name = "dbcVFamilia"
        Me.dbcVFamilia.Size = New System.Drawing.Size(253, 21)
        Me.dbcVFamilia.TabIndex = 20
        '
        'dbcVLinea
        '
        Me.dbcVLinea.Location = New System.Drawing.Point(78, 56)
        Me.dbcVLinea.Name = "dbcVLinea"
        Me.dbcVLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcVLinea.TabIndex = 21
        '
        '_lblVentas_7
        '
        Me._lblVentas_7.AutoSize = True
        Me._lblVentas_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_7, CType(7, Short))
        Me._lblVentas_7.Location = New System.Drawing.Point(16, 24)
        Me._lblVentas_7.Name = "_lblVentas_7"
        Me._lblVentas_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_7.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_7.TabIndex = 28
        Me._lblVentas_7.Text = "Familia"
        '
        '_lblVentas_8
        '
        Me._lblVentas_8.AutoSize = True
        Me._lblVentas_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_8, CType(8, Short))
        Me._lblVentas_8.Location = New System.Drawing.Point(16, 47)
        Me._lblVentas_8.Name = "_lblVentas_8"
        Me._lblVentas_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_8.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_8.TabIndex = 27
        Me._lblVentas_8.Text = "Línea"
        '
        'fraPeriodo
        '
        Me.fraPeriodo.BackColor = System.Drawing.SystemColors.Control
        Me.fraPeriodo.Controls.Add(Me.dtpFechaInicio)
        Me.fraPeriodo.Controls.Add(Me.dtpFechaFin)
        Me.fraPeriodo.Controls.Add(Me._Label1_0)
        Me.fraPeriodo.Controls.Add(Me.Label3)
        Me.fraPeriodo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraPeriodo.Location = New System.Drawing.Point(8, 56)
        Me.fraPeriodo.Name = "fraPeriodo"
        Me.fraPeriodo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraPeriodo.Size = New System.Drawing.Size(433, 57)
        Me.fraPeriodo.TabIndex = 3
        Me.fraPeriodo.TabStop = False
        Me.fraPeriodo.Text = " Período "
        '
        'dtpFechaInicio
        '
        Me.dtpFechaInicio.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicio.Location = New System.Drawing.Point(72, 20)
        Me.dtpFechaInicio.Name = "dtpFechaInicio"
        Me.dtpFechaInicio.Size = New System.Drawing.Size(99, 20)
        Me.dtpFechaInicio.TabIndex = 5
        '
        'dtpFechaFin
        '
        Me.dtpFechaFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFin.Location = New System.Drawing.Point(280, 20)
        Me.dtpFechaFin.Name = "dtpFechaFin"
        Me.dtpFechaFin.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaFin.TabIndex = 7
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(32, 23)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(34, 21)
        Me._Label1_0.TabIndex = 4
        Me._Label1_0.Text = "Del :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(248, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(33, 21)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Al :"
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(67, 21)
        Me.txtCodSucursal.MaxLength = 10
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(49, 20)
        Me.txtCodSucursal.TabIndex = 1
        Me.txtCodSucursal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(128, 21)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(299, 21)
        Me.dbcSucursales.TabIndex = 2
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(12, 24)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(73, 17)
        Me._Label1_1.TabIndex = 0
        Me._Label1_1.Text = "Sucursal :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(123, 319)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 91
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(8, 319)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 90
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(238, 320)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 89
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmRptKardexArticulo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(463, 374)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me._Frame1_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(292, 206)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRptKardexArticulo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Kardex - Movimientos por Artículos"
        Me._Frame1_0.ResumeLayout(False)
        Me._Frame1_0.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.fraJoyeria.ResumeLayout(False)
        Me.fraJoyeria.PerformLayout()
        Me.fraRelojeria.ResumeLayout(False)
        Me.fraRelojeria.PerformLayout()
        Me.fraVarios.ResumeLayout(False)
        Me.fraVarios.PerformLayout()
        Me.fraPeriodo.ResumeLayout(False)
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

End Class
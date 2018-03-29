'**********************************************************************************************************************'
'*PROGRAMA: ABC PROMOCIONES TARJETAS BANCARIAS JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018 
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualStudio.Data

Public Class frmCorpoAbcPromocionesTarjetasBanc
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: ABC a Promociones de Tarjetas Bancarias
    'Autor: Rosaura Torres López
    'Fecha de Creación: 23/Septiembre/2003

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lblActivados As System.Windows.Forms.Label
    Public WithEvents lblSuspendidos As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cmdActivarSuspender As System.Windows.Forms.Button
    Public WithEvents txtDetalle As System.Windows.Forms.TextBox
    Public WithEvents msgPromocion As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents txtDescPlan As System.Windows.Forms.Label
    Public WithEvents _lblPromoTarjetas_0 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblPromoTarjetas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray



    Const C_ColCODPLAN As Integer = 0
    Const C_ColDESCPLAN As Integer = 1
    Const C_ColPORCINTERES As Integer = 2
    Const C_ColPORCIVA As Integer = 3
    Const C_COLESTATUS As Integer = 4
    Dim mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Dim mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Dim mblnSALIR As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodBanco As Integer
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Dim i As Integer



    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoAbcPromocionesTarjetasBanc))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDescPlan = New System.Windows.Forms.Label()
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.lblActivados = New System.Windows.Forms.Label()
        Me.lblSuspendidos = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdActivarSuspender = New System.Windows.Forms.Button()
        Me.txtDetalle = New System.Windows.Forms.TextBox()
        Me.msgPromocion = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me._lblPromoTarjetas_0 = New System.Windows.Forms.Label()
        Me.lblPromoTarjetas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraGeneral.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.msgPromocion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPromoTarjetas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtDescPlan
        '
        Me.txtDescPlan.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescPlan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescPlan.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescPlan.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDescPlan.Location = New System.Drawing.Point(8, 186)
        Me.txtDescPlan.Name = "txtDescPlan"
        Me.txtDescPlan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescPlan.Size = New System.Drawing.Size(416, 21)
        Me.txtDescPlan.TabIndex = 5
        Me.txtDescPlan.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.txtDescPlan, "Descripción de Artículos")
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.Color.Silver
        Me.fraGeneral.Controls.Add(Me.Frame1)
        Me.fraGeneral.Controls.Add(Me.cmdActivarSuspender)
        Me.fraGeneral.Controls.Add(Me.txtDetalle)
        Me.fraGeneral.Controls.Add(Me.msgPromocion)
        Me.fraGeneral.Controls.Add(Me.dbcBanco)
        Me.fraGeneral.Controls.Add(Me.txtDescPlan)
        Me.fraGeneral.Controls.Add(Me._lblPromoTarjetas_0)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(14, 11)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(434, 277)
        Me.fraGeneral.TabIndex = 0
        Me.fraGeneral.TabStop = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Silver
        Me.Frame1.Controls.Add(Me.lblActivados)
        Me.Frame1.Controls.Add(Me.lblSuspendidos)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Frame1.Location = New System.Drawing.Point(208, 210)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(121, 57)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Estatus"
        '
        'lblActivados
        '
        Me.lblActivados.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblActivados.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblActivados.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblActivados.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblActivados.Location = New System.Drawing.Point(8, 16)
        Me.lblActivados.Name = "lblActivados"
        Me.lblActivados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblActivados.Size = New System.Drawing.Size(17, 17)
        Me.lblActivados.TabIndex = 11
        '
        'lblSuspendidos
        '
        Me.lblSuspendidos.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSuspendidos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSuspendidos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSuspendidos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSuspendidos.Location = New System.Drawing.Point(8, 35)
        Me.lblSuspendidos.Name = "lblSuspendidos"
        Me.lblSuspendidos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSuspendidos.Size = New System.Drawing.Size(17, 17)
        Me.lblSuspendidos.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Silver
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label2.Location = New System.Drawing.Point(35, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Activos"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Silver
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label3.Location = New System.Drawing.Point(35, 35)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(81, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Suspendidos"
        '
        'cmdActivarSuspender
        '
        Me.cmdActivarSuspender.BackColor = System.Drawing.SystemColors.Control
        Me.cmdActivarSuspender.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdActivarSuspender.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdActivarSuspender.Location = New System.Drawing.Point(344, 226)
        Me.cmdActivarSuspender.Name = "cmdActivarSuspender"
        Me.cmdActivarSuspender.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdActivarSuspender.Size = New System.Drawing.Size(81, 41)
        Me.cmdActivarSuspender.TabIndex = 6
        Me.cmdActivarSuspender.Text = "Suspender"
        Me.cmdActivarSuspender.UseVisualStyleBackColor = False
        '
        'txtDetalle
        '
        Me.txtDetalle.AcceptsReturn = True
        Me.txtDetalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtDetalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDetalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDetalle.Location = New System.Drawing.Point(200, 96)
        Me.txtDetalle.MaxLength = 0
        Me.txtDetalle.Name = "txtDetalle"
        Me.txtDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDetalle.Size = New System.Drawing.Size(65, 20)
        Me.txtDetalle.TabIndex = 3
        Me.txtDetalle.Visible = False
        '
        'msgPromocion
        '
        Me.msgPromocion.DataSource = Nothing
        Me.msgPromocion.Location = New System.Drawing.Point(8, 52)
        Me.msgPromocion.Name = "msgPromocion"
        Me.msgPromocion.OcxState = CType(resources.GetObject("msgPromocion.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgPromocion.Size = New System.Drawing.Size(416, 130)
        Me.msgPromocion.TabIndex = 4
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(56, 20)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(237, 21)
        Me.dbcBanco.TabIndex = 2
        '
        '_lblPromoTarjetas_0
        '
        Me._lblPromoTarjetas_0.AutoSize = True
        Me._lblPromoTarjetas_0.BackColor = System.Drawing.Color.Silver
        Me._lblPromoTarjetas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPromoTarjetas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPromoTarjetas_0.Location = New System.Drawing.Point(10, 22)
        Me._lblPromoTarjetas_0.Name = "_lblPromoTarjetas_0"
        Me._lblPromoTarjetas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPromoTarjetas_0.Size = New System.Drawing.Size(44, 13)
        Me._lblPromoTarjetas_0.TabIndex = 1
        Me._lblPromoTarjetas_0.Text = "Banco :"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraGeneral)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(458, 378)
        Me.Panel1.TabIndex = 1
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(14, 292)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(434, 74)
        Me.Panel3.TabIndex = 72
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(208, 14)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(160, 14)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(50, 42)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = " "
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.grabar
        Me.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnGuardar.Location = New System.Drawing.Point(11, 14)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(110, 14)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(50, 42)
        Me.btnLimpiar.TabIndex = 66
        Me.btnLimpiar.Text = " "
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Eliminar
        Me.btnEliminar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEliminar.Location = New System.Drawing.Point(61, 14)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'frmCorpoAbcPromocionesTarjetasBanc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(481, 401)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(177, 160)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoAbcPromocionesTarjetasBanc"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC  a  Promociones Tarjetas Bancarias"
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.msgPromocion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPromoTarjetas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub



    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
    End Sub

    'Sub Buscar()
    ''Esta Función se utilizará para Buscar un dato especifico de un formulario, la cual podrá realizarse por campo Codigo o Campo Descripción,
    '' y se Activará presionando la tecla F3.
    '    On Local Error GoTo MErr
    '    Dim strSQL As String
    '    Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
    '    Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
    '    Dim strControlActual As String 'Nombre del control actual
    '    Dim TipoTaller As String 'Tipo de Taller seleccionado segun los Option Button J-R-F
    '
    '    strControlActual = UCase(Screen.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
    '    strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
    '
    '    If optJoyeria.Value = True Then
    '        TipoTaller = "J"
    '    Else
    '        If optRelojeria.Value = True Then
    '            TipoTaller = "R"
    '        Else
    '            If optForaneo.Value = True Then
    '                TipoTaller = "F"
    '            Else
    '                TipoTaller = ""
    '            End If
    '        End If
    '    End If
    '
    '
    '    Select Case strControlActual
    '        Case "TXTCODTALLER"
    '            strCaptionForm = "Consulta de Talleres"
    '            If chkMostrarTodos.Value = Checked Then
    '                gStrSql = "SELECT RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,DescTaller AS DESCRIPCION,Responsable as  RESPONSABLE, " & _
    ''                        "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & _
    ''                        "FROM CatTalleres ORDER BY CodTaller "
    '            Else
    '                gStrSql = "SELECT RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,DescTaller AS DESCRIPCION,Responsable as  RESPONSABLE, " & _
    ''                        "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & _
    ''                        "FROM CatTalleres " & _
    ''                        "WHERE " & IIf((Trim(TipoTaller) <> ""), "TipoTaller= '" & TipoTaller & "'", "TipoTaller LIKE '%'") & _
    ''                        "ORDER BY CodTaller "
    '            End If
    '        Case "TXTDESCRIPCION"
    '            strCaptionForm = "Consulta de Talleres"
    '            If chkMostrarTodos.Value = vbChecked Then
    '                gStrSql = "SELECT DescTaller AS DESCRIPCION, RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,Responsable as  RESPONSABLE, " & _
    ''                    "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & _
    ''                    "FROM CatTalleres WHERE DescTaller LIKE '" & txtDescripcion & "%' " & _
    ''                    " ORDER BY DescTaller"
    '            Else
    '                gStrSql = "SELECT DescTaller AS DESCRIPCION, RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,Responsable as  RESPONSABLE, " & _
    ''                    "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & _
    ''                    "FROM CatTalleres WHERE DescTaller LIKE '" & txtDescripcion & "%' " & _
    ''                    " AND " & IIf((Trim(TipoTaller) <> ""), " TipoTaller= '" & TipoTaller & "'", "TipoTaller LIKE '%'") & _
    ''                    " ORDER BY DescTaller"
    '            End If
    '        Case Else
    '            'Sale de este sub para ke no ejecute ninguna opcion
    '            Exit Sub
    '    End Select
    '
    '    strSQL = gStrSql 'Se hace uso de una variable temporal para el query
    '
    '    'Si hubo cambios y es una modificacion entonces preguntara que si desea grabar los cambios
    '    If Cambios = True And mblnNuevo = False Then
    '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
    '            Case vbYes: 'Guardar el registro
    '                If Guardar = False Then
    '                    Exit Sub
    '                End If
    '            Case vbNo: 'No hace nada y permite que se cargue la consulta
    '            Case vbCancel: 'Cancela la consulta
    '                Exit Sub
    '        End Select
    '    End If
    '
    '    gStrSql = strSQL 'Se regresa el valor de la variavle temporal a la variable original
    '
    '    ModEstandar.BorraCmd
    '    Cmd.CommandText = "dbo.Up_Select_Datos"
    '    Cmd.CommandType = adCmdStoredProc
    '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '    Set RsGral = Cmd.Execute
    '
    '    'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
    '    If RsGral.RecordCount = 0 Then
    '        MsgBox C_msgSINDATOS & vbNewLine & "Verifique por favor...", vbExclamation, gstrNombCortoEmpresa
    '        RsGral.Close
    '        Exit Sub
    '    End If
    '
    '    'Carga el formulario de consulta
    '    Load FrmConsultas
    '    Call ConfiguraConsultas(FrmConsultas, 10200, RsGral, strTag, strCaptionForm)
    '    With FrmConsultas.Flexdet
    '        Select Case strControlActual
    '            Case "TXTCODTALLER"
    '                .ColWidth(0) = 900 'Columna del Código
    '                .ColWidth(1) = 3800 'Columna de la Descripción
    '                .ColWidth(2) = 4500 'Columna del Nombre del Responsable
    '                .ColWidth(3) = 1000 'Columna del Tipo de Taller
    '            Case "TXTDESCRIPCION"
    '                .ColWidth(0) = 3800 'Columna de la Descripción
    '                .ColWidth(1) = 900 'Columna del Código
    '                .ColWidth(2) = 4500 'Columna del Nombre del Responsable
    '                .ColWidth(3) = 1000 'Columna del Tipo de Taller
    '        End Select
    '    End With
    '    ModEstandar.CentrarForma FrmConsultas
    '    FrmConsultas.Show vbModal
    '
    'MErr:
    '    If Err.Number <> 0 Then ModEstandar.MostrarError
    'End Sub
    'Sub Eliminar()
    '    On Local Error GoTo MErr
    ''    Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción
    '
    '    gStrSql = "SELECT DescTaller FROM CatTalleres WHERE CodTaller=" & Val(txtCodTaller)
    '
    '    ModEstandar.BorraCmd
    '    Cmd.CommandText = "dbo.Up_Select_Datos"
    '    Cmd.CommandType = adCmdStoredProc
    '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '    Set RsGral = Cmd.Execute
    '
    '    If RsGral.RecordCount = 0 Then
    '        MsgBox "Proporcione un Código valido para eliminar.", vbExclamation + vbOKOnly, "Mensaje"
    '        Cnn.RollbackTrans
    '        RsGral.Close
    '        Exit Sub
    '    End If
    '
    '    'Preguntar si desea borrar el registro
    '    Select Case MsgBox(C_msgBORRAR, vbQuestion + vbYesNoCancel + vbDefaultButton2, "Mensaje")
    '        Case vbNo
    '          Exit Sub
    '        Case vbCancel
    '          Exit Sub
    '    End Select
    '    Screen.MousePointer = vbHourglass
    '    'El parametro TipoTaller no es requerido en la eliminación, por tanto le estoy mandando un Valor Fijo ("O")
    '    Cnn.BeginTrans
    '
    '    ModStoredProcedures.PR_IMECatTalleres Trim(txtCodTaller), Trim(txtDescripcion), Trim(txtResponsable), Trim(txtDomicilio), "O", C_ELIMINACION, 0
    '    Cmd.Execute
    '
    '    Cnn.CommitTrans
    '    Nuevo
    '    Limpiar
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'MErr:
    '    Cnn.RollbackTrans
    '    Screen.MousePointer = vbDefault
    '    If Err.Number <> 0 Then ModEstandar.MostrarError
    'End Sub

    Function Guardar() As Boolean
        On Error GoTo MErr
        Dim DescPlan As String
        Dim CodPlan As Integer
        Dim PorcIntereses As Decimal
        Dim PorcIva As Decimal
        Dim Estatus As String
        Dim FilaInicio As Integer

        ' Especifica si se inicia dar de alta desde la Final Uno o hasta la Dios, dependiendo, si ya existe un pla
        If ValidaDatos() = False Then Exit Function
        'If optVigente = True Then
        '    Estatus = "V"
        'ElseIf optSuspendido = True Then
        '    Estatus = "S"
        'End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Cnn.BeginTrans()

        ModStoredProcedures.PR_IMECatPlanesxBanco(CStr(intCodBanco), CStr(0), "", CStr(0), CStr(0), "", C_ELIMINACION, CStr(0))
        Cmd.Execute()

        '    ModStoredProcedures.PR_IMECatPlanesxBanco CStr(intCodBanco), CStr(1), CStr("SIN PROMOCION"), CStr(0), CStr(0), Trim(Estatus), C_INSERCION, 0
        '    Cmd.Execute

        With msgPromocion
            'FilaInicio = IIf((Numerico(.TextMatrix(1, C_ColCODPLAN)) = 1), 2, 1)


            For i = 1 To .Rows - 1
                If Trim(.get_TextMatrix(i, C_ColDESCPLAN)) = "" Then Exit For
                '            If Numerico(.TextMatrix(i, C_ColCODPLAN)) <> 1 Then
                '                If FilaInicio = 1 Then
                '                    CodPlan = i + 1
                '                Else
                CodPlan = i
                '                End If
                DescPlan = Trim(.get_TextMatrix(i, C_ColDESCPLAN))
                PorcIntereses = CDec(Numerico(Trim(.get_TextMatrix(i, C_ColPORCINTERES))))
                PorcIva = CDec(Numerico(Trim(.get_TextMatrix(i, C_ColPORCIVA))))
                Estatus = Trim(.get_TextMatrix(i, C_COLESTATUS))
                ModStoredProcedures.PR_IMECatPlanesxBanco(CStr(intCodBanco), CStr(CodPlan), Trim(DescPlan), CStr(PorcIntereses), CStr(PorcIva), Trim(Estatus), C_INSERCION, CStr(0))
                Cmd.Execute()
                '            End If
            Next
        End With

        Cnn.CommitTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        MsgBox("Los planes para el banco han sido grabados correctamente. ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")

        Guardar = True
        Nuevo()
        Limpiar()
        Exit Function
MErr:
        Cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Nuevo()
        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llave principal
        On Error GoTo MErr
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        msgPromocion.Clear()
        Encabezado()
        FueraChange = True
        txtDescPlan.Text = ""
        intCodBanco = 0
        FueraChange = False
        'optVigente.Value = True
        'System.Windows.Forms.Cursor.Current = False
        Exit Sub
MErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatos()
        On Error GoTo MErr

        gStrSql = "SELECT     P.CodBanco, B.DescBanco, P.CodPlan, P.DescPlan, P.PorcIntereses, P.PorcIva, P.Estatus " & "FROM         dbo.CatPlanesxBanco P INNER JOIN " & "dbo.CatBancos B ON P.CodBanco = B.CodBanco " & "Where (P.CodBanco =" & intCodBanco & ")"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            'If Trim(RsGral!Estatus) = "V" Then
            '    optVigente.Value = True
            'ElseIf Trim(RsGral!Estatus) = "S" Then
            '    optSuspendido.Value = True
            'End If
            With msgPromocion
                For i = 1 To RsGral.RecordCount
                    .set_TextMatrix(i, C_ColCODPLAN, RsGral.Fields("CodPlan").Value)
                    .set_TextMatrix(i, C_ColDESCPLAN, RsGral.Fields("DescPlan").Value)
                    .set_TextMatrix(i, C_ColPORCINTERES, Format(RsGral.Fields("PorcIntereses").Value, gstrFormatoCantidad))
                    .set_TextMatrix(i, C_ColPORCIVA, Format(RsGral.Fields("PorcIva").Value, gstrFormatoCantidad))
                    .set_TextMatrix(i, C_COLESTATUS, Trim(RsGral.Fields("Estatus").Value))
                    If Trim(RsGral.Fields("Estatus").Value) = "V" Then
                        PonerColor(lblActivados, i)
                    ElseIf Trim(RsGral.Fields("Estatus").Value) = "S" Then
                        PonerColor(lblSuspendidos, i)
                    End If
                    RsGral.MoveNext()
                Next
                .set_ColAlignment(C_ColDESCPLAN, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                .Col = 0
                .Row = 1
            End With
            mblnCambiosEnCodigo = False
            mblnNuevo = False
        End If
        Exit Sub
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        'Esta función Limpia todos los controles del formulario.
        'Si hubo Cambios, Pregunta si desea guardarlos.
        On Error GoTo MErr
        '    Screen.MousePointer = vbHourglass
        If Cambios() = True And mblnNuevo = False Then 'Si hubo Cambios y se trata de una consulta se hace lo siguiente
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Permite Guardar los cambios en el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No
                    'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar la pantalla
                    Exit Sub
            End Select
        End If

        Nuevo()
        dbcBanco.Text = ""
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        dbcBanco.Focus()
        Exit Sub
MErr:
        '    Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    Function Cambios() As Object
        '    'Esta Función validará si se han efectuado cambios en los controles.
        '    'lo cual es útil para la funcion de guardar. Se inicializa con True, y si se validan todos los campos y no se ha
        '    'salido del proc. entonces la variable adquiere el valor de False
        '    'se validan todos los controles existentes, excepto el de la Clave Principal
        '    On Local Error GoTo MErr
        '    '    Screen.MousePointer = vbHourglass
        '    Cambios = True
        '    If Trim(txtDescripcion) <> Trim(txtDescripcion.Tag) Then Exit Function
        '    If Trim(txtResponsable) <> Trim(txtResponsable.Tag) Then Exit Function
        '    If Trim(txtDomicilio) <> Trim(txtDomicilio.Tag) Then Exit Function
        '    If optJoyeria.Value <> optJoyeria.Tag Then Exit Function
        '    If optRelojeria.Value <> optRelojeria.Tag Then Exit Function
        '    If optForaneo.Value <> optForaneo.Tag Then Exit Function
        '    Cambios = False
        '    '    Screen.MousePointer = vbDefault
        '    Exit Function
        'MErr:
        '    '    Screen.MousePointer = vbDefault
        '    If Err.Number <> 0 Then ModEstandar.MostrarError
    End Function

    Function ValidaDatos() As Object
        On Error GoTo MErr
        Dim GridSinDAtos As Boolean
        GridSinDAtos = True
        If Trim(dbcBanco.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Nombre del Banco", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            dbcBanco.Focus()
            Exit Function
        End If
        With msgPromocion
            For i = 1 To .Rows - 1
                If Trim(.get_TextMatrix(.Row, C_ColDESCPLAN)) <> "" Then
                    GridSinDAtos = False
                End If
            Next
        End With
        If GridSinDAtos = True Then
            'MsgBox C_msgFALTADATO & "Detalle de promociones", vbExclamation, gstrNombCortoEmpresa
            msgPromocion.Focus()
            Exit Function
        End If
        ValidaDatos = True
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub cmdActivarSuspender_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdActivarSuspender.Click
        If Trim(msgPromocion.get_TextMatrix(msgPromocion.Row, 4)) = "S" Then
            PonerColor(lblActivados, msgPromocion.Row)
            msgPromocion.set_TextMatrix(msgPromocion.Row, 4, "V")
            cmdActivarSuspender.Text = "Suspender"
        ElseIf Trim(msgPromocion.get_TextMatrix(msgPromocion.Row, 4)) = "V" Then
            PonerColor(lblSuspendidos, msgPromocion.Row)
            msgPromocion.set_TextMatrix(msgPromocion.Row, 4, "S")
            cmdActivarSuspender.Text = "Activar"
        End If
    End Sub

    Private Sub frmCorpoAbcPromocionesTarjetasBanc_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoAbcPromocionesTarjetasBanc_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoAbcPromocionesTarjetasBanc_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        Nuevo()
    End Sub

    Private Sub frmCorpoAbcPromocionesTarjetasBanc_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> msgPromocion.Name Then
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmCorpoAbcPromocionesTarjetasBanc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoAbcPromocionesTarjetasBanc_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSALIR Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    'Si se cierra el formulario y existio algun cambio en el registro se
        '    'informa al usuario del cabio y si desea guardar el registro, ya sea
        '    'que sea nuevo o un registro modificado
        '    If Cambios() = True Then ' And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes 'Guardar el registro
        '                If Guardar() = False Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else 'Se quiere salir con escape
        '    mblnSALIR = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoAbcPromocionesTarjetasBanc_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        frmCorpoAbcTalleres = Nothing
    End Sub



    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcBanco" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,Ltrim(Rtrim(DescBanco)) as DescBanco FROM CatBancos WHERE ControlInterno = 0 ORDER BY DescBanco"
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodBanco,Ltrim(Rtrim(DescBanco)) as DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 0 ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
        LlenaDatos()
    End Sub

    Sub Encabezado()
        With msgPromocion
            .set_ColWidth(C_ColPORCIVA, 0, 1200)
            .set_ColWidth(C_COLESTATUS, 0, 0)
            .set_TextMatrix(0, C_ColCODPLAN, "CodPlan")
            .set_TextMatrix(0, C_ColDESCPLAN, "Plan")
            .set_TextMatrix(0, C_ColPORCINTERES, "% Interés")
            .set_TextMatrix(0, C_ColPORCIVA, "% IVA")

            .Row = 0
            For i = 0 To C_ColPORCIVA
                .Col = i
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
            Next
        End With

        If Err.Number <> 0 Then ModEstandar.MostrarError()

    End Sub

    Sub PonerColor(ByRef lblcolor As System.Windows.Forms.Control, ByRef Ren As Integer)
        Dim i As Integer
        With msgPromocion
            For i = 0 To 3
                .Col = i
                .Row = Ren
                .CellBackColor = System.Drawing.ColorTranslator.FromOle(lblcolor.BackColor.A)
            Next
            .Col = 0
        End With
    End Sub

    Private Sub msgpromocion_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgPromocion.DblClick
        msgpromocion_KeyPressEvent(msgPromocion, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub msgpromocion_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgPromocion.EnterCell
        With msgPromocion
            If Trim(.get_TextMatrix(.Row, 0)) = "" And Trim(.get_TextMatrix(.Row, 1)) = "" And Trim(.get_TextMatrix(.Row, 2)) = "" And Trim(.get_TextMatrix(.Row, 3)) = "" Then
                cmdActivarSuspender.Enabled = False
            End If
            txtDescPlan.Text = Trim(.get_TextMatrix(.Row, C_ColDESCPLAN))
            If Trim(.get_TextMatrix(.Row, C_COLESTATUS)) = "V" Then
                cmdActivarSuspender.Text = "Suspender"
                cmdActivarSuspender.Enabled = True
            ElseIf Trim(.get_TextMatrix(.Row, C_COLESTATUS)) = "S" Then
                cmdActivarSuspender.Text = "Activar"
                cmdActivarSuspender.Enabled = True
            End If
        End With
    End Sub

    Private Sub msgpromocion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgPromocion.Enter
        msgPromocion.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        msgpromocion_EnterCell(msgPromocion, New System.EventArgs())
        Pon_Tool()
    End Sub

    Private Sub msgpromocion_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgPromocion.KeyDownEvent
        'Aqui debe cvalidarse el movimiento de teclas, si es ke se va a tomar en cuenta
        With msgPromocion
            Select Case eventArgs.keyCode
                Case System.Windows.Forms.Keys.Delete
                    If .Row = 0 Then Exit Sub
                    If .get_TextMatrix(.Row, C_ColDESCPLAN) <> "" Then
                        .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                        BorraGrid(.Row) 'Cuando se Borra, se obtienen los nuevos totales, la cntidad de Articulos (Dento del proc.)                    .Col = C_ColCodplan
                        .Focus()
                        msgpromocion_EnterCell(msgPromocion, New System.EventArgs())
                    End If
            End Select
        End With
    End Sub

    Sub BorraGrid(ByRef Row As Integer)
        'Este Procediento borra un renglon del Grid
        'Si el Número de Filas que kedan en el grid, es menor de 8, se insertará una nueva fila al final del grid
        With msgPromocion
            .RemoveItem(Row)
            'Si el número de filas es menor de 10 o esta posicionado en la utlima fila, entonces, agrega una fila
            If .Rows < 11 Or .Row = .Rows - 1 Then
                .AddItem("")
                .Row = .Row
            End If
        End With
        'Al borrar se deben obtener los nuevo totales y Actualizar la cantidad de articulos
        '    ObtenerTotalDetalle
        '    ObtenerCantidadArticulos
    End Sub


    Private Sub msgpromocion_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgPromocion.KeyPressEvent
        Dim PorcDescuentoAnt As Decimal
        Dim ImpteDescuentoAnt As Decimal
        Dim TipoDescto As String '"I" "P"
        Dim ColSiguiente As Integer
        Dim rowsiguiente As Integer
        With msgPromocion
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
                Select Case .Col

                    Case C_ColDESCPLAN ''-------------- SE EDITA LA DESCRIPCION ---------------------'''''
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                        If (.Row > 1) Then
                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
                            If Trim(.get_TextMatrix(.Row - 1, C_ColDESCPLAN)) = "" Then
                                .Focus()
                                Exit Sub
                            End If
                        End If
                        'ModEstandar.MSHFlexGridEdit(msgPromocion, txtDetalle, eventArgs.keyAscii)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then
                            ModEstandar.SelTextoTxt(txtDetalle)
                        End If

                    Case C_ColPORCINTERES
                        eventArgs.keyAscii = ModEstandar.MskCantidad(txtDetalle.Text, eventArgs.keyAscii, 3, 2, (txtDetalle.SelectionStart))
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        'Validar que en el codigo y en la descripcion exista valor para editar este campo
                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColCODPLAN))) = 0 And .get_TextMatrix(.Row, C_ColDESCPLAN) = "" Then
                            .Focus()
                            Exit Sub
                        End If
                        ModEstandar.MSHFlexGridEdit(msgPromocion, txtDetalle, eventArgs.keyAscii)
                        txtDetalle.SelectionStart = Len(txtDetalle.Text)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then
                            ModEstandar.SelTextoTxt(txtDetalle)
                        End If

                    Case C_ColPORCIVA

                        eventArgs.keyAscii = ModEstandar.MskCantidad(txtDetalle.Text, eventArgs.keyAscii, 3, 2, (txtDetalle.SelectionStart))
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        'Validar que en el codigo y en la descripcion exista valor para editar este campo
                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColCODPLAN))) = 0 And .get_TextMatrix(.Row, C_ColDESCPLAN) = "" Then
                            .Focus()
                            Exit Sub
                        End If
                        ModEstandar.MSHFlexGridEdit(msgPromocion, txtDetalle, eventArgs.keyAscii)
                        txtDetalle.SelectionStart = Len(txtDetalle.Text)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then
                            ModEstandar.SelTextoTxt(txtDetalle)
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub msgpromocion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgPromocion.Leave
        msgPromocion.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub msgPromocion_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgPromocion.Scroll
        txtDetalle.Visible = False
    End Sub


    Private Sub txtDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Enter
        txtDetalle.Text = Trim(txtDetalle.Text)
        Pon_Tool()
    End Sub

    Private Sub txtDetalle_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDetalle.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Aqui se muestran los datos del control editable, en el Grid
        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
        Dim rowsiguiente As Integer
        Dim ColSiguiente As Integer
        Dim FormatoCantidad As String
        With msgPromocion
            Select Case KeyCode

                Case System.Windows.Forms.Keys.Escape
                    .Focus()
                    txtDetalle.Visible = False
                    txtDetalle.Text = ""
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                    .Focus()
                Case System.Windows.Forms.Keys.Return
                    If .Col = C_ColPORCINTERES Or .Col = C_ColPORCIVA Then
                        .set_TextMatrix(.Row, .Col, Format(Numerico(Trim(txtDetalle.Text)), gstrFormatoCantidad))
                    Else
                        If Trim(txtDetalle.Text) <> "" Then
                            .set_TextMatrix(.Row, .Col, Trim(txtDetalle.Text))
                            .set_TextMatrix(.Row, 4, "V")
                            PonerColor(lblActivados, .Row)
                        Else
                            .set_TextMatrix(.Row, 0, "")
                            .set_TextMatrix(.Row, .Col, "")
                            .set_TextMatrix(.Row, 2, "")
                            .set_TextMatrix(.Row, 3, "")
                            .set_TextMatrix(.Row, 4, "")
                            PonerColor(msgPromocion, .Row)
                            cmdActivarSuspender.Enabled = False
                        End If
                    End If
                    FueraChange = True
                    txtDetalle.Text = ""
                    txtDetalle.Visible = False
                    msgPromocion.Col = .Col
                    msgPromocion.Row = .Row
                    .Focus()
                    Exit Sub
                    If .Col = C_ColCODPLAN Then
                    ElseIf .Col = C_ColDESCPLAN Then
                        .set_ColAlignment(C_ColDESCPLAN, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        rowsiguiente = .Row
                        ColSiguiente = C_ColPORCINTERES
                    ElseIf .Col = C_ColPORCINTERES Then
                        If Trim(.get_TextMatrix(.Row, C_ColDESCPLAN)) = "" Then
                            MsgBox("La Cantidad mínima debe ser 1." & vbNewLine & "Verfique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AVISO")
                            .Focus()
                            msgpromocion_KeyPressEvent(msgPromocion, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
                            Exit Sub
                        End If
                        rowsiguiente = .Row
                        ColSiguiente = C_ColPORCIVA
                        '.TextMatrix(.Row, C_ColIMPORTE) = ObtenerImporteArticulo
                    ElseIf .Col = C_ColPORCIVA Then
                        rowsiguiente = .Row + 1
                        ColSiguiente = C_ColDESCPLAN
                        .Row = rowsiguiente
                        .Col = ColSiguiente
                    End If
                    .Row = rowsiguiente
                    .Col = ColSiguiente
            End Select
        End With
    End Sub

    Private Sub txtDetalle_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDetalle.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'En este Evento se validan los datos que se introduzcan al control txtDetalle,dependiendo de la columan en que se esté editando
        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
        With msgPromocion
            If .Col = C_ColPORCINTERES Then
                KeyAscii = ModEstandar.MskCantidad(txtDetalle.Text, KeyAscii, 3, 2, (txtDetalle.SelectionStart))
            End If
            If .Col = C_ColPORCIVA Then
                KeyAscii = ModEstandar.MskCantidad(txtDetalle.Text, KeyAscii, 3, 2, (txtDetalle.SelectionStart))
            End If
        End With
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Leave
        txtDetalle.Visible = False
    End Sub

    Private Sub dbcBanco_KeyDown(sender As Object, e As KeyEventArgs) Handles dbcBanco.KeyDown
        'If sender = System.Windows.Forms.Keys.Enter Then
        'End If
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub
End Class
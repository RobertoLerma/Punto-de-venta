Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmReimpresionCorteFinal
    Inherits System.Windows.Forms.Form
    'Programa: Reimpresión del COrte Finale de Caja
    'Autor: Rosaura Torres López
    'Fecha de Creación: 24/Julio/2003

    'Antonella Vargas - 04/MAY/2004


    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents dtpFechaCorte As System.Windows.Forms.DateTimePicker
    Public WithEvents dbcCaja As System.Windows.Forms.ComboBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents fraPeriodo As System.Windows.Forms.GroupBox
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mblnSalir As Boolean
    Dim tecla As Integer
    Dim intCodCaja As Integer
    Dim RsReimp As ADODB.Recordset

    Dim mintCodSucursal As Integer
    Dim mblnFueraChange As Boolean

    '''Private Sub chkTodas_Click()
    '''    Select Case chkTodas.Value
    '''        Case vbChecked
    '''            mblnFueraChange = True
    '''            dbcSucursal.text = "[ Todas ... ]"
    '''            dbcSucursal.Tag = ""
    '''            mintCodSucursal = 0
    '''            dbcSucursal.Enabled = False
    '''            mblnFueraChange = False
    '''        Case Else
    '''            mblnFueraChange = True
    '''            dbcSucursal.text = ""
    '''            dbcSucursal.Tag = ""
    '''            mintCodSucursal = 0
    '''            dbcSucursal.Enabled = True
    '''            mblnFueraChange = False
    '''    End Select
    '''End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraPeriodo = New System.Windows.Forms.GroupBox()
        Me.dtpFechaCorte = New System.Windows.Forms.DateTimePicker()
        Me.dbcCaja = New System.Windows.Forms.ComboBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.fraPeriodo.SuspendLayout()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraPeriodo
        '
        Me.fraPeriodo.BackColor = System.Drawing.SystemColors.Control
        Me.fraPeriodo.Controls.Add(Me.dtpFechaCorte)
        Me.fraPeriodo.Controls.Add(Me.dbcCaja)
        Me.fraPeriodo.Controls.Add(Me.dbcSucursal)
        Me.fraPeriodo.Controls.Add(Me._lblVentas_0)
        Me.fraPeriodo.Controls.Add(Me.Label3)
        Me.fraPeriodo.Controls.Add(Me._lblVentas_1)
        Me.fraPeriodo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraPeriodo.Location = New System.Drawing.Point(6, 6)
        Me.fraPeriodo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.fraPeriodo.Name = "fraPeriodo"
        Me.fraPeriodo.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.fraPeriodo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraPeriodo.Size = New System.Drawing.Size(263, 119)
        Me.fraPeriodo.TabIndex = 0
        Me.fraPeriodo.TabStop = False
        '
        'dtpFechaCorte
        '
        Me.dtpFechaCorte.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFechaCorte.Location = New System.Drawing.Point(76, 57)
        Me.dtpFechaCorte.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dtpFechaCorte.Name = "dtpFechaCorte"
        Me.dtpFechaCorte.Size = New System.Drawing.Size(125, 20)
        Me.dtpFechaCorte.TabIndex = 5
        '
        'dbcCaja
        '
        Me.dbcCaja.Location = New System.Drawing.Point(76, 88)
        Me.dbcCaja.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dbcCaja.Name = "dbcCaja"
        Me.dbcCaja.Size = New System.Drawing.Size(103, 21)
        Me.dbcCaja.TabIndex = 6
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(66, 25)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(160, 21)
        Me.dbcSucursal.TabIndex = 2
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_0, CType(0, Integer))
        Me._lblVentas_0.Location = New System.Drawing.Point(11, 28)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(51, 13)
        Me._lblVentas_0.TabIndex = 1
        Me._lblVentas_0.Text = "Sucursal:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(25, 62)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(70, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Fecha :"
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblVentas.SetIndex(Me._lblVentas_1, CType(1, Integer))
        Me._lblVentas_1.Location = New System.Drawing.Point(35, 92)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(34, 13)
        Me._lblVentas_1.TabIndex = 4
        Me._lblVentas_1.Text = "Caja :"
        '
        'frmReimpresionCorteFinal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(278, 140)
        Me.Controls.Add(Me.fraPeriodo)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(289, 216)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReimpresionCorteFinal"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reimpresión de Cortes"
        Me.fraPeriodo.ResumeLayout(False)
        Me.fraPeriodo.PerformLayout()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)

        If Trim(dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P'"
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            eventSender.KeyCode = 0
            Exit Sub
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Name Then Exit Sub
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus(dbcSucursal, gStrSql, mintCodSucursal)
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = Trim(dbcSucursal.Text)
        'If dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        dbcSucursal.Text = Aux
    End Sub

    Private Sub dbcCaja_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCaja.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcCaja.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodCaja, RIGHT('00'+LTRIM(CodCaja),2) as NumCaja From dbo.CatCajas Where CodCaja like '" & Trim(dbcCaja.Text) & "%'   And CodAlmacen =" & gintCodAlmacen & " order by NumCAja asc "
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcCaja_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCaja.Enter
        Pon_Tool()
        gStrSql = "SELECT CodCaja, RIGHT('00'+LTRIM(CodCaja),2) as NumCaja From dbo.CatCajas  Where codAlmacen= " & gintCodAlmacen & "  order by NumCAja asc "
        ModDCombo.DCGotFocus(gStrSql, dbcCaja)
    End Sub

    Private Sub dbcCaja_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCaja.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Private Sub dbcCaja_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCaja.Leave
        intCodCaja = 0
        gStrSql = "SELECT CodCaja, RIGHT('00'+LTRIM(CodCaja),2) as NumCaja From dbo.CatCajas Where CodCaja like '" & Numerico(dbcCaja.Text) & "%'   and codAlmacen= " & gintCodAlmacen & "    order by NumCAja asc "
        ModDCombo.DCLostFocus(dbcCaja, gStrSql, intCodCaja)
        'LlenaDatosCliente()
    End Sub

    Private Sub frmReimpresionCorteFinal_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                           Nuevo        Guardar       Cancelar       Eliminar      Buscar       Imprimir      Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmReimpresionCorteFinal_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                           Nuevo        Guardar       Cancelar       Eliminar      Buscar       Imprimir      Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub Form_Initialize_Renamed()
        '                           Nuevo        Guardar       Cancelar       Eliminar      Buscar       Imprimir      Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmReimpresionCorteFinal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmReimpresionCorteFinal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()
        dtpFechaCorte.Value = Format(Today, "dd/MMM/yyyy")
        mblnFueraChange = True
        dbcCaja.Text = ""
        mblnFueraChange = False
        mintCodSucursal = 0
    End Sub

    Private Sub frmReimpresionCorteFinal_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                           Nuevo        Guardar       Cancelar       Eliminar      Buscar       Imprimir      Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
    End Sub

    Private Sub frmReimpresionCorteFinal_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If Not mblnSalir Then
            'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
            ModEstandar.RestaurarForma(Me, False)
            Cancel = 0 'Para que no salga del Formulario hasta que guarde los datos, si no tiene premiso de hacerlo
        Else 'Se quiere salir con escape
            mblnSalir = False
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes 'Sale del Formulario, pero antes preguntar si desea grabar los datos registrados, solo cuando es nuevo
                    Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
                Case MsgBoxResult.No 'No sale del formulario
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmReimpresionCorteFinal_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo        Guardar       Cancelar       Eliminar      Buscar       Imprimir      Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
    End Sub

    Sub Imprime()
        On Error GoTo Merr
        If ValidaDatos() = False Then Exit Sub

        gStrSql = "Select * from CorteZ Where FechaCorte = '" & Format(dtpFechaCorte.Value, C_FORMATFECHAGUARDAR) & "' and CodSucursal = " & mintCodSucursal & " and  CodCAja= " & CShort(dbcCaja.Text) & ""

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsReimp = Cmd.Execute

        If RsReimp.RecordCount = 0 Then
            MsgBox("No Existe Información Almacenada que corresponda a los Datos Proporcionados." & vbNewLine & "Verifique Por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            dtpFechaCorte.Focus()
            Exit Sub
        Else
            ModCorporativo.TicketCorteFinalDia(RsReimp.Fields("FechaCorte").Value, mintCodSucursal, CShort(Numerico((dbcCaja.Text))), "F", True)
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub


    Function ValidaDatos() As Boolean
        On Error GoTo Merr
        ValidaDatos = False
        If dtpFechaCorte.Value > Today Then
            MsgBox("La Fecha de Corte debe ser Menor que la Fecha Actual." & vbNewLine & "Verifique Por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            dtpFechaCorte.Focus()
            Exit Function
        End If
        If Trim(dbcCaja.Text) = "" Then
            MsgBox("Proporcione el Número de Caja.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            dbcCaja.Focus()
            Exit Function
        End If
        ValidaDatos = True
        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function
End Class
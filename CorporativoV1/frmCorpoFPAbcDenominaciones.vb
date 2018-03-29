Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCorpoFPAbcDenominaciones

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnAceptar As System.Windows.Forms.Button
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents FlexDenominaciones As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Const C_ColDENOMINACION As Integer = 0
    Const C_ColDENOMINACIONTAG As Integer = 1
    Dim LnContador As Integer
    Dim rsLocal As ADODB.Recordset
    Dim i As Integer
    Dim J As Integer
    Dim CodFormaPago As Integer 'Var. Que contiene la forma de pago en uso, la toma del formulario de formas de pago
    Friend WithEvents Panel1 As Panel
    Dim Denominacion As Decimal 'Esta Variable contiene la denominacion a guardar
    'Dim mblnSALIR As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoFPAbcDenominaciones))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.FlexDenominaciones = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Frame1.SuspendLayout()
        CType(Me.FlexDenominaciones, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnAceptar
        '
        Me.btnAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.btnAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAceptar.Location = New System.Drawing.Point(130, 268)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAceptar.Size = New System.Drawing.Size(89, 31)
        Me.btnAceptar.TabIndex = 3
        Me.btnAceptar.Text = "&Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Silver
        Me.Frame1.Controls.Add(Me.txtFlex)
        Me.Frame1.Controls.Add(Me.FlexDenominaciones)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(18, 13)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(201, 249)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(32, 64)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(89, 20)
        Me.txtFlex.TabIndex = 2
        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtFlex.Visible = False
        '
        'FlexDenominaciones
        '
        Me.FlexDenominaciones.DataSource = Nothing
        Me.FlexDenominaciones.Location = New System.Drawing.Point(16, 24)
        Me.FlexDenominaciones.Name = "FlexDenominaciones"
        Me.FlexDenominaciones.OcxState = CType(resources.GetObject("FlexDenominaciones.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FlexDenominaciones.Size = New System.Drawing.Size(173, 207)
        Me.FlexDenominaciones.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.btnAceptar)
        Me.Panel1.Controls.Add(Me.Frame1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(234, 309)
        Me.Panel1.TabIndex = 4
        '
        'frmCorpoFPAbcDenominaciones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(257, 334)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCorpoFPAbcDenominaciones"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Denominaciones"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.FlexDenominaciones, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub



    Sub Encabezado()
        With FlexDenominaciones
            .set_ColWidth(C_ColDENOMINACION, 0, 2300)
            .set_ColWidth(C_ColDENOMINACIONTAG, 0, 0)
            .set_TextMatrix(0, C_ColDENOMINACION, "Denominación")
            .set_TextMatrix(0, C_ColDENOMINACIONTAG, "tag")
            'Poner el Apuntador en la Linea Cero, para posteriormente centrar el texto de las columnas
            .Row = 0
            For LnContador = 0 To (.get_Cols() - 1) Step 1
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
            Next LnContador
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub AgregarFilaFinal()
        'Agrea una Fila al Final del Grid
        With FlexDenominaciones
            If .Row = .Rows - 1 And CDbl(Numerico(.Text)) <> 0 Then
                ' Si se Presiono enter y estamos en la ultima fila, entonces se agregrara una nueva fila
                .AddItem("")
                ScrollGrid()
            End If
        End With
    End Sub

    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        Me.Visible = False
        gblnMostrarDatosGrid = False
    End Sub



    Private Sub FlexDenominaciones_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDenominaciones.DblClick
        FlexDenominaciones_KeyPressEvent(FlexDenominaciones, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(Keys.Return))
    End Sub

    Private Sub FlexDenominaciones_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDenominaciones.Enter
        FlexDenominaciones.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        If FlexDenominaciones.Row = 0 Then
            FlexDenominaciones.Row = 1
        End If
        Pon_Tool()
    End Sub


    Private Sub FlexDenominaciones_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles FlexDenominaciones.KeyDownEvent
        'Validar la únicamente la Tecla Supr, ya que las teclas (Flechas), no aplica, ya que sóo existe una columan en el grid
        With FlexDenominaciones
            Select Case eventArgs.keyCode
                Case Keys.Return
                    AgregarFilaFinal()
                Case Keys.Delete
                    'Si el cursor está en en renglón 0, que es el nombre de columna, entonces no se toma en cuenta la tecla Supr
                    If .Row = 0 Then Exit Sub
                    If .get_TextMatrix(.Row, C_ColDENOMINACION) <> "" Then
                        BorraGrid(.Row)
                    End If
            End Select
        End With
    End Sub

    Private Sub FlexDenominaciones_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles FlexDenominaciones.KeyPressEvent
        With FlexDenominaciones
            If eventArgs.keyAscii = System.Windows.Forms.Keys.Return And .Row = .Rows Then
                ' Si se Presiono enter y estamos en la ultima fila, entonces se agregrara una nueva fila
                .AddItem("")
            End If
            If eventArgs.keyAscii <> 0 Then
                If (.Col = C_ColDENOMINACION) Then
                    '''en esta parte se validará si es el rengón, columna que le
                    '''corresponde editarse
                    If (.Row > 1) Then
                        '''de tal modo que si el renglón es mayor que 1
                        '''y si un renglón antes del renglón actual está vacío,
                        '''el renglón actual no se editará
                        If Trim(.get_TextMatrix(.Row - 1, C_ColDENOMINACION)) = "" Then
                            .Focus()
                            Exit Sub
                        End If
                    End If
                    'Permite escribir en el TExtBox
                    'ModEstandar.MSHFlexGridEdit(FlexDenominaciones, txtFlex, eventArgs.keyAscii)
                    'Selecciona el contenido del TextBox
                    If Len(Trim(txtFlex.Text)) <> 1 Then
                        ModEstandar.SelTextoTxt(txtFlex)
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub frmCorpoFPAbcDenominaciones_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                           Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
    End Sub

    Private Sub frmCorpoFPAbcDenominaciones_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                           Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoFPAbcDenominaciones_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                         Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.Visible = False
        gblnMostrarDatosGrid = False
    End Sub

    Sub LlenaGrid()
        On Error GoTo MErr
        Dim i As Integer
        With FlexDenominaciones
            gStrSql = "Select * from CatDenominaciones Where CodFormaPago= " & Numerico((frmCorpoAbcFormasdePago.txtCodFormaPago).Text) & "  Order By Denominacion "
            .Clear()
            ModEstandar.BorraCmd()
            Cmd.CommandText = "Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            rsLocal = Cmd.Execute
            Encabezado()
            If rsLocal.RecordCount > 0 Then
                'Declarar un Arreglo con el numero de registros que hay el el Recorset
                If rsLocal.RecordCount < 9 Then
                    .Rows = 11
                Else
                    .Rows = rsLocal.RecordCount + 2
                End If
            Else
                'Se sale para que no muestre los valores en el grid, ya que no existe valor alguno
                Exit Sub
            End If
            'Poner los valores de los datos recopilados en las columnas TAG correspondientes
            For i = 1 To rsLocal.RecordCount
                .set_TextMatrix(i, C_ColDENOMINACION, Format((rsLocal.Fields("Denominacion").Value), "0.00"))
                .set_TextMatrix(i, C_ColDENOMINACIONTAG, Format((rsLocal.Fields("Denominacion").Value), "0.00"))
                rsLocal.MoveNext()
            Next i
            ScrollGrid()
        End With
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function CambiosGrid() As Object
        On Error GoTo MErr
        CambiosGrid = True
        With FlexDenominaciones
            For i = 1 To .Rows - 1
                If .get_TextMatrix(i, C_ColDENOMINACION) = "" And .get_TextMatrix(i, C_ColDENOMINACIONTAG) = "" Then
                    Exit For
                End If
                If .get_TextMatrix(i, C_ColDENOMINACION) <> .get_TextMatrix(i, C_ColDENOMINACIONTAG) Then Exit Function
            Next
        End With
        CambiosGrid = False
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub BorraGrid(ByRef Row As Integer)
        'Este Procediento borra un renglon del Grid
        'Si el Número de Filas que kedan en el grid, es menor de 8, se insertará una nueva fila al final del grid
        With FlexDenominaciones
            .RemoveItem(Row)
            'Si el número de filas es menor de 10 o esta posicionado en la utlima fila, entonces, agrega una fila
            If .Rows < 11 Or .Row = .Rows - 1 Then
                '            AgregarFilaFinal
                .AddItem("")
                .Row = .Row
            End If
            '        .Row = .Row - 1
        End With
    End Sub

    Public Sub ScrollGrid()
        'Procedimiento que pone el enfoque en el primer renglón vacío del Grid
        Dim i As Integer
        Dim nCont As Integer 'Cuenta los renglones que están ocupados (que no están vacíos)
        Dim nRen As Integer
        'Aparecen 7 renglones disponibles en el Grid
        'Si son menos de siete registros ocupados, no se utiliza el .TopRow
        'Pero, si son 7 ó más registros, el .TopRow manda el enfoque al primer renglón vacío
        'después de los renglones ocupados
        nRen = 9 'El máximo de renglones que aparece en el grid (Además del encabezado)
        nCont = 0
        With Me.FlexDenominaciones
            For i = 1 To .Rows
                If Trim(.get_TextMatrix(i, C_ColDENOMINACION)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next i
            If nCont < 9 Then
                'Hay menos de 9 registros
                '            .TopRow = 9
                .Row = nCont + 1
                .Col = C_ColDENOMINACION

            Else
                'Hay 9 ó más registros, hay que recorrer el grid
                .TopRow = (nCont - nRen) + 2
                .Row = nCont + 1
                .Col = C_ColDENOMINACION
            End If
        End With
    End Sub


    Sub Nuevo()
        'Este Procedmiento es diferente a el de los demas formularios.
        'Aqui, al elegir nuevo, lo que se hará es mostrar las denominaciones guardadar en la Bd
        LlenaGrid()
    End Sub

    Sub Guardar()
        '    'Este Procedimiento Guardará las denomnaciones escritas en el Grid, para lo cual.
        '    ' ELiminará todos los registros existentes en el BD correspondientes a la forma de pago en uso.
        '    ' y posteriormente guardará los nuevos datos que estén en el grid
        '    On Error GoTo MErr
        '    With frmCorpoAbcFormasdePago
        '        CodFormaPago = Numerico(.txtCodFormaPago)
        '    End With
        '    'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
        '    Screen.MousePointer = vbHourglass
        ''    Cnn.BeginTrans
        '    'En Primer lugar se hara la elimacion de los datos existentes en el BD
        '    ModStoredProcedures.PR_IECatDenominaciones CStr(CodFormaPago), CStr("10"), C_ELIMINACION, 0
        '    Cmd.Execute
        '
        '    'Ahora realizar la alta de las nuevas denominaciones ,de un apor una
        '    With FlexDenominaciones
        '        For I = 1 To .Rows
        '            Denominacion = Numerico(.TextMatrix(I, C_ColDENOMINACION))
        '            If Denominacion <> 0 Then ' Si es mayor de cetro, guardar la denominacion
        '                ModStoredProcedures.PR_IECatDenominaciones CStr(CodFormaPago), CStr(Denominacion), C_INSERCION, 0
        '                Cmd.Execute
        '            End If
        '        Next
        '    End With
        ''    Cnn.CommitTrans
        '    Screen.MousePointer = vbDefault
        '    'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.
        '
        ''    MsgBox "Las Denominaciones ha sido grabado correctamente ", vbInformation + vbOKOnly, "Mensaje"
        '    'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
        ''    Nuevo
        ''    Guardar = True
        ''    Limpiar
        '
        '    Exit Sub
        'MErr:
        ''    Cnn.RollbackTrans
        '    Screen.MousePointer = vbDefault
        ''    If Err.Number <> 0 Then ModEstandar.MostrarError
        '
    End Sub

    Sub LimpiarGrid()
        With FlexDenominaciones
            .Clear()
        End With
    End Sub

    Function ValorRepetido() As Boolean
        'Esta Fución verifica si un valor en el grid está repetido.
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Ant As String
        Dim Act As String
        Dim SalirCiclo As Boolean 'Deterina, que si al salir de un ciclo, se tendra que salir del primer ciclo tambien
        Dim FilaActual As Integer 'Guarda el Valor de la fila en que se estaba posicionado en el grid, antes de entrar al ciclo
        With FlexDenominaciones
            FilaActual = .Row
            .Col = 0
            For i = 1 To .Rows - 2
                .Row = i
                If .Text = "" Or SalirCiclo = True Then
                    .Row = FilaActual
                    Exit For
                End If
                Ant = .Text
                For J = i + 1 To .Rows - 1
                    .Row = J
                    If .Text = "" Then
                        .Row = FilaActual
                        Exit For
                    End If
                    Act = .get_TextMatrix(J, .Col)
                    If Ant = Act Then
                        ValorRepetido = True
                        SalirCiclo = True
                        Exit For
                    End If
                Next J
            Next i
            .Row = FilaActual
        End With
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Function

    Private Sub frmCorpoFPAbcDenominaciones_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If UnloadMode = 0 Then 'Cero Significa que el Usuraio seleccionó cerrar en el botón del formulario
        '    Cancel = 1 'Para que no se cierre el formulario
        '    btnAceptar_Click(btnAceptar, New System.EventArgs())
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        With FlexDenominaciones
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    txtFlex.Visible = False
                    txtFlex.Text = ""
                    If .Visible = True Then
                        FlexDenominaciones.Focus()
                    End If
                Case System.Windows.Forms.Keys.Return
                    If CDbl(Numerico(txtFlex.Text)) <> 0 Then
                        .set_TextMatrix(.Row, .Col, Format(Numerico(txtFlex.Text), "0.00"))
                        'Verificar si el Valor que se desea insertar está duplicado, para no permitir que se introduzca
                        If ValorRepetido() = True Then
                            MsgBox("Valor de Denominación Duplicado" & vbNewLine & "Verifique Porvafor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                            txtFlex.Visible = True
                            txtFlex.Focus()
                            FlexDenominaciones.set_TextMatrix(.Row, .Col, "")
                            Exit Sub
                        End If

                        'Primero se hace que el grid tome el valor del TextBox, para luego verificar si se debe agregar una fila o no..
                        AgregarFilaFinal()
                    End If
                    txtFlex.Text = ""
                    txtFlex.Visible = False
                    FlexDenominaciones.Col = .Col
                    '                FlexDenominaciones.Row = .Row + 1
            End Select
        End With
    End Sub


    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        With FlexDenominaciones
            If .Col = C_ColDENOMINACION Then
                KeyAscii = ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 5, 2, (txtFlex.SelectionStart))
            End If
        End With
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

End Class
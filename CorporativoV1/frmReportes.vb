Option Explicit On
Option Strict Off

Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Imports Microsoft.Reporting.WebForms

Public Class frmReportes

    Inherits System.Windows.Forms.Form

    Public WithEvents CrystalReportViewer1 As New CrystalDecisions.Windows.Forms.CrystalReportViewer
    Private WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Public Report As CRAXDRT.Report
    Public SubReport As CRAXDRT.Report
    Public reporteActual
    'As New Object
    Public rsReport As ADODB.Recordset
    Public rsSubReport1 As ADODB.Recordset
    Public rsSubReport2 As ADODB.Recordset
    Public rsSubReport3 As ADODB.Recordset
    Public rsSubReport4 As ADODB.Recordset
    Public rsSubReport5 As ADODB.Recordset

    Public aValues_ As Object
    Public aFormula_ As Object
    Public aParam_ As Object 'Variable usada por Paimi
    Public WithEvents btnSalir As Button
    Public I As Integer


    '    Private Sub ReportViewer1_Print(sender As Object, e As Microsoft.Reporting.WinForms.ReportPrintEventArgs) Handles ReportViewer1.Print
    '        On Error GoTo Errores

    '        Me.Report.PrinterSetup(Me.Handle.ToInt32)
    '        sender.UseDefault = False
    '        CrystalReportViewer1.PrintReport()

    'Errores:
    '        If Err.Number <> 0 Then ModErrores.Errores()
    '    End Sub


    Public Sub frmReportes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        CrystalReportViewer1.ReportSource = reporteActual

        'If (reporteActual = "rptVentasSalidaDeMercancia1") Then
        'CrystalReportViewer1.ReportSource = frmVtasRPTVentasSalidadeMercancia.rptVentasSalidaDeMercancia1
        'ElseIf (reporteActual = "rptBancosReporteMovimientosBancarios") Then
        'CrystalReportViewer1.ReportSource = frmBancosReportedeMovimientosBancarios.rptBancosReporteMovimientosBancarios
        'End If


        '        On Error GoTo Errores
        '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        '        Report.DiscardSavedData()

        '        If Not rsReport Is Nothing Then
        '            Report.Database.SetDataSource(rsReport, 3)
        '        End If
        '        If Not rsSubReport1 Is Nothing Then
        '            SubReport.Database.SetDataSource(rsSubReport1, 3)
        '        End If
        '        If Not rsSubReport2 Is Nothing Then
        '            SubReport.Database.SetDataSource(rsSubReport2, 3)
        '        End If
        '        If Not rsSubReport3 Is Nothing Then
        '            SubReport.Database.SetDataSource(rsSubReport3, 3)
        '        End If
        '        If Not rsSubReport4 Is Nothing Then
        '            SubReport.Database.SetDataSource(rsSubReport4, 3)
        '        End If
        '        If Not rsSubReport5 Is Nothing Then
        '            SubReport.Database.SetDataSource(rsSubReport5, 3)
        '        End If
        '        CrystalReportViewer1.ReportSource = Report

        '        If Not IsNothing(aFormula_) Then
        '            For I = LBound(aFormula_) To UBound(aValues_)
        '                SetFormula(aFormula_(I), aValues_(I))
        '            Next I
        '        End If
        '        If Not IsNothing(aParam_) Then
        '            For I = LBound(aParam_) To UBound(aValues_)
        '                SetParam(aParam_(I), aValues_(I))
        '            Next I
        '        End If

        '        ReportViewer1.Visible = true

        '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        '        CrystalReportViewer1.DisplayToolbar = True
        '        Icono(Me, MDIMenuPrincipalCorpo)

        'Errores:
        '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        '        If Err.Number Then
        '            ModEstandar.MostrarError()
        '        End If
    End Sub

    'Private Sub frmReportes_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
    '    If VB6.PixelsToTwipsX(Me.ClientRectangle.Width) > 0 And VB6.PixelsToTwipsY(Me.ClientRectangle.Height) > 0 Then
    '        With Me.CrystalReportViewer1
    '            .Top = VB6.TwipsToPixelsY(120)
    '            .Left = VB6.TwipsToPixelsX(120)
    '            .Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - VB6.PixelsToTwipsX(.Left) - 120)
    '            .Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(.Top) - 120)
    '        End With
    '    End If
    'End Sub

    'Private Sub frmReportes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    Report = Nothing
    '    SubReport = Nothing

    '    If Not rsReport Is Nothing Then If rsReport.State = ADODB.ObjectStateEnum.adStateOpen Then rsReport.Close()
    '    rsReport = Nothing

    '    If Not rsSubReport1 Is Nothing Then If rsSubReport1.State = ADODB.ObjectStateEnum.adStateOpen Then rsSubReport1.Close()
    '    rsSubReport1 = Nothing

    '    If Not rsSubReport2 Is Nothing Then If rsSubReport2.State = ADODB.ObjectStateEnum.adStateOpen Then rsSubReport2.Close()
    '    rsSubReport2 = Nothing

    '    If Not rsSubReport3 Is Nothing Then If rsSubReport3.State = ADODB.ObjectStateEnum.adStateOpen Then rsSubReport3.Close()
    '    rsSubReport3 = Nothing

    '    If Not rsSubReport4 Is Nothing Then If rsSubReport4.State = ADODB.ObjectStateEnum.adStateOpen Then rsSubReport4.Close()
    '    rsSubReport4 = Nothing

    '    If Not rsSubReport5 Is Nothing Then If rsSubReport5.State = ADODB.ObjectStateEnum.adStateOpen Then rsSubReport5.Close()
    '    rsSubReport5 = Nothing

    '    'Me = Nothing

    'End Sub

    '---------------------------------------------------------------------------------------------------------------------
    'Comienza Código usado por Paimí
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Imprime(ByRef cCaption As String, Optional ByRef aNomParam As Object = Nothing, Optional ByRef aValueParam As Object = Nothing)
        On Error GoTo ImprimeErr
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Not IsNothing(aNomParam) And Not IsNothing(aValueParam) Then
            aParam_ = aNomParam
            aValues_ = aValueParam
        Else
            aParam_ = Nothing
            aValues_ = Nothing
        End If
        Me.Text = cCaption
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Me.ShowDialog()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub

ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Function SetParam(ByVal ParamName As String, ByVal ParamValue As Object) As Boolean
        Dim intCounter As Integer
        For intCounter = 1 To Report.ParameterFields.Count
            If UCase(Report.ParameterFields(intCounter).ParameterFieldName) = UCase(ParamName) Then
                Report.ParameterFields(intCounter).AddCurrentValue(ParamValue)
                SetParam = True
            End If
        Next intCounter
    End Function
    '---------------------------------------------------------------------------------------------------------------------
    'Finaliza código de Paimí
    '---------------------------------------------------------------------------------------------------------------------

    Private Function SetFormula(ByVal FormulaName As String, ByVal FormulaText As String) As Boolean
        Dim intCounter As Integer
        For intCounter = 1 To Report.FormulaFields.Count
            If UCase(Report.FormulaFields(intCounter).FormulaFieldName) = UCase(FormulaName) Then
                Report.FormulaFields(intCounter).Text = "'" & FormulaText & "'"
                SetFormula = True
                Exit Function
            End If
        Next intCounter
        SetFormula = False
    End Function

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Public Sub InitializeComponent()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.ReportViewer1.Name = "ReportViewer"
        Me.ReportViewer1.Size = New System.Drawing.Size(396, 246)
        Me.ReportViewer1.TabIndex = 0
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.AutoSize = True
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalReportViewer1.ForeColor = System.Drawing.Color.Silver
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(12, 12)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ShowCloseButton = False
        Me.CrystalReportViewer1.ShowGroupTreeButton = False
        Me.CrystalReportViewer1.ShowLogo = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(965, 514)
        Me.CrystalReportViewer1.TabIndex = 0
        Me.CrystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(928, 532)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'frmReportes
        '
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(990, 584)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmReportes"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub CrystalReportViewer1_AutoSizeChanged(sender As Object, e As EventArgs) Handles CrystalReportViewer1.AutoSizeChanged
        'Dim frmReportes As frmReportes = New frmReportes()
        'CrystalReportViewer1.Size = frmReportes.Size
    End Sub
End Class
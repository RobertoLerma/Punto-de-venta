'**********************************************************************************************************************'
'*PROGRAMA: ACERCA DE GRUPO VITEK JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018    
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmAcercaDe

    Inherits System.Windows.Forms.Form

    Private Function FechaHoraUltCompilacion() As Date
        'On Error GoTo MErr
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.AppStarting
            FechaHoraUltCompilacion = FileDateTime(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe")
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Function
            'MErr:
        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MostrarError("Ocurrió un error al intentar recuperar la fecha de última compilación.", MsgBoxStyle.Exclamation)
        End Try
    End Function

    Private Sub frmAcercaDe_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then Me.Close()
    End Sub

    Private Sub frmAcercaDe_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        CentrarForma(Me)
        Text = "Acerca de " & My.Application.Info.Title
        lblVersion.Text = "Versión " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
        lblTitle.Text = UCase(My.Application.Info.Title)
        lblDescription.Text = "Sistema de control Joyería para:" & vbNewLine & gstrNombCortoEmpresa
        lblUltimaCompilacion.Text = "Última compilación: " & Format(FechaHoraUltCompilacion, "dd/MMMM/yyyy, hh:mm:ss am/pm")
        Dim FechaCompilacion As String = AgregarHoraAFecha(Today)
        lblfechaCompilacion.Text = FechaCompilacion
    End Sub

End Class
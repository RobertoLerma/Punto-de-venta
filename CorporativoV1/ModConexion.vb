'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE CONEXION JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018      
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports System.Configuration
Imports System.IO

Public Module ModConexion

    'Variables de conexion
    ''Public cnn As ADODB.Connection
    ''Public Cmd As Command
    ''Public gSql As String
    ''Public Rec As Recordset
    ''Public RecSet As Recordset
    ''Public RecTemporal As Recordset

    'Variables de Conexion para la BD de Access
    Public CnnAccess As ADODB.Connection
    Public CmdAccess As Command
    Public RsGralAccess As Recordset

    Public conexionLocalCliente As String
    Public BDUid As String = "sa"
    Public BDPwd As String = "r00tsqlsrv"
    Dim objReader As StreamReader

    'Funcion Abrir --> Abre la Conexion de Sql
    Public Function Abrir(Servidor As String, Bd As String) As Boolean
        'On Error GoTo Errores 
        Try
            Cnn = New Connection
            Cnn.ConnectionTimeout = 10
            If System.IO.File.Exists(rutaArchivoTxt) Then
                objReader = New StreamReader(rutaArchivoTxt.ToString)

                Dim sLine As String = ""
                Dim arrText As New ArrayList()
                Dim Linea As Integer = 0
                'Dim Servidor As String = ""
                'Dim BD As String = ""
                Do
                    sLine = objReader.ReadLine()
                    If Not sLine Is Nothing Then
                        arrText.Add(sLine)
                        Linea = Linea + 1
                        If Linea = 1 Then
                            Servidor = Trim(sLine)
                        End If

                        If Linea = 2 Then
                            Bd = Trim(sLine)
                        End If
                    End If
                Loop Until sLine Is Nothing
                objReader.Close()

            End If

            Cnn = New Connection
            Cnn.ConnectionTimeout = 30

            'CONEXION BD LOCAL
            conexionLocalCliente = "User ID=" & BDUid & ";Password=" & BDPwd & ";Initial Catalog=" & Bd & "; Data Source=" & Servidor & ",1433"


            'CONEXION BD SERVIDOR
            conexionServidor = "Provider=SQLOLEDB; server=" & Servidor & "; uid=" & BDUid & "; pwd=" & BDPwd & "; database=" & Bd & ";"
            Cnn.Open(conexionServidor)

            Abrir = True

            'Cnn.Open("Provider = SQLOLEDB.1;Integrated Security=SSPI;Persist Security = False ;Initial Catalog = " & Bd & "; Data Source = " & Servidor & "")
            Cnn.CursorLocation = CursorLocationEnum.adUseClient

            Cmd = New Command
            Cmd.ActiveConnection = Cnn
            Cmd.CommandText = gStrSql
            Cmd.CommandType = CommandTypeEnum.adCmdText

            RsGral = New Recordset
            With RsGral
                .ActiveConnection = Cnn
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                .LockType = LockTypeEnum.adLockReadOnly
            End With
            'Errores:  
            'Cnn.Close() 
        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            'If Err.Number <> 0 Then ModErrores.Errores()
            Abrir = False
        End Try
        Return Abrir

    End Function


    Public Function AbrirAccess() As Boolean
        'On Error GoTo Errores
        Try
            CnnAccess = New Connection
            CnnAccess.ConnectionTimeout = 30
            CnnAccess.Open("Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrCorpoDriveLocal & "\Sistema\InvElect\ImpEtiq.mdb; Persist Security Info=False")
            CnnAccess.CursorLocation = CursorLocationEnum.adUseClient
            CmdAccess = New Command
            CmdAccess.ActiveConnection = CnnAccess
            CmdAccess.CommandText = gStrSql
            CmdAccess.CommandType = CommandTypeEnum.adCmdText
            RsGralAccess = New Recordset
            With RsGralAccess
                .ActiveConnection = CnnAccess
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                .LockType = LockTypeEnum.adLockReadOnly
            End With
            AbrirAccess = True
            'Errores:
        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
        Return True
    End Function

    'Cerrar --> Destruye los Recordset y cierra la Conexion
    Public Sub Cerrar()
        'On Error GoTo Errores
        Try
            If RsGral.State = ObjectStateEnum.adStateOpen Then RsGral.Close()
            RsGral = Nothing


            If Cnn.State = ObjectStateEnum.adStateOpen Then Cnn.Close()
            Cnn = Nothing

            'Errores:
        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            If Err.Number <> 0 Then
                Err.Clear()
            End If
        End Try
    End Sub

    Public Function CerrarAccess() As Boolean
        ' On Error GoTo Errores
        Try
            If RsGralAccess.State = ObjectStateEnum.adStateOpen Then RsGralAccess.Close()
            RsGralAccess = Nothing
            If CnnAccess.State = ObjectStateEnum.adStateOpen Then CnnAccess.Close()
            CnnAccess = Nothing
            'Errores:
        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            If Err.Number <> 0 Then
                Err.Clear()
            End If
        End Try
        Return True
    End Function

    'Atributos  -->  Asigna atributos requeridos para que los Recordset funcionen correctamente
    Sub Atributos_RecordSet(Recor As Recordset)
        ' On Error GoTo Errores
        Try
            With Recor
                .ActiveConnection = Cnn
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                .LockType = LockTypeEnum.adLockReadOnly
            End With
            'Errores:
        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub

    Sub Atributos_Recordset_Access(Recor As Recordset)
        'On Error GoTo Errores
        Try
            With Recor
                .ActiveConnection = CnnAccess
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                .LockType = LockTypeEnum.adLockReadOnly
            End With
            'Errores:
        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub

    'Maquina  -->  Obtiene el nombre de la maquina por medio de windows
    Public Sub Maquina()
        'On Error GoTo Errores
        Try
            Dim strBuffer As String
            Dim lngBufSize As Long
            Dim lngStatus As Long

            lngBufSize = 255
            'strBuffer = String$(lngBufSize, " ")
            'lngStatus = getComputerName(strBuffer, lngBufSize)
            If lngStatus <> 0 Then
                'NombreMaquina = Left(strBuffer, lngBufSize)
            Else
                NombreMaquina = "X"
            End If
            'Errores:
        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub

    'Cerrar --> Destruye los Recordset y cierra la Conexion
    'Public Sub CerrarPuntoVenta()
    '    On Error GoTo Errores
    '    If RsgralPV.State = adStateOpen Then RsgralPV.Close
    '    Set RsgralPV = Nothing
    '    Set CmdPVenta = Nothing
    '    If CnnPVenta.State = adStateOpen Then CnnPVenta.Close
    '    Set CnnPVenta = Nothing
    'Errores:
    '    If Err.Number <> 0 Then
    '        Err.Clear
    '    End If
    'End Sub

    'Funcion Abrir --> Abre la Conexion de Sql
    'Public Function AbrirPuntoVenta(Servidor As String, Bd As String) As Boolean
    '    On Error GoTo Errores
    '    Set CnnPVenta = New Connection
    '    CnnPVenta.ConnectionTimeout = 30
    '    CnnPVenta.Open "Provider = SQLOLEDB.1;Integrated Security=SSPI;Persist Security = False ;Initial Catalog = " & Bd & "; Data Source = " & Servidor & ""
    ''    CnnCorpo.Close
    '    CnnPVenta.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '
    '    Set CmdPVenta = New Command
    '    Set CmdPVenta.ActiveConnection = CnnPVenta
    '    CmdPVenta.CommandText = gStrSql
    '    CmdPVenta.CommandType = adCmdText
    '
    '    Set RsgralPV = New Recordset
    '    With RsgralPV
    '        .ActiveConnection = CnnPVenta
    '        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '        .CursorType = adOpenForwardOnly
    '        .LockType = adLockReadOnly
    '    End With
    '    AbrirPuntoVenta = True
    'Errores:
    '    If Err.Number <> 0 Then ModErrores.Errores
    'End Function

End Module



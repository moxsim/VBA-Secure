Attribute VB_Name = "core_Config"
''
' VBA-Secure v1.0.0
' (c) Erukh Maksim - https://github.com/moxsim/VBA-Secure
'
' Secure VBA Config
'
' @class core_Config
' @author m0xsim@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Установите хэш-версии
Public Function VersionHash() As String
    VersionHash = "b746c46e8b93757df8b2376acd0b9703"
End Function

' Установите True если ходите отслеживать взаимодействие с API в Immediate Window
Public Function IsDebug() As Boolean
    IsDebug = True
End Function

' Установите True если ходите чтобы работала доменная аутентификация (SSPI)
' https://ru.wikipedia.org/wiki/Security_Support_Provider_Interface
Public Function IsSecure() As Boolean
    IsSecure = True
End Function

' Установите таймаут по умолчанию для выполнения запросов (в миллисекундах)
Public Function Timeout(Optional Value As Long = 0) As Long
    If Value <= 0 Then
        Timeout = 120000 ' таймаут в миллисекундах по умолчанию
    Else
        Timeout = Value
    End If
End Function

' Установите протокол. У вас есть два варианта https или http
Public Function Protocol() As String
    Protocol = "https"
    'Protocol = "http"
End Function

' Установите хост
Public Function Host() As String
    Host = "localhost:443"
End Function

' Установите API
Public Function API() As String
    API = "api"
End Function

' ------------------------------------------------------------------------------

Public Function Url() As String
    Url = Protocol() & "://" & Host() & "/" & API()
End Function

Public Sub Log(Text As String)
    If IsDebug() Then Debug.Print Text
End Sub



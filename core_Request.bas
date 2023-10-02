Attribute VB_Name = "core_Request"
''
' VBA-Secure v1.0.0
' (c) Erukh Maksim - https://github.com/moxsim/VBA-Secure
'
' Secure VBA Request
'
' @class core_Request
' @author m0xsim@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Отправка JSON-запроса
Public Function SendJson( _
    Route As String, _
    Payload As Object, _
    Optional Timeout As Long = 0, _
    Optional isAsync As Boolean = False _
) As Object
    
    Dim Url As String
    Url = core_Config.Url()
    
    ' Добавляем в запрос hash-версии
    Payload("_version") = core_Config.VersionHash()

    Dim PayloadJson As String
    PayloadJson = core_Json.ConvertToJson(Payload)
    
    Timeout = core_Config.Timeout(Timeout)
    
    Dim ResponseJson As String
    ResponseJson = private_Send("POST", Url, Route, PayloadJson, Timeout, isAsync)
    
    core_Config.Log (ResponseJson)
    
    Set SendJson = core_Json.ParseJson(ResponseJson)

End Function

' Отправка HTTP-запроса
Private Function private_Send( _
    Method As String, _
    Url As String, _
    Optional Route As String, _
    Optional Payload As String = "", _
    Optional Timeout As Long = 0, _
    Optional isAsync As Boolean = False _
) As String
    
    Dim Http
    
    Timeout = core_Config.Timeout(Timeout)
    
    ' Сообщение для лога/ошибки
    Dim msg As String
    
    If core_Config.IsSecure() Then
        ' Доменная аутентификация SSPI
        Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        Const AutoLogonPolicy_Always = 0
        Http.SetAutoLogonPolicy AutoLogonPolicy_Always ' Тут передаём Credentials
    Else
        ' Без аутентификации
        Set Http = New MSXML2.ServerXMLHTTP
        
        ' Подавить ошибки сертификата
        Http.setOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    End If
    
    ' таймауты в миллисекундах https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms760403(v=vs.85)
    '   resolveTimeout  - value is applied to mapping host names (such as "www.microsoft.com") to IP addresses
    '   connectTimeout  - value is applied to establishing a communication socket with the target server
    '   sendTimeout     - value applies to sending an individual packet of request data (if any) on the communication socket to the target server
    '   receiveTimeout  - value applies to receiving a packet of response data from the target server
    Http.setTimeouts 10000, 10000, 10000, Timeout
    
    ' Открытие соединения
    core_Config.Log ("Connection opening: " & Url)
    
    Http.Open Method, Url, isAsync
    Http.setRequestHeader "ROUTE", Route
    
    ' Нужно ли отправлять данные
    If Payload = "" Then
        core_Config.Log ("Sending...")
        Http.Send
    Else
        core_Config.Log ("Sending payload")
        core_Config.Log (Payload)
        Http.Send Payload
    End If
    
    ' Передача управления пока ожидается ответ 
    If isAsync Then
        Do While Http.readyState <> 4
            DoEvents
        Loop
    End If
    
    Dim Result As Object
    
    ' Обрабатываем статусы выполнения
    core_Config.Log (Http.Status & " " & Http.StatusText)
    If Http.Status = 200 Then
        private_Send = Http.responseText
    ElseIf Http.Status = 400 Then
        private_Send = Http.responseText
        
        Set Result = core_Json.ParseJson(private_Send)
        MsgBox Result("_error"), vbOKOnly, "Ошибка запроса"
        core_Config.Log (Result("_error"))
    Else
        Set Result = New Dictionary
        Result("_error") = Http.Status & " " & Http.StatusText
        private_Send = core_Json.ConvertToJson(Result)
        
        MsgBox Http.Status & " " & Http.StatusText, vbOKOnly, "Ошибка запроса"
        core_Config.Log (Result("_error"))
    End If
    
    ' Закрываем соединение
    Set Http = Nothing
End Function

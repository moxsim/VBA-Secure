Attribute VB_Name = "core_Example"
''
' VBA-Secure v1.0.0
' (c) Erukh Maksim - https://github.com/moxsim/VBA-Secure
'
' Secure VBA Example
'
' @class core_Example
' @author m0xsim@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Пример отправки Json в API
Sub JsonSend()
Attribute JsonSend.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Request As Object
    Dim Response As Object
    
    Set Request = New Dictionary
    Request("test1") = 1
    Set Request("data") = New Dictionary ' Или тут можно так написать: Request.Add "data", New Dictionary
        Request("data")("test2") = "Test 2"
    
    Set Response = core_Request.SendJson("demo/user", Request)
    
    'Debug.Print core_Json.ConvertToJson(Response)
    
End Sub

' Примеры конвертации JSON в объект и обратно
Sub JsonConvert()
    
    Dim Json As Object
    Set Json = core_Json.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")
    
    ' Json("a") -> 123
    ' Json("b")(2) -> 2
    ' Json("c")("d") -> 456
    Json("c")("e") = 789
    
    Debug.Print core_Json.ConvertToJson(Json)
    ' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"
    
    Debug.Print core_Json.ConvertToJson(Json, Whitespace:=2)
    ' -> "{
    '       "a": 123,
    '       "b": [
    '         1,
    '         2,
    '         3,
    '         4
    '       ],
    '       "c": {
    '         "d": 456,
    '         "e": 789
    '       }
    '     }"

End Sub

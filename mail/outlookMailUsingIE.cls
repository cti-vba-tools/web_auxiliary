Private IE As Object
Private bodyDic As Object
Private clsTo As String
Private clsCc As String
Private clsBcc As String
Private clsSubject As String
Private clsBody As String
Private Const baseUrl As String = "https://outlook.office.com/mail/deeplink/compose?"

Enum send_mode
    Default = 0
    Outlook = 1
    Gmail = 2
End Enum

Public Function Send(Optional Subject As String = "", Optional Body As Variant = "", Optional Address As String = "", Optional mode As send_mode = 0) As Long
    Dim url As String
    Select Case TypeName(Body)
    Case "String"
        Set bodyDic = CreateObject("Scripting.Dictionary")
        For Each Row In Split(Body, vbclrf)
            bodyDic.Add bodyDic.Count, Row
        Next
    Case Else
        If IsArray(Body) Then
            For Each Item In Body
                bodyDic.Add bodyDic.Count, Item
            Next
        Else
            Exit Function
        End If
    End Select
    If Subject <> "" Then
        clsSubject = Subject
    End If
    If Address <> "" Then
        clsTo = Address
    End If
    url = baseUrl & "subject=" & EncodeUriComponent(clsSubject) & "&body=" & EncodeUriComponent(Join(bodyDic.Items, vbCrLf)) & "&to=" & EncodeUriComponent(clsTo)
    SendDefBrowser url
End Function

Private Sub Class_Initialize()
    Set IE = Nothing
    Set bodyDic = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set IE = Nothing
    Set bodyDic = Nothing
End Sub

Private Function EncodeUriComponent(strText)
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        objHtmlfile.parentWindow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    End If
    EncodeUriComponent = objHtmlfile.parentWindow.encode(strText)
End Function

Private Function SendIE(url As String) As Long
    On Error Resume Next
    If url = "" Then Exit Function
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.Navigate url
    Set IE = Nothing
    SendIE = Err.Number
    On Error GoTo 0
End Function

Private Function SendDefBrowser(url As String) As Long
    On Error Resume Next
    If url = "" Then Exit Function
    Dim WSH As Object
    Set WSH = CreateObject("Wscript.Shell")
    WSH.Run url, 3
    Set WSH = Nothing
    SendDefBrowser = Err.Number
    On Error GoTo 0
End Function

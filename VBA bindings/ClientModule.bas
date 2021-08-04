Attribute VB_Name = "ClientModule"
Option Explicit

'----------------------------------------------------------------------------
'                        Messages
'----------------------------------------------------------------------------
Public Const LocalNewVersion = "New version is already downloaded, do you want to open that?"
Public Const ServerNewVersion = "A new version is available, do you want to update your current version?"
Public Const EOLLocalNewVersion = "Current version no longer supported. Do you want to open a the new version?"
Public Const EOLServerNewVersion = "Current version no longer supported. Do you want update to the new version?"
Public Const HttpServerError = "Error getting data from cloud server. Check the log file"
Public Const InternetError = "Unable to connect. Check your internet connection"
Public Const LoginError = "Unable to login. Recheck your email address and password"
Public Const BlankOtp = "Enter One Time Password"
Public Const InvalidOtp = "Enter valid One Time Password"
Public Const BlankEmailPassword = "Email Address and Password is required"

Public SessionID As String
Public AuthHash As String
Public Authenticated As Boolean
Public Const TimerIncrement As Integer = 10 ' minutes to check for timeout
Public LastRequestTime As Date

'-------------------------------------------------------------------
'               NTVX Server Request
'-------------------------------------------------------------------

' This sub is called on the client side to send the HttpRequest to server and also parse the data returned from server
' Parameters:
' clientButton: Button which the user clicked. This is used for telemetry
' templateFile: file name of the excel template on the web server that needs to opened
' copySheetsAndRanges:  how to copy client excel data on to server excel.
'                       format "Client Sheet Name~Client Excel Range~Server Sheet Name~Server Start Cell;.....
'                       example format: "Input1~A1:C7~LinkB~B1;Input2A~C5:T8~LinkB~P1"
'                       the above format means the following:
'                       Input1~A1:C7~LinkB~B1 :- From Client Excel, copy Input1 sheet, range A1 to C7 to Server LinkB sheet at B1
'                       Input2A~C5:T8~LinkB~P1 :- From Client Excel, copy Input2A sheet range C5 to T8 to server LinkB sheet at P1
' macroToExecute: name of the macro that needs to be executed on the server excel
' returnSheetsAndRanges: this is semicolon delimited excel sheets and ranges that need to be brought from server on to client
'                        in xml
'                        format: SheetName~ExcelRange (optional);SheetName~ExcelRange ....
'                        example format: "FinModel~A1:C7,E10:J18;BuyerROI~A1:H20;ROI"
'                        for the above format the following will be included in the XML
'                        FinModel~A1:C7,E10:J18 :- In FinModel sheet Ranges A1:C7 & E10:J18
'                        BuyerROI~A1:H20 :- In BuyerROI sheet A1:H20
'                        ROI :- Entire ROI sheet as no range is specified
' Returns: Parses the xml string returned by the server and populates the global xml dictionary. After this use PasteSheetRange
'          to copy parts of the server excel sheet range on to client excel sheet range
'
                         
Public Sub NTVXRequest(clientButton As String, templateFile As String, copySheetsAndRanges As String, _
            macroToExecute As String, returnSheetsAndRanges As String, Optional deleteFile As Boolean = True)
    Dim inputXmlString As String
    Dim serverResponse As String
    
    inputXmlString = GetInputXml(copySheetsAndRanges)
    
    ' sending http request to the server
     serverResponse = NTVXInternalRequest(clientButton, templateFile, inputXmlString, copySheetsAndRanges, macroToExecute, returnSheetsAndRanges, deleteFile)
            
    'parsing the xml received from server
    ParseXml (serverResponse)
    
End Sub

' This function is used to call the web server with xml data. The webserver runs the server macro and returns the xml string back
Private Function NTVXInternalRequest(clientButton As String, templateFile As String, inputXmlString As String, copySheetsAndRanges As String, _
        macroToExecute As String, returnSheetsAndRanges As String, deleteFile As Boolean) As String

    Dim formData As String
    formData = "ClientButton=" & ClientModule.URLEncode(clientButton) _
            & "&ClientFile=" & ClientModule.URLEncode(ClientConfiguration.ClientFile) _
            & "&TemplateFile=" & ClientModule.URLEncode(templateFile) _
            & "&InputXmlString=" & ClientModule.URLEncode(inputXmlString) _
            & "&CopySheetsAndRanges=" & ClientModule.URLEncode(copySheetsAndRanges) _
            & "&MacroToExecute=" & ClientModule.URLEncode(macroToExecute) _
            & "&ReturnSheetsAndRanges=" & ClientModule.URLEncode(returnSheetsAndRanges) _
            & "&DeleteFile=" & ClientModule.URLEncode(CStr(deleteFile))
            
          
    
    NTVXInternalRequest = ClientModule.HttpRequest("/excel/process", formData)
End Function

Public Function GetInputXml(copySheetsAndRanges As String) As String
    Dim sheet As Variant
    Dim sourceDest() As String
    
    Dim dicSheet As New Dictionary
    
    For Each sheet In Split(copySheetsAndRanges, ";")
        If Trim(sheet) <> "" Then
            sourceDest = Split(sheet, "~")
            If dicSheet.Exists(sourceDest(0)) Then
                dicSheet(sourceDest(0)) = dicSheet(sourceDest(0)) & "," & sourceDest(1)
            Else
                dicSheet(sourceDest(0)) = sourceDest(1)
            End If
        End If
    Next

    Dim key
    Dim inputSheetsAndRanges As String
    inputSheetsAndRanges = ""
    For Each key In dicSheet.Keys
        inputSheetsAndRanges = inputSheetsAndRanges & key & "~" & dicSheet(key) & ";"
    Next
    
    GetInputXml = SheetsAndRangesToXml(inputSheetsAndRanges)
End Function

'--------------------------------------------------------------------
'                         Authentication
'--------------------------------------------------------------------

Sub Auto_Open()
    Login
End Sub
Sub Auto_Close()
    Logout
    ThisWorkbook.Close False
End Sub
Public Sub Login()
    ClientModule.AuthHash = ""
    ClientModule.Authenticated = False
        
    LoginForm.Show
    
    If Not ClientModule.Authenticated Then
        End
    End If
    TimerAndTimeOut
End Sub

Public Sub CheckAuthentication()
    If ClientModule.AuthHash = "" Or Not ClientModule.Authenticated Then
        Login
    End If
End Sub

Public Function Authenticate(Email As String, Password As String) As String
    Dim authStr As String
    authStr = Email & ":" & Password
    ClientModule.AuthHash = "Bearer " & ClientModule.EncodeBase64(authStr)
    
    Dim responseText As String
    Dim formData As String
    formData = "clientFile=" & ClientModule.URLEncode(ClientConfiguration.ClientFile)
    
    responseText = ClientModule.HttpRequest("/excel/login", formData, True)

    If responseText = "SUCCESS" Then
        Authenticate = responseText
    ElseIf responseText <> "" Then
        ClientModule.ParseSessionStartResponse responseText
        ClientModule.Authenticated = True
        Authenticate = "BYPASS"
    End If
End Function

Public Sub Logout()
    ClientModule.SessionEnd
    ClientModule.Authenticated = False
    ClientModule.AuthHash = ""
    ClientModule.SessionID = ""
End Sub

Private Sub TimerAndTimeOut()
    Dim defaultDate As Date
    Dim nextRunDate As Date
    On Error Resume Next
    
    If defaultDate = ClientModule.LastRequestTime Then
        Exit Sub
    End If
    
    If DateTime.DateDiff("n", ClientModule.LastRequestTime, Now) > ClientConfiguration.IdleTimeout Then
        Logout
    Else
        nextRunDate = DateTime.DateAdd("n", ClientModule.TimerIncrement, Now)
        Application.OnTime nextRunDate, "TimerAndTimeOut"
    End If
End Sub

Public Function ValidateOtp(Otp As String)
    Dim responseText As String
    Dim formData As String
    formData = "otp=" & ClientModule.URLEncode(Otp) _
            & "&clientFile=" & ClientModule.URLEncode(ClientConfiguration.ClientFile)

    responseText = ClientModule.HttpRequest("/excel/validateotp", formData, True)
    If responseText <> "" Then
        ClientModule.ParseSessionStartResponse responseText
        ClientModule.Authenticated = True
    End If
End Function

Public Sub SessionStart()
    Dim responseText As String
    Dim formData As String
    formData = "clientFile=" & ClientModule.URLEncode(ClientConfiguration.ClientFile)
    
    responseText = ClientModule.HttpRequest("/excel/sessionstart", formData)
    ParseSessionStartResponse responseText
End Sub

Public Sub ParseSessionStartResponse(responseText As String)
    Dim arr() As String
    arr = Split(responseText, "~")
    
    If arr(2) = "4" And arr(1) = "" Then
        ClientModule.CriticalError (arr(0))
        End
    End If
    
    If Not arr(1) = "" Then
        NewVersion arr(1), arr(2)
    End If
End Sub

Private Sub NewVersion(newFile As String, messageType As String)
    'checking the newFile existing the current folder
    Dim newFilePath As String
    Dim message As String
    Dim title As String
    Dim button As VbMsgBoxStyle
    
    newFilePath = ThisWorkbook.Path & "\" + newFile
    title = IIf(messageType = "4", "Not Supported", "New Version")
    
    If Dir(newFilePath) <> "" Then
        message = IIf(messageType = "4", ClientModule.EOLLocalNewVersion, ClientModule.LocalNewVersion)
        button = IIf(messageType = "4", vbCritical + vbYesNo, vbExclamation + vbYesNo)
        If MsgBox(message, button, title) = vbYes Then
            OpenLocal newFile
        End If
    Else
        message = IIf(messageType = "4", ClientModule.EOLServerNewVersion, ClientModule.ServerNewVersion)
        button = IIf(messageType = "4", vbCritical + vbYesNo, vbInformation + vbYesNo)
        If MsgBox(message, button, title) = vbYes Then
            Download newFile
            OpenLocal newFile
        End If
    End If
End Sub
Private Sub OpenLocal(fileName As String)
    Dim localFile As String
    localFile = ThisWorkbook.Path & "\" & fileName
    Workbooks.Open localFile
    ThisWorkbook.Close (True)
End Sub

Private Sub Download(fileName As String)
    Dim http As New XMLHTTP60
    Dim url As String
    Dim formData As String
    
    url = ClientConfiguration.ServerUrl + "/excel/getclientfile"
    formData = "clientFile=" & ClientModule.URLEncode(fileName)
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send formData
     
    If http.Status = 200 Then
        'WriteResponseLog http.responseText, templateFile, inputXmlString, copySheetsAndRanges, macroToExecute, returnSheetsAndRanges
        Dim bytes() As Byte
        Dim ClientFile As String
        
        
        ClientFile = ThisWorkbook.Path & "\" & fileName
        Open ClientFile For Binary Access Write As #1
        bytes = http.responseBody
        Put #1, , bytes
        Close #1
    Else
        'WriteErrorLog http.responseText, templateFile, inputXmlString, copySheetsAndRanges, macroToExecute, returnSheetsAndRanges
        MsgBox "Error getting file from cloud server. Check the log file", vbOKOnly
        End ' ending the client macro execution if there is a server error
    End If
End Sub

Public Sub SessionEnd()
    Dim responseText As String
    Dim formData As String
    formData = ""
    
    responseText = ClientModule.HttpRequest("/excel/sessionend", formData)
End Sub


'-------------------------------------------------------------
'                         HTTP Methods
'-------------------------------------------------------------
Public Function HttpRequest(url As String, formData As String, Optional fromLoginPage As Boolean = False) As String

    If Not fromLoginPage Then
        ClientModule.CheckAuthentication
    End If

    ClientModule.LastRequestTime = Now
    url = ClientConfiguration.ServerUrl & url
    ClientModule.WriteRequestLog url, formData
    
    On Error GoTo InternetErr
    
    Dim http As New XMLHTTP60
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "SessionID", ClientModule.SessionID
    http.setRequestHeader "Authorization", ClientModule.AuthHash
    http.send formData

     If http.Status = 200 Then
        ClientModule.WriteResponseLog url, http.responseText, formData
        ClientModule.SessionID = http.getResponseHeader("SessionID")
        HttpRequest = http.responseText
        Exit Function
    ElseIf http.Status = 401 Then
        ClientModule.AuthHash = ""
        ClientModule.Authenticated = False
        MsgBox ClientModule.LoginError, vbCritical
    Else
        Dim errMessage As String
        errMessage = http.responseText
        
        ClientModule.WriteErrorLog url, http.responseText, formData
        If InStr(errMessage, "Error: ") <> 0 Then
            errMessage = Trim(Replace(errMessage, vbCrLf, ""))
            errMessage = Right(errMessage, Len(errMessage) - 7)
            MsgBox errMessage, vbCritical
        Else
            MsgBox ClientModule.HttpServerError, vbCritical
        End If
    End If
    If fromLoginPage Then
        Exit Function
    Else
        End ' ending the client macro execution if there is a server error
    End If
    
InternetErr:
    MsgBox ClientModule.InternetError, vbCritical
    If fromLoginPage Then
        Exit Function
    Else
        End
    End If
End Function
Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim Result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          Result(i) = Char
        Case 32
          Result(i) = Space
        Case 0 To 15
          Result(i) = "%0" & Hex(CharCode)
        Case Else
          Result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(Result, "")
  End If
End Function

Public Function URLDecode(StringToDecode As String) As String

Dim TempAns As String
Dim CurChr As Integer

CurChr = 1

Do Until CurChr - 1 = Len(StringToDecode)
  Select Case Mid(StringToDecode, CurChr, 1)
    Case "+"
      TempAns = TempAns & " "
    Case "%"
      TempAns = TempAns & Chr(Val("&h" & _
         Mid(StringToDecode, CurChr + 1, 2)))
       CurChr = CurChr + 2
    Case Else
      TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
  End Select

CurChr = CurChr + 1
Loop

URLDecode = TempAns
End Function

'----------------------------------------------------------------------------
'                      Logging & Messages
'----------------------------------------------------------------------------

Public Sub CriticalError(message As String)
    MsgBox message, vbCritical, "Error"
    End
End Sub

' This is used to log the errors returned by the server. Used to debug the issue
Public Sub WriteErrorLog(url As String, errorLog As String, formData As String)
    WriteLog "Error.log", url, errorLog, formData
End Sub

Public Sub WriteRequestLog(url As String, formData As String)
    WriteLog "Client.log", url, "", formData
End Sub

Public Sub WriteResponseLog(url As String, serverResponse As String, formData As String)
    WriteLog "Server.log", url, serverResponse, formData
End Sub

' This is generic log funciton which is used to log errors and also messages sent and received from server.
Private Sub WriteLog(fileName As String, url As String, log As String, formData As String)
    Dim i As Integer
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\log"
    If Len(Dir(filePath, vbDirectory)) = 0 Then
        MkDir filePath
    End If
    fileName = filePath & "\" + fileName
        
    Open fileName For Output As #1
    
    If log <> "" Then
        Print #1, log
        Print #1, ""
    End If
    
    Print #1, "URL:" & url
    
    Dim arr() As String
    arr = Split(formData, "&")
    
    For i = 0 To UBound(arr)
        Print #1, URLDecode(arr(i))
    Next
    
    Print #1, "DateTime: " & DateTime.Now
    Close #1
End Sub

Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument60
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function





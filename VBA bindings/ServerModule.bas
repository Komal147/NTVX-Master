Attribute VB_Name = "ServerModule"
Option Explicit

' This Function is used to process the data and instructions sent from the client. This is invoked from MVC application
' For detailed description of parameters check the HttpRequest sub above
' Parameters:
' inputXmlString - Client excel values converted to xml and sent to server
' copySheetsAndRanges - copy the data from xml (client excel data) on to the server excel sheet/range
'                          format "clientSheet-clientRange-serverSheet-serverStartCell;.....
' macroToExecute: name of the macro that needs to be executed on the server excel
' returnSheetsAndRanges - list of sheets and ranges to be converted to xml and send back to client
'                          format - "FinModel-A1:C7,E10:J18;BuyerROI-A1:H20;ROI"
Public Function ServerProcess(inputXmlString As String, copySheetsAndRanges As String, macroToExecute As String, returnSheetsAndRanges As String, _
                                downloadMacro As String, downloadFileName As String) As String
    On Error GoTo ServerError
    ' setting the input xml on to the sheets
    ParseXml (inputXmlString)
    SharedModule.PasteSheetsAndRanges (copySheetsAndRanges)
    If macroToExecute <> "" Then
        Application.Run (macroToExecute)
    End If
    
    If Not downloadMacro = vbNullString Then
        Application.Run downloadMacro, downloadFileName
    End If
    
   
    ServerProcess = SharedModule.SheetsAndRangesToXml(returnSheetsAndRanges)
    Exit Function
    
ServerError:
    ServerProcess = "Error:" & vbCrLf & "Number:" & Err.Number & vbCrLf & "Source:" & Err.Source & vbCrLf & "Description:" & Err.Description
    Err.Clear
End Function



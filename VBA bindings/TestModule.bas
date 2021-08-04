Attribute VB_Name = "TestModule"
' Testing the call of server excel over http. Used during development phase
' Not needed for production

Private Sub TestA2BOverHttp()
    Dim copySheetsAndRanges As String, returnSheetsAndRanges As String
    
    copySheetsAndRanges = "Input1~A1:C7~PPK~B1;Input2A~C5:T8~PPK~P1"
    returnSheetsAndRanges = "FinModel~E5:H8;ROI~A168:J188;Summary~A1:E5"
    
    ' sending and receiving data from server
    ClientModule.NTVXRequest "TestA2BOverHttp", "B_106.xls", copySheetsAndRanges, "", returnSheetsAndRanges
    
    ' Pasting the data from server as needed
    PasteSheetRange "FinModel", "E5:H8", "PPK", "A1"
    PasteSheetRange "ROI", "A168:J188", "PPK", "A6"
    PasteSheetRange "Summary", "A1:E5", "PPK", "E1"
        
    MsgBox "Done"
End Sub

' This is to test calling B (server excel) directly from A (client excel). Both should in the same physical folder
' Not needed for production
Private Sub TestA2B()
    Dim copySheetsAndRanges As String, returnSheetsAndRanges As String
    Dim inputXmlString As String
    
    copySheetsAndRanges = "Input1~A1:C7~PPK~M1;Input2A~C5:T8~PPK~P1;"
    returnSheetsAndRanges = "FinModel;BuyerROI;LinkB;Summary;IncStatement;ROI"
       
    inputXmlString = ClientModule.GetInputXml(copySheetsAndRanges)
       
    'Opening B excel locally
    Dim wbServer As Workbook
    Set wbServer = Workbooks.Open(ThisWorkbook.Path & "\" & "B_106.xls")
    
    
    ' calling the server macro locally
    Dim serverResponse As String
    serverResponse = Application.Run("'" & wbServer.Name & "'!NTVXModule.ServerProcess", inputXmlString, copySheetsAndRanges, "Module1.Run_DO", returnSheetsAndRanges)

    'closing the B excel
    wbServer.Close savechanges:=True
    Set wbServer = Nothing
    
    'parsing the xml received from server
    SharedModule.ParseXml (serverResponse)
    
    ' Pasting the data from server as needed
    PasteSheetRange "FinModel", "E5:H8", "PPK", "A1"
    PasteSheetRange "ROI", "A168:J188", "PPK", "A7"
    PasteSheetRange "Summary", "A1:E5", "PPK", "F1"
    
    MsgBox "Done"
    
    
End Sub



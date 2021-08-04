Attribute VB_Name = "SharedModule"
Option Explicit
Dim xmlDict As Dictionary

' This function is used to create xml string from the excel sheets and ranges
' format for input - "FinModel-A1:C7,E10:J18;BuyerROI-A1:H20;ROI"
' The following will be included in the xml
'              - In FinModel sheet Ranges A1:C7 & E10:J18
'              - In BuyerROI sheet A1:H20
'              - Entire ROI sheet as no range is specified
Public Function SheetsAndRangesToXml(sheetsAndRanges As String) As String
    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim sheetName As Variant
    Dim sheetAndRange As Variant
    
    Dim maxColumns As Long
    Dim maxRows As Long
    
    Dim sheetRange As String
    Dim cellRange As String
    Dim cell As range
    
    Dim xmlDoc As New DOMDocument60
    Dim xmlRoot As IXMLDOMElement
    Dim xmlSheet As IXMLDOMElement
    
    'creating root element
    Set xmlRoot = xmlDoc.createElement("ntvx")
    xmlDoc.appendChild xmlRoot
    
    Set wkb = ThisWorkbook
    
    For Each sheetName In Split(sheetsAndRanges, ";")
        If sheetName = "" Then
            Exit For
        End If
        If InStr(sheetName, "~") > 0 Then
            sheetAndRange = Split(sheetName, "~")
            Set wks = wkb.Sheets(sheetAndRange(0))
            cellRange = sheetAndRange(1)
        Else
            Set wks = wkb.Sheets(sheetName)
            maxRows = wks.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            maxColumns = wks.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            Set cell = wks.Cells(maxRows, maxColumns)
            cellRange = "$A$1:" & cell.address
        End If
        
        Set xmlSheet = SheetRangesToXml(xmlDoc, wks.Name, cellRange)
        xmlRoot.appendChild xmlSheet
    Next
    
    SheetsAndRangesToXml = xmlDoc.XML
End Function

' This function is an internal function which converts a sheet and range to xml string
Private Function SheetRangesToXml(xmlDoc As DOMDocument60, sheetName As String, range As String) As IXMLDOMElement
    Dim cellRange As range, cell As range
    
    Dim wkb As Workbook
    Dim wks As Worksheet
        
    Set wkb = ThisWorkbook
    Set wks = wkb.Sheets(sheetName)
    Set cellRange = wks.range(range)
    
    Dim xmlSheet As IXMLDOMElement
    Dim xmlCell As IXMLDOMElement
    
    Set xmlSheet = xmlDoc.createElement("sheet")
    xmlSheet.setAttribute "name", sheetName
    
    For Each cell In cellRange
        If Application.IsError(cell) Then
            Set xmlCell = xmlDoc.createElement("cell")
            xmlCell.setAttribute "address", Replace(cell.address, "$", "")
            xmlCell.text = CStr(cell.Value)
            xmlSheet.appendChild xmlCell
        End If
        If Not Application.IsError(cell) And Not VBA.IsEmpty(cell) Then
            Set xmlCell = xmlDoc.createElement("cell")
            xmlCell.setAttribute "address", Replace(cell.address, "$", "")
            xmlCell.text = cell.Value
            xmlSheet.appendChild xmlCell
        End If
    Next
    Set SheetRangesToXml = xmlSheet
End Function

' This is called after receiving the response from the web
' This parses the xml string and creates xml object
Public Sub ParseXml(xmlString As String)
    Dim xmlDoc As New DOMDocument60
    Dim xmlRoot As IXMLDOMElement
    Dim xmlSheet As IXMLDOMElement
    Dim xmlCell As IXMLDOMElement
    Dim dicSheet As Dictionary
    
    xmlDoc.LoadXML xmlString
    Set xmlRoot = xmlDoc.ChildNodes(0)

    Set xmlDict = New Dictionary
    
    For Each xmlSheet In xmlRoot.ChildNodes
        Set dicSheet = New Dictionary
        xmlDict.Add xmlSheet.getAttribute("name"), dicSheet
        For Each xmlCell In xmlSheet.ChildNodes
            dicSheet.Add xmlCell.getAttribute("address"), xmlCell.text
        Next
    Next
End Sub

' This method is used on the server to copy the data from xml to sheets as mentioned in the server input
' Input format "SourceSheet-SourceRange-DestSheet-DestCell;.... .
Public Sub PasteSheetsAndRanges(copySheetsAndRanges As String)
    Dim sheet As Variant
    Dim sourceDest() As String
    
    For Each sheet In Split(copySheetsAndRanges, ";")
        If Trim(sheet) <> "" Then
            sourceDest = Split(sheet, "~")
            PasteSheetRange sourceDest(0), sourceDest(1), sourceDest(2), sourceDest(3)
        End If
    Next

End Sub

' This macro is used to copy the data from the server (which is in xml format) to the client excel. This is also used on server side
' This macro internally uses the server xml data which is parsed and stored in global variable xmlDictionary
' Parameters:
' serverSheetName: The name of the sheet on the server. This is coming from the xml
' serverRange: The cell range on the server to copy the data. This data comes from xml
' clientSheetName: Name of the sheet on the client onto which server data to be copied
' clientStartCell: the starting cell to copy the server data
Public Sub PasteSheetRange(serverSheetName As String, serverRange As String, clientSheetName As String, clientStartCell As String)
    
    'Application.Calculation = xlCalculationManual

    Dim wkb As Workbook
    Dim wks As Worksheet
        
    Set wkb = ThisWorkbook
    Set wks = wkb.Sheets(clientSheetName)
    
    Dim sourceCells As range, destStartCell As range, cell As range, xmlCell As range
    Set sourceCells = wks.range(serverRange)
    Set destStartCell = wks.range(clientStartCell)
    
    If Not xmlDict.Exists(serverSheetName) Then
        Err.Raise 123, serverSheetName, "Sheet not found in XML"
        Exit Sub
    End If
     
    Dim dicSheet As Dictionary
    Set dicSheet = xmlDict(serverSheetName)
    
    Dim rowOffset As Integer, columnOffset As Integer
    rowOffset = sourceCells.Row
    columnOffset = sourceCells.Column
   
    'first store xml data into a 2D array and then later assign the array to destination range
    ReDim SourceData(1 To sourceCells.Rows.Count, 1 To sourceCells.Columns.Count) As String
    Dim cellAddress As String
    Dim cellValue As String
    
    For Each xmlCell In sourceCells.Cells
        cellAddress = Replace(xmlCell.address, "$", "")
        If dicSheet.Exists(cellAddress) Then
            cellValue = dicSheet(cellAddress)
        Else
            cellValue = ""
        End If
        SourceData(xmlCell.Row - rowOffset + 1, xmlCell.Column - columnOffset + 1) = cellValue
    Next
    
    destStartCell.Resize(sourceCells.Rows.Count, sourceCells.Columns.Count) = SourceData
End Sub




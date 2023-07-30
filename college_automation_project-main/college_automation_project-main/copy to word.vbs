Sub CopyToWord()
    ' Open the Word document
    Dim objWord As Object
    Dim objDoc As Object
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open("c:\users\krishoth\onedrive\desktop\rat college proj\SEM2_Result Analysis_Before _C SEC - Copy.doc")
    
    ' Copy the contents of the Excel sheet
    Dim objExcel As Object
    Dim objSheet As Object
    Set objExcel = CreateObject("Excel.Application")
    Set objSheet = objExcel.Workbooks.Open("C:\path\to\your\excel\workbook.xlsx").Sheets("Sheet1")
    objSheet.Range("A1:C10").Copy
    
    ' Paste the contents into the Word document
    objDoc.Activate
    objWord.Selection.GoTo wdGoToPage, wdGoToAbsolute, 3 ' Change 3 to the page number you want to paste to
    objWord.Selection.PasteExcelTable False, False, False ' Paste the table without changing the formatting
    
    ' Save and close the Word document
    objDoc.Save
    objDoc.Close
    objWord.Quit
End Sub

Sub InsertTextAfterHeaderInPage()
    'Declare variables
    Dim doc As Document
    Dim sec As Section
    Dim hdr As HeaderFooter
    Dim rng As Range
    
    'Set variables
    Set doc = ActiveDocument
    Set sec = doc.Sections(2) 'change the page number here
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    Set rng = hdr.Range
    
    'Move insertion point to end of header
    rng.Collapse wdCollapseEnd
    
    'Insert text after header
    rng.InsertAfter "Text to insert"
End Sub

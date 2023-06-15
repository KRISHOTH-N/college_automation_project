Sub CopyTableToWord()
    ' Open the Word document
    Dim objWord As Object
    Dim objDoc As Object
    Set objWord = CreateObject("Word.Application")
    MsgBox "!"


    Set objDoc = objWord.Documents.Open("C:\Users\krishoth\OneDrive\Desktop\RAT college proj\ABCD.doc")
    MsgBox "word opened"
    
    ' Go to the specific page in the Word document
    objDoc.Activate
    objWord.Selection.GoTo wdGoToPage, wdGoToAbsolute, 1 ' Change 3 to the page number you want to insert the table
    
    ' Insert text "master the blaster" with four lines space below the header

    objWord.Selection.TypeParagraph
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeParagraph 
    

    MsgBox "Excel part"
   ' Set the font size for the range
    Sheets("Sheet2").Range("A1:E10").Font.Size = 14
    
    ' Wrap the text in the range
    Sheets("Sheet2").Range("A1:E10").WrapText = True
     MsgBox "text wrap"
    
    Sheets("Sheet2").Range("A1:E10").Copy

    
    
    ' Paste the Excel table into the Word document with four lines space below the text
    objDoc.Activate
    objWord.Selection.PasteExcelTable False, False, False ' Paste the table without changing the formatting
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeParagraph
    objWord.Selection.TypeParagraph 
    
    ' Save and close the Word document
    objDoc.Save
    objDoc.Close
    objWord.Quit 
End Sub

Sub CenterTextInLine()
    MsgBox "invoked"
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Calculate the center position of the line
    Dim centerPosition As Long
    centerPosition = doc.PageSetup.PageWidth / 2
    
    ' Move the cursor to the center position and type the text
    Dim rng As Range
    Set rng = doc.Range(Start:=0, End:=0)
    rng.Collapse Direction:=wdCollapseEnd
    rng.MoveStartUntil Cset:=" ", Count:=centerPosition - rng.Start
    rng.Select
    Selection.TypeText Text:="master the blaster"
End Sub




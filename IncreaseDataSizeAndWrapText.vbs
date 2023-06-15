


Sub InsertTextBetweenHeaderAndTable()



    MsgBox "Hello, world!"

    'Declare variables
    Dim doc As Word.Document
    Dim sec As Section
    Dim tbl As Table
    Dim hdr As HeaderFooter
    Dim rng As Range
    
    'Set variables
    Set doc = ActiveDocument
    Set sec = doc.Sections(1)
    Set tbl = sec.Range.Tables(1)
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    Set rng = sec.Range
    
    'Insert text after header
    hdr.Range.Collapse wdCollapseEnd
    hdr.Range.InsertAfter vbNewLine & vbNewLine 'add 1 line space above text
    rng.Collapse wdCollapseEnd
    rng.InsertAfter "Master the blaster"
    rng.ParagraphFormat.Alignment = wdAlignParagraphCenter 'center text
    rng.InsertParagraphAfter 'add 1 line space below text
    
    'Format table
    tbl.Range.Paragraphs(1).SpaceBefore = 12 'add 1 line space above table
End Sub





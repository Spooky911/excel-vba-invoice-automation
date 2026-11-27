Attribute VB_Name = "InvoiceAutomation"
Option Explicit

Private Const FIRST_LINE As Long = 17
Private Const LAST_LINE As Long = 516
Private Const LINE_LIMIT As Long = LAST_LINE - FIRST_LINE + 1
Private Const INPUT_ERROR As Long = vbObjectError + 501

Public Sub GenerateInvoice()
    Dim invoice As Worksheet, order As Worksheet, settings As Worksheet
    Dim source As Variant, billTo As Variant, business As Variant, payment As Variant
    Dim items() As Variant
    Dim rowIndex As Long, sheetRow As Long, count As Long, col As Long, lastInput As Long
    Dim description As String, code As String, payIn3 As String, prefix As String
    Dim quantity As Double, price As Double, discount As Double
    Dim sequence As Long, terms As Long, issue As Date, invoiceNumber As String
    Dim oldScreenUpdating As Boolean
    Dim rowData(1 To 1, 1 To 3) As Variant

    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Set invoice = ThisWorkbook.Worksheets("Invoice")
    Set order = ThisWorkbook.Worksheets("Order Sheet")
    Set settings = ThisWorkbook.Worksheets("Settings")
    billTo = order.Range("B7:B14").Value2
    business = settings.Range("B6:B9").Value2
    payment = settings.Range("B13:B17").Value2

    If Len(EntryText(billTo(1, 1), "customer name")) = 0 Then
        InvalidInput "Enter a customer name on the Order Sheet."
    End If
    If Len(EntryText(business(1, 1), "business name")) = 0 Then
        InvalidInput "Enter a business name in Settings."
    End If
    For rowIndex = 1 To 8
        billTo(rowIndex, 1) = EntryText(billTo(rowIndex, 1), "customer details")
    Next rowIndex
    For rowIndex = 1 To 4
        business(rowIndex, 1) = EntryText(business(rowIndex, 1), "business details")
    Next rowIndex
    For rowIndex = 1 To 5
        payment(rowIndex, 1) = EntryText(payment(rowIndex, 1), "payment details")
    Next rowIndex

    discount = NumericEntry(order.Range("H7").Value2, "Discount on the Order Sheet", 0, 1, True)
    payIn3 = UCase$(EntryText(order.Range("H8").Value2, "Pay in 3"))
    Select Case payIn3
        Case "YES": payIn3 = "Yes"
        Case "NO", "": payIn3 = "No"
        Case Else: InvalidInput "Choose Yes or No for Pay in 3 on the Order Sheet."
    End Select
    terms = WholeEntry(settings.Range("B18").Value2, "Payment term in Settings", 0, 365)
    sequence = WholeEntry(settings.Range("B20").Value2, "Last invoice sequence in Settings", 0, 99999999) + 1
    prefix = UCase$(EntryText(settings.Range("B19").Value2, "invoice prefix"))
    If Len(prefix) = 0 Or Len(prefix) > 8 Or prefix Like "*[!A-Z0-9]*" Then
        InvalidInput "Enter an invoice prefix of 1 to 8 letters or numbers in Settings."
    End If

    For col = 1 To 4
        lastInput = order.Cells(order.Rows.Count, col).End(xlUp).Row
        If lastInput > LAST_LINE Then InvalidInput "The Order Sheet supports 500 line items. Remove entries below row 516."
    Next col
    source = order.Range("A17:D516").Value2
    ReDim items(1 To LINE_LIMIT, 1 To 4)
    For rowIndex = 1 To UBound(source, 1)
        sheetRow = rowIndex + FIRST_LINE - 1
        description = EntryText(source(rowIndex, 1), "description at row " & sheetRow)
        code = EntryText(source(rowIndex, 2), "product code at row " & sheetRow)
        If Len(description & code) > 0 Or HasEntry(source(rowIndex, 3)) Or HasEntry(source(rowIndex, 4)) Then
            If Len(description) = 0 Or Len(code) = 0 Then
                InvalidInput "Enter both a description and a product code on Order Sheet row " & sheetRow & "."
            End If
            quantity = NumericEntry(source(rowIndex, 3), "Quantity at row " & sheetRow, 0.001, 1000000)
            price = NumericEntry(source(rowIndex, 4), "Unit price at row " & sheetRow, 0, 1000000)
            If Abs(quantity - Application.Round(quantity, 3)) > 0.000000001 Then
                InvalidInput "Use up to three quantity decimal places on Order Sheet row " & sheetRow & "."
            End If
            If Abs(price - Application.Round(price, 2)) > 0.000000001 Then
                InvalidInput "Use up to two unit-price decimal places on Order Sheet row " & sheetRow & "."
            End If
            count = count + 1
            items(count, 1) = description
            items(count, 2) = code
            items(count, 3) = quantity
            items(count, 4) = price
        End If
    Next rowIndex
    If count = 0 Then InvalidInput "Add at least one product row on the Order Sheet."

    issue = Date
    invoiceNumber = prefix & "-" & Format$(issue, "yyyymmdd") & "-" & Format$(sequence, "0000")
    Application.ScreenUpdating = False
    invoice.Range("A17:F516").ClearContents
    invoice.Rows(FIRST_LINE & ":" & LAST_LINE).Hidden = True
    For rowIndex = 1 To count
        sheetRow = FIRST_LINE + rowIndex - 1
        With invoice.Range("A" & sheetRow & ":B" & sheetRow)
            If Not .MergeCells Then .Merge
        End With
        WriteText invoice.Cells(sheetRow, 1), CStr(items(rowIndex, 1))
        rowData(1, 1) = ExcelText(CStr(items(rowIndex, 2)))
        rowData(1, 2) = items(rowIndex, 3)
        rowData(1, 3) = items(rowIndex, 4)
        invoice.Cells(sheetRow, 3).NumberFormat = "@"
        invoice.Cells(sheetRow, 3).Resize(1, 3).Value2 = rowData
    Next rowIndex
    With invoice.Range("A17:F" & FIRST_LINE + count - 1)
        .EntireRow.Hidden = False
        .RowHeight = 18
    End With
    With invoice.Range("A17:C" & FIRST_LINE + count - 1)
        .ShrinkToFit = True
    End With
    invoice.Range("E17:F" & FIRST_LINE + count - 1).ShrinkToFit = True
    invoice.Range("F17:F" & FIRST_LINE + count - 1).FormulaR1C1 = "=ROUND(RC[-2]*RC[-1],2)"

    WriteText invoice.Range("A2"), business(1, 1)
    WriteText invoice.Range("A3"), business(2, 1)
    WriteText invoice.Range("A4"), business(3, 1)
    If Len(business(4, 1)) > 0 Then
        WriteText invoice.Range("E4"), "VAT registration: " & business(4, 1)
    Else
        WriteText invoice.Range("E4"), ""
    End If
    For rowIndex = 1 To 8
        WriteText invoice.Cells(6 + rowIndex, 2), billTo(rowIndex, 1)
    Next rowIndex
    WriteText invoice.Range("F7"), invoiceNumber
    invoice.Range("F8").Value = issue
    invoice.Range("F9").Value = DateAdd("d", terms, issue)
    invoice.Range("H519").Value2 = discount
    WriteText invoice.Range("C527"), payIn3
    For rowIndex = 1 To 5
        WriteText invoice.Cells(527 + rowIndex, 3), payment(rowIndex, 1)
    Next rowIndex
    invoice.Calculate
    settings.Range("B20").Value2 = sequence
    invoice.Activate
    Application.Goto invoice.Range("A1"), True
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

Failed:
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox Err.Description, vbExclamation, "Generate Invoice"
End Sub

Public Sub EmailInvoice()
    Dim invoice As Worksheet
    Dim recipient As String, invoiceNumber As String, greeting As String, business As String
    Dim pdfPath As String, outlook As Object, message As Object

    On Error GoTo ExportFailed
    Set invoice = ThisWorkbook.Worksheets("Invoice")
    invoiceNumber = EntryText(invoice.Range("F7").Value2, "invoice number")
    If Len(invoiceNumber) = 0 Then InvalidInput "Generate an invoice before opening an email draft."
    recipient = EntryText(invoice.Range("B13").Value2, "customer email")
    If Not LooksLikeEmail(recipient) Then
        InvalidInput "Enter one valid email address on the Order Sheet, then generate the invoice again."
    End If
    greeting = EntryText(invoice.Range("B7").Value2, "customer name")
    business = EntryText(invoice.Range("A2").Value2, "business name")
    pdfPath = NewPdfPath(invoiceNumber)
    WriteInvoicePdf pdfPath

    On Error GoTo OutlookFailed
    Set outlook = CreateObject("Outlook.Application")
    Set message = outlook.CreateItem(0)
    With message
        .To = recipient
        .Subject = "Invoice " & invoiceNumber
        .Body = "Hello " & greeting & "," & vbCrLf & vbCrLf & _
                "Please find invoice " & invoiceNumber & " attached." & vbCrLf & vbCrLf & _
                "Kind regards," & vbCrLf & business
        .Attachments.Add pdfPath
        .Display
    End With
    Exit Sub

OutlookFailed:
    MsgBox "Excel saved the invoice PDF. Attach it in your email app:" & vbCrLf & vbCrLf & _
           pdfPath, vbInformation, "Email Invoice"
    Exit Sub

ExportFailed:
    MsgBox Err.Description, vbExclamation, "Email Invoice"
End Sub

Private Sub WriteInvoicePdf(ByVal pdfPath As String)
    Dim originalMapping As Boolean
    Dim failureNumber As Long, failureSource As String, failureDescription As String

    originalMapping = Application.MapPaperSize
    On Error GoTo Failed
    ' Prevent the desktop's regional paper preference from converting A4 to Letter.
    Application.MapPaperSize = False
    ThisWorkbook.Worksheets("Invoice").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Application.MapPaperSize = originalMapping
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    Application.MapPaperSize = originalMapping
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Private Function NewPdfPath(ByVal invoiceNumber As String) As String
    Dim files As Object, root As String, draft As String, cleanNumber As String
    Dim charIndex As Long, letter As String
    Set files = CreateObject("Scripting.FileSystemObject")
    root = files.BuildPath(files.GetSpecialFolder(2).Path, "Excel Invoice Drafts")
    If Not files.FolderExists(root) Then files.CreateFolder root
    draft = files.BuildPath(root, files.GetTempName)
    files.CreateFolder draft
    For charIndex = 1 To Len(invoiceNumber)
        letter = Mid$(invoiceNumber, charIndex, 1)
        If letter Like "[A-Za-z0-9_-]" Then cleanNumber = cleanNumber & letter
    Next charIndex
    NewPdfPath = files.BuildPath(draft, "Invoice " & cleanNumber & ".pdf")
End Function

Private Function LooksLikeEmail(ByVal address As String) As Boolean
    Dim atSign As Long, lastDot As Long
    atSign = InStr(address, "@")
    lastDot = InStrRev(address, ".")
    If atSign < 2 Or lastDot < atSign + 2 Or lastDot >= Len(address) Then Exit Function
    If InStr(atSign + 1, address, "@") > 0 Then Exit Function
    If address Like "*[ ,;<>]*" Or InStr(address, vbCr) > 0 Or InStr(address, vbLf) > 0 Then Exit Function
    LooksLikeEmail = True
End Function

Private Function HasEntry(ByVal value As Variant) As Boolean
    If IsError(value) Then
        HasEntry = True
    Else
        HasEntry = Len(Trim$(CStr(value))) > 0
    End If
End Function

Private Function EntryText(ByVal value As Variant, ByVal label As String) As String
    If IsError(value) Then InvalidInput "Fix the Excel error in " & label & "."
    EntryText = Trim$(CStr(value))
End Function

Private Function NumericEntry(ByVal value As Variant, ByVal label As String, _
                              ByVal minimum As Double, ByVal maximum As Double, _
                              Optional ByVal allowBlank As Boolean = False) As Double
    If allowBlank And IsEmpty(value) Then Exit Function
    If Not Application.WorksheetFunction.IsNumber(value) Then InvalidInput label & " must be a number."
    If value < minimum Or value > maximum Then
        InvalidInput label & " is outside its allowed range (" & minimum & " to " & maximum & ")."
    End If
    NumericEntry = CDbl(value)
End Function

Private Function WholeEntry(ByVal value As Variant, ByVal label As String, _
                            ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim number As Double
    number = NumericEntry(value, label, minimum, maximum)
    If number <> Fix(number) Then InvalidInput label & " must be a whole number."
    WholeEntry = CLng(number)
End Function

Private Sub InvalidInput(ByVal reason As String)
    Err.Raise INPUT_ERROR, "InvoiceAutomation", reason
End Sub

Private Function ExcelText(ByVal value As String) As String
    If Len(value) = 0 Then Exit Function
    ExcelText = "'" & value
End Function

Private Sub WriteText(ByVal target As Range, ByVal value As String)
    target.NumberFormat = "@"
    target.Value2 = ExcelText(value)
End Sub

Private Sub PrepareInvoiceTemplate()
    Dim invoice As Worksheet, page As PageSetup
    Dim sampleCount As Long
    Set invoice = ThisWorkbook.Worksheets("Invoice")
    Application.ScreenUpdating = False
    On Error GoTo Failed
    invoice.Range("A17:F516").Interior.Color = RGB(246, 250, 251)
    invoice.Range("A17:F516").FormatConditions.Add _
        Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0"
    invoice.Range("A17:F516").FormatConditions(1).Interior.Color = RGB(228, 241, 244)
    sampleCount = Application.Count(invoice.Range("D17:D516"))
    invoice.Rows(FIRST_LINE + sampleCount & ":" & LAST_LINE).Hidden = True
    invoice.Range("A2:D4").ShrinkToFit = True
    invoice.Range("E4:F4").ShrinkToFit = True
    invoice.Range("B7:D14").ShrinkToFit = True
    invoice.Range("F7:F9").ShrinkToFit = True
    invoice.Range("F518:F521").ShrinkToFit = True
    invoice.Range("C527:F532").ShrinkToFit = True
    invoice.Rows("522:524").RowHeight = 4
    invoice.Range("H519").NumberFormat = ";;;"
    invoice.Range("H13").Value2 = "Save to retain the invoice sequence."
    Set page = invoice.PageSetup
    With page
        .PrintArea = "$A$1:$F$532"
        .PrintTitleRows = "$16:$16"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = 26
        .RightMargin = 26
        .TopMargin = 30
        .BottomMargin = 32
        .PrintGridlines = False
        .PrintHeadings = False
        .CenterFooter = "&P / &N"
    End With
    AddAction invoice, "GenerateInvoiceButton", "Generate Invoice", "GenerateInvoice", "H2", RGB(29, 101, 128)
    AddAction invoice, "EmailInvoiceButton", "Email Invoice", "EmailInvoice", "H4", RGB(41, 107, 152)
    invoice.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 95
    Application.Goto invoice.Range("A1"), True
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Prepare Invoice Template"
End Sub

Private Sub AddAction(ByVal invoice As Worksheet, ByVal shapeName As String, _
                      ByVal caption As String, ByVal action As String, _
                      ByVal anchor As String, ByVal color As Long)
    Dim button As Shape, position As Range
    Set position = invoice.Range(anchor)
    Set button = invoice.Shapes.AddShape(msoShapeRoundedRectangle, position.Left, _
                                         position.Top, position.Width, 22)
    With button
        .Name = shapeName
        .AlternativeText = caption
        .OnAction = action
        .Fill.ForeColor.RGB = color
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Name = "Arial"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Placement = xlFreeFloating
    End With
End Sub

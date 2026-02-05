# Automated-Invoice
<a href="https://drive.google.com/file/d/16_JaxCKoL-fvC35SjxS5Pm34jpArSTrK/view?usp=sharing"> click here for video</a>
🧾 Invoice Generator (Excel + VBA Automation)
This project is a simple yet powerful Invoice Generator tool built using Excel and VBA. With just one click, it automatically fills the invoice template, generates a PDF, and saves it to a specified folder — perfect for freelancers, small businesses, or anyone who needs to create multiple invoices quickly.

⚙️ How It Works
1.Invoice Template Setup
   Start by selecting the invoice template and customize it based on your business needs:
    Add your logo, company details, and desired styling
    Make sure all invoice data (client, amount, items, etc.) is placed in clearly defined cells
    Set the template to fit on one page for clean PDF export
2.VBA Automation Code
  The VBA code:
  Fills the invoice fields automatically
  Exports the invoice as a PDF file
  Saves the PDF to a predefined folder path 
  (You can change the path in the VBA script to match your preferred location.)
3.One-Click Output
  After setting everything up, simply click the “Generate Invoice” button to:
  Fill the invoice
  Export it as a PDF
  Automatically save the PDF file with a unique name (like Invoice_001.pdf)

📂 Files Included
   Invoice Template.xlsx – Editable invoice layout
   Invoice_Automation.bas – VBA module with full code
   Sample Output.pdf – A sample generated invoice (PDF)

🔐 Customization Tips
    Update your logo and branding in the template
    Modify cell ranges in the VBA code if your layout is different
    Adjust file path and file name formatting as needed
    
    Code:
     Sub GenerateFinalInvoices3()

    Dim wsData As Worksheet, wsTemplate As Worksheet, newInvoice As Worksheet
    Dim dict As Object, companyName As Variant
    Dim lastRow As Long, i As Long, invoiceCount As Integer
    Dim companyRows As Variant, invoiceDate As String
    Dim entryRow As Variant, itemRow As Long, pdfPath As String, invoiceSheetName As String
    Dim discountTotal As Double, shippingTotal As Double

    ' Set your worksheet names
    Set wsData = ThisWorkbook.Sheets("DataEntry")
    Set wsTemplate = ThisWorkbook.Sheets("InvoiceTemplate")
    Set dict = CreateObject("Scripting.Dictionary")
    invoiceDate = Format(wsData.Range("B2").Value, "dd-mm-yyyy")

    lastRow = wsData.Cells(wsData.Rows.Count, "E").End(xlUp).Row ' Column E = Company Name
    invoiceCount = 1

    ' Step 1: Collect all rows by company name
    For i = 2 To lastRow
        companyName = Trim(wsData.Cells(i, "E").Value)
        If Len(companyName) > 0 Then
            If Not dict.exists(companyName) Then
                dict.Add companyName, Array(i)
            Else
                companyRows = dict(companyName)
                ReDim Preserve companyRows(UBound(companyRows) + 1)
                companyRows(UBound(companyRows)) = i
                dict(companyName) = companyRows
            End If
        End If
    Next i

    ' Step 2: Loop through each company group
    For Each companyName In dict.Keys
        wsTemplate.Copy After:=Worksheets(Worksheets.Count)
        DoEvents
        Set newInvoice = ActiveSheet

        invoiceSheetName = "Invoice " & invoiceCount

        ' Check if sheet name exists
        On Error Resume Next
        If Not ThisWorkbook.Sheets(invoiceSheetName) Is Nothing Then
            invoiceSheetName = invoiceSheetName & "_" & Format(Now, "hhmmss")
        End If
        On Error GoTo 0

        newInvoice.Name = invoiceSheetName

        ' Fill header info from first matching row
        entryRow = dict(companyName)(0)
        With newInvoice
            .Range("B14").Value = wsData.Cells(entryRow, 4).Value  ' Contact No
            .Range("B15").Value = companyName                     ' Client Company
            .Range("B16").Value = wsData.Cells(entryRow, 6).Value  ' Address
            .Range("B17").Value = wsData.Cells(entryRow, 7).Value  ' Phone
            .Range("B18").Value = wsData.Cells(entryRow, 8).Value ' Email
            .Range("G6").Value = invoiceDate
            .Range("G8").Value = wsData.Cells(entryRow, 3).Value 'Invoice No
            ' Reset totals
      discountTotal = 0
      shippingTotal = 0
      totaltaxTotal = 0

    ' Loop through this company’s rows and sum properly
    For Each entryRow In dict(companyName)
    discountTotal = discountTotal + wsData.Cells(entryRow, "N").Value
    shippingTotal = shippingTotal + wsData.Cells(entryRow, "R").Value
    totaltaxTotal = totaltaxTotal + wsData.Cells(entryRow, "Q").Value
    Next entryRow

     .Range("G34").Value = discountTotal
     .Range("G38").Value = shippingTotal
     .Range("G37").Value = totaltaxTotal
           
            ' Add more mapping here if needed for line items

            
            
        End With

        ' Fill line items for this company
        itemRow = 22
        For Each entryRow In dict(companyName)
            With newInvoice
                .Cells(itemRow, 2).Value = wsData.Cells(entryRow, 9).Value  ' Description
                .Cells(itemRow, 5).Value = wsData.Cells(entryRow, 10).Value ' Qty
                .Cells(itemRow, 6).Value = wsData.Cells(entryRow, 11).Value ' Unit Price
                .Cells(itemRow, 7).Value = wsData.Cells(entryRow, 12).Value 'Total
                End With
            itemRow = itemRow + 1
        Next entryRow

        ' Step 3: Export to PDF
        pdfPath = "C:\Rinku\Linkedin Post\Invoice PDF\" & invoiceSheetName & ".pdf"
        newInvoice.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard

        invoiceCount = invoiceCount + 1
    Next companyName

    MsgBox "All invoices and PDFs have been created!", vbInformation

End Sub



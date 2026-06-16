Option Explicit

' ===========================================================
'   CONFIGURATION - edit these, not the logic below
' ===========================================================

' CRF email source
Public Const CRF_SENDER As String = "donotreply@amat.com"

' Change Coordinator exclusion
Public Const CHANGE_COORDINATOR_EMAIL As String = "change_coordinator@ichorsystems.com"

' Authorized senders for Spec Award emails
Public SPEC_AWARD_SENDERS(3) As String

' Forward recipients - defined once, used for both CRF and Spec Award
Public Const FORWARD_TO As String = "sreshmi@ichorsystems.com; schiniwar@ichorsystems.com; ukallibaddi@ichorsystems.com"
Public Const FORWARD_CC As String = "mishak@ichorsystems.com; amustrada@ichorsystems.com; kphua@ichorsystems.com; tliew@ichorsystems.com; singaporecustomerservice@ichorsystems.com; echarlie@ichorsystems.com"
Public Const FORWARD_BCC As String = "kmageshkumar@ichorsystems.com"

' Minimum body length (chars) used when deciding whether latestBody is long enough
' to use as the forward content. Does NOT affect keyword checking.
Public Const MIN_BODY_LENGTH As Integer = 200

' Spec Award staging workbook - Outlook writes the latest detected rows here.
Public Const SPEC_AWARD_STAGING_FILE As String = "C:\Users\kmageshkumar\Downloads\AMAT SGP ECO Tracker.xlsx"
Public Const SPEC_AWARD_STAGING_SHEET As String = "Sheet1"

' Python sync command - this reuses excel_to_sharepoint.py in local-file mode.
Public Const SYNC_PYTHON_COMMAND As String = "py"
Public Const SYNC_SCRIPT_PATH As String = "C:\Users\kmageshkumar\OneDrive - Ichor Systems\Scripts\ECO Tracker\excel_to_sharepoint.py"
Public Const SYNC_TARGET_WORKBOOK_FILE As String = "C:\Users\kmageshkumar\OneDrive - Ichor Systems\AMAT SGP ECO Tracker.xlsx"
Public Const SYNC_TARGET_TABLE_NAME As String = "Table1"

' ===========================================================
'   RUNTIME STATE
' ===========================================================

Public CRFSentTracker As Object       ' Scripting.Dictionary keyed on CRF number
Public SpecAwardSentTracker As Object ' Scripting.Dictionary keyed on EntryID
Public CRFRegex As Object             ' VBScript.RegExp - initialized once

' ===========================================================
'   INIT: Populate authorized Spec Award senders array
' ===========================================================

Public Sub InitSpecAwardSenders()
    SPEC_AWARD_SENDERS(0) = "faith_fan@amat.com"
    SPEC_AWARD_SENDERS(1) = "michael_leow@amat.com"
    SPEC_AWARD_SENDERS(2) = "jason_tai@amat.com"
    SPEC_AWARD_SENDERS(3) = "francesca_chang@amat.com"
End Sub

' ===========================================================
'   INIT: Pre-warm CRF tracker from Sent Items on startup
' ===========================================================

Public Sub RebuildCRFTrackerFromSentItems()
    On Error GoTo ErrorHandler

    If CRFSentTracker Is Nothing Then
        Set CRFSentTracker = CreateObject("Scripting.Dictionary")
    End If

    If CRFRegex Is Nothing Then
        Set CRFRegex = CreateObject("VBScript.RegExp")
        CRFRegex.Pattern = "CRF\s*:\s*(\d{5,})"
        CRFRegex.IgnoreCase = True
    End If

    Dim sentFolder As Outlook.Folder
    Dim filteredItems As Outlook.Items
    Dim obj As Object
    Dim Mail As Outlook.MailItem
    Dim crfKey As String
    Dim count As Long

    count = 0
    Set sentFolder = Application.Session.GetDefaultFolder(olFolderSentMail)

    Set filteredItems = sentFolder.Items.Restrict( _
        "@SQL=""urn:schemas:httpmail:subject"" LIKE '%CRF%'")

    For Each obj In filteredItems
        If obj.Class = olMail Then
            Set Mail = obj
            If CRFRegex.Test(Mail.Subject) Then
                crfKey = CRFRegex.Execute(Mail.Subject)(0).SubMatches(0)
                If Not CRFSentTracker.Exists(crfKey) Then
                    CRFSentTracker.Add crfKey, True
                    count = count + 1
                End If
            End If
        End If
    Next obj

    LogEvent "CRF tracker rebuilt from Sent Items. " & count & " unique CRFs cached."
    Exit Sub

ErrorHandler:
    LogEvent "WARNING in RebuildCRFTrackerFromSentItems: " & Err.Number & " - " & Err.Description & _
             ". Deduplication will rely on in-session tracking only."
End Sub

' ===========================================================
'   DEDUPLICATION: Has this CRF already been forwarded?
' ===========================================================

Public Function HasCRFBeenProcessed(crfKey As String) As Boolean
    If Not CRFSentTracker Is Nothing Then
        If CRFSentTracker.Exists(crfKey) Then
            HasCRFBeenProcessed = True
            Exit Function
        End If
    End If

    Dim localRegex As Object
    Set localRegex = CreateObject("VBScript.RegExp")
    localRegex.Pattern = "CRF\s*:\s*" & crfKey
    localRegex.IgnoreCase = True

    Dim outboxFolder As Outlook.Folder
    Dim filteredItems As Outlook.Items
    Dim obj As Object
    Dim Mail As Outlook.MailItem

    On Error GoTo NotFound
    Set outboxFolder = Application.Session.GetDefaultFolder(olFolderOutbox)
    Set filteredItems = outboxFolder.Items.Restrict( _
        "@SQL=""urn:schemas:httpmail:subject"" LIKE '%CRF%'")

    For Each obj In filteredItems
        If obj.Class = olMail Then
            Set Mail = obj
            If localRegex.Test(Mail.Subject) Then
                HasCRFBeenProcessed = True
                Exit Function
            End If
        End If
    Next obj

NotFound:
    HasCRFBeenProcessed = False
End Function

' ===========================================================
'   FORWARD: CRF email
' ===========================================================

Public Sub ForwardCRFEmail(ByVal Mail As Outlook.MailItem, ByVal crfKey As String)
    On Error GoTo ErrorHandler

    Dim fwdMail As Outlook.MailItem
    Set fwdMail = Mail.Forward

    Dim bodyIntro As String
    bodyIntro = "Hi," & vbCrLf & vbCrLf & _
                "Please help to sync CRF " & crfKey & " to Agile. Thanks." & vbCrLf & vbCrLf & _
                "Best regards," & vbCrLf & "Sankar"

    fwdMail.HTMLBody = "<div style='font-family:Verdana; font-size:11pt;'>" & _
                       Replace(bodyIntro, vbCrLf, "<br>") & _
                       "<hr>" & fwdMail.HTMLBody & "</div>"

    fwdMail.To = FORWARD_TO
    fwdMail.CC = FORWARD_CC
    fwdMail.BCC = FORWARD_BCC
    fwdMail.Send

    LogEvent "CRF " & crfKey & " forwarded successfully."
    Exit Sub

ErrorHandler:
    LogEvent "ERROR in ForwardCRFEmail (CRF " & crfKey & "): " & Err.Number & " - " & Err.Description
End Sub

' ===========================================================
'   MARK RELATED MAILS READ
' ===========================================================

Public Sub MarkRelatedMailsRead(ByVal crfKey As String)
    On Error GoTo ErrorHandler

    Dim Inbox As Outlook.Folder
    Dim filteredItems As Outlook.Items
    Dim obj As Object
    Dim Mail As Outlook.MailItem

    Set Inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    Set filteredItems = Inbox.Items.Restrict( _
        "@SQL=""urn:schemas:httpmail:subject"" LIKE '%CRF%'")

    Dim markRegex As Object
    Set markRegex = CreateObject("VBScript.RegExp")
    markRegex.Pattern = "CRF\s*:\s*" & crfKey
    markRegex.IgnoreCase = True

    For Each obj In filteredItems
        If obj.Class = olMail Then
            Set Mail = obj
            If Mail.UnRead And markRegex.Test(Mail.Subject) Then
                Mail.UnRead = False
                Mail.Save
            End If
        End If
    Next obj

    Exit Sub

ErrorHandler:
    LogEvent "ERROR in MarkRelatedMailsRead (CRF " & crfKey & "): " & Err.Number & " - " & Err.Description
End Sub

' ===========================================================
'   SPEC AWARD: EXTRACT, STAGE, SYNC
' ===========================================================

Public Sub ProcessSpecAwardData(ByVal latestHtml As String, ByVal subjectText As String, ByVal receivedTime As Date)
    On Error GoTo ErrorHandler

    Dim rows As Collection
    Set rows = ExtractSpecAwardRowsFromHtml(latestHtml, receivedTime)

    If rows Is Nothing Or rows.Count = 0 Then
        LogEvent "Spec Award data extraction found no table rows: " & subjectText
        Exit Sub
    End If

    SaveSpecAwardRowsToWorkbook rows, SPEC_AWARD_STAGING_FILE, SPEC_AWARD_STAGING_SHEET
    LogEvent "Spec Award staging workbook updated with " & rows.Count & " row(s)."

    RunSpecAwardSync
    Exit Sub

ErrorHandler:
    LogEvent "ERROR in ProcessSpecAwardData: " & Err.Number & " - " & Err.Description & _
             " | Subject: " & subjectText
End Sub

Public Function ExtractSpecAwardRowsFromHtml(ByVal html As String, ByVal receivedTime As Date) As Collection
    On Error GoTo ErrorHandler

    Dim doc As Object
    Dim tables As Object
    Dim tableNode As Object
    Dim rows As Collection

    Set doc = CreateObject("htmlfile")
    doc.Open
    doc.Write html
    doc.Close

    Set tables = doc.getElementsByTagName("table")

    For Each tableNode In tables
        Set rows = ReadSpecAwardTableRows(tableNode, receivedTime)
        If Not rows Is Nothing Then
            If rows.Count > 0 Then
                Set ExtractSpecAwardRowsFromHtml = rows
                Exit Function
            End If
        End If
    Next tableNode

    Set ExtractSpecAwardRowsFromHtml = New Collection
    Exit Function

ErrorHandler:
    LogEvent "ERROR in ExtractSpecAwardRowsFromHtml: " & Err.Number & " - " & Err.Description
    Set ExtractSpecAwardRowsFromHtml = New Collection
End Function

Private Function ReadSpecAwardTableRows(ByVal tableNode As Object, ByVal receivedTime As Date) As Collection
    Dim rowNodes As Object
    Set rowNodes = tableNode.Rows
    If rowNodes.Length = 0 Then Exit Function

    Dim headerMap As Object
    Set headerMap = GetSpecAwardHeaderMap(rowNodes.Item(0))
    If headerMap Is Nothing Then Exit Function

    Dim results As New Collection
    Dim rowIndex As Long

    For rowIndex = 1 To rowNodes.Length - 1
        Dim rowNode As Object
        Set rowNode = rowNodes.Item(rowIndex)

        Dim cellValues As Collection
        Set cellValues = GetRowCellValues(rowNode)
        If cellValues.Count = 0 Then GoTo ContinueRow

        Dim slotNumber As String
        Dim partNumber As String
        slotNumber = GetMappedCellValue(cellValues, headerMap, "slotnumber")
        partNumber = GetMappedCellValue(cellValues, headerMap, "partnumber")

        If slotNumber <> "" And partNumber <> "" Then
            Dim rowData As Object
            Set rowData = CreateObject("Scripting.Dictionary")
            rowData("System Number") = BuildSystemNumber(slotNumber, partNumber)
            rowData("Spec Award Date") = Format$(receivedTime, "m/d/yyyy")
            results.Add rowData
        End If

ContinueRow:
    Next rowIndex

    Set ReadSpecAwardTableRows = results
End Function

Private Function BuildSystemNumber(ByVal slotNumber As String, ByVal partNumber As String) As String
    Dim cleanedSlot As String
    cleanedSlot = CleanCellText(slotNumber)
    If cleanedSlot = "" Then Exit Function

    Dim cleanedPart As String
    cleanedPart = CleanCellText(partNumber)
    If cleanedPart = "" Then Exit Function

    If InStr(1, cleanedPart, "SEMSYS", vbTextCompare) > 0 Then
        BuildSystemNumber = Replace(cleanedPart, "SEMSYS", cleanedSlot, 1, 1, vbTextCompare)
    ElseIf InStr(1, cleanedPart, "SEMNSO", vbTextCompare) > 0 Then
        BuildSystemNumber = Replace(cleanedPart, "SEMNSO", cleanedSlot, 1, 1, vbTextCompare)
    Else
        BuildSystemNumber = cleanedPart
    End If
End Function

Private Function GetSpecAwardHeaderMap(ByVal headerRow As Object) As Object
    Dim cells As Collection
    Set cells = GetRowCellValues(headerRow)
    If cells.Count = 0 Then Exit Function

    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To cells.Count
        Dim normalizedHeader As String
        normalizedHeader = NormalizeHeaderText(CStr(cells(i)))
        If normalizedHeader <> "" Then
            headerMap(normalizedHeader) = i
        End If
    Next i

    If headerMap.Exists("slotnumber") And headerMap.Exists("partnumber") Then
        Set GetSpecAwardHeaderMap = headerMap
    End If
End Function

Private Function GetRowCellValues(ByVal rowNode As Object) As Collection
    Dim values As New Collection
    Dim cellIndex As Long

    For cellIndex = 0 To rowNode.Cells.Length - 1
        values.Add CleanCellText(rowNode.Cells.Item(cellIndex).innerText)
    Next cellIndex

    Set GetRowCellValues = values
End Function

Private Function GetMappedCellValue(ByVal cellValues As Collection, ByVal headerMap As Object, ByVal key As String) As String
    If headerMap.Exists(key) Then
        Dim indexValue As Long
        indexValue = CLng(headerMap(key))
        If indexValue >= 1 And indexValue <= cellValues.Count Then
            GetMappedCellValue = CStr(cellValues(indexValue))
        End If
    End If
End Function

Private Function NormalizeHeaderText(ByVal inputText As String) As String
    Dim normalized As String
    normalized = LCase(CollapseWhitespace(inputText))
    normalized = Replace(normalized, " ", "")
    normalized = Replace(normalized, "-", "")
    normalized = Replace(normalized, "/", "")
    NormalizeHeaderText = normalized
End Function

Private Function CleanCellText(ByVal inputText As String) As String
    Dim cleaned As String
    cleaned = Replace(inputText, vbCr, " ")
    cleaned = Replace(cleaned, vbLf, " ")
    cleaned = Replace(cleaned, Chr(160), " ")
    CleanCellText = CollapseWhitespace(cleaned)
End Function

Private Function CollapseWhitespace(ByVal inputText As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\s+"
    CollapseWhitespace = Trim(regex.Replace(inputText, " "))
End Function

Public Sub SaveSpecAwardRowsToWorkbook(ByVal rows As Collection, ByVal workbookPath As String, ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim headers As Variant
    Dim i As Long
    Dim rowIndex As Long

    headers = Array("System Number", "Spec Award Date")

    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False
    xlApp.Visible = False

    Dim workbookExists As Boolean
    workbookExists = (Dir(workbookPath) <> "")

    If workbookExists Then
        Set wb = xlApp.Workbooks.Open(workbookPath)
    Else
        Set wb = xlApp.Workbooks.Add
    End If

    Set ws = GetOrCreateWorksheet(wb, sheetName)

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    Dim existingSystemNumbers As Object
    Set existingSystemNumbers = CreateObject("Scripting.Dictionary")

    Dim lastUsedRow As Long
    lastUsedRow = GetLastUsedRow(ws, 1, 2)

    For rowIndex = 2 To lastUsedRow
        Dim existingSystemNumber As String
        existingSystemNumber = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
        If existingSystemNumber <> "" Then
            existingSystemNumbers(existingSystemNumber) = True
        End If
    Next rowIndex

    Dim nextRow As Long
    nextRow = IIf(lastUsedRow < 2, 2, lastUsedRow + 1)

    For rowIndex = 1 To rows.Count
        Dim rowData As Object
        Set rowData = rows(rowIndex)
        Dim systemNumber As String
        systemNumber = Trim$(CStr(rowData("System Number")))
        If systemNumber <> "" Then
            If Not existingSystemNumbers.Exists(systemNumber) Then
                ws.Cells(nextRow, 1).Value = systemNumber
                ws.Cells(nextRow, 2).Value = rowData("Spec Award Date")
                existingSystemNumbers(systemNumber) = True
                nextRow = nextRow + 1
            End If
        End If
    Next rowIndex

    If workbookExists Then
        wb.Save
    Else
        wb.SaveAs workbookPath
    End If
    wb.Close False
    xlApp.Quit
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    On Error GoTo 0
    Err.Raise Err.Number, , "SaveSpecAwardRowsToWorkbook failed: " & Err.Description
End Sub

Private Function GetLastUsedRow(ByVal ws As Object, ByVal firstColumn As Long, ByVal lastColumn As Long) As Long
    Dim lastRow As Long
    lastRow = 1

    Dim colIndex As Long
    For colIndex = firstColumn To lastColumn
        Dim currentLastRow As Long
        currentLastRow = ws.Cells(ws.Rows.Count, colIndex).End(-4162).Row
        If currentLastRow > lastRow Then
            lastRow = currentLastRow
        End If
    Next colIndex

    GetLastUsedRow = lastRow
End Function

Private Function GetOrCreateWorksheet(ByVal wb As Object, ByVal sheetName As String) As Object
    On Error Resume Next
    Set GetOrCreateWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = wb.Worksheets.Add
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

Public Sub RunSpecAwardSync()
    On Error GoTo ErrorHandler

    Dim shellObj As Object
    Dim commandText As String
    Dim exitCode As Long

    commandText = QuoteArg(SYNC_PYTHON_COMMAND) & " " & _
                  QuoteArg(SYNC_SCRIPT_PATH) & " " & _
                  "--target-workbook-file " & QuoteArg(SYNC_TARGET_WORKBOOK_FILE) & " " & _
                  "--table-name " & QuoteArg(SYNC_TARGET_TABLE_NAME) & " " & _
                  "--excel-file " & QuoteArg(SPEC_AWARD_STAGING_FILE) & " " & _
                  "--excel-sheet " & QuoteArg(SPEC_AWARD_STAGING_SHEET) & " " & _
                  "--retry-on-conflict " & _
                  "--retry-delay-minutes 15 " & _
                  "--clear-source-on-success"

    LogEvent "Starting SharePoint sync."
    Set shellObj = CreateObject("WScript.Shell")
    exitCode = shellObj.Run(commandText, 0, True)

    If exitCode <> 0 Then
        Err.Raise vbObjectError + 513, , "excel_to_sharepoint.py returned exit code " & exitCode
    End If

    LogEvent "SharePoint sync completed successfully."
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, , "RunSpecAwardSync failed: " & Err.Description
End Sub

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = """" & value & """"
End Function

' ===========================================================
'   LOGGING
' ===========================================================

Public Sub LogEvent(ByVal msg As String)
    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")

    Debug.Print "[" & timestamp & "] " & msg

    Dim logPath As String
    logPath = "C:\Users\" & Environ("USERNAME") & "\Documents\OutlookMacroLog.txt"

    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    Open logPath For Append As #fileNum
    Print #fileNum, "[" & timestamp & "] " & msg
    Close #fileNum
    On Error GoTo 0
End Sub

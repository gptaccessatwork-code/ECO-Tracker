Option Explicit

Public WithEvents inboxItems As Outlook.Items

' ===========================================================
'   STARTUP & INITIALIZATION
' ===========================================================

Public Sub StartCRFMonitoring()
    Set inboxItems = Session.GetDefaultFolder(olFolderInbox).Items

    Set CRFSentTracker = CreateObject("Scripting.Dictionary")
    RebuildCRFTrackerFromSentItems

    Set SpecAwardSentTracker = CreateObject("Scripting.Dictionary")

    InitSpecAwardSenders

    LogEvent "CRF and SpecAward Monitoring Started."
End Sub

Private Sub Application_Startup()
    StartCRFMonitoring
End Sub

' ===========================================================
'   NEW MAIL WATCHER
' ===========================================================

Private Sub inboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

    Dim Mail As Outlook.MailItem
    Set Mail = Item

    Dim senderSMTP As String
    senderSMTP = GetSMTPFromSender(Mail)

    If senderSMTP = CRF_SENDER Then
        HandleCRFEmail Mail
    End If

    HandleSpecAwardEmail Mail

    Exit Sub

ErrorHandler:
    LogEvent "ERROR in inboxItems_ItemAdd: " & Err.Number & " - " & Err.Description
End Sub

' ===========================================================
'   CRF HANDLER
' ===========================================================

Private Sub HandleCRFEmail(ByVal Mail As Outlook.MailItem)
    On Error GoTo ErrorHandler

    Dim recip As Outlook.Recipient
    For Each recip In Mail.Recipients
        If RecipientMatchesChangeCoordinator(recip) Then
            LogEvent "CRF skipped - Change Coordinator already on thread: " & Mail.Subject
            Exit Sub
        End If
    Next recip

    Dim crfKey As String
    crfKey = ExtractCRFKey(Mail.Subject)
    If crfKey = "" Then Exit Sub

    If HasCRFBeenProcessed(crfKey) Then
        LogEvent "CRF already processed, skipping: " & crfKey
        Exit Sub
    End If

    If Not CRFSentTracker.Exists(crfKey) Then
        CRFSentTracker.Add crfKey, True
    End If

    ForwardCRFEmail Mail, crfKey
    MarkRelatedMailsRead crfKey

    On Error Resume Next
    ProcessCRFData crfKey, Mail.Subject, Mail.ReceivedTime
    If Err.Number <> 0 Then
        LogEvent "ERROR while processing CRF data: " & Err.Number & " - " & Err.Description & _
                 " | Subject: " & Mail.Subject
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    LogEvent "ERROR in HandleCRFEmail: " & Err.Number & " - " & Err.Description & _
             " | Subject: " & Mail.Subject
End Sub

Private Function RecipientMatchesChangeCoordinator(ByVal recip As Outlook.Recipient) As Boolean
    Dim smtpAddress As String
    smtpAddress = LCase$(Trim$(GetSMTPAddress(recip)))
    If smtpAddress <> "" Then
        If InStr(1, smtpAddress, LCase$(CHANGE_COORDINATOR_EMAIL), vbTextCompare) > 0 Then
            RecipientMatchesChangeCoordinator = True
            Exit Function
        End If
    End If

    Dim recipientName As String
    recipientName = LCase$(Trim$(recip.Name))
    If recipientName <> "" Then
        If InStr(1, recipientName, LCase$(CHANGE_COORDINATOR_NAME), vbTextCompare) > 0 Then
            RecipientMatchesChangeCoordinator = True
            Exit Function
        End If
    End If

    Dim recipientAddress As String
    recipientAddress = LCase$(Trim$(recip.Address))
    If recipientAddress <> "" Then
        If InStr(1, recipientAddress, LCase$(CHANGE_COORDINATOR_EMAIL), vbTextCompare) > 0 Then
            RecipientMatchesChangeCoordinator = True
            Exit Function
        End If
        If InStr(1, recipientAddress, LCase$(CHANGE_COORDINATOR_NAME), vbTextCompare) > 0 Then
            RecipientMatchesChangeCoordinator = True
            Exit Function
        End If
    End If
End Function

' ===========================================================
'   SPEC AWARD HANDLER
' ===========================================================

Private Sub HandleSpecAwardEmail(ByVal Mail As Outlook.MailItem)
    On Error GoTo ErrorHandler

    Dim senderSMTP As String
    senderSMTP = GetSMTPFromSender(Mail)
    If Not IsAuthorizedSpecAwardSender(senderSMTP) Then Exit Sub

    If InStr(LCase(Mail.Subject), "spec award") = 0 Then Exit Sub

    Dim entryKey As String
    entryKey = Mail.EntryID
    If SpecAwardSentTracker.Exists(entryKey) Then
        LogEvent "Spec Award already processed, skipping: " & Mail.Subject
        Exit Sub
    End If

    Dim latestBody As String
    latestBody = GetLatestBodyOnly(Mail.HTMLBody)

    Dim bodyToCheck As String
    bodyToCheck = LCase(latestBody)

    Dim hasKeywords As Boolean
    hasKeywords = (InStr(bodyToCheck, "part number") > 0) And _
                  (InStr(bodyToCheck, "qtr") > 0) And _
                  (InStr(bodyToCheck, "need by date") > 0 Or InStr(bodyToCheck, "nbd") > 0)

    If Not hasKeywords Then
        LogEvent "Spec Award subject matched but keywords not found in latest body: " & Mail.Subject
        Exit Sub
    End If

    LogEvent "SPEC AWARD detected: " & Mail.Subject

    Dim fwd As Outlook.MailItem
    Set fwd = Mail.Forward

    Dim intro As String
    intro = "Hi guys," & vbCrLf & vbCrLf & _
            "Please help to release below GPs to Agile. Thanks." & vbCrLf & vbCrLf & _
            "Best regards," & vbCrLf & "Sankar"

    Dim finalBody As String
    If Len(latestBody) < MIN_BODY_LENGTH Then
        finalBody = Mail.HTMLBody
    Else
        finalBody = latestBody
    End If

    fwd.HTMLBody = "<div style='font-family:Verdana; font-size:11pt;'>" & _
                   Replace(intro, vbCrLf, "<br>") & _
                   "<hr>" & finalBody & "</div>"

    fwd.To = FORWARD_TO
    fwd.CC = FORWARD_CC
    fwd.BCC = FORWARD_BCC
    fwd.Send

    SpecAwardSentTracker.Add entryKey, True
    LogEvent "Spec Award forwarded: " & Mail.Subject

    On Error Resume Next
    ProcessSpecAwardData latestBody, Mail.Subject, Mail.ReceivedTime
    If Err.Number <> 0 Then
        LogEvent "ERROR while processing Spec Award data: " & Err.Number & " - " & Err.Description & _
                 " | Subject: " & Mail.Subject
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    LogEvent "ERROR in HandleSpecAwardEmail: " & Err.Number & " - " & Err.Description & _
             " | Subject: " & Mail.Subject
End Sub

' ===========================================================
'   HELPER: SMTP ADDRESS RESOLVERS
' ===========================================================

Private Function GetSMTPAddress(r As Outlook.Recipient) As String
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    On Error Resume Next
    GetSMTPAddress = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    On Error GoTo 0
End Function

Private Function GetSMTPFromSender(Mail As Outlook.MailItem) As String
    On Error Resume Next
    Dim sender As Outlook.AddressEntry
    Set sender = Mail.Sender
    If Not sender Is Nothing Then
        If sender.AddressEntryUserType = olExchangeUserAddressEntry Or _
           sender.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
            Dim exchUser As Outlook.ExchangeUser
            Set exchUser = sender.GetExchangeUser
            If Not exchUser Is Nothing Then
                GetSMTPFromSender = LCase(exchUser.PrimarySmtpAddress)
                On Error GoTo 0
                Exit Function
            End If
        End If
        GetSMTPFromSender = LCase(sender.Address)
    Else
        GetSMTPFromSender = LCase(Mail.SenderEmailAddress)
    End If
    On Error GoTo 0
End Function

' ===========================================================
'   HELPER: CHECK AUTHORIZED SPEC AWARD SENDER
' ===========================================================

Private Function IsAuthorizedSpecAwardSender(senderSMTP As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(SPEC_AWARD_SENDERS)
        If senderSMTP = SPEC_AWARD_SENDERS(i) Then
            IsAuthorizedSpecAwardSender = True
            Exit Function
        End If
    Next i
    IsAuthorizedSpecAwardSender = False
End Function

' ===========================================================
'   HELPER: STRIP QUOTED REPLY HISTORY FROM HTML BODY
' ===========================================================

Private Function GetLatestBodyOnly(html As String) As String
    Dim p As Long
    Dim lowerHTML As String
    lowerHTML = LCase(html)

    p = InStr(lowerHTML, "<div id=""divrplyfwdmsg""")
    If p = 0 Then p = InStr(lowerHTML, "<div id=divrplyfwdmsg")

    If p > 0 Then
        GetLatestBodyOnly = Left(html, p - 1)
        Exit Function
    End If

    p = InStr(lowerHTML, "<hr")
    If p > 10000 Then
        GetLatestBodyOnly = Left(html, p - 1)
        Exit Function
    End If

    GetLatestBodyOnly = html
End Function

' ===========================================================
'   HELPER: EXTRACT CRF KEY FROM SUBJECT
' ===========================================================

Private Function ExtractCRFKey(subject As String) As String
    If CRFRegex Is Nothing Then
        Set CRFRegex = CreateObject("VBScript.RegExp")
        CRFRegex.Pattern = "CRF\s*[:#-]?\s*(\d{5,})"
        CRFRegex.IgnoreCase = True
    End If
    If CRFRegex.Test(subject) Then
        ExtractCRFKey = CRFRegex.Execute(subject)(0).SubMatches(0)
    Else
        ExtractCRFKey = ""
    End If
End Function

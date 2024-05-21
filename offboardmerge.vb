Sub MailMergeWithCC()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim oDoc As Document
    Dim oData As MailMergeDataSource
    Dim i As Integer
    Dim UserName As String
    Dim UserEmail As String
    Dim OffboardingDate As String
    Dim OffboardingTime As String
    Dim CCAddresses As String
    Dim SelectedRecords As String
    Dim RecordArray() As String
    Dim RecordNumber As Integer
    Dim OffboardType As String
    Dim EmailBody As String

    ' Prompt for the user's name and email, offboarding date and time
    UserName = InputBox("Enter the user's name:", "User Information")
    UserEmail = InputBox("Enter the user's email:", "User Information")
    OffboardingDate = InputBox("Enter the offboarding date (e.g., May 17, 2024):", "Offboarding Information")
    OffboardingTime = InputBox("Enter the offboarding time (e.g., 5:00 PM):", "Offboarding Information")
    
    ' Prompt for the type of offboarding
    OffboardType = InputBox("Is this offboarding voluntary or involuntary? (Enter V for voluntary or I for involuntary):", "Offboarding Type")
    
    ' Check the type of offboarding and set the email body accordingly
    If UCase(Trim(OffboardType)) = "V" Then
        EmailBody = "Hello," & vbCrLf & vbCrLf & _
                    "The following user has been offboarded as of " & OffboardingDate & " @ " & OffboardingTime & ". Please block any access they may have by end of day for:" & vbCrLf & vbCrLf & _
                    "<<software>>" & vbCrLf & vbCrLf & _
                    UserName & vbCrLf & _
                    UserEmail & vbCrLf & vbCrLf & _
                    "Thank you,"
    ElseIf UCase(Trim(OffboardType)) = "I" Then
        EmailBody = "Hello," & vbCrLf & vbCrLf & _
                    "The following user has been offboarded as of " & OffboardingDate & " @ " & OffboardingTime & ". Please block any access they may have ASAP for:" & vbCrLf & vbCrLf & _
                    "<<software>>" & vbCrLf & vbCrLf & _
                    UserName & vbCrLf & _
                    UserEmail & vbCrLf & vbCrLf & _
                    "Thank you,"
    Else
        MsgBox "Invalid offboarding type entered. Please enter 'V' for voluntary or 'I' for involuntary.", vbExclamation
        Exit Sub
    End If

    ' Prompt for the selected records (comma-separated list)
    SelectedRecords = InputBox("Enter the record numbers to include (comma-separated, e.g., 1,3,5):", "Select Records")
    RecordArray = Split(SelectedRecords, ",")

    ' Create Outlook application object
    Set OutApp = CreateObject("Outlook.Application")
    Set oDoc = ActiveDocument
    Set oData = oDoc.MailMerge.DataSource

    ' Ensure the mail merge is active
    If oDoc.MailMerge.State = wdMainAndDataSource Then
        ' Iterate through each selected record
        For i = LBound(RecordArray) To UBound(RecordArray)
            RecordNumber = CInt(Trim(RecordArray(i)))
            oData.ActiveRecord = RecordNumber

            ' Get CC addresses, if any
            CCAddresses = oData.DataFields("cc_email").Value

            ' Create a new email
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = oData.DataFields("main_email").Value
                .Subject = "Offboarding User"
                .Body = Replace(EmailBody, "<<software>>", oData.DataFields("software").Value)

                ' Add CC addresses if they exist
                If Len(Trim(CCAddresses)) > 0 Then
                    .CC = CCAddresses
                End If

                ' Mark as important if involuntary
                If UCase(Trim(OffboardType)) = "I" Then
                    .Importance = 2 ' 2 = high importance
                End If

                .Send
            End With
        Next i
    Else
        MsgBox "The mail merge is not properly set up. Please check your data source.", vbExclamation
    End If

    ' Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

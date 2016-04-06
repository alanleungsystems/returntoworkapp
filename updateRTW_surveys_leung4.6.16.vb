
'The following section was implemented to include survey letters from the Workforce Well-Being Services group'
Private Sub Survey_Click()
On Error GoTo Err_Survey_Click

'Variables'
    Dim stDocName As String
    Dim strSurv As String
    Dim strPath As String
    Dim strSQL As String
    Dim strSurveysLoc As String
    Dim strLtrType As String
    
    strPath = Application.CurrentProject.Path
'Make sure this points to the C:\Program Files\RTW2003oracle\Letters\Surveys when moving to production'
    strSurveysLoc = "I:\RTW2003oracle\Letters\Surveys\"

'This macro creates the deletes an old table and creates a new table of people who left the RTW system in the last 7 days
    stDocName = "mcrSurveys"
    DoCmd.RunMacro stDocName
    DoCmd.SetWarnings False

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    strLtrType = Me.cboLetters.Value
    Me.lblLtrType.Caption = strLtrType

'Deletes existing files before writing new ones'
    If fso.FileExists(strSurveysLoc & "mailmergeSurveys" & Me.lblLtrType.Caption) Then
        fso.DeleteFile strSurveysLoc & "mailmergeSurveys" & Me.lblLtrType.Caption
    End If
    Set fso = Nothing

'Exports tblSurveys from Access to the specified drive'
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tblSurveys", strSurveysLoc & "mailmergeSurveys" & Me.lblLtrType.Caption, True
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim myMerge As Word.MailMerge
    Dim strSurveyFileName As String

    strSurveyFileName = strSurveysLoc & "SatisfactionSurvey"

'Opens up MSWord and starts the mail merge process'
    Set oWord = CreateObject("Word.Application")
        With oWord
            .Visible = True
            .Activate
            .WindowState = wdWindowStateNormal
        End With

    Set oDoc = oWord.Documents.Open(strSurveyFileName)
    If oDoc.MailMerge.STATE = wdMainAndDataSource Then
        oDoc.MailMerge.Execute
    End If

'Sends the document to the default printer and prints the letters'
    Set myMerge = oDoc.MailMerge
    With myMerge
        If .Destination <> 1 Then
        .Destination = wdSendToPrinter
        .Execute
        End If
    End With

'Closes the MSWord doc'
    oDoc.Close SaveChanges:=wdDoNotSaveChanges

    oWord.Quit
    Set oWord = Nothing

Exit_Survey_Click:
    Exit Sub

Err_Survey_Click:
    MsgBox Err.Description
    Resume Exit_Survey_Click
    
End Sub

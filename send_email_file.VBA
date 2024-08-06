' VBA  send files with email
' Column A: To
' Column B: CC -> separated by semicolons ";"
' Column C: Subject
' Column D: body, HTML script
' Column E: name of files separated by semicolon ";"
' Column F: VBA write status if email, successful or unsuccessful
' Column G: VBA write name of file with error

Sub sendemail()
    Dim I, j As Integer
    Dim page1 As Worksheet
    Set page1 = ActiveWorkbook.Worksheets("Hoja1")
    Dim OutApp As Object
    Dim Email As Object
    Dim FilePath As String
    Dim File As Variant
    Dim Pass As Integer
    Dim FileName As String
    
    ' App
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    ' For each row with data
    For To_ = 1 To WorksheetFunction.CountA(Columns(1))
    
        ' variables
        Debug.Print Cells(To_, 1).Value & " | " & Cells(To_, 2).Value & " | " & Cells(To_, 3).Value & " | " & Cells(To_, 4).Value
        
        ' open Outlook if this close
        On Error Resume Next
        Set OutApp = GetObject("", "Outlook.Application")
        Err.Clear
        
        ' App
        If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")

        ' folder path to files to attach in the email (files PDFs)
        FilePath = "C:\" 
        
        ' Item email
        OutApp.Visible = True
        Set Email = OutApp.CreateItem(0)
        
        ' create object mail
        With Email

            ' If Pass = then send eamil
            Pass = 1

            ' add subjects
            .To = Cells(To_, 1).Value
            .CC = Cells(To_, 2).Value
            .Subject = Cells(To_, 3).Value
            .HTMLBody = Cells(To_, 4).Value

            ' add files
            For Each File In Split(Cells(To_, 5).Value, ";")
                If File <> "" Then ' Last value return char space
                    ' Drop char space
                    FileName = Trim(File)
                    If Dir(FilePath & FileName & vbNormal) <> "" Then 
                        ' Add file
                        .Attachments.Add (FilePath & FileName)
                    Else
                        Debug.Print "Not exists <" & FileName & ">"
                        ' Write name of the file with error
                        Cells(To_, 7).Value = Cells(To_, 7).Value & " - not exists <" & FileName & ">"
                        Cells(To_, 6).Value = "Not successful !!!"
                        ' email is not sent
                        Pass = 0
                    End If
            Next File
            ' Check if is Attachments all files
            If Pass = 1 Then
                ' view previsualization of email 
                .Display
                ' Send email
                ' .Send
                Cells(To_, 6).Value = "was successful !!!"
                Cells(To_, 7).Value = ""
            End If
        End With
    
        ' App
        With Application
            .EnableEvents = True
            .ScreenUpdating = True
        End With

    ' next row with data
    Next To_

End Sub
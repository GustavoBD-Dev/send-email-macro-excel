' VBA  send files with email
' Column A: To
' Column B: CC -> separated by semicolons ';' 
' Column C: Subject
' Column D: body, HTML script
' Column E: name of files separated by space ' '

Sub enviarcorreo()
    Dim I, j As Integer
    Dim pagina1 As Worksheet
    Set pagina1 = ActiveWorkbook.Worksheets("Hoja1")
    Dim OutApp As Object
    Dim Correo As Object
    Dim FilePath As String
    Dim File As Variant
    
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

        ' folder path to files
        FilePath = "C:\Users\A8065955\OneDrive - GFBANORTE\Documents\Correcciones\PDFs\"
        
        ' Item email
        OutApp.Visible = True
        Set Correo = OutApp.CreateItem(0)
        
        ' create object mail
        With Correo

            ' add files
            For Each File In Split(Cells(To_, 5).Value)
                .Attachments.Add (FilePath & File)
            Next File

            ' add subjects
            .To = Cells(To_, 1).Value
            .CC = Cells(To_, 2).Value
            .Subject = Cells(To_, 3).Value
            .HTMLBody = Cells(To_, 4).Value
            '.Send 'send email
            .Display ' show windows send whitout send message
        End With
    
        ' App
        With Application
            .EnableEvents = True
            .ScreenUpdating = True
        End With

    ' next row with data
    Next To_

End Sub
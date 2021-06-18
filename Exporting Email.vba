Option Explicit

Sub ExportEmail()

    Dim objfile As FileSystemObject
    Dim xNewFolder
    Dim xDir As String, xMonth As String, xFile As String, xPath As String
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim NameX As Name, xStp As Long
    Dim xDate As Date, AWBookPath As String
    Dim currentWB As Workbook, newWB As Workbook
    Dim strEmailTo As String, strEmailCC As String, strEmailBCC As String, strDistroList As String
    
    AWBookPath = ActiveWorkbook.Path & "\"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Creating Email and Attachment for " & Format(Date, "dddd dd mmmm yyyy")
    
    Set currentWB = ActiveWorkbook
    
    xDate = Date
    
    '******************************Grabbing New WorkBook and Formatting*************
    
    Sheets(Array("Cover", "Graph Page", "My Data")).Copy
    
    Set newWB = ActiveWorkbook
    
    Range("A1").Select
    Sheets("My Data").Visible = False
    Sheets("Cover").Select
    
   
    '******************************Creating Pathways*********************************
    
    xDir = AWBookPath
    xMonth = Format(xDate, "mm mmmm yy") & "\"
    
    xFile = "My Report" & Format(xDate, "dd-mm-yyyy") & ".xlsx"
    
    xPath = xDir & xMonth & xFile
    
    '******************************Saving File in Pathway*********************************
    
    Set objfile = New FileSystemObject
    
    If objfile.FolderExists(xDir & xMonth) Then
        If objfile.FileExists(xPath) Then
            objfile.DeleteFile (xPath)
            newWB.SaveAs Filename:=xPath, FileFormat:=xlOpenXMLWorkbook, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
            , CreateBackup:=False
            
            Application.ActiveWorkbook.Close
        Else
            newWB.SaveAs Filename:=xPath, FileFormat:=xlOpenXMLWorkbook, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
            , CreateBackup:=False
            Application.ActiveWorkbook.Close
        End If
    Else
        xNewFolder = xDir & xMonth
        MkDir xNewFolder
        newWB.SaveAs Filename:=xPath, FileFormat:=xlOpenXMLWorkbook, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
            , CreateBackup:=False
        Application.ActiveWorkbook.Close
    End If
    
    '******************************Preparing Distribution List *********************************

    currentWB.Activate
    Sheets("Email").Visible = True
    Sheets("Email").Select
    
    strEmailTo = ""
    strEmailCC = ""
    strEmailBCC = ""
    
    xStp = 1
    
    Do Until xStp = 4
    
        Cells(2, xStp).Select
        
        Do Until ActiveCell = ""
        
            strDistroList = ActiveCell.Value
        
            If xStp = 1 Then strEmailTo = strEmailTo & strDistroList & "; "
            If xStp = 2 Then strEmailCC = strEmailCC & strDistroList & "; "
            If xStp = 3 Then strEmailBCC = strEmailBCC & strDistroList & "; "
            
            ActiveCell.Offset(1, 0).Select
            
        Loop
        
        xStp = xStp + 1
    
    Loop
    
    Range("A1").Select
    
    '******************************Preparing Email*********************************
    
    Set olApp = New Outlook.Application
       Dim olNs As Outlook.Namespace
       Set olNs = olApp.GetNamespace("MAPI")
       olNs.Logon
    Set olMail = olApp.CreateItem(olMailItem)
    olMail.To = strEmailTo
    olMail.CC = strEmailCC
    olMail.BCC = strEmailBCC
    
        
        olMail.Subject = Mid(xFile, 1, Len(xFile) - 4)
        olMail.Body = vbCrLf & "Hello Everyone," _
                            & vbCrLf & vbCrLf & "Please find attached the " & Mid(xFile, 1, Len(xFile) - 4) & "." _
                            & vbCrLf & vbCrLf & "Regards," _
                            & vbCrLf & "Ramyaa Prasath - 2019503547"
    
    
    olMail.Attachments.Add xPath
    olMail.Display
   
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub SaveClose()
    ActiveWorkbook.Close True
End Sub

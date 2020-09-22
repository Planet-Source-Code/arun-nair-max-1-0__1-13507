VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Address eXtractor(MAX) ver 1.2"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ForeColor       =   &H00800000&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDomainCheck 
      Height          =   345
      Left            =   1320
      TabIndex        =   13
      Top             =   1620
      Width           =   1785
   End
   Begin VB.CommandButton bnSaveToFile 
      Height          =   405
      Left            =   5310
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Export results into a text file"
      Top             =   900
      Width           =   405
   End
   Begin VB.CommandButton bnExportToCSV 
      Height          =   405
      Left            =   4800
      Picture         =   "Form1.frx":07C8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Export results into a Microsoft Outlook compatible CSV file"
      Top             =   900
      Width           =   405
   End
   Begin VB.CommandButton bnBrowseFile 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3510
      TabIndex        =   9
      Top             =   480
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   390
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton bnCopyToClipboard 
      Height          =   405
      Left            =   4290
      Picture         =   "Form1.frx":0B7C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copy results to clipboard"
      Top             =   900
      Width           =   405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Mail Separator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   3210
      TabIndex        =   6
      Top             =   1470
      Width           =   2625
      Begin VB.OptionButton mailSepOpLBr 
         BackColor       =   &H00C00000&
         Caption         =   "Line break"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1380
         TabIndex        =   3
         Top             =   210
         Width           =   1215
      End
      Begin VB.OptionButton mailSepOpComma 
         BackColor       =   &H00C00000&
         Caption         =   "Comma"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox txtDisplayBox 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2235
      Width           =   5655
   End
   Begin VB.CommandButton bnScan 
      Cancel          =   -1  'True
      Caption         =   "Click here to start the scan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   210
      TabIndex        =   1
      Tag             =   "stop"
      Top             =   900
      Width           =   2475
   End
   Begin VB.TextBox txtFileScan 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   3195
   End
   Begin MSComctlLib.ProgressBar prMailCtr 
      Height          =   225
      Left            =   180
      Negotiate       =   -1  'True
      TabIndex        =   10
      Top             =   4980
      Visible         =   0   'False
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Having name (domain)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Left            =   150
      TabIndex        =   14
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   4410
      Picture         =   "Form1.frx":0F57
      ToolTipText     =   "This utility was developed by Arun Nair www.techinnova.com"
      Top             =   4950
      Width           =   1440
   End
   Begin VB.Label status 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   210
      TabIndex        =   8
      Top             =   4980
      Width           =   4035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the file to be scanned. Separate multiple files to be scanned by a comma."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   7
      Top             =   30
      Width           =   5595
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' mailStorage - variable that has all the mails addresses parsed
' temporarily stored
Dim mailStorage
Dim fileList
Dim filesListArr()
' Array where all the mail addresses are stored
Dim mailArr(2000)
Private Function checkDomain(getMail) As Boolean
  ' function to check if the domain name exists in the
  ' mail list
    
    If InStr(1, getMail, txtDomainCheck) > 0 Then
        'return true if the domain matches
        checkDomain = True
    End If

End Function

Private Sub bnScan_Click()
    ' The tag of the button decides whether to start
    ' or to stop the scan
    Call toggleBn
    If bnScan.Tag = "start" Then
        Erase mailArr
        Me.addFileNamesToArray (Me.txtFileScan)

        mailStorage = ""
        txtDisplayBox.Text = ""
        
        If Me.mailSepOpComma.Value Then
            Me.scanMailAdd (",")
        Else
            Me.scanMailAdd (vbNewLine)
        End If
    End If
    
End Sub

Public Function scanMailAdd(sepVal)
' This function is the entry point
' for the scan.
' This one is heavily nested and co-ordinates with the
' other functions
' sepVal - is the separator that separates every mail

On Error GoTo chkErr ' error handler

Dim mailCnt, fileCtr, lnCtr
lnCtr = mailCnt = 0

Dim lineRead, spcCtr, mailStart, mailEnd
    
    txtDisplayBox = "" ' Initialize the display box

If Trim(filesListArr(UBound(filesListArr, 1))) = "" Then Exit Function

For fileCtr = 1 To UBound(filesListArr) 'parse all files 1 by 1
   
   'update the progress bar with the maximum no of lines
    progInit (fileCtr)
    
    'open the file and assign filenumber 1 to it
    Open filesListArr(fileCtr) For Input As #1
        'set the mail count to 0
        mailCnt = 0
        
        'parse the file till not EOF or the tag on the button is scan
        Do While Not EOF(1)
'            MsgBox "Entering the outer loop"
            'increment the line counter by 1
            lnCtr = lnCtr + 1
            'increment the progress bar
            prMailCtr.Value = lnCtr
            status.Visible = True
            status.Caption = lnCtr
            'read the line and store in the lineread var
            Line Input #1, lineRead
            
            mailStart = 0
            mailEnd = 0
               
               'continue parsing the line till the mailend value
               'is smaller than lineread
               ' the logic here is simple , after reading a certain chunk,
               ' truncate the chunk and assign the remaining characters of
               ' lineRead as the new value for lineRead.
               ' This avoid having to read all through the line once again and
               ' can save substantial time especially if the line is pretty long.
               ' This method also avoids redundancy.
               
'               MsgBox "Mail end >> " & mailEnd & vbNewLine & "Len lineread " & Len(lineRead)
               
               Do While mailEnd <= Len(lineRead)
                If bnScan.Tag = "stop" Then GoTo showMails
               
'                MsgBox "Entering the inner loop"
 '               MsgBox InStr(1, lineRead, "@")
                'if an @ is found in the line, start the process
                If Len(lineRead) > 1 And InStr(1, lineRead, "@") > 0 Then
                    
                    spcCtr = InStr(1, lineRead, "@")
                
                    'first do a reverse check and capture the start pos of the
                    'email id
                    
                    Do While Trim(Mid(lineRead, spcCtr, 1)) <> ""
                         'MsgBox Asc(Trim(Mid(lineRead, spcCtr, 1)))
                         Select Case Asc(Trim(Mid(lineRead, spcCtr, 1)))
                            Case 95, 45, 46, 64, 97 To 122, 65 To 90, 48 To 57
                             'MsgBox Trim(Mid(lineRead, spcCtr, 1)) & " valid"
                             spcCtr = spcCtr - 1
                            Case Else
                                
                            Exit Do
                         End Select
                    
                        If spcCtr = 0 Then Exit Do
                        DoEvents
                    Loop
                
                
                    spcCtr = spcCtr + 1
                    mailStart = spcCtr
              
                ' now parse the characters after the @ for illegal chars
                ' if an unacceptable/illegal char is found, you have found
                ' the end pos of the email id!
                
                    Do While spcCtr > 0 And Trim(Mid(lineRead, spcCtr, 1)) <> ""
                         Select Case Asc(Trim(Mid(lineRead, spcCtr, 1)))
                            Case 95, 45, 46, 64, 97 To 122, 65 To 90, 48 To 57
                    
                                spcCtr = spcCtr + 1
                            Case Else
                                Exit Do
                        End Select
                        DoEvents
                    Loop
                
                    
                    mailEnd = spcCtr
                
                    Dim dispMail
                
                    'the chars between mailstart to mailend is where your
                    ' email id is
                    
                    'now pass the extracted chars to extractMailVal for
                    ' verification and final modification
                    dispMail = extractMailVal(Mid(lineRead, mailStart, mailEnd - mailStart))
                
                     If dispMail <> "" Then
                    
                        ' here I'm limiting the amt of mails to 1000
                        ' you can increment this value if you want
                        
                        If mailCnt >= 1000 Then
                            MsgBox "Sorry!, But this version supports only upto 1000 mails. For more information mail at arunn@techinnova.com", vbExclamation, "Limit reached"
                            'directly jump to the showMails label
                            GoTo showMails
                        End If
                    
                        mailCnt = mailCnt + 1
                        
                        'if the mailStorage var is empty
                        'then the sepval is not needed
                        'or you'll have something like >> ,xyz@abc.com !
                        
                        If Trim(mailStorage) = "" Then
                            mailStorage = dispMail
                        Else
                            mailStorage = mailStorage + sepVal + dispMail
                        End If
                        
                        mailArr(mailCnt) = dispMail
                        
                     
                     End If
                        
                    'The caption of the form should reflect the scan status
                    Caption = "Scanning... " + Str(mailCnt) + " mail(s) found"
                
                Else
                  Exit Do
                End If
                
'                    MsgBox "LInread >> " & lineRead & " << Mail end " & mailEnd & ">> Len(lineread) " & Len(lineRead)
                    lineRead = Mid(lineRead, mailEnd, Len(lineRead) - (mailEnd - 1))
                
            DoEvents
            
            Loop
            'Close #1
        Loop
        
        'the scan is over, call the toggleBn
showMails:
        If bnScan.Tag = "start" Then
            toggleBn
        End If
        
        bnScan.Visible = True
        txtDisplayBox.Text = mailStorage
    Close #1
Next


    Caption = "Mail Address eXtractor(MAX) 1.0"
    
    prMailCtr.Visible = False
    
    status = Str(UBound(filesListArr)) + " file(s) scanned. " + Str(mailCnt) + " mail(s) found"

Exit Function

'Error handler

chkErr:
    If Err.Number = 55 Then
        Exit Function
    End If
     
     
    Close #1
    
    bnScan.Visible = True
    'call the toggleBn
    MsgBox "togbn called by error"
    toggleBn
    
    Caption = "Mail Address eXtractor(MAX) 1.0"
    prMailCtr.Visible = False
    status = "Done "
    
    'Display the error
    MsgBox Err.Description & Err.Number, vbCritical, "Error encountered"
    
End Function


Private Sub bnCopyToClipboard_Click()

    'copy the results to the clipboard
    If Trim(txtDisplayBox) = "" Then
        MsgBox "There are no results to be copied to the clipboard.", vbInformation, "No results"
    Else
        Clipboard.Clear
        Clipboard.SetText (txtDisplayBox)
    End If

End Sub

Private Sub bnBrowseFile_Click()
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    
    CommonDialog1.Flags = cdlOFNHideReadOnly

    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|HTML Files (*.htm,*.html)|*.htm;*.html|Email Files (*.eml)|*.eml"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file


    

    
    txtFileScan.Text = Me.CommonDialog1.FileName
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    MsgBox Err.Description
    

End Sub
Function extractMailVal(getPreVal) As String
        
        'check the first and last character of the extracted text
        ' if they are alphanumeric then proceed with further evaluation
        
        Select Case Asc(Mid(getPreVal, 1, 1)) And Asc(Mid(getPreVal, Len(getPreVal), 1))
            Case 97 To 122, 65 To 90, 48 To 57
                '97-122(a-z), 65-90(A-Z)
                '48-59(0-9)
                ' resume if these chars are
                ' involved

            Case Else
               'invalid mail
               extractMailVal = ""
               Exit Function
        End Select


                Dim atCtr, pos
                atCtr = 0
                
                For pos = 1 To Len(getPreVal)
                    
                    If Mid(getPreVal, pos, 1) = "@" Then
                        'increment the atCtr(ctr which
                        'counts the @ in the text
                        atCtr = atCtr + 1
                    End If
                    
                    'not more than 1 @ is allowed in the
                    'mail address!!
                    If atCtr > 1 Then
                        extractMailVal = ""
                        Exit Function
                    End If
                Next
                'the email is structurely right if it has reached so far
                
                ' next check if the "." lies within the last 4 chars
                ' if not - then eliminate the chars after the "."
                ' At first I thought of using the truncation algorithm
                ' but it would be impractical and could also misjudge
                ' a real email ID
                               
                Dim revPos, valParse
                
                For valParse = 1 To Len(getPreVal)
                    
                    If Mid(getPreVal, Len(getPreVal) - (valParse - 1), 1) = "." Then
                        If valParse > 4 Then
                            getPreVal = Mid(getPreVal, 1, Len(getPreVal) - valParse)
                        End If
                        Exit For
                    End If
                
                Next
             
                'check for duplicate entries
                If Me.chkDupl(getPreVal) = True Then
                    extractMailVal = ""
                    Exit Function
                End If
                'check if the mail lies in the domain
                'MsgBox getPreVal
                
                If Trim(txtDomainCheck) <> "" Then
                    If checkDomain(getPreVal) = False Then
                        extractMailVal = ""
                        Exit Function
                    End If
                End If
                    
                
    
    
    'all is well...well almost :)
    'now return back the modified input as proper mail ID
    extractMailVal = getPreVal

End Function
Function chkDupl(getMail) As Boolean
        'this function checks the list for duplicate entries
        If InStr(1, mailStorage, getMail) > 0 Then
                    chkDupl = True
                    Exit Function
        End If
    
    
    
End Function

Sub addFileNamesToArray(fileList)
'****** This Procedure will accept the filelist
' from the textbox and seggregates(if any) into
' the filesListArr array



' Erase the contents of the array
' this is to prevent accumulation
Erase filesListArr

Dim startPos, initPos, arrCtr
    'initialize the counters
    arrCtr = 0: startPos = 1: initPos = 0

    If InStr(1, fileList, ",") > 0 Then
            'if the filelist contains commas(,)
            ' seggregate them into arrays
            
            While InStr(initPos + 1, fileList, ",") > 0
        
                arrCtr = arrCtr + 1
                ReDim Preserve filesListArr(arrCtr)
                
                startPos = initPos
                initPos = InStr(initPos + 1, fileList, ",")
                        
                'the array filesListArr where the files are
                'seggregated
                filesListArr(arrCtr) = Mid(fileList, startPos + 1, (initPos - startPos) - 1)
'                MsgBox filesListArr(arrCtr)

            Wend
            
            'the last file in the list to be appended
            ' to the array
            
            arrCtr = arrCtr + 1
            ReDim Preserve filesListArr(arrCtr)
    
            filesListArr(arrCtr) = Mid(fileList, initPos + 1, (Len(fileList) - initPos))
    
            
    Else
        'if the number of files in the box is only one
        arrCtr = arrCtr + 1
        ReDim filesListArr(arrCtr)
        
        filesListArr(arrCtr) = Trim(fileList)

    
    End If

End Sub

Sub progInit(fCtr)
    'set the visible property of the progress bar to True
    prMailCtr.Visible = True
        
    Dim lnCtr, dummyVar
    lnCtr = 0
    
    Open filesListArr(fCtr) For Input As #2
        'count the number of lines in the file
        While Not EOF(2)
            Line Input #2, dummyVar
            lnCtr = lnCtr + 1
            
        Wend
        
        prMailCtr.MAX = lnCtr

        
        
    Close #2

End Sub

Private Sub bnExportToCSV_Click()
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    
    CommonDialog1.Flags = cdlOFNHideReadOnly

    ' Set filters
    CommonDialog1.Filter = "Outlook CSV Format(*.csv)|*.csv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowSave
    ' Display name of selected file

    
    Open Me.CommonDialog1.FileName For Output As #3
    Print #3, "E-mail Address"
    Dim cnt
    cnt = 1
    
    While mailArr(cnt) <> "" And cnt <= UBound(mailArr)
        Print #3, mailArr(cnt)
        cnt = cnt + 1
    Wend
    Close #3

ErrHandler:
    'User pressed the Cancel button

End Sub


Private Sub bnSaveToFile_Click()
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    
    CommonDialog1.Flags = cdlOFNHideReadOnly

    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowSave
    ' Display name of selected file

    
    Open Me.CommonDialog1.FileName For Output As #4
    Print #4, "Extracted by MAX(Mail Address Extractor) 1.2 "
    Print #4, "---------------------------------------------"
    Dim cnt
    cnt = 1
    
    While mailArr(cnt) <> "" And cnt <= UBound(mailArr)
        Print #4, mailArr(cnt)
        cnt = cnt + 1
    Wend
    Close #4

ErrHandler:
    'User pressed the Cancel button

End Sub


Public Sub toggleBn()
    ' this function will toggle the button tags and start the scan
'    MsgBox "Toggle called"
'    MsgBox bnScan.Tag
    If bnScan.Tag = "stop" Then
'        MsgBox "stop"
        bnScan.Caption = "Click here to stop the scan"
        bnScan.Tag = "start"
        
    ElseIf bnScan.Tag = "start" Then
'        MsgBox "start"
        bnScan.Caption = "Click here to start the scan"
        bnScan.Tag = "stop"
    End If

End Sub


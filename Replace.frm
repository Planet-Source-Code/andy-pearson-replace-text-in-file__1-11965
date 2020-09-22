VERSION 5.00
Begin VB.Form frmReplace 
   Caption         =   "Replace Text in File"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReplace 
      Height          =   495
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "Replace.frx":0000
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdEntireFile 
      Caption         =   "Entire File (multi-line)"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton cmdLineByLine 
      Caption         =   "Line By Line (better for large files)"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtFind 
      Height          =   495
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Replace.frx":000F
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "c:\data_out.txt"
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "c:\data.txt"
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblLineNum 
      Caption         =   "Current Line Number: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Replace Text"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Find Text"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Output File"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Input File"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEntireFile_Click()
    Dim fso As Object
    Dim strFile As String
    Const ForReading = 1
    
    If do_before_replace_stuff Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.OpenTextFile(txtInput, ForReading)
        strFile = Replace(f.ReadAll & "", txtFind, txtReplace)
        f.Close
        Print #2, strFile;
    End If
    do_after_replace_stuff
End Sub

Private Sub cmdLineByLine_Click()
    Dim strLine As String
    Dim lngLine As Long
    
    If Not do_before_replace_stuff Then
        If InStr(txtFind, vbCrLf) > 0 Or InStr(txtReplace, vbCrLf) > 0 Then
            MsgBox "Can't have carriage returns in find/replace text."
            Exit Sub
        End If
        
        Open txtInput For Input As #1
        lngLine = 0
        While Not EOF(1)
            lngLine = lngLine + 1
            Line Input #1, strLine
            strLine = Replace(strLine, txtFind.Text, txtReplace.Text)
            Print #2, strLine
            lblLineNum.Caption = "Current Line Number: " & lngLine
            DoEvents
        Wend
        Close #1
    End If
    do_after_replace_stuff
End Sub

Private Function do_before_replace_stuff() As Boolean
    do_before_replace_stuff = True
    If Dir(txtInput) = "" Then
        MsgBox "Can't find input file."
        do_before_replace_stuff = False
        Exit Function
    End If
    If Dir(txtOutput) <> "" Then
        If MsgBox("Delete old output file?", vbYesNo) = vbYes Then
            Kill txtOutput
        Else
            do_before_replace_stuff = False
            Exit Function
        End If
    End If
    MousePointer = vbHourglass 'set cursor to hourglass
    txtInput.Enabled = False
    txtOutput.Enabled = False
    txtFind.Enabled = False
    txtReplace.Enabled = False
    cmdLineByLine.Enabled = False
    cmdEntireFile.Enabled = False
    Open txtOutput For Output As #2
End Function

Private Sub do_after_replace_stuff()
    Close #2
    txtInput.Enabled = True
    txtOutput.Enabled = True
    txtFind.Enabled = True
    txtReplace.Enabled = True
    cmdLineByLine.Enabled = True
    cmdEntireFile.Enabled = True
    MousePointer = vbDefault
    MsgBox "done"
End Sub


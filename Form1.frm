VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FS Network Troubleshooting Tools"
   ClientHeight    =   6375
   ClientLeft      =   510
   ClientTop       =   195
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   5520
   End
   Begin VB.Frame Frame3 
      Caption         =   "Command Output"
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   9255
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3135
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5530
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"Form1.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8160
      TabIndex        =   11
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Traceroute Tool"
      Height          =   1935
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command6 
         Caption         =   "More Info"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Execute"
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Remote IP Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   $"Form1.frx":0082
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ping Tool"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command5 
         Caption         =   "More Info"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ping Self"
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Execute"
         Height          =   495
         Left            =   3240
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Remote IP Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"Form1.frx":0134
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Idle"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   5880
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I've been trying to get DOS command line output to show up in
'my Visual Basic programs.  I finally figured it out.  It works
'really simple.  The program creates a batch file that is to be
'executed, instead of using the "Shell" command to execute single
'commands.  The shell command also writes the output to a file.
'At the end of each batch file, the progams writes
'the word "DONE" to a text file called "finished."  When a user
'starts a command, a timer constantly checks for the word "DONE"
'in the "finished" file.  When the word finally is written to
'the file, the batch file is finished executing, and the program
'can read the output file it created.

Private Sub CreatePingBatchFile(command As String)
    'This function first clears the word "DONE" from the "finished"
    'file.  It then creates the new batch file to be executed.
    
    Dim intfile As Integer
    Dim intfile2 As Integer
    
    intfile = FreeFile
    intfile2 = FreeFile

    'This clears the finished file.
    Open "c:\finished.txt" For Output As #intfile
        Print #intfile, ""
    Close #intfile

    'This writes the batch file.
    Open "c:\NewBatch.bat" For Output As #intfile2
        Print #intfile2, "command.com /c ping " & command & " > " & "c:\OutputText.txt"
        Print #intfile2, vbNewLine & "echo DONE > " & "c:\finished.txt"
    Close #intfile2
End Sub

Private Sub CreateTracertBatchFile(command As String)
    'This function first clears the word "DONE" from the "finished"
    'file.  It then creates the new batch file to be executed.
    
    Dim intfile As Integer
    Dim intfile2 As Integer

    intfile = FreeFile
    intfile2 = FreeFile

    'This clears the finished file.
    Open "c:\finished.txt" For Output As #intfile
        Print #intfile, ""
    Close #intfile
    
    'This writes the batch file.
    Open "c:\NewBatch.bat" For Output As #intfile2
        Print #intfile2, "command.com /c tracert " & command & " > " & "c:\OutputText.txt"
        Print #intfile2, vbNewLine & "echo DONE > " & "c:\finished.txt"
    Close #intfile2
End Sub

Private Sub Command1_Click()
    'When this button is clicked, it calls the CreatePingBatchFile
    'using the text from text1 as the IP.  It then executes the
    'new batch file and starts the timer to begin checking the
    '"finished" file for the word "DONE."
    Call CreatePingBatchFile(Text1.Text)
    
    Dim strbatchfile As String
    
    strbatchfile = "c:\NewBatch.bat"
    Shell (strbatchfile), vbMinimizedNoFocus
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    'When this button is clicked, it first changes text1.text to the
    'loop back address for its own IP.  It then calls the CreatePingBatchFile
    'using the text from text1 as the IP.  It then executes the
    'new batch file and starts the timer to begin checking the
    '"finished" file for the word "DONE."
    Text1.Text = "127.0.0.1"
    
    Call CreatePingBatchFile(Text1.Text)
    
    Dim strbatchfile As String
    
    strbatchfile = "c:\NewBatch.bat"
    Shell (strbatchfile), vbMinimizedNoFocus
    Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
    'Does this need explained?
    End
End Sub

Private Sub Command4_Click()
    'When this button is clicked, it calls the CreateTracertBatchFile
    'using the text from text1 as the IP.  It then executes the
    'new batch file and starts the timer to begin checking the
    '"finished" file for the word "DONE."
    Call CreateTracertBatchFile(Text2.Text)
    
    Dim strbatchfile As String
    
    strbatchfile = "c:\NewBatch.bat"
    Shell (strbatchfile), vbMinimizedNoFocus
    Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
    'Shows form3 for the ping information.
    Form3.Show
End Sub

Private Sub Command6_Click()
    'Shows form3 for the tracert information.
    Form2.Show
End Sub

Private Sub Timer1_Timer()
    'This timer, when enabled, checks the file called "finished" for
    'the word "DONE."  When it finds it, it will open the output file
    'from the batch file and display it in the rich text box.

    Dim intfile3 As Integer
    Dim strline As String

    intfile3 = FreeFile
    Label5.Caption = "Processing..."
    
    Open "c:\finished.txt" For Input As #intfile3
        Line Input #intfile3, strline
    Close #intfile3
    
    If Left(strline, 4) = "DONE" Then
        RichTextBox1.LoadFile ("c:\outputtext.txt")
        Timer1.Enabled = False
        Label5.Caption = "Done"
    End If
End Sub

'Created by Frank Scarberry

VERSION 5.00
Begin VB.Form frmQuestionaire 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Questionaire"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAns 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   900
      Width           =   495
   End
   Begin VB.Label lblQuiz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label lblQuiz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmQuestionaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program displays a question and series of options
'stored in a data file and informs the user whether or
'not the option picked is correct.
'I made use of an array of labels to display the options
'and question, and an array of command buttons to
'identify the options.

'***************************************
'Written by 'seyi Ogunbanjo in MS VB6.0
'(theguru@programmer.net)
'Date Created: 04-09-02, western
'Last updated: 06-07-02, western
'***************************************

Dim qtn As String, opt(1 To 4) As String
Dim ans As Integer 'id of correct answer.
Dim rights As Integer 'num of correctly answered qtns
Dim all As Integer    'total qtns attempted
Option Explicit

Private Sub cmdAns_Click(Index As Integer)
 Dim confirm As Integer, prompt As String
 If Index = ans Then
    MsgBox "Absolutely correct."
    rights = rights + 1
 Else
    MsgBox "Nice try, but u got it wrong."
 End If
 all = all + 1
 If EOF(1) Then
    prompt = "No more questions. Would you like to view your performance?"
    confirm = MsgBox(prompt, vbYesNo, "Out of questions.")
    If confirm = vbYes Then Call ShowPerformance(rights, all)
    'Close
    Unload Me
    End
    'The program ends when there are no more questions.
 Else
    Call Questionaire
 End If
End Sub

Private Sub cmdQuit_Click()
 Dim confirm  As Integer, prompt As String
 prompt = "Would you like to view your performance?"
 confirm = MsgBox(prompt, vbYesNo, "Out of questions.")
 If confirm = vbYes Then Call ShowPerformance(rights, all)
 'Close
 Unload Me
 End
End Sub

Private Sub Form_Load()
 On Error GoTo ErrorHandler:
 Dim file As String
 Call LoadForm
 file = "qtns.txt"
 Open file For Input As #1
 If EOF(1) Then
    MsgBox "Data File is empty. Quiting program.", , "Bye"
    Unload Me: End
 Else
    Call Questionaire
 End If
 Exit Sub
ErrorHandler:
 If Err.Number = 53 Then
    MsgBox "File not found."
 End If
 Resume Next
 
End Sub

Private Sub LoadForm()
 'Load and arrange controls on form
 Dim x As Integer
 For x = 2 To 4
    Load cmdAns(x)
    Load lblQuiz(x)
    cmdAns(x).Top = cmdAns(x - 1).Top + _
    cmdAns(x - 1).Height * 1.5
    cmdAns(x).Caption = x
    cmdAns(x).TabIndex = x
    cmdAns(x).Visible = True
    lblQuiz(x).Top = cmdAns(x).Top + 50
    lblQuiz(x).Left = lblQuiz(1).Left
    lblQuiz(x).Visible = True
 Next x
End Sub

Private Sub Questionaire()
'Read the next question, set of options,
'and correct answer id from the data file.

'You can add to the data file at any time,
'just follow the order.
 Dim x As Integer
 Input #1, ans, qtn
 lblQuiz(0).Caption = qtn
 For x = 1 To 4
    Input #1, opt(x)
    lblQuiz(x).Caption = opt(x)
 Next x
   
End Sub

Private Sub ShowPerformance(r As Integer, a As Integer)
 Dim prompt As String
 prompt = "You got " + Str(r) + " of " + Str(a) & " questions right."
 MsgBox prompt, , "Great Job!"
End Sub

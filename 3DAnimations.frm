VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "3D Animations !"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   1320
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Erase "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6840
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Typing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Scroll "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   -120
      ScaleHeight     =   1515
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "This label is used for animations in Picturebox !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   4890
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===========================================================
'================3D ANIMATIONS MADE EASY====================
'===========================================================
'This is to show how to animate 3D text in a simple and nice way
'This is completely my idea co'z I have never seen anything like this
'in PSC so I made one
'I hope you like this code and it helps you somewhat
'This is cool to put in your project
'And if you think this deseves a vote then vote for it !
'Thanks and any comments or suggestions then email at
'mark_anthony_dinglasa@yahoo.com hope to hear from you !

Dim Word As String, Sel As Integer 'variables used for Animations

Private Sub Command1_Click() 'Scroll Animation
Label1.Caption = "               Mark Anthony Dinglasa "
Timer3.Enabled = False
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Command2_Click() 'Typing Animation
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Command3_Click() 'Erase Animation
Label1.Caption = "Mark Anthony Dinglasa"
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Form_Activate() 'Just to display the text in Picturebox
For i = 0 To 255        '250 will be the width of the text, you can costumized it
    Picture1.ForeColor = RGB(i, i + 1, i + 90) 'Color of the text
    Picture1.CurrentX = i       'posisition of x-axis
    Picture1.CurrentY = i          'posisiton of y-axis
    Picture1.Print " Mark Anthony Dinglasa "  'The text that will be printed in picturebox
Next
End Sub

Private Sub Form_Load()
Word = " Mark Anthony Dinglasa " 'Used in Typing Animation
End Sub

Private Sub Label1_change()
    Picture1.Cls        'Refreshes Picturebox every label event is change
For i = 0 To 255        '250 will be the width of the text, you can costumized it
    Picture1.ForeColor = RGB(i, i + 1, i + 90) 'Color of the text
    Picture1.CurrentX = i       'posisition of x-axis
    Picture1.CurrentY = i          'posisiton of y-axis
    Picture1.Print Label1.Caption  'The text that will be printed in picturebox
Next
End Sub

Private Sub Timer1_Timer() 'This is for Scroll Animation
Dim Pass As String
    Pass = Label1.Caption
    Pass = Mid(Pass, 2, Len(Pass)) + Left(Pass, 1) 'scrolling from right to left
    Label1.Caption = Pass
End Sub

Private Sub Timer2_Timer() 'This is for Typing Animation
Label1.Caption = Left(Word, Sel)
    Sel = Sel + 1
        If Sel > Len(Word) Then
            Sel = 0
        End If
End Sub

Private Sub Timer3_Timer() 'This for Erase Animations
On Error GoTo err
Label1.Caption = Mid(Label1.Caption, 1, Len(Label1.Caption) - 1) 'Erasing text from right to left
err:
 If err.Number = 5 Then Call Form_Activate: Exit Sub 'Handles error when error occurs
End Sub

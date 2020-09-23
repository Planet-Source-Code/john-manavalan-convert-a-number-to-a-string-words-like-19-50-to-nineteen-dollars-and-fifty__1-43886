VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3720
   ClientTop       =   3180
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.Timer animation 
      Interval        =   200
      Left            =   1560
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About Me"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      MaxLength       =   9
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Number in the Text Box "
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s

Private Sub Text1_Change()
'eg :- Msgbox numWor("120")
Text2 = numWor((Text1.Text))
End Sub
Private Sub Command1_Click()
MsgBox "Written by" & vbCrLf & "    John Manavalan" & vbCrLf & "    66, Priyadarshini Nagar, Paravattani, " & vbCrLf & "    Trichur, Kerala, India-680 005 " & vbCrLf & "Email:" & vbCrLf & "    johnmanavalan@hotmail.com", vbOKOnly, "Account Manager 3.07 v"
End Sub

Private Sub animation_Timer()
txt = Space(60) & "Written By John Manavalan Email:-johnmanavalan@hotmail.com" & Space(5)
s = s + 1
Me.Caption = Mid(txt, s, 60)
If s = Len(txt) Then s = 0
End Sub

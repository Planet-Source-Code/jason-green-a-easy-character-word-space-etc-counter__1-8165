VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "3"
   ClientHeight    =   4350
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   3375
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label12 
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Number of spaces:"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Number of other characters:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Number of characters:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Number of words:"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Number of vowels:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Number of consonants:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
On Error Resume Next
Label2 = 0
Label4 = 0
Label6 = 0
Label8 = 0
Label10 = 0
Label12 = 0
If Text1 <> "" Then
If Mid(Text1.Text, 1, 1) <> " " Then
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 1) = " " Then
Label2.Caption = Val(Label2.Caption) + 1
End If
Next
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 2) = "  " Then
Label2.Caption = Label2.Caption - 1
End If
Next
Label2 = Label2 + 1
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 1) = "a" Or Mid(Text1.Text, i, 1) = "e" Or Mid(Text1.Text, i, 1) = "i" Or Mid(Text1.Text, i, 1) = "o" Or Mid(Text1.Text, i, 1) = "u" Then
Label4.Caption = Val(Label4.Caption) + 1
ElseIf Mid(Text1.Text, i, 1) = "b" Or Mid(Text1.Text, i, 1) = "c" Or Mid(Text1.Text, i, 1) = "d" Or Mid(Text1.Text, i, 1) = "f" Or Mid(Text1.Text, i, 1) = "g" Or Mid(Text1.Text, i, 1) = "h" Or Mid(Text1.Text, i, 1) = "j" Or Mid(Text1.Text, i, 1) = "k" Or Mid(Text1.Text, i, 1) = "l" Or Mid(Text1.Text, i, 1) = "m" Or Mid(Text1.Text, i, 1) = "n" Or Mid(Text1.Text, i, 1) = "p" Or Mid(Text1.Text, i, 1) = "q" Or Mid(Text1.Text, i, 1) = "r" Or Mid(Text1.Text, i, 1) = "s" Or Mid(Text1.Text, i, 1) = "t" Or Mid(Text1.Text, i, 1) = "v" Or Mid(Text1.Text, i, 1) = "w" Or Mid(Text1.Text, i, 1) = "x" Or Mid(Text1.Text, i, 1) = "y" Or Mid(Text1.Text, i, 1) = "z" Then
Label6.Caption = Val(Label6.Caption) + 1
ElseIf Mid(Text1.Text, i, 1) = " " Then
Label12 = Label12 + 1
Else
Label10 = Label10 + 1
End If
Next
If Mid(Text1.Text, Len(Text1.Text), 1) = " " Then
Label2 = Label2 - 1
End If
End If
End If
If Text1.Text = "" Then
Label2 = 0
End If
Label8 = Len(Text1.Text)
End Sub



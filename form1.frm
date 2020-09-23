VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Random Number Generator"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "form1.frx":0442
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generate my number"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Number:"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Enter Maximum number here:"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox texty 
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Command1_Click()
    If texty.Text = "" Then
    MsgBox ("You have not filled out the maximum number you desire!")
    On Error GoTo goat
goat:
    End
    End If
    
    Randomize
    Value = texty.Text * Rnd
    Label1.Caption = Str$(Value)

End Sub



VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Backgroundpicture"
   ClientHeight    =   4368
   ClientLeft      =   1320
   ClientTop       =   1392
   ClientWidth     =   6384
   LinkTopic       =   "Form2"
   ScaleHeight     =   4368
   ScaleWidth      =   6384
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   372
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1212
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4596
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   4596
      ScaleWidth      =   4488
      TabIndex        =   0
      Top             =   0
      Width           =   4488
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form2.Hide
End Sub


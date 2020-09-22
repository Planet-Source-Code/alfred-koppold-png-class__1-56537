VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4368
   ClientLeft      =   1320
   ClientTop       =   1392
   ClientWidth     =   6384
   LinkTopic       =   "Form1"
   ScaleHeight     =   4368
   ScaleWidth      =   6384
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   1800
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "*.png|*.png"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "open png-picture"
      Height          =   732
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pngClass As New LoadPNG
Private Sub Command1_Click()
Dim filename As String
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
If filename <> "" Then
pngClass.PicBox = Form1 'or Picturebox
pngClass.SetToBkgrnd True, 100, 50 'set to Background (True or false), x and y
pngClass.BackgroundPicture = Form1 'same Backgroundpicture
pngClass.SetAlpha = True 'when Alpha then alpha
pngClass.SetTrans = True 'when transparent Color then transparent Color
pngClass.OpenPNG filename 'Open and display Picture
End If
End Sub


VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Open Picture Files"
   ClientHeight    =   5364
   ClientLeft      =   1440
   ClientTop       =   1968
   ClientWidth     =   7728
   LinkTopic       =   "Form1"
   ScaleHeight     =   5364
   ScaleWidth      =   7728
   WindowState     =   2  'Maximiert
   Begin VB.Frame Frame1 
      Caption         =   "Background (only for png)"
      Height          =   2292
      Left            =   4560
      TabIndex        =   8
      Top             =   0
      Width           =   4332
      Begin VB.CommandButton Command3 
         Caption         =   "Reset Picture"
         Height          =   252
         Left            =   2280
         TabIndex        =   24
         Top             =   1920
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.TextBox txtY 
         Height          =   288
         Left            =   1680
         TabIndex        =   23
         Text            =   "0"
         Top             =   1920
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.TextBox txtX 
         Height          =   288
         Left            =   480
         TabIndex        =   20
         Text            =   "0"
         Top             =   1920
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   132
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   804
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change Picture"
         Height          =   252
         Left            =   2280
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Set Background Picture"
         Height          =   252
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2172
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Set Own BackgroundColor"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2172
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Set Png BackgroundColor"
         Height          =   250
         Left            =   120
         TabIndex        =   12
         Top             =   1000
         Width           =   2172
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Set Alpha for png"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   320
         Width           =   1572
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Set Transparence for png"
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2292
      End
      Begin VB.PictureBox Picture3 
         Height          =   132
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   804
         TabIndex        =   9
         Top             =   1000
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label10 
         Caption         =   "y"
         Height          =   252
         Left            =   1320
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label Label9 
         Caption         =   "x"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.Label Label8 
         Caption         =   "has no transparence"
         ForeColor       =   &H000000FF&
         Height          =   252
         Left            =   2400
         TabIndex        =   19
         Top             =   680
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label7 
         Caption         =   "has no alpha"
         ForeColor       =   &H000000FF&
         Height          =   252
         Left            =   1800
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.Label Label2 
         Caption         =   "has no bkgnd chunk"
         ForeColor       =   &H000000FF&
         Height          =   250
         Left            =   2520
         TabIndex        =   17
         Top             =   1000
         Visible         =   0   'False
         Width           =   1572
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000004&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4872
      Left            =   0
      ScaleHeight     =   4872
      ScaleWidth      =   4212
      TabIndex        =   1
      Top             =   0
      Width           =   4212
      Begin VB.PictureBox pic1 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   456
         Left            =   0
         ScaleHeight     =   38
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   2
         Top             =   0
         Width           =   576
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   372
      Left            =   4440
      TabIndex        =   0
      Top             =   2400
      Width           =   1572
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   960
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   $"Form1.frx":0000
   End
   Begin VB.Label Label6 
      Height          =   252
      Left            =   1320
      TabIndex        =   7
      Top             =   5160
      Width           =   1932
   End
   Begin VB.Label Label5 
      Caption         =   "gama:"
      Height          =   252
      Left            =   0
      TabIndex        =   6
      Top             =   5160
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "Last Modification:"
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Top             =   4920
      Width           =   1452
   End
   Begin VB.Label Label3 
      Height          =   252
      Left            =   1440
      TabIndex        =   4
      Top             =   4920
      Width           =   1692
   End
   Begin VB.Label Label1 
      Height          =   3612
      Left            =   4560
      TabIndex        =   3
      Top             =   2880
      Width           =   4092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cPicC As clsPicContainer
Attribute cPicC.VB_VarHelpID = -1
Private PicturePath As String
Dim PicType As String

Private Sub Check1_Click()
If Check1 = 1 Then
If Option1 = False And Option2 = False And Option3 = False Then
Option1 = True
End If
Else
If Check2 = 0 Then
Option1 = False
Option2 = False
Option3 = False
End If
End If
TestVisible
End Sub

Private Sub Check2_Click()
If Check2 = 1 Then
If Option1 = False And Option2 = False And Option3 = False Then
Option1 = True
End If
Else
If Check1 = 0 Then
Option1 = False
Option2 = False
Option3 = False
End If
End If
TestVisible
End Sub

Private Sub Command1_Click()
Dim Ende As Long
Dim Beendet As Boolean
Dim Anfang As Long
Dim png As New LoadPNG
Dim tif As New clsTiff
Dim Teststring As String
Dim tga As New LoadTGA
Dim pcx As New LoadPCX
Dim Test As Long
Dim Testtxt As String
Label1 = ""
Label3 = ""
Label6 = ""
Picture3.Visible = False
Label2.Visible = False
CommonDialog1.Filter = "*.tif *.png *.bmp *.pcx *.tga *.dib *.gif *.jpg *.wmf *.emf *.ico *.cur|*.tif; *.png; *.bmp; *.pcx; *.tga; *.dib; *.gif; *.jpg; *.wmf; *.emf; *.ico; *.cur"
CommonDialog1.Filename = ""
CommonDialog1.ShowOpen
If UCase(Right(CommonDialog1.Filename, 3)) = "PCX" Then
pic1.Picture = LoadPicture("")
End If
Select Case UCase(Right(CommonDialog1.Filename, 3))
Case "BMP"
PicType = "BMP"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "DIB"
PicType = "DIB"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "GIF"
PicType = "GIF"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "JPG"
PicType = "JPG"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "WMF"
PicType = "WMF"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "EMF"
PicType = "EMF"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "ICO"
PicType = "ICO"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "CUR"
PicType = "CUR"
pic1 = LoadPicture(CommonDialog1.Filename)
Case "PCX"
PicType = "PCX"
pcx.LoadPCX CommonDialog1.Filename
pcx.DrawPCX pic1
Case "TIF"
PicType = "TIF"
tif.LoadTIFF CommonDialog1.Filename
Case "TGA"
PicType = "TGA"
'tga.DrawAlpha = True
'tga.SetToBkgrnd True, 10, 10
tga.LoadTGA CommonDialog1.Filename
tga.DrawTGA pic1
Case "PNG"
PicType = "PNG"
If Option3 = True Then
png.PicBox = pic1
png.BackgroundPicture = pic1
png.SetToBkgrnd True, txtX.Text, txtY.Text
picContainer.AutoRedraw = True
Else
png.PicBox = pic1
picContainer.Picture = LoadPicture("")
End If
If Option2 Then
png.SetOwnBkgndColor True, Picture1.BackColor
Else
png.SetOwnBkgndColor False
End If
If Check1.Value = 1 Then
png.SetAlpha = True
End If
If Check2.Value = 1 Then
png.SetTrans = True
End If
Test = png.OpenPNG(CommonDialog1.Filename)
Label3.Caption = png.ModiTime
If png.gama <> 0 Then Label6.Caption = png.gama
If png.Text <> "" Then
Testtxt = png.Text
Anfang = 1
Do While Beendet = False
Ende = InStr(Anfang, Testtxt, Chr(0))
If Ende = 0 Then Exit Do
Teststring = Teststring & Mid(Testtxt, Anfang, Ende - Anfang) & ": "
Anfang = Ende + 1
Ende = InStr(Anfang, Testtxt, Chr(0))
If Ende = 0 Then Exit Do
Teststring = Teststring & Mid(Testtxt, Anfang, Ende - Anfang) & vbCrLf
Anfang = Ende + 1
Loop
End If
If png.zText <> "" Then
Testtxt = png.zText
Anfang = 1
Do While Beendet = False
Ende = InStr(Anfang, Testtxt, Chr(0))
If Ende = 0 Then Exit Do
Teststring = Teststring & Mid(Testtxt, Anfang, Ende - Anfang) & ": "
Anfang = Ende + 1
Ende = InStr(Anfang, Testtxt, Chr(0))
If Ende = 0 Then Exit Do
Teststring = Teststring & Mid(Testtxt, Anfang, Ende - Anfang) & vbCrLf
Anfang = Ende + 1
Loop
End If
If Teststring <> "" Then Label1.Caption = Teststring

If png.ErrorNumber <> 0 Then MsgBox "Error Nr. " & png.ErrorNumber
If png.HasBKGDChunk Then
Picture3.BackColor = png.BkgdColor
Picture3.Visible = True
Else
Label2.Visible = True
End If
If png.HaveAlpha Then
Label7.Visible = False
Else
Label7.Visible = True
End If
If png.HaveTransparence Then
Label8.Visible = False
Else
Label8.Visible = True
End If
End Select
RepaintPB

If PicType = "PNG" Then
If Option1 = True Then
If png.HasBKGDChunk Then
picContainer.BackColor = png.BkgdColor
Else
picContainer.BackColor = RGB(150, 150, 150)
End If
End If
If Option2 = True Then picContainer.BackColor = Picture1.BackColor
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
CommonDialog1.Filename = ""
CommonDialog1.Filter = "*.bmp|*.bmp; *.jpg|*.jpg"
CommonDialog1.ShowOpen
If CommonDialog1.Filename <> "" Then
pic1.Picture = LoadPicture(CommonDialog1.Filename)
PicturePath = CommonDialog1.Filename
End If
RepaintPB
End Sub

Private Sub Command3_Click()
If PicturePath = "" Then
pic1.Picture = Form2.Picture2.Picture
Else
pic1.Picture = LoadPicture(PicturePath)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set cPicC = Nothing
End

End Sub

Private Sub Option1_Click()
If Option1 = True Then
If Check1 = 0 And Check2 = 0 Then
Check1 = 1
Check2 = 1
Else
End If
End If
TestVisible
End Sub

Private Sub Option2_Click()
If Option2 = True Then
If Check1 = 0 And Check2 = 0 Then
Check1 = 1
Check2 = 1
End If
End If
TestVisible
End Sub

Private Sub Option3_Click()
If Option3 = True Then
If Check1 = 0 And Check2 = 0 Then
Check1 = 1
Check2 = 1
End If
pic1.Picture = Form2.Picture2.Picture
RepaintPB
End If
TestVisible
End Sub
Private Sub TestVisible()
If Option2 = True Then
Picture1.Visible = True
Else
Picture1.Visible = False
End If
If Option3 = True Then
Command2.Visible = True
Command3.Visible = True
txtX.Visible = True
txtY.Visible = True
Label9.Visible = True
Label10.Visible = True
Else
PicturePath = ""
Command2.Visible = False
Command3.Visible = False
txtX.Visible = False
txtY.Visible = False
Label9.Visible = False
Label10.Visible = False
End If
End Sub

Private Sub Picture1_Click()
CommonDialog1.ShowColor
Picture1.BackColor = CommonDialog1.Color
End Sub

Public Sub RepaintPB()
      If cPicC Is Nothing Then
         Set cPicC = New clsPicContainer
         cPicC.Init picContainer, pic1
         Else
         Set cPicC = Nothing
         Set cPicC = New clsPicContainer
        cPicC.Init picContainer, pic1
      End If
End Sub

Private Sub txtX_Change()
TestNumber txtX
End Sub

Private Sub txtY_Change()
TestNumber txtY
End Sub
Private Sub TestNumber(TextO As Object)
If IsNumeric(TextO.Text) = False Then
TextO.Text = "0"
End If
End Sub

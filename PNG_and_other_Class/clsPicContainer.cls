VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPicContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-------------------------------------------------------
'
'        clsPicContainer   (c) peter.k.sauer@web.de
'                              05.2003
'        eine Picturebox als Container für ein
'        verschiebbares Object wie Picture, Frame,
'        das als Behälter für andere Objecte dienen kann
'-------------------------------------------------------
'        Ereignisse: Change
'-------------------------------------------------------

Option Explicit

Private WithEvents picContainer As PictureBox   'das Fenster
Attribute picContainer.VB_VarHelpID = -1
Private PicMove As Object                       'Objecteträger
Private Parentform As Form                      'Mutterform

Private WithEvents HScrol As HScrollBar         'Scrollbars
Attribute HScrol.VB_VarHelpID = -1
Private WithEvents VScrol As VScrollBar
Attribute VScrol.VB_VarHelpID = -1
Private PicEdge As PictureBox                   'Ecke

Public Event Change()

Private InitDone As Boolean                     'für Event Change

Private mvarStretchPicture As Boolean           'Symbol für ziehen

Public Property Let StretchPicture(ByVal vData As Boolean)
      mvarStretchPicture = vData
      'Bild aktualisieren
      Refresh
End Property

Public Property Get StretchPicture() As Boolean
      StretchPicture = mvarStretchPicture
End Property


Public Sub Init(Parentbox As PictureBox, Containerbox As Object)
'initialisieren der notwendigen Elemente

      If InitDone Then
         Exit Sub
      End If

      'zuweisen der beiden Boxen
      Set picContainer = Parentbox
      Set PicMove = Containerbox
      Set Parentform = Parentbox.Parent
      
      Set HScrol = LoadControl("VB.HScrollBar", "HScrol")
      Set VScrol = LoadControl("VB.VScrollBar", "VScrol")
      Set PicEdge = LoadControl("VB.PictureBox", "PicEdge")
      
      'Controls auf PicContainer setzen
      Set PicMove.Container = picContainer
      Set HScrol.Container = picContainer
      Set VScrol.Container = picContainer
      Set PicEdge.Container = picContainer
      
      picContainer.BackColor = vbApplicationWorkspace
      PicMove.Left = 0
      PicMove.Top = 0
      'Scrollbars einrichten
      With HScrol
         .Height = 240
         .Min = 0
         .Max = 100
         .Value = 0
         .LargeChange = .Max / 5
         .SmallChange = .Max / 3
      End With
      With VScrol
         .Width = 240
         .Min = 0
         .Max = 100
         .Value = 0
         .LargeChange = .Max / 5
         .SmallChange = .Max / 3
      End With
      'Ecke
      With PicEdge
         .BorderStyle = 0
         .Height = 240
         .Width = 240
         .AutoRedraw = True
         .FontName = "Marlett"
         .FontSize = 12
      End With
      
      'visualisieren
      Refresh
      
      InitDone = True
End Sub

Public Sub Refresh()

      SetScrollBars
      SetScrollValues
      
      If PicMove.Width <= (picContainer.ScaleWidth - VScrol.Width) Then
         PicMove.Left = 0
         HScrol.Enabled = False
      Else
         HScrol.Enabled = True
      End If
      If PicMove.Height <= (picContainer.ScaleHeight - HScrol.Height) Then
         PicMove.Top = 0
         VScrol.Enabled = False
      Else
         VScrol.Enabled = True
      End If
      
      If InitDone Then
         'Event auslösen
         RaiseEvent Change
      End If
End Sub

Private Sub SetScrollValues()

   Dim x As Single
   
      'Scrollbalken (Sprung)Breite
      With HScrol
      If picContainer.ScaleWidth < PicMove.Width Then
         x = picContainer.ScaleWidth / PicMove.Width
         x = Int(x * .Max)
         .SmallChange = x * 0.5
         .LargeChange = x
         Else
         .SmallChange = 10
         .LargeChange = 100
         End If
      End With
      
      With VScrol
      If picContainer.ScaleHeight < PicMove.Height Then
         x = picContainer.ScaleHeight / PicMove.Height
         x = Int(x * .Max)
         .SmallChange = x * 0.5
         .LargeChange = x
         Else
         .SmallChange = 10
         .LargeChange = 100
         End If
      End With
End Sub

Private Sub SetScrollBars()

      With HScrol
         .Left = 0
         .Top = picContainer.ScaleHeight - .Height
         .Width = picContainer.ScaleWidth - VScrol.Width
      End With
      With VScrol
         .Left = picContainer.ScaleWidth - .Width
         .Top = 0
         .Height = picContainer.ScaleHeight - HScrol.Height
      End With
      
      With PicEdge
         .Top = HScrol.Top
         .Left = HScrol.Width
         .BackColor = vbApplicationWorkspace
         .Cls
      End With
      
      If mvarStretchPicture Then
         PicEdge.BackColor = vbButtonFace
         PicEdge.CurrentX = 0
         PicEdge.CurrentY = 0
         PicEdge.Print Chr(111)
      End If
      
      'zum Test
      HScrol.Visible = True
      VScrol.Visible = True
      PicEdge.Visible = True
End Sub

Private Function LoadControl(LibraryName As String, _
                             objCtlName As String) As Control
'ein Control auf die Parentform beamen

   Dim i As Long
   Dim s As String
   Dim objCtl As Control
   
      i = 1
      Do
         On Error Resume Next
         'durch Anhängen einer Laufnummer eindeutige Bezeichner bilden
         s = objCtlName & i
         'Control in Parentform Collection aufnehmen
         Set objCtl = Parentform.Controls.Add(LibraryName, s, Parentform)
         If Err.Number = 0 Then
            'kein Fehler
            Exit Do
         ElseIf Err.Number = 727 Then
            'Control mit Bezeichner schon vorhanden, Laufnummer + 1
            i = i + 1
            Err.Clear
         Else
            'sonstiger Fehler
            FehlerAnzeige Err.Number, Err.Description, "clsFrameAsForm.LoadControl"
            Exit Do
         End If
      Loop
      
      'Control zuweisen
      Set LoadControl = objCtl
      Set objCtl = Nothing
End Function

Private Sub FehlerAnzeige(ErrNumber As Long, ErrDescription As String, _
                         Optional Titel As String = "")

   Dim Msg As String
   
      Msg = "Fehler " & ErrNumber & vbCrLf & vbCrLf & _
            ErrDescription
      MsgBox Msg, vbCritical, Titel
End Sub

Private Sub Class_Initialize()


      mvarStretchPicture = False       'kein Ziehen Symbol
End Sub

Private Sub Class_Terminate()
      
      'aufräumen
      Parentform.Controls.Remove HScrol
      Parentform.Controls.Remove VScrol
      Parentform.Controls.Remove PicEdge
      
      Set HScrol = Nothing
      Set VScrol = Nothing
      Set PicEdge = Nothing
      Set PicMove = Nothing
      Set picContainer = Nothing
      Set Parentform = Nothing
End Sub

Private Sub HScrol_Change()
   
   Dim x As Single
   
      'Scrollen Horizontal
      If HScrol.Value = 0 Then
         PicMove.Left = 0
      ElseIf HScrol.Value = HScrol.Max Then
         PicMove.Left = picContainer.ScaleWidth - PicMove.Width - VScrol.Width
      Else
         x = HScrol.Value / HScrol.Max
         PicMove.Left = (picContainer.ScaleWidth - PicMove.Width - VScrol.Width) * x
      End If
      
      Refresh
End Sub

Private Sub HScrol_GotFocus()

      picContainer.SetFocus
End Sub

Private Sub VScrol_Change()

   Dim x As Single
   
      'Scrollen Vertikal
      If VScrol.Value = 0 Then
         PicMove.Top = 0
      ElseIf VScrol.Value = VScrol.Max Then
         PicMove.Top = picContainer.ScaleHeight - PicMove.Height - HScrol.Height
      Else
         x = VScrol.Value / VScrol.Max
         PicMove.Top = (picContainer.ScaleHeight - PicMove.Height - HScrol.Height) * x
      End If
      
      Refresh
End Sub

Private Sub VScrol_GotFocus()

      picContainer.SetFocus
End Sub

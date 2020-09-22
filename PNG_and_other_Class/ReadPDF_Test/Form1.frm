VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4368
   ClientLeft      =   1320
   ClientTop       =   1392
   ClientWidth     =   6384
   LinkTopic       =   "Form1"
   ScaleHeight     =   4368
   ScaleWidth      =   6384
   Begin VB.CommandButton Command1 
      Caption         =   "Read Sample-PDF"
      Height          =   612
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   2052
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "*.pdf|*.pdf"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type ZI
Offset As Long
Inhalt As String
Other As Long
Typ As String
End Type
Dim Zeileninhalt() As ZI


Private Sub Command1_Click()
Dim Filename As String
Dim Filenumber As Long
Dim Filearray() As Byte
Dim Übersicht As Long
Dim suchbegriff As String
Dim Wo As Long
Dim Anfangzahl As Long
Dim Endezahl As Long
Dim Anzahl As Long
Dim Hilfslong As Long
Dim Längezahl As Long
Dim z As Long
Dim Beginstream As Long
Dim Zahlenarray() As Byte
Dim Zahlenstring As String
Dim Zahllong As Long
Dim Textstream() As Byte
Dim testarray() As Byte
Dim teststring As String
Dim Anfang As Long
Dim testbool As Long
Dim Ende As Long
Dim Länge As Long
Dim Zeilen() As String
Dim Zeilenbyte() As Byte
Dim AnzahlObjekte As Long
Dim Seitenzahl As Long
Dim BeginnStream As Long
Dim Endebefehl As Long
Dim i As Long
Dim Test As Long
Dim Stand As Long
Dim Befehlsarray() As String
Dim GrößeArray As Long
Dim EndeGefunden As Boolean
Dim Typ As String
Dim Befehleinzeln() As String
CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
Seitenzahl = 1
Filenumber = FreeFile
Open Filename For Binary As Filenumber
ReDim Filearray(LOF(Filenumber) - 1)
Get Filenumber, 1, Filearray
Close Filenumber
suchbegriff = StrConv("startxref", vbFromUnicode)
Wo = InStrB(Filearray, suchbegriff)
If Filearray(Wo - 2) = 10 Or Filearray(Wo - 2) = 13 Then '13 oder 10
Anfangzahl = Wo + 9
Endezahl = FindeEnde(Filearray, Wo + 9)
Längezahl = Endezahl - Anfangzahl
ReDim Zahlenarray(Längezahl - 1)
CopyMemory Zahlenarray(0), Filearray(Anfangzahl), Längezahl
Zahlenstring = Zahlenarray
Zahlenstring = StrConv(Zahlenstring, vbUnicode)
If IsNumeric(Zahlenstring) Then Zahllong = CLng(Zahlenstring)
ReDim testarray(3)
CopyMemory testarray(0), Filearray(Zahllong), 4
teststring = StrConv(testarray, vbUnicode) 'xref
Anfang = FindeEnde(Filearray, Zahllong) + 1
If Filearray(Anfang) = 10 Or Filearray(Anfang) = 13 Then
Anfang = Anfang + 1
End If
Ende = FindeEnde(Filearray, Anfang)
Länge = Ende - Anfang
ReDim testarray(Länge - 1)
CopyMemory testarray(0), Filearray(Anfang), Länge
teststring = StrConv(testarray, vbUnicode) '1. Zeile
Anfang = Ende + 1
If Filearray(Ende + 1) = 10 Or Filearray(Ende + 1) = 13 Then Anfang = Anfang + 1
Ende = InStr(teststring, " ")
Zahlenstring = Left(teststring, Ende - 1) 'Nummer xref
Zahlenstring = Mid(teststring, Ende + 1) 'Anzahl Objekte
If IsNumeric(Zahlenstring) Then AnzahlObjekte = CLng(Zahlenstring)
ReDim Zeilen(AnzahlObjekte - 1)
ReDim Zeileninhalt(AnzahlObjekte - 1)
ReDim Zeilenbyte(19)
For i = 0 To AnzahlObjekte - 1
CopyMemory Zeilenbyte(0), Filearray(Anfang), 20
Zeilen(i) = StrConv(Zeilenbyte, vbUnicode)
If Zeilenbyte(18) = 13 Then
Zeilen(i) = Left(Zeilen(i), 18)
Anfang = Anfang + 19
End If
If Zeilenbyte(18) = 32 Then
Zeilen(i) = Left(Zeilen(i), 18)
Anfang = Anfang + 20
End If
teststring = Mid(Zeilen(i), 18, 1)
If teststring = "n" Then
Zahlenstring = Mid(Zeilen(i), 1, 10)
If IsNumeric(Zahlenstring) Then Zeileninhalt(i).Offset = CLng(Zahlenstring)
Zeileninhalt(i).Inhalt = "n"
Zahlenstring = Mid(Zeilen(i), 12, 5)
If IsNumeric(Zahlenstring) Then Zeileninhalt(i).Other = CLng(Zahlenstring)
Else
Zahlenstring = Mid(Zeilen(i), 1, 10)
If IsNumeric(Zahlenstring) Then Zeileninhalt(i).Offset = CLng(Zahlenstring)
Zeileninhalt(i).Inhalt = "f"
Zahlenstring = Mid(Zeilen(i), 12, 5)
If IsNumeric(Zahlenstring) Then Zeileninhalt(i).Other = CLng(Zahlenstring)
End If
Next i
For i = 0 To AnzahlObjekte - 1
If Zeileninhalt(i).Inhalt = "n" Then
Test = FindeEnde(Filearray, Zeileninhalt(i).Offset, True)
Länge = Test - Zeileninhalt(i).Offset
ReDim Zeilenbyte(Länge - 1)
CopyMemory Zeilenbyte(0), Filearray(Zeileninhalt(i).Offset), Länge
teststring = StrConv(Zeilenbyte, vbUnicode) 'objektname
EndeGefunden = False
Do While EndeGefunden = False
Stand = Test + 1
If Filearray(Stand) = 10 Or Filearray(Stand) = 13 Then Stand = Stand + 1
Test = FindeEnde(Filearray, Stand)
Länge = Test - Stand
ReDim Zeilenbyte(Länge - 1)
CopyMemory Zeilenbyte(0), Filearray(Stand), Länge
Endebefehl = Stand + Länge
teststring = StrConv(Zeilenbyte, vbUnicode) 'inhalt
If InStr(teststring, ">>") Then
If Mid(teststring, Len(teststring) - 2 + 1, 2) = ">>" Then EndeGefunden = True
End If
If InStr(teststring, "endobj") Then
If Mid(teststring, Len(teststring) - 6 + 1, 6) = "endobj" Then EndeGefunden = True
End If
If InStr(teststring, "/") Then
Typ = SplitBefehle(teststring, Befehlsarray)
If Typ = "Text" Then
For z = 0 To UBound(Befehlsarray)
If InStr(Befehlsarray(z), "Length") Then
If Right(Befehlsarray(z), 2) = " R" Or IsNumeric(Right(Befehlsarray(z), 1)) Then
Anzahl = SplitBefehl(Befehlsarray(z), Befehleinzeln)
If Anzahl = 4 Then
If IsNumeric(Befehleinzeln(1)) Then
Hilfslong = Befehleinzeln(1)
Ende = FindeEnde(Filearray, Zeileninhalt(Hilfslong).Offset) 'objektname
Anfangzahl = Ende + 1
Ende = FindeEnde(Filearray, Ende + 1) 'objetinhalt
Endezahl = Ende
Ende = FindeEnde(Filearray, Ende + 1) 'Objektende

Längezahl = Endezahl - Anfangzahl
testbool = False
Do While testbool = False
If Filearray(Endebefehl) = 10 Or Filearray(Endebefehl) = 13 Then
Endebefehl = Endebefehl + 1
Else
testbool = True
End If
Loop

ReDim Zahlenarray(Längezahl - 1)
CopyMemory Zahlenarray(0), Filearray(Anfangzahl), Längezahl
Zahlenstring = Zahlenarray
Zahlenstring = StrConv(Zahlenstring, vbUnicode) 'Größe des komprimierten Strings
End If
End If
If Anzahl = 2 Then
If IsNumeric(Befehleinzeln(1)) Then Zahlenstring = Befehleinzeln(1)
End If
If IsNumeric(Zahlenstring) Then GrößeArray = CLng(Zahlenstring)
ReDim testarray(5)
CopyMemory testarray(0), Filearray(Endebefehl), 6
teststring = StrConv(testarray, vbUnicode)
If teststring = "stream" Then
Beginstream = Endebefehl + 6
End If
testbool = False
Do While testbool = False
If Filearray(Beginstream) = 10 Or Filearray(Beginstream) = 13 Then
Beginstream = Beginstream + 1
Else
testbool = True
End If
Loop
ReDim Textstream(GrößeArray - 1)
CopyMemory Textstream(0), Filearray(Beginstream), GrößeArray
UnCompressBytes Textstream, GrößeArray, GrößeArray * 12
cleantext Textstream
Hilfslong = FreeFile
Open App.Path & "\Seite" & CStr(Seitenzahl) & ".rtf" For Binary As Hilfslong
Put Hilfslong, , Textstream
Close Hilfslong
Seitenzahl = Seitenzahl + 1
End If

End If
Next z
End If
Else
End If
Loop

End If
Next i
Else
MsgBox "Error"
End If
End Sub

Private Function FindeEnde(testarray() As Byte, Anfang As Long, Optional Endzeichen As Boolean = False) As Long
Dim Test As Boolean
Dim Standpunkt As Long
Dim Zeichenende As String
Dim Testzeichen() As Byte
ReDim Testzeichen(1)
Standpunkt = Anfang
If Endzeichen = False Then
Do While Test = False
If testarray(Standpunkt) = 10 Or testarray(Standpunkt) = 13 Then
Test = True
Exit Do
Else
Standpunkt = Standpunkt + 1
End If
Loop
Else
Do While Test = False
If Standpunkt < UBound(testarray) Then
CopyMemory Testzeichen(0), testarray(Standpunkt), 2
Zeichenende = StrConv(Testzeichen, vbUnicode)
Else
ReDim Testzeichen(1)
End If
If testarray(Standpunkt) = 10 Or testarray(Standpunkt) = 13 Or Zeichenende = "<<" Then
Test = True
If Zeichenende = ">>" Then Standpunkt = Standpunkt + 1
Exit Do
Else
Standpunkt = Standpunkt + 1
End If
Loop
End If
FindeEnde = Standpunkt
End Function

Public Function SplitBefehle(teststring As String, Befehlsarray() As String) As String
Dim Endpunkt As Long
Dim Anfangspunkt As Long
Dim AnzahlBefehle As Long
Dim Befehlsteil As String
If Left(teststring, 2) = "<<" Then
teststring = Mid(teststring, 3)
End If
Anfangspunkt = 1
Anfangspunkt = InStr(Anfangspunkt, teststring, "/")
ReDim Befehlsarray(0)
Do While Anfangspunkt < Len(teststring) And Anfangspunkt > 0
Endpunkt = InStr(Anfangspunkt + 1, teststring, "/")
If Endpunkt <> 0 Then
Befehlsteil = Mid(teststring, Anfangspunkt + 1, Endpunkt - Anfangspunkt - 1)
CleanBefehl Befehlsteil
If Right(Befehlsteil, 1) = " " Then Befehlsteil = Mid(Befehlsteil, 1, Len(Befehlsteil) - 1)
ReDim Preserve Befehlsarray(AnzahlBefehle)
Befehlsarray(AnzahlBefehle) = Befehlsteil
AnzahlBefehle = AnzahlBefehle + 1

Else
Befehlsteil = Mid(teststring, Anfangspunkt + 1, Len(teststring) - Anfangspunkt)
CleanBefehl Befehlsteil
If Right(Befehlsteil, 1) = " " Then Befehlsteil = Mid(Befehlsteil, 1, Len(Befehlsteil) - 1)
ReDim Preserve Befehlsarray(AnzahlBefehle)
Befehlsarray(AnzahlBefehle) = Befehlsteil
AnzahlBefehle = AnzahlBefehle + 1
End If
If InStr(Befehlsteil, "Length") Then
SplitBefehle = "Text"
End If
If InStr(Befehlsteil, "Image") Then
SplitBefehle = "Bild"
End If
Anfangspunkt = Endpunkt
Loop
End Function

Public Function SplitBefehl(Befehl As String, Back() As String) As Long
Dim Anzahl As Long
Dim Wo As Long
Dim Startpunkt As Long
Dim Teil As String
Dim nichtdrin As Boolean
Startpunkt = 1
ReDim Back(0)
If InStr(Befehl, " ") Then
Do While nichtdrin = False
Wo = InStr(Startpunkt, Befehl, " ")
If Wo = 0 Then
nichtdrin = True
Teil = Mid(Befehl, Startpunkt, Len(Befehl) + 1 - Startpunkt)
ReDim Preserve Back(Anzahl)
Back(Anzahl) = Teil
Anzahl = Anzahl + 1
Else
Teil = Mid(Befehl, Startpunkt, Wo - Startpunkt)
ReDim Preserve Back(Anzahl)
Back(Anzahl) = Teil
Anzahl = Anzahl + 1

Startpunkt = Wo + 1
End If
Loop
Else
Back(0) = Befehl
Anzahl = 1
End If
SplitBefehl = Anzahl
End Function


Public Sub cleantext(Textstream() As Byte)
Dim AnfangsString As String
Dim EndString As String
Dim Anfangspunkt As Long
Dim Endpunkt As Long
Dim EndeString As Boolean
EndeString = False
Dim Übergabestream() As Byte
Dim Standpunkt As Long
Dim Stringlänge As Long

Anfangspunkt = 1
ReDim Übergabestream(0)
AnfangsString = StrConv("(", vbFromUnicode)
EndString = StrConv(")", vbFromUnicode)
Do While EndeString = False
Anfangspunkt = InStrB(Anfangspunkt, Textstream, AnfangsString)
If Anfangspunkt = 0 Then
EndeString = True
Exit Do
End If
Endpunkt = InStrB(Anfangspunkt, Textstream, EndString)
If Endpunkt = 0 Then
EndeString = True
Exit Do
End If
Stringlänge = Endpunkt - Anfangspunkt - 1
ReDim Preserve Übergabestream(Standpunkt + Stringlänge - 1)
CopyMemory Übergabestream(Standpunkt), Textstream(Anfangspunkt), Stringlänge
Anfangspunkt = Endpunkt
Standpunkt = Standpunkt + Stringlänge
Loop
ReDim Textstream(UBound(Übergabestream))
CopyMemory Textstream(0), Übergabestream(0), UBound(Übergabestream) + 1
End Sub


Public Sub CleanBefehl(Befehl As String)
Dim Wo As Long
Dim EndeGefunden As Boolean

Wo = InStr(Befehl, ">>")
If Wo <> 0 Then
Befehl = Left(Befehl, Wo - 1)
End If
Do While EndeGefunden = False
If Right(Befehl, 1) = " " Then
Befehl = Left(Befehl, Len(Befehl) - 1)
Else
EndeGefunden = True
End If
Loop
End Sub


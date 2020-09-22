VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type TIF_HEADER
     byte_order As Integer
     Version As Integer
     Offset As Long
End Type
Private Type TIF_ENTRY1
      tag As Integer
      type As Integer
      Length As Long
End Type
Private Type TIFENTRY2L
      Offset As Long
End Type
Private Type TIF_ENTRY2I
      Offset As Integer
End Type

Private m_width As Long
Private m_height As Long
Private m_BitsPerPixel As Long
Private m_compression As Long
Private SamplesPerPixel As Long
Private BitsPerSample As Long
Private PhotometricInterpretation As Integer
Private HasPalette As Boolean
Private RowsPerStrip As Long
Public Sub LoadTIFF(Filename As String)
Dim Palette() As Byte
Dim Red() As Integer
Dim Green() As Integer
Dim Blue() As Integer
Dim StripOffsets() As Long
Dim StripBytesCount() As Long
    Dim buf() As Byte
    Dim i As Long, k As Long, j As Long, entry As Integer, nc As Long, Nr As Long
    Dim tif_head As TIF_HEADER
    Dim tifen As TIF_ENTRY1
    Dim tife2I As TIF_ENTRY2I
    Dim tife2L As TIFENTRY2L
    Dim FGröße As Long
    Dim bw As Long, unreg As Long, Tmp As Long
    Dim intel As Boolean
    Dim offset1 As Long, Offset2 As Long
    Dim Filepointer As Long
    Dim Stand As Long
    Dim Filenumber As Long
    Dim Feldtypen As Integer
    Dim Typen As Integer
    Dim LongArray() As Long
    Dim Hilfslong As Long
    Dim z As Long
    Dim Größe As Long
    Dim Bilddaten() As Byte
    Dim Anzahl As Long
    Dim Bildstand As Long
    Dim Hilfsbyte As Byte
    Dim Hilfsint As Integer
    Dim Reihenrest As Long
    Dim Reihenfertig As Long
    
    HasPalette = False
    Form1.pic1.Picture = LoadPicture("")
    Filenumber = FreeFile
    Filepointer = 1
    Open Filename For Binary Access Read As Filenumber
        Get Filenumber, 1, tif_head
        Hilfsint = tif_head.byte_order
        SwapBytesInteger Hilfsint
        If Hilfsint = 18761 Then                       '"II" or 0x4949
            intel = True
        ElseIf Hilfsint = 19789 Then                   '"MM" or 0x4d4d
            intel = False
        Else                                    ' Not a TIFF file
            Close Filenumber
            Exit Sub
        End If
        If intel = False Then SwapBytesInteger tif_head.Version
        If tif_head.Version <> &H2A Then            ' Not a TIFF file
            Close Filenumber
            Exit Sub
        End If
        If intel = False Then SwapBytesLong tif_head.Offset
        offset1 = tif_head.Offset               'get the direction offset
        Filepointer = offset1 + 1
        Get Filenumber, Filepointer, entry
         If intel = False Then SwapBytesInteger entry
        Stand = Seek(Filenumber)
        For i = 0 To entry - 1                 'Deal with Entry
        Get Filenumber, Stand, tifen
        If intel = False Then
        SwapBytesInteger tifen.tag
        SwapBytesInteger tifen.type
        SwapBytesLong tifen.Length
        End If
        Stand = Stand + 8
        Select Case tifen.type
        Case 1, 2, 3, 6, 7, 7
        Get Filenumber, Stand, tife2I
                tife2L.Offset = tife2I.Offset
        Stand = Stand + 4
        Case Else
        Get Filenumber, Stand, tife2L
        
        Stand = Stand + 4
        End Select


    Select Case tifen.tag
    Case &H100 'ImageWidth
                    m_width = tife2L.Offset
    Case &H101 'ImageHeight
                    m_height = tife2L.Offset
    Case &H102 'BitsPerSample
    If tifen.Length = 1 Then
          BitsPerSample = tife2L.Offset
          Else
          BitsPerSample = tifen.Length
    End If
    Case &H103 'Compression
    m_compression = tife2L.Offset
    Case &H111 'StripOffsets
            ReDim StripOffsets(tifen.Length - 1)
            Select Case tifen.Length
            Case 1
            StripOffsets(0) = tife2L.Offset
            Case Else
            Get Filenumber, tife2L.Offset + 1, StripOffsets
            If intel = False Then
            For z = 0 To tifen.Length - 1
            SwapBytesLong (StripOffsets(z))
            Next z
            End If
            End Select
    Case &H106 'PhotometricInterpretation
    PhotometricInterpretation = tife2L.Offset
            Case &H115 'SamplesPerPixel
        SamplesPerPixel = tife2L.Offset
        Case &H116
        RowsPerStrip = tife2L.Offset
    Case &H117 'StripByteCount
        ReDim StripBytesCount(tifen.Length - 1)
        Select Case tifen.Length
        Case 1
        StripBytesCount(0) = tife2L.Offset
        Case Else
            Get Filenumber, tife2L.Offset + 1, StripBytesCount
            If intel = False Then
            For z = 0 To tifen.Length - 1
            SwapBytesLong (StripBytesCount(z))
            Next z
            End If
        End Select
    Case &H140 'ColorMap
    HasPalette = True
    ReDim Palette(tifen.Length - 1)
    ReDim Red(tifen.Length / 3 - 1)
    ReDim Green(tifen.Length / 3 - 1)
    ReDim Blue(tifen.Length / 3 - 1)

    Get Filenumber, tife2L.Offset + 2, Blue
    Get Filenumber, , Green
    Get Filenumber, , Red
    Hilfslong = 0
    For z = 0 To UBound(Red)
    CopyMemory Palette(Hilfslong), ByVal VarPtr(Blue(z)), 1
    CopyMemory Palette(Hilfslong + 1), ByVal VarPtr(Green(z)), 1
    CopyMemory Palette(Hilfslong + 2), ByVal VarPtr(Red(z)), 1
    Hilfslong = Hilfslong + 3
    Next z
    End Select
    Next i
        If BitsPerSample = 3 Then
    Select Case SamplesPerPixel
    Case 1
    m_BitsPerPixel = 8
    Case 3
    m_BitsPerPixel = 24
    End Select
    Else
    m_BitsPerPixel = BitsPerSample * SamplesPerPixel
    End If

    If UBound(StripOffsets) = UBound(StripBytesCount) Then
    Anzahl = UBound(StripBytesCount) + 1
    'For i = 0 To Anzahl - 1
    'Größe = Größe + StripBytesCount(i)
    'Next i
    ReDim Bilddaten(0)
    FGröße = -1
    Bildstand = 0
    Reihenrest = Me.ImageHeight
    For i = 0 To Anzahl - 1
    ReDim buf(StripBytesCount(i) - 1)
    Get Filenumber, StripOffsets(i) + 1, buf
    Select Case Me.Compression
    Case 5
    
    Case -32763
    If Reihenrest > RowsPerStrip Then
    DecompressRLE buf, 8, RowsPerStrip, Me.ImageWidth
    Reihenrest = Reihenrest - RowsPerStrip
    Else
    DecompressRLE buf, 8, Reihenrest, Me.ImageWidth
    Reihenrest = Reihenrest - Reihenrest
    End If
    End Select
    ReDim Preserve Bilddaten(Bildstand + UBound(buf))
    CopyMemory Bilddaten(Bildstand), buf(0), UBound(buf) + 1
    Bildstand = Bildstand + UBound(buf) + 1
    Next i
    
    End If
    Close Filenumber
    Dim b As Long
    b = FreeFile
    Open "c:\komp.txt" For Binary As b
    Put b, , Bilddaten
    Close b
    Select Case Me.Compression
    Case 1, -32763
    Select Case Me.BitsPerPixel
    Case 1
    MakeBitmap Bilddaten, Me.ImageHeight, DataPerRow
InitColorTable_1 PhotometricInterpretation
CreateBitmap_1 Bilddaten, Me.ImageWidth, Me.ImageHeight, 0
    DrawBitmap Me.ImageWidth, Me.ImageHeight, Form1.pic1, True
    Case 8
    MakeBitmap Bilddaten, Me.ImageHeight, Me.ImageWidth
    If HasPalette = True Then
    Größe = UBound(Palette) + 1
    InitColorTable_8 Palette
    Else
    Größe = 768
    InitColorTable_Grey 8
    End If
    CreateBitmap_8 Bilddaten, Me.ImageWidth, Me.ImageHeight, 0, 1
    DrawBitmap Me.ImageWidth, Me.ImageHeight, Form1.pic1, True
    Case 24
    Größe = UBound(Bilddaten) + 1
    For i = 0 To Größe - 2 Step 3
    Hilfsbyte = Bilddaten(i)
    Bilddaten(i) = Bilddaten(i + 2)
    Bilddaten(i + 2) = Hilfsbyte
    Next i
    MakeBitmap Bilddaten, Me.ImageHeight, Me.ImageWidth * 3
    CreateBitmap_24 Bilddaten, Me.ImageWidth, Me.ImageHeight, 0, 1
    DrawBitmap Me.ImageWidth, Me.ImageHeight, Form1.pic1, True
    End Select
    End Select
End Sub

Public Sub SwapBytesLong(ByteValue As Long)
Dim Übergabe As Long
Dim i As Long
For i = 0 To 3
CopyMemory ByVal VarPtr(Übergabe) + i, ByVal VarPtr(ByteValue) + (3 - i), 1
Next i
ByteValue = Übergabe
End Sub
Private Sub SwapBytesInteger(Val As Integer)
Dim Bytef(1) As Byte
Dim u As Byte
CopyMemory Bytef(0), ByVal VarPtr(Val), 2
u = Bytef(0)
Bytef(0) = Bytef(1)
Bytef(1) = u
CopyMemory ByVal VarPtr(Val), Bytef(0), 2
End Sub
Public Property Get ImageWidth() As Long
ImageWidth = m_width
End Property

Public Property Get ImageHeight() As Long
ImageHeight = m_height
End Property
Private Sub MakeBitmap(ImageArray() As Byte, Lines As Long, BytesLine As Long)
Dim Übergabe() As Byte
Dim Größe As Long
Dim GrößeBMP As Long
Dim i As Long
Dim Zugabe As Integer
Dim Standort As Long
Dim nBitmapX As Long
    
    If (BytesLine) Mod Len(nBitmapX) = 0 Then
    
        nBitmapX = BytesLine - 1
   
    Else
   
        nBitmapX = (BytesLine \ 4) * 4 + 3
   
    End If
If nBitmapX + 1 <> BytesLine Then
Zugabe = nBitmapX - BytesLine + 1
End If
Größe = Lines * BytesLine

GrößeBMP = Lines * (nBitmapX + 1) - 1
ReDim Übergabe(UBound(ImageArray))

CopyMemory Übergabe(0), ImageArray(0), UBound(ImageArray) + 1

Dim l As Long
ReDim ImageArray(GrößeBMP)
For i = 0 To BytesLine * Lines - BytesLine Step BytesLine
CopyMemory ImageArray(Standort), Übergabe(Größe - i - BytesLine), BytesLine
Standort = Standort + nBitmapX + 1

Next i

End Sub
Public Property Get BitsPerPixel() As Long
BitsPerPixel = m_BitsPerPixel
End Property

Public Property Get Compression() As Long
Compression = m_compression
End Property
Private Function DataPerRow() As Long
DataPerRow = (Me.ImageWidth * Me.BitsPerPixel() + 7) \ 8 ' + 1
End Function

Private Sub DecompressRLE(RLEStream() As Byte, Bits As Long, Höhe As Long, Breite As Long)
    Dim InitSize As Long
    Dim Temp() As Byte
    Dim n As Long
    Dim k As Boolean
    Dim b As Long
    Dim l As Long
    Dim fertig As Long
    Dim z As Long
    Dim Länge As Long
    Dim Byteanzahl As Long
    On Error Resume Next
    Byteanzahl = Bits / 8
       InitSize = CLng(Höhe * Breite * Byteanzahl)
    ReDim Temp(0 To InitSize)
    Do While fertig < InitSize
    
    If l > UBound(RLEStream) Then GoTo Ende
        z = 0
        If RLEStream(l) > 128 Then
            n = 257 - RLEStream(l)
            For b = 0 To n - 1
            
            CopyMemory Temp(fertig), RLEStream(l + 1), Byteanzahl
                           fertig = fertig + Byteanzahl
            
            Next b
        k = True
        ElseIf RLEStream(l) = 128 Then
        'Mache nichts
        Else
           n = RLEStream(l) + 1
           Länge = n * Byteanzahl

            CopyMemory Temp(fertig), RLEStream(l + 1), Länge
            
            k = False
            z = z + Länge

            fertig = fertig + z
            End If
            If k = True Then
                l = Byteanzahl + 1 + l
                n = z + (n * Byteanzahl) + 1
            Else
                l = (n * Byteanzahl + 1) + l
                n = z + n
            End If
    Loop
Ende:
ReDim Preserve Temp(UBound(Temp) - 1)
    RLEStream = Temp
    
End Sub


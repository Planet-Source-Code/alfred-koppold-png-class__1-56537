VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadTGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Autor: ALKO
'e-mail: alfred.koppold@freenet.de

Option Explicit
'Types
Private Type ScTw
TwipWidth As Long
TwipHeight As Long
End Type
Private Type BITMAPINFOHEADER
 Size As Long
 Width As Long
 Height As Long
 Planes As Integer
 BitCount As Integer
 Compression As Long
 SizeImage As Long
 XPelsPerMeter As Long
 YPelsPerMeter As Long
 ClrUsed As Long
 ClrImportant As Long
End Type

Private Type RGBTriple
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type RGBQUAD
rgbBlue As Byte
rgbGreen As Byte
rgbRed As Byte
rgbReserved As Byte
End Type

Private Type tgaheader                  'declare the header
  info       As Byte      'length of image information block
  colortyp   As Byte      'DAC table or BGR format
  imagetyp   As Byte      'compressed or uncompressed
  origin     As Integer         'first entry in the DAC table
  colnumber  As Integer         'number of colors in the DAC table
  entrybits  As Byte      'entry size in the DAC table
  xvalue     As Integer         'x co-ordinate lower left corner
  yvalue     As Integer         'y co-ordinate lower left corner
  widt       As Integer         'image width
  Height     As Integer         'image height
  pixelsize  As Byte      'number of bits per pixel
  descriptor As Byte      'image descriptor
End Type
Private Type TGA_FOOTER
ExtOffset As Long
DevOffset As Long
TGASignatur As String * 18
End Type
Private Type Palette2
Byte1 As Byte
Byte2 As Byte
End Type

Private Type Palette3
Blau As Byte
Grün As Byte
Rot As Byte
End Type

Private Type Palette4
Blau As Byte
Grün As Byte
Rot As Byte
Attribut As Byte
End Type

'Functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long

'Variables
Private Orientation As Integer
Private Übergabe() As Byte
Private BitmapData() As Byte
Private Header As tgaheader         'define the header
Private Pal1() As Palette3
Private Pal2() As Palette4
Private Pal3() As Palette2
Private i As Long
Private nWidth As Long
Private nHeight As Long
Private Scales As Integer
Private Automatic As Boolean
Private TW As ScTw
Private tgafooter As TGA_FOOTER
Private m_alpha As Boolean
Private m_OBCol As Long
Private m_Bgx As Long
Private m_Bgy As Long
Private m_settoBG As Boolean
Private m_isalpha As Boolean
Private m_bkgdColor As Long
Public Function LoadTGA(ByVal Filename As String) As StdPicture
    Dim nFreefile As Integer
    Dim filesize As Long
    nFreefile = FreeFile
    m_isalpha = False
    Open Filename For Binary Lock Write As #nFreefile
    filesize = LOF(nFreefile)
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    Get #nFreefile, filesize - 25, tgafooter
    Close #nFreefile
TW = PixelToTwips(CLng(Header.widt), CLng(Header.Height))
Select Case Header.pixelsize
Case 8
Read8Bit Filename
Case 16
Read16bit Filename
Case 24
Read24bit Filename
Case 32
Read32bit Filename
End Select
End Function




Private Sub Read8Bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Palette8(0 To 255) As RGBTriple
    Dim PalByte As Byte
    Dim result As Long
    Dim NYSize As Long
    
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
Select Case Header.entrybits / 8
    Case 2
    ReDim Pal3(Header.colnumber - 1)
    Get #nFreefile, , Pal3
    'Palette erstellen
    For i = 0 To UBound(Pal3)
    If GetByte(Pal3(i).Byte1, 1) = 1 Then Palette8(i).Blue = 16
    If GetByte(Pal3(i).Byte1, 2) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 8
    If GetByte(Pal3(i).Byte1, 3) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 4
    If GetByte(Pal3(i).Byte1, 4) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 2
    If GetByte(Pal3(i).Byte1, 5) = 1 Then Palette8(i).Blue = Palette8(i).Blue + 1
    If GetByte(Pal3(i).Byte1, 6) = 1 Then Palette8(i).Green = 16
    If GetByte(Pal3(i).Byte1, 7) = 1 Then Palette8(i).Green = Palette8(i).Green + 8
    If GetByte(Pal3(i).Byte1, 8) = 1 Then Palette8(i).Green = Palette8(i).Green + 4
    If GetByte(Pal3(i).Byte2, 1) = 1 Then Palette8(i).Green = Palette8(i).Green + 2
    If GetByte(Pal3(i).Byte2, 2) = 1 Then Palette8(i).Green = Palette8(i).Green + 1
    If GetByte(Pal3(i).Byte2, 3) = 1 Then Palette8(i).Red = 16
    If GetByte(Pal3(i).Byte2, 4) = 1 Then Palette8(i).Red = Palette8(i).Red + 8
    If GetByte(Pal3(i).Byte2, 5) = 1 Then Palette8(i).Red = Palette8(i).Red + 4
    If GetByte(Pal3(i).Byte2, 6) = 1 Then Palette8(i).Red = Palette8(i).Red + 2
    If GetByte(Pal3(i).Byte2, 7) = 1 Then Palette8(i).Red = Palette8(i).Red + 1
Palette8(i).Red = Palette8(i).Red * 4
Palette8(i).Green = Palette8(i).Green * 4
Palette8(i).Blue = Palette8(i).Blue * 4
    Next i
    Case 3
    ReDim Pal1(Header.colnumber - 1)
    Get #nFreefile, , Pal1
    'Farbfolge ändern
    For i = 0 To UBound(Pal1)
    Palette8(i).Red = Pal1(i).Rot
    Palette8(i).Blue = Pal1(i).Blau
    Palette8(i).Green = Pal1(i).Grün
    Next i

    Case 4
    ReDim Pal2(Header.colnumber - 1)
    Get #nFreefile, , Pal2
    'Farbfolge ändern
    For i = 0 To UBound(Pal2)
    Palette8(i).Red = Pal2(i).Rot
    Palette8(i).Blue = Pal2(i).Blau
    Palette8(i).Green = Pal2(i).Grün
    Next i


End Select
  If Header.entrybits <> 0 Then
    ReDim BitmapData(LOF(nFreefile) - Len(Header) - (UBound(Pal1) * 3))
    Else
    'graue Palette erstellen
    For i = 0 To 255
    Palette8(i).Blue = i
    Palette8(i).Green = i
    Palette8(i).Red = i
    Next i
        ReDim BitmapData(LOF(nFreefile) - Len(Header)) 'keine Palette vorhanden

    End If
     Get #nFreefile, , BitmapData()
        Orientation = GetByte(Header.descriptor, 3)

    Close #nFreefile
        With Header
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt
    End With
        ReDim Übergabe(UBound(Palette8) * 3 + 3)
        CopyMemory Übergabe(0), Palette8(0), UBound(Übergabe)
        InitColorTable_8 Übergabe
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    BitmapData = DecompressTGA(BitmapData(), 8, nHeight, nWidth)
    End If
    MakeBitmap BitmapData, nHeight, nLineSize
    CreateBitmap_8 BitmapData, nWidth, nHeight, Orientation

End Sub


Private Sub Read16bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Data() As Byte
    
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
    With Header
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt * .pixelsize
    End With
    ReDim Data(LOF(nFreefile) - Len(Header))
    
    Get #nFreefile, , Data()
       Orientation = GetByte(Header.descriptor, 3)

    Close #nFreefile
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    Data = DecompressTGA(Data(), 24, nHeight, nWidth)
    End If
    MakeBitmap Data, nHeight, nWidth * 2
CreateBitmap_16 Data, nWidth, nHeight, Orientation

End Sub

Private Sub Read24bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Data() As Byte
    Dim quad() As RGBQUAD
    
    nFreefile = FreeFile

    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
    With Header
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt * .pixelsize
    End With
    ReDim Data(LOF(nFreefile) - Len(Header))
    
    Get #nFreefile, , Data()
   Orientation = GetByte(Header.descriptor, 3)
    Close #nFreefile
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    Data = DecompressTGA(Data(), 24, nHeight, nWidth)
    End If
    'Sort from planes into a bitmap
'==================================================
    Dim x As Long, y As Long, nStartPos As Long
    ReDim quad(UBound(Data) / 3)
    For x = 0 To UBound(Data) / 3 - 1
        
        
            With quad(x)
                .rgbBlue = Data(x * 3)
                .rgbGreen = Data(x * 3 + 1)
                .rgbRed = Data(x * 3 + 2)
            End With
        
        Next x
           ReDim BitmapData(UBound(quad) * 4 + 4)
   CopyMemory BitmapData(0), quad(0), UBound(BitmapData)

CreateBitmap_24 BitmapData, nWidth, nHeight, Orientation

End Sub

Private Sub Read32bit(Filename As String)
    Dim nFreefile As Integer
    Dim nLineSize As Long
    Dim Data() As Byte
    Dim quad() As RGBQUAD
    Dim Bildgröße As Long
    nFreefile = FreeFile
    Dim x As Long, y As Long, nStartPos As Long
    m_isalpha = True
    Open Filename For Binary Lock Write As #nFreefile
    
    Seek #nFreefile, 1
    Get #nFreefile, , Header
    
    With Header
    'Bildgröße = CLng(.widt) * CLng(.Height) * CLng(.pixelsize) / 8
        nWidth = .widt - .xvalue ' + 1
        nHeight = .Height - .yvalue ' + 1
        nLineSize = .widt * .pixelsize
    End With
    ReDim Data(LOF(nFreefile) - Len(Header))
    Get #nFreefile, , Data()
       Orientation = GetByte(Header.descriptor, 3)

    Close #nFreefile
If Header.imagetyp = 9 Or Header.imagetyp = 10 Then
    Data = DecompressTGA(Data(), 32, nHeight, nWidth)
    End If
    'Sort from planes into a bitmap
'==================================================
    If m_alpha = True Then
    MakeAlpha Form1.pic1, Data
    ReDim BitmapData(UBound(Data))
    CopyMemory BitmapData(0), Data(0), UBound(Data) + 1
    CreateBitmap_24 BitmapData, nWidth, nHeight, Orientation, 1
    Else
    ReDim quad(UBound(Data) / 4)
    For x = 0 To UBound(Data) / 4 - 1
            With quad(x)
                .rgbBlue = Data(x * 4)
                .rgbGreen = Data(x * 4 + 1)
                .rgbRed = Data(x * 4 + 2)
            End With
                Next x
                           ReDim BitmapData(UBound(quad) * 4 + 4)
   CopyMemory BitmapData(0), quad(0), UBound(BitmapData)
CreateBitmap_24 BitmapData, nWidth, nHeight, Orientation
    End If
End Sub


Private Function DecompressTGA(RLEStream() As Byte, Bits As Long, Höhe As Long, Breite As Long) As Byte()
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
        If RLEStream(l) > 127 Then
            n = RLEStream(l) - 127
            For b = 0 To n - 1
            
            CopyMemory Temp(fertig), RLEStream(l + 1), Byteanzahl
                           fertig = fertig + Byteanzahl
            
            Next b
        k = True
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

    DecompressTGA = Temp
    
End Function

Private Sub MakeBitmap(ImageArray() As Byte, Lines As Long, BytesLine As Long)
Dim Übergabe() As Byte
Dim GrößeBMP As Long
Dim Zugabe As Integer
Dim Standort As Long
Dim nBitmapX As Long
Dim l As Long

    If (BytesLine) Mod 4 = 0 Then
            'wenn duch 4 teilbar dann ok
        nBitmapX = BytesLine - 1
       Else
           'ansonsten ändern
        nBitmapX = (BytesLine \ 4) * 4 + 3
    End If
If nBitmapX + 1 <> BytesLine Then
Zugabe = nBitmapX - BytesLine + 1
End If
GrößeBMP = Lines * (nBitmapX + 1) - 1
ReDim Übergabe(UBound(ImageArray))

CopyMemory Übergabe(0), ImageArray(0), UBound(ImageArray) + 1
ReDim ImageArray(GrößeBMP)
For i = 0 To BytesLine * Lines - BytesLine Step BytesLine
CopyMemory ImageArray(Standort), Übergabe(i), BytesLine
Standort = Standort + nBitmapX + 1

Next i

End Sub


Private Function GetByte(Bytes As Byte, Position As Long) As Integer
GetByte = 0
Select Case Position
Case 1
If Bytes And 128 Then GetByte = 1
Case 2
If Bytes And 64 Then GetByte = 1
Case 3
If Bytes And 32 Then GetByte = 1
Case 4
If Bytes And 16 Then GetByte = 1
Case 5
If Bytes And 8 Then GetByte = 1
Case 6
If Bytes And 4 Then GetByte = 1
Case 7
If Bytes And 2 Then GetByte = 1
Case 8
If Bytes And 1 Then GetByte = 1
End Select

End Function



Public Function DrawTGA(PicObj As Object)
If Me.IsAlpha = True And m_alpha = True Then
DrawBitmap nWidth, nHeight, PicObj, Automatic, m_Bgx, m_Bgy, m_settoBG
Else
DrawBitmap nWidth, nHeight, PicObj, Automatic
End If
End Function

Public Property Get TGAWidth() As Long
Select Case Scales
Case 0
TGAWidth = nWidth
Case 1
TGAWidth = TW.TwipWidth
End Select
End Property



Public Property Get TGAHeight() As Long
Select Case Scales
Case 0
TGAHeight = nHeight
Case 1
TGAHeight = TW.TwipHeight
End Select
End Property

Public Property Get Compressed() As Boolean
Select Case Header.imagetyp
Case 9
Compressed = True
Case 10
Compressed = True
Case Else
Compressed = False
End Select

End Property


Public Property Get Bpp() As Integer
Bpp = CInt(Header.pixelsize)
End Property


Public Property Get IsTGA() As Boolean
Dim Test As Boolean
Test = True
    Select Case Header.imagetyp
    Case 1
    Case 2
    Case 3
    Case 9
    Case 10
    Case 11
    Case 32
    Case 33
    Case Else
    Test = False
    End Select

    Select Case Header.pixelsize
    Case &H8
    '8bpp
    Case &H10
    '16bpp
    Case &H18
    '24bpp
    Case &H20
    '32bpp
    Case Else
    Test = False
    End Select
    IsTGA = Test
End Property


Public Property Get ScaleMode() As Integer
ScaleMode = Scales
End Property

Public Property Let ScaleMode(ByVal vNewValue As Integer)
If vNewValue > 0 Then vNewValue = 1
If vNewValue <> Scales Then
Scales = vNewValue
End If
End Property


Public Property Get Autoscale() As Boolean
Autoscale = Automatic
End Property

Public Property Let Autoscale(ByVal vNewValue As Boolean)
Automatic = vNewValue
End Property

Private Sub Class_Initialize()
Automatic = True
Scales = 1
End Sub
Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
PixelToTwips.TwipWidth = xwert * ux
uy = Screen.TwipsPerPixelY
PixelToTwips.TwipHeight = ywert * uy
End Function

Private Sub MakeAlpha(PicObject As Object, Buffer() As Byte, Optional x As Long = 0, Optional y As Long = 0)
Dim Myx As Long, Myy As Long, DatOff As Long
Dim R As Long, G As Long, b As Long, a As Long
Dim sR As Long, sG As Long, sB As Long
Dim dR As Long, dG As Long, dB As Long
Dim DestData() As Byte, bytesperrow As Long
Dim DestOff As Long, DestHdr As BITMAPINFOHEADER
Dim MemDC As Long, hBmp As Long, hOldBmp As Long
Dim SrcData() As Byte
Dim hdc As Long
Dim Hohe As Long
Dim Breite As Long
Hohe = nHeight
Breite = nWidth
On Error Resume Next
hdc = PicObject.hdc
If Err.Number = 91 Then
 ReDim SrcData(UBound(Buffer))
  bytesperrow = LineBytes(Breite, 24)
If m_OBCol = False Then
  FillColorArray SrcData, Me.BkgdColor, bytesperrow
Else
  FillColorArray SrcData, m_OBCol, bytesperrow
End If
 ReDim DestData(bytesperrow * Hohe - 1)
Err.Clear
Else
If PicObject.Width < Breite * Screen.TwipsPerPixelX Then
 PicObject.Width = Screen.TwipsPerPixelX * Breite + 100
End If
If PicObject.Height < Hohe * Screen.TwipsPerPixelY Then
 PicObject.Height = Screen.TwipsPerPixelY * Hohe + 100
End If
 hdc = PicObject.hdc
 bytesperrow = LineBytes(Breite, 24)
 ReDim DestData(bytesperrow * Hohe - 1)
 ReDim SrcData(UBound(Buffer))
 DestHdr.BitCount = 24
 DestHdr.Height = Hohe
 DestHdr.Width = Breite
 DestHdr.Planes = 1
 DestHdr.Size = 40
 MemDC = CreateCompatibleDC(hdc)
 hBmp = CreateCompatibleBitmap(hdc, Breite, Hohe)
 hOldBmp = SelectObject(MemDC, hBmp)
 BitBlt MemDC, 0, 0, Breite, Hohe, hdc, x, y, vbSrcCopy
 GetDIBits MemDC, hBmp, 0, Hohe, SrcData(0), DestHdr, 0
SelectObject hOldBmp, MemDC
 DeleteObject hBmp
 DeleteDC MemDC
End If
 For Myy = 0 To Hohe - 1
 For Myx = 0 To Breite - 1
 DestOff = Myy * bytesperrow + Myx * 3
 sR = SrcData(DestOff + 2)
 sG = SrcData(DestOff + 1)
 sB = SrcData(DestOff)
 b = Buffer(DatOff)
 G = Buffer(DatOff + 1)
 R = Buffer(DatOff + 2)
 a = Buffer(DatOff + 3)
If a = 255 Then
 DestData(DestOff + 2) = R
 DestData(DestOff + 1) = G
 DestData(DestOff) = b
 ElseIf a = 0 Then
 DestData(DestOff + 2) = sR
 DestData(DestOff + 1) = sG
 DestData(DestOff) = sB
 Else
 dR = R * a + (255 - a) * sR + 255
 dG = G * a + (255 - a) * sG + 255
 dB = b * a + (255 - a) * sB + 255
 CopyMemory DestData(DestOff + 2), ByVal VarPtr(dR) + 1, 1
 CopyMemory DestData(DestOff + 1), ByVal VarPtr(dG) + 1, 1
 CopyMemory DestData(DestOff), ByVal VarPtr(dB) + 1, 1
End If
 DatOff = DatOff + 4
Next Myx
Next Myy
 Buffer = DestData
End Sub

Private Function LineBytes(Width As Long, BitCount As Integer) As Long
LineBytes = ((Width * BitCount + 31) \ 32) * 4
End Function


Public Property Get BkgdColor() As Long
BkgdColor = m_bkgdColor
End Property

Private Sub FillColorArray(FArray() As Byte, Color As Long, bytesperrow As Long)
Dim DA(3) As Byte
Dim i As Long
Dim u As Byte
Dim Zähler As Long
CopyMemory DA(0), ByVal VarPtr(Color), 3
If DA(3) = 0 Then
u = DA(0)
DA(0) = DA(2)
DA(2) = u
u = DA(1)
If DA(0) = DA(1) And DA(1) = DA(2) Then
FillMemory FArray(0), UBound(FArray) + 1, DA(0)
Else
Zähler = 1
For i = 0 To UBound(FArray) - 2 Step 3
CopyMemory FArray(i), DA(0), 3
If i = ((Zähler * bytesperrow) - 1) Or i = ((Zähler * bytesperrow) - 2) Then
i = Zähler * bytesperrow
i = bytesperrow * Zähler
Zähler = Zähler + 1
End If
Next i
End If
End If
End Sub

Public Sub SetToBkgrnd(SetToBG As Boolean, Optional x As Long = 0, Optional y As Long = 0)
m_Bgx = x
m_Bgy = y
m_settoBG = SetToBG
End Sub
Public Property Let DrawAlpha(ByVal vNewValue As Boolean)
m_alpha = vNewValue
End Property
Public Property Get IsAlpha() As Boolean
IsAlpha = m_isalpha
End Property


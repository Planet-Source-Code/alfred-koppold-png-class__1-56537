Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Function UnCompressBytes(Buffer() As Byte, CompressedSize As Long, UncompressedSize As Long) As Boolean
   
   Dim b() As Byte

   Dim BufferSize As Long
   Dim FileSize As Long
   
   Dim crc As Long
   Dim fh As Long
   Dim r As Long

If Buffer(0) <> 120 Then
   ReDim b(UBound(Buffer) + 2)
   'Zlib's Uncompress method expects the 2 byte head that the Compress method adds
   'so we put that on first. Luckily it's always the same value.
   b(0) = 120
   b(1) = 156
      CopyMemory b(2), Buffer(0), UBound(Buffer) + 1
    Else
       ReDim b(UBound(Buffer))
      CopyMemory b(0), Buffer(0), UBound(Buffer) + 1
    End If
    
   FileSize = UBound(Buffer) + 3
   BufferSize = UncompressedSize * 1.01 + 12
   ReDim Buffer(BufferSize - 1) As Byte
   
   'r = UnCompress(Buffer(0), BufferSize, b(0), FileSize)
 Inflate b, BufferSize
 Buffer = b
   'ReDim Preserve Buffer(CentralFileHeader.UnCompressedSize - 1)
   'crc = lCRC32(0&, Buffer(0), UBound(Buffer) + 1)
   'If crc = CRC32 Then
      'UnCompressBytes = True
   'End If

End Function


' Name: LZW (Lempel, Ziv, Welch) Dictionary Compression
' Description:The Lempel, Ziv, Welch compression algorithm i
'     s considered the most efficcient all purpose compression alg
'     orithm there is.
' By: Asgeir Bjarni Ingvarsson
'
' Inputs:None
' Returns:None
' Assumes:None
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************

'     ' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'     '| LZW - Compression/Uncompression|
'     '|-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-|
'     '|Author: Asgeir B. Ingvarsson |
'     '| |
'     '|E-Mail: abi@islandia.is |
'     '| |
'     '|Address: Hringbraut 119 |
'     '| IS-107, Reykjavik|
'     '| ICELAND |
'     '|-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-|
'     '|For any comments or questions, please contact me |
'     '|using either of the above measures. |
'     '|-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-|
'     '|This code has one flaw, it can't process characters |
'     '|higher than 127. |
'     '|For the code that can compress all 256 ascii chars. |
'     '|please e-mail me.|
'     '|If you use this code or modify it, I would appreciate|
'     '|it if you would mention my name somewhere and send me|
'     '|a copy of the code (if it has been modified).|
'     '|-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-|
'     '|LZW is property of Unisys and is free for|
'     '|noncommercial software. |
'     ' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Private Dict(0 To 255) As String
Private Count As Integer

Private Sub Init()


             For i = 0 To 127
                    Dict(i) = Chr(i)
             Next

End Sub


Private Function Search(inp As String) As Integer


             For i = 0 To 255

                           If Dict(i) = inp Then Search = i: Exit Function
                           Next

                    Search = 256
             End Function


Private Sub Add(inp As String)


             If Count = 256 Then Wipe
                    Dict(Count) = inp
                    Count = Count + 1
             End Sub


Private Sub Wipe()


             For i = 128 To 255
                    Dict(i) = ""
             Next

      Count = 128
End Sub


Public Function Deflate(inp As String) As String

      '     'Begin Error Checking

             If Len(inp) = 0 Then Exit Function

                           For i = 1 To Len(inp)

                                         If Asc(Mid(inp, i, 1)) > 127 Then MsgBox "Illegal Character Value", vbCritical, "Error:": Exit Function
                                         Next

                                  '     'End Error Checking
                                  Init
                                  Wipe
                                  p = ""
                                  i = 1

                                         Do Until i > Len(inp)
                                                c = Mid(inp, i, 1)
                                                i = i + 1
                                                temp = p & c

                                                       If Not Search(CStr(temp)) = 256 Then
                                                              p = temp
                                                       Else
                                                              o = o & Chr(Search(CStr(p)))
                                                              Add CStr(temp)
                                                              p = c
                                                       End If

                                         Loop

                                  o = o & Chr(Search(CStr(p)))
                                  Deflate = o
                           End Function


Public Function Inflate(inp As String) As String


             If Len(inp) = 0 Then Exit Function
                    Init
                    Wipe
                    cW = Asc(Mid(inp, 1, 1))
                    o = Dict(cW)
                    i = 2

                           Do Until i > Len(inp)
                                  pW = cW
                                  cW = Asc(Mid(inp, i, 1))
                                  i = i + 1

                                         If Not Dict(cW) = "" Then
                                                o = o & Dict(cW)
                                                p = Dict(pW)
                                                c = Mid(Dict(cW), 1, 1)
                                                Add (CStr(p) & CStr(c))
                                         ElseIf Dict(cW) = "" Then
                                                p = Dict(pW)
                                                c = Mid(Dict(pW), 1, 1)
                                                o = o & p & c
                                                Add (CStr(p) & CStr(c))
                                         End If

                           Loop

                    Inflate = o
             End Function


Public Sub main()

      inp = "Hello World, Hello World"
      d = Deflate(CStr(inp)) 'Compress
      q = Inflate(CStr(d)) 'Uncompress
      MsgBox "Uncompressed: " & q & vbCrLf & vbCrLf & _
      "Compressed: " & d & vbCrLf & vbCrLf & _
      "Compressed Size: " & Len(d) & vbCrLf & vbCrLf & _
      "Uncompressed Size: " & Len(q) & vbCrLf & vbCrLf & _
      "Compression Ratio: " & (100 - (((Len(d) / Len(q)) * 100) \ 1)) & "%", vbOKOnly, "Results:"
End Sub

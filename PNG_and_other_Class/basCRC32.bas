Attribute VB_Name = "basCRC32"
Option Explicit
Option Base 0

' basCRC32: Calculates CRC-32 checksum for a given message string
' Version 1. Published 6 May 2001.
'************************COPYRIGHT NOTICE*************************
' Copyright (C) 2001 DI Management Services Pty Ltd,
' Sydney Australia <www.di-mgt.com.au>. All rights reserved.
' This code was originally written in Visual Basic by David Ireland.
' You are free to use this code in your applications without liability
' or compensation, but the courtesy of both notification of use and
' inclusion of due credit are requested. You must keep this copyright
' notice intact.
' It is PROHIBITED to distribute or reproduce this code for profit
' or otherwise, on any web site, ftp server or BBS, or by any
' other means, including CD-ROM or other physical media, without the
' EXPRESS WRITTEN PERMISSION of the author.
' Use at your own risk.
' David Ireland and DI Management Services Pty Limited
' offer no warranty of its fitness for any purpose whatsoever,
' and accept no liability whatsoever for any loss or damage
' incurred by its use.
' If you use it, or found it useful, or can suggest an improvement
' please let us know at <code@di-mgt.com.au>.
'*****************************************************************

Private aCRC32Table(255) As Long

Public Function CRC32(sMessage As String) As Long
' Given table is already setup.
' Set iCRC = 0xffffffff
' For each byte in message do:
'   calculate iCRC = (iCRC >> 8) ^ Table[(iCRC & 0xFF) ^ byte]
' Return iCRC ^ 0xffffffff
    
    Dim iCRC As Long
    Dim i As Integer
    Dim bytT As Byte
    Dim bytC As Byte
    Dim lngA As Long
    
    Call CRC32Setup     ' Note static flag in setup to avoid doing twice
    
    iCRC = &HFFFFFFFF
    For i = 1 To Len(sMessage)
        bytC = Asc(Mid(sMessage, i, 1))
        bytT = (iCRC And &HFF) Xor bytC
        lngA = ulShiftRightBy8(iCRC)
        iCRC = lngA Xor aCRC32Table(bytT)
    Next
    
    CRC32 = iCRC Xor &HFFFFFFFF

End Function

Public Function ulShiftRightBy8(x As Long) As Long
    ' Shift 32-bit long value to right by 8 bits
    ' Avoiding problem with sign bit
    Dim iNew As Long
    iNew = (x And &H7FFFFFFF) \ 256
    If (x And &H80000000) <> 0 Then
        iNew = iNew Or &H800000
    End If
    ulShiftRightBy8 = iNew
End Function

Public Function CRC32Setup()

    Static bDone As Boolean

    Dim vntA As Variant
    Dim i As Integer, iOffset As Integer
    Dim nLen As Integer
    
    If bDone Then
        Exit Function
    End If

    iOffset = 0
    nLen = 32
    ' Use variant array kludge to set up table
    vntA = Array( _
        &H0, &H77073096, &HEE0E612C, &H990951BA, _
        &H76DC419, &H706AF48F, &HE963A535, &H9E6495A3, _
        &HEDB8832, &H79DCB8A4, &HE0D5E91E, &H97D2D988, _
        &H9B64C2B, &H7EB17CBD, &HE7B82D07, &H90BF1D91, _
        &H1DB71064, &H6AB020F2, &HF3B97148, &H84BE41DE, _
        &H1ADAD47D, &H6DDDE4EB, &HF4D4B551, &H83D385C7, _
        &H136C9856, &H646BA8C0, &HFD62F97A, &H8A65C9EC, _
        &H14015C4F, &H63066CD9, &HFA0F3D63, &H8D080DF5)

    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen
    
    vntA = Array( _
        &H3B6E20C8, &H4C69105E, &HD56041E4, &HA2677172, _
        &H3C03E4D1, &H4B04D447, &HD20D85FD, &HA50AB56B, _
        &H35B5A8FA, &H42B2986C, &HDBBBC9D6, &HACBCF940, _
        &H32D86CE3, &H45DF5C75, &HDCD60DCF, &HABD13D59, _
        &H26D930AC, &H51DE003A, &HC8D75180, &HBFD06116, _
        &H21B4F4B5, &H56B3C423, &HCFBA9599, &HB8BDA50F, _
        &H2802B89E, &H5F058808, &HC60CD9B2, &HB10BE924, _
        &H2F6F7C87, &H58684C11, &HC1611DAB, &HB6662D3D)

    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen
    
    vntA = Array( _
        &H76DC4190, &H1DB7106, &H98D220BC, &HEFD5102A, _
        &H71B18589, &H6B6B51F, &H9FBFE4A5, &HE8B8D433, _
        &H7807C9A2, &HF00F934, &H9609A88E, &HE10E9818, _
        &H7F6A0DBB, &H86D3D2D, &H91646C97, &HE6635C01, _
        &H6B6B51F4, &H1C6C6162, &H856530D8, &HF262004E, _
        &H6C0695ED, &H1B01A57B, &H8208F4C1, &HF50FC457, _
        &H65B0D9C6, &H12B7E950, &H8BBEB8EA, &HFCB9887C, _
        &H62DD1DDF, &H15DA2D49, &H8CD37CF3, &HFBD44C65)

    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen
    
    vntA = Array( _
        &H4DB26158, &H3AB551CE, &HA3BC0074, &HD4BB30E2, _
        &H4ADFA541, &H3DD895D7, &HA4D1C46D, &HD3D6F4FB, _
        &H4369E96A, &H346ED9FC, &HAD678846, &HDA60B8D0, _
        &H44042D73, &H33031DE5, &HAA0A4C5F, &HDD0D7CC9, _
        &H5005713C, &H270241AA, &HBE0B1010, &HC90C2086, _
        &H5768B525, &H206F85B3, &HB966D409, &HCE61E49F, _
        &H5EDEF90E, &H29D9C998, &HB0D09822, &HC7D7A8B4, _
        &H59B33D17, &H2EB40D81, &HB7BD5C3B, &HC0BA6CAD)
        
    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &HEDB88320, &H9ABFB3B6, &H3B6E20C, &H74B1D29A, _
        &HEAD54739, &H9DD277AF, &H4DB2615, &H73DC1683, _
        &HE3630B12, &H94643B84, &HD6D6A3E, &H7A6A5AA8, _
        &HE40ECF0B, &H9309FF9D, &HA00AE27, &H7D079EB1, _
        &HF00F9344, &H8708A3D2, &H1E01F268, &H6906C2FE, _
        &HF762575D, &H806567CB, &H196C3671, &H6E6B06E7, _
        &HFED41B76, &H89D32BE0, &H10DA7A5A, &H67DD4ACC, _
        &HF9B9DF6F, &H8EBEEFF9, &H17B7BE43, &H60B08ED5)

    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &HD6D6A3E8, &HA1D1937E, &H38D8C2C4, &H4FDFF252, _
        &HD1BB67F1, &HA6BC5767, &H3FB506DD, &H48B2364B, _
        &HD80D2BDA, &HAF0A1B4C, &H36034AF6, &H41047A60, _
        &HDF60EFC3, &HA867DF55, &H316E8EEF, &H4669BE79, _
        &HCB61B38C, &HBC66831A, &H256FD2A0, &H5268E236, _
        &HCC0C7795, &HBB0B4703, &H220216B9, &H5505262F, _
        &HC5BA3BBE, &HB2BD0B28, &H2BB45A92, &H5CB36A04, _
        &HC2D7FFA7, &HB5D0CF31, &H2CD99E8B, &H5BDEAE1D)

    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &H9B64C2B0, &HEC63F226, &H756AA39C, &H26D930A, _
        &H9C0906A9, &HEB0E363F, &H72076785, &H5005713, _
        &H95BF4A82, &HE2B87A14, &H7BB12BAE, &HCB61B38, _
        &H92D28E9B, &HE5D5BE0D, &H7CDCEFB7, &HBDBDF21, _
        &H86D3D2D4, &HF1D4E242, &H68DDB3F8, &H1FDA836E, _
        &H81BE16CD, &HF6B9265B, &H6FB077E1, &H18B74777, _
        &H88085AE6, &HFF0F6A70, &H66063BCA, &H11010B5C, _
        &H8F659EFF, &HF862AE69, &H616BFFD3, &H166CCF45)

    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &HA00AE278, &HD70DD2EE, &H4E048354, &H3903B3C2, _
        &HA7672661, &HD06016F7, &H4969474D, &H3E6E77DB, _
        &HAED16A4A, &HD9D65ADC, &H40DF0B66, &H37D83BF0, _
        &HA9BCAE53, &HDEBB9EC5, &H47B2CF7F, &H30B5FFE9, _
        &HBDBDF21C, &HCABAC28A, &H53B39330, &H24B4A3A6, _
        &HBAD03605, &HCDD70693, &H54DE5729, &H23D967BF, _
        &HB3667A2E, &HC4614AB8, &H5D681B02, &H2A6F2B94, _
        &HB40BBE37, &HC30C8EA1, &H5A05DF1B, &H2D02EF8D)
        
    For i = iOffset To iOffset + nLen - 1
        aCRC32Table(i) = vntA(i - iOffset)
    Next
    iOffset = iOffset + nLen
    
    bDone = True
    
End Function

Public Function TestCRC32()

' Test suite answers
'CRC32(123456789) = CBF43926
'CRC32(hello world)=D4A1185
'CRC32(Hello world)=8BD69E52
'CRC32(a) = E8B7BE43
'CRC32() = E96CCF45
    
    Dim sMessage As String
    Dim iCRC As Long
    
    Call CRC32Setup
    
    sMessage = "123456789"
    iCRC = CRC32(sMessage)
    Debug.Print "CRC32(" & sMessage & ")=" & Hex(iCRC)
    
    sMessage = "hello world"
    iCRC = CRC32(sMessage)
    Debug.Print "CRC32(" & sMessage & ")=" & Hex(iCRC)

    sMessage = "Hello world"
    iCRC = CRC32(sMessage)
    Debug.Print "CRC32(" & sMessage & ")=" & Hex(iCRC)

    sMessage = "a"
    iCRC = CRC32(sMessage)
    Debug.Print "CRC32(" & sMessage & ")=" & Hex(iCRC)

    sMessage = " "
    iCRC = CRC32(sMessage)
    Debug.Print "CRC32(" & sMessage & ")=" & Hex(iCRC)


End Function





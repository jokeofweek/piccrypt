Attribute VB_Name = "modBuffer"
' Project: PicCrypt
' File: modBuffer

' Copyright (C) 2011 by Dominic Charley-Roy

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Const CHUNK_SIZE As Long = 52428800 '50 MB

Private Declare Function ZCompress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function ZUncompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

' Basic zLib Compression integration
' Found on this thread: http://key2heaven.net/ms/forums/viewtopic.php?f=177&t=898&start=0
Public Function Compress(ByRef bData() As Byte) As Byte()
    Dim lKey As Long  'original size
    Dim sTmp As String  'string buffer
    
    Dim bRet() As Byte  'output buffer
    Dim lCSz As Long  'compressed size

    If StrPtr(bData) <> 0 Then 'if data buffer contains data
        lKey = UBound(bData) + 1 'get data size
        lCSz = lKey + (lKey * 0.01) + 12 'estimate compressed size
        ReDim bRet(lCSz - 1) 'allocate output buffer
        Call ZCompress(bRet(0), lCSz, bData(0), lKey) 'compress data (lCSz returns actual size)
        ReDim Preserve bRet(lCSz - 1) 'resize output buffer to actual size
        Erase bData 'deallocate data buffer
        ReDim bData(lCSz + 3) As Byte 'allocate data buffer
        CopyMemory bData(0), lKey, 4 'copy key to buffer
        CopyMemory bData(4), bRet(0), lCSz 'copy data to buffer
        Erase bRet 'deallocate output buffer
        bRet = bData 'copy to output buffer
        Erase bData 'deallocate data buffer
         
          Compress = bRet 'return output buffer
        Erase bRet 'deallocate output buffer
    End If
End Function

Public Function Uncompress(ByRef bData() As Byte) As Byte()
   Dim lKey As Long  'original size
   Dim sTmp As String  'string buffer

   Dim bRet() As Byte  'output buffer
   Dim lCSz As Long  'compressed size
 

   If StrPtr(bData) <> 0 Then 'if there is data
      
         lCSz = UBound(bData) - 3 'get actual data size
         CopyMemory lKey, bData(0), 4 'copy key value to key
         ReDim bRet(lCSz - 1) 'allocate output buffer
         CopyMemory bRet(0), bData(4), lCSz 'copy data to output buffer
         Erase bData 'deallocate data buffer
         bData = bRet 'copy to data buffer
         Erase bRet 'deallocate output buffer

      ReDim bRet(lKey - 1) 'allocate output buffer
      Call ZUncompress(bRet(0), lKey, bData(0), lCSz) 'decompress to output buffer

         Uncompress = bRet 'return output buffer

      Erase bRet 'deallocate return buffer
   End If
End Function
Public Sub CreateChunkFile(ByVal oldFileName As String)
Dim oldFile As Long, chunkFile As Long
Dim oldData() As Byte, newData() As Byte, lngLen(0 To 3) As Byte
Dim newSize As Long, Chunks As Integer, I As Long, Counter As Long

oldFile = FreeFile
Open oldFileName For Binary As oldFile

    chunkFile = FreeFile
    
    If LOF(oldFile) > CHUNK_SIZE Then
        Chunks = (LOF(oldFile) \ CHUNK_SIZE) + 1
    Else
        Chunks = 1
    End If
    
    Open oldFileName & ".chunk" For Binary As chunkFile
        ReDim oldData(0 To 1) As Byte
        CopyMemory oldData(0), Chunks, 2
        Put chunkFile, , oldData
        
        If Chunks = 1 Then
            ReDim oldData(0 To LOF(oldFile) - 1) As Byte
            Get oldFile, 1, oldData
            newData = Compress(oldData)
            newSize = UBound(newData)
            
            
            CopyMemory lngLen(0), newSize, 4
            Put chunkFile, 3, lngLen
            Put chunkFile, 7, newData

            Erase newData
        Else
            For I = 1 To Chunks
                If I <> Chunks Then
                    ReDim oldData(0 To CHUNK_SIZE - 1) As Byte
                Else
                    ReDim oldData(0 To (LOF(oldFile) - ((Chunks - 1) * CHUNK_SIZE)) - 1) As Byte
                End If
                Get oldFile, ((I - 1) * CHUNK_SIZE) + 1, oldData
                newData = Compress(oldData)
                newSize = UBound(newData)
                
                
                CopyMemory lngLen(0), newSize, 4
                
                
                Put chunkFile, 3 + Counter + ((I - 1) * 4), lngLen
                Put chunkFile, 3 + Counter + (I * 4), newData
                
                Counter = Counter + newSize
    
                Erase newData
                DoEvents
            Next I
        End If
        

        

    Close chunkFile
    
Close oldFile

Erase oldData

End Sub
Public Sub DeCreateChunkFile(ByVal chunkFileName As String)
Dim newFile As Long, chunkFile As Long
Dim oldData() As Byte, newData() As Byte
Dim chunkSize As Long, Chunks As Integer, I As Long
Dim lngLen() As Byte, Counter As Long

chunkFile = FreeFile
Open chunkFileName For Binary As chunkFile

    newFile = FreeFile
    Open Left(chunkFileName, Len(chunkFileName) - 6) For Binary As newFile
        
        ReDim lngLen(0 To 1) As Byte
        Get chunkFile, 1, lngLen
        CopyMemory Chunks, lngLen(0), 2
        
        ReDim lngLen(0 To 3) As Byte
        For I = 1 To Chunks
            Get chunkFile, Counter + ((I - 1) * 4) + 3, lngLen
            CopyMemory chunkSize, lngLen(0), 4
            
            ReDim oldData(0 To chunkSize) As Byte
            Get chunkFile, Counter + (I * 4) + 3, oldData
            
            newData = Uncompress(oldData)
            Put newFile, , newData
            
            Erase newData
            Erase oldData
            
            Counter = Counter + chunkSize
            DoEvents
        Next I
        
    Close newFile
    
Close chunkFile

End Sub
Public Sub AddByte(ByRef Arr() As Byte, ByVal StartLoc As Byte, ByVal nVal As Byte)

CopyMemory Arr(StartLoc), nVal, 1

End Sub
Public Sub AddInt(ByRef Arr() As Byte, ByVal StartLoc As Byte, ByVal nVal As Integer)

CopyMemory Arr(StartLoc), nVal, 2

End Sub

Public Sub AddLong(ByRef Arr() As Byte, ByVal StartLoc As Byte, ByVal nVal As Long)

CopyMemory Arr(StartLoc), nVal, 4

End Sub


Public Sub Encrypt(ByRef origBytes() As Byte, ByVal keyStr As String)
    Dim keyBytes() As Byte
    Dim I As Long, Y As Long
    
    keyBytes = StrConv(keyStr, vbFromUnicode)
    
    ' Simple XOR encrypt with every character of the key (shifting)
    For Y = 0 To UBound(keyBytes)
        For I = 0 To UBound(origBytes)
            origBytes(I) = origBytes(I) Xor keyBytes((I + Y) Mod UBound(keyBytes))
        Next I
    Next Y
    
    ' Simple XOR encrypt with the location
    For I = 0 To UBound(origBytes)
        origBytes(I) = origBytes(I) Xor (I Mod 255)
    Next I
    
    ' Bit manipulation based on the character, and then more simple XOR encryption
    For I = 0 To UBound(origBytes)
        If (origBytes(I) Xor keyBytes(I Mod UBound(keyBytes))) Mod 2 = 0 Then
            origBytes(I) = origBytes(I) Xor 4  ' Toggle third bit
            origBytes(I) = origBytes(I) Xor 32 ' Toggle sixth bit
        Else
            origBytes(I) = origBytes(I) Xor 8  ' Toggle fourth bit
            origBytes(I) = origBytes(I) Xor 64 ' Toggle seventh bit
        End If
    
        origBytes(I) = origBytes(I) Xor keyBytes((I \ UBound(keyBytes)) Mod UBound(keyBytes))
    Next I
    
    ' One final key encryption round
        For I = 0 To UBound(origBytes)
        origBytes(I) = origBytes(I) Xor keyBytes(I Mod UBound(keyBytes))
    Next I
    
End Sub

Public Sub Decrypt(ByRef origBytes() As Byte, ByVal keyStr As String)
    Dim keyBytes() As Byte
    Dim I As Long, Y As Long
    
    keyBytes = StrConv(keyStr, vbFromUnicode)
    
    ' Reverse the byte order
    For I = 0 To UBound(origBytes)
        origBytes(I) = origBytes(I) Xor keyBytes(I Mod UBound(keyBytes))
    Next I
    
    For I = 0 To UBound(origBytes)
        origBytes(I) = origBytes(I) Xor keyBytes((I \ UBound(keyBytes)) Mod UBound(keyBytes))
        If (origBytes(I) Xor keyBytes(I Mod UBound(keyBytes))) Mod 2 = 0 Then
            origBytes(I) = origBytes(I) Xor 4  ' Toggle third bit
            origBytes(I) = origBytes(I) Xor 32 ' Toggle sixth bit
        Else
            origBytes(I) = origBytes(I) Xor 8  ' Toggle fourth bit
            origBytes(I) = origBytes(I) Xor 64 ' Toggle seventh bit
        End If
    Next I
   
    For I = 0 To UBound(origBytes)
        origBytes(I) = origBytes(I) Xor (I Mod 255)
    Next I
    
    For Y = UBound(keyBytes) To 0 Step -1
        For I = 0 To UBound(origBytes)
            origBytes(I) = origBytes(I) Xor keyBytes((I + Y) Mod UBound(keyBytes))
        Next I
    Next Y
    
    
End Sub



VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PicCrypt v0.3"
   ClientHeight    =   5610
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   5505
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Progress"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   5295
      Begin VB.Shape shpProgress 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   480
         Width           =   2775
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   120
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Convert Back"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   5295
      Begin VB.CommandButton cmdDecrypt 
         Caption         =   "Choose file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtDecryptKey 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   $"frmMain.frx":1CCA
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "PicCrypt"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Convert To"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
      Begin VB.HScrollBar scrlWidth 
         Height          =   255
         Left            =   1200
         Max             =   3
         TabIndex        =   13
         Top             =   1440
         Value           =   3
         Width           =   3135
      End
      Begin VB.TextBox txtEncryptKey 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   5055
      End
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "Choose file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label lblWidth 
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Image Width :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   $"frmMain.frx":1D70
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":1E86
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project: PicCrypt
' File: frmMain

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

Public PROGRESS_WIDTH As Integer

Private Sub cmdEncrypt_Click()

    Dim F As Long, I As Long
    Dim oldFile As Integer, newFile As Integer
    Dim fileLen As Long, titleLen As Byte
    Dim Data() As Byte
    Dim tmpByte As Byte, tmpInt As Integer, tmpLong As Long
    Dim titleStream() As Byte, oldTitle As String
    
    ' Open and choose a file
    Me.CommonDialog1.DialogTitle = "Choose A File"
    Me.CommonDialog1.Filter = "All Files (*.*)|*.*"
    Me.CommonDialog1.ShowOpen
    
    ' Checks for -
    ' - Does file exist?
    ' - Did user pick a file?
    If Dir(Me.CommonDialog1.FileName) = vbNullString Or Me.CommonDialog1.FileName = vbNullString Then
        ' File does not exist
        MsgBox "Error: You have not picked a file, or the file you have picked does not exist.", vbOKOnly + vbExclamation, "PicCrypt"
        Exit Sub
    End If
    
    ' Build the original title stream
    titleStream = StrConv(Me.CommonDialog1.FileTitle, vbFromUnicode)
    titleLen = Len(Me.CommonDialog1.FileTitle)
    oldTitle = Me.CommonDialog1.FileName & ".chunk"
    
    ' Compressing progress bar
    lblProgress.Caption = "Compressing..."
    shpProgress.Width = 0
    Me.Height = 7280
    
    CreateChunkFile Me.CommonDialog1.FileName
    
    oldFile = FreeFile
    
    ' Load into a byte array, reading every byte of the file
    Open Me.CommonDialog1.FileName & ".chunk" For Binary As #oldFile

        ' Get the file length
        fileLen = LOF(oldFile) + 5 + titleLen
        
        ' Cut out the old extension and add in the .bmp
        For F = Len(Me.CommonDialog1.FileName) To 1 Step -1
            If Mid(Me.CommonDialog1.FileName, F, 1) = "." Then
                Me.CommonDialog1.FileName = Left(Me.CommonDialog1.FileName, F - 1) & ".bmp"
                Exit For
            End If
        Next F
        
        ' Let's offer the user the save path
        Me.CommonDialog1.DialogTitle = "Choose The Destination File"
        Me.CommonDialog1.Filter = "Bitmap Files (*.bmp)|*.bmp"
        Me.CommonDialog1.ShowOpen
        
        ' Progress bar4
        lblProgress.Caption = "Progress : 0%"
        shpProgress.Width = 0
        'Me.Height = 7280
    
        newFile = FreeFile
        Open Me.CommonDialog1.FileName For Output As newFile
        Close newFile
        
        ' Open the bitmap
        newFile = FreeFile
        Open Me.CommonDialog1.FileName For Binary As #newFile
    
            ' Write the bitmap description.
            ReDim Data(0 To 53) As Byte
            AddInt Data, 0, 19778   'bfType
            AddLong Data, 2, (((fileLen \ (GetWidthByID(frmMain.scrlWidth.Value) * 3)) + 1) * (GetWidthByID(frmMain.scrlWidth.Value) * 3)) + 54 'bfSize
            AddLong Data, 10, 54 'bfOffBits
            AddLong Data, 14, 40 'biSize
            AddLong Data, 18, GetWidthByID(frmMain.scrlWidth.Value) 'biWidth
            AddLong Data, 22, ((fileLen \ (GetWidthByID(frmMain.scrlWidth.Value) * 3)) + 1) 'biHeight
            AddInt Data, 26, 1 'biPlanes
            AddInt Data, 28, 24 'biBitcount
            AddLong Data, 34, (GetWidthByID(frmMain.scrlWidth.Value) * 3) * ((fileLen \ (GetWidthByID(frmMain.scrlWidth.Value) * 3)) + 1) 'biSizeImage
            Put newFile, , Data
            
            ' Prepare the data file to handle each line (3072 bytes)
            ReDim Data(0 To 3071) As Byte
            
            ' Process the first line sepeartely, since we are adding the title stream and the file length
            Get oldFile, 1, Data
            CopyMemory ByVal VarPtr(Data(5 + titleLen)), ByVal VarPtr(Data(0)), 3067 - titleLen
            CopyMemory Data(0), fileLen, 4
            CopyMemory Data(4), titleLen, 1
            CopyMemory Data(5), titleStream(0), titleLen
            If Me.txtEncryptKey.Text <> vbNullString Then
                Call Encrypt(Data, Me.txtEncryptKey.Text)
            End If
            Put newFile, , Data
            Erase titleStream
            
            If fileLen > 3072 Then
                I = (3068 - titleLen)
                For F = 1 To (fileLen \ 3072)
                    ZeroMemory Data(0), 3072
                    Get oldFile, I, Data
                    
                    If Me.txtEncryptKey.Text <> vbNullString Then
                        Call Encrypt(Data, Me.txtEncryptKey.Text)
                    End If
                    
                    Put newFile, , Data
                    I = I + 3072
                    
                    lblProgress.Caption = "Progress: " & CByte((F / (fileLen \ 3072)) * 100) & "%"
                    shpProgress.Width = ((F / (fileLen \ 3072)) * PROGRESS_WIDTH) * Screen.TwipsPerPixelX
                    DoEvents
                Next F
            End If
           
        
        Close #newFile
    Close #oldFile

    Kill oldTitle

    shpProgress.Width = PROGRESS_WIDTH * Screen.TwipsPerPixelX
    lblProgress.Caption = "Conversion completed!"

End Sub

Private Sub cmdDecrypt_Click()
    
    Dim oldFile As Long, newFile As Long
    Dim F As Long, Data() As Byte, I As Long
    Dim fileLen As Long
    Dim titleLength As Byte, titleByt() As Byte, titleString As String

    ' Open and get the file
    Me.CommonDialog1.DialogTitle = "Choose A Picture File"
    Me.CommonDialog1.Filter = "Bitmap Files (*.bmp)|*.bmp"
    Me.CommonDialog1.ShowOpen
    
    ' Checks for -
    ' - Does file exist?
    ' - Did user pick a file?
    If Dir(Me.CommonDialog1.FileName) = vbNullString Or Me.CommonDialog1.FileName = vbNullString Then
        ' File does not exist
        MsgBox "Error: You have not picked a file, or the file you have picked does not exist.", vbOKOnly + vbExclamation, "PicCrypt"
        Exit Sub
    End If
    
    ' Load the bitmap's data.
    oldFile = FreeFile
    
    ' Progress bar
    lblProgress.Caption = "Progress : 0%"
    shpProgress.Width = 0
    Me.Height = 7280
        
    ReDim Data(0 To 3071) As Byte
    
    Open Me.CommonDialog1.FileName For Binary Access Read As oldFile
    
        Get oldFile, 55, Data
        
        ' In case it is encrypted, that means file length will be encrypted
        If frmMain.txtDecryptKey.Text <> vbNullString Then
            Call Decrypt(Data, txtDecryptKey.Text)
        End If
                    
        ' File Length
        CopyMemory fileLen, Data(0), 4
        
        ' Title
        titleLength = Data(4)
        ReDim titleByt(0 To (titleLength - 1)) As Byte
        CopyMemory titleByt(0), Data(5), titleLength
        
        For F = Len(Me.CommonDialog1.FileName) To 1 Step -1
            If Mid(Me.CommonDialog1.FileName, F, 1) = "\" Then
                titleString = Left(Me.CommonDialog1.FileName, F) & StrConv(titleByt, vbUnicode)
                Exit For
            End If
        Next F
        
        ' Open and choose a file
        Me.CommonDialog1.DialogTitle = "Choose Where To Save The File"
        Me.CommonDialog1.FileName = titleString
        Me.CommonDialog1.Filter = "All Files (*.*)|*.*"
        Me.CommonDialog1.ShowOpen
        
        newFile = FreeFile
        Open titleString For Output As newFile
        Close newFile
        
        newFile = FreeFile
        Open titleString & ".chunk" For Binary Access Write As newFile
            
            ' Read the first line since it is a different length
            Dim tmpData() As Byte
            ReDim tmpData(0 To 3066 - titleLength) As Byte
            CopyMemory tmpData(0), Data(5 + titleLength), 3067 - titleLength
       
            If fileLen > 3072 Then
            
                Put newFile, , tmpData ' Put  first line since it must be a full line
                                
                For F = 1 To (fileLen \ 3072)
                    ZeroMemory Data(0), 3072
                    Get oldFile, 55 + (3072 * F), Data
                    
                    If frmMain.txtDecryptKey.Text <> vbNullString Then
                        Call Decrypt(Data, txtDecryptKey.Text)
                    End If
                    
                    If F = (fileLen \ 3072) Then
                        ReDim Preserve Data(0 To 3071 - 6 - titleLength) As Byte
                    End If
                    
                    Put newFile, , Data
                    
                    lblProgress.Caption = "Progress: " & CByte((F / (fileLen \ 3072)) * 100) & "%"
                    shpProgress.Width = ((F / (fileLen \ 3072)) * PROGRESS_WIDTH) * Screen.TwipsPerPixelX
                    DoEvents
                    
                Next F
            Else
                ' Trim first line if it isnt full
                ReDim Preserve tmpData(0 To fileLen - 6 - titleLength) As Byte
                Put newFile, , tmpData
            End If
        Close newFile
        
    Close oldFile
    lblProgress.Caption = "Decompressing..."
    DeCreateChunkFile titleString & ".chunk"
    Kill titleString & ".chunk"

    shpProgress.Width = PROGRESS_WIDTH * Screen.TwipsPerPixelX
    lblProgress.Caption = "Conversion completed!"
    Me.CommonDialog1.FileName = ""
End Sub

Private Sub Form_Load()
    PROGRESS_WIDTH = Shape1.Width / Screen.TwipsPerPixelX
    If Command <> vbNullString Then
        Me.CommonDialog1.FileName = Command
        cmdEncrypt_Click
    End If
    frmMain.lblWidth.Caption = "1024 Px"
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub scrlWidth_Change()
    frmMain.lblWidth.Caption = GetWidthByID(scrlWidth.Value) & " Px"
End Sub

Public Function GetWidthByID(ByVal ID As Byte) As Long
    Select Case ID
        Case 0
            GetWidthByID = 64
        Case 1
            GetWidthByID = 256
        Case 2
            GetWidthByID = 512
        Case 3
            GetWidthByID = 1024
    End Select
End Function

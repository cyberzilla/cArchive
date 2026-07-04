VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "cArchive v0.6.1 Demo"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Maple Mono NF"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo ErrHandler

'=========================================================================
' Demo 1: ZIP Operations
'=========================================================================
With New cArchive
    .AddFromFolder App.Path & "\input\*.*", Recursive:=True
    .CompressArchive App.Path & "\output\archive.zip"
End With

With New cArchive
    .AddFromFolder App.Path & "\input\*.*", Recursive:=True, IncludeEmptyFolders:=True, Password:="!PasswordSuperKuat123#"
    .CompressArchive App.Path & "\output\archive_password.zip"
End With

With New cArchive
    .AddFromFolder App.Path & "\input\*.*", Recursive:=True, IncludeEmptyFolders:=True, Password:="!PasswordSuperKuat123#", EncrStrength:=3
    .CompressArchive App.Path & "\output\archive_password_aes.zip"
End With

With New cArchive
    .OpenArchive App.Path & "\output\archive_password.zip"
    .Extract App.Path & "\extract", Password:="!PasswordSuperKuat123#"
End With

With New cArchive
    .OpenArchive App.Path & "\output\archive_password_aes.zip"
    .Extract App.Path & "\extract\aes", Password:="!PasswordSuperKuat123#"
End With

'=========================================================================
' Demo 2: GZ Compression - Engine Comparison
'=========================================================================
Dim sSource As String
sSource = App.Path & "\gz\dummy.log"

If Dir$(sSource) <> "" Then
    Dim lOriginal As Long
    lOriginal = FileLen(sSource)
    
    '--- Thunk (Fixed Huffman)
    With New cArchive
        .GzEngine = geThunk
        .GzCompress sSource, App.Path & "\gz\dummy_thunk.gz", acMaxCompression
    End With
    
    '--- Zlib (Dynamic Huffman)
    With New cArchive
        .GzEngine = geZlib
        .GzCompress sSource, App.Path & "\gz\dummy_zlib.gz", acMaxCompression
    End With
    
    '--- Auto (default: zlib jika ada, fallback thunk)
    With New cArchive
        .GzCompress sSource, App.Path & "\gz\dummy.log.gz", acMaxCompression
    End With
End If

MsgBox "Demo selesai!", vbInformation, "cArchive v0.6.1"
Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
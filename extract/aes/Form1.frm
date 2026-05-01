VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
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
With New cArchive
    .AddFromFolder App.Path & "\input\*.*", Recursive:=True
    .CompressArchive App.Path & "\output\archive.zip"
End With

With New cArchive
    .AddFromFolder App.Path & "\input\*.*", Recursive:=True, IncludeEmptyFolders:=True, Password:="!PasswordSuperKuat123#"
    .CompressArchive App.Path & "\output\archive_password.zip"
End With

With New cArchive
    .AddFromFolder App.Path & "\input\*.*", Recursive:=True, IncludeEmptyFolders:=True, Password:="!PasswordSuperKuat123#", EncrStrength:=3 '3=AES-256,2=AES-192,1=AES-128,0-ZipCrypto default
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

With New cArchive
    If .GzCompress(App.Path & "\gz\dummy.log", App.Path & "\gz\dummy.log.gz", acMaxCompression) Then
        MsgBox "Compressed to GZ"
    Else
        MsgBox "Failed"
    End If
    If .GzDecompress(App.Path & "\gz\dummy.log.gz", App.Path & "\gz\dummy.log") Then
        MsgBox "Extracted"
    End If
End With


End Sub

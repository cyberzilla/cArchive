# cArchive

A single, dependency-free Visual Basic 6 class for creating and extracting **ZIP** archives and compressing/decompressing **GZ** (Gzip) files.

This class is a refactored combination of two VB6 projects:
* [**ZipArchive**](https://github.com/wqweto/ZipArchive) by wqweto
* [**cGZip**](https://github.com/cyberzilla/cGZip) by cyberzilla

`cArchive` merges their functionalities into a single, cohesive class, removing code duplication and leveraging a flexible Virtual File System (VFS) to handle both file-based and in-memory operations seamlessly.

---

## ## Features

* **ZIP & GZ Support**: Full capabilities for both archive formats in one class.
* **No Dependencies**: Pure VB6 code. No need for external DLLs or OCXs.
* **Flexible I/O**: Works directly with file paths (`String`), in-memory data (`Byte Array`), and stream objects thanks to its VFS architecture.
* **Strong Encryption**: Supports **AES-256**, AES-192, AES-128, and legacy ZipCrypto 2.0 for password-protected ZIP files.
* **ZIP64 Support**: Create and extract ZIP archives larger than 4 GB.
* **Progress Events**: Monitor compression and extraction progress for long operations.

---

## ## How to Use

1.  Add the `cArchive.cls` file to your VB6 project.
2.  Instantiate the class and start using its methods.

### ### ZIP Operations

#### Creating/Extract a ZIP Archive

The process involves adding files one by one and then compressing the archive.

```vb
Private Sub CreateZip()
    ' Example: Create a simple ZIP file from folder input
    With New cArchive
    .AddFromFolder App.Path & "\input\*.*", Recursive:=True
    .CompressArchive App.Path & "\output\archive.zip"
    End With
    
    ' Example: Create a simple ZIP file from folder input with password include empty folder encrypt using ZipCrypto (by default)
    With New cArchive
        .AddFromFolder App.Path & "\input\*.*", Recursive:=True, IncludeEmptyFolders:=True, Password:="!Password123#"
        .CompressArchive App.Path & "\output\archive_password.zip"
    End With
    
    ' Example: Create a simple ZIP file from folder input with password include empty folder encrypt using Aes-256 (by default)
    With New cArchive
        .AddFromFolder App.Path & "\input\*.*", Recursive:=True, IncludeEmptyFolders:=True, Password:="!Password123#", EncrStrength:=3 '3=AES-256,2=AES-192,1=AES-128,0-ZipCrypto default
        .CompressArchive App.Path & "\output\archive_password_aes.zip"
    End With
    
    'Extract File to folder
    With New cArchive
        .OpenArchive App.Path & "\output\archive_password.zip"
        .Extract App.Path & "\extract", Password:="!Password123#"
    End With

End Sub
```

---

### ### Gzip Operations

Gzip is a stream-based format, typically used for compressing a single file or data stream. The VFS-enabled methods allow for flexible source and destination types.

#### Compressing/Decompressing File (File-to-File)

```vb
Private Sub GzFileOperation()
    With New cArchive
        'Compress File to GZ
        If .GzCompress(App.Path & "\gz\dummy.log", App.Path & "\gz\dummy.log.gz", acMaxCompression) Then
            MsgBox "Compressed to GZ"
        Else
            MsgBox "Failed"
        End If
        'Decompress GZ to Original File
        If .GzDecompress(App.Path & "\gz\dummy.log.gz", App.Path & "\gz\dummy.log") Then
            MsgBox "Extracted"
        End If
    End With
End Sub
```

---

## ## License

This project is licensed under the **MIT License**.

```text
MIT License

Copyright (c) 2025 cArchive
```

---

## ## Acknowledgements

This class would not be possible without the foundational work of the original authors. It stands on the shoulders of giants.

* **ZipArchive for VB6**: A comprehensive ZIP archive management class by **wqweto**.
    * [https://github.com/wqweto/ZipArchive](https://github.com/wqweto/ZipArchive)

* **cGZip for VB6**: A Gzip compression class adapted from ZipArchive by **cyberzilla**.
    * [https://github.com/cyberzilla/cGZip](https://github.com/cyberzilla/cGZip)
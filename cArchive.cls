VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   cArchive v0.6.1-merged (c) 2025-2026
'
'   Kelas arsip serbaguna untuk VB6: ZIP + GZ + DEFLATE.
'   Menggabungkan kemampuan cZipArchive dan cGZip dalam satu file.
'
'   Fitur Utama:
'   - Kompresi/dekompresi ZIP (termasuk ZIP64 untuk file >4GB)
'   - Kompresi/dekompresi GZ (RFC 1952, streaming per-chunk)
'   - Enkripsi ZIP: AES-128/192/256 dan ZipCrypto (WinZip-compatible)
'   - Raw DEFLATE/INFLATE untuk penggunaan low-level
'   - Embedded zlib1.dll: auto-extract ke %TEMP%, zero-dependency
'   - Tri-mode GZ: zlib (Dynamic Huffman), thunk (Fixed Huffman), libarchive
'   - GzEngine property: pilih engine manual (geAuto/geThunk/geZlib)
'   - Runtime-safe zlib via DispCallFunc (tidak crash jika DLL tidak ada)
'   - VFS abstraction: input/output bisa file path atau byte array
'   - Streaming: file besar diproses per-chunk tanpa memuat ke memori
'
'   Berdasarkan cZipArchive (c) 2017-2022 by wqweto@gmail.com
'   DEFLATE thunk dari PuTTY sshzlib oleh Simon Tatham
'   zlib1.dll dari zlib library (https://zlib.net/)
'
'   LICENSE: MIT
'
'=========================================================================
'
'   RIWAYAT VERSI:
'
'   0.6.1   05/07/2026  Runtime-safe zlib + GzEngine selector
'                       - Declare Function -> DispCallFunc (tidak crash tanpa DLL)
'                       - GzEngine property: geAuto, geThunk, geZlib
'                       - GzEngineType enum
'                       - ZlibDiagnostic() method
'                       - Fix: Kill output file sebelum compress
'   0.6.0   04/07/2026  Multi-format via archiveint.dll (Windows 11 23H2+)
'                       - 12 format: ZIP, 7z, TAR, RAR, CAB, ISO, dll.
'                       - Arc API: ArcOpen, ArcExtract, ArcCreate, ArcSave, dll.
'                       - Auto-detect format dari ekstensi
'                       - ArcAvailable, ArcVersion properties
'   0.5.0   04/07/2026  Embedded zlib1.dll (Dynamic Huffman, zero-dependency)
'                       - Auto-extract zlib1.dll ke %TEMP%
'                       - pvGzCompressZlib: streaming deflate via zlib
'                       - pvGzDecompressZlib: streaming inflate via zlib
'                       - Properties: ZlibAvailable, ZlibVersion
'                       - czStorage: default acMaxCompression (level 9)
'   0.4.0   04/07/2026  Merge cGZip, streaming GZ, RFC 1952 full flags
'   0.3.0   2025        ZIP64 support, AES encryption
'   0.2.0   2022        ZIP crypto (ZipCrypto + AES)
'   0.1.0   2018        Initial cZipArchive (wqweto)
'
'=========================================================================
'
'   API REFERENCE:
'
'   ---- Properties ----
'
'   SemVersion      As String       [R]   Versi semantik class (e.g., "0.5.0-merged")
'   ThunkBuildDate  As String       [R]   Tanggal build thunk DEFLATE engine
'   LastError       As String       [R]   Pesan error terakhir (kosong jika sukses)
'   ZlibAvailable   As Boolean      [R]   True jika zlib1.dll berhasil di-load
'   ZlibVersion     As String       [R]   Versi zlib (e.g., "1.2.3"), kosong jika tidak tersedia
'   GzEngine        As GzEngineType [RW]  Engine untuk GzCompress:
'                                         geAuto(0)=otomatis, geThunk(1)=Fixed Huffman,
'                                         geZlib(2)=Dynamic Huffman
'   FileCount       As Long         [R]   Jumlah file dalam arsip yang sedang dibuka
'   FileInfo(idx, [col]) As Variant [R]   Metadata file ke-idx (nama, ukuran, CRC32, dll.)
'                                         Kolom: ucsIdxFileName, ucsIdxSize, ucsIdxCrc32, dll.
'   CodePage        As Long         [RW]  Code page untuk konversi nama file (default: OEM)
'   Comment         As String       [R]   Komentar arsip ZIP (jika ada)
'
'   ---- GZ Methods ----
'
'   GzCompress(Source, Destination, [Level]) As Boolean
'       Kompresi data ke format GZ.
'       Source:       String (file path) atau Byte() (data di memori)
'       Destination:  String (file path) atau Byte() (output buffer)
'       Level:        acNoCompression(0), acFastestCompression(1),
'                     acNormalCompression(6), acMaxCompression(9)
'       Engine ditentukan oleh GzEngine property (default: geAuto).
'       geAuto: zlib jika tersedia, fallback thunk.
'       geThunk: selalu Fixed Huffman (embedded, zero-dependency).
'       geZlib: selalu Dynamic Huffman (zlib1.dll, kompresi 40% lebih baik).
'
'   GzDecompress(SourceGz, Destination) As Boolean
'       Dekompresi file/data GZ.
'       SourceGz:     String (file path) atau Byte() (data GZ)
'       Destination:  String (file path) atau Byte() (output buffer)
'       Mendukung semua flag GZ: FEXTRA, FNAME, FCOMMENT, FHCRC.
'
'   ---- ZIP Methods ----
'
'   AddFile(File, [Name], [Comment], [Password], [EncrStrength], [Level]) As Boolean
'       Tambahkan file ke antrian kompresi ZIP.
'       File:          String (file path) atau Byte() (data)
'       Name:          Nama file di dalam arsip (opsional, auto-detect)
'       Password:      Password untuk enkripsi (opsional)
'       EncrStrength:  1=AES-128, 2=AES-192, 3=AES-256
'       Level:         0-9 (-1 = default)
'
'   AddFromFolder(Folder, [Spec], [Recursive]) As Boolean
'       Tambahkan semua file dari folder ke antrian.
'       Folder:     Path folder sumber
'       Spec:       Filter wildcard (e.g., "*.txt")
'       Recursive:  True untuk include subfolder
'
'   CompressArchive(ArchiveFile, [Comment], [Password], [EncrStrength], [Level]) As Boolean
'       Tulis antrian file ke arsip ZIP.
'       ArchiveFile:  String (file path) atau Byte() (output buffer)
'
'   OpenArchive(ArchiveFile) As Boolean
'       Buka arsip ZIP untuk dibaca.
'       ArchiveFile:  String (file path) atau Byte() (data ZIP)
'
'   Extract(ArchiveFile, [TargetFolder], [Filter], [Password]) As Boolean
'       Ekstrak file dari arsip ZIP.
'       ArchiveFile:   String (file path) atau Byte() (data ZIP)
'       TargetFolder:  Folder tujuan (opsional)
'       Filter:        Wildcard filter (e.g., "*.txt")
'       Password:      Password dekripsi (opsional)
'
'   ---- Raw DEFLATE Methods ----
'
'   Deflate(baInput(), baOutput(), [Level]) As Boolean
'       Raw DEFLATE compress byte array.
'
'   Inflate(baBuffer(), baOutput()) As Boolean
'       Raw INFLATE decompress byte array.
'
'   DeflateBase64(baInput(), sText, [Level]) As Boolean
'       DEFLATE + encode ke Base64 string.
'
'   InflateBase64(sText, baOutput()) As Boolean
'       Decode Base64 + INFLATE decompress.
'
'   DeflateInit() As Long
'       Inisialisasi konteks DEFLATE untuk multi-block.
'       Kembalikan handle (hCtx) untuk digunakan dengan DeflateBlob.
'
'   DeflateBlob(hCtx, lInputPtr, lInputSize, lOutputPtr, lOutputSize, [Level], [Final]) As Boolean
'       Kompres satu blok data menggunakan konteks DEFLATE.
'       hCtx:        Handle dari DeflateInit()
'       lInputPtr:   Pointer ke data input
'       lInputSize:  Ukuran input (bytes)
'       lOutputPtr:  [Out] Pointer ke data output (perlu CoTaskMemFree)
'       lOutputSize: [Out] Ukuran output (bytes)
'       Final:       True jika ini blok terakhir
'
'   DeflateEnd(hCtx)
'       Lepaskan konteks DEFLATE.
'
'   ---- Utility Methods ----
'
'   CalcCrc32Array(baData()) As Long
'       Hitung CRC32 dari byte array.
'
'   CalcCrc32Ptr(lPtr, lSize, lCrc32)
'       Hitung CRC32 dari pointer memori (akumulatif).
'
'   ToUtf8Array(sText) As Byte()
'       Konversi string Unicode ke byte array UTF-8.
'
'   FromUtf8Array(baText()) As String
'       Konversi byte array UTF-8 ke string Unicode.
'
'
'   ---- Multi-Format Archive Methods (archiveint.dll) ----
'
'   ArcAvailable       As Boolean  [R]   True jika archiveint.dll tersedia
'   ArcVersion         As String   [R]   Versi libarchive (e.g., "libarchive 3.8.4")
'   ArcFileCount       As Long     [R]   Jumlah entry setelah ArcOpen
'
'   ArcOpen(ArchiveFile, [ArcType], [Password]) As Boolean
'       Buka arsip multi-format untuk dibaca.
'       ArcType: arAuto(0), arZip(1), arSevenZip(2), arTar(3),
'                arTarGz(4), arTarBz2(5), arTarXz(6), arTarZst(7),
'                arGz(8), arRar(9), arCab(10), arIso(11)
'
'   ArcFileInfo(Index) As Variant
'       Array(Name, Size, FileType) untuk entry ke-Index.
'
'   ArcExtract(ArchiveFile, TargetFolder, [Password]) As Boolean
'       Ekstrak semua file ke folder tujuan.
'
'   ArcCreate(ArchiveFile, [ArcType], [Level]) As Boolean
'       Buat arsip baru (ZIP, 7z, TAR.*, CPIO, XAR).
'
'   ArcAddFile(FilePath, [ArcName]) As Boolean
'       Tambahkan file ke arsip.
'
'   ArcAddFolder(FolderPath, [Recursive], [BaseFolder]) As Boolean
'       Tambahkan semua file dari folder.
'
'   ArcAddBytes(Name, Data()) As Boolean
'       Tambahkan data byte array.
'
'   ArcSave() As Boolean
'       Simpan dan tutup arsip.
'
'   ArcClose()
'       Tutup arsip yang sedang dibuka.
'   ---- Events ----
'
'   Progress(FileIdx, Current, Total, Cancel)
'       Dipanggil selama kompresi/dekompresi untuk reporting progress.
'
'   Error(FileIdx, Source, Description, Cancel)
'       Dipanggil saat terjadi error pada file tertentu.
'
'   BeforeCompress(FileIdx, Level, SkipFile, Cancel)
'       Dipanggil sebelum file dikompres (hanya jika ImplCompress aktif).
'
'   BeforeExtract(FileIdx, File, SkipFile, Cancel)
'       Dipanggil sebelum file diekstrak (hanya jika ImplDecompress aktif).
'
'   ExtractComplete(FileIdx, File)
'       Dipanggil setelah file berhasil diekstrak.
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cArchive"

#Const ImplCompress = ZIP_NOCOMPRESS = 0
#Const ImplDecompress = ZIP_NOEXTRACT = 0
#Const ImplCrypto = ZIP_CRYPTO <> 0
#Const ImplUseShared = ZIP_USESHARED <> 0
#Const ImplInflate = ZIP_INFLATE <> 0
#Const ImplLogging = DEBUG_MODE

'=========================================================================
' Public enums
'=========================================================================

Public Enum GzEngineType
    geAuto = 0          '--- Otomatis: zlib > thunk
    geThunk = 1         '--- Force thunk (Fixed Huffman)
    geZlib = 2          '--- Force zlib (Dynamic Huffman)
End Enum

Public Enum ArchiveCompressionLevel
    acNoCompression = 0         ' Tanpa kompresi (Store)
    acFastestCompression = 1    ' Paling cepat, kompresi minimal
    acNormalCompression = 6     ' Keseimbangan terbaik (Default)
    acMaxCompression = 9        ' Kompresi maksimal, paling lambat
End Enum

Public Enum ArchiveType
    arAuto = 0              ' Auto-detect dari ekstensi file
    arZip = 1               ' ZIP archive
    arSevenZip = 2          ' 7-Zip archive (LZMA/LZMA2)
    arTar = 3               ' TAR (ustar, tanpa kompresi)
    arTarGz = 4             ' TAR + GZIP
    arTarBz2 = 5            ' TAR + BZip2
    arTarXz = 6             ' TAR + XZ/LZMA2
    arTarZst = 7            ' TAR + Zstandard
    arGz = 8                ' GZ tunggal (built-in + zlib)
    arRar = 9               ' RAR (baca saja)
    arCab = 10              ' CAB (baca saja)
    arIso = 11              ' ISO 9660 (baca saja)
    arCpio = 12             ' CPIO
    arXar = 13              ' XAR
End Enum

Public Enum UcsFileInfoIndexEnum
    ucsIdxFileName = 0
    ucsIdxAttributes
    ucsIdxCrc32
    ucsIdxSize
    ucsIdxCompressedSize
    ucsIdxComment
    ucsIdxLastModified
    ucsIdxMethod
    ucsIdxOffset
    ucsIdxFlags
End Enum

'=========================================================================
' Public events
'=========================================================================

Event Progress(ByVal FileIdx As Long, ByVal Current As Currency, ByVal Total As Currency, Cancel As Boolean)
Event Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
#If ImplCompress Then
    Event BeforeCompress(ByVal FileIdx As Long, Level As Long, SkipFile As Boolean, Cancel As Boolean)
#End If ' ImplCompress
#If ImplDecompress Then
    Event BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    Event ExtractComplete(ByVal FileIdx As Long, File As Variant)
#End If ' ImplDecompress

'=========================================================================
' Thunk data
'=========================================================================
' Auto-generated on 12.1.2018 17:15:52, CodeSize=6640, DataSize=984, ALIGN_SIZE=16
Private Const STR_THUNK1 As String = _
    "UYtEJAhTi1wkEFWLbCQYVleLeEQD64tEJCSJbCQQiwD2wwN0HDvddBQPthNDD7bIM9HB6AgzBJf2wwN16IlcJByL1cdEJBgAAAAAK9OD4vyNDBqDwgPB6gI7yxvJ99EjyolMJCB0eYvpjaQkAAAAAIsbi8vB6QgPtvGLyMHpCA+2yTPxi8vB6RAPttGLyMHpEIu0twAIAAAPtskz0YvIwekYM7SXAAQAAIvTweoYM9EPtsgPtsOLXCQcM8iDwwQzNJeJXCQcM7SPAAwAAItMJBiLxkGJTCQYO811lItsJBCLzTP2K8s76xvt99Uj6XQWD7YTjVsBD7bIRjPRwegIMwSXO/V16otMJCRfXl2JAVtZwhAAzMzMzMzMzMzMzMzMVot0JBCF9n41i1QkDE6LRCQIwe4ERoMCAYsKdQP/QgSJCItKBIlIBMdACAAAAADHQAwAAAAAg8AQg+4BddhewhAAzMzMzMzMzMzMzMzMzMyLVCQMhdJ0GotEJAhWi3Qk" & _
    "CCvwigwGjUABMEj/g+oBdfJewhAAzMzMzMzMzMzMzMyDfCQMAItEJASLSAyJTCQED46OAAAAU1VWi3QkFFeLfCQgi++D5QGD5wKNmwAAAACKHoXtdBWLSAiDyQKL0YPyAQ+v0cHqCDLTiBaF/3QCih4PthBGD7bLi1wkFDPRiwjB6QiLFJOLWAgz0Q+2ygNIBGnJBYQICIkQD7bTwesIQYlIBA+2SAcz0YtMJBQzHJGLTCQcSYlYCIlMJByFyX+QX15dW8IQAMzMzMzMzMzMzFaLdCQIV2oMi0Yo/9CL+FeJN+gJCQAAi04oahT/0cdAEAAAAADHQAwAAAAAiUcIi8dfXsIQAMzMzMzMzMzMzMxWi3QkCFeLPv92CItHMP/Q/3YEi0cw/9CLRzBW/9BfXsIQAMzMzMzMzMzMzMzMzMyLRCQMVot0JAxXi3wkDIXAdA1Q/3YE/zb/N+gh/f///3YcjUYM/3YY/3YU/3YQUI1GCFD/dgT/NlfoIgkAAF9e" & _
    "whAAzMzMzMzMzMzMzMzMzP90JATo5w8AAMIQAMzMzMxWi3QkCItODI1GDFeLPoXJdAw7TgR0B1BW6IIRAACLThCNRhCFyXQMO04IdAdQVuhsEQAAg34UAI1GFHQHUFboXBEAAI1GBFBW6FIRAACNRghQVuhIEQAAi0cwVv/QX17CEADMzMzMzMzMzMzMzMzMi0QkCFNVi2wkDFZXjVgMU414CFf/cAT/MFXoIQkAAItMJByL8IXJdBiF9nQUiweFwHQMUf8zUP91AOgx/P//i8ZfXl1bwhAAzMzMzMzMzMxWi3QkCIPK/4uGBIAEAA+3TMYCjQTGZjvKdAkPv8FmiRTG6xCLQASD+P90CGaJlEYIgAQAi4YEgAQAi0wkEIlMxgSLhgSABABmiVTGAo0UTouOBIAEAA+3ggiABABmiQTOi4YEgAQAD78Mxg+3wGaJggiABACD+f90DGaLhgSABABmiUTOAouOBIAEAIpEJAyIhA4AAAQAi4YEgAQAQCX/" & _
    "fwAAiYYEgAQAXsIMAMzMzIPsDFNVi2wkIFZXi3wkIDP2i38EiXwkGDm3DIAGAA+O4gAAAIuPDIAGAIvBK8YDxYP4BA+MnAAAAItUJCQ78X0KD7aEPgiABgDrCIvGK8EPtgQQjV4CiEQkFI1D/zvBfQoPtoQ+CYAGAOsJi8YrwQ+2RBABiEQkFTvZfQoPtoQ+CoAGAOsJi8YrwQ+2RBACiEQkFo1DATvBfQoPtoQ+C4AGAOsJi8YrwQ+2RBADiEQkF4tEJBRpyL2nNR7B6RBRUFfoiv7//0Y7twyABgAPjFH////rMYvOO7cMgAYAfSeNlwiABgDrCY2kJAAAAACL/4qEDwiABgCNUgGIQv9BO48MgAYAfOoptwyABgCF7Q+O9QEAAOsKjaQkAAAAAI1JAItcJCSD/QQPjFkBAACLhwAABAAz7YmHAIAEAGkDvac1HsHoEA+/lEcIgAQAg/r/D4QxAQAAi48EgAQAi8Irwb4AgAAAJf9/AAAr8CvOgeH/" & _
    "fwAAi4Q5AAAEADsDdQ6JtK8QgAYARTtsJCx9CQ+/FNeD+v91wIXtD47pAAAAi0QkKDlEJDB+BolEJDDrBItEJDC+BAAAADvGfnGNmwAAAACKBB4zyTPbiEQkE4XtfliQi5SPEIAGAIvGK8J5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrC4t8JCSKBDiLfCQYOEQkE3UIiZSfEIAGAENBO818vYP7AX4Pi0QkMEaL64tcJCQ78HyZi1wkJItsJCg79X09i5cQgAYAi84ryo0sGoXJeRiLhwSABAArwgPGJf9/AACKhDgAAAQA6wOKBBk4BCl1CEZBO3QkKHzUi2wkKFb/txCABgD/dCQo6KAOAACF9n596ziLTCQgigM8j4sRi3EID7bIi0I8dwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFZS6EcSAAC+AQAAAItsJCiD/QR8FmkDvac1HsHoEFAPtgNQV+h1/P//6xWLjwyABgCKA4iEDwiABgD/hwyA" & _
    "BgBOQ02F9n/JiWwkKIlcJCSF7Q+PF/7//19eXVuDxAzCFADMzMzMzMzMzMzMzMzMg+wUU1WLbCQoVleLfCQoM/aLfwSJfCQgObcMgAYAD47iAAAAi48MgAYAi8ErxgPFg/gED4ycAAAAi1QkLDvxfQoPtoQ+CIAGAOsIi8YrwQ+2BBCNXgKIRCQYjUP/O8F9Cg+2hD4JgAYA6wmLxivBD7ZEEAGIRCQZO9l9Cg+2hD4KgAYA6wmLxivBD7ZEEAKIRCQajUMBO8F9Cg+2hD4LgAYA6wmLxivBD7ZEEAOIRCQbi0QkGGnIvac1HsHpEFFQV+hq+///Rju3DIAGAA+MUf///+sxi847twyABgB9J42XCIAGAOsJjaQkAAAAAIv/ioQPCIAGAI1SAYhC/0E7jwyABgB86im3DIAGADPAM8mJRCQUiUwkGIlEJByF7Q+OCQIAAOsKjaQkAAAAAI1JAItcJCyD/QQPjMABAACLhwAABAAz7YmHAIAEAGkDvac1" & _
    "HsHoEA+/lEcIgAQAg/r/D4SYAQAAi48EgAQAi8Irwb4AgAAAJf9/AAAr8CvOgeH/fwAAi4Q5AAAEADsDdQ6JtK8QgAYARTtsJDR9CQ+/FNeD+v91wIXtD45MAQAAi1QkMItEJDg7wn4Gi8KJRCQ4vgQAAAA7xn53igQeM8kz24hEJBOF7X5g6weNpCQAAAAAi5SPEIAGAIvGK8J5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrC4t8JCyKBDiLfCQgOEQkE3UIiZSfEIAGAENBO818vYP7AX4Pi0QkOEaL64tcJCw78HyRi1wkLItUJDA78n05i5cQgAYAi84ryo0sGoXJeRiLhwSABAArwgPGJf9/AACKhDgAAAQA6wOKBBk4BCl1CEZBO3QkMHzUi2wkFIXtfl6NRQE78H5Pi0QkKIsQi2gIi0QkHDyPD7bIi0I8dwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFVS6CAPAAAPtgOLjxCABgCJTCQYiXQk" & _
    "FIlEJBzpkAAAAItEJBhVUOsli4cQgAYAiUQkGA+2A4l0JBSJRCQc63GLTCQYi2wkFIXtfjJVUf90JDDo8AoAADPAjXX/iUQkFIX2f1KLbCQwi0wkGIXtD48D/v//X15dW4PEFMIUAItMJCiKAzyPixGLcQgPtsiLQjx3CQ+2RAgwagjrDQ+2BAhqCY0ERQEAAABQVlLodg4AAL4BAAAAi2wkMIP9BHwWaQO9pzUewegQUA+2A1BX6KT4///rFYuPDIAGAIoDiIQPCIAGAP+HDIAGAE5DTYX2f8mJbCQwiVwkLOlq////zMzMzMzMzMzMVleLfCQMaLCPBgCLB4tAKP/Qi/CF9nUFX17CBACJdwSNTgK6AIAAAIPI/+sDjUkAx0EC/////41JCGaJQfhmiUH2g+oBdemNvgiABAC5AIAAAPOrX4mWBIAEAI1CAYmWDIAGAF7CBADMzMzMzMzMzFZXi3wkDLkSAAAAvgBQ51W4kCnnVfOli0wkEIPAYF9e" & _
    "xwEAEOdViUEEx0EIoEDnVcdBDHhE51XCCADMzMzMzMxTi1wkCFZXaACAAACLO4tzCItHKP/QiQYzwDlEJCRqAw+VwMdGCACAAACDwALHRgQAAAAAUFZX6DQNAACDfCQoAP90JDD/dCQw/3QkIP90JCBTdAfoF/j//+sF6DD7//9qB2oAVlfoBQ0AAIN8JCQAdBiLThCFyXQRuAgAAAArwVBqAFZX6OYMAACLRCQciw5fiQiLRCQci04EXluJCLgBAAAAwiQAzMzMzMzMzMzMzFFTVVaLdCQUV2gAgAAAix6JXCQci0Mo/9CLbCQgiYakgQAAx4asgQAAAIAAAMeGqIEAAAAAAACF7X8Ng76cAQAAAA+OtgUAAIO+nAEAABiNvpwBAAB9KotUJBzrA41JAIXtfhkPtgJNiw9C0+AJhpgBAACNQQiJB4P4GHzjiVQkHItGGIXAdQnHRhgBAAAA66aD+AF1c4sXg/oDD4xdBQAAi4aYAQAAg8L90eiLyIkX" & _
    "wegCiYaYAQAAg+EDdR2LysdGGAoAAACD4Qcr0dPoiReJhpgBAADpXP///4P5AXUYi0YEiUYMi0YIiUYQx0YYBgAAAOk/////g/kCD4U2////iU4Y6S7///+D+AJ1bosXg/oOD4zlBAAAi46YAQAAi8GD4B/B6QUFAQEAAMdGLAAAAACJRiCLwYPgH8HpBUDHRhgDAAAAiUYki8GD4A/B6QSDwASJjpgBAACJRiiNQvKJBzPAiUZEiUZIiUZMiUZQZolGVIhGVum7/v//g/gDdW85Bw+MdQQAAItGLDtGKH0xgz8DfCyLS0CLRiyKlpgBAACA4gcPtgQBiFQwRP9GLItGLIMH/cGumAEAAAM7Rih8z4tGLDtGKA+Faf7//2oTjUZEUFboxgkAAIlGFMdGGAQAAADHRiwAAAAA6Uf+//+D+AQPheIAAACLRiSLTiADwTlGLHw/UY1GV1BW6I8JAAD/diSJRgyLRiCDwFcDxlBW6HoJAACNfhSJRhBXVui9"
Private Const STR_THUNK2 As String = _
    "BQAAxwcAAAAAx0YYBgAAAOny/f///3YUjYaYAQAAV1DoGwYAAIvIg/n/D4SeAwAAg/n+D4RiAwAAg/kQfQ+LRiyITDBX/0Ys6br9//91B7gCAAAA6w8zwIP5EQ+VwI0EhQMAAACJRjAzwIP5Eg+UwI0ExQMAAACJRjSD+RB1G4tGLIXAfhQPtkQwVolGPMdGGAUAAADpbf3//zPAx0YYBQAAAIlGPOlc/f//g/gFdVyLH4tOMDvZD4wRAwAAi4aYAQAAugEAAADT4ivZSokfI9DT6ANWNImGmAEAAIXSfhyLRiQDRiCLTiw7yH0PikY8SohEMVf/RiyF0n/ki1wkGMdGGAQAAADp+/z//4P4Bg+FnAAAAP92DI1eDFeNhpgBAABQ6BgFAACD+P8PhJ0CAACD+P4PhF0CAAA9AAEAAH0QUFboCAQAAItcJBjptvz//3VAiwPHRhgBAAAAO0YEdA1TVuhYBAAAxwMAAAAAi0YQjX4Qi1wkGDtGCA+Ehvz/" & _
    "/1dW6DgEAADHBwAAAADpdPz//4tcJBg9HgEAAA+NZfz//8dGGAcAAACJRhzpVvz//4P4B3VUi0YcLQEBAACNDECLQzSNHIgPv0sCOQ8PjPsBAACLlpgBAAC4AQAAANPgSCPCA0MEiUY4D79DAikHiksCi1wkGNPqiZaYAQAAx0YYCAAAAOn9+///g/gIdTr/dhCNhpgBAABXUOghBAAAg/j/D4SmAQAAg/j+D4RqAQAAg/geD41hAQAAx0YYCQAAAIlGHOm++///g/gJD4WPAAAAi0YcixeNDECLQziNBIgPv0gCiUQkIIlMJBA70Q+MWgEAAIuGmAEAALsBAAAAK1QkENPji0wkIEsj2ANZBIkXikkC0+iDfjgAiYaYAQAAx0YYBgAAAHQsjaQkAAAAAIuGoIEAAP9OOCvDJf9/AAAPtoQwoAEAAFBW6IECAACDfjgAddv/TjiLXCQY6Sb7//+D+Ap1MYsXg/oQD4zdAAAAi46YAQAAD7fBiUZAjULw" & _
    "wekQiQeJjpgBAADHRhgLAAAA6fD6//+D+At1P4sHg/gQD4ynAAAAi46YAQAAg8DwD7fRiQeB8v//AACLRkDB6RCJjpgBAAA7wnVQ99gbwIPgC0CJRhjprPr//4P4DA+Fo/r//4M/CHxlD7aGmAEAAFBW6NkBAACDB/jBrpgBAAAIg0ZA/w+FfPr//8dGGAEAAADpcPr//4tcJBj/tqSBAACLQzD/0ItEJCTHhqSBAAAAAAAAX17HAAAAAACLRCQgXVvHAAAAAAAzwFnCFACLTCQkuAEAAACLlqSBAABfiRGLlqiBAACLTCQkXl1biRFZwhQAzMzMzMzMzMzMzMzMzIHsIAEAAFNWi7QkLAEAAFdosIEAAItGKP/Qi9iNfCQMuAgICAi5JAAAAGggAQAAiTPzq7gJCQkJjbwkoAAAALkcAAAA86uNRCQQx4QkEAEAAAcHBwdQU8eEJBwBAAAHBwcHx4QkIAEAAAcHBwfHhCQkAQAABwcHB8eEJCgBAAAH" & _
    "BwcHx4QkLAEAAAcHBwfHhCQwAQAACAgICMeEJDQBAAAICAgI6McEAACJQwSNRCQMaiBQU8dEJBgFBQUFx0QkHAUFBQXHRCQgBQUFBcdEJCQFBQUFx0QkKAUFBQXHRCQsBQUFBcdEJDAFBQUFx0QkNAUFBQXodwQAAF+JQwiLw17HQxgAAAAAx0MUAAAAAMdDEAAAAADHQwwAAAAAx4OYAQAAAAAAAMeDnAEAAAAAAADHg6CBAAAAAAAAW4HEIAEAAMIEAMzMzMzMzMzMzMzMzFOLXCQMVot0JAyLhqCBAACLDoicMKABAACLhqCBAABAJf9/AACJhqCBAACLhqiBAAA7hqyBAAB8GgPAUP+2pIEAAImGrIEAAItBLP/QiYakgQAAi4aogQAAi46kgQAAiBwB/4aogQAAXlvCCADMzMxTi1wkCFWLbCQQiwOJRCQMhe11CF2DyP9bwggAV4t9AIX/dEtWM/Y5N3wj6wONSQCLRwSNBPCDwASDOAB0B1BT" & _
    "6Lv///9GOzd+5otEJBT/dwSLQDD/0ItEJBTHRwQAAAAAV4tAMP/Qx0UAAAAAAF5fXTPAW8IIAMzMzMzMzMzMzItEJAxTi1wkDFaLCItABIszV4t8JBCLFyPKjQTID7YIO85/Jg+2CNPqK/Fmg3gC/3Usi0AEhcB0GosIi0AEI8qNBMgPtgg7zn7aX16DyP9bwgwAX164/v///1vCDACJF4kzD79AAl9eW8IMAMzMzMyD7AiLRCQMVYsoi0AIiUQkEItEJBiJbCQEhcAPji0BAABTVlc9BAEAAH4HvgIBAADrDIvwPQIBAAB+A41w/SvGux0AAACJRCQkg8//i0U0i+iJRCQUjQQ7mSvC0fiNDEA7dI0EfQSL2OvrO3SNCH4Ei/jr4YtcJByNBECNPIUAAAAAuRcBAAAPtwQvA/2LbCQQZjvBD7/Ii0U8fwwPtoRIAP7//2oH6wcPtkQBqGoIUFNV6CkDAAAPt0cCZoXAdA0rdwSYUFZTVegTAwAAi0U4" & _
    "g87/i1wkIL8eAAAAi+iL/40EN5krwtH4jQxAO1yNBH0Ei/jr6ztcjQh+BIvw6+GLfCQcjQRAjTSFAAAAAAP1i2wkEGoFD78Oi0U8D7YEyFBXVei5AgAAD7dGAmaFwHQPmFCLwytGBFBXVeihAgAAi0QkJIXAD4/Z/v//X15bXYPECMIMAMzMzMzMzMzMzMzMU1WLbCQMVldqCItFKP/Qi1wkKIv4i8u4AQAAANPgvgEAAACLTCQsSNPmiUQkKI0E9QAAAABQi0Uo/9CJRwQzyY1G/4PK/4kHhcB4I41kJACLRwRmiVTIAotHBMYEyACLRwTHRMgEAAAAAEE7D37hM+05bCQgfmmLdCQc6wONSQAPtgQuO8N+UYtEJBiLBKiLyCNMJCg7TCQkdT6Ly9P4Iwc7B380i08EZolswQKLTwQPthQuK9ONNMEPtg47yn0CiBaLdCQcugEAAAAPtgwuK8vT4gPCOwd+zEU7bCQgfKAz9jk3fGeLVCQsi2wkHIv/" & _
    "i0cEjQzwD7YBO8J+S4PI/2aJQQKLRwSNDPAPtgErwoP4B34FuAcAAABQjQQTiBFQi8uLxtPgC0QkLFD/dCQsVf90JCz/dCQs6MP+//+LTwSLVCQsiUTxBEY7N36ji8dfXl1bwhwAzMzMzMzMzMzMzIHsBAUAADPAM9JVi6wkFAUAAIlUJASJRCQMiUQkEIlEJBSJRCQYiUQkHIlEJCCJRCQkiUQkKIlEJCyJRCQwiUQkNIlEJDiJRCQ8iUQkQIlEJERWi7QkFAUAAIXtfhmL/w+2DDD/RIwMO9F9AovRQDvFfO2JVCQIM9IzyesDjUkAiVQMUANUDBCDwQQD0oP5PHzuVzP/he1+SVONnCSUAAAAD7Y0N8cDAAAAAItUtFSNQgGJRLRUhfZ0GDPAjWQkAIvKA8CD4QHR+gvBg+4BdfCJA4u0JBwFAABHg8MEO/18wFuLRCQMX4P4CXwFuAkAAABQagBqAFX/tCQkBQAAjYQkoAAAAFCLhCQoBQAA/zDo" & _
    "jP3//15dgcQEBQAAwgwAzItEJAxWi3QkDItOENPgCUYMi0QkFAPBiUYQg/gIfDlXi3wkDItGCDlGBHwPA8BQ/zaJRgiLRyz/0IkGi1YEiw6KRgyIBAr/RgTBbgwIg0YQ+IN+EAh9zV9ewhAAVYvsiwCAQMAgoGDgEJBQ0DCwcPAIiEjIKKho6BiYWNg4uHj4BIRExCSkZOQUlFTUNLR09AyMTMwsrGzsHJxc3Dy8fPwCgkLCIqJi4hKSUtIysnLyCopKyiqqauoamlraOrp6+gaGRsYmpmbmFpZW1ja2dvYOjk7OLq5u7h6eXt4+vn7+AYFBwSGhYeERkVHRMbFx8QmJSckpqWnpGZlZ2Tm5efkFhUXFJaVl5RWVVdU1tXX1DY1NzS2tbe0dnV3dPb19/QODQ8Mjo2PjE5NT0zOzc/MLi0vLK6tr6xubW9s7u3v7B4dHxyenZ+cXl1fXN7d39w+PT88vr2/vH59f3z+/f/8BAQAAAwAAAAMAAAACAQAA" & _
    "BAAAAAQAAAADAQAABQAAAAUAAAAEAQAABgAAAAYAAAAFAQAABwAAAAcAAAAGAQAACAAAAAgAAAAHAQAACQAAAAkAAAAIAQAACgAAAAoAAAAJAQEACwAAAAwAAAAKAQEADQAAAA4AAAALAQEADwAAABAAAAAMAQEAEQAAABIAAAANAQIAEwAAABYAAAAOAQIAFwAAABoAAAAPAQIAGwAAAB4AAAAQAQIAHwAAACIAAAARAQMAIwAAACoAAAASAQMAKwAAADIAAAATAQMAMwAAADoAAAAUAQMAOwAAAEIAAAAVAQQAQwAAAFIAAAAWAQQAUwAAAGIAAAAXAQQAYwAAAHIAAAAYAQQAcwAAAIIAAAAZAQUAgwAAAKIAAAAaAQUAowAAAMIAAAAbAQUAwwAAAOIAAAAcAQUA4wAAAAEBAAAdAQAAAgEAAAIBAAAQERIACAcJBgoFCwQMAw0CDgEPAAAAAAABAAAAAQAAAAEAAAACAAAAAgAAAAIAAAADAAAA" & _
    "AwAAAAMAAAAEAAAABAAAAAQAAQAFAAAABgAAAAUAAQAHAAAACAAAAAYAAgAJAAAADAAAAAcAAgANAAAAEAAAAAgAAwARAAAAGAAAAAkAAwAZAAAAIAAAAAoABAAhAAAAMAAAAAsABAAxAAAAQAAAAAwABQBBAAAAYAAAAA0ABQBhAAAAgAAAAA4ABgCBAAAAwAAAAA8ABgDBAAAAAAEAABAABwABAQAAgAEAABEABwCBAQAAAAIAABIACAABAgAAAAMAABMACAABAwAAAAQAABQACQABBAAAAAYAABUACQABBgAAAAgAABYACgABCAAAAAwAABcACgABDAAAABAAABgACwABEAAAABgAABkACwABGAAAACAAABoADAABIAAAADAAABsADAABMAAAAEAAABwADQABQAAAAGAAAB0ADQABYAAAAIAAAA=="
Private Const STR_THUNK_OFFSETS As String = "592|656|704|784|800|912|0|288|368|416|0|0|0|6896|7264|6640|7244"
Private Const STR_THUNK_BUILDDATE As String = "12.1.2018 17:15:52"
' end of generated code

'=========================================================================
' Embedded zlib1.dll (Base64-encoded, 75264 bytes)
'=========================================================================
' Auto-generated zlib1.dll embed, Size=75264 bytes, Base64=100352 chars
Private Const STR_ZLIB_1 As String = _
    "TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAABQRQAATAEHAFd23kIAAAAAAAAAAOAADiILAQI4AP4AAAAiAQAAAgAAABAAAAAQAAAAEAEAAAC4YQAQAAAAAgAABAAAAAEAAgAEAAAAAAAAAACAAQAABAAAnhICAAMAAAAAACAAABAAAAAAEAAAEAAAAAAAABAAAAAAMAEAVgYAAABAAQBIAwAAAFABAFgQAAAAAAAAAAAAAAAAAAAAAAAAAHABAEgDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC50ZXh0AAAA" & _
    "RPwAAAAQAAAA/gAAAAQAAAAAAAAAAAAAAAAAAGAAAGAuZGF0YQAAAHAAAAAAEAEAAAIAAAACAQAAAAAAAAAAAAAAAABAAADALmJzcwAAAACQAAAAACABAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAwC5lZGF0YQAAVgYAAAAwAQAACAAAAAQBAAAAAAAAAAAAAAAAAAAAAMAuaWRhdGEAAEgDAAAAQAEAAAQAAAAMAQAAAAAAAAAAAAAAAABAAADALnJzcmMAAABYEAAAAFABAAASAAAAEAEAAAAAAAAAAAAAAAAAQAAAwC5yZWxvYwAASAMAAABwAQAABAAAACIBAAAAAAAAAAAAAAAAAAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFWJ5YPsGIl1/It1DIld+IP+AXRciXQkBItVEIlUJAiLVQiJFCTo5vMAAIPsDIP+AYnDdCyF9nUMiw0AILlhhcl1EDHbidiLXfiLdfyJ7F3CDADoWQAAAOvrjbQmAAAAAIXAddDoRwAAAOvJkI10JgDHBCSAAAAA6DT6AACjACC5" & _
    "YYXAdBrHAAAAAACjECC5YeiL9AAA6Bb0AADpdf///+j8+QAAxwAMAAAAMcDrmIn2VYnlg+wIiV38ixUAILlhhdJ0KYsdECC5YYPrBDnTcg2LA4XAdSqD6wQ503PziRQk6Kv5AAAxwKMAILlhxwQkAAAAAOiI+QAAi138iexdw5D/0IsVACC5YevMjbYAAAAAVYnlg+wYx0QkCBAguWGLVQjHRCQEACC5YYkUJOg/+QAAhcAPlcKJ7F0PtsJIw4n2VYnlg+wYx0QkCBAguWGLVQjHRCQEACC5YYkUJOgP+QAAiexdw5CQkJCQkJCQkJCQVYnlXemX9AAAkJCQkJCQkFVXVlOLTCQUi3wkHIt0JBiJy8HrEIHh//8AAIP/AQ+E/gEAAIX2uAEAAAB0QoP/D3dCT4P//3QOD7YWT0YB0QHLg///dfKB+fD/AAB2BoHp8f8AAInYvnGAB4D35onYweoPadLx/wAAKdDB4BAJyFteX13Dgf+vFQAAD4bDAAAA" & _
    "vXGAB4CB77AVAAC6WwEAAA+2BgHBAcsPtkYBAcEByw+2RgIBwQHLD7ZGAwHBAcsPtkYEAcEByw+2RgUBwQHLD7ZGBgHBAcsPtkYHAcEByw+2RggBwQHLD7ZGCQHBAcsPtkYKAcEByw+2RgsBwQHLD7ZGDAHBAcsPtkYNAcEByw+2Rg4BwQHLD7ZGD4PGEAHBActKD4V3////icj35YnYweoPadLx/wAAKdH35cHqD2nS8f8AACnTgf+vFQAAD4dC////hf8PhNAAAACD/w8Pho4AAAAPthaD7xAPtm4BAdEByw+2VgIB6QHLD7ZuAwHRAcsPtlYEAekByw+2bgUB0QHLD7ZWBgHpAcsPtm4HAdEByw+2VggB6QHLD7ZuCQHRAcsPtlYKAekByw+2bgsB0QHLD7ZWDAHpAcsPtm4NAdEByw+2Vg4B6QHLD7ZuDwHRAcsB6QHLg8YQg/8PD4dy////T4P//3QOD7YuT0YB6QHLg///dfKJyL1xgAeA9+WJ" & _
    "0MHoD2nA8f8AACnBidj35cHqD2nS8f8AACnTidjpSP7//w+2FgHRgfnw/wAAdgaB6fH/AAABy4H78P8AAHbcgevx/wAA69SJ9o28JwAAAABVV1ZTg+wci0wkPMdEJAjx/wAAi1QkOItcJDCJTCQEi3QkNMdEJAwAAAAAD7frwesQiRQk6Gf0AACJx4nBD6/9unGAB4CJ+PfiweoPadLx/wAAKdeJ8oHi//8AAMHuEI0EKgHzjajw/wAAKcuB/fH/AACNvDvx/wAAdhWNaP+B/fH/AAB2Co2oDgD//410JgCB/+L/AQB2BoHv4v8BAIH/8f8AAHYGge/x/wAAg8QcwecQCf1bieheX13DkJCQkJAxLjIuMwCNdgCNvCcAAAAAg+xci0wkYIl8JFiLfCRkiVwkUItcJGyJdCRUixeJTCQci3QkaIlcJBSJVCQgi1QkcIl0JBCNdCQQx0QkMAAAAADHRCQ0AAAAAMdEJDgAAAAAx0QkDDgAAADHRCQIkBS4" & _
    "YYlUJASJNCTos0UAAIXAdBCLXCRQi3QkVIt8JFiDxFzDx0QkBAQAAACJNCToL00AAIP4AYnDdBWJNCToQFkAAIXbidh1y7j7////68SLVCQkiReJNCToJVkAAOu0jXYAg+wci1QkLMdEJBD/////i0wkKIlUJAyLVCQkiUwkCIlUJASLVCQgiRQk6BH///+DxBzDjbYAAAAAjbwnAAAAAItUJASJ0MHoDI0MEMHqDo1ECgvDkJCQkJCQkJCQkJCQAAAAAJYwB3csYQ7uulEJmRnEbQeP9GpwNaVj6aOVZJ4yiNsOpLjceR7p1eCI2dKXK0y2Cb18sX4HLbjnkR2/kGQQtx3yILBqSHG5895BvoR91Noa6+TdbVG11PTHhdODVphsE8Coa2R6+WL97Mllik9cARTZbAZjYz0P+vUNCI3IIG47XhBpTORBYNVycWei0eQDPEfUBEv9hQ3Sa7UKpfqotTVsmLJC1sm720D5vKzjbNgydVzfRc8N1txZPdGr" & _
    "rDDZJjoA3lGAUdfIFmHQv7X0tCEjxLNWmZW6zw+lvbieuAIoCIgFX7LZDMYk6Quxh3xvLxFMaFirHWHBPS1mtpBB3HYGcdsBvCDSmCoQ1e+JhbFxH7W2BqXkv58z1LjooskHeDT5AA+OqAmWGJgO4bsNan8tPW0Il2xkkQFcY+b0UWtrYmFsHNgwZYVOAGLy7ZUGbHulARvB9AiCV8QP9cbZsGVQ6bcS6ri+i3yIufzfHd1iSS3aFfN804xlTNT7WGGyTc5RtTp0ALyj4jC71EGl30rXldg9bcTRpPv01tNq6WlD/NluNEaIZ63QuGDacy0EROUdAzNfTAqqyXwN3TxxBVCqQQInEBALvoYgDMkltWhXs4VvIAnUZrmf5GHODvneXpjJ2SkimNCwtKjXxxc9s1mBDbQuO1y9t61susAgg7jttrO/mgzitgOa0rF0OUfV6q930p0VJtsEgxbccxILY+OEO2SUPmptDahaanoLzw7knf8JkyeuAAqxngd9" & _
    "RJMP8NKjCIdo8gEe/sIGaV1XYvfLZ2WAcTZsGecGa252G9T+4CvTiVp62hDMSt1nb9+5+fnvvo5DvrcX1Y6wYOij1tZ+k9GhxMLYOFLy30/xZ7vRZ1e8pt0GtT9LNrJI2isN2EwbCq/2SgM2YHoEQcPvYN9V32eo745uMXm+aUaMs2HLGoNmvKDSbyU24mhSlXcMzANHC7u5FgIiLyYFVb47usUoC72yklq0KwRqs1yn/9fCMc/QtYue2Swdrt5bsMJkmybyY+yco2p1CpNtAqkGCZw/Ng7rhWcHchNXAAWCSr+VFHq44q4rsXs4G7YMm47Skg2+1eW379x8Id/bC9TS04ZC4tTx+LPdaG6D2h/NFr6BWya59uF3sG93R7cY5loIiHBqD//KOwZmXAsBEf+eZY9prmL40/9rYUXPbBZ44gqg7tIN11SDBE7CswM5YSZnp/cWYNBNR2lJ23duPkpq0a7cWtbZZgvfQPA72DdTrrypxZ673n/Pskfp/7Uw" & _
    "HPK9vYrCusowk7NTpqO0JAU20LqTBtfNKVfeVL9n2SMuemazuEphxAIbaF2UK28qN74LtKGODMMb3wVaje8CLQAAAABBMRsZgmI2MsNTLSsExWxkRfR3fYanWlbHlkFPCIrZyEm7wtGK6O/6y9n04wxPtaxNfq61ji2Dns8cmIdREsJKECPZU9Nw9HiSQe9hVdeuLhTmtTfXtZgcloSDBVmYG4IYqQCb2/otsJrLNqldXXfmHGxs/98/QdSeDlrNoiSEleMVn4wgRrKnYXepvqbh6PHn0PPoJIPew2Wyxdqqrl1d659GRCjMa29p/XB2rmsxOe9aKiAsCQcLbTgcEvM2Rt+yB13GcVRw7TBla/T38yq7tsIxonWRHIk0oAeQ+7yfF7qNhA553qklOO+yPP9583O+SOhqfRvFQTwq3lgFT3nwRH5i6YctT8LGHFTbAYoVlEC7Do2D6COmwtk4vw3FoDhM9Lshj6eWCs6WjRMJAMxcSDHXRYti+m7KU+F3"
Private Const STR_ZLIB_2 As String = _
    "VF27uhVsoKPWP42Ilw6WkVCY194RqczH0vrh7JPL+vVc12JyHeZ5a961VECfhE9ZWBIOFhkjFQ/acDgkm0EjPadr/WXmWuZ8JQnLV2Q40E6jrpEB4p+KGCHMpzNg/bwqr+Ekre7QP7QtgxKfbLIJhqskSMnqFVPQKUZ++2h3ZeL2eT8vt0gkNnQbCR01KhIE8rxTS7ONSFJw3mV5Me9+YP7z5ue/wv3+fJHQ1T2gy8z6NoqDuweRmnhUvLE5ZaeoS5iDOwqpmCLJ+rUJiMuuEE9d718ObPRGzT/ZbYwOwnRDElrzAiNB6sFwbMGAQXfYR9c2lwbmLY7FtQClhIQbvBqKQXFbu1pomOh3Q9nZbFoeTy0VX342DJwtGyfdHAA+EgCYuVMxg6CQYq6L0VO1khbF9N1X9O/ElKfC79WW2fbpvAeuqI0ct2veMZwq7yqF7XlryqxIcNNvG134LipG4eE23magB8V/Y1ToVCJl803l87ICpMKpG2eRhDAmoJ8p" & _
    "uK7F5Pmf3v06zPPWe/3oz7xrqYD9WrKZPgmfsn84hKuwJBws8RUHNTJGKh5zdzEHtOFwSPXQa1E2g0Z6d7JdY07X+ssP5uHSzLXM+Y2E1+BKEpavCyONtshwoJ2JQbuERl0jAwdsOBrEPxUxhQ4OKEKYT2cDqVR+wPp5VYHLYkwfxTiBXvQjmJ2nDrPclhWqGwBU5VoxT/yZYmLX2FN5zhdP4UlWfvpQlS3Xe9QczGITio0tUruWNJHoux/Q2aAG7PN+Xq3CZUdukUhsL6BTdeg2EjqpBwkjalQkCCtlPxHkeaeWpUi8j2YbkaQnKoq94LzL8qGN0Oti3v3AI+/m2b3hvBT80KcNP4OKJn6ykT+5JNBw+BXLaTtG5kJ6d/1btWtl3PRafsU3CVPudjhI97GuCbjwnxKhM8w/inL9JJMAAAAAN2rCAW7UhANZvkYC3KgJB+vCywayfI0EhRZPBbhREw6PO9EP1oWXDeHvVQxk+RoJU5PYCAotngo9R1wL" & _
    "cKMmHEfJ5B0ed6IfKR1gHqwLLxubYe0awt+rGPW1aRnI8jUS/5j3E6YmsRGRTHMQFFo8FSMw/hR6jrgWTeR6F+BGTTjXLI85jpLJO7n4Czo87kQ/C4SGPlI6wDxlUAI9WBdeNm99nDc2w9o1AakYNIS/VzGz1ZUw6mvTMt0BETOQ5Wskp4+pJf4x7yfJWy0mTE1iI3snoCIimeYgFfMkISi0eCof3rorRmD8KXEKPij0HHEtw3azLJrI9S6tojcvwI2acPfnWHGuWR5zmTPcchwlk3crT1F2cvEXdEWb1XV43Il+T7ZLfxYIDX0hYs98pHSAeZMeQnjKoAR6/crGe7AuvGyHRH5t3vo4b+mQ+m5shrVrW+x3agJSMWg1OPNpCH+vYj8VbWNmqythUcHpYNTXpmXjvWRkugMiZo1p4Gcgy9dIF6EVSU4fU0t5dZFK/GPeT8sJHE6St1pMpd2YTZiaxEav8AZH9k5ARcEkgkREMs1Bc1gPQCrmSUIdjItD" & _
    "UGjxVGcCM1U+vHVXCda3VozA+FO7qjpS4hR8UNV+vlHoOeJa31MgW4btZlmxh6RYNJHrXQP7KVxaRW9ebS+tX4AbNeG3cffg7s+x4tmlc+Ncszzma9n+5zJnuOUFDXrkOEom7w8g5O5WnqLsYfRg7eTiL+jTiO3pijar671caerwuBP9x9LR/J5sl/6pBlX/LBAa+ht62PtCxJ75da5c+EjpAPN/g8LyJj2E8BFXRvGUQQn0oyvL9fqVjffN/0/2YF142Vc3utgOifzaOeM+27z1cd6Ln7Pf0iH13eVLN9zYDGvX72ap1rbY79SBsi3VBKRi0DPOoNFqcObTXRok0hD+XsUnlJzEfiraxklAGMfMVlfC+zyVw6KC08GV6BHAqK9Ny5/Fj8rGe8nI8RELyXQHRMxDbYbNGtPAzy25As5Alq+Rd/xtkC5CK5IZKOmTnD6mlqtUZJfy6iKVxYDglPjHvJ/PrX6elhM4nKF5+p0kb7WYEwV3mUq7MZt90fOa" & _
    "MDWJjQdfS4xe4Q2OaYvPj+ydgIrb90KLgkkEibUjxoiIZJqDvw5YguawHoDR2tyBVMyThGOmUYU6GBeHDXLVhqDQ4qmXuiCozgRmqvlupKt8eOuuSxIprxKsb60lxq2sGIHxpy/rM6Z2VXWkQT+3pcQp+KDzQzqhqv18o52XvqLQc8S15xkGtL6nQLaJzYK3DNvNsjuxD7NiD0mxVWWLsGgi17tfSBW6BvZTuDGckbm0it68g+AcvdpeWr/tNJi+AAAAAGVnvLiLyAmq7q+1EleXYo8y8N433F9rJbk4153vKLTFik8IfWTgvW8BhwHXuL/WSt3YavIzd9/gVhBjWJ9XGVD6MKXoFJ8Q+nH4rELIwHvfrafHZ0MIcnUmb87NcH+tlRUYES37t6Q/ntAYhyfozxpCj3OirCDGsMlHegg+rzKgW8iOGLVnOwrQAIeyaThQLwxf7Jfi8FmFh5flPdGHhmW04DrdWk+Pzz8oM3eGEOTq43dYUg3Y7UBov1H4" & _
    "ofgr8MSfl0gqMCJaT1ee4vZvSX+TCPXHfadA1RjA/G1O0J81K7cjjcUYlp+gfyonGUf9unwgQQKSj/QQ9+hIqD1YFJtYP6gjtpAdMdP3oYlqz3YUD6jKrOEHf76EYMMG0nCgXrcXHOZZuKn0PN8VTIXnwtHggH5pDi/Le2tId8OiDw3Lx2ixcynHBGFMoLjZ9ZhvRJD/0/x+UGbuGzfaVk0nuQ4oQAW2xu+wpKOIDBwasNuBf9dnOZF40iv0H26TA/cmO2aQmoOIPy+R7ViTKVRgRLQxB/gM36hNHrrP8abs35L+ibguRmcXm1QCcCfsu0jwcd4vTMkwgPnbVedFY5ygP2v5x4PTF2g2wXIPinnLN13krlDhXED/VE4lmOj2c4iLrhbvNxb4QIIEnSc+vCQf6SFBeFWZr9fgi8qwXDM7tlntXtHlVbB+UEfVGez/bCE7YglGh9rn6TLIgo6OcNSe7Six+VGQX1bkgjoxWDqDCY+n5m4zHwjBhg1tpjq1" & _
    "pOFAvcGG/AUvKUkXSk71r/N2IjKWEZ6KeL4rmB3ZlyBLyfR4Lq5IwMAB/dKlZkFqHF6W93k5Kk+Xlp9d8vEj5QUZa01gftf1jtFi5+u23l9SjgnCN+m1etlGAGi8IbzQ6jHfiI9WYzBh+dYiBJ5qmr2mvQfYwQG/Nm60rVMJCBWaTnId/ynOpRGGe7d04ccPzdkQkqi+rCpGERk4I3algHVmxtgQAXpg/q7PcpvJc8oi8aRXR5YY76k5rf3MXhFFBu5NdmOJ8c6NJkTc6EH4ZFF5L/k0HpNB2rEmU7/WmuvpxvmzjKFFC2IO8BkHaUyhvlGbPNs2J4Q1mZKWUP4uLpm5VCb83uieEnFdjHcW4TTOLjapq0mKEUXmPwMggYO7dpHg4xP2XFv9WelJmD5V8SEGgmxEYT7Uqs6Lxs+pN344QX/WXSbDbrOJdnzW7srEb9YdWQqxoeHkHhTzgXmoS9dpyxOyDnerXKHCuTnGfgGA/qmc5ZkVJAs2oDZuURyO" & _
    "pxZmhsJx2j4s3m8sSbnTlPCBBAmV5rixe0kNox4usRtIPtJDLVlu+8P22+mmkWdRH6mwzHrODHSUYblm8QYF3gAAAAB3BzCW7g5hLJkJUboHbcQZcGr0j+ljpTWeZJWjDtuIMnncuKTg1ekel9LZiAm2TCt+sXy957gtB5C/HZEdtxBkarAg8vO5cUiEvkHeGtrUfW3d5Ov01LVRg9OFxxNsmFZka6jA/WL5eoplyewUAVxPYwZs2foPPWONCA31O24gyExpEF7VYEHkomdxcjwD5NFLBNRH0g2F/aUKtWs1taj6QrKYbNu7ydasvPlAMths40XfXHXc1g3Pq9E9WSbZMKxR3gA6yNdRgL/QYRYhtPS1VrPEI8+6lZm4vaUPKAK4nl8FiAjGDNmysQvpJC9vfIdYaEwRwWEdq7ZmLT123EGQAdtxBpjSILzv1RAqcbGFiQa2tR+fv+Sl6LjUM3gHyaIPAPk0lgmojuEOmBh/ag27CG09LZFkbJfmY1wB" & _
    "a2tR9BxsYWKFZTDY8mIATmwGle0bAaV7ggj0wfUPxFdlsNnGErfpUIu+uOr8uYh8Yt0d3xXaLUmM03zz+9RMZU2yYVg6tVHOo7wAdNS7MOJK36VBPdiV16TRxG3T1vT7Q2npajRu2fytZ4hG2mC40EQELXMzAx3lqgpMX90NfMlQBXE8JwJBqr4LEBDJDCCGV2i1JSBvhbO5ZtQJzmHkn17e+Q4p2cmYsNCYIsfXqLRZsz0XLrQNgbe9XDvAumyt7biDIJq/s7YDtuIMdLHSmurVRzmd0nevBNsmFXPcFoPjYwsSlGQ7hA1taj56alqo5A7PC5MJ/50KAK4nfQeesfAPk0SHCKPSHgHyaGkGwv73YlddgGVnyxlsNnFuawbn/tQbdonTK+AQ2npaZ91KzPm532+Ovu/5F7e+Q2CwjtXW1qPoodGTfjjYwsRP3/JS0btn8aa8V2c/tQbdSLI2S9gNK9qvChtMNgNK9kEEemDfYO/DqGffVTFuju9Gab55" & _
    "y2GzjLxmgxolb9KgUmjiNswMd5W7C0cDIgIWuVUFJi/Fuju+sr0LKCu0WpJcs2oEwtf/p7XQzzEs2Z6LW96uHZtkwrDsY/ImdWqjnAJtkwqcCQap6w42P3IHZ4UFAFcTlb9KguK4ehR7sSuuDLYbOJLSjpvl1b4NfNzvtwvb3yGG09LU8dTiQmjds/gf2oNugb4Wzfa5JltvsHfhGLdHd4gIWub/D2pwZgY7yhEBC1yPZZ7/+GKuaWFr/9MWbM9FoArieNcN0u5OBINUOQOzwqdnJmHQYBb3SWlHTT5ud9uu0WpK2dZa3EDfC2Y32DvwqbyuU967nsVHss9/MLX/6b298hzKusKKU7OTMCS0o6a60DYFzdcGk1TeVykj2We/s2Z6LsRhSrhdaBsCKm8rlLQLvjfDDI6hWgXfGy0C740AAAAAGRsxQTI2YoIrLVPDZGzFBH139EVWWqeGT0GWx8jZigjRwrtJ+u/oiuP02custU8Mta5+TZ6DLY6HmBzP"
Private Const STR_ZLIB_3 As String = _
    "SsISUVPZIxB49HDTYe9Bki6u11U3teYUHJi11wWDhJaCG5hZmwCpGLAt+tupNsua5nddXf9sbBzUQT/fzVoOnpWEJKKMnxXjp7JGIL6pd2Hx6OGm6PPQ58PegyTaxbJlXV2uqkRGn+tva8wodnD9aTkxa64gKlrvCwcJLBIcOG3fRjbzxl0Hsu1wVHH0a2Uwuyrz96IxwraJHJF1kAegNBefvPsOhI26JaneeTyy7zhz83n/auhIvkHFG31Y3io88HlPBelifkTCTy2H21QcxpQVigGNDrtApiPog7842cI4oMUNIbv0TAqWp48TjZbOXMwACUXXMUhu+mKLd+FTyrq7XVSjoGwViI0/1pGWDpfe15hQx8ypEezh+tL1+suTcmLXXGt55h1AVLXeWU+EnxYOElgPFSMZJDhw2j0jQZtl/WunfOZa5lfLCSVO0DhkAZGuoxiKn+Izp8whKrz9YK0k4a+0P9DunxKDLYYJsmzJSCSr0FMV6vt+RiniZXdo" & _
    "Lz959jYkSLcdCRt0BBIqNUtTvPJSSI2zeWXecGB+7zHn5vP+/v3Cv9XQkXzMy6A9g4o2+pqRB7uxvFR4qKdlOTuDmEsimKkKCbX6yRCuy4hf711PRvRsDm3ZP810wg6M81oSQ+pBIwLBbHDB2HdBgJc210eOLeYGpQC1xbwbhIRxQYoaaFq7W0N36JhabNnZFS1PHgw2fl8nGy2cPgAc3bmYABKggzFTi65ikJK1U9Hd9MUWxO/0V+/Cp5T22ZbVrge86bccjaicMd5rhSrvKspree3TcEis+F0bb+FGKi5m3jbhf8UHoFToVGNN82UiArLz5RupwqQwhJFnKZ+gJuTFrrj93p/51vPMOs/o/XuAqWu8mbJa/bKfCT6rhDh/LBwksDUHFfEeKkYyBzF3c0hw4bRRa9D1ekaDNmNdsnfL+tdO0uHmD/nMtczg14SNr5YSSraNIwudoHDIhLtBiQMjXUYaOGwHMRU/xCgODoVnT5hCflSpA1V5+sBMYsuB" & _
    "gTjFH5gj9F6zDqedqhWW3OVUABv8TzFa12Jimc55U9hJ4U8XUPp+VnvXLZVizBzULY2KEzSWu1Ifu+iRBqDZ0F5+8+xHZcKtbEiRbnVToC86EjboIwkHqQgkVGoRP2Urlqd55I+8SKWkkRtmvYoqJ/LLvODr0I2hwP3eYtnm7yMUvOG9DafQ/CaKgz8/kbJ+cNAkuWnLFfhC5kY7W/13etxla7XFflr07lMJN/dIOHa4Ca6xoRKf8Io/zDOTJP1yAAAAAAHCajcDhNRuAka+WQcJqNwGy8LrBI18sgVPFoUOE1G4D9E7jw2XhdYMVe/hCRr5ZAjYk1MKni0KC1xHPRwmo3Ad5MlHH6J3Hh5gHSkbLwusGu1hmxir38IZabX1EjXyyBP3mP8RsSamEHNMkRU8WhQU/jAjFriOehd65E04TUbgOY8s1zvJko46C/i5P0TuPD6GhAs8wDpSPQJQZTZeF1g3nH1vNdrDNjQYqQExV7+EMJXVszLTa+ozEQHd" & _
    "JGvlkCWpj6cn7zH+Ji1bySNiTUwioCd7IOaZIiEk8xUqeLQoK7reHyn8YEYoPgpxLXEc9CyzdsMu9ciaLzeirXCajcBxWOf3cx5ZrnLcM5l3kyUcdlFPK3QX8XJ11ZtFfonceH9Ltk99DQgWfM9iIXmAdKR4Qh6TegSgynvGyv1svC6wbX5Eh284+t5u+pDpa7WGbGp37FtoMVICafM4NWKvfwhjbRU/YSurZmDpwVFlptfUZGS942YiA7pn4GmNSNfLIEkVoRdLUx9OSpF1eU/eY/xOHAnLTFq3kk2Y3aVGxJqYRwbwr0VATvZEgiTBQc0yREAPWHNCSeYqQ4uMHVTxaFBVMwJnV3W8Pla31glT+MCMUjqqu1B8FOJRvn7VWuI56FsgU99ZZu2GWKSHsV3rkTRcKfsDXm9FWl+tL23hNRuA4Pdxt+Kxz+7jc6XZ5jyzXOf+2WvluGcy5HoNBe8mSjju5CAP7KKeVu1g9GHoL+Lk6e2I0+urNorqaVy9" & _
    "/RO48PzR0sf+l2ye/1UGqfoaECz72Hob+Z7EQvhcrnXzAOlI8sKDf/CEPSbxRlcR9AlBlPXLK6P3jZX69k//zdl4XWDYujdX2vyJDts+4znecfW837Ofi931IdLcN0vl12sM2NapZu/U79i21S2ygdBipATRoM4z0+ZwatIkGl3FXv4QxJyUJ8baKn7HGEBJwldWzMOVPPvB04KiwBHolctNr6jKj8WfyMl7xskLEfHMRAd0zYZtQ8/A0xrOArktka+WQJBt/HeSK0Iuk+koGZamPpyXZFSrlSLq8pTggMWfvMf4nn6tz5w4E5ad+nmhmLVvJJl3BRObMbtKmvPRfY2JNTCMS18Hjg3hXo/Pi2mKgJ3si0L324kESYKIxiO1g5pkiIJYDr+AHrDmgdza0YSTzFSFUaZjhxcYOobVcg2p4tCgqCC6l6pmBM6rpG75rut4fK8pEkutb6wSrK3GJafxgRimM+svpHVVdqW3P0Gg+CnEoTpD86N8/aqivped" & _
    "tcRz0LQGGee2QKe+t4LNibLN2wyzD7E7sUkPYrCLZVW71yJouhVIX7hT9ga5kZwxvN6KtL0c4IO/Wl7avpg07QAAAAC4vGdlqgnIixK1r+6PYpdXN97wMiVrX9yd1zi5xbQo730IT4pvveBk1wGHAUrWv7jyatjd4N93M1hjEFZQGVef6KUw+voQnxRCrPhx33vAyGfHp611cghDzc5vJpWtf3AtERgVP6S3+4cY0J4az+gnonOPQrDGIKwIekfJoDKvPhiOyFsKO2e1socA0C9QOGmX7F8MhVnw4j3ll4dlhofR3TrgtM+PT1p3Myg/6uQQhlJYd+NA7dgN+FG/aPAr+KFIl5/EWiIwKuKeV09/SW/2x/UIk9VAp31t/MAYNZ/QTo0jtyuflhjFJyp/oLr9RxkCQSB8EPSPkqhI6PebFFg9I6g/WDEdkLaJoffTFHbPaqzKqA++fwfhBsNghF6gcNLmHBe39Km4WUwV3zzRwueFaX6A4HvLLw7Dd0hr" & _
    "yw0PonOxaMdhBMcp2bigTERvmPX80/+Q7mZQflbaNxsOuSdNtgVAKKSw78YcDIijgduwGjln138r0niRk24f9Dsm9wODmpBmkS8/iCmTWO20RGBUDPgHMR5NqN+m8c+6/pLf7EYuuIlUmxdn7CdwAnHwSLvJTC/e2/mAMGNF51VrP6Cc04PH+cE2aBd5ig9y5F03y1zhUK5OVP9A9uiYJa6LiHMWN+8WBIJA+Lw+J50h6R8kmVV4QYvg168zXLDK7Vm2O1Xl0V5HUH6w/+wZ1WI7IWzah0YJyDLp53COjoIo7Z7UkFH5sYLkVl86WDE6p48Jgx8zbuYNhsEItTqmbb1A4aQF/IbBF0kpL6/1TkoyInbzip4Rlpgrvnggl9kdePTJS8BIri7S/QHAakFmpfeWXhxPKjl5XZ+Wl+Uj8fJNaxkF9dd+YOdi0Y5f3rbrwgmOUnq16TdoAEbZ0LwhvIjfMeowY1aPItb5YZpqngQHvaa9vwHB2K20bjYVCAlT" & _
    "HXJOmqXOKf+3e4YRD8fhdJIQ2c0qrL6oOBkRRoCldiPYxmZ1YHoBEHLPrv7Kc8mbV6TxIu8Ylkf9rTmpRRFezHZN7gbO8Ylj3EQmjWT4Qej5L3lRQZMeNFMmsdrrmta/s/nG6QtFoYwZ8A5ioUxpBzybUb6EJzbblpKZNS4u/lAmVLmZnuje/IxdcRI04RZ3qTYuzhGKSasDP+ZFu4OBIOPgkXZbXPYTSelZ/fFVPphsggYh1D5hRMaLzqp+N6nP1n9BOG7DJl18domzxMru1lkd1m/hobEK8xQe5EuoeYETy2nXq3cOsrnCoVwBfsY5nKn+gCQVmeU2oDYLjhxRboZmFqc+2nHCLG/eLJTTuUkJBIHwsbjmlaMNSXsbsS4eQ9I+SPtuWS3p2/bDUWeRpsywqR90DM56ZrlhlN4FBvG4wBW4YcONdgCNvCcAAAAAi1QkCIXSdAjrFo22AAAAADHAw422AAAAAI28JwAAAABXVlOLfCQYi1wkEItUJBSF" & _
    "//fTdCv2wgN0Jo22AAAAAA+2AkIw2A+2yMHrCIs0jcAVuGEx8090CPbCA3XjjXYAg/8fidYPhvEBAACQjXQmAIsWg8YEMdMPtssPtseLFIXAHbhhiwSNwCG4YcHrEDHCD7bDiwyFwBm4YYnYwegIMcqLDIXAFbhhidOLFjHLg8YEMdMPtssPtseLFIXAHbhhwesQiwSNwCG4YTHCD7bDiwyFwBm4YYnYwegIMcqLDIXAFbhhidOLFoPGBDHLMdMPtssPtseLFIXAHbhhwesQiwSNwCG4YTHCD7bDiwyFwBm4YYnYwegIMcqLDIXAFbhhidOLFoPGBDHLMdMPtssPtseLFIXAHbhhiwSNwCG4YcHrEDHCD7bDiwyFwBm4YYnYMcrB6AiJ04sMhcAVuGGLFoPGBDHLMdMPtssPtseLFIXAHbhhwesQiwSNwCG4YTHCD7bDiwyFwBm4YYnYwegIMcqLDIXAFbhhidOLFoPGBDHLMdMPtssPtseLFIXAHbhh" & _
    "wesQiwSNwCG4YTHCD7bDiwyFwBm4YYnYwegIMcqLDIXAFbhhidOLFoPGBDHLMdMPtssPtseLFIXAHbhhiwSNwCG4YcHrEDHCD7bDiwyFwBm4YYnYwegIMcqLDIXAFbhhidOLFjHLg8YEMdMPtssPtseLFIXAHbhhiwSNwCG4YTHCwesQg+8gD7bDiwyFwBm4YYnYwegIMcqLDIXAFbhhidMxy4P/Hw+HFP7//4P/A3ZT6w2QkJCQkJCQkJCQkJCQixaD7wSDxgQx0w+2yw+2x4sUhcAduGHB6xCLBI3AIbhhMcIPtsOLDIXAGbhhidjB6AgxyosMhcAVuGGJ0zHLg/8Dd7yF/4nydB6NtgAAAAAPtgJCMNgPtsjB6wiLNI3AFbhhMfNPdej304nYW15fw1dWU4tMJBCLfCQYi3QkFInLicrB6wiB4wD/AADB6hgB2onLgeMA/wAAweMIAdrB4RiNHBGF//fTdEb3xgMAAAB0Po20JgAAAACNvCcAAAAA"
Private Const STR_ZLIB_4 As String = _
    "D7YOidrB6hjB4whGMcqLDJXAJbhhMctPdBb3xgMAAAB13o20JgAAAACNvCcAAAAAg+4Eg/8fD4bwAQAAjXQmAIPGBIsWg8YEMdMPtssPtseLFIXAKbhhwesQiwSNwCW4YTHCD7bDiwyFwC24YYnYwegIMcqLDIXAMbhhidOLFoPGBDHLMdMPtssPtseLFIXAKbhhiwSNwCW4YcHrEDHCD7bDiwyFwC24YYnYwegIMcqLDIXAMbhhidOLFjHLg8YEMdMPtssPtseLFIXAKbhhwesQiwSNwCW4YTHCD7bDiwyFwC24YYnYwegIMcqLDIXAMbhhidOLFjHLMdMPtssPtseLFIXAKbhhiwSNwCW4YcHrEDHCD7bDiwyFwC24YYnYMcrB6AiDxgSLDIXAMbhhidOLFoPGBDHLMdMPtssPtseLFIXAKbhhiwSNwCW4YcHrEDHCD7bDiwyFwC24YYnYwegIMcqLDIXAMbhhidOLFjHLg8YEMdMPtssPtseLFIXA" & _
    "KbhhwesQiwSNwCW4YTHCD7bDiwyFwC24YYnYwegIMcqLDIXAMbhhidOLFoPGBDHLMdMPtssPtseLFIXAKbhhwesQiwSNwCW4YTHCD7bDiwyFwC24YYnYwegIMcqLDIXAMbhhidOLFjHLMdMPtssPtseLFIXAKbhhiwSNwCW4YTHCwesQg+8gD7bDiwyFwC24YYnYwegIMcqLDIXAMbhhidMxy4P/Hw+HFP7//4P/A3ZT6w2QkJCQkJCQkJCQkJCQg8YEg+8EixYx0w+2yw+2x4sUhcApuGHB6xCLBI3AJbhhMcIPtsOLDIXALbhhidjB6AgxyosMhcAxuGGJ0zHLg/8Dd7yDxgSF/3QfkI10JgAPtg6J2sHqGMHjCEYxyosMlcAluGExy0915vfTid+J2Inewe8IgecA/wAAwegYgeYA/wAAAfjB5ggB8MHjGAHYW15fw422AAAAAI2/AAAAAFVXVlOB7BwBAACLjCQ4AQAAi4QkPAEAAIusJDABAACJ" & _
    "TCQIiUQkDAnID4R0AQAAx0QkECCDuO26AQAAALgBAAAAiVSEEEAB0oP4H370jbwkkAAAADHbjXQkEIsEnjHJifKFwHQYjXQmAI28JwAAAACoAXQCMwqDwgTR6HXziQyfQ4P7H37UMduLBJ8xyYn6hcB0Go22AAAAAI28JwAAAACoAXQCMwqDwgTR6HXziQyeQ4P7H37SMduQjbQmAAAAAIsEnjHJifKFwHQSkI10JgCoAXQCMwqDwgTR6HXziQyfQ4P7H37ai1QkCIPiAYXSdBsxyYn6he2J6HQPqAF0BIsqMemDwgTR6HXxic2LXCQMi0wkCA+s2QHR+4lcJAwJy4lMJAh0dDHbkI20JgAAAACLBJ8xyYn6hcB0EpCNdCYAqAF0AjMKg8IE0eh184kMnkOD+x9+2otUJAiD4gGF0nQbMcmJ8oXtieh0D6gBdASLKjHpg8IE0eh18YnNi0QkDItMJAgPrMEB0fiJRCQMCciJTCQID4Uc////i7QkNAEA" & _
    "ADH1gcQcAQAAiehbXl9dw5CQkJCQkJCQkJCQkJCQkJAfAAAAiwAAAJCNtCYAAAAAg+wMi1QkFMdEJAj/////iVQkBItUJBCJFCTosRYAAIPEDMM8ZmQ6JWQ+AJCNdCYAg+xMMcCJdCRIi3QkUIlcJESF9ngsiXQkCI1cJBDHRCQE0z24YYkcJOhDzQAAiXQkCItUJFSJHCSJVCQE6F8WAACLXCREi3QkSIPETMONdgCD7ByJXCQQi1wkIIl0JBSLdCQohduJfCQYi3wkJHQGgHtcd3Qbi1wkELj+////i3QkFIt8JBiDxBzDjbYAAAAAi1MQhdJ0IIl0JCiLdCQUiXwkJIt8JBiJXCQgi1wkEIPEHOkpIgAAi0tIi0NAiUsMiUQkDMdEJAgAQAAAx0QkBAEAAACJDCTohMwAAD0AQAAAdAfHQzj/////x0MQAEAAAOukjbYAAAAAjbwnAAAAAIPsHIl0JBSLdCQgiVwkEIl8JBiLXgSD+wEPh5EAAACF" & _
    "23QKiz6LTkQPtgeIAeiOywAAxwAAAAAAi35AiNnHRCQEAQAAALoAQAAA0/qJfCQMiVQkCItORAHLiRwk6O/LAACFwInDdRmLTkD2QQwgdBDHRjj/////ifaNvCcAAAAAi1YEAdqLXkSJVgSD+gGJHnceiVZYjXQmAI28JwAAAACLXCQQi3QkFIt8JBiDxBzDiwaAOB90CcdGWAEAAADr4IB4AYt18YtWBIPAArv/////iQaLTjyD6gKJVgSFyXUThdIPhHkEAAD/TgSLFg+2GkKJFoXJv/////91F4tGBIXAD4T7AwAAiw5IiUYED7Y5QYkOg/sIdAzHRjj9////6Xz////3x+AAAAB17ItOPDHbhcl1EYtGBIXAD4RgAwAA/wZIiUYEQ4P7BXbl98cEAAAAdGmFybv/////dReLRgSFwA+EzwIAAIsWSIlGBA+2GkKJFoXJuv////91F4tGBIXAD4RKAgAASIlGBIsGD7YQQIkGweIIAdNLg/v/dB6F" & _
    "yXUai0YEhcAPhMQBAABIiUYEiwYPthBAiQZCddz3xwgAAAB0J4XJuv////91F4tGBIXAD4Q0AQAAix5IiUYED7YTQ4kehdJ0A0J12ffHEAAAAHQnhcm6/////3UXi0YEhcAPhKAAAACLHkiJRgQPthNDiR6F0nQDQnXZg+cCdBkx24XJdQ2LRgSFwHQf/wZIiUYEQ4P7AXbpMdKFyQ+Uwo1cUv2JXjjpVP7//+hvyQAAxwAAAAAAi1ZAx0QkCABAAADHRCQEAQAAAIlUJAyLfkSJPCTo18kAAIlGBIXAdR7HRjwBAAAAi05A9kEMIHQHx0Y4/////7kBAAAA65WLfkSLTjyJPuuF6BPJAADHAAAAAACLTkDHRCQIAEAAAMdEJAQBAAAAiUwkDItWRIkUJOh7yQAAiUYEhcB1JMdGPAEAAACLXkD2QwwgdAfHRjj/////i048uv/////pFP///4tORIkOi0486fv+///orsgAAMcAAAAAAIteQMdEJAgA" & _
    "QAAAx0QkBAEAAACJXCQMi1ZEiRQk6BbJAACJRgSFwHUkx0Y8AQAAAItOQPZBDCB0B8dGOP////+LTjy6/////+mA/v//i15Ei048iR7pZ/7//+hJyAAAxwAAAAAAi05Ax0QkCABAAADHRCQEAQAAAIlMJAyLVkSJFCToscgAAIlGBIXAdA2LTkSJDotOPOn7/f//x0Y8AQAAAIteQPZDDCB0B8dGOP////+LTjzp6/3//+jpxwAAxwAAAAAAi1ZAx0QkCABAAADHRCQEAQAAAIlUJAyLTkSJDCToUcgAAIlGBIXAdSTHRjwBAAAAi05A9kEMIHQHx0Y4/////4tOPLr/////6Wr9//+LTkSJDotOPOlR/f//6ITHAADHAAAAAACLVkDHRCQIAEAAAMdEJAQBAAAAiVQkDIteRIkcJOjsxwAAiUYEhcB1JMdGPAEAAACLTkD2QQwgdAfHRjj/////i048u//////p5fz//4teRItOPIke6cz8//+NdCYA" & _
    "6BvHAADHAAAAAACLTkDHRCQIAEAAAMdEJAQBAAAAiUwkDItWRIkUJOiDxwAAiUYEhcB1IcdGPAEAAACLVkD2QgwgdAfHRjj/////uQEAAADpUfz//4tORIkOi0486T78///oucYAAMcAAAAAAItOQMdEJAgAQAAAx0QkBAEAAACJTCQMi35EiTwk6CHHAACJRgSFwHUhx0Y8AQAAAItWQPZCDCB0B8dGOP////+//////+m8+///i35EiT7ppvv//+haxgAAxwAAAAAAi35Ax0QkCABAAADHRCQEAQAAAIl8JAyLVkSJFCTowsYAAIlGBIXAdSTHRjwBAAAAi05A9kEMIHQHx0Y4/////4tOPLv/////6Tr7//+LXkSLTjyJHuki+///jXYAjbwnAAAAAFa4/v///zH2U4PsBItcJBCF23RXi0NQhcAPhcEAAACLUxyF0nQUD7ZDXDx3D4SjAAAAPHIPhIwAAACLQ0CFwHVji0M4hcB4WItDRIXAdUGL" & _
    "Q0iFwHUqi0NUhcB1E4XbdAiJHCTod8UAAInwWVtew5CJBCToaMUAAOvjjbYAAAAAiQQk6FjFAADrzI22AAAAAIkEJOhIxQAA67WNtgAAAACJxuukiQQk6MTFAACFwHSR6DvFAACDOB0PlMEPttFKCdbpe////4kcJOhSkAAAicbpZf///4kcJOjzKAAA6++QiQQk6PjEAADpMv///412AFdWU4PsEItcJCCLdCQkhdt0BoB7XHJ0DLj+////g8QQW15fw4tTOIP6/Q+EHwEAAIP6/w+EFgEAADHASnTfiXMMi1QkKIn3iVMQhdJ0QoN7eP90PA+2S3iNfgGIDotTfIn+/0MM/0sQg0NwAcdDeP////+DU3QAhdJ0FcdDOAEAAAC4AQAAAOuUjbQmAAAAAItTEIXSdHyJ9o28JwAAAACLQ1iFwA+FlQEAAItLBIXJdQuLUzyF0g+EIwEAAItTBDHJAVNoi1MQEUtsMckBU3ARS3THRCQEAAAAAIkcJOit" & _
    "dgAAiUM4i1MEMckpU2iLUxAZS2wxySlTcBlLdIP4AXRYhcB1DotLPIXJdQeLUxCF0nWNi1MMiXQkBCnyiVQkCIt7TIk8JOgF7///iUNMi0MQOUQkKHQNi1wkKCnDidjp2f7//4tTOIP6/XQDQnXouP/////pxP7//4tLDIl0JAQp8YlMJAiLc0yJNCTov+7//4lDTItzDIkcJOjhCAAAO0NMdBHHQzj9////uP3////pbP///4kcJOjDCAAAiRwk6Jv3//+LUziF0g+FY////4kcJOgIaQAAx0QkCAAAAADHRCQEAAAAAMcEJAAAAADoXO7//4lDTItDOOki////kOgbwwAAxwAAAAAAi1NAx0QkCABAAADHRCQEAQAAAIlUJAyLS0SJDCTog8MAAIlDBIXAdRDHQzwBAAAAi0tA9kEMIHUKi1NEiRPpj/7//8dDOP/////p2f7//422AAAAAItzBDnWdgKJ1oX2dVeF0nUuKVQkKDHJi1QkKAFTaIt8"
Private Const STR_ZLIB_5 As String = _
    "JCgRS2wBU3ARS3SF/3UHx0M8AQAAAItEJCjpnv3//4tzQIlUJAjHRCQEAQAAAIl0JAyJPCTo9sIAAItTECnCiVMQ662LSwwB94sTiXQkCIkMJIlUJATotcIAAAEzi0sQKXMEiXsMKfGJyolLEOl7////jbQmAAAAAI28JwAAAACD7ByLTCQgjVQkG4lUJATHRCQIAQAAAIkMJOgB/f//D7ZMJBtID5TAg8QcD7bQSgnRicjDjbYAAAAAjb8AAAAAi1QkCIXSdAaAelxydAa4/////8ODfCQE/3Tzg3p4/3Xtg0Jw/4tMJASDUnT/iUp4McmDejgBD5TBhcmJSnx0B8dCOAAAAADHQjwAAAAAi0QkBMONtgAAAACNvCcAAAAAVVdWU4PsDItcJCSLbCQgi3QkKIXbid90RYX2fkGNdgBOhfZ+IMdEJAgBAAAAiVwkBIksJOhH/P//SHUJD7YTQ4D6CnXbxgMAOd90Con4g8QMW15fXcMxwIX2f/Lr7jHA" & _
    "6+yNtCYAAAAAjbwnAAAAAFdWU4PsEItcJCCLfCQki3QkKIXbdAaAe1x3dAy4/v///4PEEFteX8OJO4X2iXMED4STAAAAjXQmAI28JwAAAACLUxCF0nUzi0tIi0NAiUsMiUQkDMdEJAgAQAAAx0QkBAEAAACJDCToRMEAAD0AQAAAdXfHQxAAQAAAi1MEMckBU2iLUxARS2wxyQFTcBFLdMdEJAQAAAAAiRwk6DAYAACJQziLUwQxySlTaItTEBlLbDHJKVNwGUt0hcB1C4tDBIXAD4V4////iXQkCIl8JASLe0yJPCToZev//4lDTInwi3MEKfDpN////8dDOP/////r1Y22AAAAAI28JwAAAAC4HBAAAOiWvwAAiZwkGBAAAIuUJCQQAACNjCQoEAAAiUwkDI1cJBDGhCQPEAAAAIlUJAjHRCQEABAAAIkcJOjNwAAAjUj/gfn+DwAAdwqAvCQPEAAAAHQQMcCLnCQYEAAAgcQcEAAAw4lcJASLnCQg" & _
    "EAAAiUQkCIkcJOiB/v//69nrDZCQkJCQkJCQkJCQkJCD7ByLTCQkjVQkG4lUJATHRCQIAQAAAIhMJBuLTCQgiQwk6En+//8PtkwkG0gPlMCDxBwPttBKCdGJyMONdCYAg+wciVwkGItcJCSJHCToDcAAAIlEJAiLVCQgiVwkBIkUJOgJ/v//i1wkGIPEHMOQg+wMi1QkFIlcJAiLXCQQiVQkBIkcJOh1DAAAhcB0CItcJAiDxAzDi1NAiRQk6L6+AACLSzgxwIP5AQ+UwEghyOvdjbQmAAAAAI28JwAAAABXVlOD7BCLfCQgi1wkJIt0JCiF/4tUJCwPhLICAACD+gIPhKkCAACLRziD+P8PhJ0CAACD+P0PhJQCAACAf1x3D4SqAQAASg+EmAEAAIX2uv////+5/////w+IwQAAAItXWIXSD4UmAQAAi1d0i0dwOdYPjPgAAAB/CDnDD4LuAAAAKcMZ1onxCdl0C4tXSIXSD4S0AAAAifEJ2XQpg394" & _
    "/3Qjg0dwAYtXfMdHeP////+DV3QAg8P/g9b/hdJ0B8dHOAEAAACF9n52jbYAAAAAhfa4AEAAAH5XiUQkCItPSIk8JIlMJATo1Pj//4XAuv////+5/////34mmSnDGdaF9n/NhfZ4BYP7AHfEi1dwi090jbQmAAAAAI28JwAAAACDxBCJ0InKW15fw5CNdCYAhfZ4CIH7/z8AAHedidjrmYX2eMiD+wB3h+vBxwQkAEAAAOh5vQAAiUdIhcC6/////7n/////D4Ur////67KJ9ok8JOhoAQAAhcC6/////7n/////D4n8/v//65SNdCYAx0d4/////4tXRMdHBAAAAACJF8dEJAwAAAAAiVwkBIl0JAiLT0CJDCToorUAAIXAuv////+5/////w+IUP///4lfcInaifGJd3SJX2iJd2zpO////wNfcBN3dOld/v//hdJ1BitfaBt3bIX2uv////+5/////w+IFP///4tXRIXSdHiF9n5pifaNvCcAAAAA" & _
    "hfa4AEAAAH5HiUQkCItXRIk8JIlUJATodPv//4XAuv////+5/////w+E0v7//zHJKcMZzoX2f8iF9ngFg/sAd7+LV2iLT2zptP7//410JgCF9ngIgfv/PwAAd62J2OuphfZ43YP7AHeX69bHBCQAQAAA6Dm8AACJR0SFwLr/////uf////8PhHT+///HRCQIAEAAAMdEJAQAAAAAiQQk6Fy8AADpSv///420JgAAAAC6/////7n/////6UH+//+Qg+wciVwkGItcJCCF23QGgHtccnQNuP////+LXCQYg8Qcw8dDOAAAAACLU0THQzwAAAAAx0N4/////4kTx0MEAAAAAMdEJAgAAAAAx0QkBAAAAADHBCQAAAAA6LHm//+JQ0yLU1iF0nRCx0NoAAAAAMdDbAAAAADHQ3AAAAAAx0N0AAAAAMdEJAwAAAAAi0tki1NgiUwkCIlUJASLS0CJDCTo2rMAAOlv////iRwk6O1gAADrtI10JgCNvCcAAAAA" & _
    "g+wci1QkIMdEJAwBAAAAx0QkBAAAAADHRCQIAAAAAIkUJOg5/P//g8Qcw5CNdCYAi0QkBIXAdAaAeFxydAMxwMOLSDyFyXQIuAEAAADDifaDeDgBD5TCD7bCw5CNdCYAi0QkBIXAdAaAeFxydAMxwMOLQFjDjXQmAI28JwAAAABWU4PsFItcJCCLQzyFwA+FHAIAAItDBIXAD4TBAQAAixNIiUMED7YyQokTi1M8uf////+F0nUXi0MEhcAPhDMBAABIiUMEiwMPtghAiQPB4QgBzoXSuf////91F4tDBIXAD4SpAAAASIlDBIsDD7YIQIkDweEQAc6F0rn/////dROLQwSFwHQoixNIiUMED7YKQokTg/n/dA2DxBTB4RgBzluJ8F7Dx0M4/f///+vqkOjruQAAxwAAAAAAi0tAx0QkCABAAADHRCQEAQAAAIlMJAyLU0SJFCToU7oAAIlDBIXAdR7HQzwBAAAAi1NA9kIMIHQHx0M4/////7n/////" & _
    "65KLS0SJC+l8////jXQmAOiLuQAAxwAAAAAAi0tAx0QkCABAAADHRCQEAQAAAIlMJAyLU0SJFCTo87kAAIlDBIXAdSTHQzwBAAAAi1NA9kIMIHQHx0M4/////4tTPLn/////6Qv///+LS0SLUzyJC+ny/v//6Ca5AADHAAAAAACLU0DHRCQIAEAAAMdEJAQBAAAAiVQkDItLRIkMJOiOuQAAiUMEhcB1JMdDPAEAAACLS0D2QQwgdAfHQzj/////i1M8uf/////pgf7//4tLRItTPIkL6Wj+//+NtgAAAADou7gAAMcAAAAAAItLQMdEJAgAQAAAx0QkBAEAAACJTCQMi1NEiRQk6CO5AACJQwSFwHUmx0M8AQAAAItzQPZGDCB0DMdDOP////+QjXQmAL7/////6fH9//+LS0SJC+nb/f//jbYAAAAAjb8AAAAAg+wciVwkDItcJCCJdCQQiXwkFIXbiWwkGA+EkAAAAIB7XHd0HIlcJCCLdCQQi1wk" & _
    "DIt8JBSLbCQYg8Qc6R/y///HRCQEBAAAAIkcJOh/BQAAhcB10ItrQL8DAAAAi3NMiWwkBInygeL/AAAAiRQkwe4I6Cm4AABPeeaLa0C/AwAAAItzaI10JgCNvCcAAAAAiWwkBInygeL/AAAAiRQkwe4I6Pm3AABPeebpdv///4tcJAy4/v///4t0JBCLfCQUi2wkGIPEHMMAOiAAjXQmAIPsHItEJCSJdCQUi3QkIIlcJBCJfCQYhfYPhOEAAACLVjiF0okQuBhTuGEPhI8AAABCD4SsAAAAi34Yhf90BYA/AHURi0Y4ugIAAAApwos8laCvuGGLRlCFwHV7i15UiRwk6Pm3AACJPCSJw+jvtwAAjUwYA4kMJOgTtwAAiUZQhcB0TItOVIkEJIlMJATovbcAAMdEJAQZU7hhi1ZQiRQk6Aq3AACJfCQEi35QiTwk6Pu2AACLRlCQjbQmAAAAAItcJBCLdCQUi3wkGIPEHMOhuK+4YevpiQQk6JG2AADp" & _
    "eP///+iXtgAAixCJFCTozbYAAInH6UH///+NtgAAAADHAP7///+hsK+4YeuzjXYAi0QkBIXAdCCDeDgBdAfHQDgAAAAAx0A8AAAAAItQQIlUJATpaLYAAMMxLjIuMwAlYyVjJWMlYyVjJWMlYyVjJWMlYwCNtgAAAACNvwAAAABVMe1XVlOB7JwAAACLlCSwAAAAx0QkPP////+LnCS0AAAAjXwkQIXSD4SYAgAAhdsPhJACAADHBCSAAAAA6O61AACJxjHAhfYPhPIBAADHRiAAAAAAx0YkAAAAAMdGKAAAAADHRkQAAAAAxwYAAAAAx0ZIAAAAAMdGDAAAAADHRhAAAAAAx0YEAAAAAMdGQAAAAADHRjgAAAAAx0Y8AAAAAMdGaAAAAADHRmwAAAAAx0ZwAAAAAMdGdAAAAADHRnj/////x0QkCAAAAADHRCQEAAAAAMcEJAAAAADocOD//4lGTIuUJLAAAADHRlAAAAAAx0ZYAAAAAIkUJOgAtgAA" & _
    "QIkEJOgntQAAiUZUhcAPhKoBAACJBCSLjCSwAAAAiUwkBOjJtQAAxkZcAJCNdCYAD7YTgPpyD4RUAgAAgPp3D4Q7AgAAgPphD4QyAgAAiNGA6TCA+QkPh+gBAAAPvsqD6TCJTCQ8Q4TSdAuNlCSQAAAAOdd1ug+2RlyEwA+EPAEAADx3D4RfAQAAxwQkAEAAAOiatAAAiUZEiQbHRCQMOAAAAMdEJAhZVLhhx0QkBPH///+JNCToJVsAAIXAD4X7AAAAi15EhdsPhPAAAADHRhAAQAAA6Ea0AADHAAAAAACLjCS4AAAAhckPiN4AAACLhCS4AAAAjWwkQIlsJASJBCTo2rMAAIlGQIXAD4StAAAAgH5cd3Q0iTQkMf/oP+j//4tOQIkMJOgkrQAAi14EKdgZ+olGYIlWZInwkI10JgCBxJwAAABbXl9dw8dEJCwLAAAAx0QkKAAAAADHRCQkAAAAAMdEJCAAAAAAx0QkHAAAAADHRCQYAAAAAMdEJBQA"
Private Const STR_ZLIB_6 As String = _
    "AAAAx0QkEAgAAADHRCQMiwAAAMdEJAgfAAAAx0QkBF9UuGGJBCToNbQAAMdGYAoAAADHRmQAAAAA6Xv///+JNCToau3//zHA6XP///+LvCSwAAAAjVQkQIlUJASJPCToLLQAAOkd////x0QkHDgAAACLRCQ8x0QkGFlUuGGJbCQUx0QkEAgAAADHRCQM8f///8dEJAgIAAAAiUQkBIk0JOh7AwAAxwQkAEAAAInH6P2yAACJRkiF/4lGDHWBhcDphP7//4D6ZnQngPpodBeA+lJ0CIgXR+kL/v//vQMAAADpAf7//70CAAAA6ff9//+QvQEAAADp7P3//422AAAAAMZGXHfpxf3//420JgAAAADGRlxy6aP9//+NtCYAAAAAVVcx/1ZTg+wci1wkMItsJDSF23QGgHtcd3QNuP7///+DxBxbXl9dw8dDBAAAAACLUxCNtCYAAAAAjbwnAAAAALkAQAAAic4p1nV1hf91aot7EDHJAXtwEUt0Mf+JbCQE" & _
    "iRwk6NgJAACJQziLUxApU3AZe3SF9nUFg/j7dDOLcxAx/4X2ifJ1H4tDOIP4AXQQg/gBdquD+AF1hjHA64KJ9r8BAAAA6+mLQzjr9I10JgDHQzgAAAAA68SLQzjr14n2i0tAiXQkCMdEJAQBAAAAiUwkDItTSIkUJOhCsgAAOfB1EsdDEABAAACLU0iJUwzpV////8dDOP////+4/////+kb////kJCQkJCQkJCQkJAxLjIuMwCNdgCNvCcAAAAAg+xci1QkYIl8JFiLfCRkiVwkUItMJGyJdCRUi1wkaI10JBCJVCQcixeJXCQQiUwkFIlUJCDHRCQwAAAAAMdEJDQAAAAAx0QkCDgAAADHRCQEMFm4YYk0JOjDWAAAhcB0EItcJFCLdCRUi3wkWIPEXMPHRCQEBAAAAIk0JOhvYwAAg/gBicN0JYk0JOgAfAAAg/sCdBGD+/t0BInY68OLdCQUhfZ19Lj9////67SLVCQkiReJNCTo1XsAAOukkJCQ" & _
    "IGRlZmxhdGUgMS4yLjMgQ29weXJpZ2h0IDE5OTUtMjAwNSBKZWFuLWxvdXAgR2FpbGx5IACNdCYAjbwnAAAAAAAAAAAAAAAAYHK4YQQABAAIAAQA0HS4YQQABQAQAAgA0HS4YQQABgAgACAA0HS4YQQABAAQABAA4Hm4YQgAEAAgACAA4Hm4YQgAEACAAIAA4Hm4YQgAIACAAAAB4Hm4YSAAgAACAQAE4Hm4YSAAAgECAQAQ4Hm4YZCNtCYAAAAAg+wsi1QkPMdEJBQAAAAAi0wkOMdEJBAIAAAAiVQkHItUJDSJTCQYx0QkDA8AAACJVCQEi1QkMMdEJAgIAAAAiRQk6BkAAACDxCzDMS4yLjMA6w2QkJCQkJCQkJCQkJCQg+wsi0QkRIlcJByLVCQ0iXQkIItMJDi7AQAAAIlEJBSLdCQ8i0QkSIl8JCSLfCQwhcCJbCQoi2wkQIlUJBiLVCRMdAWAODF0Gbj6////i1wkHIt0JCCLfCQki2wkKIPE" & _
    "LMOD+jh14oX/uP7///903sdHGAAAAACLVyCF0nUOx0cgELC4YcdHKAAAAACLVySF0nUHx0ckMLC4YYN8JBj/D4QZAgAAhfYPiAgCAACD/g9+CLsCAAAAg+4QjVX/g/oIdwWD+Qh0B7j+////64CD/gd+9IP+D3/vi0wkGIXJeOeDfCQYCX/gi0QkFIXAeNiDfCQUBH/Rg/4ID4SsAQAAx0QkCMAWAADHRCQEAQAAAItXKIkUJP9XIIlEJBCLTCQQuPz///+FyQ+EIf///4tEJBCJ8YlYGLsBAAAAidqJcDDT4ot0JBCJRxyNTQeJOMdAHAAAAACJ2NPgiVYsSolWNInCSolOUI1NCYlGTLirqqqqiVZU9+GLRCQQ0eqJUFjHRCQIAgAAAItwLIl0JASLVyiJFCT/VyCLTCQQiUE4x0QkCAIAAACLcSyJdCQEi08oiQwk/1cgi1QkEIlCQMdEJAgCAAAAi3JMiXQkBItXKIkUJP9XIIt0JBCNTQbT44me" & _
    "nBYAAIlGRMdEJAgEAAAAi46cFgAAiUwkBItfKIkcJP9XIItcJBCLk5wWAACJQwiLcziNLJUAAAAAiWsMhfZ0ZYtrQIXtdF6Lc0SF9nRXhcB0U4nVjRxSg+X+jXQFAItUJBiLbCQQibWkFgAAjTQDi0QkFIm1mBYAAImVhAAAAImFiAAAAMZFJAiLXCQciXwkMIt0JCCLfCQki2wkKIPELOmrAQAAi0wkEIsduK+4YcdBBJoCAACJXxiJPCTovxAAALj8////6aP9//++CQAAAOlK/v//Mdv33un8/f//x0QkGAYAAADp2v3//5CNdCYAVVdWU4PsLIt8JECLdCRIi2wkRIX/iXQkKA+EMQEAAItfHIXbD4QmAQAAhe0PhB4BAACLQxiD+AIPhBIBAABID4T+AAAAi0MYhcAPhdgAAAAxwIP+Ag+GxQAAAItDLC0GAQAAOcZ2CIlEJCgpxgH1i0s4iWwkBIt0JCiJDCSJdCQI6JqsAACLeziLbCQoi0NY" & _
    "iXwkEItTVIlrbIjBiWtcD7Y3iUQkJIlzSNPmi0s0D7ZHAYt8JCiJVCQgiUwkGDHGg+8DiXwkHItDQCHWiXNIi3tEMfaJRCQUkI20JgAAAAAPtkwkJItDSItUJBCLbCQg0+APtkwWAjHIIegPtxRHiUNIi0wkGItsJBQh8WaJVE0AZok0R0Y7dCQcdsQxwIn2g8QsW15fXcOJdCQIiWwkBItXMIkUJOhVsv//iUcw6Q3///+DewQqD4T4/v//jXYAuP7////ryYn2jbwnAAAAAIPsHIl0JBiLdCQgiVwkFIX2dB2LXhyF23QWi1YghdJ0D4tWJIXSdRmQjbQmAAAAALj+////i1wkFIt0JBiDxBzDx0YsAgAAAItLCItTGMdGFAAAAADHRggAAAAAhdLHRhgAAAAAx0MUAAAAAIlLEHhxMcmF0g+UwUmD4bmDwXGD+gKJSwR0PcdEJAgAAAAAx0QkBAAAAADHBCQAAAAA6I2x//+JRjDHQygAAAAAiRwk" & _
    "6Is7AACJHCTooxEAADHA6XH////HRCQIAAAAAMdEJAQAAAAAxwQkAAAAAOig1f//68GJ0ffZicqJSxjrhI12AItEJASFwHQYi1AchdJ0EYN6GAJ1C4tMJAgxwIlKHMOQuP7////DjXYAjbwnAAAAAFOLRCQIi1wkEIXAdAeLUByF0nUMuP7///9bw5CNdCYAi0wkDLgBAAAAiYq8FgAAD7ZMJAzT4EghwzHAZomauBYAAOvVjbYAAAAAjb8AAAAAg+wciXwkFIt8JCCJXCQMi1wkJIX/iWwkGItsJCiJdCQQx0QkCAAAAAB0JYt3HIX2dB6D+/8PhMUAAACD+wl3EIXteAyD/QR+J420JgAAAAC4/v///4tcJAyLdCQQi3wkFItsJBiDxBzDjbQmAAAAAIuOhAAAAI0UW4sUlUhauGGNBEk5FIVIWrhhdAeLVwiF0nVQOdl0PYmehAAAAI08W8HnAg+3j0JauGEPt59AWrhhD7eXRFq4YYmOgAAAAA+3" & _
    "j0ZauGGJnowAAACJlpAAAACJTnyJrogAAACLRCQI6XL////HRCQEAQAAAIk8JOitAAAAiUQkCIuOhAAAAOuUkLsGAAAA6TH///+NtgAAAACLRCQEhcB0B4tQHIXSdQa4/v///8OLTCQIMcCJiowAAACLTCQMiYqAAAAAi0wkEImKkAAAAItMJBSJSnzDjXYAU4tMJAyLXCQIjUEHwegDjRQIjUE/wegGhduNRBALdBCLUxyF0nQJg3owD3QFjXYAW8ODelAPdfiJTCQIW+lOs///jbQmAAAAAI28JwAAAABVV1ZTg+wci2wkMIXtD4T9CwAAi30chf8PhPILAACDfCQ0BA+P5wsAAItUJDSF0g+I2wsAAItVDIXSD4S8CwAAi0UAhcB1C4tNBIXJD4WqCwAAi0cEPZoCAACJwg+EjwsAAIt1EIX2D4QrAwAAiS+LdyiD+iqLXCQ0iXQkGIlfKA+EIAgAAIP4RQ+ElwYAAIt3FIP4SQ+EWgUAAIP4Ww+E" & _
    "LgQAAIP4Zw+EYwMAAIX2D4Q1AwAAi00ci0UQi1kUOcN2AonDhdsPhd8CAACFwHUYx0co/////zHAjbQmAAAAAIPEHFteX13Di0UEi1cEgfqaAgAAD4SWAgAAhcB1F4t3dIX2dRCLTCQ0hcl0TIH6mgIAAHREi5eEAAAAiTwki0wkNI00UolMJAT/FLVIWrhhjVj+icKD+wF3B8dHBJoCAACF0g+ENwIAAIP6Ag+ELgIAAEoPhGABAAAxwIN8JDQED4V6////i1cYuAEAAACF0g+Oav///4P6Ag+ErQAAAItNMIt3FItXCInLwekYwesQiAwWi3cUi08IRol3FIgcDot3FItPCA+3XTBGiXcUidrB6giIFA6LdxSLVwhGiXcUiBwW/0cUi00ci0UQi1kUOcN2AonDhdt1GotHGIXAfgX32IlHGDHAg38UAA+UwOnt/v//i3UMi1EQiVwkCIk0JIlUJATop6YAACldEItVHAFdDAFdFItqFAFaECndhe2J" & _
    "ahR1tYtaCIlaEOutD7ZNMIt3FItXCIgMFot3FItdMItPCEbB6wiJdxSIHA6LdxSLXwgPt1UyRol3FIgUHot3FItfCA+2TTNGiXcUiAwei18Ui08ID7ZVCEOJXxSIFAuLXxSLVQiLdwhDweoIiV8UiBQzi18Ui3cID7dNCkOJXxSIDDOLdxSLTwgPtlULRol3FIgUDukK////g3wkNAEPhK0AAADHRCQMAAAAAMdEJAgAAAAAx0QkBAAAAACJPCTofCgAAIN8JDQDdFyLTRyLRRCLWRQ5w3YCicOF23UNhcAPhVP+///pxf3//4t1DItREIlcJAiJNCSJVCQE6I+lAAApXRCLVRwBXQwBXRSLchQBWhAp3oX2iXIUdQaLWgiJWhCLRRDrt4tXTIt3RI1MEv5mx0RW/gAAiUwkCMdEJAQAAAAAiTQk6DSlAADpev///4k8JOjHKAAA6W3///+J9otdEIXbD4VM/f//6UD9//+FwA+EZv3//4s9vK+4Ybj7"
Private Const STR_ZLIB_7 As String = _
    "////iX0Y6TX9//+LdQyLURCJXCQIiTQkiVQkBOjvpAAAKV0Qi1UcAV0MAV0Ui3IUAVoQKd6F9olyFHUGi0oIiUoQi0UQ6eL8//+LRQSFwA+F9vz//4tcJBg5XCQ0D4/o/P//g3wkNAQPhN38///riItXHItaLIXbdGqLVwyNTgI50XYTi00ci0UQi1kUOcN2AonDhdt1WY1OAjnRD4dr/P//D7ZdMItPCIgcDot3FItVMItfCEbB6giJdxSIFB7/RxTHRCQIAAAAAMdEJAQAAAAAxwQkAAAAAOjyzv//iUUwi3cUx0cEcQAAAOkd/P//i3UMi1EQiVwkCIk0JIlUJAToCqQAACldEItVHAFdDAFdFItyFAFaECnehfaJchR1BotaCIlaEIt3FItXDOll////i18ci1MkhdJ0Q4nxifY7dwx0e4tXIItDJA+2HAJCiVcgi1cIiBwWi0cUQIlHFIXbdAeLXxyJxuvVicaLRxyLUCyF0nQEOc53HYXbdRHH" & _
    "RwRnAAAAuGcAAADpdPv//4tHBOls+///Kc6JdCQIi0cIAcGJTCQEi00wiQwk6B7O//+JRTCLdxTrwY22AAAAAItTLIXSdAQ5zndoi00ci0UQi1kUOcN2AonDhdt1HIt3FDt3DInxdAiLXxzpVf///7sBAAAA6XP///+LURCLdQyJXCQIiVQkBIk0JOgAowAAKV0Qi1UcAV0MAV0Ui0oUAVoQKdmFyYlKFHWzi1oIiVoQ66spzol0JAiLdwgB8YlMJASLXTCJHCTog83//4lFMOl2////i18ci1MchdJ0V4nx6w2QkJCQkJCQkJCQkJCQO3cMdHyLVyCLQxwPthwCQolXIItXCIgcFotHFECJRxSF23QHi18cicbr1YnGi0cci1AshdJ0BDnOdySF23UYx0cgAAAAAMdHBFsAAAC4WwAAAOk0+v//i0cE6Sz6//8pzol0JAiLdwgB8YlMJASLTTCJDCTo58z//4lFMIt3FOu6i0MshcB0BDnOd2iLTRyL" & _
    "RRCLWRQ5w3YCicOF23Uci3cUO3cMifF0CItfHOlU////uwEAAADpcv///4tREIt1DIlcJAiJVCQEiTQk6M+hAAApXRCLVRwBXQwBXRSLShQBWhAp2YXJiUoUdbOLWgiJWhDrqynOiXQkCItfCAHZiUwkBIt1MIk0JOhSzP//iUUw6Xb///+LXxyLSxCJXCQQhckPhF0BAAAPt1MUi3cUOVcgiXQkFHNEkI20JgAAAAA7dwwPhJcAAACLRyCLUxCLTwgPthwQiBwOi08Ui1cgi18cQUKJTxSJzolXIA+3QxSJXCQQOcJzBInO68SLSyyFyXQGO3QkFHcoi0sUOU8gdAiLRwTp6fj//8dHIAAAAAC4SQAAAMdHBEkAAADp0fj//4tcJBQp3ol0JAiLVwgBVCQUi0QkFIlEJASLdTCJNCToi8v//4lFMIt3FItfHOupi0wkEItZLIXbdAY7dCQUd2iLTRyLRRCLWRQ5w3YCicOF23Uci3cUO3cMiXQkFHQI" & _
    "i18c6TH///+LXxzpXP///4tREIt1DIlcJAiJVCQEiTQk6GqgAAApXRCLVRwBXQwBXRSLShQBWhAp2YXJiUoUdbOLWgiJWhDrq4tUJBQp1ol0JAiLXwgBXCQUi0QkFIlEJASLdTCJNCTo48r//4lFMOls////x0cESQAAALhJAAAA6en3//+DfxgCD4QQAQAAi1cwg+oIweIMg7+IAAAAAY2aAAgAAA+OxwAAADHAweAGCcOLR2yFwHQDg8sgx0cEcQAAALiFEEII9+Mp09HrjQQawegEi1cIicbB5gUpxo1eH4t3FInZwekIiAwWi3cUi1cIRol3FIgcFotfFItPbEOFyYlfFHREi1Uwi3cIidHB6hjB6RCIFDOLXxSLdwhDiV8UiAwzi3cUi08ID7ddMEaJdxSJ2sHqCIgUDotPFItXCEGJTxSIHBH/RxTHRCQIAAAAAMdEJAQAAAAAxwQkAAAAAOiYpf//iUUwi0cE6fP2//+Ll4QAAACD+gEPjir/" & _
    "//+D+gW4AQAAAA+OHv///zHAg/oGD5XAg8AC6Q7////HRCQIAAAAAMdEJAQAAAAAxwQkAAAAAOiUyf//iUUwi08Ui0cIxgQBH4tfFIt3CEOJXxTGBDOLi08Ui1cIQYlPFMYEEQiLRxSLVxyNSAGF0olPFA+FjgAAAItfCLACxgQZAItPFItXCEGJTxTGBBEAi18Ui3cIQ4lfFMYEMwCLTxSLVwhBiU8UxgQRAItfFIt3CEOJXxTGBDMAi1cUi18IjUoBg8ICiVcUi5eEAAAAg/oJdAuDv4gAAAABfiOwBIgEGYtHCIt3FMYEBgu4cQAAAP9HFMdHBHEAAADp5vX//zHASn/a69aDwAKLXwiJRxSLciwxwIM6AA+VwIX2dAODwAKLchCF9nQDg8AEi3IchfZ0A4PACItyJIX2dAODwBCIBBmLdxywAotfFItXCA+2TgSIDBOLdxyLXxSLVwiLTgRDiV8UwekIiAwTi3cci18Ui08ID7dWBkOJXxSIFAuL" & _
    "dxyLXxSLTwgPtlYHQ4lfFIgUC4tXFItfCI1KAYPCAolXFIuXhAAAAIP6CXQPg7+IAAAAAQ+OiQAAALAEiAQZi3cci18Ui1cID7ZODIgME4tfHItPFItDEEGJTxSFwHQmD7ZTFIt3CIgUMYtfFIt3HItXCEOJXxSLThTB6QiIDBOLXxz/RxSLQyyFwHUYx0cgAAAAALhFAAAAx0cERQAAAOnB9P//i3cUiXQkCItPCIlMJASLXTCJHCToksf//4lFMOvKMcBKD49w////6Wn///+DfCQ0BA+EZvT//4sdsK+4YYldGI10JgCNvCcAAAAAuP7////pxvT//422AAAAAIPsHIlcJBSLXCQgiXQkGIXbdAmLUxyF0onRdRS4/v///4tcJBSLdCQYg8Qcw412AItyBIP+KnQog/5FdCOD/kl0HoP+W3QZg/5ndBSD/nF0D4H+mgIAALj+////dcOJ9otBCIXAdW+LQkSFwHVWi0JAhcB1PYtCOIXAdSSJVCQE" & _
    "i0soiQwk/1Mkx0McAAAAADHbg/5xD5XDjURb/euDifaJRCQEi1MoiRQk/1Mki1Mc68qJRCQEi0soiQwk/1Mki1Mc67GJRCQEi1MoiRQk/1Mki1Mc65iJRCQEi1MoiRQk/1Mki1Mc6Xz///+QjXQmAIPsLIl0JCCLdCQ0iVwkHItcJDCF9ol8JCSJbCQodAuF23QHi24che11Gbj+////i1wkHIt0JCCLfCQki2wkKIPELMP8uQ4AAACJ3/Olx0QkCMAWAADHRCQEAQAAAItTKIkUJP9TIIlEJBiLTCQYuPz///+FyXS49kQkGASJ6ItMJBi6wBYAAInPiUscdBCLVQCDxwSNRQSJEbq8FgAA/InRwekCicbzpYtUJBiJGsdEJAgCAAAAi3osiXwkBItzKIk0JP9TIIt0JBiJRjjHRCQIAgAAAItWLIlUJASLSyiJDCT/UyCJRkDHRCQIAgAAAIt+TIl8JASLUyiJFCT/UyCJRkTHRCQIBAAAAIuOnBYA" & _
    "AIlMJASLeyiJPCT/UyCLTjiFyYnHiUYID4TyAAAAi1ZAhdIPhOcAAACLRkSFwA+E3AAAAIX/D4TUAAAAi0QkGIt1OItYLIl0JASJDCQB24lcJAjoOJoAAIt0JBiLTUCLViyLXkCJTCQEAdKJVCQIiRwk6BmaAACLVkyLTUSLXkQB0olUJAiJTCQEiRwk6P6ZAACLTQiLXgiLVgyJTCQEiRwkiVQkCOjlmQAAi00Qi0UIi14IKcGNLBkxwIluEIuunBYAAInpg+H+jRQ5jUxtAImWpBYAAAHLifeJnpgWAACJ9YHHlAAAAIm+GAsAAInzgcWICQAAia4kCwAAgcN8CgAAiZ4wCwAA6RL+//+NdgCJHCTo6Pz//7j8////6f39//+NtCYAAAAAjbwnAAAAAIPsHIlcJBiLXCQgi0ssi0NMAcmNVAD+iUs8i0tEZsdEQf4AAIlUJAiJDCTHRCQEAAAAAOgYmQAAx0NsAAAAAIuThAAAAMdDXAAAAADHQ3QA" & _
    "AAAAjQxSweECD7eRQlq4YcdDeAIAAADHQ2ACAAAAx0NoAAAAAImTgAAAAA+3kUBauGHHQ0gAAAAAiZOMAAAAD7eRRFq4YYmTkAAAAA+3kUZauGGJU3yLXCQYg8Qcw4n2Vb3//wAAV1ZTg+wci3wkMItHDIPoBTnFdgaJxY10JgCLR3SD+AEPhmQBAADHR3QAAAAAi09sAciJR2yLT1yFwI1UDQB0BDnQclaJV2wp0CnKiUd0McCFycdEJAwAAAAAiVQkCHgFi0c4AciJRCQEiTwk6GkeAACLN4tXbItOHIlXXItGEItZFDnDdgKJw4XbD4W6AAAAi0YQhcB0YotPXItXbIt3LCnKge4GAQAAOfIPgmr////HRCQMAAAAADHAhcmJVCQIeAWLRzgByIlEJASJPCToBx4AAItPbIs3iU9ci04ci0YQi1kUOcN2AonDhdt1HItWEIXSD4Uh////McCDxBxbXl9dw420JgAAAACLVgyLQRCJXCQIiRQkiUQk" & _
    "BOiKlwAAKV4Qi1YcAV4MAV4Ui3IUAVoQKd6F9olyFHUGi1oIiVoQizfrqZCNdCYAi1YMi0EQiVwkCIkUJIlEJAToSpcAACleEItWHAFeDAFeFItyFAFaECnehfaJchR1BotaCIlaEIs36Qj///+J9ok8JOgoDAAAi0d0hcAPhYn+//+LVCQ0McCF0g+EUP///zHJg3wkNAQPlMExwIlMJAyLV1yLb2wp1YXSiWwkCHgFi0c4AdCJRCQEiTwk6P4cAACLd2yJd1yLN4tOHItGEItZFDnDdgKJw4XbdTKLfhCF/3URMcCDfCQ0BA+UwAHA6ez+//+DfCQ0BLgDAAAAD4Tc/v//uAEAAADp0v7//4tuDItREIlcJAiJLCSJVCQE6GuWAAApXhCLVhwBXgwBXhSLahQBWhAp3YXtiWoUdQaLWgiJWhCLN+uTjbYAAAAAVVdWU4PsPIt8JFDHRCQ4AAAAAI22AAAAAI28JwAAAACLb3SB/QUBAAAPhgYEAACD"
Private Const STR_ZLIB_8 As String = _
    "/QJ2PItXSItPWIt3ONPii09sD7ZcMQKLd1Qx2iHyi180i3dEiVdIIdmLX0APtwRWZokESw+3T2yJRCQ4ZokMVotUJDiF0nQ4i19si0QkOIt3LInZKcGB7gYBAAA58Xcgi5eIAAAAjXL+g/4BD4d+AwAAg/oDD4S7AgAAkI10JgCDf2ACD4ZKAgAAD7dfbIu3pBYAAIuXmBYAAA+3R3APtk9gKcMPt+uLn6AWAACA6QNmiSxejXX/D7bpiAwTD7fe/4egFgAAD7a1wIu4YQ+3lLeYBAAAQmaB+/8AZomUt5gEAAAPh9oBAAAPtoPAibhhZoH7/wCNFIWACQAAD4eyAQAAD7aDwIm4YY0MhYAJAAAPt1wPCENmiVwXCIu3nBYAAItHYE45t6AWAAAPlMIPtuqLV3SJbCQ0KcI7h4AAAACJV3QPhzkBAACD+gIPhjABAACLd1hIi184iUdgi1dUi280iXQkJItHQIt3RIlcJCyJVCQoiWwkMIlEJByNdCYA" & _
    "D7ZMJCSLV0iLX2zT4otMJCyNawGJb2wPtkQNAotMJDAxwotEJCghzYtMJBwhwg+3BFaJV0hmiQRpi09gD7dvbIlEJDhJhclmiSxWiU9gdbCNUwKJV2yNdgCNvCcAAAAAi0wkNIXJD4QE/v//x0QkDAAAAAAxwItXXIt3bCnWhdKJdCQIeAWLRzgB0IlEJASJPCToKRoAAIs3i1dsi04ciVdci0YQi1kUOcN2AonDhdt1FYtGEIXAD4Wz/f//McCDxDxbXl9dw4tuDItREIlcJAiJLCSJVCQE6LOTAAApXhCLVhwBXgwBXhSLahQBWhAp3YXtiWoUdQaLWgiJWhCLN+uwx0dgAAAAAItXbItPWItfVAHCi0c4iVdsD7Y0Aol3SNPmD7ZsAgEx7iHeiXdI6Sz////B6wcPtoPAirhh6Ub+//+J2cHpBw+2gcCKuGHpHP7//4tvbItHOIu3oBYAAIufpBYAAIuXmBYAAA+2TAUAZscEcwAAD7bZiAwW/4eg" & _
    "FgAAD7esn5QAAABFZomsn5QAAACLt5wWAABOObegFgAAD5TC/090/0dsD7bKiUwkNOmq/v//SQ+FQ/3//4tHOIt0JDiNFAONDAYPtgKNmgIBAAA4AXQNuAIAAACJR2DpHP3//w+2QgE4QQF16oPCAoPBAkFCD7YBOAJ1Q0FCD7YBOAJ1OkFCD7YBOAJ1MUFCD7YBOAJ1KEFCD7YBOAJ1H0FCD7YBOAJ1FkFCD7YBOAJ1DUFCD7YBOAJ1BDnacrQp07gCAAAAugIBAAAp2oP6An6Ki0wkODnqidCJT3APhnn///+J6Oly////jbYAAAAAiTwki2wkOIlsJATowAgAAOlX////iTwk6BMHAACLb3SB/QUBAAB3DotMJFQxwIXJD4QZ/v//he0PhdH7//8x0oN8JFQED5TCMcCJVCQMi1dci09sKdGF0olMJAh4BYtHOAHQiUQkBIk8JOjhFwAAi3dsiXdcizeLThyLRhCLWRQ5w3YCicOF23Uri34Qhf91" & _
    "ETHAg3wkVAQPlMABwOmt/f//McmDfCRUBA+UwY1ECQHpmv3//4tuDItREIlcJAiJLCSJVCQE6FWRAAApXhCLVhwBXgwBXhSLahQBWhAp3YXtiWoUdQaLWgiJWhCLN+uaVVdWU4PsHIt8JDDHRCQYAAAAAI22AAAAAI28JwAAAACLV3SB+gUBAACJVCQQD4bVBAAAg3wkEAIPhsIEAACLXziLb2yLV0iLT1gPtnQdAotfNNPii09UMfKJ7iHeIcqLX0SJV0iLT0APtwRTZokEcQ+3d2yJRCQYZok0U4tEJBiLV2DHR2ACAAAAi09whcCJV3iJT2R0WTuXgAAAAHNRi3QkGInpi1csKfGB6gYBAAA50Xc8i5eIAAAAjVr+g/sBD4ceBAAAg/oDD4RcAwAAi0dgg/gFdxmDv4gAAAABD4Q7AwAAg/gDD4QfAwAAjXYAi0d4g/gCD4YEAgAAOUdgD4f7AQAAD7dfbItPdIuXpBYAAA+3R2QBzYPtAw+2T3gp" & _
    "w0sPt/OLn6AWAACA6QNmiTRai5eYFgAAiAwTjVb/D7bx/4egFgAAD7faD7aOwIu4YQ+3lI+YBAAAQmaB+/8AZomUj5gEAAAPh4IBAAAPtoPAibhhZoH7/wCNFIWACQAAD4daAQAAD7aDwIm4YY0chYAJAAAPt0wfCEFmiUwXCIu3nBYAAItXdE45t6AWAAAPlMMPtsuLX3iJTCQUKdpCiVd0g+sCiV94kI20JgAAAACLd2xGOe6Jd2x3OYtXSItPWItHNNPii084D7ZcDgKLT1QhxjHai19EIcqJV0iLT0APtwRTZokEcQ+3d2yJRCQYZok0U4tXeEqF0olXeHWxx0doAAAAAItHbItsJBTHR2ACAAAAQIXtiUdsD4Tg/f//x0QkDAAAAACLV1wp0IlEJAgxwIXSeAWLRzgB0IlEJASJPCTo+BQAAIsfi3dsi0sciXdci0MQi3EUOcZ2AonGhfZ1Io22AAAAAI28JwAAAACLQxCFwA+Fhf3//zHAg8Qc" & _
    "W15fXcOLawyLURCJdCQIiSwkiVQkBOh1jgAAKXMQi1McAXMMAXMUi1oUAXIQKfOF24laFHUGi0oIiUoQix/rsMHrBw+2g8CKuGHpnv7//4newe4HD7aGwIq4Yel0/v//i3dohfYPhOwAAACLVziLj6AWAACLt5gWAAAPtlwV/4uvpBYAAGbHRE0AAAAPtuuIHDH/h6AWAAAPt5SvlAAAAEJmiZSvlAAAAIufnBYAAEs5n6AWAAAPhZQAAADHRCQMAAAAADHAi1dci19sKdOF0olcJAh4BYtHOAHQiUQkBIk8JOjZEwAAix+Lb2yLSxyJb1yLQxCLcRQ5xnYCicaF9nUO/090jU0BiU9s6eD+//+LURCLawyJdCQIiVQkBIksJOhqjQAAKXMQi1McAXMMAXMUi0oUAXIQKfGFyYlKFHUPi3IIiXIQifaNvCcAAAAAi29six/rq/9PdI1dAcdHaAEAAACJX2zpFPz//4tHcInuKcaB/gAQAAAPhtH8///H" & _
    "R2ACAAAA6cX8//9JD4Wd/P//i0c4i3QkGI1UBQCNDAYPtgKNmgIBAAA4AXQNuAIAAACJR2Dpdfz//w+2QgE4QQF16oPCAoPBAkFCD7YBOAJ1Q0FCD7YBOAJ1OkFCD7YBOAJ1MUFCD7YBOAJ1KEFCD7YBOAJ1H0FCD7YBOAJ1FkFCD7YBOAJ1DUFCD7YBOAJ1BDnacrQp07gCAAAAugIBAAAp2oP6An6KO1QkEInQi0wkGIlPcA+Gd////4tEJBDpbv///4k8JItsJBiJbCQE6OwCAACJR2CLb2zp0Pv//5CLb2zpdPv//4k8JOgwAQAAi0d0PQUBAACJRCQQdw6LTCQ0McCFyQ+Ecf3//4tUJBCF0g+F+/r//4tvaIXtdEeLT2yLRziLt6QWAACLr5gWAAAPtlwB/4uPoBYAAGbHBE4AAA+284gcKf+HoBYAAA+3lLeUAAAAQmaJlLeUAAAAx0doAAAAADHSg3wkNAQPlMIxwIlUJAyLV1yLX2wp04XS" & _
    "iVwkCHgFi0c4AdCJRCQEiTwk6KkRAACLH4tvbItLHIlvXItDEItxFDnGdgKJxoX2dSuLexCF/3URMcCDfCQ0BA+UwAHA6bP8//8x24N8JDQED5TDjUQbAemg/P//i1EQi2sMiXQkCIlUJASJLCToHYsAAClzEItTHAFzDAFzFItKFAFyECnxhcmJShR1BotyCIlyEIsf65qQjbQmAAAAAFVXVlOD7ByLbCQwi30si1V0ifiNtgAAAACNvCcAAAAAi3U8jYw4+v7//4tdbCnWKd45y4l0JBgPg/MAAACLdQCLTgSFyQ+EpwAAAItdbItFOItVdAHDAdM7TCQYiVwkFInLdgSLXCQYMcCF23Q5i1YcKdmJTgSLQhiD+AEPhJUAAACD+AJ0c4sWiVwkCItMJBSJVCQEiQwk6FOKAAABHonYi1V0AV4IjTQQg/4CiXV0diGLRWyLVTiLTVgPtjQQiXVI0+YPtkwQAYtFVDHOIcaJdUiLXXSB+wUBAACJ2ncS" & _
    "i00Ai3EEhfZ0CItFLOku////g8QcW15fXcOJXCQIixaJVCQEi04wiQwk6KG0//+JRjDpcP///4lcJAiLFolUJASLTjCJDCToNJD//+vhifaLTTiJfCQIjRwPiQwkiVwkBOiqiQAAKX1wi01Mi1VEKX1sKX1cjQRKkI20JgAAAACD6AIx2w+3EDn6cgSJ0yn7ZokYSXXqi3VAifmNBH6J9oPoAjHbD7cQOfpyBInTKftmiRhJdeoBfCQY6Z7+//+QVVdWU4PsHItEJDDHRCQQAAAAAItMJDSLWDiLUGyLqJAAAACLeHyJXCQYi3B4i0QkMIlsJBSLXCQYi2gsAdONhfr+//85wnYNidAp6AUGAQAAiUQkEItEJDCLaECLRCQwiWwkDItoNItEJBiJbCQIjawCAgEAAIlsJAQPtkQe/4hEJAOLRCQwD7YUHjuwjAAAAIhUJAJyA8HvAotUJDCLanQ5bCQUdgmJbCQUkI10JgAPtkQkAotUJBgByjgEFnQn" & _
    "i0QkCItUJAwhyA+3DEI7TCQQdgNPddk57onwdgKJ6IPEHFteX13DD7ZEJAM4RBb/dc4PtgM4AnXHD7ZDAUI4AnW+g8MCQkJDD7YCOAN1RUJDD7YCOAN1PEJDD7YCOAN1M0JDD7YCOAN1KkJDD7YCOAN1IUJDD7YCOAN1GEJDD7YCOAN1D0JDD7YCOAN1BjtcJARysotEJAS6AgEAACnYi1wkBCnCgesCAQAAOfIPjk3///87VCQUi3QkMIlOcInWD41R////D7ZEGv+IRCQDD7YUGohUJALpJP///5CQkJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABAAAAAQAAAAEAAAACAAAAAgAAAAIAAAACAAAAAwAAAAMAAAADAAAAAwAAAAQAAAAEAAAABAAAAAQAAAAFAAAABQAAAAUAAAAFAAAAAAAAAI22AAAAAI2/AAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAQAAAAIAAAACAAAA"
Private Const STR_ZLIB_9 As String = _
    "AwAAAAMAAAAEAAAABAAAAAUAAAAFAAAABgAAAAYAAAAHAAAABwAAAAgAAAAIAAAACQAAAAkAAAAKAAAACgAAAAsAAAALAAAADAAAAAwAAAANAAAADQAAAJCNtCYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAADAAAABwAAABAREgAIBwkGCgULBAwDDQIOAQ+QDAAIAIwACABMAAgAzAAIACwACACsAAgAbAAIAOwACAAcAAgAnAAIAFwACADcAAgAPAAIALwACAB8AAgA/AAIAAIACACCAAgAQgAIAMIACAAiAAgAogAIAGIACADiAAgAEgAIAJIACABSAAgA0gAIADIACACyAAgAcgAIAPIACAAKAAgAigAIAEoACADKAAgAKgAIAKoACABqAAgA6gAIABoACACaAAgAWgAIANoACAA6AAgAugAIAHoACAD6AAgA" & _
    "BgAIAIYACABGAAgAxgAIACYACACmAAgAZgAIAOYACAAWAAgAlgAIAFYACADWAAgANgAIALYACAB2AAgA9gAIAA4ACACOAAgATgAIAM4ACAAuAAgArgAIAG4ACADuAAgAHgAIAJ4ACABeAAgA3gAIAD4ACAC+AAgAfgAIAP4ACAABAAgAgQAIAEEACADBAAgAIQAIAKEACABhAAgA4QAIABEACACRAAgAUQAIANEACAAxAAgAsQAIAHEACADxAAgACQAIAIkACABJAAgAyQAIACkACACpAAgAaQAIAOkACAAZAAgAmQAIAFkACADZAAgAOQAIALkACAB5AAgA+QAIAAUACACFAAgARQAIAMUACAAlAAgApQAIAGUACADlAAgAFQAIAJUACABVAAgA1QAIADUACAC1AAgAdQAIAPUACAANAAgAjQAIAE0ACADNAAgALQAIAK0ACABtAAgA7QAIAB0ACACdAAgAXQAIAN0ACAA9AAgAvQAIAH0ACAD9AAgA" & _
    "EwAJABMBCQCTAAkAkwEJAFMACQBTAQkA0wAJANMBCQAzAAkAMwEJALMACQCzAQkAcwAJAHMBCQDzAAkA8wEJAAsACQALAQkAiwAJAIsBCQBLAAkASwEJAMsACQDLAQkAKwAJACsBCQCrAAkAqwEJAGsACQBrAQkA6wAJAOsBCQAbAAkAGwEJAJsACQCbAQkAWwAJAFsBCQDbAAkA2wEJADsACQA7AQkAuwAJALsBCQB7AAkAewEJAPsACQD7AQkABwAJAAcBCQCHAAkAhwEJAEcACQBHAQkAxwAJAMcBCQAnAAkAJwEJAKcACQCnAQkAZwAJAGcBCQDnAAkA5wEJABcACQAXAQkAlwAJAJcBCQBXAAkAVwEJANcACQDXAQkANwAJADcBCQC3AAkAtwEJAHcACQB3AQkA9wAJAPcBCQAPAAkADwEJAI8ACQCPAQkATwAJAE8BCQDPAAkAzwEJAC8ACQAvAQkArwAJAK8BCQBvAAkAbwEJAO8ACQDvAQkA" & _
    "HwAJAB8BCQCfAAkAnwEJAF8ACQBfAQkA3wAJAN8BCQA/AAkAPwEJAL8ACQC/AQkAfwAJAH8BCQD/AAkA/wEJAAAABwBAAAcAIAAHAGAABwAQAAcAUAAHADAABwBwAAcACAAHAEgABwAoAAcAaAAHABgABwBYAAcAOAAHAHgABwAEAAcARAAHACQABwBkAAcAFAAHAFQABwA0AAcAdAAHAAMACACDAAgAQwAIAMMACAAjAAgAowAIAGMACADjAAgAAAAFABAABQAIAAUAGAAFAAQABQAUAAUADAAFABwABQACAAUAEgAFAAoABQAaAAUABgAFABYABQAOAAUAHgAFAAEABQARAAUACQAFABkABQAFAAUAFQAFAA0ABQAdAAUAAwAFABMABQALAAUAGwAFAAcABQAXAAUAkI20JgAAAAAAAQIDBAQFBQYGBgYHBwcHCAgICAgICAgJCQkJCQkJCQoKCgoKCgoKCgoKCgoKCgoLCwsLCwsLCwsLCwsLCwsL" & _
    "DAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwNDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQ4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4ODg4PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PDw8PAAAQERISExMUFBQUFRUVFRYWFhYWFhYWFxcXFxcXFxcYGBgYGBgYGBgYGBgYGBgYGRkZGRkZGRkZGRkZGRkZGRoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxscHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwcHBwc" & _
    "HR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHR0dHQABAgMEBQYHCAgJCQoKCwsMDAwMDQ0NDQ4ODg4PDw8PEBAQEBAQEBARERERERERERISEhISEhISExMTExMTExMUFBQUFBQUFBQUFBQUFBQUFRUVFRUVFRUVFRUVFRUVFRYWFhYWFhYWFhYWFhYWFhYXFxcXFxcXFxcXFxcXFxcXGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgZGRkZGRkZGRkZGRkZGRkZGRkZGRkZGRkZGRkZGRkZGRoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxwAAAAAAQAAAAIAAAADAAAABAAAAAUAAAAGAAAABwAAAAgAAAAKAAAADAAAAA4AAAAQAAAAFAAAABgAAAAcAAAA" & _
    "IAAAACgAAAAwAAAAOAAAAEAAAABQAAAAYAAAAHAAAACAAAAAoAAAAMAAAADgAAAAAAAAAI22AAAAAI2/AAAAAAAAAAABAAAAAgAAAAMAAAAEAAAABgAAAAgAAAAMAAAAEAAAABgAAAAgAAAAMAAAAEAAAABgAAAAgAAAAMAAAAAAAQAAgAEAAAACAAAAAwAAAAQAAAAGAAAACAAAAAwAAAAQAAAAGAAAACAAAAAwAAAAQAAAAGAAAJCNtCYAAAAAg+wQiRwki1wkFIl0JASLdCQgiXwkCIt8JByJbCQMi4u8FgAAi2wkGIP5DQ+OjwAAAInwi1MU0+APt4u4FgAACciLSwhmiYO4FgAAD7aDuBYAAIgECotLFItTCA+2g7kWAABBiUsUiAQRi4O8FgAAuRAAAAD/QxQPt9YpwdP6g+gNZomTuBYAAImDvBYAAIt0JASJfCQci3wkCIlsJBiLbCQMiVwkFIscJMdEJCABAAAAg8QQ6XcMAACNtCYAAAAA" & _
    "D7eTuBYAANPmjUEDCdZmibO4FgAA67GJ9o28JwAAAABVV1ZTg+wMi3QkIIuOvBYAAIP5DQ+OWQIAAA+3hrgWAAC/AgAAAIn90+WLXggJxWaJrrgWAACLbhQPto64FgAAiEwdAIteFItOCA+2lrkWAABDiV4UiBQLi4a8FgAAuRAAAAD/RhQpwdP/Zom+uBYAAIPoDQ+3PcKIuGG9EAAAAInqiYa8FgAAi468FgAAKfo50Q+OugEAAA+3HcCIuGGLVhSJ2NPgD7eOuBYAAAnIi04IZomGuBYAAA+2hrgWAACIBAqLThSLVggPtoa5FgAAQYlOFIgEEYuWvBYAAInp/0YUKdGNRBfw0/tmiZ64FgAAiYa8FgAAiTQk6GEKAACLvrQWAACLjrwWAAApz4PHC4P/CA+P6QAAAIP5DQ+OEQEAAA+3hrgWAAC9AgAAAInv0+eLXggJx2aJvrgWAACLfhQPto64FgAAiAwfi14Ui04ID7aWuRYAAEOJXhSIFAuL" & _
    "hrwWAAC5EAAAAP9GFCnB0/1mia64FgAAg+gND7c9woi4Yb0QAAAAieqJhrwWAACLjrwWAAAp+jnRfnoPtx3AiLhhi1YUidjT4A+3jrgWAAAJyItOCGaJhrgWAAAPtoa4FgAAiAQKi04Ui1YID7aGuRYAAEGJThSIBBGLlrwWAACJ6f9GFCnRjUQX8NP7ZomeuBYAAImGvBYAAIk0JOheCQAAvwcAAACJvrQWAACDxAxbXl9dww+3LcCIuGEPt4a4FgAA0+UJxWaJrrgWAACNBA/rwQ+3nrgWAAC9AgAAANPljUEDCd1mia64FgAA6Sn///8Pty3AiLhhD7eGuBYAANPlCcVmia64FgAAjQQP6X7+//8Pt564FgAAvwIAAADT541BAwnfZom+uBYAAOni/f//jbQmAAAAAI28JwAAAABVV1ZTg+wci3QkMMdEJBgAAAAAi2wkNItcJDiLloQAAACLfCQ8hdIPjhwDAACF23QMixaDeiwCD4T/AgAAiTQk" & _
    "jY4YCwAAiUwkBOjRCwAAiTQkjZYkCwAAiVQkBOi/CwAAi44cCwAAjZaUAAAAiVQkBIk0JIlMJAjoIxEAAIuOKAsAAI2WiAkAAIlUJASJNCSJTCQI6AcRAACJNCSNjjALAACJTCQE6HULAAC5EgAAAA+2kayEuGFmg7yWfgoAAAB1BkmD+QJ/6IlMJBiLhqgWAACNFEmLjqwWAAABwo1CEYmGqBYAAIPBCoPCG8HqA8HpAznRdwKJyo1DBDnQD4e3AAAAhe0PhK8AAACJfCQMiVwkCIlsJASJNCToYfv//zHA6w2QkJCQkJCQkJCQkJCQMclmiYyGlAAAAEA9HQEAAH7uMcCNtgAAAACNvwAAAAAx7WaJrIaICQAAQIP4HX7wMcAx22aJnIZ8CgAAQIP4En7wMdsxyTHSiZ6sFgAAMcC9AQAAAGaJrpQEAACF/4mOqBYAAImWsBYAAImGoBYAAHUIg8QcW15fXcOJdCQwg8QcW15fXemGBwAAg76IAAAA"
Private Const STR_ZLIB_10 As String = _
    "BA+E2QAAADnRD4TRAAAAi468FgAAg/kND46oAAAAD7eGuBYAAI1vBInr0+OLVggJw2aJnrgWAACLXhQPto64FgAAiAwTi14Ui04ID7aWuRYAAEOJXhSIFAuLhrwWAAC5EAAAAP9GFA+31SnB0/qD6A1miZa4FgAAiYa8FgAAi1QkGEKJVCQMi54oCwAAQ4lcJAiLrhwLAACJNCRFiWwkBOhrEAAAjY6ICQAAjZaUAAAAiUwkCIlUJASJNCTo3wAAAOmZ/v//D7eeuBYAAI1vBNPljUEDCd1mia64FgAA65iLjrwWAACD+Q1+dw+3hrgWAACNXwKJ3dPli1YICcVmia64FgAAi24UD7aOuBYAAIhMFQCLbhSLTggPtpa5FgAARYluFIhUDQCLhrwWAAC5EAAAAP9GFA+36ynB0/2D6A1mia64FgAAiYa8FgAAx0QkCECJuGHHRCQEwIS4YelX////D7euuBYAAI1fAtPjjUEDCetmiZ64FgAA68uJNCTo" & _
    "7AQAAOn0/P//jUsF6aX9///rDZCQkJCQkJCQkJCQkJBVV1ZTg+wMi3QkIMdEJAQAAAAAi56gFgAAhdsPhNAAAADrDZCQkJCQkJCQkJCQkJCLTCQEi66kFgAAi5aYFgAAD7d8TQCJfCQIi0QkCA+2LBFBhcCJTCQED4WCAQAAi0wkJLoQAAAAD7d8qQKLjrwWAAAp+jnRD442AQAAi0QkJItWCA+3HKgPt4a4FgAAid3T5QnFZomuuBYAAItuFA+2jrgWAACITBUAi24Ui04ID7aWuRYAAEWJbhSIVA0Ai668FgAAuRAAAAD/RhQp6Y1EL/DT+2aJnrgWAACJhrwWAACLnqAWAAA5XCQED4I/////i0wkJL0QAAAAiesPt7kCBAAAi468FgAAKfs52Q+OfQAAAItEJCSLVhQPt5gABAAAidjT4A+3jrgWAAAJyItOCGaJhrgWAAAPtoa4FgAAiAQKi04Ui1YID7aGuRYAAEGJThSIBBGLlrwWAACJ6f9G" & _
    "FCnRjUQX8NP7ZomeuBYAAImGvBYAAItMJCQPt7kCBAAAib60FgAAg8QMW15fXcOQD7eeuBYAAItEJCQPt6gABAAAjQQP0+UJ3WaJrrgWAADrvotUJCQPtwSqjRQPD7euuBYAAImWvBYAANPgCehmiYa4FgAA6Q3///+NtgAAAAAPtr3Ai7hhi1wkJA+3lLsGBAAAuxAAAACJFCSLDCQpy4uOvBYAADnZD46SAgAAi0QkJItWFA+3nLgEBAAAidjT4A+3jrgWAAAJyItOCGaJhrgWAAAPtoa4FgAAiAQKi04Ui1YID7aGuRYAAEGJThSIBBGLhrwWAAC5EAAAAP9GFCnB0/tmiZ64FgAAiwwkjUQB8ImGvBYAAIscvWCDuGGF23R7iwy9wIy4Yb8QAAAAifopzSnai468FgAAOdEPjuABAACJ6ItWFNPgD7eOuBYAAAnIi04IZomGuBYAAA+2hrgWAACIBAqLThSLVggPtoa5FgAAQYlOFIgEEQ+3xYn5" & _
    "/0YUi668FgAAKenT+GaJhrgWAACNRCvwiYa8FgAA/0wkCIF8JAj/AAAAD4dgAQAAi1wkCA+2u8CJuGGLVCQouxAAAACLjrwWAAAPt2y6AinrOdkPjg4BAAAPtxy6i1YUidjT4A+3jrgWAAAJyItOCGaJhrgWAAAPtoa4FgAAiAQKi04Ui1YID7aGuRYAAEGJThSIBBGLlrwWAAC5EAAAAP9GFCnRjUQV8NP7ZomeuBYAAImGvBYAAIscveCDuGGF2w+EOv3//4sUvUCNuGG/EAAAAIuOvBYAAClUJAiJ/SndOel+YA+3hrgWAACLbCQIi1YI0+UJxWaJrrgWAACLbhQPto64FgAAiEwVAItuFItOCA+2lrkWAABFiW4UiFQNAIuuvBYAAIn5D7dEJAj/RhQp6dP4ZomGuBYAAI1EK/Dptvz//9NkJAiNBAsPt5a4FgAACVQkCIt8JAhmib64FgAA6ZT8//+LRCQoD7ccuA+3hrgWAADT4wnDZomeuBYA" & _
    "AI1EDQDpKP///422AAAAAItUJAjB6gcPtrrAirhh6Zj+//8Pt764FgAA0+WNBAsJ/WaJrrgWAADpW/7//412AItEJCQPt5y4BAQAAA+3hrgWAADT4wnDZomeuBYAAIscJI0EC+mp/f//jbQmAAAAAItEJAQx0o12AI28JwAAAABmg7yQlAAAAAB1BkKD+gh+74P6CXQPg/ogiwgPlMIPttKJUSzDug4AAACNtgAAAABmg7yQlAAAAAB120KD+h9+7+vTjbYAAAAAjbwnAAAAAFOLTCQIMduLVCQMkI10JgCJyEqD4AEJw9HpAduF0n/w0euJ2FvDjXYAjbwnAAAAAFOLXCQIi4O8FgAAg/gQdDCD+Ad+KA+2g7gWAACLSwiLUxSIBAoPtou5FgAA/0MUg6u8FgAACGaJi7gWAABbw5APtoO4FgAAi1MUi0sIiAQKi0sUi1MID7aDuRYAAEGJSxSIBBEx0v9DFGaJk7gWAAAx0omTvBYAAOvDjbYAAAAA" & _
    "U4tcJAiLg7wWAACD+Ah+QA+2g7gWAACLUxSLSwiIBAqLSxSLUwgPtoO5FgAAQYlLFIgEEf9DFDHSZomTuBYAADHSiZO8FgAAW8ONtgAAAACFwH7jD7aDuBYAAItLCItTFIgECuvOjXYAjbwnAAAAAFdWU4PsEIt0JCCLfCQki1wkKIk0JOh2////uggAAACJlrQWAACLVCQshdJ0QYtOFItGCIgcAYtOFA+2x4tWCEGJThSIBBGLThSI2ItWCPbQQYlOFIgEEYtOFInYi1YI99APtsRBiU4UiAQR/0YUS4P7/3QijbYAAAAAjb8AAAAAD7YHS4tWFItOCEeIBAr/RhSD+/916oPEEFteX8ONdgCLVCQEMcCNipQAAACJihgLAAC5ABC5YYmKIAsAAI2KiAkAAImKJAsAALkUELlhiYosCwAAjYp8CgAAiYowCwAAuSgQuWGJijgLAAAxyWaJirgWAAAxyYmKvBYAALkIAAAAiYq0FgAAifaNvCcAAAAA" & _
    "MclmiYyClAAAAEA9HQEAAH7uMcCNtgAAAACNvwAAAAAxyWaJjIKICQAAQIP4HX7wMcAxyWaJjIJ8CgAAQIP4En7wuQEAAAAxwGaJipQEAAAxyYmKrBYAADHJiYqoFgAAMcmJirAWAACJgqAWAADDkIPsDItMJBSJHCSLXCQQiXQkBIt0JBiJfCQIi4OgFgAAi7ukFgAAifJmiQxHi7uYFgAAiBQ4/4OgFgAAhcl1Mw+3lLOUAAAAQmaJlLOUAAAAi7OcFgAAi3wkCE45s6AWAACLdCQED5TDD7bDixwkg8QMww+2vsCLuGFJ/4OwFgAAD7e0u5gEAABGgfn/AAAAZom0u5gEAAB3Rw+2gcCJuGGB+f8AAACNFIWACQAAdyEPtoHAibhhjTyFgAkAAA+3TDsIQWaJTBMI64aNtgAAAADB6QcPtoHAirhh69qNdCYAicrB6gcPtoLAirhh67KJ9lVXVlOD7Hwx28dEJEgAAAAAi5QklAAAAIu0JJAAAADH" & _
    "RCRE/////4tqCIs6i1UMiZ5QFAAAi00AvT0CAAA5VCRIiVQkTImuVBQAAH1TjXYAi1wkSGaDPJ8AD4QOBQAAi5ZQFAAAMcCLXCRIQolcJESJllAUAACJnJZcCwAAiIQzWBQAAP9EJEiLXCRMOVwkSHy+jbQmAAAAAI28JwAAAACLhlAUAACD+AGJRCQcf1GNUAExwIN8JEQBiZZQFAAAfjWJRCRAiYSWXAsAAGbHBIcBAItUJEAxwP+OqBYAAIXJiIQyWBQAAHS3D7dskQIprqwWAADrqv9EJESLRCRE68GLRCREi6wklAAAAIlFBIueUBQAAInZiVwkHMHpHwHZiVwkPNH5hcmJTCRID462AAAA6w2QkJCQkJCQkJCQkJCQi2wkSInpi5SuXAsAAAHJO0wkPIlUJDh/c420JgAAAAA7TCQ8D433AwAAi5SOYAsAAIucjlwLAACJVCQ0D7cUlw+3BJ9mOcIPgtMDAABmOcIPhKgDAACNdgCNvCcAAAAA" & _
    "i0QkOA+3FIcPtwSfZjnCchtmOcIPhGYDAACJnK5cCwAAic2NDAk7TCQ8fpT/TCRIi1wkOItMJEiJnK5cCwAAhckPj1n///+LbCRMiWwkQJDHRCQsAQAAAItMJBy7AgAAAIuWYAsAAIuEjlwLAACJVCRIiUQkMImGYAsAAItEJBxIiYZQFAAAOcN/dolEJByQO1wkHA+N1wIAAIusnmALAACLjJ5cCwAAD7cUrw+3BI9mOcIPgrcCAABmOcIPhIwCAACNtCYAAAAAjbwnAAAAAA+3BI+LbCQwD7cUr2Y5wnIgZjnCD4RGAgAAi1QkLIlcJCwB2ztcJByJjJZcCwAAfo+LXCQwi0wkLIuWVBQAAItsJEiJnI5cCwAAi45gCwAAiayWWAsAAIPqAomWVBQAAImMllwLAAAPtxyvi2wkQA+3BI8Bw2aJHK+LXCRID7aEMVgUAAAPtpQzWBQAADjCcgKI0ItcJED+wIiEM1gUAABmiVyPAotEJEC7AgAAAItM" & _
    "JEhmiUSPAotUJEDHRCQkAQAAAInVRYlsJECLrlAUAACJVCQoiZZgCwAAOeuJbCQcf32NtCYAAAAAjbwnAAAAADnrD41ZAQAAi4SeYAsAAIuMnlwLAACJRCQgD7cUhw+3BI9mOcIPgjUBAABmOcIPhBEBAACNtgAAAACNvwAAAACLRCQoD7cUhw+3BI9mOcJyHmY5wg+EzQAAAItUJCSJXCQkAds564mMllwLAAB+kYtcJChNi0wkJImcjlwLAAAPjwv+//+LjlQUAACNnjwLAACLlmALAACLrCSUAAAASYmUjlwLAACJjlQUAACJbCQEiTQk6GgFAAAxyboBAAAAkA+3bFP+jXQNAAH2Zol0VFAPt85Cg/oPfucx2ztcJER/H+sNkJCQkJCQkJCQkJCQkA+3TJ8Chcl1D0M7XCREfvCDxHxbXl9dw4lMJAQPt3RMUI1GAWaJRExQiTQk6O/3//9miQSf69IPtoQxWBQAAItUJCg4hDJYFAAAD4ca////"
Private Const STR_ZLIB_11 As String = _
    "6Sr///8PtoQxWBQAAItUJCA4hDJYFAAAD4fi/v//ifZDi4yeXAsAAOnT/v//jXYAD7aEMVgUAACLbCQwOIQ1WBQAAA+Hof3//+mz/f//D7aUMVgUAAA4lDVYFAAAD4dt/f//jbYAAAAAjbwnAAAAAEOLjJ5cCwAA6VP9//+NdgAPtoQzWBQAAItUJDg4hDJYFAAAD4eB/P//6Y78//8PtoQzWBQAAItUJDQ4hDJYFAAAD4dJ/P//ifaNvCcAAAAAQYucjlwLAADpM/z//4tsJEhmx0SvAgAA6Qf7//+NdgBVMclXvwcAAABWU4PsCItUJCDHRCQE/////4tsJBwPt1oCxwQkBAAAAIXbdQzHBCQDAAAAv4oAAACLRCQkMfaLVCQgOcZmx0SCBv//f1mJ9otEJCBBido5+Q+3XLAGfQQ52nQ8OwwkfWeLvJV8CgAAAc9mibyVfAoAAIlUJAQxyb+KAAAAhdt0NznadCPHBCQEAAAAvwcAAACNtCYAAAAA" & _
    "Rjt0JCR+qYPECFteX13DkL8GAAAAjXQmAI28JwAAAADHBCQDAAAA69eNtCYAAAAAhdJ0LDtUJAR0EIuMlXwKAABBZomMlXwKAACLvbwKAABHZom9vAoAAOuAjbYAAAAAg/kKfxOLjcAKAABBZomNwAoAAOli////i73ECgAAR2aJvcQKAADpT////5CNdCYAVVdWU4PsHIt0JDCLRCQ0i3wkOIuOvBYAAIlEJBiLbCQ8g/kLD456AgAAicOLVhSB6wEBAACJ2NPgD7eOuBYAAAnIi04IZomGuBYAAA+2hrgWAACIBAqLThSLVggPtoa5FgAAQYlOFIgEEYuGvBYAALkQAAAA/0YUD7fTKcHT+oPoC2aJlrgWAACJhrwWAACLjrwWAACD+QsPjuUBAAAPt4a4FgAAT4n7iXwkFNPji34UCcNmiZ64FgAAi14ID7aOuBYAAIgMH4teFItOCA+2lrkWAABDiV4UiBQLi4a8FgAAuRAAAAAPt3wkFP9GFCnB" & _
    "g+gL0/9mib64FgAAiYa8FgAAi468FgAAg/kMD45VAQAAD7eGuBYAAI19/In70+OLVggJw2aJnrgWAACLXhQPto64FgAAiAwTi14Ui04ID7aWuRYAAEOJXhSIFAuLhrwWAAC5EAAAAP9GFA+31ynB0/qD6AxmiZa4FgAAiYa8FgAAMf857w+NggAAAJCNdCYAi468FgAAg/kND46xAAAAD7aXrIS4YQ+3nJZ+CgAAi1YUidjT4A+3jrgWAAAJyItOCGaJhrgWAAAPtoa4FgAAiAQKi04Ui1YID7aGuRYAAEGJThSIBBGLhrwWAAC5EAAAAP9GFCnB0/tmiZ64FgAAg+gNiYa8FgAARznvfIOJNCSLXCQYjY6UAAAAiUwkBI2uiAkAAEuJXCQI6OICAACJbCQ0i3wkFIl0JDCJfCQ4g8QcW15fXenGAgAAjbYAAAAAD7aXrIS4YY1BAw+3nJZ+CgAAD7eWuBYAANPjCdNmiZ64FgAA64wPt564FgAAjX38" & _
    "0+eNQQQJ32aJvrgWAADp6P7//0+NQQWJ+4l8JBTT4w+3vrgWAAAJ+2aJnrgWAADpWP7//w+3hrgWAACLXCQYgesBAQAA0+MJw41BBWaJnrgWAADpwf3//5CNtCYAAAAAVVdWU4PsIItcJDjHRCQEAAAAAItsJDSLM4tDBIl0JByLcwgx24lEJBiLVgiLPotOBIlUJAyLVhCJfCQUiUwkEIlUJAiNtgAAAACNvwAAAAAx0maJlF08CwAAQ4P7D37wi51UFAAAi0wkHIu8nVwLAABmx0S5AgAAi71UFAAAR4H/PAIAAA+P+gAAAJCNdCYAi0QkHIu0vVwLAAAPt1ywAg+3TJgCjVkBO1wkCH4I/0QkBItcJAiLVCQcZolcsgI7dCQYD473AAAAR4H/PAIAAH6+i3QkBIX2D4SnAAAAi0wkCI1R/410JgCNvCcAAAAAD7eEVTwLAACJ02aFwHUPkEsPt4RdPAsAAGaFwHTyD7eMXT4LAABIg2wkBAJmiYRd" & _
    "PAsAAIPBAmaJjF0+CwAAi0wkCItcJAQPt7RNPAsAAE6F22aJtE08CwAAf6SFyYnLdDONtCYAAAAAjbwnAAAAAA+3tF08CwAAhfZ0Fo10JgBPi4y9XAsAADtMJBh+EoX2de5LdduDxCBbXl9dw412AItUJBwPt0SKAjnYdCCJ2inCidCLVCQcD7cUig+vwotUJBwBhagWAABmiVyKAk7rvg+3jF08CwAAQWaJjF08CwAAMck7dCQMfA+LTCQMifCLVCQQKciLDIKLRCQcD7cUsI0EGYtcJBQPr8IBhagWAACF2w+Ewf7//4tcJBQPt0SzAgHID6/QAZWsFgAA6aj+///rDZCQkJCQkJCQkJCQkJBVuf////8x7Ve4BAAAAFZTg+wki1wkPIt8JDgPt1MCiVQkGIt0JBi6BwAAAIX2dQq6igAAALgDAAAAx0QkIAAAAACLdCRAOXQkIA+P/QAAAItcJBhFi3QkPDnViVwkHItcJCAPt1yeBolcJBh9Cjlc" & _
    "JBwPhMMAAAA5xQ+NFQEAALsQAAAAjXQmAI28JwAAAACLTCQcidoPt7SPfgoAAIuPvBYAACnyOdEPjsQAAACLRCQcD7eUh3wKAACJ0NPgD7ePuBYAAIlUJBSLVxQJyGaJh7gWAACLTwgPtoe4FgAAiAQKi08Ui1cID7aHuRYAAEGJTxSIBBGLh7wWAACJ2f9HFCnBjUQG8NN8JBSLTCQUZomPuBYAAImHvBYAAE0PhW3///+LVCQYMe2LTCQchdJ1JLqKAAAAuAMAAAD/RCQgi3QkQDl0JCAPjgP///+DxCRbXl9dw4tEJBg5RCQcdAy6BwAAALgEAAAA69C6BgAAAOvEi1QkHA+3hJd8CgAAD7eXuBYAANPgCdBmiYe4FgAAjQQO64GLXCQchdsPhAMCAAA5TCQcD4SNAAAAi1QkHL4QAAAAi4+8FgAAD7ecl34KAAAp3jnxiVwkEA+OqgEAAItUJBwPt7SXfAoAAItXFInw0+APt4+4FgAACciLTwhm" & _
    "iYe4FgAAD7aHuBYAAIgECotPFItXCA+2h7kWAABBiU8UiAQRi5e8FgAAuRAAAAD/RxQp0Y1EE/DT/maJt7gWAACJh7wWAABND7ePvgoAALoQAAAAKcqJTCQMi4+8FgAAOdEPjv0AAAAPt7e8CgAAi1cID7eHuBYAAInz0+MJw2aJn7gWAACLXxQPto+4FgAAiAwTi18Ui08ID7aXuRYAAEOJXxSIFAuLn7wWAAC5EAAAAP9HFCnZ0/5mibe4FgAAi3QkDI1EHvCJh7wWAACLj7wWAACD+Q5+cg+3h7gWAACNXf2J3dPli3cICcVmia+4FgAAi28UD7aPuBYAAIhMNQCLdxQPt+sPtpe5FgAAi08IRol3FIgUDouHvBYAALkQAAAA/0cUKcHT/WaJr7gWAACD6A6NtgAAAACNvwAAAACJh7wWAADp6P3//w+3t7gWAACNXf3T441BAgnzZomfuBYAAOvbD7efvAoAAIt0JAwPt4e4FgAA0+MJw2aJn7gW" & _
    "AACNBA7pPv///w+3h7gWAACLdCQcD7ect3wKAACLdCQQ0+MJw40EDmaJn7gWAADpjf7//4P9Cg+PKgEAAA+3j8IKAAC6EAAAACnKiUwkBIuPvBYAADnRD47lAAAAD7e3wAoAAItXCA+3h7gWAACJ89PjCcNmiZ+4FgAAi18UD7aPuBYAAIgME4tfFItPCA+2l7kWAABDiV8UiBQLi5+8FgAAuRAAAAD/RxQp2dP+Zom3uBYAAIt0JASNRB7wiYe8FgAAi4+8FgAAg/kNfmAPt4e4FgAAjV39id3T5Yt3CAnFZomvuBYAAItvFA+2j7gWAACITDUAi3cUD7frD7aXuRYAAItPCEaJdxSIFA6Lh7wWAAC5EAAAAP9HFCnB0/1mia+4FgAAg+gN6ZL+//8Pt7e4FgAAjV390+ONQQPpmP7//w+3n8AKAACLdCQED7eHuBYAANPjCcNmiZ+4FgAAjQQO6Vb///8Pt4/GCgAAuhAAAAApyokMJIuPvBYAADnR" & _
    "D47kAAAAD7e3xAoAAItXCA+3h7gWAACJ89PjCcNmiZ+4FgAAi18UD7aPuBYAAIgME4tfFItPCA+2l7kWAABDiV8UiBQLi5+8FgAAuRAAAAD/RxQp2dP+Zom3uBYAAIs0JI1EHvCJh7wWAACLj7wWAACD+Ql+YA+3h7gWAACNXfWJ3dPli3cICcVmia+4FgAAi28UD7aPuBYAAIhMNQCLdxQPt+sPtpe5FgAAi08IRol3FIgUDouHvBYAALkQAAAA/0cUKcHT/WaJr7gWAACD6Anpav3//w+3t7gWAACNXfXT441BB+lw/f//D7efxAoAAIs0JA+3h7gWAADT4wnDZomfuBYAAI0EDulX////kJCQkJCQkJCQkJCQkJCQkJCQbmVlZCBkaWN0aW9uYXJ5AHN0cmVhbSBlbmQAAGZpbGUgZXJyb3IAc3RyZWFtIGVycm9yAGRhdGEgZXJyb3IAaW5zdWZmaWNpZW50IG1lbW9yeQBidWZmZXIgZXJyb3IA" & _
    "aW5jb21wYXRpYmxlIHZlcnNpb24AjXQmAI28JwAAAAAgr7hhMK+4YTuvuGE8r7hhR6+4YVSvuGFfr7hhc6+4YYCvuGE7r7hhMS4yLjMAifa4yK+4YcONdgCNvCcAAAAAuJUAAADDjXYAjbwnAAAAAItEJAS5AgAAACnBiwSNoK+4YcONtgAAAACNvCcAAAAAi1QkDItEJAgPr9CJVCQE6YxaAACNtgAAAACNvwAAAACD7AyLVCQUiRQk6FFaAACDxAzDkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCLRCQEhcB0B4tQHIXSdQa4/v///8PHQhwAAAAAjYowBQAAx0AUAAAAAMdACAAAAADHQDABAAAAx0AYAAAAADHAxwIAAAAAx0IEAAAAAMdCDAAAAADHQhQAgAAAx0IgAAAAAMdCKAAAAADHQiwAAAAAx0IwAAAAAMdCOAAAAADHQjwAAAAAiUpsiUpQiUpMw5CD7AyLRCQQiRwki0wkFIl0"
Private Const STR_ZLIB_12 As String = _
    "JASLXCQYhcCJfCQIdCKLUByF0nQbg/kQfxaLcjyNPDGD/yB2H410JgCNvCcAAAAAuP7///+LHCSLdCQEi3wkCIPEDMOJejy4AQAAANPgSInxIcPT4zHAAVo469mNdCYAg+wci0QkKIlcJBSLVCQsiXQkGItcJCCFwIt0JCR0BYA4MXQRuPr///+LXCQUi3QkGIPEHMOD+jh16oXbuP7///905sdDGAAAAACLUyCF0nUTx0MgELC4YboQsLhhx0MoAAAAAItLJIXJdQfHQyQwsLhhx0QkCDAlAADHRCQEAQAAAItLKIkMJP/SicKF0rj8////dJSJUxyF9nhVifHB+QRBiUoIg/4vfwOD5g+NTviD+Qd3HolyJMdCNAAAAACLdCQYiVwkIItcJBSDxBzpLf7//4lUJASLUyiJFCT/UyTHQxwAAAAAuP7////pOP///8dCCAAAAAD33uuxjXQmAIPsHItUJCTHRCQEDwAAAItMJCiJVCQIi1QkIIlMJAyJ" & _
    "FCTo2f7//4PEHMOQkJCQkJCQkJCQkJCQkJCQkJCQkJAQABEAEgAAAAgABwAJAAYACgAFAAsABAAMAAMADQACAA4AAQAPAJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQYAcAAAAIUAAACBAAFAhzABIHHwAACHAAAAgwAAAJwAAQBwoAAAhgAAAIIAAACaAAAAgAAAAIgAAACEAAAAngABAHBgAACFgAAAgYAAAJkAATBzsAAAh4AAAIOAAACdAAEQcRAAAIaAAACCgAAAmwAAAICAAACIgAAAhIAAAJ8AAQBwQAAAhUAAAIFAAVCOMAEwcrAAAIdAAACDQAAAnIABEHDQAACGQAAAgkAAAJqAAACAQAAAiEAAAIRAAACegAEAcIAAAIXAAACBwAAAmYABQHUwAACHwAAAg8AAAJ2AASBxcAAAhsAAAILAAACbgAAAgMAAAIjAAACEwAAAn4ABAHAwAACFIAAAgSABUIowATByMAAAhyAAAIMgAACcQA" & _
    "EQcLAAAIYgAACCIAAAmkAAAIAgAACIIAAAhCAAAJ5AAQBwcAAAhaAAAIGgAACZQAFAdDAAAIegAACDoAAAnUABIHEwAACGoAAAgqAAAJtAAACAoAAAiKAAAISgAACfQAEAcFAAAIVgAACBYAQAgAABMHMwAACHYAAAg2AAAJzAARBw8AAAhmAAAIJgAACawAAAgGAAAIhgAACEYAAAnsABAHCQAACF4AAAgeAAAJnAAUB2MAAAh+AAAIPgAACdwAEgcbAAAIbgAACC4AAAm8AAAIDgAACI4AAAhOAAAJ/ABgBwAAAAhRAAAIEQAVCIMAEgcfAAAIcQAACDEAAAnCABAHCgAACGEAAAghAAAJogAACAEAAAiBAAAIQQAACeIAEAcGAAAIWQAACBkAAAmSABMHOwAACHkAAAg5AAAJ0gARBxEAAAhpAAAIKQAACbIAAAgJAAAIiQAACEkAAAnyABAHBAAACFUAAAgVABAIAgETBysAAAh1AAAINQAACcoA" & _
    "EQcNAAAIZQAACCUAAAmqAAAIBQAACIUAAAhFAAAJ6gAQBwgAAAhdAAAIHQAACZoAFAdTAAAIfQAACD0AAAnaABIHFwAACG0AAAgtAAAJugAACA0AAAiNAAAITQAACfoAEAcDAAAIUwAACBMAFQjDABMHIwAACHMAAAgzAAAJxgARBwsAAAhjAAAIIwAACaYAAAgDAAAIgwAACEMAAAnmABAHBwAACFsAAAgbAAAJlgAUB0MAAAh7AAAIOwAACdYAEgcTAAAIawAACCsAAAm2AAAICwAACIsAAAhLAAAJ9gAQBwUAAAhXAAAIFwBACAAAEwczAAAIdwAACDcAAAnOABEHDwAACGcAAAgnAAAJrgAACAcAAAiHAAAIRwAACe4AEAcJAAAIXwAACB8AAAmeABQHYwAACH8AAAg/AAAJ3gASBxsAAAhvAAAILwAACb4AAAgPAAAIjwAACE8AAAn+AGAHAAAACFAAAAgQABQIcwASBx8AAAhwAAAIMAAACcEA" & _
    "EAcKAAAIYAAACCAAAAmhAAAIAAAACIAAAAhAAAAJ4QAQBwYAAAhYAAAIGAAACZEAEwc7AAAIeAAACDgAAAnRABEHEQAACGgAAAgoAAAJsQAACAgAAAiIAAAISAAACfEAEAcEAAAIVAAACBQAFQjjABMHKwAACHQAAAg0AAAJyQARBw0AAAhkAAAIJAAACakAAAgEAAAIhAAACEQAAAnpABAHCAAACFwAAAgcAAAJmQAUB1MAAAh8AAAIPAAACdkAEgcXAAAIbAAACCwAAAm5AAAIDAAACIwAAAhMAAAJ+QAQBwMAAAhSAAAIEgAVCKMAEwcjAAAIcgAACDIAAAnFABEHCwAACGIAAAgiAAAJpQAACAIAAAiCAAAIQgAACeUAEAcHAAAIWgAACBoAAAmVABQHQwAACHoAAAg6AAAJ1QASBxMAAAhqAAAIKgAACbUAAAgKAAAIigAACEoAAAn1ABAHBQAACFYAAAgWAEAIAAATBzMAAAh2AAAINgAACc0A" & _
    "EQcPAAAIZgAACCYAAAmtAAAIBgAACIYAAAhGAAAJ7QAQBwkAAAheAAAIHgAACZ0AFAdjAAAIfgAACD4AAAndABIHGwAACG4AAAguAAAJvQAACA4AAAiOAAAITgAACf0AYAcAAAAIUQAACBEAFQiDABIHHwAACHEAAAgxAAAJwwAQBwoAAAhhAAAIIQAACaMAAAgBAAAIgQAACEEAAAnjABAHBgAACFkAAAgZAAAJkwATBzsAAAh5AAAIOQAACdMAEQcRAAAIaQAACCkAAAmzAAAICQAACIkAAAhJAAAJ8wAQBwQAAAhVAAAIFQAQCAIBEwcrAAAIdQAACDUAAAnLABEHDQAACGUAAAglAAAJqwAACAUAAAiFAAAIRQAACesAEAcIAAAIXQAACB0AAAmbABQHUwAACH0AAAg9AAAJ2wASBxcAAAhtAAAILQAACbsAAAgNAAAIjQAACE0AAAn7ABAHAwAACFMAAAgTABUIwwATByMAAAhzAAAIMwAACccA" & _
    "EQcLAAAIYwAACCMAAAmnAAAIAwAACIMAAAhDAAAJ5wAQBwcAAAhbAAAIGwAACZcAFAdDAAAIewAACDsAAAnXABIHEwAACGsAAAgrAAAJtwAACAsAAAiLAAAISwAACfcAEAcFAAAIVwAACBcAQAgAABMHMwAACHcAAAg3AAAJzwARBw8AAAhnAAAIJwAACa8AAAgHAAAIhwAACEcAAAnvABAHCQAACF8AAAgfAAAJnwAUB2MAAAh/AAAIPwAACd8AEgcbAAAIbwAACC8AAAm/AAAIDwAACI8AAAhPAAAJ/wAQBQEAFwUBARMFEQAbBQEQEQUFABkFAQQVBUEAHQUBQBAFAwAYBQECFAUhABwFASASBQkAGgUBCBYFgQBABQAAEAUCABcFgQETBRkAGwUBGBEFBwAZBQEGFQVhAB0FAWAQBQQAGAUBAxQFMQAcBQEwEgUNABoFAQwWBcEAQAUAAGluY29ycmVjdCBoZWFkZXIgY2hlY2sAdW5rbm93biBj" & _
    "b21wcmVzc2lvbiBtZXRob2QAaW52YWxpZCB3aW5kb3cgc2l6ZQB1bmtub3duIGhlYWRlciBmbGFncyBzZXQAaGVhZGVyIGNyYyBtaXNtYXRjaABpbnZhbGlkIGJsb2NrIHR5cGUAaW52YWxpZCBzdG9yZWQgYmxvY2sgbGVuZ3RocwCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkHRvbyBtYW55IGxlbmd0aCBvciBkaXN0YW5jZSBzeW1ib2xzAGludmFsaWQgY29kZSBsZW5ndGhzIHNldABpbnZhbGlkIGxpdGVyYWwvbGVuZ3RocyBzZXQAaW52YWxpZCBkaXN0YW5jZXMgc2V0AGludmFsaWQgbGl0ZXJhbC9sZW5ndGggY29kZQBpbnZhbGlkIGRpc3RhbmNlIGNvZGUAaW52YWxpZCBkaXN0YW5jZSB0b28gZmFyIGJhY2sAaW52YWxpZCBiaXQgbGVuZ3RoIHJlcGVhdABpbmNvcnJl" & _
    "Y3QgZGF0YSBjaGVjawBpbmNvcnJlY3QgbGVuZ3RoIGNoZWNrAI10JgCNvCcAAAAAVVdWU4PsfIuEJJAAAACFwHQui4QkkAAAAItoHIXtdCCLUAyF0nQZiwiFyQ+FVhgAAItwBIX2iXQkKHQUjXQmALj+////g8R8W15fXcONdgCLRQCD+AsPhBoYAACJTCRwi3QkKIucJJAAAACJVCRsi004iXQkaIt7EMdEJFgAAAAAi3U8iUwkYIl8JGSJfCRcg/gdd6v/JIXMvbhhRL64YZvBuGE5wrhhk8K4YVDHuGFYx7hhYMe4YWjHuGFwx7hheMe4YffHuGEyyLhhQMi4YWDJuGHZybhhMsq4Yc3KuGHNy7hhn824YQLPuGFXz7hhetC4YfbQuGHQ07hh9NO4YZHVuGFZ1bhhmdW4YYLBuGFwvbhhi1UIhdKJ13URx0UADAAAALgMAAAA6WL///+D/g93K4tEJGiFwA+EwQEAAP9MJGiJ8YtEJHCDxggPthj/" & _
    "RCRw0+MBXCRgg/4PdtWD5wJ0DoF8JGAfiwAAD4QbAQAAx0UQAAAAAItFIIXAdArHQDD/////i1UI9sIBdDwPtlwkYL+FEEIIi1QkYMHjCMHqCAHTidiJ2ffnKdHR6YlUJCSJRCQgi0QkJAHBwekEic/B5wUpzzn7dCeLnCSQAAAAx0MYYLu4YZCNtCYAAAAAx0UAGwAAALgbAAAA6Z/+//+LTCRgg+EPg/kIdBCLhCSQAAAAx0AYd7u4YevTwWwkYASD7gSLVCRgg+IPjXoIO30kdhCLvCSQAAAAx0cYkru4Yeusifm6AQAAANPiiVUUMdvHRCQIAAAAAMdEJAQAAAAAxwQkAAAAAOjyUf//90QkYAACAACJRRiLtCSQAAAAx0QkYAAAAAAPlMOJRjAx9o1EGwmJRQDpBP7//8dEJAgAAAAAjXwkdDH2x0QkBAAAAADHBCQAAAAA6PJ1//+JRRiLTCR0x0QkCAIAAACJfCQEgeEA////g8kfiUwkdMZE"
Private Const STR_ZLIB_13 As String = _
    "JHWLi1UYiRQk6MJ1//+JRRi4AQAAAMdEJGAAAAAAx0UAAQAAAOmW/f//jbYAAAAAi4QkkAAAAIt8JGSLXCRoi0wkbItUJHCJeBCLfCRgiVgEiXU8i3UoiUgMiRCF9ol9OHUag30AFw+HHwEAAIu0JJAAAACLRhA5RCRcdCmLXCRci4QkkAAAAIlcJASJBCTojBoAAIXAD4XfAAAAi7wkkAAAAItHEClEJFyLjCSQAAAAi1wkKIt1CIt5BCn7AVkIi3wkXAF5FAF9HIX2dDKF/3Qui1UQhdJ0eol8JAiLcQwp/ol0JASLTRiJDCTo33T//4lFGIuMJJAAAACJQTCJ9ot9BItFPIX/dAODwECDfQALdDuLrCSQAAAAhduJRSx1CItcJFyF23QTg7wklAAAAAR0CYtEJFjpOfz//4t0JFiF9nXvx0QkWPv////r5YPogOvAi0QkXIuUJJAAAACJRCQIi3oMKceJfCQEi3UYiTQk6ApQ///pdv///8dFABwA" & _
    "AAC4/P///+np+///i5QkkAAAAItCEOkL////g/4PdyyLTCRohckPhIT+////TCRoifGLXCRwg8YID7Y7Q4lcJHDT5wF8JGCD/g921ItEJGA8CIlFEHQTi5QkkAAAAMdCGHe7uGHpJv3//4tNEIHhAOD//2aFyXQTi5wkkAAAAMdDGKa7uGHpBf3//4tVIIXSdAyLdCRgwe4Ig+YBiTL2RRECD4XtBAAAx0QkYAAAAAAx9sdFAAIAAACD/h93LItcJGiF2w+E5v3///9MJGiJ8Yt8JHCDxggPthdHiXwkcNPiAVQkYIP+H3bUi0UghcB0B4t0JGCJcAT2RRECD4U9BAAAx0QkYAAAAAAx9sdFAAMAAACD/g93NJCNtCYAAAAAi3wkaIX/D4SE/f///0wkaInxi1QkcIPGCA+2GkKJVCRw0+MBXCRgg/4PdtSLVSCF0nQSD7ZMJGCLdCRgwe4IiUoIiXIMi1UQ9sYCD4WRAwAAx0QkYAAAAAAx9sdFAAQA" & _
    "AAD2xgQPhGADAACD/g93K4tMJGiFyQ+EFP3///9MJGiJ8YtcJHCDxggPtjv/RCRw0+cBfCRgg/4PdtWLRSCLXCRghcCJXUB0BolYFItVEPbGAg+F1gIAAMdEJGAAAAAAMfbHRQAFAAAA9sYEdHuLXUA7XCRoidh2BItcJGiF23Rdi00ghcl0PYt5EIX/iXwkVHQyi3kUi1EYidkpx40EOznQdgSJ0Sn5iUwkCItEJFQB+IkEJIt8JHCJfCQE6E5HAACLVRD2xgIPhUECAAApXCRoi0VAAVwkcCnYiUVAhcAPhUr8//+LVRDHRUAAAAAAx0UABgAAAPbGCA+E+AEAAItcJGiF2w+EJPz//zHbifaLTCRwi0UgD7Y8C0OFwHQgi1AchdJ0GYtNQDtIIHMRifiIBBH/RUCJ9o28JwAAAACF/3QGO1wkaHLG9kURAg+FhQEAAClcJGgBXCRwhf8Phcz7//+LVRDHRUAAAAAAx0UABwAAAPbGEA+ERAEAAItc" & _
    "JGiF2w+Epvv//zHbjXQmAItUJHCLRSAPtjwTQ4XAdCCLUCSF0nQZi01AO0gocxGJ+IgEEf9FQIn2jbwnAAAAAIX/dAY7XCRocsb2RRECD4XPAAAAKVwkaAFcJHCF/w+FTPv//4tVEMdFAAgAAAD2xgJ0X4P+D3czkI20JgAAAACLfCRohf8PhCT7////TCRoifGLRCRwg8YID7YY/0QkcNPjAVwkYIP+D3bVD7dNGDlMJGB0E4uUJJAAAADHQhi/u7hh6cj5///HRCRgAAAAADH2i0UghcB0EMdAMAEAAADB+gmD4gGJUCzHRCQIAAAAAMdEJAQAAAAAxwQkAAAAAOhLcP//iUUYi5wkkAAAAIlDMMdFAAsAAAC4CwAAAOkd+P//iVwkCItEJHCJRCQEi00YiQwk6BZw//+JRRjpEv///4tFIIXAD4Qa////x0AkAAAAAOkO////iVwkCItEJHCJRCQEi1UYiRQk6OBv//+JRRjpXP7//4tFIIXAD4Rk" & _
    "/v//x0AcAAAAAOlY/v//iVwkCItEJHCJRCQEi1UYiRQk6Kpv//+JRRjpoP3//w+2TCRgjVQkdMFsJGAIiVQkBMdEJAgCAAAAD7ZEJGCITCR0iEQkdYt1GIk0JOhwb///iUUYi1UQ6e38//+LRSCFwA+E7Pz//8dAEAAAAADp4Pz//w+2RCRgjVQkdMFsJGAIiVQkBMdEJAgCAAAAD7ZcJGCIRCR0iFwkdYt9GIk8JOgcb///iUUYi1UQ6TL8//8PtlwkYMFsJGAIx0QkCAQAAAAPtkwkYMFsJGAIiFwkdI1cJHQPtlQkYMFsJGAIiEwkdYlcJAQPtkQkYIhUJHaIRCR3i00YiQwk6MNu//+JRRjpbfv//w+2TCRgjVQkdMFsJGAIx0QkCAIAAACJVCQED7ZEJGCITCR0iEQkdYt9GIk8JOiJbv//iUUY6dn6//+Qi1UQ6ar7//+LVRDpCfz//4tVEOmP/P//i1UQ6QX9//+LVRDpdv3//4P+H3csi3wk" & _
    "aIX/D4Sn+P///0wkaInxi1wkcIPGCA+2E0OJXCRw0+IBVCRgg/4fdtTHRQAKAAAAi1wkYItMJGCLfCRgwesIi7QkkAAAAMFkJGAYgeMA/wAAwekYgecA/wAAAdmLRCRgx0QkYAAAAADB5wgB+QHBiU0YiU4wMfaLVQyF0g+EIgEAAMdEJAgAAAAAx0QkBAAAAADHBCQAAAAA6GJJ//+JRRiLjCSQAAAAx0UACwAAAIlBMIO8JJQAAAAFD4Tw9///i30Ehf90HMdFABgAAACJ8bgYAAAAg+EHKc7TbCRg6V31//+D/gJ3NJCNtCYAAAAAi1wkaIXbD4S09////0wkaInxi3wkcIPGCA+2F0eJfCRw0+IBVCRgg/4CdtSLTCRgTtFsJGCD4QGLRCRgiU0Eg+ADg/gBdEiD+AFyOoP4AnQsg/gDdBDBbCRgAoPuAotFAOnq9P//x0UAGwAAAIuEJJAAAADHQBjTu7hh69nHRQAPAAAA69DHRQANAAAA68fH" & _
    "RUzgsrhhx0VUCQAAAMdFUOC6uGHHRVgFAAAAx0UAEgAAAOuii4QkkAAAAItcJGSLTCRsi3wkcItUJGiJWBCLXCRgiVAEiUgMiTi4AgAAAIldOIl1POkb9P//jbYAAAAAifGD4QcpztNsJGCD/h93LItcJGiF2w+EtPb///9MJGiJ8YtUJHCDxggPtjpCiVQkcNPnAXwkYIP+H3bUD7dUJGCLTCRgwekQgfH//wAAOcp0E4uEJJAAAADHQBjmu7hh6Uv1//+JVUAx9sdEJGAAAAAAx0UADgAAAItdQIXbD4Su+///O1wkaHYEi1wkaDtcJGR2BItcJGSF2w+EMPb//4lcJAiLfCRwi1QkbIl8JAQB34kUJOj2QAAAKVwkaClcJGQBXCRsKV1AiXwkcOmc/v//g/4NdyyLfCRohf8PhO31////TCRoifGLXCRwg8YID7YTQ4lcJHDT4gFUJGCD/g121It8JGCD7g7BbCRgBYPnH4tUJGDBbCRgBYHHAQEA" & _
    "AIl9YIPiH4tcJGDBbCRgBI1KAYPjD4lNZIPDBIH/HgEAAIldXHcFg/kedhOLjCSQAAAAx0EYILy4YelR9P//x0VoAAAAAMdFABAAAACLRWiLXVyJRCQsOdhzXoP+Ancsi3wkaIX/D4RE9f///0wkaInxi3wkcIPGCA+2F0eJfCRw0+IBVCRgg/4CdtSLfWiD7gOLVCRgwWwkYAONRwGD4gcPt4w/oLK4YTnYiUVoZolUTXBypolEJCyLVCQsg/oSdxwPt5wSoLK4YY1CAYP4EonCZsdEXXAAAHbniUVox0VUBwAAAI29MAUAAI1VVIl9bI1NbI2d8AIAAIl9TI19cIlcJBSJVCQQiUwkDMdEJAgTAAAAiXwkBMcEJAAAAADonCwAAIlEJFiFwHQTi4QkkAAAAMdAGES8uGHpUfP//8dFaAAAAADHRQARAAAAi01gi11ki0VoiUwkPInPAd+JXCQ4OfiJRCQsD4PpAAAAi01UugEAAACLRUzT4kqJRCQ0" & _
    "iVQkUI22AAAAAI2/AAAAAItcJFCLTCRgi3wkNCHZixyPD7bPD7b5Ofd2KItUJGiF0g+E+fP///9MJGiJ8YtEJHCDxggPtjj/RCRw0+cBfCRg672J2MHoEGaD+A8Ph7IFAAAPttcPtso5znM1jXQmAIt8JGiF/w+EtPP///9MJGiJ8YtUJHCDxggPtjpCiVQkcA+219PnAXwkYA+2yjnOcs8Pts+LRCQswesQ02wkYA+2ySnOZolcRXCJw0OJXWiJXCQsi3wkPItEJDgBxzl8JCwPgjn///+LRQCD+BsPhN3w///HRVQJAAAAjb0wBQAAjV1UiX1sjVVwiX1Mjb3wAgAAiVwkEI1dbIl8JBSJXCQMi01giVQkBMcEJAEAAACJTCQI6BgrAACJRCRYhcB0E4uEJJAAAADHQBhdvLhh6c3x///HRVgGAAAAi1VsiVVQiXwkFI19WIl8JBCJXCQMi11kiVwkCItNYMcEJAIAAACNVE1wiVQkBOjDKgAAiUQk" & _
    "WIXAdBOLhCSQAAAAx0AYeby4Yel48f//x0UAEgAAAIN8JGgFdg6BfCRkAQEAAA+H1wMAAItNVLsBAAAAi31M0+ONU/+JfCQ0kI20JgAAAACLTCRgi3wkNCHRixyPD7bHD7bIOfF2KYtcJGiF2w+EPfL///9MJGiJ8Yt8JHCDxggPth9HiXwkcNPjAVwkYOvAhNsPhJIAAAD2w/APhYkAAACJ34n5wesQiVwkTA+20YlUJEiJ+rgBAAAAD7beD7bLi1wkSAHZ0+CLXCRMi0wkYEghyA+2ztPoi0wkNAHYixyBD7bGD7bID7bXD7bSAdE58XYpi0QkaIXAD4Sx8f///0wkaInxi1QkcIPGCA+2GkKJVCRw0+MBXCRg65eJ+g+2zg+2+dNsJGAp/g+2zw+2+Sn+02wkYInZwekQiU1AhNt1EcdFABcAAAC4FwAAAOnt7v//D7bDqCAPhbT2//+oQHQTi4QkkAAAAMdAGI+8uGHpG/D//8dFABMAAACD4A+J"
Private Const STR_ZLIB_14 As String = _
    "RUiLXUiF23RHOd5zKot8JGiF/w+EF/H///9MJGiJ8YtEJHCDxggPthD/RCRw0+IBVCRgOd5y1ojZi0QkYL8BAAAA02wkYCne0+dPIccBfUDHRQAUAAAAi01YuwEAAACLVVDT44lUJDCNU/+QjXQmAIt8JGCLTCQwIdeLHLkPtscPtsg58XYoi1wkaIXbD4Sd8P///0wkaInxi0QkcIPGCA+2OP9EJHDT5wF8JGDrwfbD8A+FiQAAAInfifnB6xCJXCRED7bRiVQkQIn6uAEAAAAPtt4PtsuLXCRAAdnT4ItcJESLTCRgSCHID7bO0+iLTCQwAdiLHIEPtsYPtsgPttcPttIB0TnxdimLVCRohdIPhBrw////TCRoifGLVCRwg8YID7YaQolUJHDT4wFcJGDrl4n6D7bOD7b502wkYCn+D7bPD7bD02wkYA+2ySnOqEB0E4uEJJAAAADHQBirvLhh6anu///HRQAVAAAAwesQg+APiV1EiUVIi11Ihdt0" & _
    "RznecyqLfCRohf8PhJ/v////TCRoifGLRCRwg8YID7YQ/0QkcNPiAVQkYDnectaI2YtEJGC6AQAAANNsJGAp3tPiSiHCAVVEi1wkXIt9LItMJGQB+ynLOV1EdhOLlCSQAAAAx0IYwby4Yekh7v//x0UAFgAAAItcJGSF2w+ELu///4tcJFyLTCRki0VEKcs52HZrKdiJw4tFMDnDicJ2BYtVKCnDi300KdqLRUAB+jnDdgKJwztcJGR2BItcJGQpXCRkKdiJRUCQjbQmAAAAAA+2AkKLTCRsiAFBS4lMJGx17otVQIXSD4Vh9///x0UAEgAAALgSAAAA6ULs//+LVCRsi11AKcKJ2Ouqi5wkkAAAAItEJGyLVCRki3wkaItMJHCJQwyLRCRgiVMQi1QkXIl7BIu8JJAAAACJC4lFOIl1PIlUJASJPCTooiwAAIucJJAAAACLtCSQAAAAi5QkkAAAAItLDItGEIucJJAAAACJTCRsi3U4izqJRCRki0sE" & _
    "iXQkYIt1PIl8JHCJTCRo6br2//9mg/gQD4QuAQAAZoP4EQ+EvgAAAA+2zw+2+YPHBzn+czSLRCRohcAPhO/t////TCRoifGLVCRwg8YID7Y6QolUJHDT5w+2zwF8JGAPttGDwgc51nLMD7bPMf8PttnTbCRgKd6D7geLTCRgwWwkYAeD4X+NWQuLVCQsi0QkPItMJDgB2gHIOcJ3L0uD+/8PhBz6//+J9o28JwAAAACLTCQsS2aJfE1wQYP7/4lNaIlMJCx16en3+f//i5QkkAAAAMdCGN+8uGHHRQAbAAAA6fH5//8PttcPtvqDxwM5/nM0i0wkaIXJD4Qx7f///0wkaInxi1QkcIPGCA+2OkKJVCRwD7bX0+cBfCRgD7b6g8cDOf5yzA+2zzH/D7bZ02wkYCneg+4Di0wkYMFsJGADg+EHjVkD6T3///8Pts8PtvmDxwI5/nM0i1QkaIXSD4TL7P///0wkaInxi3wkcIPGCA+2F0eJfCRw0+IPts8B" & _
    "VCRgD7b5g8cCOf5yzA+2z4tcJCzTbCRgD7bJKc6F23Qai0QkLIPuAotMJGDBbCRgAg+3fEVug+ED64mLjCSQAAAAx0EY37y4YekM////ifaLRCRkhcAPhFTs//8Ptk1Ai1wkbIgLQ/9MJGSJXCRs6Xn9//+LVQiF0g+E6gAAAIP+H3cri3wkaIX/D4Qg7P///0wkaInxi0QkcIPGCA+2EP9EJHDT4gFUJGCD/h921YtMJGSLnCSQAAAAKUwkXItEJFwBQxQBRRyFwHQ1i30Qhf8PhBgBAACJRCQIi1wkbItEJFwpw4lcJASLVRiJFCToXGH//4lFGIuUJJAAAACJQjCLfRCLRCRki00Yhf+JRCRcD4XPAAAAi1QkYItcJGCLfCRgweoIgeIA/wAAwesYAdOB5wD/AACLVCRgwecIAfvB4hgB0znLdBOLjCSQAAAAx0EY+by4Yek06v//x0QkYAAAAACLVQgx9sdFABkAAACF0nRei0UQhcB0V4P+H3cs" & _
    "i3wkaIX/D4Qk6////0wkaInxi1wkcIPGCA+2E0OJXCRw0+IBVCRgg/4fdtSLTCRgO00cdBOLhCSQAAAAx0AYDr24YenI6f//x0QkYAAAAAAx9sdFABoAAADHRCRYAQAAAOnK6v//OUwkYOlY////i1wkXIt8JGyJXCQIKd+JfCQEi00YiQwk6PQ7///p4/7//4tVCOlX////x0QkWP3////piur//8dFAAwAAAC4DAAAAOnV5///i5wkkAAAAItDBIlEJCjptuf//422AAAAAIPsDIlcJAiLXCQQhdt0EYtDHIXAdAqLUySF0nUTjXYAuP7///+LXCQIg8QMw412AItINIXJdReJRCQEi0soiQwk/9LHQxwAAAAAMcDr14lMJASLSyiJDCT/0otDHItTJOvVjbQmAAAAAI28JwAAAACD7ByJdCQQi3QkIIl8JBSLfCQohfaJbCQYi2wkJIlcJAwPhNkAAACLXhyF2w+EzgAAAItTCIXSD4TKAAAAgzsK" & _
    "D4W6AAAAx0QkCAAAAADHRCQEAAAAAMcEJAAAAADo3jr//4kEJIl8JAiJbCQE6M46//87QxgPhX4AAACLVhCJNCSJVCQE6FYEAACFwHVei0soOc92O4tzNI0ELynIiUwkCIlEJASJNCToIzQAAItLKIlLLMdDDAEAAAAxwItcJAyLdCQQi3wkFItsJBiDxBzDi1M0iXwkCIlsJAQB0Sn5iQwk6OkzAACJeyzrx8cDHAAAALj8////68O4/f///+u8uP7////rtYM7Cg+Fa////+kx////jXQmAI28JwAAAACLRCQEi1QkCIXAdBqLQByFwHQT9kAIAnQNiVAgMcDHQjAAAAAAw7j+////w410JgBVV1ZTg+wci2wkMIXtdAeLdRyF9nUNuP7///+DxBxbXl9dw4t9BIX/dQqDfjwHD4Z/AQAAgz4dD4RqAQAAxwYdAAAAi1Y8Mf+J0YPhBynK02Y4g/oHiVY8diGNtgAAAAAPtk44iEw8GItWPEfBbjgI" & _
    "g+oIg/oHiVY8d+XHRmgAAAAAjV5oMdKJXCQUiVwkDDHbOfpzRI10JgCNvCcAAAAAD7ZMHBiD+gEPtsEPhu8AAAA9/wAAAA+E2AAAALgEAAAAKdCJwjHAhMkPlcBIIcJDOftzBYP6A3bHi3wkFIkXi30Ei00Ai1ZoiUwkEDHJOflzToP6A3dJifaNvCcAAAAAi0QkEIP6AQ+2HAEPtsMPhnwAAAA9/wAAAHRtuAQAAAAp0InCMcCE2w+VwEghwkE5+XMNg/oDdsiQjbQmAAAAAIt8JAyLXCQQiReLfQgByylNBIldAAHPg35oBIl9CHUei10UiSwk6GXX//+JfQgxwIldFMcGCwAAAOmd/v//uP3////pk/7//0Lro5CNdCYAhcDrhULpNf///422AAAAAIXA6Q////+NVmiJVCQM6S////+4+////+le/v//jXYAi0QkBIXAdAeLUByF0nUGuP7////DMcCDOg10BMONdgCLSjyFyXX1uAEAAADr7on2" & _
    "g+wsiWwkKItsJDCJdCQgi3QkNIXtiVwkHIl8JCR0IYX2dB2LXhyF23QWi1YghdJ0D4tGJIXAdSGQjbQmAAAAALj+////i1wkHIt0JCCLfCQki2wkKIPELMPHRCQIMCUAAMdEJAQBAAAAi34oiTwk/9KJRCQYi0wkGLj8////hcl0w8dEJBQAAAAAi0M0hcAPhbcAAAD8uQ4AAACJ7/OluDAlAACJ3ot8JBj2RCQYBHQTizO4LCUAAIk3i3wkGI1zBIPHBPyJwcHpAvOljYswBQAAi1NMOcpyI427LCUAADn6dxmLTCQYKdqLdCQYAdGLU1CJTkwp2gHyiVZQi1Nsi0wkGItEJBSLfCQYKdoByoXAiVdsdCGLSyS/AQAAAItzNItcJBTT54l8JAiJdCQEiRwk6FkwAACLTCQYMcCLVCQUiU0ciVE06fn+///HRCQIAQAAAL8BAAAAi0sk0+eJfCQEi1YoiRQk/1YgiUQkFIXAD4Ue////i2wkGIlsJASL" & _
    "XiiJHCT/ViS4/P///+mz/v//jbQmAAAAAI28JwAAAABVV1ZTg+wMi1QkIItyHItuNIXtD4TpAAAAi04ohcl1HcdGMAAAAACLTiS/AQAAAMdGLAAAAADT54n5iX4oi1QkJInTi1QkIItCECnDOcsPg5AAAACLVjCJzynXOd92Aonfi0wkIItBDI0MKol8JAgp2IlEJASJDCTocy8AACn7dT+LXjCLVigB+4leMDnTdCCLRiw50HMMAfiNtCYAAAAAiUYsMdKDxAyJ0FteX13DkMdGMAAAAADr1420JgAAAACLRCQgi240i3gMiVwkCIksJCnfiXwkBOgULwAAiV4wi0Yo67yLWgyJTCQIiSwkKcuJXCQE6PcuAADHRjAAAAAA693HRCQIAQAAALsBAAAAi04k0+OLTCQgiVwkBItRKIkUJP9RILoBAAAAhcCJRjQPhHD///+Jxend/v//kJCQkIPsHItEJCyJXCQUi1wkIIXAiXQkGIt0JCh0BYA4MXQR" & _
    "uPr///+LXCQUi3QkGIPEHMODfCQwOHXohdt0EoX2dA6DfCQkB34Hg3wkJA9+B7j+////69DHQxgAAAAAi1MghdJ1E8dDIBCwuGG6ELC4YcdDKAAAAACLSySFyXUHx0MkMLC4YcdEJAgwJQAAx0QkBAEAAACLSyiJDCT/0onChdK4/P///w+Eev///8dCFACAAAAxwItMJCSJUxy7AQAAAIlKJA+2TCQkiXI0x0IwAAAAAMdCLAAAAADT44laKOlC////kJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJAQABEAEgAAAAgABwAJAAYACgAFAAsABAAMAAMADQACAA4AAQAPAJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQYAcAAAAIUAAACBAAFAhzABIHHwAACHAAAAgwAAAJwAAQBwoAAAhgAAAIIAAACaAAAAgAAAAIgAAACEAAAAngABAHBgAACFgAAAgYAAAJkAATBzsAAAh4AAAIOAAACdAA"
Private Const STR_ZLIB_15 As String = _
    "EQcRAAAIaAAACCgAAAmwAAAICAAACIgAAAhIAAAJ8AAQBwQAAAhUAAAIFAAVCOMAEwcrAAAIdAAACDQAAAnIABEHDQAACGQAAAgkAAAJqAAACAQAAAiEAAAIRAAACegAEAcIAAAIXAAACBwAAAmYABQHUwAACHwAAAg8AAAJ2AASBxcAAAhsAAAILAAACbgAAAgMAAAIjAAACEwAAAn4ABAHAwAACFIAAAgSABUIowATByMAAAhyAAAIMgAACcQAEQcLAAAIYgAACCIAAAmkAAAIAgAACIIAAAhCAAAJ5AAQBwcAAAhaAAAIGgAACZQAFAdDAAAIegAACDoAAAnUABIHEwAACGoAAAgqAAAJtAAACAoAAAiKAAAISgAACfQAEAcFAAAIVgAACBYAQAgAABMHMwAACHYAAAg2AAAJzAARBw8AAAhmAAAIJgAACawAAAgGAAAIhgAACEYAAAnsABAHCQAACF4AAAgeAAAJnAAUB2MAAAh+AAAIPgAACdwA" & _
    "EgcbAAAIbgAACC4AAAm8AAAIDgAACI4AAAhOAAAJ/ABgBwAAAAhRAAAIEQAVCIMAEgcfAAAIcQAACDEAAAnCABAHCgAACGEAAAghAAAJogAACAEAAAiBAAAIQQAACeIAEAcGAAAIWQAACBkAAAmSABMHOwAACHkAAAg5AAAJ0gARBxEAAAhpAAAIKQAACbIAAAgJAAAIiQAACEkAAAnyABAHBAAACFUAAAgVABAIAgETBysAAAh1AAAINQAACcoAEQcNAAAIZQAACCUAAAmqAAAIBQAACIUAAAhFAAAJ6gAQBwgAAAhdAAAIHQAACZoAFAdTAAAIfQAACD0AAAnaABIHFwAACG0AAAgtAAAJugAACA0AAAiNAAAITQAACfoAEAcDAAAIUwAACBMAFQjDABMHIwAACHMAAAgzAAAJxgARBwsAAAhjAAAIIwAACaYAAAgDAAAIgwAACEMAAAnmABAHBwAACFsAAAgbAAAJlgAUB0MAAAh7AAAIOwAACdYA" & _
    "EgcTAAAIawAACCsAAAm2AAAICwAACIsAAAhLAAAJ9gAQBwUAAAhXAAAIFwBACAAAEwczAAAIdwAACDcAAAnOABEHDwAACGcAAAgnAAAJrgAACAcAAAiHAAAIRwAACe4AEAcJAAAIXwAACB8AAAmeABQHYwAACH8AAAg/AAAJ3gASBxsAAAhvAAAILwAACb4AAAgPAAAIjwAACE8AAAn+AGAHAAAACFAAAAgQABQIcwASBx8AAAhwAAAIMAAACcEAEAcKAAAIYAAACCAAAAmhAAAIAAAACIAAAAhAAAAJ4QAQBwYAAAhYAAAIGAAACZEAEwc7AAAIeAAACDgAAAnRABEHEQAACGgAAAgoAAAJsQAACAgAAAiIAAAISAAACfEAEAcEAAAIVAAACBQAFQjjABMHKwAACHQAAAg0AAAJyQARBw0AAAhkAAAIJAAACakAAAgEAAAIhAAACEQAAAnpABAHCAAACFwAAAgcAAAJmQAUB1MAAAh8AAAIPAAACdkA" & _
    "EgcXAAAIbAAACCwAAAm5AAAIDAAACIwAAAhMAAAJ+QAQBwMAAAhSAAAIEgAVCKMAEwcjAAAIcgAACDIAAAnFABEHCwAACGIAAAgiAAAJpQAACAIAAAiCAAAIQgAACeUAEAcHAAAIWgAACBoAAAmVABQHQwAACHoAAAg6AAAJ1QASBxMAAAhqAAAIKgAACbUAAAgKAAAIigAACEoAAAn1ABAHBQAACFYAAAgWAEAIAAATBzMAAAh2AAAINgAACc0AEQcPAAAIZgAACCYAAAmtAAAIBgAACIYAAAhGAAAJ7QAQBwkAAAheAAAIHgAACZ0AFAdjAAAIfgAACD4AAAndABIHGwAACG4AAAguAAAJvQAACA4AAAiOAAAITgAACf0AYAcAAAAIUQAACBEAFQiDABIHHwAACHEAAAgxAAAJwwAQBwoAAAhhAAAIIQAACaMAAAgBAAAIgQAACEEAAAnjABAHBgAACFkAAAgZAAAJkwATBzsAAAh5AAAIOQAACdMA" & _
    "EQcRAAAIaQAACCkAAAmzAAAICQAACIkAAAhJAAAJ8wAQBwQAAAhVAAAIFQAQCAIBEwcrAAAIdQAACDUAAAnLABEHDQAACGUAAAglAAAJqwAACAUAAAiFAAAIRQAACesAEAcIAAAIXQAACB0AAAmbABQHUwAACH0AAAg9AAAJ2wASBxcAAAhtAAAILQAACbsAAAgNAAAIjQAACE0AAAn7ABAHAwAACFMAAAgTABUIwwATByMAAAhzAAAIMwAACccAEQcLAAAIYwAACCMAAAmnAAAIAwAACIMAAAhDAAAJ5wAQBwcAAAhbAAAIGwAACZcAFAdDAAAIewAACDsAAAnXABIHEwAACGsAAAgrAAAJtwAACAsAAAiLAAAISwAACfcAEAcFAAAIVwAACBcAQAgAABMHMwAACHcAAAg3AAAJzwARBw8AAAhnAAAIJwAACa8AAAgHAAAIhwAACEcAAAnvABAHCQAACF8AAAgfAAAJnwAUB2MAAAh/AAAIPwAACd8A" & _
    "EgcbAAAIbwAACC8AAAm/AAAIDwAACI8AAAhPAAAJ/wAQBQEAFwUBARMFEQAbBQEQEQUFABkFAQQVBUEAHQUBQBAFAwAYBQECFAUhABwFASASBQkAGgUBCBYFgQBABQAAEAUCABcFgQETBRkAGwUBGBEFBwAZBQEGFQVhAB0FAWAQBQQAGAUBAxQFMQAcBQEwEgUNABoFAQwWBcEAQAUAAGludmFsaWQgYmxvY2sgdHlwZQBpbnZhbGlkIHN0b3JlZCBibG9jayBsZW5ndGhzAJCQkJCQkJCQkJCQkJCQkJB0b28gbWFueSBsZW5ndGggb3IgZGlzdGFuY2Ugc3ltYm9scwBpbnZhbGlkIGNvZGUgbGVuZ3RocyBzZXQAaW52YWxpZCBsaXRlcmFsL2xlbmd0aHMgc2V0AGludmFsaWQgZGlzdGFuY2VzIHNldABpbnZhbGlkIGxpdGVyYWwvbGVuZ3RoIGNvZGUAaW52YWxpZCBkaXN0YW5jZSBjb2Rl" & _
    "AGludmFsaWQgZGlzdGFuY2UgdG9vIGZhciBiYWNrAGludmFsaWQgYml0IGxlbmd0aCByZXBlYXQAjbQmAAAAAFVXVlOD7FyLVCRwhdJ0C4tUJHCLehyF/3UNuP7///+DxFxbXl9dw8cHCwAAAItUJHDHRwQAAAAAx0csAAAAAIsKx0IYAAAAAMdEJFAAAAAAhcmJTCRYdAeLQgSJRCRQx0QkSAAAAACLdzQx7YtfKIl0JFSJXCRMiweD6AuD+BAPh3sOAAD/JIXM57hhEOi4YUD2uGE66bhhQPa4YbjquGFA9rhhQPa4YR7uuGFA9rhhQPa4YUD2uGFA9rhhQPa4YUD2uGFA9rhhTfa4YY72uGGLVwSF0nQYxwcaAAAAiem4GgAAAIPhBynN02wkSOuKg/0CdziNtgAAAACNvwAAAACLTCRQhckPhJ4AAAD/TCRQiemLdCRYg8UID7YWRol0JFjT4gFUJEiD/QJ21ItMJEhN0WwkSIPhAYtEJEiJTwSD" & _
    "4AOD+AF0P4P4AXIyg/gCdCWD+AN0DcFsJEgCg+0C6RT////HBxsAAACLXCRwx0MYIOa4YevgxwcPAAAA69jHBw0AAADr0MdHTKDduGHHR1QJAAAAx0dQoOW4YcdHWAUAAADHBxIAAADrrItEJHiNXCRYiVwkBIkEJP9UJHSJRCRQhcAPhUP////HRCRYAAAAAMdEJDz7////jbQmAAAAAIt0JHCLRCRQi0wkWIlGBItEJDyJDukl/v//iemD4QcpzdNsJEiD/R93Mo22AAAAAIt0JFCF9g+EOAEAAP9MJFCJ6YtcJFiDxQgPtjNDiVwkWNPmAXQkSIP9H3bUD7dUJEiLTCRIwekQgfH//wAAOcp0G4tUJHDHQhgz5rhhxwcbAAAAuBsAAADpDP7//4lXQDHthdLHRCRIAAAAAHRhifaLdCRQi19AhfYPhKEAAACLVCRMhdJ0WDtcJFB2BItcJFA7XCRMdgSLXCRMiVwkCItUJFiLRCRUiVQkBIkEJOgN" & _
    "IQAAKVwkUIt3QAFcJFgpXCRMKd4BXCRUhfaJd0B1occHCwAAALgLAAAA6Yr9//+LTyiLVzSLtCSAAAAAiU8siVQkVIlMJEyJTCQIiVQkBIk0JP9UJHzHRCQ8+////4XAD4Rx////6bX+//+QjXQmAItEJHiNTCRYiUwkBIkEJP9UJHSJRCRQhcAPhUD////pdf7//4tEJHiNVCRYiVQkBIkEJP9UJHSJRCRQhcAPhan+///pUf7//4P9DXcvjXYAi1wkUIXbD4RMCwAA/0wkUInpi1wkWIPFCA+2E0OJXCRY0+IBVCRIg/0NdtSLdCRIg+0OwWwkSAWD5h+LRCRIwWwkSAWBxgEBAACJd2CD4B+LVCRIwWwkSASNSAGD4g+JT2SDwgSB/h4BAACJV1x3BYP5HnYQi0wkcMdBGGDmuGHpWP7//8dHaAAAAACF0nRkg/0CdzeNdCYAjbwnAAAAAItUJFCF0g+EiAoAAP9MJFCJ6YtcJFiDxQgPthNDiVwk" & _
    "WNPiAVQkSIP9AnbUi3dog+0Di0wkSMFsJEgDD7eENmDduGFGg+EHO3dciXdoZolMR3BynItXaIP6EnccD7ecEmDduGGNQgGD+BKJwmbHRF9wAAB254lHaMdHVAcAAACNnzAFAACNj/ACAACJX2yNR1SNV2yJX0yNd3CJXCQ4iUwkNIlMJBSJRCQwiUQkEIlUJCyJVCQMx0QkCBMAAACJdCQoiXQkBMcEJAAAAADoDgwAAIlEJDyFwHQQi3QkcMdGGITmuGHpU/3//8dHaAAAAACLX2SLd2AB8zlcJDwPg/wAAACLV1SLd0yJVCQYjXYAD7ZMJBi4AQAAAItcJEjT4Egh2Ischg+21w+2yjnpdlCLTCRQhcl0Hf9MJFCJ6YtcJFiDxQgPthNDiVwkWNPiAVQkSOu7i0QkeI10JFiJdCQEiQQk/1QkdIlEJFCFwA+ENfz//4tXVIt3TIlUJBjruInYwegQZoP4Dw+H5wYAAA+21w+2yjnNczGLdCRQhfYP"
Private Const STR_ZLIB_16 As String = _
    "hK0GAAD/TCRQiemLdCRYg8UID7YWRol0JFjT4g+2zwFUJEgPttE51XLPi3doD7bPwesQ02wkSA+2ySnNZolcd3BGi19giXdoifGLd2SJXCQci1wkHAHzOdkPggT///+LB4P4Gw+ET/r//8dHVAkAAACLdCQ4i1wkNItEJDCLVCQsiXdsiXdMi3QkKIlcJBSJRCQQiVQkDItPYIl0JATHBCQBAAAAiUwkCOiSCgAAhcB0EItMJHDHQRid5rhh6dv7///HR1gGAAAAi19sjVdYi0QkNItMJCyJX1CJVCQQiUQkFIlMJAyLd2SJdCQIi19gxwQkAgAAAI1UX3CJVCQE6DwKAACFwHQQi3QkcMdGGLnmuGHphfv//8cHEgAAAIN8JFAFD4aRAAAAgXwkTAEBAAAPhoMAAACLXCRwi1QkUItEJEyLTCRUiVMEiUMQi1coiUsMi3QkWIlvPItMJEg5VyyJM4lPOHMLi1wkTInVKd2JbyyJVCQEi2wkcIksJOjt" & _
    "DwAAi1QkcItsJHCLXCRwi3IMi00Ai1UEiXQkVItDEIt3OIlMJFiLbzyJRCRMiVQkUIl0JEjp/fj//4tHVIt3TIlEJBiNtgAAAACNvwAAAAAPtkwkGLgBAAAAi1wkSNPgSCHYixyGD7bXD7bKOel2UItEJFCFwHQd/0wkUInpi1wkWIPFCA+2E0OJXCRY0+IBVCRI67uLdCR4jUwkWIlMJASJNCT/VCR0iUQkUIXAD4TV+f//i19Ui3dMiVwkGOu4hNsPhKsAAAD2w/APhaIAAACJXCREidnB6RCJTCQkifaLXCREuAEAAAAPttcPtsoPttMB0YtUJEjT4EgPts+LXCQkIdDT6A+2yQHYixyGD7bXD7bSAdE56XZJi0wkUIXJdB3/TCRQiemLXCRYg8UID7YTQ4lcJFjT4gFUJEjroYtEJHiNdCRYiXQkBIkEJP9UJHSJRCRQhcAPhCv5//+Ld0zrv4tUJEQPts7TbCRID7bJKc0Pts+J3sHuENNsJEgP" & _
    "tskpzYl3QITbdWWLXCRMhdt0Iw+2X0C4EgAAAItMJFSIGUH/TCRMiUwkVMcHEgAAAOmD9///i1coi0c0i7QkgAAAAIlXLIlEJFSJVCRMiVQkCIlEJASJNCT/VCR8x0QkPPv///+FwHSq6bL4//+J9g+2w6ggD4Wk+f//qEB0EIt0JHDHRhjP5rhh6Q75//+D4A+FwIlHSHRNOcUPgwkDAACLRCRQhcAPhNkCAAD/TCRQiemLXCRYg8UID7YzQ4lcJFjT5otPSAF0JEg5zXLSi0QkSLoBAAAAKc3TbCRI0+JKIcIBV0CLT1iLd1CJTCQYD7ZMJBi4AQAAAItcJEjT4Egh2Ischg+2zw+20TnqdlCLRCRQhcB0Hf9MJFCJ6YtcJFiDxQgPthNDiVwkWNPiAVQkSOu7i3QkeI1UJFiJVCQEiTQk/1QkdIlEJFCFwA+Etff//4tPWIt3UIlMJBjruPbD8A+FoAAAAIlcJESJ2cHpEIlMJCCLXCREuAEAAAAP" & _
    "ttcPtsoPttMB0YtUJEjT4EgPts+LXCQgIdDT6A+2yQHYixyGD7bXD7bSAdE56XZJi0QkUIXAdB3/TCRQiemLXCRYg8UID7YTQ4lcJFjT4gFUJEjroYt0JHiNTCRYiUwkBIk0JP9UJHSJRCRQhcAPhBX3//+Ld1Drv4tUJEQPts4PtvHTbCRIKfUPts8PtsPTbCRID7bJKc2oQHQQi0QkcMdAGOvmuGHpcff//4PgD8HrEIXAiV9EiUdIdE05xQ+DOgEAAItcJFCF2w+ECgEAAP9MJFCJ6YtcJFiDxQgPthNDiVwkWNPii09IAVQkSDnNctKLRCRIvgEAAAApzdNsJEjT5k4hxgF3RItXKDlXLInQcwaLTCRMKcg5R0R2EotcJHDHQxgB57hh6e/2//+J9ot0JEyF9nRmi0dEidMpwztcJExzTYt0JEyLVCRUKd4B2onzi0dAOcN2AonDKVwkTCnYiUdAjXQmAI28JwAAAAAPtgpCi3QkVIgORkuJdCRU" & _
    "de6LV0CF0g+EqvT//4tXKOuei1QkVItcJEwpwuu1iVcsi180i4QkgAAAAIlUJEyJXCRUiVQkCIlcJASJBCT/VCR8x0QkPPv///+FwA+FzvX//4tXKOle////i3QkeI1MJFiJTCQEiTQk/1QkdIlEJFCFwA+F1/7//+mL9f//i09I6ez+//+LVCR4jUwkWIlMJASJFCT/VCR0iUQkUIXAD4UI/f//6V/1//+LT0jpHf3//4tEJHiNTCRYiUwkBIkEJP9UJHSJRCRQhcAPhTT5///pM/X//2aD+BAPhHABAABmg/gRD4TdAAAAD7bPD7bxg8YHOfVzNItEJFCFwA+EoAAAAP9MJFCJ6Yt0JFiDxQgPthZGiXQkWNPiAVQkSA+21w+2yoPBBznNcszHRCRAAAAAAA+2zw+22dNsJEgp3YPtB4tMJEjBbCRIB4Phf41ZC4tPaItHYI0UC4t3ZIlEJBwB8DnCdyZLg/v/D4Tg+P//i0QkQEuNUQGJV2iD+/9m" & _
    "iURPcInRdenpxPj//4tEJHDHQBgf57hhxwcbAAAA6bz4//+LdCR4jVQkWIlUJASJNCT/VCR0iUQkUIXAD4VB////6UL0//8PttcPtsqDwQM5zXMwi3QkUIX2dFP/TCRQiemLdCRYg8UID7YWRol0JFjT4g+2zwFUJEgPttGDwgM51XLQx0QkQAAAAAAPts8PttnTbCRIKd2D7QOLTCRIwWwkSAOD4QeNWQPpIv///4tEJHiNTCRYiUwkBIkEJP9UJHSJRCRQhcB1kum68///kA+2zw+28YPGAjn1czCLVCRQhdJ0a/9MJFCJ6Yt0JFiDxQgPthZGiXQkWNPiAVQkSA+21w+2yoPBAjnNctAPts8Pttkp3dNsJEiLT2iFyXQgD7dcT26D7QKLdCRIwWwkSAKD5gOJXCRAjV4D6ZT+//+LVCRwx0IYH+e4YenI/v//i0QkeI10JFiJdCQEiQQk/1QkdIlEJFCFwA+Fdv///+kV8///i1wkeI10JFiJdCQE" & _
    "iRwk/1QkdIlEJFCFwA+FWfX//+nx8v//i0wkeI10JFiJdCQEiQwk/1QkdIlEJFCFwA+FlfT//+nN8v//jXQmAMdEJDz+////6dPy///HRCQ8AQAAAItHKDlEJEwPg77y//+LXCRMKdiJRCQIi280i7wkgAAAAIlsJASJPCT/VCR8hcAPhJfy///pg/L//8dEJDz9////6YXy//+QjXQmAIPsDIlcJAiLXCQQhdt0DotDHIXAdAeLUySF0nUTuP7///+LXCQIg8QMw422AAAAAIlEJASLSyiJDCT/0sdDHAAAAAAxwOvbkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkCBpbmZsYXRlIDEuMi4zIENvcHlyaWdodCAxOTk1LTIwMDUgTWFyayBBZGxlciAAkJCQkJCQkJCQkJCQkJCQkJADAAQABQAGAAcACAAJAAoACwANAA8AEQATABcAGwAfACMAKwAzADsAQwBTAGMAcwCDAKMAwwDjAAIBAAAAAIn2" & _
    "EAAQABAAEAAQABAAEAAQABEAEQARABEAEgASABIAEgATABMAEwATABQAFAAUABQAFQAVABUAFQAQAMkAxACJ9gEAAgADAAQABQAHAAkADQARABkAIQAxAEEAYQCBAMEAAQGBAQECAQMBBAEGAQgBDAEQARgBIAEwAUABYAAAAAAQABAAEAAQABEAEQASABIAEwATABQAFAAVABUAFgAWABcAFwAYABgAGQAZABoAGgAbABsAHAAcAB0AHQBAAEAAVVdWMfZTgeycAAAAi5wkuAAAAI22AAAAAI28JwAAAABmx0R0cAAARoP+D3bzx0QkTAAAAAA5XCRMcyqQjXQmAItMJEyLrCS0AAAA/0QkTA+3dE0AD7dUdHBCOVwkTGaJVHRwctvHRCREDwAAAIuUJMAAAACLMol0JECNtgAAAACLVCREZoN8VHAAdQdKiVQkRHXti0wkRDlMJEB2BIlMJEDHRCRIAQAAAItsJESF7XVPi4QkvAAAAIHnAP///4PP" & _
    "QIn6ixi2AYnVgeX//wAAiSuLtCS8AAAAiw6NeQSJPolpBIucJMAAAACDBgTHAwEAAAAxwIHEnAAAAFteX13DkItUJEhmg3xUcAB1CkKD+g+JVCRIduqLTCRIOUwkQHMEiUwkQLoBAAAAvgEAAACJ9g+3THRwAdIpyngcRoP+D3bvhdJ+GYu0JLAAAACF9nQHg3wkRAF0B7j/////65Zmx0QkUgAAvgEAAAAPt2x0UA+3RHRwAcVmiWx0UkaD/g526cdEJEwAAAAAOVwkTHM+i0wkTIu0JLQAAAAPtwROZoXAdCAPt+gPt3RsUI1WAWaJVGxQi1QkTIusJMQAAABmiVR1AP9EJEw5XCRMcsKLnCSwAAAAhdsPhNsDAACDvCSwAAAAAQ+EsAMAAMdEJCTA97hhx0QkIAD4uGHHRCQc/////w+2TCRAx0QkOAEAAACLrCS8AAAAx0QkNAAAAACLdCRI02QkOItFAItsJEDHRCRMAAAAAItUJDiJRCQox0Qk" & _
    "PAAAAABKg7wksAAAAAHHRCQw/////4lUJCwPhCIDAACJ8otEJEyJ+SpUJDyLnCTEAAAAiNUPtxRDic87VCQcD427AgAAweIQgecA/wAACdeLXCQ8ifG6AQAAACnZ0+KLXCQ0iVQkDInpugEAAADT4g+2TCQ8iVQkSNPrkI10JgCLTCQMKcqLTCQojQQahdKJPIF17I1W/7sBAAAAiNHT44VcJDSJXCQMdBSNtgAAAADRbCQMi0QkDIVEJDR18otcJAyF2w+EKQIAAItcJAyLTCQMSyFcJDQBTCQ0D7dMdHD/RCRMSWaFyWaJTHRwdSQ7dCRED4QPAQAAi7QkxAAAAItMJEyLnCS0AAAAD7cETg+3NEM7dCRAD4YD////i1wkNItEJCwhwztcJDAPhO/+//+LbCQ8he11CItUJECJVCQ8i2wkKItUJEiLRCQ8i1wkPI1MlQCJ9YlMJCgpxboBAAAAiemNBCvT4jtEJERzHJAPt1xEcCnahdJ+EItcJDxF"
Private Const STR_ZLIB_17 As String = _
    "AdKNBCs7RCREcuW6AQAAAInp0+IBVCQ4g7wksAAAAAF0UotEJCyLXCQ0i4wkvAAAACHDieiJXCQwi1QkMIsZiASTi4QkvAAAAIsYD7ZEJECIRJMBi5QkvAAAAItMJCiLGotUJDAp2cH5AmaJTJMC6Sz+//+BfCQ4rwUAALgBAAAAdp/prPz//4nxgecA////g89AKkwkPIn7i3QkNIjPid+B5///AACF9g+EmAAAALsBAAAAi0QkPIXAdDWLbCQ0i0wkLCHNO2wkMHQlx0QkPAAAAACLrCS8AAAAifiLdCRAimQkQItVAInHiVQkKI1W/w+2TCQ8i3QkNItsJCjT7ojRiXy1AIne0+aFdCQ0iXQkDHQO0WwkDItEJAyFRCQ0dfKLbCQMhe10P4t0JAyLTCQMTiF0JDQBTCQ0i2wkNIXtD4Vt////i1QkOIuMJLwAAACLXCRAweICi7wkwAAAAAERiR/pyfv//8dEJDQAAAAA68jHRCQ0AAAAAOnb/f//" & _
    "O1QkHH4oi0QkIIHnAP///w+2HFCLRCQkCd8PtxxQgef//wAAweMQCd/pIv3//4HnAP///4PPYIHn//8AAOkO/f//gXwkOK8FAAC4AQAAAA+HX/v//+nG/P//x0QkJD71uGHHRCQgfvW4YcdEJBwAAQAA6Uv8///HRCQcEwAAAIuMJMQAAACJTCQgiUwkJOkv/P//kJCQkJCQkJCQkJCQkGludmFsaWQgbGl0ZXJhbC9sZW5ndGggY29kZQBpbnZhbGlkIGRpc3RhbmNlIGNvZGUAaW52YWxpZCBkaXN0YW5jZSB0b28gZmFyIGJhY2sAVVdWU4PsQItcJFSLVCRUi3QkWItrHItKBItaEIlsJDyLKineTYnvAc+D7wWLTCQ8iXwkOIt6DE+J+CnwjbQ7//7//4lEJDSLXCQ8i0EsiXQkMItMJDyLUyiJRCQoi3QkPIlUJCyLUTSLXjCJVCQgi1QkPIlcJCSLQkyLWjiLcjyJRCQci0pQuAEAAACJTCQY" & _
    "i0pUicLT4onRi1QkPEmJTCQUi0pY0+BIiUQkEIP+DnccRYnxg8YID7ZVAEXT4gHTifGDxggPtlUA0+IB04tEJBQh2I22AAAAAI28JwAAAACLTCQcixSBD7bOD7bBiMEpxtPrD7bKhcmJDCQPhIQDAAD2BCQQD4XVAAAA9gQkQHUWD7YMJLgBAAAAweoQ0+BIIdiNBBDrufYEJCAPhaAAAACLTCRUx0EYIP64YYtEJDzHABsAAACJ8MHoA4nCweIDKcUp1onxugEAAADT4kqLTCRUIdM7bCQ4jVUBiRGNVwGJUQxzTSlsJDiLRCQ4g8AFO3wkMItsJFSJRQRzJil8JDCLRCQwBQEBAACLfCQ8i0wkVIlfOIlBEIl3PIPEQFteX13Di1QkMLgBAQAAKdcp+OvYi0wkOLgFAAAAKc0p6Ouvi1QkPMcCCwAAAOlm////weoQgyQkD4lUJAx0LDs0JHMORYnxg8YID7ZVANPiAdMPtgwkugEAAADT4koh2tPr" & _
    "AVQkDIsMJCnOg/4OdxxFifGDxggPtlUARdPiAdOJ8YPGCA+2VQDT4gHTi0QkECHYi0wkGIsUgQ+2zg+2wYjBKcYPtsLT66gQiQQkdSz2BCRAdRYPtgwkuAEAAADB6hDT4Egh2I0EEOvHi1QkVMdCGDz+uGHps/7//4PgD8HqEDnGiVQkCIkEJHMhRYnxg8YID7ZVANPiAdM7NCRzDkWJ8YPGCA+2VQDT4gHTD7YMJLoBAAAAiwQk0+Ipxkoh2on40+sBVCQIi0wkNCnIOUQkCIkEJA+GbgEAAItMJAgpwYtEJCiJDCQ5wQ+HSQEAAItUJCCLTCQkSolUJASFyQ+FjwAAAItEJCyLDCQpyAHCi0QkDDkEJHMdiwwkKciJRCQMQkcPtgqID/8MJHX0i0wkCIn6KcqDfCQMAnYhQkcPtgpCiA9HD7YKQogPRw+2CogPg2wkDAODfCQMAnffi0QkDIXAdBVCRw+2CogPg3wkDAF2Bw+2SgFHiA87bCQ4D4Oz" & _
    "/f//O3wkMA+CD/3//+mk/f//iwwkOUwkJHNqi0wkJItEJCwByIsMJCnIAcKLRCQkKQQki0QkDDkEJA+DdP///4sMJCnIiUQkDEJHD7YKiA//DCR19ItEJAyLVCQEOUQkJA+DTf///4tMJCQpyIkMJIlEJAxCRw+2CogP/wwkdfTpJ////4tEJCSLDCQpyAHCi0QkDDkEJA+DF////4sMJCnIiUQkDEJHD7YKiA//DCR19On1/v//i0wkVMdBGFL+uGHp5fz//4tMJAiJ+inKQkcPtgpCiA9HD7YKQogPRw+2CogPg2wkDAODfCQMAnff6ef+//9HweoQiBfp+f7//4tEJASFwHQFgzggdAa4VwAHgMPHQAQBAAAAx0AIAgAAAMdADAMAAADHQBABAAAAx0AY6wcDAMdAHAIAAQAxwMOD7CyLRCQ8iVwkHItcJDSD+AGJdCQgi3QkOIl8JCSLfCQwiWwkKHRNg/gCdDOFwI1sJBB1CIlcJBCJdCQUiWwk" & _
    "BIk8JOgnCAAAi1wkHIt0JCCLfCQki2wkKIPELMONdgCLVxCNbCQQiRQk6BEIAAABwxHW68CJPCSNbCQQiWwkBOgLCAAAAVwkEBF0JBTrrpCD7ByNVCQQiVQkBItUJCCJFCTo6QcAAIXAuv////+5/////3UIi1QkEItMJBSJ0IPEHInKw5CQkJCQkJCQkJCQVYnluAEAAABdwgwAkJCQkFWJ5YPsCKFEELlhixCF0nUEiexdw/8QixVEELlhi0oEjUIEo0QQuWGFyXXp6+OJ9lWJ5YPsCIld/KEwDLlhg/j/dCmFwInDdRfHBCQgBLlh6JsM//+LXfyJ7F3DjXQmAP8UnTAMuWFLdfbr3YsNNAy5YTHAhcl0y7owDLlhjXYAQItMggSFyXX367iQjXQmAFWJ5YPsCIld/IsVQBC5YYXSdSehMAy5YboBAAAAiRVAELlhg/j/dCWFwInDdRPHBCQgBLlh6CYM//+LXfyJ7F3D/xSdMAy5YUt19uvhiw00" & _
    "DLlhMcCFyXTPujAMuWGJ9kCLXIIEhdt19+u9kJCQkJBVuQAguWGB+QAguWGJ5XMWi1EEiwGDwQgBggAAuGGB+QAguWFy6l3DkJCQkJCQkJAtTElCR0NDVzMyLUVILVNKTEotR1RIUi1NSU5HVzMyAFWJ5YPsCKFQILlh/1AEiexdw420JgAAAACNvCcAAAAAVTHAieX8g+wIuQ4AAACJHCSLXQiJfCQEid/zq8cDOAAAAKEgILlhixVkELlhx0ME4Au5YYlDFKFgELlhiVMcx0Ms/////4lDGKEwILlhx0MIcAW5YcdDJAAAAACJQyihaBC5YYsVbBC5YYlDMIlTNIscJIt8JASJ7F3DkI10JgBVieWD7FiJXfihUCC5YYl1/IXAdAuLXfiLdfyJ7F3DkMdFuEFBQUGhUAW5YY11uMdFvEFBQUHHRcBBQUFBiUXYoVQFuWHHRcRBQUFBx0XIQUFBQYlF3KFYBblhx0XMQUFBQcdF0EFBQUGJReChXAW5" & _
    "YcdF1EFBQUGJNCSJReShYAW5YYlF6KFkBblhiUXsoWgFuWGJRfChbAW5YYlF9Og9BQAAD7fAg+wEZoXAdWLHBCQ4AAAA6OYDAACFwInDdFyJBCTouP7//4kcJOhgAAAAZoXAdBuJHVAguWGNQwSjQCC5YY1DCKNgILlh6SX///+JHCToiAMAAIk0JOjgBAAAg+wED7fAjXYAjbwnAAAAAIkEJOjIAAAAicPrueivBAAA6w2QkJCQkJCQkJCQkJCQVbgfAAAAieVWugEAAABTg+xQi3UIjXQmAI28JwAAAACF8rFBdQKxYYhMBbgB0kh576FQBblhiUXYoVQFuWGJRdyhWAW5YYlF4KFcBblhiUXkoWAFuWGJReihZAW5YYlF7KFoBblhiUXwoWwFuWGJRfSNRbiJBCToRAQAAA+32IPsBGaF23ULMdKNZfiJ0FteXcOJHCToFgAAADnwidp15+vnjbQmAAAAAI28JwAAAABVieWNVbhTg+xUMdsPt0UI" & _
    "x0QkCEAAAACJVCQEiQQk6PwDAACD7AyFwHQxuB8AAAC6AQAAAI10JgCNvCcAAAAAgHwFuEF0EQHSSHn0gzs4dQuJ2Itd/MnDCdPr6+iPAwAAkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkFWJ5VdWU4PsNItdDMdF2AAAAACLTQiLRRDHRdwAAAAAhduLVRTHReQAAAAAD4hyAQAAhdIPiF4BAACJVdSF0onGiU3Qid91aTnYdkiJyIna9/aJVdCNTeiFyXQZx0XcAAAAAItF0IlF2ItV2ItN3IlV6IlN7ItF5IXAdAr3XeiDVewA913si0Xoi1Xsg8Q0W15fXcOFwHULuAEAAAAx0vf2icaJ+ItV1Pf2i0XQ9/broTld1HYaiU3YiX3ci0XYi1XciUXoiVXs66uNtgAAAAAPvUXUicOD8x91KDt91HcFOXXQcguLVdAp8ht91IlV0I1N6IXJD4R6////iX3c6WD///+LVdS4IAAAAIjZKdiJ" & _
    "RczT4onwD7ZNzNPoiNnT5gnCi0XQD7ZNzIlV1In60+qI2dPnD7ZNzNPoCceJ+IjZ93XU02XQidf35olFxDn6dzo5+nQujUXohcAPhBX///8Ptk3Mi0XQK0XEGdeJRdCJ+tPiiNnT6AnCiVXY0+/pNv///4tF0DlFxHbKi03EKfEbVdSJTcTrvffYg9IA99rplv7//8dF5P/////32YPTAPfb6Xv+//+QkJCQkJCQkJBRieGDwQg9ABAAAHIQgekAEAAAgwkALQAQAADr6SnBgwkAieCJzIsIi0AE/+CQkJD/JfBAuWGQkAAAAAAAAAAA/yX8QLlhkJAAAAAAAAAAAP8lGEG5YZCQAAAAAAAAAAD/JTBBuWGQkAAAAAAAAAAA/yUAQblhkJAAAAAAAAAAAP8lPEG5YZCQAAAAAAAAAAD/JRBBuWGQkAAAAAAAAAAA/yVMQblhkJAAAAAAAAAAAP8lVEG5YZCQAAAAAAAAAAD/JShBuWGQkAAAAAAAAAAA"
Private Const STR_ZLIB_18 As String = _
    "/yVEQblhkJAAAAAAAAAAAP8lQEG5YZCQAAAAAAAAAAD/JRRBuWGQkAAAAAAAAAAA/yUsQblhkJAAAAAAAAAAAP8lOEG5YZCQAAAAAAAAAAD/JUhBuWGQkAAAAAAAAAAA/yUkQblhkJAAAAAAAAAAAP8lUEG5YZCQAAAAAAAAAAD/JVhBuWGQkAAAAAAAAAAA/yUgQblhkJAAAAAAAAAAAP8lCEG5YZCQAAAAAAAAAAD/JTRBuWGQkAAAAAAAAAAA/yUEQblhkJAAAAAAAAAAAP8lHEG5YZCQAAAAAAAAAAD/JQxBuWGQkAAAAAAAAAAA/yXgQLlhkJAAAAAAAAAAAP8l3EC5YZCQAAAAAAAAAAD/JeRAuWGQkAAAAAAAAAAAVYnlXelHBf//kJCQkJCQkP////8gDLlhAAAAAP////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
    "wIS4YWCDuGEBAQAAHgEAAA8AAABAibhh4IO4YQAAAAAeAAAADwAAAAAAAABghLhhAAAAABMAAAAHAAAAAAAAAAAAAABADLlhAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP////8AAAAA/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAV3beQgAAAAACMwEAAQAAAEkAAABJAAAAKDABAEwxAQBwMgEAAAMBAMCJAADAiwAAoI4AAECRAACQmwAAwI0AAHCcAACAEQAA0BMAAGAVAACgFAAAoBUAANA1AADgOwAAYGIAABBiAABwbwAAgG4AACBbAADAWgAAwGAAAHBgAABQXwAA8F0AAEBgAADQYQAAAFoAAMA1AAAwVAAAUFIAAOBPAADgPQAAsE8AACBTAACQSwAAgEgAACBJAACwPQAAkEoAACBLAABgSwAAoEUAAMBOAADgSwAAMD4AAIBPAADASAAAkEkAADC9AABA5wAAoPYAAGDcAACA2QAA" & _
    "0NUAAGDXAABgsQAAYLIAAPCwAABgsAAAQNYAAJDXAABQ2QAAAPcAAHD+AABA+AAAQFkAAPCvAACgrwAAELAAADCwAADgrwAA0K8AAAwzAQAaMwEAJTMBADIzAQA8MwEATDMBAFUzAQBmMwEAcDMBAHgzAQCIMwEAkTMBAJszAQCpMwEArzMBAL0zAQDFMwEA0jMBAN4zAQDpMwEA9zMBAAQ0AQASNAEAHzQBACw0AQBBNAEAUjQBAF40AQBwNAEAfjQBAIk0AQCRNAEAmjQBAKI0AQCoNAEAsDQBALg0AQC/NAEAxjQBAM00AQDWNAEA3TQBAOQ0AQDrNAEA9DQBAPs0AQAHNQEADjUBABc1AQAfNQEAJzUBADM1AQBCNQEAUzUBAF81AQBqNQEAezUBAIk1AQCWNQEAozUBALA1AQDFNQEA0TUBAOI1AQD0NQEAATYBAA82AQAaNgEAITYBACo2AQAyNgEAOTYBAEo2AQAAAAEAAgADAAQABQAGAAcA" & _
    "CAAJAAoACwAMAA0ADgAPABAAEQASABMAFAAVABYAFwAYABkAGgAbABwAHQAeAB8AIAAhACIAIwAkACUAJgAnACgAKQAqACsALAAtAC4ALwAwADEAMgAzADQANQA2ADcAOAA5ADoAOwA8AD0APgA/AEAAQQBCAEMARABFAEYARwBIAHpsaWIxLmRsbABEbGxHZXRWZXJzaW9uAF9kaXN0X2NvZGUAX2xlbmd0aF9jb2RlAF90cl9hbGlnbgBfdHJfZmx1c2hfYmxvY2sAX3RyX2luaXQAX3RyX3N0b3JlZF9ibG9jawBfdHJfdGFsbHkAYWRsZXIzMgBhZGxlcjMyX2NvbWJpbmUAY29tcHJlc3MAY29tcHJlc3MyAGNvbXByZXNzQm91bmQAY3JjMzIAY3JjMzJfY29tYmluZQBkZWZsYXRlAGRlZmxhdGVCb3VuZABkZWZsYXRlQ29weQBkZWZsYXRlRW5kAGRlZmxhdGVJbml0Ml8AZGVmbGF0ZUlu" & _
    "aXRfAGRlZmxhdGVQYXJhbXMAZGVmbGF0ZVByaW1lAGRlZmxhdGVSZXNldABkZWZsYXRlU2V0RGljdGlvbmFyeQBkZWZsYXRlU2V0SGVhZGVyAGRlZmxhdGVUdW5lAGRlZmxhdGVfY29weXJpZ2h0AGdldF9jcmNfdGFibGUAZ3pjbGVhcmVycgBnemNsb3NlAGd6ZGlyZWN0AGd6ZG9wZW4AZ3plb2YAZ3plcnJvcgBnemZsdXNoAGd6Z2V0YwBnemdldHMAZ3pvcGVuAGd6cHJpbnRmAGd6cHV0YwBnenB1dHMAZ3pyZWFkAGd6cmV3aW5kAGd6c2VlawBnenNldHBhcmFtcwBnenRlbGwAZ3p1bmdldGMAZ3p3cml0ZQBpbmZsYXRlAGluZmxhdGVCYWNrAGluZmxhdGVCYWNrRW5kAGluZmxhdGVCYWNrSW5pdF8AaW5mbGF0ZUNvcHkAaW5mbGF0ZUVuZABpbmZsYXRlR2V0SGVhZGVyAGluZmxh" & _
    "dGVJbml0Ml8AaW5mbGF0ZUluaXRfAGluZmxhdGVQcmltZQBpbmZsYXRlUmVzZXQAaW5mbGF0ZVNldERpY3Rpb25hcnkAaW5mbGF0ZVN5bmMAaW5mbGF0ZVN5bmNQb2ludABpbmZsYXRlX2NvcHlyaWdodABpbmZsYXRlX2Zhc3QAaW5mbGF0ZV90YWJsZQB1bmNvbXByZXNzAHpFcnJvcgB6X2Vycm1zZwB6Y2FsbG9jAHpjZnJlZQB6bGliQ29tcGlsZUZsYWdzAHpsaWJWZXJzaW9uAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFRAAQAAAAAAAAAAALxCAQDcQAEAaEABAAAAAAAAAAAA0EIBAPBAAQB0QAEAAAAAAAAAAAA8QwEA/EABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGBBAQBsQQEAeEEBAAAAAAAAAAAAiEEBAAAAAAAAAAAAlEEBAKRBAQCwQQEA" & _
    "xEEBANRBAQDcQQEA6EEBAPRBAQAAQgEADEIBABRCAQAgQgEAKEIBADBCAQA4QgEAREIBAFBCAQBcQgEAaEIBAHRCAQCAQgEAjEIBAJhCAQCkQgEAAAAAAAAAAABgQQEAbEEBAHhBAQAAAAAAAAAAAIhBAQAAAAAAAAAAAJRBAQCkQQEAsEEBAMRBAQDUQQEA3EEBAOhBAQD0QQEAAEIBAAxCAQAUQgEAIEIBAChCAQAwQgEAOEIBAERCAQBQQgEAXEIBAGhCAQB0QgEAgEIBAIxCAQCYQgEApEIBAAAAAAABAEFkZEF0b21BAACvAEZpbmRBdG9tQQDcAEdldEF0b21OYW1lQQAAGABfZmRvcGVuAAAAJABfX2RsbG9uZXhpdAAAAJgAX2Vycm5vAAAAAK0AX2ZpbGVsZW5ndGhpNjQAAAAAvQFfdnNucHJpbnRmAAAAABUCYWJvcnQAIwJjbGVhcmVycgAALQJmY2xvc2UAAAAAMAJmZmx1c2gAAAAA"
Private Const STR_ZLIB_19 As String = _
    "MgJmZ2V0cG9zAAAAOAJmb3BlbgA5AmZwcmludGYAAAA6AmZwdXRjAD4CZnJlYWQAPwJmcmVlAABEAmZzZXRwb3MAAABHAmZ3cml0ZQAAAAByAm1hbGxvYwAAAAB4Am1lbWNweQAAAAB6Am1lbXNldAAAAACTAnNwcmludGYAAACXAnN0cmNhdAAAAACbAnN0cmNweQAAAACdAnN0cmVycm9yAACfAnN0cmxlbgAAAAAAQAEAAEABAABAAQBLRVJORUwzMi5kbGwAAAAAFEABAG1zdmNydC5kbGwAAChAAQAoQAEAKEABAChAAQAoQAEAKEABAChAAQAoQAEAKEABAChAAQAoQAEAKEABAChAAQAoQAEAKEABAChAAQAoQAEAKEABAChAAQAoQAEAKEABAChAAQAoQAEAKEABAG1zdmNydC5kbGwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVnbeQgAAAAAAAAMAAwAAACgAAIAOAAAAWAAAgBAAAACIAACAAAAAAFZ23kIAAAAAAAABAAEAAABAAACAAAAAAFZ23kIAAAAAAAABAAkEAAC4AAAAAAAAAFZ23kIAAAAAAAABAPQBAABwAACAAAAAAFZ23kIAAAAAAAABAAkEAADIAAAAAAAAAFZ23kIAAAAAAAABAAEAAACgAACAAAAAAFZ23kIAAAAAAAABAAkEAADYAAAA6FABAOgCAAAAAAAAAAAAANBTAQAUAAAAAAAAAAAAAADkUwEAdAwAAAAAAAAAAAAAKAAAACAAAABAAAAAAQAEAAAAAACAAgAA" & _
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAIAAAACAgACAAAAAgACAAICAAADAwMAAgICAAAAA/wAA/wAAAP//AP8AAAD/AP8A//8AAP///wD//wAP//D///D////w/////wAAAP/wL//wj///AA////IiIiD/8gD/8gj/8AIi///yIiIgD/Ig//Ig//AiIi//8iIiIg/yIv/yIv/wIiIv//IiIiIA8iL/8iL/IiIiIv/3Ii/yJ/Ii//Ii/yIiL/Iv/yII///yIv/yIv8iIv/yL/9yAP//8iL/8iL/IiL/8ij/8iD///Ii//Ii/yIi//Ai//8gD//yIv/yIv8iIv8AIv//IA//8iL/8iL/Iij/ACL/AAIA//Ii//Ii/yIv/wIi/wACII/yIv8iIv8iL/ACIvIiIiAP8iL/IiL/IiAAAiLyIiIiAvIi/yIi/yIvACIi8iIiIg/yIv8iIv8iLyIiIv8iIiIv8iL/+iL/Ii9yIi//co/4L/Ii//Ki/yIv8iIv" & _
    "//L////yKv/yov8iL/+C////////8iL/IiD/Io////////////Iq/yog/yIv///////////yKv8qov8iL///////////Air/qqoAIi///////////wIq//iiICIv//////////ACKv//r4AiL//////////yIir////yIi///////////yIi////8iIv///////////yIv////8oL////////////yL/////8i/////////////y//////8v/////////////////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAEAICAQAAEABADoAgAAAQB0DDQAAABWAFMAXwBWAEUAUgBTAEkATwBOAF8A" & _
    "SQBOAEYATwAAAAAAvQTv/gAAAQACAAEA6wcDAAIAAQDrBwMAPwAAACAAAAAEAAAAAgAAAAAAAAAAAAAAAAAAANQLAAAAAFMAdAByAGkAbgBnAEYAaQBsAGUASQBuAGYAbwAAALALAAAAADAANAAwADkAMAA0AEUANAAAACoABQABAEMAbwBtAHAAYQBuAHkATgBhAG0AZQAAAAAAWgBsAGkAYgAAAAAASAeYAwEATABpAGMAZQBuAHMAZQAAAAAAKABDACkAIAAxADkAOQA1AC0AMgAwADAAMgAgAEoAZQBhAG4ALQBsAG8AdQBwACAARwBhAGkAbABsAHkAIABhAG4AZAAgAE0AYQByAGsAIABBAGQAbABlAHIAIABUAGgAaQBzACAAcwBvAGYAdAB3AGEAcgBlACAAaQBzACAAcAByAG8AdgBpAGQAZQBkACAAJwBhAHMALQBpAHMAJwAsACAAdwBpAHQAaABvAHUAdAAgAGEAbgB5ACAAZQB4AHAA" & _
    "cgBlAHMAcwAgAG8AcgAgAGkAbQBwAGwAaQBlAGQAIAB3AGEAcgByAGEAbgB0AHkALgAgAEkAbgAgAG4AbwAgAGUAdgBlAG4AdAAgAHcAaQBsAGwAIAB0AGgAZQAgAGEAdQB0AGgAbwByAHMAIABiAGUAIABoAGUAbABkACAAbABpAGEAYgBsAGUAIABmAG8AcgAgAGEAbgB5ACAAZABhAG0AYQBnAGUAcwAgAGEAcgBpAHMAaQBuAGcAIABmAHIAbwBtACAAdABoAGUAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABzAG8AZgB0AHcAYQByAGUALgAgAFAAZQByAG0AaQBzAHMAaQBvAG4AIABpAHMAIABnAHIAYQBuAHQAZQBkACAAdABvACAAYQBuAHkAbwBuAGUAIAB0AG8AIAB1AHMAZQAgAHQAaABpAHMAIABzAG8AZgB0AHcAYQByAGUAIABmAG8AcgAgAGEAbgB5ACAAcAB1AHIAcABvAHMA" & _
    "ZQAsACAAaQBuAGMAbAB1AGQAaQBuAGcAIABjAG8AbQBtAGUAcgBjAGkAYQBsACAAYQBwAHAAbABpAGMAYQB0AGkAbwBuAHMALAAgAGEAbgBkACAAdABvACAAYQBsAHQAZQByACAAaQB0ACAAYQBuAGQAIAByAGUAZABpAHMAdAByAGkAYgB1AHQAZQAgAGkAdAAgAGYAcgBlAGUAbAB5ACwAIABzAHUAYgBqAGUAYwB0ACAAdABvACAAdABoAGUAIABmAG8AbABsAG8AdwBpAG4AZwAgAHIAZQBzAHQAcgBpAGMAdABpAG8AbgBzADoAIAAxAC4AIABUAGgAZQAgAG8AcgBpAGcAaQBuACAAbwBmACAAdABoAGkAcwAgAHMAbwBmAHQAdwBhAHIAZQAgAG0AdQBzAHQAIABuAG8AdAAgAGIAZQAgAG0AaQBzAHIAZQBwAHIAZQBzAGUAbgB0AGUAZAA7ACAAeQBvAHUAIABtAHUAcwB0ACAAbgBvAHQA" & _
    "IABjAGwAYQBpAG0AIAB0AGgAYQB0ACAAeQBvAHUAIAB3AHIAbwB0AGUAIAB0AGgAZQAgAG8AcgBpAGcAaQBuAGEAbAAgAHMAbwBmAHQAdwBhAHIAZQAuACAASQBmACAAeQBvAHUAIAB1AHMAZQAgAHQAaABpAHMAIABzAG8AZgB0AHcAYQByAGUAIABpAG4AIABhACAAcAByAG8AZAB1AGMAdAAsACAAYQBuACAAYQBjAGsAbgBvAHcAbABlAGQAZwBtAGUAbgB0ACAAaQBuACAAdABoAGUAIABwAHIAbwBkAHUAYwB0ACAAZABvAGMAdQBtAGUAbgB0AGEAdABpAG8AbgAgAHcAbwB1AGwAZAAgAGIAZQAgAGEAcABwAHIAZQBjAGkAYQB0AGUAZAAgAGIAdQB0ACAAaQBzACAAbgBvAHQAIAByAGUAcQB1AGkAcgBlAGQALgAgADIALgAgAEEAbAB0AGUAcgBlAGQAIABzAG8AdQByAGMAZQAgAHYA" & _
    "ZQByAHMAaQBvAG4AcwAgAG0AdQBzAHQAIABiAGUAIABwAGwAYQBpAG4AbAB5ACAAbQBhAHIAawBlAGQAIABhAHMAIABzAHUAYwBoACwAIABhAG4AZAAgAG0AdQBzAHQAIABuAG8AdAAgAGIAZQAgAG0AaQBzAHIAZQBwAHIAZQBzAGUAbgB0AGUAZAAgAGEAcwAgAGIAZQBpAG4AZwAgAHQAaABlACAAbwByAGkAZwBpAG4AYQBsACAAcwBvAGYAdAB3AGEAcgBlAC4AIAAzAC4AIABUAGgAaQBzACAAbgBvAHQAaQBjAGUAIABtAGEAeQAgAG4AbwB0ACAAYgBlACAAcgBlAG0AbwB2AGUAZAAgAG8AcgAgAGEAbAB0AGUAcgBlAGQAIABmAHIAbwBtACAAYQBuAHkAIABzAG8AdQByAGMAZQAgAGQAaQBzAHQAcgBpAGIAdQB0AGkAbwBuAC4AIABKAGUAYQBuAC0AbABvAHUAcAAgAEcAYQBpAGwA" & _
    "bAB5ACAATQBhAHIAawAgAEEAZABsAGUAcgAgAGoAbABvAHUAcABAAGcAegBpAHAALgBvAHIAZwAgAG0AYQBkAGwAZQByAEAAYQBsAHUAbQBuAGkALgBjAGEAbAB0AGUAYwBoAC4AZQBkAHUAIAAAAKYAPwABAEYAaQBsAGUARABlAHMAYwByAGkAcAB0AGkAbwBuAAAAAABaAGwAaQBiADoAIABnAGUAbgBlAHIAYQBsACAAcAB1AHIAcABvAHMAZQAgAGQAYQB0AGEAIABjAG8AbQBwAHIAZQBzAHMAaQBvAG4AIAAvACAAZABlAGMAbwBtAHAAcgBlAHMAcwBpAG8AbgAgAGwAaQBiAHIAYQByAHkAAAAAADYACwABAEYAaQBsAGUAVgBlAHIAcwBpAG8AbgAAAAAAMQAuADIALgAzAC4AMgAwADIANwAAAAAAqgA/AAEATABpAGIAVABvAG8AbABGAGkAbABlAFYAZQByAHMAaQBvAG4AAABaAGwA"
Private Const STR_ZLIB_20 As String = _
    "aQBiADoAIABnAGUAbgBlAHIAYQBsACAAcAB1AHIAcABvAHMAZQAgAGQAYQB0AGEAIABjAG8AbQBwAHIAZQBzAHMAaQBvAG4AIAAvACAAZABlAGMAbwBtAHAAcgBlAHMAcwBpAG8AbgAgAGwAaQBiAHIAYQByAHkAAAAAACwABgABAEkAbgB0AGUAcgBuAGEAbABOAGEAbQBlAAAAegBsAGkAYgAxAAAAyABSAAEATABlAGcAYQBsAEMAbwBwAHkAcgBpAGcAaAB0AAAAqQAgADIAMAAwADUAIABKAGUAYQBuAC0AbABvAHUAcAAgAEcAYQBpAGwAbAB5ACAAPABqAGwAbwB1AHAAQABnAHoAaQBwAC4AbwByAGcAIAA+ACwAIABNAGEAcgBrACAAQQBkAGwAZQByACAAPABtAGEAZABsAGUAcgBAAGEAbAB1AG0AbgBpAC4AYwBhAGwAdABlAGMAaAAuAGUAZAB1AD4AAABSABUAAQBMAGUAZwBhAGwA" & _
    "VAByAGEAZABlAG0AYQByAGsAcwAAAAAAWgBsAGkAYgCuACwAIABaAGwAaQBiAK4ALAAgAHoAbABpAGIAMQCuAAAAAAA8AAoAAQBPAHIAaQBnAGkAbgBhAGwARgBpAGwAZQBuAGEAbQBlAAAAegBsAGkAYgAxAC4AZABsAGwAAAAqAAUAAQBQAHIAbwBkAHUAYwB0AE4AYQBtAGUAAAAAAFoAbABpAGIAAAAAADoACwABAFAAcgBvAGQAdQBjAHQAVgBlAHIAcwBpAG8AbgAAADEALgAyAC4AMwAuADIAMAAyADcAAAAAAHIAKQABAFMAcABlAGMAaQBhAGwAQgB1AGkAbABkAAAARwBOAFUAIABmAG8AcgAgAFcAaQBuADMAMgAgADwAZwBuAHUAdwBpAG4AMwAyAC4AcwBvAHUAcgBjAGUAZgBvAHIAZwBlAC4AbgBlAHQAPgAAAAAAOAAUAAEAVwBXAFcAAAAAAGgAdAB0AHAAOgAvAC8AdwB3AHcA" & _
    "LgB6AGwAaQBiAC4AbgBlAHQAAABEAAAAAABWAGEAcgBGAGkAbABlAEkAbgBmAG8AAAAAACQABAAAAFQAcgBhAG4AcwBsAGEAdABpAG8AbgAAAAAACQTkBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAIAAAADowfTCMMLswxTDoMAQxGjElMUoxVTH9NAAwAACkAAAAwTUfNlA2VzZmNnQ2jDaWNqI2sDbINtI23jbsNgQ3CzcaNyo3QDdKN1Y3ZDd8N4Y3kjegN7g3vzfON9w39Df7Nw04GzhTOF04aTh3OJ84ETlTOV05aTl3OY85ljmlObM5yznVOeE57zkEOgs6GjorOkM6SjpZOmc6fzqJOpU6ozq7OsU60TrfOvQ6+zoNOxs7UztdO2k7dzuhOwE+AFAAADQAAABHM3EztzPxMyc0JzYfN3k3kTlIOlQ6YDpsOng6hDqQOpw6qDq0OqE7tjurPQBgAAAcAAAA" & _
    "PDFGMWYxbTF0MYExjDP6NV4+AAAAcAAAJAAAAAwyLjJCMk8yyTXsNQU2ujfLNy07UDtpO8Y81zwAgAAADAAAABc/Oz8AkAAANAAAAB4wPjC4MPUw8zFRNFk0gzYbNyY3tzc7OEo4GjmjO7o70TvlPAw9Ij1GPVg9AKAAACQAAACiNVM2oD+kP6g/rD+wP7Q/uD+8P8A/xD/RP/4/ALAAAFQAAAC0MbkxzjHIPcw90D3UPdg93D3gPeQ96D3sPfA99D34Pfw9AD4EPgg+DD4QPhQ+GD4cPiA+JD4oPiw+MD40Pjg+PD5APgQ/Nz9ePwAAAMAAACQAAADhMQIyPzXnOAI5EDm8ObY6JTtGO7Y7Oj2PPew+ANAAABwAAABeMOYw1jLFM9M0PzXGPMs84DwAAADgAABAAAAAyDfMN9A31DfYN9w34DfkN+g37DfwN/Q3+Df8NwA4BDgIOAw4sDjJONc4mTk8O587wDtBPLk9Dz4A8AAAHAAAAIYwIzKlMpQ0" & _
    "xzUuOjY63j3mPag/AAABAKgAAADwML4yJzQ5NEQ0WjRsNIM0jjSZNLs0xDTPNOE09DT/NAo1IjUoNTo1QDV3NbM1uTXANcg12jXhNfA19jUaNjg2UTZnNn02jzaXNp82pzbnNu829zZyN3o3gjeKN5I3mjeiN6o3YjpyOoI6kjqiOrI6wjrSOuI68joCOxI7IjsyO0I7UjtiO3I7gjuSO6I7sjvCO9I74jvyOwI8Ejw0PAAAABABABQAAAAAMAQwFDAYMCwwRDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
Private Const STR_ZLIB_COUNT As Long = 20


'=========================================================================
' API
'=========================================================================

'--- zlib1.dll API (opsional, untuk kompresi yang lebih baik)
'--- zlib function pointers (resolved at runtime via GetProcAddress)
'--- Ini menggantikan Declare Function ... Lib "zlib1.dll" agar tidak crash jika DLL tidak ada
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Long, ByVal cActuals As Long, prgvt As Any, prgpvarg As Any, pvargResult As Variant) As Long
Private Const CC_CDECL As Long = 1
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SetDllDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

'--- archiveint.dll API (libarchive, bawaan Windows 11 23H2+)
Private Declare Function arc_read_new Lib "archiveint.dll" Alias "archive_read_new" () As Long
Private Declare Function arc_read_free Lib "archiveint.dll" Alias "archive_read_free" (ByVal h As Long) As Long
Private Declare Function arc_read_close Lib "archiveint.dll" Alias "archive_read_close" (ByVal h As Long) As Long
Private Declare Function arc_read_support_format_all Lib "archiveint.dll" Alias "archive_read_support_format_all" (ByVal h As Long) As Long
Private Declare Function arc_read_support_filter_all Lib "archiveint.dll" Alias "archive_read_support_filter_all" (ByVal h As Long) As Long
Private Declare Function arc_read_open_filename Lib "archiveint.dll" Alias "archive_read_open_filename" (ByVal h As Long, ByVal fn As Long, ByVal blockSize As Long) As Long
Private Declare Function arc_read_next_header Lib "archiveint.dll" Alias "archive_read_next_header" (ByVal h As Long, hEntry As Long) As Long
Private Declare Function arc_read_data Lib "archiveint.dll" Alias "archive_read_data" (ByVal h As Long, ByVal pData As Long, ByVal DataSize As Long) As Long
Private Declare Function arc_read_add_passphrase Lib "archiveint.dll" Alias "archive_read_add_passphrase" (ByVal h As Long, ByVal pw As Long) As Long
Private Declare Function arc_write_new Lib "archiveint.dll" Alias "archive_write_new" () As Long
Private Declare Function arc_write_free Lib "archiveint.dll" Alias "archive_write_free" (ByVal h As Long) As Long
Private Declare Function arc_write_close Lib "archiveint.dll" Alias "archive_write_close" (ByVal h As Long) As Long
Private Declare Function arc_write_open_filename Lib "archiveint.dll" Alias "archive_write_open_filename" (ByVal h As Long, ByVal fn As Long) As Long
Private Declare Function arc_write_header Lib "archiveint.dll" Alias "archive_write_header" (ByVal h As Long, ByVal entry As Long) As Long
Private Declare Function arc_write_data Lib "archiveint.dll" Alias "archive_write_data" (ByVal h As Long, ByVal pData As Long, ByVal DataSize As Long) As Long
Private Declare Function arc_write_finish_entry Lib "archiveint.dll" Alias "archive_write_finish_entry" (ByVal h As Long) As Long
Private Declare Function arc_write_set_format_zip Lib "archiveint.dll" Alias "archive_write_set_format_zip" (ByVal h As Long) As Long
Private Declare Function arc_write_set_format_7zip Lib "archiveint.dll" Alias "archive_write_set_format_7zip" (ByVal h As Long) As Long
Private Declare Function arc_write_set_format_ustar Lib "archiveint.dll" Alias "archive_write_set_format_ustar" (ByVal h As Long) As Long
Private Declare Function arc_write_set_format_pax Lib "archiveint.dll" Alias "archive_write_set_format_pax_restricted" (ByVal h As Long) As Long
Private Declare Function arc_write_set_format_cpio Lib "archiveint.dll" Alias "archive_write_set_format_cpio" (ByVal h As Long) As Long
Private Declare Function arc_write_set_format_xar Lib "archiveint.dll" Alias "archive_write_set_format_xar" (ByVal h As Long) As Long
Private Declare Function arc_write_add_filter_none Lib "archiveint.dll" Alias "archive_write_add_filter_none" (ByVal h As Long) As Long
Private Declare Function arc_write_add_filter_gzip Lib "archiveint.dll" Alias "archive_write_add_filter_gzip" (ByVal h As Long) As Long
Private Declare Function arc_write_add_filter_bzip2 Lib "archiveint.dll" Alias "archive_write_add_filter_bzip2" (ByVal h As Long) As Long
Private Declare Function arc_write_add_filter_xz Lib "archiveint.dll" Alias "archive_write_add_filter_xz" (ByVal h As Long) As Long
Private Declare Function arc_write_add_filter_zstd Lib "archiveint.dll" Alias "archive_write_add_filter_zstd" (ByVal h As Long) As Long
Private Declare Function arc_write_set_options Lib "archiveint.dll" Alias "archive_write_set_options" (ByVal h As Long, ByVal opts As Long) As Long
Private Declare Function arc_entry_new Lib "archiveint.dll" Alias "archive_entry_new" () As Long
Private Declare Sub arc_entry_free Lib "archiveint.dll" Alias "archive_entry_free" (ByVal entry As Long)
Private Declare Sub arc_entry_set_pathname_utf8 Lib "archiveint.dll" Alias "archive_entry_set_pathname_utf8" (ByVal entry As Long, ByVal pathname As Long)
Private Declare Sub arc_entry_set_size Lib "archiveint.dll" Alias "archive_entry_set_size" (ByVal entry As Long, ByVal sz As Long)
Private Declare Sub arc_entry_set_filetype Lib "archiveint.dll" Alias "archive_entry_set_filetype" (ByVal entry As Long, ByVal ft As Long)
Private Declare Sub arc_entry_set_perm Lib "archiveint.dll" Alias "archive_entry_set_perm" (ByVal entry As Long, ByVal perm As Long)
Private Declare Function arc_entry_pathname_w Lib "archiveint.dll" Alias "archive_entry_pathname_w" (ByVal hEntry As Long) As Long
Private Declare Function arc_entry_size Lib "archiveint.dll" Alias "archive_entry_size" (ByVal hEntry As Long) As Long
Private Declare Function arc_entry_filetype Lib "archiveint.dll" Alias "archive_entry_filetype" (ByVal hEntry As Long) As Long
Private Declare Function arc_error_string Lib "archiveint.dll" Alias "archive_error_string" (ByVal h As Long) As Long
Private Declare Function arc_version_string Lib "archiveint.dll" Alias "archive_version_string" () As Long
'--- archiveint constants
Private Const ARC_OK As Long = 0
Private Const ARC_EOF As Long = 1
Private Const ARC_WARN As Long = -20
Private Const ARC_IFREG As Long = &H8000&
Private Const ARC_IFDIR As Long = &H4000&
Private Const ARC_PERM_FILE As Long = &H1A4
Private Const ARC_PERM_DIR As Long = &H1ED
Private Const ARC_BLOCK_SIZE As Long = 65536

'--- zlib constants
Private Const Z_OK                          As Long = 0
Private Const Z_STREAM_END                  As Long = 1
Private Const Z_FINISH                      As Long = 4
Private Const Z_NO_FLUSH                    As Long = 0
Private Const Z_DEFAULT_STRATEGY            As Long = 0
Private Const Z_DEFLATED                    As Long = 8
Private Const ZLIB_VERSION                  As String = "1.3.1"

Private Const VT_BYREF                      As Long = &H4000
'--- for FindFirstFile
Private Const MAX_PATH                      As Long = 260
Private Const INVALID_HANDLE_VALUE          As Long = -1
'--- for CreateFile
Private Const CREATE_ALWAYS                 As Long = 2
Private Const OPEN_EXISTING                 As Long = &H3
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
'--- for VirtualAlloc
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
'--- for SetFilePointer
Private Const FILE_BEGIN                    As Long = 0
Private Const FILE_CURRENT                  As Long = 1
Private Const FILE_END                      As Long = 2
'--- for MultiByteToWideChar
Private Const CP_UTF8                       As Long = 65001
Private Const CP_OEMCP                      As Long = 1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for CryptBinaryToString/CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for CNG
Private Const MS_PRIMITIVE_PROVIDER         As String = "Microsoft Primitive Provider"
Private Const BCRYPT_SHA1_ALGORITHM         As String = "SHA1"
Private Const BCRYPT_AES_ALGORITHM          As String = "AES"
Private Const BCRYPT_HASH_LENGTH            As String = "HashDigestLength"
Private Const BCRYPT_OBJECT_LENGTH          As String = "ObjectLength"
Private Const BCRYPT_CHAINING_MODE          As String = "ChainingMode"
Private Const BCRYPT_CHAIN_MODE_ECB         As String = "ChainingModeECB"
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8
'--- for GetStdHandle
Private Const STD_INPUT_HANDLE              As Long = -10
Private Const STD_OUTPUT_HANDLE             As Long = -11
Private Const STD_ERROR_HANDLE              As Long = -12

Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMove As Currency, lpNewFilePointer As Any, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32" (lpSystemTime As SYSTEMTIME, pvTime As Date) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32" (ByVal vTime As Date, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cbMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If ImplCompress Then
    Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
#End If ' ImplCompress
#If ImplDecompress Then
    Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
#End If ' ImplDecompress
#If ImplCrypto Then
    Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, phKey As Long, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
    Private Declare Function BCryptEncrypt Lib "bcrypt" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal pPrf As Long, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Long, ByVal dwDummy As Long, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long
    Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
#End If ' ImplCrypto
#If ImplUseShared = 0 Then
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageW" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Args As Any) As Long
    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
    Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
    Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
#End If ' Not ImplUseShared

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

'--- ZIP Constants
Private Const LNG_SIG_LOCAL                 As Long = &H4034B50
Private Const LNG_SIG_CENTRAL               As Long = &H2014B50
Private Const LNG_SIG_END                   As Long = &H6054B50
Private Const LNG_SIG_DATADESC              As Long = &H8074B50
Private Const LNG_SIG_EXTLOCAL              As Long = &H8074B50
Private Const LNG_SIG_EOCDR                 As Long = &H6064B50
Private Const LNG_SIG_LOCATOR               As Long = &H7064B50
Private Const LNG_LEN_LOCAL                 As Long = 30
Private Const LNG_LEN_CENTRAL               As Long = 46
Private Const LNG_LEN_END                   As Long = 22
Private Const LNG_LEN_EXTLOCAL              As Long = 16
Private Const LNG_LEN_EOCDR                 As Long = 56
Private Const LNG_LEN_LOCATOR               As Long = 20
Private Const LNG_VER_DEFAULT               As Long = 10
Private Const LNG_VER_DEFLATE               As Long = 20
Private Const LNG_VER_ZIP64                 As Long = 45
Private Const LNG_METHOD_DEFLATE            As Long = 8
Private Const LNG_METHOD_STORE              As Long = 0
Private Const LNG_IOBUF_SIZE                As Long = 65536
Private Const LNG_METHOD_AES                As Long = 99
Private Const LNG_EXTRADATA_AES_HEADER      As Integer = &H9901
Private Const LNG_EXTRADATA_AES_VENDOR      As Integer = &H4541 '--- "AE"
Private Const LNG_EXTRADATA_AES_SIZE        As Long = 7
Private Const LNG_EXTRADATA_AE_2_MAXSIZE    As Long = 20
Private Const LNG_EXTRADATA_ZIP64_HEADER    As Long = 1
Private Const LNG_EXTRADATA_ZIP64_LOCAL_SIZE As Long = 16
Private Const LNG_ENC_HEADER_SIZE           As Long = 12
Private Const LNG_AES_AUTHCODE_SIZE         As Long = 10
Private Const LNG_AES_BLOCK_SIZE            As Long = 16
Private Const LNG_PIPE_ENDED_ERROR          As Long = 109
Private Const LNG_SIZE_LIMIT                As Currency = 2 ^ 32 - 2
Private Const LNG_ENTRIES_LIMIT             As Long = 2 ^ 16 - 2
'--- errors
Private Const ERR_USER_CANCEL               As String = "User cancelled"
Private Const ERR_INIT_COMPRESSOR           As String = "Cannot init deflate compressor"
Private Const ERR_COMPRESSING               As String = "Error compressing"
Private Const ERR_EMPTY_ARCHIVE             As String = "Empty archive"
Private Const ERR_INVALID_ARCHIVE           As String = "Invalid archive"
Private Const ERR_ARCHIVE_NOT_OPEN          As String = "Archive not opened"
Private Const ERR_INVALID_LOCALHDR          As String = "Invalid local header"
Private Const ERR_UNSUPPORTED_METHOD        As String = "Unsupported compression method %1"
Private Const ERR_INIT_DECOMPRESSOR         As String = "Cannot init inflate decompressor"
Private Const ERR_READING_ARCHIVE           As String = "Error reading archive"
Private Const ERR_DECOMPRESSING             As String = "Error decompressing"
Private Const ERR_CRC_CHECK                 As String = "CRC check failed"
Private Const ERR_ENTRY_INVALID_SIG         As String = "Entry %1 has invalid signature"
Private Const ERR_ENTRY_NO_FILENAME         As String = "Entry %1 has no filename"
Private Const ERR_UNSUPPORTED_ENCRYPTION    As String = "Unsupported encryption"
Private Const ERR_PASSWORD_REQUIRED         As String = "Password required"
Private Const ERR_INVALID_PASSWORD          As String = "Invalid password"
Private Const ERR_INVALID_AUTHCODE          As String = "Invalid authentication code"
Private Const ERR_INVALID_BYTEARRAY         As String = "Invalid byte array"
Private Const ERR_MISSING_ZIP64             As String = "Missing Zip64 extension above 4GB limit"
Private Const ERR_NO_CRYPTO_COMPILED        As String = "Crypto not compiled (ZIP_CRYPTO = 1)"
Private Const STR_BUFFER                    As String = "[buffer]"
Private Const STR_STREAM                    As String = "[stream]"
Private Const STR_ENTRY                     As String = "[entry %1]"
'--- GZ Constants & Errors
Private Const ERR_GZ_INVALID_HEADER         As String = "Invalid GZ data: invalid header"
Private Const ERR_GZ_TOO_SHORT              As String = "Invalid GZ data: too short"
Private Const ERR_GZ_SIZE_MISMATCH          As String = "Decompression failed: size mismatch"
Private Const ERR_GZ_CRC_MISMATCH           As String = "Decompression failed: CRC32 mismatch"


Private m_uRtbl                     As UcsZlibRelocTableType
Private m_sLastError                As String
Private m_bCancel                   As Boolean
Private m_lFileCount                As Long
Private m_uFiles()                  As UcsFileInfo
Private m_lCurrentFile              As Long
Private m_lCodePage                 As Long
Private m_hZlib                     As Long     '--- handle ke zlib1.dll (0 jika tidak tersedia)
Private m_eGzEngine             As GzEngineType
Private m_bZlibChecked              As Boolean  '--- flag apakah sudah dicek
Private m_pfnDeflateInit2           As Long     '--- zlib deflateInit2_
Private m_pfnDeflate                As Long     '--- zlib deflate
Private m_pfnDeflateEnd             As Long     '--- zlib deflateEnd
Private m_pfnInflateInit2           As Long     '--- zlib inflateInit2_
Private m_pfnInflate                As Long     '--- zlib inflate
Private m_pfnInflateEnd             As Long     '--- zlib inflateEnd
Private m_pfnCrc32                  As Long     '--- zlib crc32
Private m_pfnZlibVersion            As Long     '--- zlib zlibVersion
Private m_pfnCompress2              As Long     '--- zlib compress2
Private m_pfnUncompress             As Long     '--- zlib uncompress
Private m_pfnCompressBound          As Long     '--- zlib compressBound
Private m_hArcInt                   As Long     '--- handle ke archiveint.dll
Private m_bArcChecked               As Boolean  '--- flag apakah sudah dicek
Private m_hArcRead                  As Long     '--- handle read archive (ArcOpen)
Private m_hArcWrite                 As Long     '--- handle write archive (ArcCreate)
Private m_lArcEntryCount            As Long     '--- jumlah entry di arsip
Private m_sArcEntries()             As String   '--- nama entry
Private m_lArcSizes()               As Long     '--- ukuran entry
Private m_lArcTypes()               As Long     '--- tipe entry (file/dir)
#If ImplDecompress Then
    Private m_sComment              As String
    Private m_vArchiveFile          As Variant
#End If ' ImplDecompress
#If ImplCrypto Then
    Private m_uCrypto               As UcsZipCryptoType
#End If ' ImplCrypto

'--- zlib stream structure (matches z_stream in zlib.h for 32-bit)
Private Type z_stream
    next_in             As Long     ' pointer to input buffer
    avail_in            As Long     ' bytes available at next_in
    total_in            As Long     ' total bytes read so far
    next_out            As Long     ' pointer to output buffer
    avail_out           As Long     ' bytes available at next_out
    total_out           As Long     ' total bytes written so far
    msg                 As Long     ' pointer to error message
    state               As Long     ' internal state (opaque)
    zalloc              As Long     ' custom allocator (0 = default)
    zfree               As Long     ' custom free (0 = default)
    opaque              As Long     ' private data for allocator
    data_type           As Long     ' best guess about data type
    adler               As Long     ' Adler-32 / CRC-32 checksum
    reserved            As Long     ' reserved for future use
End Type

Private Type UcsVfsFileType
    Handle              As Long
    Stream              As Object
    BufferArray         As Variant
    BufferBase          As LongPtr
    BufferSize          As Long
    BufferPtr           As LongPtr
    Data                As WIN32_FIND_DATA
    FileName            As String
End Type

Private Type UcsZlibRelocTableType
    CompressInit        As Long
    CompressCleanup     As Long
    CompressBlock       As Long
    DecompressInit      As Long
    DecompressCleanup   As Long
    DecompressBlock     As Long
    CalcCrc32           As Long
    MemNonce            As Long
    MemXor              As Long
    ZipCrypt            As Long
    MallocImpl          As Long
    ReallocImpl         As Long
    FreeImpl            As Long
    LenCodes            As Long
    DistCodes           As Long
    MirrorBytes         As Long
    LenLenMap           As Long
    Crc32Table          As Long
End Type

Private Type UcsZlibBuffersType
    InBlock             As Long
    InLen               As Long
    OutBlock            As Long
    OutLen              As Long
    Final               As Long
    Greedy              As Long
    MaxMatch            As Long
    NiceLen             As Long
End Type

Private Enum UcsRelocIndexesEnum
    ucsIdx_CompressInit = 0
    ucsIdx_CompressCleanup
    ucsIdx_CompressBlock
    ucsIdx_DecompressInit
    ucsIdx_DecompressCleanup
    ucsIdx_DecompressBlock
    ucsIdx_CalcCrc32
    ucsIdx_MemNonce
    ucsIdx_MemXor
    ucsIdx_ZipCrypt
    ucsIdx_MallocImpl
    ucsIdx_ReallocImpl
    ucsIdx_FreeImpl
    ucsIdx_LenCodes
    ucsIdx_DistCodes
    ucsIdx_MirrorBytes
    ucsIdx_LenLenMap
End Enum

Private Enum UcsHeaderFlagsEnum
    ucsZcfEncrypted = 2 ^ 0             ' If set, indicates that the file is encrypted
    ucsZcfCompressOptionMask = 2 ^ 1 Or 2 ^ 2 ' 0 - normal, 1 - maximum, 2 - fast, 3 - super fast
    ucsZcfHasDataDescriptor = 2 ^ 3     ' Values are put in the data descriptor immediately following the compressed data
    ucsZcfStrongEncrypted = 2 ^ 6       ' Strong encryption
    ucsZcfUseUtf8 = 2 ^ 11              ' Language encoding flag (EFS)
End Enum

Private Type UcsZip64ExtraDataType
    HeaderId            As Integer
    DataSize            As Integer
    USize               As Currency
    CSize               As Currency
    Offset              As Currency
    DiskNum             As Long
End Type

Private Type UcsZip64EndHeaderType
    Signature           As Long
    Size                As Currency
    VerMade             As Integer
    VerExt              As Integer
    DiskNum             As Long
    DiskStart           As Long
    Entries             As Currency
    TotEntr             As Currency
    CenSize             As Currency
    CenOff              As Currency
End Type

Private Type UcsZip64EndLocatorType
    Signature           As Long
    DiskEocdr           As Long
    Offset              As Currency
    DiskNum             As Long
End Type

Private Type UcsLocalHeaderType
    Signature           As Long                 ' Signature
    VerExt              As Integer              ' version needed to extract
    Flags               As Integer              ' encrypt and compression flags
    Method              As Integer              ' compression method
    FTime               As Integer              ' time last modifies, dos format
    FDate               As Integer              ' date last modifies, dos format
    '--- padding
    Crc32               As Long                 ' CRC32 for uncompressed file
    CSize               As Long                 ' compressed size
    USize               As Long                 ' uncompressed size
    LenFname            As Integer              ' Length filename
    LenExt              As Integer              ' Length for extra field
    Ext                 As UcsZip64ExtraDataType
End Type

Private Type UcsCentralHeaderType
    Signature           As Long                 ' Signature
    VerMade             As Integer              ' version made by
    VerExt              As Integer              ' version needed to extract
    Flags               As Integer              ' encrypt and compression flags
    Method              As Integer              ' compression method
    FTime               As Integer              ' time last modifies, dos format
    FDate               As Integer              ' date last modifies, dos format
    '--- padding
    Crc32               As Long                 ' CRC32 for uncompressed file
    CSize               As Long                 ' compressed size
    USize               As Long                 ' uncompressed size
    LenFname            As Integer              ' Length filename
    LenExt              As Integer              ' Length for extra field
    LenCom              As Integer              ' Length for comment field
    DiskStart           As Integer              ' start disk number
    AttribI             As Integer              ' internal file attributes
    '--- padding
    AttribX             As Long                 ' external file attributes
    Offset              As Long                 ' relative offset of local header
    Ext                 As UcsZip64ExtraDataType
End Type

Private Type UcsEndHeaderType
    Signature           As Long                 ' Signature
    DiskNum             As Integer              ' This disk number
    DiskStart           As Integer              ' Start disk number
    Entries             As Integer              ' Entries on this disk
    TotEntr             As Integer              ' Number of total entries
    CenSize             As Long                 ' Size of entire cetral directory
    CenOff              As Long                 ' Offset of central on starting disk
    LenCom              As Integer              ' Length of comment field
    Ext                 As UcsZip64EndHeaderType
End Type

Private Type UcsAesExtraDataType
    HeaderId            As Integer              ' Extra field header ID (0x9901)
    DataSize            As Integer              ' Data size (currently 7, but subject to possible increase in the future)
    Version             As Integer              ' Integer version number specific to the zip vendor
    VendorId            As Integer              ' 2-character vendor ID
    Strength            As Byte                 ' Integer mode value indicating AES encryption strength
    '--- padding
    Method              As Integer              ' The actual compression method used to compress the file
End Type

Private Type UcsFileInfo
    FileName            As String
    Attributes          As Long
    Crc32               As Long
    Size                As Currency
    CompressedSize      As Currency
    Comment             As String
    LastModified        As Date
    Method              As Long
    Offset              As Currency
    Flags               As Long
    Extra()             As Byte
    SourceFile          As Variant
    Level               As Long
#If ImplCrypto Then
    Aes                 As UcsAesExtraDataType
    DecDat              As Byte
    Password            As String
#End If ' ImplCrypto
End Type

Private Type UcsZipCryptoType
    hPbkdf2Alg          As Long
    hHmacAlg            As Long
    hHmacHash           As Long
    HmacLen             As Long
    hAesAlg             As Long
    hAesKey             As Long
    AesKeyObjData()     As Byte
    AesKeyObjLen        As Long
    Nonce(0 To 1)       As Long
    EncrData()          As Byte
    EncrPos             As Long
    TradKey(0 To 3)     As Long
End Type

'=========================================================================
' Properties
'=========================================================================

Public Property Get SemVersion() As String
    SemVersion = "0.6.1-merged"
End Property

Public Property Get ThunkBuildDate() As String
    ThunkBuildDate = STR_THUNK_BUILDDATE
End Property

Public Property Get LastError() As String
    LastError = m_sLastError
End Property

Public Property Get ZlibAvailable() As Boolean
    ZlibAvailable = pvZlibAvailable()
End Property

Public Property Get ZlibVersion() As String
    If pvZlibAvailable() Then
        Dim lPtr As Long
        lPtr = zlib_zlibVersion()
        If lPtr <> 0 Then
            Dim lLen As Long
            lLen = lstrlenA(lPtr)
            If lLen > 0 And lLen < 50 Then
                Dim baVer() As Byte
                ReDim baVer(0 To lLen - 1)
                Call CopyMemory(baVer(0), ByVal lPtr, lLen)
                ZlibVersion = StrConv(baVer, vbUnicode)
            End If
        End If
    End If
End Property

Public Function ZlibDiagnostic() As String
    '--- Test zlib functions via DispCallFunc untuk diagnostik
    Dim s As String
    If Not pvZlibAvailable() Then
        ZlibDiagnostic = "Zlib not available"
        Exit Function
    End If
    
    On Error Resume Next
    
    '--- Test 1: compressBound(1000) - harus ~1013
    Dim lCBound As Long
    lCBound = zlib_compressBound(1000)
    s = "compressBound(1000)=" & lCBound & vbCrLf
    If Err.Number <> 0 Then s = s & "ERR: " & Err.Description & vbCrLf: Err.Clear
    
    '--- Test 2: compress2 - kompres 1000 byte "A" 
    Dim baSrc(0 To 999) As Byte
    Dim i As Long
    For i = 0 To 999: baSrc(i) = 65: Next
    
    Dim lDestLen As Long
    If lCBound > 0 Then lDestLen = lCBound Else lDestLen = 2000
    Dim baDest() As Byte
    ReDim baDest(0 To lDestLen - 1)
    
    Dim lRet As Long
    lRet = zlib_compress2(VarPtr(baDest(0)), lDestLen, VarPtr(baSrc(0)), 1000, 9)
    s = s & "compress2 ret=" & lRet & ", outLen=" & lDestLen & vbCrLf
    If Err.Number <> 0 Then s = s & "ERR: " & Err.Description & vbCrLf: Err.Clear
    
    '--- Test 3: crc32
    Dim lCrc As Long
    lCrc = zlib_crc32(0, VarPtr(baSrc(0)), 1000)
    s = s & "crc32(1000 A)=" & lCrc & vbCrLf
    If Err.Number <> 0 Then s = s & "ERR: " & Err.Description & vbCrLf: Err.Clear
    
    '--- Function pointers
    s = s & "pfnDeflateInit2=" & m_pfnDeflateInit2 & vbCrLf
    s = s & "pfnDeflate=" & m_pfnDeflate & vbCrLf
    s = s & "pfnCrc32=" & m_pfnCrc32 & vbCrLf
    
    ZlibDiagnostic = s
End Function

Public Property Get GzEngine() As GzEngineType
    GzEngine = m_eGzEngine
End Property

Public Property Let GzEngine(ByVal eEngine As GzEngineType)
    m_eGzEngine = eEngine
End Property

Public Property Get FileCount() As Long
    FileCount = m_lFileCount
End Property

Public Property Get FileInfo(ByVal FileIdx As Long, Optional ByVal Index As UcsFileInfoIndexEnum = -1) As Variant
    If FileIdx >= 0 And FileIdx < m_lFileCount Then
        With m_uFiles(FileIdx)
            Select Case Index
            Case Is < 0
                FileInfo = Array(.FileName, .Attributes, .Crc32, .Size, .CompressedSize, .Comment, .LastModified, .Method, .Offset, .Flags)
            Case ucsIdxFileName
                FileInfo = .FileName
            Case ucsIdxAttributes
                FileInfo = .Attributes
            Case ucsIdxCrc32
                FileInfo = .Crc32
            Case ucsIdxSize
                FileInfo = .Size
            Case ucsIdxCompressedSize
                FileInfo = .CompressedSize
            Case ucsIdxComment
                FileInfo = .Comment
            Case ucsIdxLastModified
                FileInfo = .LastModified
            Case ucsIdxMethod
                FileInfo = .Method
            Case ucsIdxOffset
                FileInfo = .Offset
            Case ucsIdxFlags
                FileInfo = .Flags
            End Select
        End With
    End If
End Property

Public Property Get CodePage() As Long
    CodePage = m_lCodePage
End Property

Public Property Let CodePage(ByVal lValue As Long)
    m_lCodePage = lValue
End Property

#If ImplDecompress Then

Public Property Get Comment() As String
    Comment = m_sComment
End Property
    
#End If ' ImplDecompress

'=========================================================================
'
'   GZIP Methods
'
'=========================================================================

Public Function GzCompress(ByVal Source As Variant, ByRef Destination As Variant, Optional ByVal Level As ArchiveCompressionLevel = acNormalCompression) As Boolean
    ' Deskripsi:
    ' Mengkompresi sumber menjadi format GZ ke tujuan.
    ' Menggunakan streaming (per-chunk) untuk mendukung file besar.
    ' Jika zlib1.dll tersedia, menggunakan zlib (Dynamic Huffman, kompresi lebih baik).
    ' Jika tidak, fallback ke embedded thunk (Fixed Huffman).
    '
    ' Parameter:
    ' Source (Variant): Sumber data yang akan dikompres. Bisa berupa:
    '   - String: Path lengkap ke file (e.g., "C:\data.txt")
    '   - Byte Array: Data dalam bentuk byte array di memori.
    '
    ' Destination (Variant, ByRef): Tujuan penyimpanan hasil kompresi. Bisa berupa:
    '   - String: Path lengkap file tujuan (e.g., "C:\arsip.gz")
    '   - Byte Array: Variabel byte array kosong yang akan diisi dengan hasil kompresi.
    '
    Const FUNC_NAME As String = "GzCompress"
    Dim hCtx As Long, lResult As Long
    Dim uBuf As UcsZlibBuffersType
    Dim baBuffer() As Byte, baNext() As Byte
    Dim lCrc32 As Long, cTotalSize As Currency
    Dim gzHeader(0 To 9) As Byte, gzFooter(0 To 7) As Byte
    Dim uSourceFile As UcsVfsFileType, uDestFile As UcsVfsFileType
    Dim lSize As Long
    Dim bToByteArray As Boolean
    Dim baGzOut() As Byte
    
    On Error GoTo EH
    pvSetError
    
    ' 0. Pilih engine berdasarkan GzEngine property
    Dim bUseZlib As Boolean
    Select Case m_eGzEngine
        Case geZlib:  bUseZlib = pvZlibAvailable()   '--- Force zlib
        Case geThunk: bUseZlib = False                '--- Force thunk
        Case Else:    bUseZlib = pvZlibAvailable()    '--- geAuto
    End Select
    
    If bUseZlib Then
        GzCompress = pvGzCompressZlib(Source, Destination, Level)
        Exit Function
    End If
    
    ' === FALLBACK: Thunk DEFLATE (Fixed Huffman) ===
    
    ' 1. Validasi dan siapkan level kompresi
    Level = IIf(Level < acNoCompression, acNormalCompression, IIf(Level > acMaxCompression, acMaxCompression, Level))
    
    ' 2. Buka sumber menggunakan VFS
    uSourceFile = pvVfsOpen(Source)
    pvVfsSeek uSourceFile, 0, FILE_BEGIN
    
    ' 3. Buat GZ Header dengan metadata yang benar (RFC 1952)
    gzHeader(0) = &H1F: gzHeader(1) = &H8B  ' Magic number
    gzHeader(2) = 8                           ' CM: deflate
    gzHeader(3) = 0                           ' FLG: no extras
    ' gzHeader(4-7) = 0                       ' MTIME: not set
    gzHeader(8) = IIf(Level >= acMaxCompression, 2, IIf(Level <= acFastestCompression, 4, 0)) ' XFL
    gzHeader(9) = 11                          ' OS: NTFS/Windows
    
    ' 4. Siapkan output — file atau byte array
    bToByteArray = IsMissing(Destination) Or VarType(Destination) = (vbArray Or vbByte)
    If bToByteArray Then
        ReDim baGzOut(0 To LNG_IOBUF_SIZE - 1) As Byte
        uDestFile = pvVfsCreate(baGzOut)
    Else
        uDestFile = pvVfsCreate(Destination)
    End If
    
    ' 5. Tulis GZ Header
    pvVfsWrite uDestFile, VarPtr(gzHeader(0)), 10
    
    ' 6. Inisialisasi DEFLATE compressor
    hCtx = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
    If hCtx = 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_INIT_COMPRESSOR

    uBuf.Greedy = (Level <= 4)
    uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), Level)
    uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), Level)
    lCrc32 = -1
    cTotalSize = 0
    
    ' 7. Streaming compression — baca per-chunk, kompres, tulis langsung
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baNext(0 To LNG_IOBUF_SIZE - 1) As Byte
    
    lSize = pvVfsRead(uSourceFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
    If lSize > 0 Then
        Do
            Call CopyMemory(baBuffer(0), baNext(0), lSize)
            uBuf.InBlock = VarPtr(baBuffer(0))
            uBuf.InLen = lSize
            lSize = pvVfsRead(uSourceFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
            uBuf.Final = (lSize = 0)
            lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf), VarPtr(lCrc32), CLng(Level))
            If lResult = 0 Or uBuf.OutBlock = 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_COMPRESSING
            pvVfsWrite uDestFile, uBuf.OutBlock, uBuf.OutLen
            Call CoTaskMemFree(uBuf.OutBlock)
            uBuf.OutBlock = 0
            cTotalSize = cTotalSize + uBuf.InLen
        Loop While lSize <> 0
    Else
        ' Tangani sumber kosong — tulis blok DEFLATE final kosong
        uBuf.InBlock = VarPtr(baBuffer(0))
        uBuf.InLen = 0
        uBuf.Final = -1
        lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf), VarPtr(lCrc32), CLng(Level))
        If lResult = 0 Or uBuf.OutBlock = 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_COMPRESSING
        pvVfsWrite uDestFile, uBuf.OutBlock, uBuf.OutLen
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    
    ' 8. Tulis GZ Footer (CRC32 + ISIZE mod 2^32)
    lCrc32 = lCrc32 Xor -1
    Dim lISize As Long
    lISize = pvToInt32(cTotalSize)  ' ISIZE = ukuran asli mod 2^32 (RFC 1952)
    Call CopyMemory(gzFooter(0), lCrc32, 4)
    Call CopyMemory(gzFooter(4), lISize, 4)
    pvVfsWrite uDestFile, VarPtr(gzFooter(0)), 8
    pvVfsSetEof uDestFile, uDestFile.FileName
    
    ' 9. Jika output ke byte array, assign hasilnya
    If bToByteArray Then
        Destination = baGzOut
    End If

    GzCompress = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then Call CoTaskMemFree(uBuf.OutBlock)
    If hCtx <> 0 Then Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
    pvVfsClose uSourceFile
    pvVfsClose uDestFile
    Exit Function
EH:
    m_sLastError = Err.Description
    GzCompress = False
    Resume QH
End Function

Private Function pvGzCompressZlib(ByVal Source As Variant, ByRef Destination As Variant, ByVal Level As ArchiveCompressionLevel) As Boolean
    '--- GzCompress menggunakan zlib1.dll (Dynamic Huffman, kompresi lebih baik)
    '--- windowBits = -15 (raw deflate, tanpa zlib header)
    Const FUNC_NAME As String = "GzCompress[zlib]"
    Dim uStrm As z_stream
    Dim baBuffer() As Byte, baOut() As Byte
    Dim lCrc32 As Long, cTotalSize As Currency
    Dim gzHeader(0 To 9) As Byte, gzFooter(0 To 7) As Byte
    Dim uSourceFile As UcsVfsFileType, uDestFile As UcsVfsFileType
    Dim lSize As Long, lRet As Long
    Dim bToByteArray As Boolean
    Dim baGzOut() As Byte
    Dim bDeflateInit As Boolean
    
    On Error GoTo EH
    
    ' 1. Validasi level
    Level = IIf(Level < acNoCompression, acNormalCompression, IIf(Level > acMaxCompression, acMaxCompression, Level))
    
    ' 2. Buka sumber
    uSourceFile = pvVfsOpen(Source)
    pvVfsSeek uSourceFile, 0, FILE_BEGIN
    
    ' 3. GZ Header (RFC 1952)
    gzHeader(0) = &H1F: gzHeader(1) = &H8B
    gzHeader(2) = 8                           ' CM: deflate
    gzHeader(3) = 0                           ' FLG: no extras
    gzHeader(8) = IIf(Level >= acMaxCompression, 2, IIf(Level <= acFastestCompression, 4, 0))
    gzHeader(9) = 11                          ' OS: NTFS/Windows
    
    ' 4. Siapkan output (hapus file lama agar tidak ada sisa data)
    bToByteArray = IsMissing(Destination) Or VarType(Destination) = (vbArray Or vbByte)
    If Not bToByteArray Then
        On Error Resume Next
        If Dir$(CStr(Destination)) <> "" Then Kill CStr(Destination)
        On Error GoTo EH
    End If
    If bToByteArray Then
        ReDim baGzOut(0 To LNG_IOBUF_SIZE - 1) As Byte
        uDestFile = pvVfsCreate(baGzOut)
    Else
        uDestFile = pvVfsCreate(Destination)
    End If
    pvVfsWrite uDestFile, VarPtr(gzHeader(0)), 10
    
    ' 5. Init deflate (windowBits = -15 → raw deflate tanpa zlib/gzip wrapper)
    lRet = zlib_deflateInit2(uStrm, CLng(Level), Z_DEFLATED, -15, 8, Z_DEFAULT_STRATEGY, ZLIB_VERSION, LenB(uStrm))
    If lRet <> Z_OK Then Err.Raise vbObjectError, FUNC_NAME, ERR_INIT_COMPRESSOR
    bDeflateInit = True
    
    ' 6. Streaming deflate
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baOut(0 To LNG_IOBUF_SIZE - 1) As Byte
    lCrc32 = 0  ' zlib crc32 mulai dari 0
    cTotalSize = 0
    
    Do
        lSize = pvVfsRead(uSourceFile, VarPtr(baBuffer(0)), LNG_IOBUF_SIZE)
        
        ' Hitung CRC32 dan total size
        If lSize > 0 Then
            lCrc32 = zlib_crc32(lCrc32, VarPtr(baBuffer(0)), lSize)
            cTotalSize = cTotalSize + lSize
        End If
        
        uStrm.next_in = VarPtr(baBuffer(0))
        uStrm.avail_in = lSize
        
        ' Flush mode: Z_FINISH jika ini chunk terakhir
        Dim lFlush As Long
        lFlush = IIf(lSize < LNG_IOBUF_SIZE, Z_FINISH, Z_NO_FLUSH)
        
        Do
            uStrm.next_out = VarPtr(baOut(0))
            uStrm.avail_out = LNG_IOBUF_SIZE
            
            lRet = zlib_deflate(uStrm, lFlush)
            If lRet < 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_COMPRESSING
            
            Dim lHave As Long
            lHave = LNG_IOBUF_SIZE - uStrm.avail_out
            If lHave > 0 Then
                pvVfsWrite uDestFile, VarPtr(baOut(0)), lHave
            End If
        Loop While uStrm.avail_out = 0
    Loop While lRet <> Z_STREAM_END
    
    ' 7. GZ Footer
    Dim lISize As Long
    lISize = pvToInt32(cTotalSize)
    Call CopyMemory(gzFooter(0), lCrc32, 4)
    Call CopyMemory(gzFooter(4), lISize, 4)
    pvVfsWrite uDestFile, VarPtr(gzFooter(0)), 8
    pvVfsSetEof uDestFile, uDestFile.FileName
    
    ' 8. Output
    If bToByteArray Then
        Destination = baGzOut
    End If
    
    pvGzCompressZlib = True
QH:
    On Error Resume Next
    If bDeflateInit Then zlib_deflateEnd uStrm
    pvVfsClose uSourceFile
    pvVfsClose uDestFile
    Exit Function
EH:
    m_sLastError = Err.Description
    pvGzCompressZlib = False
    Resume QH
End Function

Public Function GzDecompress(ByVal SourceGz As Variant, ByRef Destination As Variant) As Boolean
    ' Deskripsi:
    ' Mendekompresi sumber GZ ke tujuan menggunakan streaming (per-chunk).
    ' Mendukung file GZ dengan flag FEXTRA, FNAME, FCOMMENT, FHCRC (RFC 1952).
    ' Mendukung file .gz besar — tidak perlu memuat seluruh file ke memori.
    ' Jika zlib1.dll tersedia, menggunakan zlib inflate (mendukung Dynamic Huffman).
    '
    ' Parameter:
    ' SourceGz (Variant): Sumber data GZ yang akan didekompres. Bisa berupa:
    '   - String: Path lengkap ke file .gz (e.g., "C:\arsip.gz")
    '   - Byte Array: Data GZ dalam bentuk byte array di memori.
    '
    ' Destination (Variant, ByRef): Tujuan penyimpanan hasil dekompresi. Bisa berupa:
    '   - String: Path lengkap file tujuan (e.g., "C:\data.txt")
    '   - Byte Array: Variabel byte array kosong yang akan diisi dengan hasil dekompresi.
    '
    Const FUNC_NAME As String = "GzDecompress"
    Const LNG_HEADER_BUF As Long = 4096
    Dim hCtx As Long, lResult As Long
    Dim uBuf As UcsZlibBuffersType
    Dim lCrc32_Stored As Long, lSize_Stored As Long, lCrc32_Calculated As Long
    Dim uSourceFile As UcsVfsFileType, uDestFile As UcsVfsFileType
    Dim lDataOffset As Long
    Dim bFlags As Byte
    Dim lExtraLen As Integer
    Dim baHeader() As Byte
    Dim baBuffer() As Byte
    Dim baFooter(0 To 7) As Byte
    Dim cFileSize As Currency
    Dim cDataSize As Currency
    Dim cRead As Currency
    Dim cTotalOut As Currency
    Dim lSize As Long
    Dim lHeaderLen As Long
    Dim bToByteArray As Boolean
    Dim baDecompOut() As Byte
    
    On Error GoTo EH
    m_sLastError = ""
    
    ' 0. Jika zlib1.dll tersedia, gunakan jalur zlib (mendukung Dynamic Huffman)
    If pvZlibAvailable() Then
        GzDecompress = pvGzDecompressZlib(SourceGz, Destination)
        Exit Function
    End If
    
    ' === FALLBACK: Thunk INFLATE ===
    
    ' 1. Buka sumber GZ dan tentukan ukuran
    uSourceFile = pvVfsOpen(SourceGz)
    cFileSize = pvVfsSeek(uSourceFile, 0, FILE_END)
    If cFileSize < 18 Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_TOO_SHORT
    
    ' 2. Baca header saja (maks 4096 byte, cukup untuk header GZ terpanjang)
    lHeaderLen = LNG_HEADER_BUF
    If cFileSize < lHeaderLen Then lHeaderLen = pvToInt32(cFileSize)
    ReDim baHeader(0 To lHeaderLen - 1)
    pvVfsSeek uSourceFile, 0, FILE_BEGIN
    pvVfsRead uSourceFile, VarPtr(baHeader(0)), lHeaderLen
    
    ' 3. Validasi magic dan compression method
    If baHeader(0) <> &H1F Or baHeader(1) <> &H8B Or baHeader(2) <> 8 Then
        Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_INVALID_HEADER
    End If
    
    ' 4. Parse flag byte dan hitung offset data kompresi (RFC 1952)
    bFlags = baHeader(3)
    lDataOffset = 10
    '--- FEXTRA (bit 2)
    If (bFlags And &H4) <> 0 Then
        If lDataOffset + 2 > lHeaderLen Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_INVALID_HEADER
        Call CopyMemory(lExtraLen, baHeader(lDataOffset), 2)
        lDataOffset = lDataOffset + 2 + lExtraLen
    End If
    '--- FNAME (bit 3): null-terminated string
    If (bFlags And &H8) <> 0 Then
        Do While lDataOffset < lHeaderLen
            If baHeader(lDataOffset) = 0 Then
                lDataOffset = lDataOffset + 1: Exit Do
            End If
            lDataOffset = lDataOffset + 1
        Loop
    End If
    '--- FCOMMENT (bit 4): null-terminated string
    If (bFlags And &H10) <> 0 Then
        Do While lDataOffset < lHeaderLen
            If baHeader(lDataOffset) = 0 Then
                lDataOffset = lDataOffset + 1: Exit Do
            End If
            lDataOffset = lDataOffset + 1
        Loop
    End If
    '--- FHCRC (bit 1): 2-byte CRC16
    If (bFlags And &H2) <> 0 Then
        lDataOffset = lDataOffset + 2
    End If
    If lDataOffset >= pvToInt32(cFileSize) - 8 Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_INVALID_HEADER
    
    ' 5. Baca footer (8 byte terakhir): CRC32 + ISIZE
    pvVfsSeek uSourceFile, cFileSize - 8, FILE_BEGIN
    pvVfsRead uSourceFile, VarPtr(baFooter(0)), 8
    Call CopyMemory(lCrc32_Stored, baFooter(0), 4)
    Call CopyMemory(lSize_Stored, baFooter(4), 4)
    
    ' 6. Hitung ukuran data kompresi
    cDataSize = cFileSize - lDataOffset - 8
    
    ' 7. Siapkan output — file atau byte array
    bToByteArray = IsMissing(Destination) Or VarType(Destination) = (vbArray Or vbByte)
    If bToByteArray Then
        ReDim baDecompOut(0 To LNG_IOBUF_SIZE - 1) As Byte
        uDestFile = pvVfsCreate(baDecompOut)
    Else
        uDestFile = pvVfsCreate(Destination)
    End If
    
    ' 8. Inisialisasi INFLATE decompressor
    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
    If hCtx = 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_INIT_DECOMPRESSOR
    
    ' 9. Streaming decompression — baca per-chunk, dekompresi, tulis langsung
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    pvVfsSeek uSourceFile, CCur(lDataOffset), FILE_BEGIN
    lCrc32_Calculated = -1
    cRead = 0
    cTotalOut = 0
    
    Do While cRead < cDataSize
        lSize = LNG_IOBUF_SIZE
        If cDataSize - cRead < lSize Then
            lSize = pvToInt32(cDataSize - cRead)
        End If
        lResult = pvVfsRead(uSourceFile, VarPtr(baBuffer(0)), lSize)
        If lResult = 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_DECOMPRESSING
        cRead = cRead + lResult
        
        uBuf.InBlock = VarPtr(baBuffer(0))
        uBuf.InLen = lResult
        lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf), VarPtr(lCrc32_Calculated))
        If lResult = 0 Or uBuf.OutBlock = 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_DECOMPRESSING
        
        pvVfsWrite uDestFile, uBuf.OutBlock, uBuf.OutLen
        cTotalOut = cTotalOut + uBuf.OutLen
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    Loop
    
    pvVfsSetEof uDestFile, uDestFile.FileName
    
    ' 10. Validasi CRC32 dan ISIZE (mod 2^32, sesuai RFC 1952)
    lCrc32_Calculated = lCrc32_Calculated Xor -1
    If lCrc32_Calculated <> lCrc32_Stored Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_CRC_MISMATCH
    If pvToInt32(cTotalOut) <> lSize_Stored Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_SIZE_MISMATCH
    
    ' 11. Jika output ke byte array, assign hasilnya
    If bToByteArray Then
        Destination = baDecompOut
    End If
    
    GzDecompress = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then Call CoTaskMemFree(uBuf.OutBlock)
    If hCtx <> 0 Then Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
    pvVfsClose uSourceFile
    pvVfsClose uDestFile
    Exit Function
EH:
    m_sLastError = Err.Description
    GzDecompress = False
    Resume QH
End Function

Private Function pvGzDecompressZlib(ByVal SourceGz As Variant, ByRef Destination As Variant) As Boolean
    '--- GzDecompress menggunakan zlib1.dll (mendukung Dynamic Huffman blocks)
    '--- windowBits = -15 (raw inflate)
    Const FUNC_NAME As String = "GzDecompress[zlib]"
    Const LNG_HEADER_BUF As Long = 4096
    Dim uStrm As z_stream
    Dim baBuffer() As Byte, baOut() As Byte
    Dim lCrc32_Stored As Long, lSize_Stored As Long
    Dim lCrc32_Calculated As Long
    Dim uSourceFile As UcsVfsFileType, uDestFile As UcsVfsFileType
    Dim lDataOffset As Long
    Dim bFlags As Byte
    Dim lExtraLen As Integer
    Dim baHeader() As Byte
    Dim baFooter(0 To 7) As Byte
    Dim cFileSize As Currency
    Dim cDataSize As Currency
    Dim cRead As Currency
    Dim cTotalOut As Currency
    Dim lSize As Long, lRet As Long
    Dim lHeaderLen As Long
    Dim bToByteArray As Boolean
    Dim baDecompOut() As Byte
    Dim bInflateInit As Boolean
    
    On Error GoTo EH
    
    ' 1. Buka sumber dan baca header
    uSourceFile = pvVfsOpen(SourceGz)
    cFileSize = pvVfsSeek(uSourceFile, 0, FILE_END)
    If cFileSize < 18 Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_TOO_SHORT
    
    lHeaderLen = LNG_HEADER_BUF
    If cFileSize < lHeaderLen Then lHeaderLen = pvToInt32(cFileSize)
    ReDim baHeader(0 To lHeaderLen - 1)
    pvVfsSeek uSourceFile, 0, FILE_BEGIN
    pvVfsRead uSourceFile, VarPtr(baHeader(0)), lHeaderLen
    
    ' 2. Validasi magic
    If baHeader(0) <> &H1F Or baHeader(1) <> &H8B Or baHeader(2) <> 8 Then
        Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_INVALID_HEADER
    End If
    
    ' 3. Parse GZ header flags (RFC 1952) — sama dengan thunk path
    bFlags = baHeader(3)
    lDataOffset = 10
    If (bFlags And &H4) <> 0 Then
        If lDataOffset + 2 > lHeaderLen Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_INVALID_HEADER
        Call CopyMemory(lExtraLen, baHeader(lDataOffset), 2)
        lDataOffset = lDataOffset + 2 + lExtraLen
    End If
    If (bFlags And &H8) <> 0 Then
        Do While lDataOffset < lHeaderLen
            If baHeader(lDataOffset) = 0 Then lDataOffset = lDataOffset + 1: Exit Do
            lDataOffset = lDataOffset + 1
        Loop
    End If
    If (bFlags And &H10) <> 0 Then
        Do While lDataOffset < lHeaderLen
            If baHeader(lDataOffset) = 0 Then lDataOffset = lDataOffset + 1: Exit Do
            lDataOffset = lDataOffset + 1
        Loop
    End If
    If (bFlags And &H2) <> 0 Then lDataOffset = lDataOffset + 2
    If lDataOffset >= pvToInt32(cFileSize) - 8 Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_INVALID_HEADER
    
    ' 4. Baca footer
    pvVfsSeek uSourceFile, cFileSize - 8, FILE_BEGIN
    pvVfsRead uSourceFile, VarPtr(baFooter(0)), 8
    Call CopyMemory(lCrc32_Stored, baFooter(0), 4)
    Call CopyMemory(lSize_Stored, baFooter(4), 4)
    cDataSize = cFileSize - lDataOffset - 8
    
    ' 5. Siapkan output
    bToByteArray = IsMissing(Destination) Or VarType(Destination) = (vbArray Or vbByte)
    If bToByteArray Then
        ReDim baDecompOut(0 To LNG_IOBUF_SIZE - 1) As Byte
        uDestFile = pvVfsCreate(baDecompOut)
    Else
        uDestFile = pvVfsCreate(Destination)
    End If
    
    ' 6. Init inflate (windowBits = -15 → raw inflate)
    lRet = zlib_inflateInit2(uStrm, -15, ZLIB_VERSION, LenB(uStrm))
    If lRet <> Z_OK Then Err.Raise vbObjectError, FUNC_NAME, ERR_INIT_DECOMPRESSOR
    bInflateInit = True
    
    ' 7. Streaming inflate
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baOut(0 To LNG_IOBUF_SIZE - 1) As Byte
    pvVfsSeek uSourceFile, CCur(lDataOffset), FILE_BEGIN
    lCrc32_Calculated = 0
    cRead = 0
    cTotalOut = 0
    
    Do While cRead < cDataSize
        lSize = LNG_IOBUF_SIZE
        If cDataSize - cRead < lSize Then lSize = pvToInt32(cDataSize - cRead)
        lSize = pvVfsRead(uSourceFile, VarPtr(baBuffer(0)), lSize)
        If lSize = 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_DECOMPRESSING
        cRead = cRead + lSize
        
        uStrm.next_in = VarPtr(baBuffer(0))
        uStrm.avail_in = lSize
        
        Do
            uStrm.next_out = VarPtr(baOut(0))
            uStrm.avail_out = LNG_IOBUF_SIZE
            
            lRet = zlib_inflate(uStrm, Z_NO_FLUSH)
            If lRet < 0 Then Err.Raise vbObjectError, FUNC_NAME, ERR_DECOMPRESSING
            
            Dim lHave As Long
            lHave = LNG_IOBUF_SIZE - uStrm.avail_out
            If lHave > 0 Then
                lCrc32_Calculated = zlib_crc32(lCrc32_Calculated, VarPtr(baOut(0)), lHave)
                pvVfsWrite uDestFile, VarPtr(baOut(0)), lHave
                cTotalOut = cTotalOut + lHave
            End If
        Loop While uStrm.avail_out = 0
        
        If lRet = Z_STREAM_END Then Exit Do
    Loop
    
    pvVfsSetEof uDestFile, uDestFile.FileName
    
    ' 8. Validasi CRC32 dan ISIZE
    If lCrc32_Calculated <> lCrc32_Stored Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_CRC_MISMATCH
    If pvToInt32(cTotalOut) <> lSize_Stored Then Err.Raise vbObjectError, FUNC_NAME, ERR_GZ_SIZE_MISMATCH
    
    If bToByteArray Then
        Destination = baDecompOut
    End If
    
    pvGzDecompressZlib = True
QH:
    On Error Resume Next
    If bInflateInit Then zlib_inflateEnd uStrm
    pvVfsClose uSourceFile
    pvVfsClose uDestFile
    Exit Function
EH:
    m_sLastError = Err.Description
    pvGzDecompressZlib = False
    Resume QH
End Function

'=========================================================================
'
'   ZIP Methods
'
'=========================================================================

#If ImplCompress Then

Public Function AddFile( _
        File As Variant, _
        Optional Name As String, _
        Optional Comment As String, _
        Optional Password As String, _
        Optional EncrStrength As Long, _
        Optional Level As Long = -1) As Boolean
    Const FUNC_NAME     As String = "AddFile"
    Dim uFile           As UcsVfsFileType
    
    On Error GoTo EH
    pvSetError
    If m_lFileCount = 0 Then
        ReDim m_uFiles(0 To 2) As UcsFileInfo
    ElseIf m_lFileCount > UBound(m_uFiles) Then
        ReDim Preserve m_uFiles(0 To 2 * UBound(m_uFiles)) As UcsFileInfo
    End If
    uFile = pvVfsOpen(File)
    With m_uFiles(m_lFileCount)
        If LenB(Name) <> 0 Then
            .FileName = Name
        Else
            .FileName = Mid$(uFile.FileName, InStrRev(uFile.FileName, "\") + 1)
        End If
        If Right$(.FileName, 1) = "\" Then
            .Size = 0
            .Attributes = vbDirectory
        Else
            .Size = pvToInt64(uFile.Data.nFileSizeLow, uFile.Data.nFileSizeHigh)
            .Attributes = uFile.Data.dwFileAttributes
        End If
        .Comment = Comment
        .LastModified = pvFromFileTime(uFile.Data.ftLastWriteTime)
        .Extra = vbNullString
        If IsObject(File) Then
            '--- remove VT_BYREF if any
            Set .SourceFile = C_Obj(File)
        Else
            .SourceFile = File
        End If
        .Level = Level
        #If ImplCrypto Then
            .Password = Password
            .Aes.Strength = EncrStrength
        #Else
            If LenB(Password) <> 0 Then
                On Error GoTo 0
                Err.Raise vbObjectError, , ERR_NO_CRYPTO_COMPILED
            End If
        #End If ' ImplCrypto
    End With
    m_lFileCount = m_lFileCount + 1
    '--- success
    AddFile = True
QH:
    On Error Resume Next
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function CompressArchive( _
        ArchiveFile As Variant, _
        Optional Comment As String, _
        Optional ByVal Level As Long = 6, _
        Optional ByVal UseUtf8 As VbTriState = vbUseDefault, _
        Optional ByVal UseZip64 As VbTriState = vbUseDefault) As Boolean
    Const FUNC_NAME     As String = "CompressArchive"
    Dim baBuffer()      As Byte
    Dim baNext()        As Byte
    Dim uArchiveFile    As UcsVfsFileType
    Dim uLocal          As UcsLocalHeaderType
    Dim lIdx            As Long
    Dim uFile           As UcsVfsFileType
    Dim lSize           As Long
    Dim cHeaderOffset   As Currency
    Dim cBeginOffset    As Currency
    Dim hCtx            As Long
    Dim uBuf            As UcsZlibBuffersType
    Dim lResult         As Long
    Dim uCentral        As UcsCentralHeaderType
    Dim uEndHdr         As UcsEndHeaderType
    Dim lEntries        As Long
    Dim lComprLevel     As Long
    Dim bSkip           As Boolean
#If ImplCrypto Then
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim nPassVer        As Integer
    Dim lJdx            As Long
    Dim uAes            As UcsAesExtraDataType
#End If ' ImplCrypto
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baNext(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsCreate(ArchiveFile)
    uCentral.VerExt = LNG_VER_DEFAULT
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            lComprLevel = IIf(.Level >= 0, .Level, Level)
            '--- OPTIM-3: Skip DEFLATE untuk file yang sudah terkompresi
            If lComprLevel > 0 And pvIsPreCompressed(.FileName) Then
                lComprLevel = 0
            End If
            RaiseEvent BeforeCompress(lIdx, lComprLevel, bSkip, m_bCancel)
            If m_bCancel Then
                m_sLastError = ERR_USER_CANCEL
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            lComprLevel = Clamp(lComprLevel, 0, 9)
            uBuf.Greedy = (lComprLevel <= 4)
            uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), lComprLevel)
            uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), lComprLevel)
            .Offset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            On Error GoTo EH_Continue
            uFile = pvVfsOpen(.SourceFile)
            pvToDosDateTime .LastModified, uLocal.FDate, uLocal.FTime
            uLocal.Method = LNG_METHOD_DEFLATE
            uLocal.VerExt = LNG_VER_DEFLATE
            uLocal.Crc32 = -1
            uLocal.Ext.USize = 0
            uLocal.Ext.CSize = 0
            If UseUtf8 = vbUseDefault Then
                uLocal.Flags = IIf(pvFromOemString(pvToOemString(.FileName, m_lCodePage), m_lCodePage) <> .FileName, ucsZcfUseUtf8, 0)
            Else
                uLocal.Flags = IIf(UseUtf8 = vbTrue, ucsZcfUseUtf8, 0)
            End If
            #If ImplCrypto Then
                If LenB(.Password) <> 0 Then
                    uLocal.Flags = uLocal.Flags Or ucsZcfEncrypted
                    If .Aes.Strength = 0 Then
                        uLocal.Flags = uLocal.Flags Or ucsZcfHasDataDescriptor
                        If Not pvCryptoTradInit(ToUtf8Array(.Password)) Then
                            Err.Raise vbObjectError, , m_sLastError
                        End If
                        '--- encrypt 12-byte random header w/ last byte used for password check
                        pvCryptoGetRandom LNG_ENC_HEADER_SIZE, baSalt
                        Call CopyMemory(baSalt(LNG_ENC_HEADER_SIZE - 1), ByVal UnsignedAdd(VarPtr(uLocal.FDate), -1), 1)
                        If Not pvCryptoTradCrypt(baSalt) Then
                            Err.Raise vbObjectError, , m_sLastError
                        End If
                    Else
                        lSaltSize = (.Aes.Strength + 1) * 4
                        pvCryptoGetRandom lSaltSize, baSalt
                        If Not pvCryptoAesInit(ToUtf8Array(.Password), baSalt, lSaltSize * 2, nPassVer) Then
                            Err.Raise vbObjectError, , m_sLastError
                        End If
                        '--- reserve space in extra data for AE-1/2
                        lJdx = UBound(.Extra) + 1
                        ReDim Preserve .Extra(0 To lJdx + LNG_EXTRADATA_AES_SIZE + 3) As Byte
                    End If
                End If
            #End If ' ImplCrypto
            cHeaderOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
            #If ImplCrypto Then
                If LenB(.Password) <> 0 Then
                    '--- prepend salt/enc_header before file data
                    pvVfsWrite uArchiveFile, VarPtr(baSalt(0)), UBound(baSalt) + 1
                    If .Aes.Strength > 0 Then
                        '--- prepend Password Verification Value for AES
                        pvVfsWrite uArchiveFile, VarPtr(nPassVer), 2
                    End If
                End If
            #End If ' ImplCrypto
            cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            If (.Attributes And vbDirectory + vbVolume) = 0 Then
                If lComprLevel > 0 Then
                    hCtx = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
                    If hCtx = 0 Then
                        Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
                    End If
                    lSize = pvVfsRead(uFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
                    Do
                        Call CopyMemory(baBuffer(0), baNext(0), lSize)
                        uBuf.InBlock = VarPtr(baBuffer(0))
                        uBuf.InLen = lSize
                        lSize = pvVfsRead(uFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
                        uBuf.Final = (lSize = 0)
                        lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf), VarPtr(uLocal.Crc32), lComprLevel) '--- level ignored
                        If lResult = 0 Or uBuf.OutBlock = 0 Then
                            Err.Raise vbObjectError, , ERR_COMPRESSING
                        End If
                        #If ImplCrypto Then
                            If LenB(.Password) <> 0 Then
                                If UBound(baBuffer) < uBuf.OutLen Then
                                    ReDim baBuffer(0 To (uBuf.OutLen And -2048) + 2047) As Byte
                                End If
                                Call CopyMemory(baBuffer(0), ByVal uBuf.OutBlock, uBuf.OutLen)
                                If .Aes.Strength = 0 Then
                                    If Not pvCryptoTradCrypt(baBuffer, Size:=uBuf.OutLen) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                Else
                                    If Not pvCryptoAesCrypt(baBuffer, Size:=uBuf.OutLen, HashAfter:=True) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                End If
                                pvVfsWrite uArchiveFile, VarPtr(baBuffer(0)), uBuf.OutLen
                                Call CoTaskMemFree(uBuf.OutBlock)
                                uBuf.OutBlock = 0
                            End If
                        #End If ' ImplCrypto
                        If uBuf.OutBlock <> 0 Then
                            pvVfsWrite uArchiveFile, uBuf.OutBlock, uBuf.OutLen
                            Call CoTaskMemFree(uBuf.OutBlock)
                            uBuf.OutBlock = 0
                        End If
                        uLocal.Ext.USize = uLocal.Ext.USize + uBuf.InLen
                        uLocal.Ext.CSize = uLocal.Ext.CSize + uBuf.OutLen
                        .CompressedSize = uLocal.Ext.CSize
                        RaiseEvent Progress(lIdx, uLocal.Ext.USize, .Size, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                    Loop While lSize <> 0
                    If hCtx <> 0 Then
                        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                        hCtx = 0
                    End If
                    If uLocal.Ext.CSize = 0 And uLocal.Ext.USize = 0 Then
                        uLocal.Method = LNG_METHOD_STORE
                        uLocal.VerExt = LNG_VER_DEFAULT
                    End If
                End If
                '--- if data is incompressible and not encrypted -> store file
                If lComprLevel = 0 Or (uLocal.Ext.CSize > uLocal.Ext.USize And (uLocal.Flags And ucsZcfEncrypted) = 0) Then
                    pvVfsSeek uFile, 0, FILE_BEGIN
                    pvVfsSeek uArchiveFile, cBeginOffset, FILE_BEGIN
                    uLocal.Method = LNG_METHOD_STORE
                    uLocal.VerExt = LNG_VER_DEFAULT
                    uLocal.Crc32 = -1
                    uLocal.Ext.USize = 0
                    uLocal.Ext.CSize = 0
                    #If ImplCrypto Then
                        If LenB(.Password) <> 0 Then
                            If .Aes.Strength = 0 Then
                                pvVfsSeek uArchiveFile, cHeaderOffset, FILE_BEGIN
                                uLocal.Flags = uLocal.Flags And Not ucsZcfEncrypted And Not ucsZcfHasDataDescriptor
                                pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
                                .Password = vbNullString
                            Else
                                If Not pvCryptoAesInit(ToUtf8Array(.Password), baSalt, lSaltSize * 2, nPassVer) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            End If
                        End If
                    #End If ' ImplCrypto
                    Do
                        RaiseEvent Progress(lIdx, uLocal.Ext.USize, .Size, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                        lSize = pvVfsRead(uFile, VarPtr(baBuffer(0)), LNG_IOBUF_SIZE)
                        If lSize = 0 Then
                            Exit Do
                        End If
                        Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lSize, VarPtr(uLocal.Crc32))
                        #If ImplCrypto Then
                            If LenB(.Password) <> 0 Then
                                If .Aes.Strength = 0 Then
                                    If Not pvCryptoTradCrypt(baBuffer, Size:=lSize) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                Else
                                    If Not pvCryptoAesCrypt(baBuffer, Size:=lSize, HashAfter:=True) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                End If
                            End If
                        #End If ' ImplCrypto
                        pvVfsWrite uArchiveFile, VarPtr(baBuffer(0)), lSize
                        uLocal.Ext.USize = uLocal.Ext.USize + lSize
                        uLocal.Ext.CSize = uLocal.Ext.CSize + lSize
                    Loop
                End If
            End If
            pvVfsClose uFile
            uLocal.Crc32 = uLocal.Crc32 Xor -1
            #If ImplCrypto Then
                If LenB(.Password) <> 0 Then
                    If .Aes.Strength = 0 Then
                        uLocal.Ext.CSize = uLocal.Ext.CSize + LNG_ENC_HEADER_SIZE
                    Else
                        '--- append HMAC-SHA1 authentication tag after file data
                        baSalt = pvCryptoAesGetFinalHash(LNG_AES_AUTHCODE_SIZE)
                        pvVfsWrite uArchiveFile, VarPtr(baSalt(0)), LNG_AES_AUTHCODE_SIZE
                        '--- update AE-1/2 in extra data
                        uAes.HeaderId = LNG_EXTRADATA_AES_HEADER
                        uAes.DataSize = LNG_EXTRADATA_AES_SIZE
                        uAes.Version = IIf(.Size <= LNG_EXTRADATA_AE_2_MAXSIZE, 2, 1)
                        uAes.VendorId = LNG_EXTRADATA_AES_VENDOR
                        uAes.Strength = lSaltSize / 4 - 1
                        uAes.Method = uLocal.Method
                        lJdx = UBound(.Extra) - LNG_EXTRADATA_AES_SIZE - 3
                        Debug.Assert VarPtr(uAes.Strength) + LenB(uAes.Strength) - VarPtr(uAes.HeaderId) = 9
                        Call CopyMemory(.Extra(lJdx + 0), uAes.HeaderId, 9)
                        Call CopyMemory(.Extra(lJdx + 9), uAes.Method, LenB(uAes.Method))
                        '--- update local header
                        uLocal.Method = LNG_METHOD_AES
                        uLocal.Ext.CSize = uLocal.Ext.CSize + lSaltSize + 2 + LNG_AES_AUTHCODE_SIZE
                        If uAes.Version = 2 Then
                            uLocal.Crc32 = 0
                        End If
                    End If
                End If
            #End If ' ImplCrypto
            .Crc32 = uLocal.Crc32
            Debug.Assert .Size = uLocal.Ext.USize
            .Size = uLocal.Ext.USize
            .CompressedSize = uLocal.Ext.CSize
            .Method = uLocal.Method
            .Flags = uLocal.Flags
            If (uLocal.Flags And ucsZcfHasDataDescriptor) <> 0 Then
                pvOutputDataDecriptor uArchiveFile, uLocal, UseZip64
                uLocal.Crc32 = 0
            End If
            cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
            pvVfsSeek uArchiveFile, cBeginOffset, FILE_BEGIN
            If uLocal.VerExt > uCentral.VerExt Then
                uCentral.VerExt = uLocal.VerExt
            End If
            If False Then
SkipFile:
                On Error GoTo EH
                pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
                .Offset = -1
                If hCtx <> 0 Then
                    Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                    hCtx = 0
                End If
            End If
            On Error GoTo EH
        End With
    Next
    m_lCurrentFile = -1
    cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
    uCentral.VerMade = LNG_VER_ZIP64
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            If .Offset >= 0 Then
                uCentral.Method = .Method
                pvToDosDateTime .LastModified, uCentral.FDate, uCentral.FTime
                uCentral.Crc32 = .Crc32
                uCentral.AttribX = .Attributes
                uCentral.Flags = .Flags
                uCentral.Ext.CSize = .CompressedSize
                uCentral.Ext.USize = .Size
                uCentral.Ext.Offset = .Offset
                pvOutputCentralHeader uArchiveFile, uCentral, .FileName, .Extra, .Comment, UseZip64
                lEntries = lEntries + 1
            End If
        End With
    Next
    m_lCurrentFile = -1
    If lEntries = 0 Then
        Err.Raise vbObjectError, , ERR_EMPTY_ARCHIVE
    End If
    uEndHdr.Ext.Entries = lEntries
    uEndHdr.Ext.TotEntr = lEntries
    uEndHdr.Ext.CenSize = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT) - cBeginOffset
    uEndHdr.Ext.CenOff = cBeginOffset
    pvOutputEndHeader uArchiveFile, uEndHdr, Comment, UseZip64
    pvVfsSetEof uArchiveFile, "[output_archive]"
    '--- success
    CompressArchive = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uFile
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description, CanContinue:=True) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

Public Function AddFromFolder( _
        sFolderAndMask As String, _
        Optional Recursive As Boolean, _
        Optional TargetFolder As String, _
        Optional IncludeEmptyFolders As Boolean, _
        Optional Password As String, _
        Optional EncrStrength As Long, _
        Optional Level As Long = -1) As Boolean
    Const FUNC_NAME     As String = "AddFromFolder"
    Dim lIdx            As Long
    Dim sFolder         As String
    Dim sMask           As String
    Dim lRootOffset     As Long
    Dim cFolders        As Collection
    Dim vElem           As Variant
    Dim sName           As String
    
    On Error GoTo EH
    pvSetError
    Set cFolders = New Collection
    lIdx = InStrRev(sFolderAndMask, "\")
    If lIdx > 0 Then
        sFolder = Left$(sFolderAndMask, lIdx - 1)
    End If
    sMask = Mid$(sFolderAndMask, lIdx + 1)
    If LenB(sFolder) = 0 Then
        sFolder = CurDir$()
    End If
    lRootOffset = Len(sFolder)
    cFolders.Add sFolder
    Do While cFolders.Count > 0
        sFolder = cFolders.Item(1)
        cFolders.Remove 1
        lIdx = 0
        For Each vElem In pvEnumFiles(sFolder, sMask, -vbDirectory)
            If AddFile(vElem, PathCombine(TargetFolder, Mid$(vElem, lRootOffset + 2)), Password:=Password, EncrStrength:=EncrStrength, Level:=Level) Then
                lIdx = lIdx + 1
                '--- success (entries added)
                AddFromFolder = True
            End If
        Next
        If lIdx = 0 And IncludeEmptyFolders Then
            sName = PathCombine(TargetFolder, Mid$(sFolder, lRootOffset + 2))
            If LenB(sName) <> 0 Then
                If AddFile(sFolder & "\", sName & "\", Password:=Password, EncrStrength:=EncrStrength, Level:=Level) Then
                    '--- success (entries added)
                    AddFromFolder = True
                End If
            End If
        End If
        If Recursive Then
            For Each vElem In pvEnumFiles(sFolder, "*.*", vbDirectory)
                cFolders.Add vElem
            Next
        End If
    Loop
QH:
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

#End If ' ImplCompress

#If ImplDecompress Then

Public Function OpenArchive(ArchiveFile As Variant) As Boolean
    Const FUNC_NAME     As String = "OpenArchive"
    Const MAX_END_SEEK  As Long = 10000
    Dim uArchiveFile    As UcsVfsFileType
    Dim uEndHdr         As UcsEndHeaderType
    Dim lIdx            As Long
    Dim sArchiveComment As String
    Dim uFiles()        As UcsFileInfo
    Dim uCentral        As UcsCentralHeaderType
    Dim baComment()     As Byte
    Dim cOffset         As Currency
    Dim baBuffer()      As Byte
    Dim uLocator        As UcsZip64EndLocatorType
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    uArchiveFile = pvVfsOpen(ArchiveFile)
    cOffset = pvVfsSeek(uArchiveFile, 0, FILE_END)
    If cOffset >= LNG_LEN_END Then
        For lIdx = 0 To IIf(cOffset - LNG_LEN_END < MAX_END_SEEK, cOffset - LNG_LEN_END, MAX_END_SEEK)
            cOffset = pvVfsSeek(uArchiveFile, -LNG_LEN_END - lIdx, FILE_END)
            Debug.Assert VarPtr(uEndHdr.LenCom) + 2 - VarPtr(uEndHdr.Signature) = LNG_LEN_END
            pvVfsRead uArchiveFile, VarPtr(uEndHdr), LNG_LEN_END
            If uEndHdr.Signature = LNG_SIG_END And uEndHdr.LenCom = lIdx Then
                If uEndHdr.LenCom > 0 Then
                    ReDim baComment(0 To uEndHdr.LenCom - 1) As Byte
                    pvVfsRead uArchiveFile, VarPtr(baComment(0)), uEndHdr.LenCom
                    sArchiveComment = pvFromOemString(baComment, m_lCodePage)
                End If
                If uEndHdr.CenOff = -1 Or uEndHdr.Ext.Entries = -1 Then
                    If cOffset <= LNG_LEN_LOCATOR + LNG_LEN_EOCDR Then
                        Exit For
                    End If
                    pvVfsSeek uArchiveFile, cOffset - LNG_LEN_LOCATOR, FILE_BEGIN
                    pvVfsRead uArchiveFile, VarPtr(uLocator), LNG_LEN_LOCATOR
                    If uLocator.Signature <> LNG_SIG_LOCATOR Or uLocator.DiskEocdr <> uEndHdr.DiskNum Then
                        Exit For
                    End If
                    pvVfsSeek uArchiveFile, uLocator.Offset * 10000@, FILE_BEGIN
                    pvVfsRead uArchiveFile, VarPtr(uEndHdr.Ext), LNG_LEN_EOCDR
                    If uEndHdr.Ext.Signature <> LNG_SIG_EOCDR Then
                        Exit For
                    End If
                    uEndHdr.Ext.CenOff = uEndHdr.Ext.CenOff * 10000@
                    uEndHdr.Ext.Entries = uEndHdr.Ext.Entries * 10000@
                Else
                    uEndHdr.Ext.CenOff = pvToInt64(uEndHdr.CenOff)
                    uEndHdr.Ext.Entries = uEndHdr.Entries And &HFFFF&
                End If
                lIdx = -1
                Exit For
            End If
        Next
    End If
    If lIdx < 0 Then
        '--- note: redim one more (last ignored)
        ReDim uFiles(0 To uEndHdr.Ext.Entries) As UcsFileInfo
        pvVfsSeek uArchiveFile, uEndHdr.Ext.CenOff, FILE_BEGIN
        For lIdx = 0 To uEndHdr.Ext.Entries - 1
            m_lCurrentFile = lIdx
            Debug.Assert VarPtr(uCentral.FDate) + LenB(uCentral.FDate) - VarPtr(uCentral.Signature) + VarPtr(uCentral.AttribI) + 2 - VarPtr(uCentral.Crc32) + VarPtr(uCentral.Offset) + 4 - VarPtr(uCentral.AttribX) = LNG_LEN_CENTRAL
            Debug.Assert VarPtr(uCentral.FDate) + LenB(uCentral.FDate) - VarPtr(uCentral.Signature) = 16
            pvVfsRead uArchiveFile, VarPtr(uCentral.Signature), 16
            Debug.Assert VarPtr(uCentral.AttribI) + LenB(uCentral.AttribI) - VarPtr(uCentral.Crc32) = 22
            pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), 22
            Debug.Assert VarPtr(uCentral.Offset) + LenB(uCentral.Offset) - VarPtr(uCentral.AttribX) = 8
            pvVfsRead uArchiveFile, VarPtr(uCentral.AttribX), 8
            If uCentral.Signature <> LNG_SIG_CENTRAL Then
                pvSetError MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", lIdx + 1)
                GoTo QH
            End If
            If Not pvLoadFileInfo(uArchiveFile, uCentral, uFiles(lIdx), lIdx + 1) Then
                GoTo QH
            End If
        Next
    Else
        '--- fall-back to sequential lookup from start of file
        pvVfsSeek uArchiveFile, 0, FILE_BEGIN
        uEndHdr.Ext.Entries = 0
        Do
            uCentral.Offset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            pvVfsRead uArchiveFile, VarPtr(uCentral.Signature), 4
            Select Case uCentral.Signature
            Case LNG_SIG_LOCAL
                pvVfsRead uArchiveFile, VarPtr(uCentral.VerExt), 10
                pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), 16
                uEndHdr.Ext.Entries = uEndHdr.Ext.Entries + 1
                If uEndHdr.Ext.Entries = 1 Then
                    ReDim uFiles(0 To 3) As UcsFileInfo
                ElseIf uEndHdr.Ext.Entries > UBound(uFiles) Then
                    ReDim Preserve uFiles(0 To 2 * uEndHdr.Ext.Entries) As UcsFileInfo
                End If
                If Not pvLoadFileInfo(uArchiveFile, uCentral, uFiles(uEndHdr.Ext.Entries - 1), uEndHdr.Ext.Entries) Then
                    GoTo QH
                End If
                If uCentral.Ext.CSize = 0 And uCentral.Ext.USize <> 0 Then
                    cOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
                    ReDim baBuffer(0 To uCentral.Ext.USize - LNG_LEN_EXTLOCAL) As Byte
                    pvVfsRead uArchiveFile, VarPtr(baBuffer(0)), UBound(baBuffer) + 1
                    For lIdx = 0 To UBound(baBuffer) - LNG_LEN_EXTLOCAL
                        If Peek(VarPtr(baBuffer(lIdx))) = LNG_SIG_EXTLOCAL Then
                            If Peek(VarPtr(baBuffer(lIdx + 4))) = uCentral.Crc32 Or uCentral.Crc32 = 0 Then
                                uCentral.Crc32 = Peek(VarPtr(baBuffer(lIdx + 4)))
                                uCentral.Ext.CSize = Peek(VarPtr(baBuffer(lIdx + 8)))
                                uFiles(uEndHdr.Ext.Entries - 1).Crc32 = uCentral.Crc32
                                uFiles(uEndHdr.Ext.Entries - 1).CompressedSize = uCentral.Ext.CSize
                                Exit For
                            End If
                        End If
                    Next
                    If uCentral.Ext.CSize = 0 Then
                        Exit Do
                    End If
                    pvVfsSeek uArchiveFile, cOffset, FILE_BEGIN
                End If
                pvVfsSeek uArchiveFile, uCentral.Ext.CSize, FILE_CURRENT
            Case LNG_SIG_EXTLOCAL
                If uEndHdr.Ext.Entries > 0 Then
                    pvVfsSeek uArchiveFile, LNG_LEN_EXTLOCAL - 4, FILE_CURRENT
                End If
            Case Else
                Exit Do
            End Select
        Loop
        '--- note: redim one more (last ignored)
        ReDim Preserve uFiles(0 To uEndHdr.Ext.Entries) As UcsFileInfo
    End If
    '--- commit archive info to member vars
    AssignVariant m_vArchiveFile, ArchiveFile
    m_lFileCount = uEndHdr.Ext.Entries
    m_sComment = sArchiveComment
    m_uFiles = uFiles
    '--- success
    OpenArchive = True
QH:
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function Extract( _
        OutputTarget As Variant, _
        Optional Filter As Variant, _
        Optional Password As String) As Boolean
    Const FUNC_NAME     As String = "Extract"
    Dim baBuffer()      As Byte
    Dim uArchiveFile    As UcsVfsFileType
    Dim lIdx            As Long
    Dim uLocal          As UcsLocalHeaderType
    Dim baFName()       As Byte
    Dim baExt()         As Byte
    Dim cRead           As Currency
    Dim lSize           As Long
    Dim lResult         As Long
    Dim hCtx            As Long
    Dim uBuf            As UcsZlibBuffersType
    Dim vFileName       As Variant
    Dim uFile           As UcsVfsFileType
    Dim lCrc32          As Long
    Dim pCrc32          As Long
    Dim bSkip           As Boolean
    Dim sFilterFolder   As String
    Dim sFilterMask     As String
    Dim sOutputFolder   As String
    Dim sOutputFile     As String
#If ImplCrypto Then
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim nPassVer        As Integer
#End If ' ImplCrypto
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    If IsEmpty(m_vArchiveFile) Then
        Err.Raise vbObjectError, , ERR_ARCHIVE_NOT_OPEN
    End If
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsOpen(m_vArchiveFile)
    If VarType(Filter) = vbString Then
        sFilterMask = Replace(Filter, "/", "\")
        lIdx = InStrRev(sFilterMask, "\")
        If lIdx > 0 Then
            sFilterFolder = Left$(sFilterMask, lIdx) & "*"
            sFilterMask = Mid$(sFilterMask, lIdx + 1)
        End If
    End If
    If VarType(OutputTarget) = vbString Then
        sOutputFolder = OutputTarget
        If LenB(sOutputFolder) <> 0 Then
            If (FileAttribs(sOutputFolder) And vbDirectory) = 0 Then
                sOutputFolder = pvPathGetFolder(sOutputFolder, sOutputFile)
                If (FileAttribs(sOutputFolder) And vbDirectory) = 0 Then
                    sOutputFolder = OutputTarget
                    sOutputFile = vbNullString
                End If
            End If
        End If
    End If
    '--- touch to open for read
    pvVfsRead uArchiveFile, 0, 0
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            '--- check if current entry is filtered
            If IsMissing(Filter) Then
                '--- do nothing
            ElseIf IsArray(Filter) Then
                If Not At(Filter, lIdx) Then
                    GoTo SkipFile
                End If
            ElseIf VarType(Filter) = vbString Then
                If LenB(sFilterFolder) <> 0 Then
                    If Not .FileName Like sFilterFolder Then
                        GoTo SkipFile
                    End If
                End If
                If LenB(sFilterMask) <> 0 Then
                    If PathMatchSpecW(StrPtr(Mid$(.FileName, InStrRev(.FileName, "\") + 1)), StrPtr(sFilterMask)) = 0 Then
                        GoTo SkipFile
                    End If
                End If
            ElseIf Not lIdx = Filter Then
                GoTo SkipFile
            End If
            On Error GoTo EH_Continue
            '--- read local header
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) + VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = LNG_LEN_LOCAL
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) = 14
            pvVfsRead uArchiveFile, VarPtr(uLocal.Signature), 14
            Debug.Assert VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = 16
            pvVfsRead uArchiveFile, VarPtr(uLocal.Crc32), 16
            Debug.Assert uLocal.LenFname > 0
            If uLocal.LenFname > 0 Then
                ReDim baFName(0 To uLocal.LenFname - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baFName(0)), uLocal.LenFname
            End If
            If uLocal.LenExt > 0 Then
                ReDim baExt(0 To uLocal.LenExt - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baExt(0)), uLocal.LenExt
            End If
            '--- sanity checks
            If uLocal.Signature <> LNG_SIG_LOCAL Then
                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", lIdx + 1)
            End If
'            If pvToWinFileName(baFName, .Flags) <> .FileName Then
'                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & " (Filename)"
'            End If
            If .Method <> LNG_METHOD_DEFLATE And .Method <> LNG_METHOD_STORE Then
                If pvSetError(MODULE_NAME & "." & FUNC_NAME, Replace(ERR_UNSUPPORTED_METHOD, "%1", .Method), CanContinue:=True) Then
                    GoTo QH
                Else
                    .Method = LNG_METHOD_STORE
                End If
            End If
            bSkip = False
            If VarType(OutputTarget) = vbString Then
                If LenB(sOutputFile) = 0 Then
                    vFileName = .FileName
                ElseIf InStr(sOutputFile, "*.*") > 0 Then
                    vFileName = Replace(sOutputFile, "*.*", .FileName)
                Else
                    vFileName = Replace(sOutputFile, "*", pvFileGetFName(.FileName))
                End If
                vFileName = PathCombine(sOutputFolder, CStr(vFileName))
                RaiseEvent BeforeExtract(lIdx, vFileName, bSkip, m_bCancel)
            Else
                RaiseEvent BeforeExtract(lIdx, OutputTarget, bSkip, m_bCancel)
            End If
            If m_bCancel Then
                m_sLastError = ERR_USER_CANCEL
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            '--- allow output to buffer, stream object or (customized) filename
            If VarType(OutputTarget) = vbString Then
                uFile = pvVfsCreate(vFileName)
            Else
                If IsArray(OutputTarget) And .Size > 0 Then
                    ReDim OutputTarget(0 To .Size - 1) As Byte
                End If
                uFile = pvVfsCreate(OutputTarget)
            End If
            If .CompressedSize > 0 Then
                lCrc32 = -1
                pCrc32 = VarPtr(lCrc32)
                cRead = 0
                If (.Flags And ucsZcfEncrypted) <> 0 Then
                    RaiseEvent Progress(lIdx, cRead, .CompressedSize, m_bCancel)
                    If m_bCancel Then
                        m_sLastError = ERR_USER_CANCEL
                        GoTo QH
                    End If
                    #If ImplCrypto Then
                        If (.Flags And ucsZcfStrongEncrypted) <> 0 Then
                            Err.Raise vbObjectError, , ERR_UNSUPPORTED_ENCRYPTION
                        End If
                        If LenB(Password) = 0 Then
                            Err.Raise vbObjectError, , ERR_PASSWORD_REQUIRED
                        End If
                        If .Aes.Strength = 0 Then
                            '--- init traditional ZipCrypt
                            If Not pvCryptoTradInit(ToUtf8Array(Password)) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), LNG_ENC_HEADER_SIZE)
                            If lResult = 0 Then
                                Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                            End If
                            If Not pvCryptoTradCrypt(baBuffer, 0, LNG_ENC_HEADER_SIZE, UseAfter:=True) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            If baBuffer(LNG_ENC_HEADER_SIZE - 1) <> .DecDat Then
                                Err.Raise vbObjectError, , ERR_INVALID_PASSWORD
                            End If
                            cRead = LNG_ENC_HEADER_SIZE
                        Else
                            '--- init WinZip AES
                            lSaltSize = (.Aes.Strength + 1) * 4
                            lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSaltSize + 2)
                            If lResult < lSaltSize + 2 Then
                                Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                            End If
                            ReDim baSalt(0 To lSaltSize - 1) As Byte
                            Call CopyMemory(baSalt(0), baBuffer(0), UBound(baSalt) + 1)
                            If Not pvCryptoAesInit(ToUtf8Array(Password), baSalt, lSaltSize * 2, nPassVer) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            If nPassVer <> PeekInt(VarPtr(baBuffer(lSaltSize))) Then
                                Err.Raise vbObjectError, , ERR_INVALID_PASSWORD
                            End If
                            '--- don't calc CRC32 for AE-2
                            If .Aes.Version = 2 Then
                                pCrc32 = 0
                            End If
                            '--- note: LNG_AES_AUTHCODE_SIZE chunk comes after file data
                            cRead = lSaltSize + 2 + LNG_AES_AUTHCODE_SIZE
                        End If
                    #Else
                        Err.Raise vbObjectError, , ERR_NO_CRYPTO_COMPILED
                    #End If ' ImplCrypto
                End If
                If .Method = LNG_METHOD_DEFLATE Then
                    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
                    If hCtx = 0 Then
                        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
                    End If
                End If
                Do ' While cRead < .CompressedSize
                    RaiseEvent Progress(lIdx, cRead, .CompressedSize, m_bCancel)
                    If m_bCancel Then
                        m_sLastError = ERR_USER_CANCEL
                        GoTo QH
                    End If
                    If cRead >= .CompressedSize Then
                        Exit Do
                    End If
                    lSize = LNG_IOBUF_SIZE
                    If lSize > .CompressedSize - cRead Then
                        lSize = .CompressedSize - cRead
                    End If
                    lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSize)
                    If lResult = 0 Then
                        Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                    End If
                    #If ImplCrypto Then
                        If (.Flags And ucsZcfEncrypted) <> 0 Then
                            If .Aes.Strength = 0 Then
                                If Not pvCryptoTradCrypt(baBuffer, 0, lResult, UseAfter:=True) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            Else
                                If Not pvCryptoAesCrypt(baBuffer, 0, lResult, HashBefore:=True) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            End If
                        End If
                    #End If ' ImplCrypto
                    cRead = cRead + lResult
                    If .Method = LNG_METHOD_DEFLATE Then
                        uBuf.InBlock = VarPtr(baBuffer(0))
                        uBuf.InLen = lResult
                        lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf), pCrc32)
                        If lResult = 0 Or uBuf.OutBlock = 0 Then
                            Err.Raise vbObjectError, , ERR_DECOMPRESSING
                        End If
                        pvVfsWrite uFile, uBuf.OutBlock, uBuf.OutLen
                        Call CoTaskMemFree(uBuf.OutBlock)
                        uBuf.OutBlock = 0
                    Else ' If .Method = LNG_METHOD_STORE Then
                        If pCrc32 <> 0 Then
                            Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lResult, pCrc32)
                        End If
                        pvVfsWrite uFile, VarPtr(baBuffer(0)), lResult
                    End If
                Loop
                pvVfsSetEof uFile, .FileName
                #If ImplCrypto Then
                    If .Aes.Strength <> 0 Then
                        '--- check HMAC-SHA1 auth code
                        ReDim baSalt(0 To LNG_AES_AUTHCODE_SIZE - 1) As Byte
                        lResult = pvVfsRead(uArchiveFile, VarPtr(baSalt(0)), LNG_AES_AUTHCODE_SIZE)
                        If lResult < LNG_AES_AUTHCODE_SIZE Then
                            Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                        End If
                        If Not pvArrayEqual(baSalt, pvCryptoAesGetFinalHash(LNG_AES_AUTHCODE_SIZE)) Then
                            Err.Raise vbObjectError, , ERR_INVALID_AUTHCODE
                        End If
                    End If
                #End If ' ImplCrypto
                If pCrc32 <> 0 Then
                    If .Crc32 <> (lCrc32 Xor -1) Then
                        Err.Raise vbObjectError, , ERR_CRC_CHECK
                    End If
                End If
            Else
                If (.Attributes And vbDirectory + vbVolume) = 0 And Right$(.FileName, 1) <> "\" Then
                    pvVfsSetEof uFile, .FileName
                End If
                '--- note: Total=1 to prevent division by zero in client code
                RaiseEvent Progress(lIdx, 1, 1, m_bCancel)
                If m_bCancel Then
                    m_sLastError = ERR_USER_CANCEL
                    GoTo QH
                End If
            End If
            If VarType(OutputTarget) = vbString Then
                RaiseEvent ExtractComplete(lIdx, vFileName)
            Else
                RaiseEvent ExtractComplete(lIdx, OutputTarget)
            End If
            '--- success (entries extracted)
            Extract = True
SkipFile:
            On Error GoTo EH
            pvVfsClose uFile
            If hCtx <> 0 Then
                Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
                hCtx = 0
            End If
        End With
    Next
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uArchiveFile
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description, CanContinue:=True) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

#End If ' ImplDecompress

#If ImplInflate Then

Public Function InflateBase64(sText As String, baOutput() As Byte) As Boolean
    InflateBase64 = Inflate(FromBase64Array(sText), baOutput)
End Function

Public Function Inflate(baBuffer() As Byte, baOutput() As Byte) As Boolean
    Const FUNC_NAME     As String = "Inflate"
    Dim hCtx            As Long
    Dim uBuf            As UcsZlibBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
    If hCtx = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
    End If
    uBuf.InBlock = VarPtr(baBuffer(0))
    uBuf.InLen = UBound(baBuffer) + 1
    lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_DECOMPRESSING
    End If
    ReDim baOutput(0 To uBuf.OutLen - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal uBuf.OutBlock, uBuf.OutLen)
    '--- success
    Inflate = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
        hCtx = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function DeflateBase64(baInput() As Byte, sText As String, Optional ByVal Level As Long = 6) As Boolean
    Dim baOutput()      As Byte
    
    DeflateBase64 = Deflate(baInput, baOutput, Level)
    If DeflateBase64 Then
        sText = ToBase64String(baOutput)
    End If
End Function

Public Function Deflate(baBuffer() As Byte, baOutput() As Byte, Optional ByVal Level As Long = 6) As Boolean
    Const FUNC_NAME     As String = "Deflate"
    Dim hCtx            As Long
    Dim lOutputPtr      As Long
    Dim lOutputSize     As Long
    
    On Error GoTo EH
    hCtx = DeflateInit()
    If Not DeflateBlob(hCtx, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, lOutputPtr, lOutputSize, Level, 1) Then
        GoTo QH
    End If
    ReDim baOutput(0 To lOutputSize - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal lOutputPtr, lOutputSize)
    '--- success
    Deflate = True
QH:
    On Error Resume Next
    If lOutputPtr <> 0 Then
        Call CoTaskMemFree(lOutputPtr)
    End If
    If hCtx <> 0 Then
        DeflateEnd hCtx
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function DeflateInit() As Long
    DeflateInit = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
    If DeflateInit = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
    End If
End Function

Public Function DeflateBlob(ByVal hCtx As Long, ByVal lInputPtr As Long, ByVal lInputSize As Long, lOutputPtr As Long, lOutputSize As Long, Optional ByVal Level As Long = 6, Optional ByVal Final As Boolean) As Boolean
    Const FUNC_NAME     As String = "DeflateBlob"
    Dim uBuf            As UcsZlibBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    uBuf.Greedy = (Level <= 4)
    uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), Level)
    uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), Level)
    uBuf.InBlock = lInputPtr
    uBuf.InLen = lInputSize
    uBuf.Final = -Final
    lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_COMPRESSING
    End If
    '--- commit
    lOutputPtr = uBuf.OutBlock
    lOutputSize = uBuf.OutLen
    uBuf.OutBlock = 0
    '--- success
    DeflateBlob = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Sub DeflateEnd(ByVal hCtx As Long)
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
    End If
End Sub

#End If ' ImplInflate

'=========================================================================
'
'   Private Helper Methods
'
'=========================================================================

'== Private Crypto Helpers ==
#If ImplCrypto Then

Private Sub pvCryptoGetRandom(ByVal lSize As Long, baOutput() As Byte)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call RtlGenRandom(baOutput(0), lSize)
End Sub

Private Function pvCryptoAesInit(baPass() As Byte, baSalt() As Byte, ByVal lKeyLen As Long, nPassVer As Integer) As Boolean
    Dim baDerivedKey()  As Byte
    Dim lDummy          As Long '--- discarded
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- init member vars
    m_uCrypto.Nonce(0) = 0
    m_uCrypto.Nonce(1) = 0
    m_uCrypto.EncrData = vbNullString
    m_uCrypto.EncrPos = 0
    '--- generate RFC 2898 based derived key
    If m_uCrypto.hPbkdf2Alg = 0 Then
        '--- CNG API missing on XP
        On Error GoTo EH_Unsupported
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hPbkdf2Alg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_SHA1_ALGORITHM & ")"
            GoTo QH
        End If
        On Error GoTo 0
    End If
    ReDim baDerivedKey(0 To 2 * lKeyLen + 1) As Byte
    '--- PBKDF2 API missing on Vista
    On Error GoTo EH_Unsupported
    hResult = BCryptDeriveKeyPBKDF2(m_uCrypto.hPbkdf2Alg, baPass(0), UBound(baPass) + 1, baSalt(0), UBound(baSalt) + 1, 1000, 0, baDerivedKey(0), UBound(baDerivedKey) + 1, 0)
    If hResult < 0 Then
        sApiSource = "BCryptDeriveKeyPBKDF2"
        GoTo QH
    End If
    On Error GoTo 0
    '--- extract Password Verification Value from last 2 bytes of derived key
    Call CopyMemory(nPassVer, baDerivedKey(2 * lKeyLen), 2)
    '--- init AES key from first half of the derived key
    If m_uCrypto.hAesAlg = 0 Then
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hAesAlg, StrPtr(BCRYPT_AES_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), 0)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_AES_ALGORITHM & ")"
            GoTo QH
        End If
        hResult = BCryptGetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_OBJECT_LENGTH), m_uCrypto.AesKeyObjLen, 4, lDummy, 0)
        If hResult < 0 Then
            sApiSource = "BCryptGetProperty(" & BCRYPT_OBJECT_LENGTH & ")"
            GoTo QH
        End If
        hResult = BCryptSetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_CHAINING_MODE), StrPtr(BCRYPT_CHAIN_MODE_ECB), LenB(BCRYPT_CHAIN_MODE_ECB), 0)
        If hResult < 0 Then
            sApiSource = "BCryptSetProperty(" & BCRYPT_CHAINING_MODE & ")"
            GoTo QH
        End If
    End If
    If m_uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(m_uCrypto.hAesKey)
        m_uCrypto.hAesKey = 0
    End If
    ReDim m_uCrypto.AesKeyObjData(0 To m_uCrypto.AesKeyObjLen - 1) As Byte
    hResult = BCryptGenerateSymmetricKey(m_uCrypto.hAesAlg, m_uCrypto.hAesKey, m_uCrypto.AesKeyObjData(0), m_uCrypto.AesKeyObjLen, baDerivedKey(0), lKeyLen, 0)
    If hResult < 0 Then
        sApiSource = "BCryptGenerateSymmetricKey"
        GoTo QH
    End If
    '--- init HMAC key from second half of the derived key
    If m_uCrypto.hHmacAlg = 0 Then
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_SHA1_ALGORITHM & ")"
            GoTo QH
        End If
        hResult = BCryptGetProperty(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_HASH_LENGTH), m_uCrypto.HmacLen, 4, lDummy, 0)
        If hResult < 0 Then
            sApiSource = "BCryptGetProperty(" & BCRYPT_HASH_LENGTH & ")"
            GoTo QH
        End If
    End If
    If m_uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(m_uCrypto.hHmacHash)
        m_uCrypto.hHmacHash = 0
    End If
    hResult = BCryptCreateHash(m_uCrypto.hHmacAlg, m_uCrypto.hHmacHash, 0, 0, baDerivedKey(lKeyLen), lKeyLen, 0)
    If hResult < 0 Then
        sApiSource = "BCryptCreateHash"
        GoTo QH
    End If
    '--- success
    pvCryptoAesInit = True
    Exit Function
QH:
    If hResult < 0 Then
        m_sLastError = GetSystemMessage(hResult) & " [" & sApiSource & "]"
    ElseIf Err.LastDllError <> 0 Then
        m_sLastError = GetSystemMessage(Err.LastDllError)
    End If
    Exit Function
EH_Unsupported:
    m_sLastError = ERR_UNSUPPORTED_ENCRYPTION
End Function

Private Sub pvCryptoTerminate()
    If m_uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(m_uCrypto.hHmacHash)
        m_uCrypto.hHmacHash = 0
    End If
    If m_uCrypto.hHmacAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(m_uCrypto.hHmacAlg, 0)
        m_uCrypto.hHmacAlg = 0
    End If
    If m_uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(m_uCrypto.hAesKey)
        m_uCrypto.hAesKey = 0
    End If
    If m_uCrypto.hAesAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(m_uCrypto.hAesAlg, 0)
        m_uCrypto.hAesAlg = 0
    End If
End Sub

Private Function pvCryptoAesCrypt( _
        baData() As Byte, _
        Optional ByVal Offset As Long, _
        Optional ByVal Size As Long, _
        Optional ByVal HashBefore As Boolean, _
        Optional ByVal HashAfter As Boolean) As Boolean
    Dim lIdx            As Long
    Dim lPadSize        As Long
    Dim lDummy          As Long '--- discarded
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    If HashBefore Then
        hResult = BCryptHashData(m_uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult < 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    With m_uCrypto
        '--- reuse EncrData from prev call until next LNG_AES_BLOCK_SIZE boundary
        For lIdx = Offset To Offset + Size - 1
            If (.EncrPos And (LNG_AES_BLOCK_SIZE - 1)) = 0 Then
                Exit For
            End If
            baData(lIdx) = baData(lIdx) Xor .EncrData(.EncrPos)
            .EncrPos = .EncrPos + 1
        Next
        If lIdx < Offset + Size Then
            '--- pad remaining input size to LNG_AES_BLOCK_SIZE
            lPadSize = (Offset + Size - lIdx + LNG_AES_BLOCK_SIZE - 1) And -LNG_AES_BLOCK_SIZE
            If UBound(.EncrData) + 1 < lPadSize Then
                ReDim .EncrData(0 To lPadSize - 1) As Byte
            End If
            '--- encrypt incremental nonces in EncrData
            Call CallWindowProc(m_uRtbl.MemNonce, VarPtr(.EncrData(0)), VarPtr(.Nonce(0)), lPadSize)
            hResult = BCryptEncrypt(.hAesKey, .EncrData(0), lPadSize, 0, 0, 0, .EncrData(0), lPadSize, lDummy, 0)
            If hResult < 0 Then
                sApiSource = "BCryptEncrypt"
                GoTo QH
            End If
            '--- xor remaining input and leave anything extra of EncrData for reuse
            .EncrPos = Offset + Size - lIdx
            Call CallWindowProc(m_uRtbl.MemXor, VarPtr(.EncrData(0)), VarPtr(baData(lIdx)), .EncrPos)
        End If
    End With
    If HashAfter Then
        hResult = BCryptHashData(m_uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult < 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    '--- success
    pvCryptoAesCrypt = True
    Exit Function
QH:
    If hResult < 0 Then
        m_sLastError = GetSystemMessage(hResult) & " [" & sApiSource & "]"
    ElseIf Err.LastDllError <> 0 Then
        m_sLastError = GetSystemMessage(Err.LastDllError)
    End If
End Function

Private Function pvCryptoAesGetFinalHash(ByVal lSize As Long) As Byte()
    Dim baResult()      As Byte
    
    ReDim baResult(0 To m_uCrypto.HmacLen - 1) As Byte
    Call BCryptFinishHash(m_uCrypto.hHmacHash, baResult(0), m_uCrypto.HmacLen, 0)
    ReDim Preserve baResult(0 To lSize - 1) As Byte
    pvCryptoAesGetFinalHash = baResult
End Function

Private Function pvCryptoTradInit(baPass() As Byte) As Boolean
    With m_uCrypto
        .TradKey(0) = &H12345678
        .TradKey(1) = &H23456789
        .TradKey(2) = &H34567890
        .TradKey(3) = m_uRtbl.Crc32Table
        Call CallWindowProc(m_uRtbl.ZipCrypt, VarPtr(.TradKey(0)), VarPtr(baPass(0)), UBound(baPass) + 1)
    End With
    '--- success
    pvCryptoTradInit = True
End Function

Private Function pvCryptoTradCrypt( _
        baData() As Byte, _
        Optional ByVal Offset As Long, _
        Optional ByVal Size As Long = -1, _
        Optional ByVal UseAfter As Boolean) As Boolean
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    Call CallWindowProc(m_uRtbl.ZipCrypt, VarPtr(m_uCrypto.TradKey(0)), VarPtr(baData(Offset)), Size, IIf(UseAfter, 3, 1))
    '--- success
    pvCryptoTradCrypt = True
End Function

#End If ' ImplCrypto

'== VFS Helpers ==
Private Function pvVfsOpen(File As Variant) As UcsVfsFileType
    Dim hFind           As Long
    
    If IsArray(File) Then
        pvVfsOpen.FileName = STR_BUFFER
        If (Peek(VarPtr(File)) And VT_BYREF) <> 0 Then
            Call CopyMemory(pvVfsOpen.BufferArray, File, 16)
        Else
            pvVfsOpen.BufferArray = File
        End If
        pvVfsOpen.BufferBase = pvArrPtr(File)
        pvVfsOpen.BufferPtr = pvVfsOpen.BufferBase
        If pvVfsOpen.BufferBase <> 0 Then
            pvVfsOpen.BufferSize = UBound(File) + 1
        End If
        pvVfsOpen.Data.nFileSizeLow = pvVfsOpen.BufferSize
        pvVfsOpen.Data.dwFileAttributes = vbArchive
        pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
    ElseIf IsObject(File) Then
        pvVfsOpen.FileName = STR_STREAM
        Set pvVfsOpen.Stream = File
        If Not pvVfsOpen.Stream Is Nothing Then
            pvVfsOpen.Data.nFileSizeLow = pvVfsOpen.Stream.VfsSetFilePointer(0, FILE_END)
        End If
        pvVfsOpen.Data.dwFileAttributes = vbArchive
        pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
    Else
        pvVfsOpen.FileName = File
        If Right$(File, 1) = ":" Then
            pvVfsOpen.Data.dwFileAttributes = vbArchive
            pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
        ElseIf Right$(File, 1) <> "\" Then
            hFind = FindFirstFile(StrPtr(File), VarPtr(pvVfsOpen.Data))
        Else
            hFind = FindFirstFile(StrPtr(Left$(File, Len(File) - 1)), VarPtr(pvVfsOpen.Data))
        End If
        If hFind <> INVALID_HANDLE_VALUE Then
            Call FindClose(hFind)
        Else
            Err.Raise vbObjectError, "pvVfsOpen", GetSystemMessage(Err.LastDllError) & " (" & File & ")"
        End If
    End If
End Function

Private Function pvCreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Select Case LCase$(lpFileName)
    Case "stdin:"
        pvCreateFile = GetStdHandle(STD_INPUT_HANDLE)
    Case "stdout:"
        pvCreateFile = GetStdHandle(STD_OUTPUT_HANDLE)
    Case "stderr:"
        pvCreateFile = GetStdHandle(STD_ERROR_HANDLE)
    Case Else
        pvCreateFile = CreateFile(StrPtr(lpFileName), dwDesiredAccess, dwShareMode, NoSecurity, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    End Select
End Function

Private Function pvVfsCreate(File As Variant) As UcsVfsFileType
    Dim sError          As String
    
    If IsArray(File) Then
        pvVfsCreate.FileName = STR_BUFFER
        If (Peek(VarPtr(File)) And VT_BYREF) <> 0 Then
            Call CopyMemory(pvVfsCreate.BufferArray, File, 16)
        Else
            Err.Raise vbObjectError, "pvVfsCreate", ERR_INVALID_BYTEARRAY
        End If
        pvVfsCreate.BufferBase = pvArrPtr(pvVfsCreate.BufferArray)
        pvVfsCreate.BufferPtr = pvVfsCreate.BufferBase
        If pvVfsCreate.BufferBase <> 0 Then
            pvVfsCreate.BufferSize = UBound(pvVfsCreate.BufferArray) + 1
        End If
    ElseIf IsObject(File) Then
        pvVfsCreate.FileName = STR_STREAM
        Set pvVfsCreate.Stream = File
    Else
        pvVfsCreate.FileName = File
        If InStrRev(File, "\") > 0 Then
            If Not MkPath(Left$(File, InStrRev(File, "\") - 1), sError) Then
                Err.Raise vbObjectError, , sError
            End If
        End If
        If Right$(File, 1) <> "\" Then
            pvVfsCreate.Handle = pvCreateFile(File, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive, 0)
            If pvVfsCreate.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "pvVfsCreate", GetSystemMessage(Err.LastDllError) & " (" & File & ")"
            End If
        End If
    End If
End Function

Private Sub pvVfsClose(uFile As UcsVfsFileType)
    Dim vEmpty          As Variant
    
    If Not IsEmpty(uFile.BufferArray) Then
        If (Peek(VarPtr(uFile.BufferArray)) And VT_BYREF) <> 0 Then
            Call CopyMemory(uFile.BufferArray, vEmpty, 16)
        Else
            uFile.BufferArray = Empty
        End If
    End If
    If uFile.Handle <> 0 And uFile.Handle <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(uFile.Handle)
        uFile.Handle = 0
    End If
    Set uFile.Stream = Nothing
End Sub

Private Function pvVfsRead(uFile As UcsVfsFileType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Not IsEmpty(uFile.BufferArray) Then
        pvVfsRead = Clamp(lSize, 0, uFile.BufferSize - (uFile.BufferPtr - uFile.BufferBase))
        Call CopyMemory(ByVal lPtr, ByVal uFile.BufferPtr, pvVfsRead)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsRead
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsRead = uFile.Stream.VfsReadFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = pvCreateFile(uFile.FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "pvVfsRead", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If ReadFile(uFile.Handle, ByVal lPtr, lSize, pvVfsRead, 0) = 0 Then
            If Err.LastDllError <> LNG_PIPE_ENDED_ERROR Then
                Err.Raise vbObjectError, "pvVfsRead", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
    End If
End Function

Private Function pvVfsWrite(uFile As UcsVfsFileType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const MAX_STEP      As Long = 2 ^ 24 '--- 16MB
    Dim lOffset         As Long
    
    If Not IsEmpty(uFile.BufferArray) Then
        lOffset = uFile.BufferPtr - uFile.BufferBase
        If lOffset + lSize > uFile.BufferSize Then
            If lOffset + lSize <= MAX_STEP Then
                uFile.BufferSize = Clamp(2 ^ Int(Log(lOffset + lSize) / Log(2) + 1), 8192)
            Else
                uFile.BufferSize = (lOffset + lSize + MAX_STEP - 1) And -MAX_STEP
            End If
            ReDim Preserve uFile.BufferArray(0 To uFile.BufferSize - 1) As Byte
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + lOffset
        End If
        pvVfsWrite = lSize
        Call CopyMemory(ByVal uFile.BufferPtr, ByVal lPtr, pvVfsWrite)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsWrite
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsWrite = uFile.Stream.VfsWriteFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = pvCreateFile(uFile.FileName, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "pvVfsWrite", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If WriteFile(uFile.Handle, ByVal lPtr, lSize, pvVfsWrite, 0) = 0 Then
            Err.Raise vbObjectError, "pvVfsWrite", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Function

Private Function pvVfsSeek(uFile As UcsVfsFileType, ByVal cPosition As Currency, ByVal lMoveMethod As Long) As Currency
    If Not IsEmpty(uFile.BufferArray) Then
        Select Case lMoveMethod
        Case FILE_BEGIN
            pvVfsSeek = cPosition
        Case FILE_CURRENT
            pvVfsSeek = uFile.BufferPtr - uFile.BufferBase + cPosition
        Case FILE_END
            pvVfsSeek = uFile.BufferSize + cPosition
        End Select
        uFile.BufferPtr = uFile.BufferBase + pvVfsSeek
    ElseIf Not uFile.Stream Is Nothing Then
        pvVfsSeek = uFile.Stream.VfsSetFilePointer(cPosition, lMoveMethod)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            '--- touch to open for read
            pvVfsRead uFile, 0, 0
        End If
        If SetFilePointerEx(uFile.Handle, cPosition / 10000@, pvVfsSeek, lMoveMethod) = 0 Then
            Err.Raise vbObjectError, "pvVfsSeek", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
        pvVfsSeek = pvVfsSeek * 10000@
    End If
End Function

Private Sub pvVfsSetEof(uFile As UcsVfsFileType, sMetaData As String)
    If Not IsEmpty(uFile.BufferArray) Then
        If uFile.BufferSize <> uFile.BufferPtr - uFile.BufferBase Then
            uFile.BufferSize = uFile.BufferPtr - uFile.BufferBase
            If uFile.BufferSize > 0 Then
                ReDim Preserve uFile.BufferArray(0 To uFile.BufferSize - 1) As Byte
            Else
                Erase uFile.BufferArray
            End If
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + uFile.BufferSize
        End If
    ElseIf Not uFile.Stream Is Nothing Then
        uFile.Stream.VfsSetEndOfFile sMetaData
    ElseIf Right$(uFile.FileName, 1) <> ":" Then
        If SetEndOfFile(uFile.Handle) = 0 Then
            Err.Raise vbObjectError, "pvVfsSetEof", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Sub

Private Function pvArrPtr(vArray As Variant) As Long
    Dim lPtr            As Long
    
    lPtr = Peek(UnsignedAdd(VarPtr(vArray), 8))
    If (Peek(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = Peek(lPtr)
    End If
    If lPtr <> 0 Then
        pvArrPtr = Peek(UnsignedAdd(lPtr, 12))
    End If
End Function

'== Common Helpers ==
Private Function pvFromOemString(baBuffer() As Byte, ByVal lCodePage As Long) As String
    Dim lSize           As Long
    
    If UBound(baBuffer) >= 0 Then
        lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, 0, 0)
        If lSize > 0 Then
            pvFromOemString = String$(lSize, 0)
            lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, ByVal StrPtr(pvFromOemString), lSize)
        End If
    End If
End Function

Private Function pvToOemString(sText As String, ByVal lCodePage As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    If Len(sText) <> 0 Then
        lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
        If lSize > 0 Then
            ReDim baRetVal(0 To lSize - 1) As Byte
            lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), UBound(baRetVal) + 1, 0, 0)
            pvToOemString = baRetVal
        Else
            pvToOemString = vbNullString
        End If
    Else
        pvToOemString = vbNullString
    End If
End Function

Private Function pvFromFileTime(uFileTime As FILETIME) As Date
    Dim uLocalTime      As FILETIME
    Dim uSysTime        As SYSTEMTIME
    
    If FileTimeToLocalFileTime(uFileTime, uLocalTime) <> 0 Then
        If FileTimeToSystemTime(uLocalTime, uSysTime) <> 0 Then
            Call SystemTimeToVariantTime(uSysTime, pvFromFileTime)
        End If
    End If
End Function

Private Function pvToFileTime(dDate As Date) As FILETIME
    Dim uSysTime        As SYSTEMTIME
    Dim uLocalTime      As FILETIME

    If VariantTimeToSystemTime(dDate, uSysTime) <> 0 Then
        If SystemTimeToFileTime(uSysTime, uLocalTime) <> 0 Then
            Call LocalFileTimeToFileTime(uLocalTime, pvToFileTime)
        End If
    End If
End Function

Private Function pvToInt64(ByVal lLowPart As Long, Optional ByVal lHiPart As Long) As Currency
    Call CopyMemory(pvToInt64, lLowPart, 4)
    If lHiPart <> 0 Then
        Call CopyMemory(ByVal UnsignedAdd(VarPtr(pvToInt64), 4), lHiPart, 4)
    End If
    pvToInt64 = pvToInt64 * 10000@
End Function

Private Function pvToInt32(ByVal cValue As Currency) As Long
    cValue = cValue / 10000@
    Call CopyMemory(pvToInt32, cValue, 4)
End Function

Private Function pvSetError(Optional Source As String, Optional Description As String, Optional ByVal CanContinue As Boolean) As Boolean
    If StrPtr(Source) = 0 And StrPtr(Description) = 0 Then
        m_sLastError = vbNullString
        m_bCancel = False
    Else
        #If ImplLogging Then
            If m_lCurrentFile >= 0 Then
                Debug.Print "Error in " & Source & ": " & Description, m_uFiles(m_lCurrentFile).FileName, Timer
            Else
                Debug.Print "Error in " & Source & ": " & Description, Timer
            End If
        #End If ' ImplLogging
        m_sLastError = Description
        If Not CanContinue Then
            m_bCancel = True
        End If
        RaiseEvent Error(m_lCurrentFile, Source, m_sLastError, m_bCancel)
        pvSetError = m_bCancel
    End If
End Function

Private Function pvPathGetFolder(sPath As String, Optional sFile As String) As String
    Dim lPos            As Long
    
    lPos = InStrRev(sPath, "\")
    If lPos > 0 Then
        pvPathGetFolder = Left$(sPath, lPos - 1)
        sFile = Mid$(sPath, lPos + 1)
    Else
        pvPathGetFolder = sPath
        sFile = vbNullString
    End If
End Function

Private Function pvFileGetFName(sFile As String) As String
    Dim lPos            As Long
    
    lPos = InStrRev(sFile, ".")
    If lPos > InStrRev(sFile, "\") Then
        pvFileGetFName = Left$(sFile, lPos - 1)
    Else
        pvFileGetFName = sFile
    End If
End Function

Private Function pvArrayEqual(baFirst() As Byte, baSecond() As Byte) As Boolean
    If UBound(baFirst) = UBound(baSecond) Then
        pvArrayEqual = (InStrB(baFirst, baSecond) = 1)
    End If
End Function

'== Thunk Initialization ==
Private Function pvInitRelocTable(uRtbl As UcsZlibRelocTableType) As Long
    Dim lpThunk         As Long
    Dim vSplit          As Variant
    
    lpThunk = pvGetThunkAddress()
    vSplit = Split(STR_THUNK_OFFSETS, "|")
    With uRtbl
        .CompressInit = lpThunk + vSplit(ucsIdx_CompressInit)
        .CompressCleanup = lpThunk + vSplit(ucsIdx_CompressCleanup)
        .CompressBlock = lpThunk + vSplit(ucsIdx_CompressBlock)
        .DecompressInit = lpThunk + vSplit(ucsIdx_DecompressInit)
        .DecompressCleanup = lpThunk + vSplit(ucsIdx_DecompressCleanup)
        .DecompressBlock = lpThunk + vSplit(ucsIdx_DecompressBlock)
        .CalcCrc32 = lpThunk + vSplit(ucsIdx_CalcCrc32)
        .MemNonce = lpThunk + vSplit(ucsIdx_MemNonce)
        .MemXor = lpThunk + vSplit(ucsIdx_MemXor)
        .ZipCrypt = lpThunk + vSplit(ucsIdx_ZipCrypt)
        .MallocImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemAlloc")
        .ReallocImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemRealloc")
        .FreeImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemFree")
        .LenCodes = lpThunk + vSplit(ucsIdx_LenCodes)
        .DistCodes = lpThunk + vSplit(ucsIdx_DistCodes)
        .MirrorBytes = lpThunk + vSplit(ucsIdx_MirrorBytes)
        .LenLenMap = lpThunk + vSplit(ucsIdx_LenLenMap)
        .Crc32Table = pvGetCrc32Table()
    End With
End Function

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    Dim sBuffer         As String
    
    If lpThunk = 0 Then
        sBuffer = String$(50, 0)
        '--- Menggunakan satu nama environment variable yang konsisten
        Call GetEnvironmentVariable(StrPtr("_ARCHIVE_THUNK_" & GetCurrentProcessId() & "_" & STR_THUNK_BUILDDATE), StrPtr(sBuffer), Len(sBuffer) - 1)
        lpThunk = Val(sBuffer)
        If lpThunk = 0 Then
            baThunk = FromBase64Array(STR_THUNK1 & STR_THUNK2)
            lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
            Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
            Call SetEnvironmentVariable(StrPtr("_ARCHIVE_THUNK_" & GetCurrentProcessId() & "_" & STR_THUNK_BUILDDATE), StrPtr(lpThunk))
        End If
    End If
    pvGetThunkAddress = lpThunk
End Function

Private Function pvGetCrc32Table() As Long
    Static aTable()      As Long
    Static bIsInit       As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lReminder       As Long
    Dim lValue          As Long

    If Not bIsInit Then
        '--- table mem allocated: 4KB
        ReDim aTable(0 To &H3FF) As Long
        For lIdx = 0 To &H3FF
            If lIdx < &H100 Then
                lReminder = 0
                lValue = lIdx
            Else
                lReminder = aTable(lIdx - &H100)
                lValue = 0
            End If
            For lJdx = 1 To 8
                If ((lReminder Xor lValue) And 1) <> 0 Then
                    lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF Xor &HEDB88320
                Else
                    lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF
                End If
                lValue = (lValue And &HFFFFFFFE) \ 2 And &H7FFFFFFF
            Next
            aTable(lIdx) = lReminder
        Next
        bIsInit = True
    End If
    pvGetCrc32Table = VarPtr(aTable(0))
End Function

Public Function CalcCrc32Array(baData() As Byte) As Long
    CalcCrc32Array = -1
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baData(0)), UBound(baData) + 1, VarPtr(CalcCrc32Array))
    CalcCrc32Array = CalcCrc32Array Xor -1
End Function

Public Sub CalcCrc32Ptr(ByVal lPtr As Long, ByVal lSize As Long, lCrc32 As Long)
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), lPtr, lSize, VarPtr(lCrc32))
End Sub

'== Shared Helpers ==
#If ImplUseShared = 0 Then

Private Function FileAttribs(sFile As String) As VbFileAttribute
    FileAttribs = GetFileAttributes(StrPtr(sFile))
    If FileAttribs = -1 Then
        FileAttribs = &H8000&
    End If
End Function

Private Function MkPath(sPath As String, sError As String) As Boolean
    Dim lAttrib         As Long
    
    lAttrib = GetFileAttributes(StrPtr(sPath))
    If lAttrib = -1 Then
        If InStrRev(sPath, "\") > 0 Then
            If Not MkPath(Left$(sPath, InStrRev(sPath, "\") - 1), sError) Then
                Exit Function
            End If
        End If
        If CreateDirectory(StrPtr(sPath), 0) = 0 Then
            sError = GetSystemMessage(Err.LastDllError) & " (" & sPath & ")"
            Exit Function
        End If
    ElseIf (lAttrib And vbDirectory + vbVolume) = 0 Then
        sError = "File already exists (" & sPath & ")"
        Exit Function
    End If
    '--- success
    MkPath = True
End Function

Private Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize           As Long
    
    GetSystemMessage = String$(2000, 0)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDllError, 0, StrPtr(GetSystemMessage), Len(GetSystemMessage), 0)
    If lSize > 2 Then
        If Mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, lSize) & " &H" & Hex$(lLastDllError)
End Function

Private Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, 0, lSize, 0, dwDummy)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    FromBase64Array = baOutput
End Function

Private Function ToBase64String(baInput() As Byte) As String
    Dim lSize As Long
    If (CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, 0&, lSize) <> 0) Then
        ToBase64String = String$(lSize - 1, 0)
        Call CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64String), lSize)
    End If
End Function

Private Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function

Private Function Clamp( _
        ByVal lValue As Long, _
        Optional ByVal Min As Long = -2147483647, _
        Optional ByVal Max As Long = 2147483647) As Long
    If lValue < Min Then
        Clamp = Min
    ElseIf lValue > Max Then
        Clamp = Max
    Else
        Clamp = lValue
    End If
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function PeekInt(ByVal lPtr As Long) As Integer
    Call CopyMemory(PeekInt, ByVal lPtr, 2)
End Function

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * (UBound(baText) + 1), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, ByVal StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function C_Obj(Value As Variant) As Object
    On Error Resume Next
    Set C_Obj = Value
    On Error GoTo 0
End Function

Private Function SanitizeFileName(sFileName As String, Optional InvalidChar As String = " ") As String
    Const STR_INVALID   As String = "[\x00-\x1F""<>\|:\*\?\\/]"
    
    With InitRegExp(STR_INVALID)
        SanitizeFileName = .Replace(sFileName, InvalidChar)
    End With
End Function

Private Function InitRegExp(sPattern As String) As Object
    Static oCache        As Object
    
    If Not oCache Is Nothing Then
        If oCache.Pattern = sPattern Then
            GoTo QH
        End If
    End If
    Set oCache = VBA.CreateObject("VBScript.RegExp")
    oCache.Global = True
    oCache.Pattern = sPattern
QH:
    Set InitRegExp = oCache
End Function

Private Function pvEnumFiles(sFolder As String, sMask As String, ByVal eAttrib As VbFileAttribute) As Collection
    Dim sFile           As String
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    
    On Error GoTo EH
    Set pvEnumFiles = New Collection
    sFile = PathCombine(sFolder, "*.*")
    hFind = FindFirstFile(StrPtr(sFile), VarPtr(uData))
    If hFind = INVALID_HANDLE_VALUE Then
        If Err.LastDllError <> 2 Then
            On Error GoTo 0
            Err.Raise vbObjectError, "pvEnumFiles", GetSystemMessage(Err.LastDllError) & " (" & sFile & ")"
        End If
    Else
        Do
            If eAttrib > 0 And (uData.dwFileAttributes And eAttrib) <> 0 _
                    Or eAttrib < 0 And (uData.dwFileAttributes And -eAttrib) = 0 Then
                sFile = Left$(uData.cFileName, InStr(uData.cFileName, Chr$(0)) - 1)
                If sFile <> "." And sFile <> ".." Then
                    If PathMatchSpecW(StrPtr(sFile), StrPtr(sMask)) <> 0 Then
                        pvEnumFiles.Add PathCombine(sFolder, sFile)
                    End If
                End If
            End If
        Loop While FindNextFile(hFind, VarPtr(uData)) <> 0
        Call FindClose(hFind)
        hFind = 0
    End If
    Exit Function
EH:
    If hFind <> 0 And hFind <> INVALID_HANDLE_VALUE Then
        Call FindClose(hFind)
        hFind = 0
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub pvOutputLocalHeader(uFile As UcsVfsFileType, uHdr As UcsLocalHeaderType, FileName As String, baExtra() As Byte, ByVal CSize As Currency, ByVal UseZip64 As VbTriState)
    Dim baFName()       As Byte
    
    If UseZip64 = vbTrue Or uHdr.Ext.CSize > LNG_SIZE_LIMIT Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or CSize > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        UseZip64 = vbTrue
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.CSize = -1
        uHdr.USize = -1
    Else
        uHdr.CSize = pvToInt32(uHdr.Ext.CSize)
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
    End If
    baFName = pvToOemString(Replace(FileName, "\", "/"), IIf((uHdr.Flags And ucsZcfUseUtf8) <> 0, CP_UTF8, m_lCodePage))
    uHdr.Signature = LNG_SIG_LOCAL
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = IIf(UseZip64 = vbTrue, 4 + LNG_EXTRADATA_ZIP64_LOCAL_SIZE, 0) + UBound(baExtra) + 1
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) + VarPtr(uHdr.LenExt) + 2 - VarPtr(uHdr.Crc32) = LNG_LEN_LOCAL
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) = 14
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 14
    Debug.Assert VarPtr(uHdr.LenExt) + LenB(uHdr.LenExt) - VarPtr(uHdr.Crc32) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 16
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
    If UseZip64 = vbTrue Then
        uHdr.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER
        uHdr.Ext.DataSize = LNG_EXTRADATA_ZIP64_LOCAL_SIZE
        uHdr.Ext.USize = uHdr.Ext.USize / 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize / 10000@
        Debug.Assert VarPtr(uHdr.Ext.CSize) + LenB(uHdr.Ext.CSize) - VarPtr(uHdr.Ext.HeaderId) = 20
        pvVfsWrite uFile, VarPtr(uHdr.Ext.HeaderId), 20
        uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize * 10000@
    End If
    If UBound(baExtra) >= 0 Then
        pvVfsWrite uFile, VarPtr(baExtra(0)), UBound(baExtra) + 1
    End If
End Sub

Private Sub pvOutputDataDecriptor(uFile As UcsVfsFileType, uHdr As UcsLocalHeaderType, ByVal UseZip64 As VbTriState)
    pvVfsWrite uFile, VarPtr(LNG_SIG_DATADESC), 4
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 4
    If UseZip64 = vbTrue Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or uHdr.Ext.CSize > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.Ext.USize = uHdr.Ext.USize / 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize / 10000@
        Debug.Assert VarPtr(uHdr.Ext.USize) + LenB(uHdr.Ext.USize) - VarPtr(uHdr.Ext.CSize) = 16
        pvVfsWrite uFile, VarPtr(uHdr.Ext.CSize), 16
        uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        uHdr.Ext.CSize = uHdr.Ext.CSize * 10000@
    Else
        uHdr.CSize = pvToInt32(uHdr.Ext.CSize)
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
        Debug.Assert VarPtr(uHdr.USize) + LenB(uHdr.USize) - VarPtr(uHdr.CSize) = 8
        pvVfsWrite uFile, VarPtr(uHdr.CSize), 8
        uHdr.CSize = 0
        uHdr.USize = 0
    End If
End Sub

Private Sub pvOutputCentralHeader(uFile As UcsVfsFileType, uHdr As UcsCentralHeaderType, FileName As String, baExtra() As Byte, Comment As String, ByVal UseZip64 As VbTriState)
    Dim lCodePage       As Long
    Dim baFName()       As Byte
    Dim baComment()     As Byte
    
    If UseZip64 = vbTrue Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or uHdr.Ext.CSize > LNG_SIZE_LIMIT Or uHdr.Ext.Offset > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        UseZip64 = vbTrue
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.Ext.DataSize = 0
    End If
    If uHdr.Ext.CSize > LNG_SIZE_LIMIT Then
        uHdr.CSize = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.CSize = pvToInt32(uHdr.Ext.CSize)
    End If
    If uHdr.Ext.USize > LNG_SIZE_LIMIT Then
        uHdr.USize = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
    End If
    If uHdr.Ext.Offset > LNG_SIZE_LIMIT Then
        uHdr.Offset = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.Offset = pvToInt32(uHdr.Ext.Offset)
    End If
    lCodePage = IIf((uHdr.Flags And ucsZcfUseUtf8) <> 0, CP_UTF8, m_lCodePage)
    baFName = pvToOemString(Replace(FileName, "\", "/"), lCodePage)
    baComment = pvToOemString(Comment, lCodePage)
    uHdr.Signature = LNG_SIG_CENTRAL
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = IIf(UseZip64 = vbTrue, 4 + uHdr.Ext.DataSize, 0) + UBound(baExtra) + 1
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) + VarPtr(uHdr.AttribI) + 2 - VarPtr(uHdr.Crc32) + VarPtr(uHdr.Offset) + 4 - VarPtr(uHdr.AttribX) = LNG_LEN_CENTRAL
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 16
    Debug.Assert VarPtr(uHdr.AttribI) + LenB(uHdr.AttribI) - VarPtr(uHdr.Crc32) = 22
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 22
    Debug.Assert VarPtr(uHdr.Offset) + LenB(uHdr.Offset) - VarPtr(uHdr.AttribX) = 8
    pvVfsWrite uFile, VarPtr(uHdr.AttribX), 8
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
    If UseZip64 = vbTrue Then
        uHdr.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER
        Debug.Assert VarPtr(uHdr.Ext.DataSize) + LenB(uHdr.Ext.DataSize) - VarPtr(uHdr.Ext.HeaderId) = 4
        pvVfsWrite uFile, VarPtr(uHdr.Ext.HeaderId), 4
        If uHdr.USize = -1 Then
            uHdr.Ext.USize = uHdr.Ext.USize / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.USize), LenB(uHdr.Ext.USize)
            uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        End If
        If uHdr.CSize = -1 Then
            uHdr.Ext.CSize = uHdr.Ext.CSize / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.CSize), LenB(uHdr.Ext.CSize)
            uHdr.Ext.CSize = uHdr.Ext.CSize * 10000@
        End If
        If uHdr.Offset = -1 Then
            uHdr.Ext.Offset = uHdr.Ext.Offset / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.Offset), LenB(uHdr.Ext.Offset)
            uHdr.Ext.Offset = uHdr.Ext.Offset * 10000@
        End If
    End If
    If UBound(baExtra) >= 0 Then
        pvVfsWrite uFile, VarPtr(baExtra(0)), UBound(baExtra) + 1
    End If
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvOutputEndHeader(uFile As UcsVfsFileType, uHdr As UcsEndHeaderType, Comment As String, ByVal UseZip64 As VbTriState)
    Dim baComment()     As Byte
    Dim uLocator        As UcsZip64EndLocatorType
    
    
    If UseZip64 = vbTrue Or uHdr.Ext.Entries > LNG_ENTRIES_LIMIT Or uHdr.Ext.TotEntr > LNG_ENTRIES_LIMIT Or uHdr.Ext.CenSize > LNG_SIZE_LIMIT Or uHdr.Ext.CenOff > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        uHdr.CenOff = -1
        '--- Zip64 End of Central Directory record
        uHdr.Ext.Signature = LNG_SIG_EOCDR
        uHdr.Ext.Size = (LNG_LEN_EOCDR - 12) / 10000@
        uHdr.Ext.VerMade = LNG_VER_ZIP64
        uHdr.Ext.VerExt = LNG_VER_ZIP64
        uHdr.Ext.Entries = uHdr.Ext.Entries / 10000@
        uHdr.Ext.TotEntr = uHdr.Ext.TotEntr / 10000@
        uHdr.Ext.CenSize = uHdr.Ext.CenSize / 10000@
        uHdr.Ext.CenOff = uHdr.Ext.CenOff / 10000@
        uLocator.Offset = pvVfsSeek(uFile, 0, FILE_CURRENT) / 10000@
        Debug.Assert VarPtr(uHdr.Ext.CenOff) + 8 - VarPtr(uHdr.Ext.Signature) = LNG_LEN_EOCDR
        pvVfsWrite uFile, VarPtr(uHdr.Ext.Signature), LNG_LEN_EOCDR
        uHdr.Ext.Entries = uHdr.Ext.Entries * 10000@
        uHdr.Ext.TotEntr = uHdr.Ext.TotEntr * 10000@
        uHdr.Ext.CenSize = uHdr.Ext.CenSize * 10000@
        uHdr.Ext.CenOff = uHdr.Ext.CenOff * 10000@
        '--- Zip64 End of Central Directory Locator record
        uLocator.Signature = LNG_SIG_LOCATOR
        Debug.Assert VarPtr(uLocator.DiskNum) + 4 - VarPtr(uLocator.Signature) = LNG_LEN_LOCATOR
        pvVfsWrite uFile, VarPtr(uLocator.Signature), LNG_LEN_LOCATOR
    Else
        uHdr.CenOff = pvToInt32(uHdr.Ext.CenOff)
    End If
    If uHdr.Ext.Entries > LNG_ENTRIES_LIMIT Then
        uHdr.Entries = -1
    Else
        uHdr.Entries = pvToInt32(uHdr.Ext.Entries) And &HFFFF&
    End If
    If uHdr.Ext.TotEntr > LNG_ENTRIES_LIMIT Then
        uHdr.TotEntr = -1
    Else
        uHdr.TotEntr = pvToInt32(uHdr.Ext.TotEntr) And &HFFFF&
    End If
    If uHdr.Ext.CenSize > LNG_SIZE_LIMIT Then
        uHdr.CenSize = -1
    Else
        uHdr.CenSize = pvToInt32(uHdr.Ext.CenSize)
    End If
    baComment = pvToOemString(Comment, m_lCodePage)
    uHdr.Signature = LNG_SIG_END
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.LenCom) + 2 - VarPtr(uHdr.Signature) = LNG_LEN_END
    pvVfsWrite uFile, VarPtr(uHdr.Signature), LNG_LEN_END
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvToDosDateTime(dDate As Date, nDate As Integer, nTime As Integer)
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME
    
    With uSysTime
        .wYear = Year(dDate)
        .wMonth = Month(dDate)
        .wDay = Day(dDate)
        .wHour = Hour(dDate)
        .wMinute = Minute(dDate)
        .wSecond = Second(dDate)
    End With
    Call SystemTimeToFileTime(uSysTime, uFileTime)
    Call FileTimeToDosDateTime(uFileTime, VarPtr(nDate), VarPtr(nTime))
End Sub

Private Function pvFromDosDateTime(ByVal nDate As Integer, ByVal nTime As Integer) As Date
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME

    If nDate <> 0 Or nTime <> 0 Then
        Call DosDateTimeToFileTime(nDate, nTime, uFileTime)
        Call FileTimeToSystemTime(uFileTime, uSysTime)
        With uSysTime
            pvFromDosDateTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        End With
    End If
End Function

Private Function pvToWinFileName(baBuffer() As Byte, ByVal lCodePage As Long) As String
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    '--- replaces non-ANSI symbols with '?' too
    vSplit = Split(Replace(StrConv(StrConv(pvFromOemString(baBuffer, lCodePage), _
        vbFromUnicode), vbUnicode), "/", "\"), "\")
    For lIdx = 0 To UBound(vSplit)
        vSplit(lIdx) = SanitizeFileName(CStr(vSplit(lIdx)), "_")
    Next
    pvToWinFileName = Join(vSplit, "\")
End Function

Private Function pvLoadFileInfo(uArchiveFile As UcsVfsFileType, uCentral As UcsCentralHeaderType, uFile As UcsFileInfo, ByVal lEntry As Long) As Boolean
    Const FUNC_NAME     As String = "pvLoadFileInfo"
    Dim baFName()       As Byte
    Dim baExtra()       As Byte
    Dim baComment()     As Byte
    Dim lCodePage       As Long
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim aSizes(0 To 2)  As Currency
    
    With uFile
        .Method = uCentral.Method
        .LastModified = pvFromDosDateTime(uCentral.FDate, uCentral.FTime)
        .Crc32 = uCentral.Crc32
        .Attributes = uCentral.AttribX
        .Flags = uCentral.Flags
        lCodePage = IIf((.Flags And ucsZcfUseUtf8) <> 0, CP_UTF8, m_lCodePage)
        Debug.Assert uCentral.LenFname > 0
        If uCentral.LenFname > 0 Then
            ReDim baFName(0 To uCentral.LenFname - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baFName(0)), uCentral.LenFname
            .FileName = pvToWinFileName(baFName, lCodePage)
        Else
            If pvSetError(MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_NO_FILENAME, "%1", lEntry), CanContinue:=True) Then
                GoTo QH
            End If
            .FileName = Replace(STR_ENTRY, "%1", lEntry)
        End If
        uCentral.Ext.USize = uCentral.USize
        uCentral.Ext.CSize = uCentral.CSize
        uCentral.Ext.Offset = uCentral.Offset
        If uCentral.LenExt > 0 Then
            ReDim baExtra(0 To uCentral.LenExt - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baExtra(0)), uCentral.LenExt
            .Extra = baExtra
            lPos = 0
            Do While lPos + 3 < uCentral.LenExt
                Call CopyMemory(uCentral.Ext.HeaderId, baExtra(lPos + 0), 4)
                If uCentral.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER Then
                    If uCentral.Ext.DataSize >= 8 Then
                        Call CopyMemory(aSizes(0), baExtra(lPos + 4), Clamp(uCentral.Ext.DataSize, 0, 24))
                        If uCentral.USize = -1 Then
                            uCentral.Ext.USize = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                        If uCentral.CSize = -1 Then
                            uCentral.Ext.CSize = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                        If uCentral.Offset = -1 Then
                            uCentral.Ext.Offset = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                    End If
                    Exit Do
                End If
                lPos = lPos + uCentral.Ext.DataSize + 4
            Loop
        End If
        .CompressedSize = uCentral.Ext.CSize
        .Size = uCentral.Ext.USize
        .Offset = uCentral.Ext.Offset
        If uCentral.LenCom > 0 Then
            ReDim baComment(0 To uCentral.LenCom - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baComment(0)), uCentral.LenCom
            .Comment = pvFromOemString(baComment, lCodePage)
        End If
        #If ImplCrypto Then
            If (.Flags And ucsZcfEncrypted) <> 0 Then
                If .Method = LNG_METHOD_AES Then
                    lPos = 0
                    Do While lPos + 3 < UBound(baExtra) + 1
                        Call CopyMemory(.Aes.HeaderId, baExtra(lPos + 0), 4)
                        If .Aes.HeaderId = LNG_EXTRADATA_AES_HEADER Then
                            If .Aes.DataSize >= LNG_EXTRADATA_AES_SIZE Then
                                Debug.Assert VarPtr(.Aes.Strength) + 1 - VarPtr(.Aes.HeaderId) = 9
                                Call CopyMemory(.Aes.HeaderId, baExtra(lPos + 0), 9)
                                Call CopyMemory(.Aes.Method, baExtra(lPos + 9), 2)
                                .Method = .Aes.Method
                            End If
                            Exit Do
                        End If
                        lPos = lPos + .Aes.DataSize + 4
                    Loop
                End If
                '--- store for traditional ZipCrypt password check
                If (.Flags And ucsZcfHasDataDescriptor) <> 0 Then
                    Call CopyMemory(.DecDat, ByVal UnsignedAdd(VarPtr(uCentral.FDate), -1), 1)
                Else
                    Call CopyMemory(.DecDat, ByVal UnsignedAdd(VarPtr(uCentral.CSize), -1), 1)
                End If
            End If
        #End If ' ImplCrypto
    End With
    '--- success
    pvLoadFileInfo = True
QH:
End Function

#End If ' Not ImplUseShared

'=========================================================================
' Pre-compressed file detection
'=========================================================================

Private Function pvIsPreCompressed(sFileName As String) As Boolean
    Dim lPos As Long
    Dim sExt As String
    
    lPos = InStrRev(sFileName, ".")
    If lPos = 0 Then Exit Function
    sExt = LCase$(Mid$(sFileName, lPos + 1))
    Select Case sExt
    '--- Gambar terkompresi
    Case "jpg", "jpeg", "png", "gif", "webp", "avif", "heic"
        pvIsPreCompressed = True
    '--- Video terkompresi
    Case "mp4", "mkv", "avi", "mov", "wmv", "webm", "flv"
        pvIsPreCompressed = True
    '--- Audio terkompresi
    Case "mp3", "aac", "ogg", "flac", "wma", "m4a", "opus"
        pvIsPreCompressed = True
    '--- Arsip terkompresi
    Case "zip", "rar", "7z", "gz", "bz2", "xz", "zst", "cab", "lzma", "lz4"
        pvIsPreCompressed = True
    '--- Dokumen terkompresi (Office/PDF)
    Case "docx", "xlsx", "pptx", "odt", "ods", "odp", "epub"
        pvIsPreCompressed = True
    End Select
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    pvInitRelocTable m_uRtbl
    m_lCurrentFile = -1
    m_lCodePage = CP_OEMCP
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    #If ImplCrypto Then
        pvCryptoTerminate
    #End If
    '--- lepas zlib1.dll
    If m_hZlib <> 0 Then
        FreeLibrary m_hZlib
        m_hZlib = 0
    End If
    '--- tutup arsip yang masih terbuka
    ArcClose
    '--- lepas archiveint.dll
    If m_hArcInt <> 0 Then
        FreeLibrary m_hArcInt
        m_hArcInt = 0
    End If
End Sub

'=========================================================================
'
'   Multi-Format Archive Methods (via archiveint.dll / libarchive)
'   Tersedia jika Windows 11 23H2+ (archiveint.dll di SysWOW64)
'
'=========================================================================

Public Property Get ArcAvailable() As Boolean
    ArcAvailable = pvArcAvailable()
End Property

Public Property Get ArcVersion() As String
    If pvArcAvailable() Then
        Dim lPtr As Long
        lPtr = arc_version_string()
        If lPtr <> 0 Then
            Dim lLen As Long
            lLen = lstrlenA(lPtr)
            If lLen > 0 And lLen < 100 Then
                Dim baVer() As Byte
                ReDim baVer(0 To lLen - 1)
                Call CopyMemory(baVer(0), ByVal lPtr, lLen)
                ArcVersion = StrConv(baVer, vbUnicode)
            End If
        End If
    End If
End Property

Public Function ArcOpen(ByVal ArchiveFile As String, Optional ByVal ArcType As ArchiveType = arAuto, Optional ByVal Password As String = "") As Boolean
    '--- Buka arsip multi-format untuk dibaca.
    '--- Otomatis deteksi format. Mendukung: ZIP, 7z, TAR, RAR, CAB, ISO, dll.
    Const FUNC_NAME As String = "ArcOpen"
    Dim hEntry As Long, lRet As Long
    Dim lPtr As Long, lLen As Long
    
    On Error GoTo EH
    pvSetError
    If Not pvArcAvailable() Then Err.Raise vbObjectError, FUNC_NAME, "archiveint.dll tidak tersedia"
    
    '--- Tutup arsip sebelumnya jika ada
    ArcClose
    
    '--- Buat handle baru
    m_hArcRead = arc_read_new()
    If m_hArcRead = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Gagal membuat archive reader"
    
    '--- Support semua format dan filter
    arc_read_support_format_all m_hArcRead
    arc_read_support_filter_all m_hArcRead
    
    '--- Set password jika ada
    If LenB(Password) <> 0 Then
        Dim baPw() As Byte
        baPw = ToUtf8Array(Password)
        ReDim Preserve baPw(0 To UBound(baPw) + 1)
        baPw(UBound(baPw)) = 0
        arc_read_add_passphrase m_hArcRead, VarPtr(baPw(0))
    End If
    
    '--- Buka file
    Dim baFn() As Byte
    baFn = ToUtf8Array(ArchiveFile)
    ReDim Preserve baFn(0 To UBound(baFn) + 1)
    baFn(UBound(baFn)) = 0
    lRet = arc_read_open_filename(m_hArcRead, VarPtr(baFn(0)), ARC_BLOCK_SIZE)
    If lRet <> ARC_OK Then
        Dim sErr As String
        sErr = pvArcErrorString(m_hArcRead)
        arc_read_free m_hArcRead: m_hArcRead = 0
        Err.Raise vbObjectError, FUNC_NAME, "Gagal membuka arsip: " & sErr
    End If
    
    '--- Scan semua entry
    m_lArcEntryCount = 0
    ReDim m_sArcEntries(0 To 255)
    ReDim m_lArcSizes(0 To 255)
    ReDim m_lArcTypes(0 To 255)
    
    Do
        lRet = arc_read_next_header(m_hArcRead, hEntry)
        If lRet = ARC_EOF Then Exit Do
        If lRet < ARC_OK And lRet <> ARC_WARN Then Exit Do
        
        '--- Grow arrays jika perlu
        If m_lArcEntryCount > UBound(m_sArcEntries) Then
            ReDim Preserve m_sArcEntries(0 To m_lArcEntryCount * 2)
            ReDim Preserve m_lArcSizes(0 To m_lArcEntryCount * 2)
            ReDim Preserve m_lArcTypes(0 To m_lArcEntryCount * 2)
        End If
        
        '--- Baca nama (wide string)
        lPtr = arc_entry_pathname_w(hEntry)
        If lPtr <> 0 Then
            m_sArcEntries(m_lArcEntryCount) = pvWideToString(lPtr)
        End If
        m_lArcSizes(m_lArcEntryCount) = arc_entry_size(hEntry)
        m_lArcTypes(m_lArcEntryCount) = arc_entry_filetype(hEntry)
        m_lArcEntryCount = m_lArcEntryCount + 1
    Loop
    
    '--- Tutup dan buka ulang (libarchive bersifat streaming)
    arc_read_close m_hArcRead
    arc_read_free m_hArcRead
    
    '--- Buka ulang untuk extraction nanti
    m_hArcRead = arc_read_new()
    arc_read_support_format_all m_hArcRead
    arc_read_support_filter_all m_hArcRead
    If LenB(Password) <> 0 Then arc_read_add_passphrase m_hArcRead, VarPtr(baPw(0))
    arc_read_open_filename m_hArcRead, VarPtr(baFn(0)), ARC_BLOCK_SIZE
    
    ArcOpen = True
QH:
    Exit Function
EH:
    m_sLastError = Err.Description
    ArcOpen = False
    Resume QH
End Function

Public Sub ArcClose()
    On Error Resume Next
    If m_hArcRead <> 0 Then
        arc_read_close m_hArcRead
        arc_read_free m_hArcRead
        m_hArcRead = 0
    End If
    If m_hArcWrite <> 0 Then
        arc_write_close m_hArcWrite
        arc_write_free m_hArcWrite
        m_hArcWrite = 0
    End If
    m_lArcEntryCount = 0
End Sub

Public Property Get ArcFileCount() As Long
    ArcFileCount = m_lArcEntryCount
End Property

Public Function ArcFileInfo(ByVal Index As Long) As Variant
    '--- Kembalikan Array(Name, Size, Type) untuk entry ke-Index
    If Index < 0 Or Index >= m_lArcEntryCount Then
        ArcFileInfo = Array("", 0, 0)
        Exit Function
    End If
    ArcFileInfo = Array(m_sArcEntries(Index), m_lArcSizes(Index), m_lArcTypes(Index))
End Function

Public Function ArcExtract(ByVal ArchiveFile As String, ByVal TargetFolder As String, _
        Optional ByVal Password As String = "") As Boolean
    '--- Ekstrak semua file dari arsip ke folder tujuan.
    Const FUNC_NAME As String = "ArcExtract"
    Dim hRead As Long, hEntry As Long
    Dim lRet As Long
    Dim sName As String, sOutPath As String
    Dim baBuffer() As Byte
    Dim iFile As Integer
    Dim lRead As Long
    
    On Error GoTo EH
    pvSetError
    If Not pvArcAvailable() Then Err.Raise vbObjectError, FUNC_NAME, "archiveint.dll tidak tersedia"
    
    '--- Pastikan folder tujuan ada
    If Right$(TargetFolder, 1) <> "\" Then TargetFolder = TargetFolder & "\"
    pvMkDirRecursive TargetFolder
    
    '--- Buka arsip
    hRead = arc_read_new()
    If hRead = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Gagal membuat archive reader"
    arc_read_support_format_all hRead
    arc_read_support_filter_all hRead
    
    If LenB(Password) <> 0 Then
        Dim baPw() As Byte
        baPw = ToUtf8Array(Password)
        ReDim Preserve baPw(0 To UBound(baPw) + 1)
        baPw(UBound(baPw)) = 0
        arc_read_add_passphrase hRead, VarPtr(baPw(0))
    End If
    
    Dim baFn() As Byte
    baFn = ToUtf8Array(ArchiveFile)
    ReDim Preserve baFn(0 To UBound(baFn) + 1)
    baFn(UBound(baFn)) = 0
    lRet = arc_read_open_filename(hRead, VarPtr(baFn(0)), ARC_BLOCK_SIZE)
    If lRet <> ARC_OK Then Err.Raise vbObjectError, FUNC_NAME, "Gagal membuka: " & pvArcErrorString(hRead)
    
    '--- Ekstrak semua entry
    ReDim baBuffer(0 To ARC_BLOCK_SIZE - 1)
    Do
        lRet = arc_read_next_header(hRead, hEntry)
        If lRet = ARC_EOF Then Exit Do
        If lRet < ARC_OK And lRet <> ARC_WARN Then Exit Do
        
        Dim lPtr As Long
        lPtr = arc_entry_pathname_w(hEntry)
        If lPtr = 0 Then GoTo NextEntry
        sName = pvWideToString(lPtr)
        If LenB(sName) = 0 Then GoTo NextEntry
        
        '--- Ganti "/" jadi "\"
        sName = Replace$(sName, "/", "\")
        sOutPath = TargetFolder & sName
        
        '--- Cek apakah directory
        Dim lType As Long
        lType = arc_entry_filetype(hEntry)
        If lType = ARC_IFDIR Then
            pvMkDirRecursive sOutPath
            GoTo NextEntry
        End If
        
        '--- Pastikan parent folder ada
        Dim sDir As String
        sDir = Left$(sOutPath, InStrRev(sOutPath, "\"))
        If LenB(sDir) > 0 Then pvMkDirRecursive sDir
        
        '--- Tulis file
        iFile = FreeFile
        Open sOutPath For Binary Access Write As #iFile
        Do
            lRead = arc_read_data(hRead, VarPtr(baBuffer(0)), ARC_BLOCK_SIZE)
            If lRead <= 0 Then Exit Do
            Dim baChunk() As Byte
            ReDim baChunk(0 To lRead - 1)
            CopyMemory baChunk(0), baBuffer(0), lRead
            Put #iFile, , baChunk
        Loop
        Close #iFile
        iFile = 0
        
NextEntry:
    Loop
    
    arc_read_close hRead
    arc_read_free hRead
    
    ArcExtract = True
QH:
    On Error Resume Next
    If iFile > 0 Then Close #iFile
    Exit Function
EH:
    m_sLastError = Err.Description
    ArcExtract = False
    On Error Resume Next
    If hRead <> 0 Then arc_read_close hRead: arc_read_free hRead
    Resume QH
End Function

Public Function ArcCreate(ByVal ArchiveFile As String, Optional ByVal ArcType As ArchiveType = arAuto, _
        Optional ByVal Level As ArchiveCompressionLevel = acNormalCompression) As Boolean
    '--- Buat arsip baru untuk menulis.
    Const FUNC_NAME As String = "ArcCreate"
    Dim lRet As Long
    
    On Error GoTo EH
    pvSetError
    If Not pvArcAvailable() Then Err.Raise vbObjectError, FUNC_NAME, "archiveint.dll tidak tersedia"
    
    '--- Tutup arsip sebelumnya
    ArcClose
    
    '--- Auto-detect format dari ekstensi
    If ArcType = arAuto Then ArcType = pvDetectArchiveType(ArchiveFile)
    
    '--- Buat writer
    m_hArcWrite = arc_write_new()
    If m_hArcWrite = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Gagal membuat archive writer"
    
    '--- Set format dan filter
    pvSetArcFormat m_hArcWrite, ArcType
    
    '--- Set compression level
    Dim baOpt() As Byte, sOpt As String
    sOpt = "compression-level=" & CStr(CLng(Level))
    baOpt = ToUtf8Array(sOpt)
    ReDim Preserve baOpt(0 To UBound(baOpt) + 1)
    baOpt(UBound(baOpt)) = 0
    arc_write_set_options m_hArcWrite, VarPtr(baOpt(0))
    
    '--- Buka file output
    Dim baFn() As Byte
    baFn = ToUtf8Array(ArchiveFile)
    ReDim Preserve baFn(0 To UBound(baFn) + 1)
    baFn(UBound(baFn)) = 0
    lRet = arc_write_open_filename(m_hArcWrite, VarPtr(baFn(0)))
    If lRet <> ARC_OK Then Err.Raise vbObjectError, FUNC_NAME, "Gagal membuat file: " & pvArcErrorString(m_hArcWrite)
    
    ArcCreate = True
QH:
    Exit Function
EH:
    m_sLastError = Err.Description
    ArcCreate = False
    Resume QH
End Function

Public Function ArcAddFile(ByVal FilePath As String, Optional ByVal ArcName As String = "") As Boolean
    '--- Tambahkan file ke arsip yang sedang ditulis.
    Const FUNC_NAME As String = "ArcAddFile"
    Dim hEntry As Long, lRet As Long
    Dim iFile As Integer, lFileSize As Long
    Dim baBuffer() As Byte
    Dim lRead As Long
    
    On Error GoTo EH
    If m_hArcWrite = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Arsip belum dibuka untuk menulis"
    
    '--- Tentukan nama di arsip
    If LenB(ArcName) = 0 Then
        ArcName = Mid$(FilePath, InStrRev(FilePath, "\") + 1)
    End If
    
    '--- Dapatkan ukuran file
    lFileSize = FileLen(FilePath)
    
    '--- Buat entry
    hEntry = arc_entry_new()
    If hEntry = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Gagal membuat entry"
    
    Dim baName() As Byte
    baName = ToUtf8Array(Replace$(ArcName, "\", "/"))
    ReDim Preserve baName(0 To UBound(baName) + 1)
    baName(UBound(baName)) = 0
    arc_entry_set_pathname_utf8 hEntry, VarPtr(baName(0))
    arc_entry_set_size hEntry, lFileSize
    arc_entry_set_filetype hEntry, ARC_IFREG
    arc_entry_set_perm hEntry, ARC_PERM_FILE
    
    lRet = arc_write_header(m_hArcWrite, hEntry)
    If lRet < ARC_OK And lRet <> ARC_WARN Then
        arc_entry_free hEntry
        Err.Raise vbObjectError, FUNC_NAME, "Gagal menulis header: " & pvArcErrorString(m_hArcWrite)
    End If
    
    '--- Tulis data file
    ReDim baBuffer(0 To ARC_BLOCK_SIZE - 1)
    iFile = FreeFile
    Open FilePath For Binary Access Read As #iFile
    Do While Not EOF(iFile)
        lRead = LOF(iFile) - Seek(iFile) + 1
        If lRead > ARC_BLOCK_SIZE Then lRead = ARC_BLOCK_SIZE
        If lRead <= 0 Then Exit Do
        Dim baChunk() As Byte
        ReDim baChunk(0 To lRead - 1)
        Get #iFile, , baChunk
        arc_write_data m_hArcWrite, VarPtr(baChunk(0)), lRead
    Loop
    Close #iFile: iFile = 0
    
    arc_write_finish_entry m_hArcWrite
    arc_entry_free hEntry
    
    ArcAddFile = True
QH:
    On Error Resume Next
    If iFile > 0 Then Close #iFile
    Exit Function
EH:
    m_sLastError = Err.Description
    ArcAddFile = False
    Resume QH
End Function

Public Function ArcAddFolder(ByVal FolderPath As String, Optional ByVal Recursive As Boolean = True, _
        Optional ByVal BaseFolder As String = "") As Boolean
    '--- Tambahkan semua file dari folder ke arsip.
    Const FUNC_NAME As String = "ArcAddFolder"
    Dim sFile As String, sRelative As String
    
    On Error GoTo EH
    If m_hArcWrite = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Arsip belum dibuka untuk menulis"
    
    If Right$(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    If LenB(BaseFolder) = 0 Then BaseFolder = FolderPath
    If Right$(BaseFolder, 1) <> "\" Then BaseFolder = BaseFolder & "\"
    
    '--- Scan files
    sFile = Dir$(FolderPath & "*.*")
    Do While LenB(sFile) <> 0
        sRelative = Mid$(FolderPath & sFile, Len(BaseFolder) + 1)
        ArcAddFile FolderPath & sFile, sRelative
        sFile = Dir$
    Loop
    
    '--- Scan subdirectories
    If Recursive Then
        Dim sDir As String
        sDir = Dir$(FolderPath & "*.*", vbDirectory)
        Do While LenB(sDir) <> 0
            If sDir <> "." And sDir <> ".." Then
                If (GetAttr(FolderPath & sDir) And vbDirectory) = vbDirectory Then
                    ArcAddFolder FolderPath & sDir & "\", True, BaseFolder
                End If
            End If
            sDir = Dir$
        Loop
    End If
    
    ArcAddFolder = True
QH:
    Exit Function
EH:
    m_sLastError = Err.Description
    ArcAddFolder = False
    Resume QH
End Function

Public Function ArcAddBytes(ByVal Name As String, ByRef Data() As Byte) As Boolean
    '--- Tambahkan data dari byte array ke arsip.
    Const FUNC_NAME As String = "ArcAddBytes"
    Dim hEntry As Long, lRet As Long
    Dim lSize As Long
    
    On Error GoTo EH
    If m_hArcWrite = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Arsip belum dibuka untuk menulis"
    
    lSize = UBound(Data) - LBound(Data) + 1
    
    hEntry = arc_entry_new()
    If hEntry = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Gagal membuat entry"
    
    Dim baName() As Byte
    baName = ToUtf8Array(Replace$(Name, "\", "/"))
    ReDim Preserve baName(0 To UBound(baName) + 1)
    baName(UBound(baName)) = 0
    arc_entry_set_pathname_utf8 hEntry, VarPtr(baName(0))
    arc_entry_set_size hEntry, lSize
    arc_entry_set_filetype hEntry, ARC_IFREG
    arc_entry_set_perm hEntry, ARC_PERM_FILE
    
    lRet = arc_write_header(m_hArcWrite, hEntry)
    If lRet < ARC_OK And lRet <> ARC_WARN Then
        arc_entry_free hEntry
        Err.Raise vbObjectError, FUNC_NAME, "Gagal menulis header"
    End If
    
    arc_write_data m_hArcWrite, VarPtr(Data(LBound(Data))), lSize
    arc_write_finish_entry m_hArcWrite
    arc_entry_free hEntry
    
    ArcAddBytes = True
QH:
    Exit Function
EH:
    m_sLastError = Err.Description
    ArcAddBytes = False
    Resume QH
End Function

Public Function ArcSave() As Boolean
    '--- Simpan dan tutup arsip yang sedang ditulis.
    Const FUNC_NAME As String = "ArcSave"
    On Error GoTo EH
    If m_hArcWrite = 0 Then Err.Raise vbObjectError, FUNC_NAME, "Tidak ada arsip yang sedang ditulis"
    arc_write_close m_hArcWrite
    arc_write_free m_hArcWrite
    m_hArcWrite = 0
    ArcSave = True
QH:
    Exit Function
EH:
    m_sLastError = Err.Description
    ArcSave = False
    Resume QH
End Function

'--- Helper: detect archive type dari ekstensi file
Private Function pvDetectArchiveType(ByVal sFile As String) As ArchiveType
    Dim sExt As String
    sExt = LCase$(Mid$(sFile, InStrRev(sFile, ".") + 1))
    Select Case sExt
        Case "zip":     pvDetectArchiveType = arZip
        Case "7z":      pvDetectArchiveType = arSevenZip
        Case "tar":     pvDetectArchiveType = arTar
        Case "tgz":     pvDetectArchiveType = arTarGz
        Case "tbz2", "tb2": pvDetectArchiveType = arTarBz2
        Case "txz":     pvDetectArchiveType = arTarXz
        Case "tzst":    pvDetectArchiveType = arTarZst
        Case "gz":      pvDetectArchiveType = arGz
        Case "rar":     pvDetectArchiveType = arRar
        Case "cab":     pvDetectArchiveType = arCab
        Case "iso":     pvDetectArchiveType = arIso
        Case "cpio":    pvDetectArchiveType = arCpio
        Case "xar":     pvDetectArchiveType = arXar
        Case Else:      pvDetectArchiveType = arZip
    End Select
End Function

'--- Helper: set format + filter di archive writer
Private Sub pvSetArcFormat(ByVal hWrite As Long, ByVal ArcType As ArchiveType)
    Select Case ArcType
        Case arZip
            arc_write_set_format_zip hWrite
            arc_write_add_filter_none hWrite
        Case arSevenZip
            arc_write_set_format_7zip hWrite
            arc_write_add_filter_none hWrite
        Case arTar
            arc_write_set_format_ustar hWrite
            arc_write_add_filter_none hWrite
        Case arTarGz
            arc_write_set_format_pax hWrite
            arc_write_add_filter_gzip hWrite
        Case arTarBz2
            arc_write_set_format_pax hWrite
            arc_write_add_filter_bzip2 hWrite
        Case arTarXz
            arc_write_set_format_pax hWrite
            arc_write_add_filter_xz hWrite
        Case arTarZst
            arc_write_set_format_pax hWrite
            arc_write_add_filter_zstd hWrite
        Case arCpio
            arc_write_set_format_cpio hWrite
            arc_write_add_filter_none hWrite
        Case arXar
            arc_write_set_format_xar hWrite
            arc_write_add_filter_none hWrite
        Case Else
            arc_write_set_format_zip hWrite
            arc_write_add_filter_none hWrite
    End Select
End Sub

'--- Helper: cek apakah archiveint.dll tersedia
Private Function pvArcAvailable() As Boolean
    If Not m_bArcChecked Then
        m_bArcChecked = True
        On Error Resume Next
        m_hArcInt = LoadLibrary(StrPtr("archiveint.dll"))
        On Error GoTo 0
        If m_hArcInt <> 0 Then
            On Error Resume Next
            Dim lPtr As Long
            lPtr = arc_version_string()
            If Err.Number <> 0 Or lPtr = 0 Then
                FreeLibrary m_hArcInt
                m_hArcInt = 0
            End If
            On Error GoTo 0
        End If
    End If
    pvArcAvailable = (m_hArcInt <> 0)
End Function

'--- Helper: error string dari libarchive
Private Function pvArcErrorString(ByVal hArc As Long) As String
    On Error Resume Next
    Dim lPtr As Long
    lPtr = arc_error_string(hArc)
    If lPtr <> 0 Then
        Dim lLen As Long
        lLen = lstrlenA(lPtr)
        If lLen > 0 And lLen < 1024 Then
            Dim baMsg() As Byte
            ReDim baMsg(0 To lLen - 1)
            CopyMemory baMsg(0), ByVal lPtr, lLen
            pvArcErrorString = StrConv(baMsg, vbUnicode)
        End If
    End If
    If LenB(pvArcErrorString) = 0 Then pvArcErrorString = "Unknown error"
End Function

'--- Helper: wide string (LPCWSTR) ke VB String
Private Function pvWideToString(ByVal lPtr As Long) As String
    Dim lLen As Long
    lLen = lstrlenW(lPtr)
    If lLen > 0 And lLen < 32768 Then
        pvWideToString = String$(lLen, 0)
        CopyMemory ByVal StrPtr(pvWideToString), ByVal lPtr, lLen * 2
    End If
End Function

'--- Helper: buat folder secara rekursif
Private Sub pvMkDirRecursive(ByVal sPath As String)
    If Right$(sPath, 1) = "\" Then sPath = Left$(sPath, Len(sPath) - 1)
    If Dir$(sPath, vbDirectory) <> "" Then Exit Sub
    Dim sParent As String
    sParent = Left$(sPath, InStrRev(sPath, "\") - 1)
    If LenB(sParent) > 2 Then pvMkDirRecursive sParent
    On Error Resume Next
    MkDir sPath
    On Error GoTo 0
End Sub

Private Function pvZlibAvailable() As Boolean
    '--- Cek dan load zlib1.dll secara dinamis (GetProcAddress, tanpa Declare)
    If Not m_bZlibChecked Then
        m_bZlibChecked = True
        '--- 1. Coba dari system/app path standar
        m_hZlib = LoadLibrary(StrPtr("zlib1.dll"))
        '--- 2. Coba dari App.Path
        If m_hZlib = 0 Then
            On Error Resume Next
            Dim sAppDir As String
            sAppDir = App.Path
            If LenB(sAppDir) <> 0 Then
                SetDllDirectoryW StrPtr(sAppDir)
                m_hZlib = LoadLibrary(StrPtr("zlib1.dll"))
                SetDllDirectoryW 0
            End If
            On Error GoTo 0
        End If
        '--- 3. Extract dari embedded Base64
        If m_hZlib = 0 Then
            Dim sTempPath As String
            sTempPath = pvExtractEmbeddedZlib()
            If LenB(sTempPath) <> 0 Then
                Dim sTempDir As String
                sTempDir = Environ$("TEMP")
                SetDllDirectoryW StrPtr(sTempDir)
                m_hZlib = LoadLibrary(StrPtr("zlib1.dll"))
                SetDllDirectoryW 0
            End If
        End If
        '--- 4. Resolve function pointers via GetProcAddress
        If m_hZlib <> 0 Then
            m_pfnDeflateInit2 = GetProcAddress(m_hZlib, "deflateInit2_")
            m_pfnDeflate = GetProcAddress(m_hZlib, "deflate")
            m_pfnDeflateEnd = GetProcAddress(m_hZlib, "deflateEnd")
            m_pfnInflateInit2 = GetProcAddress(m_hZlib, "inflateInit2_")
            m_pfnInflate = GetProcAddress(m_hZlib, "inflate")
            m_pfnInflateEnd = GetProcAddress(m_hZlib, "inflateEnd")
            m_pfnCrc32 = GetProcAddress(m_hZlib, "crc32")
            m_pfnZlibVersion = GetProcAddress(m_hZlib, "zlibVersion")
            m_pfnCompress2 = GetProcAddress(m_hZlib, "compress2")
            m_pfnUncompress = GetProcAddress(m_hZlib, "uncompress")
            m_pfnCompressBound = GetProcAddress(m_hZlib, "compressBound")
            '--- Verifikasi: semua fungsi harus ditemukan
            If m_pfnDeflateInit2 = 0 Or m_pfnDeflate = 0 Or m_pfnCrc32 = 0 Or m_pfnZlibVersion = 0 Then
                FreeLibrary m_hZlib
                m_hZlib = 0
            End If
        End If
    End If
    pvZlibAvailable = (m_hZlib <> 0)
End Function

'--- Zlib wrapper: panggil fungsi zlib via DispCallFunc (cdecl, runtime-safe)
'--- prgpvarg harus array of VarPtr(Variant), bukan VarPtr(rawValue)

Private Function zlib_zlibVersion() As Long
    Dim vRet As Variant
    DispCallFunc 0, m_pfnZlibVersion, CC_CDECL, vbLong, 0, ByVal 0&, ByVal 0&, vRet
    zlib_zlibVersion = CLng(vRet)
End Function

Private Function zlib_deflateInit2(strm As z_stream, ByVal Level As Long, ByVal Method As Long, ByVal windowBits As Long, ByVal memLevel As Long, ByVal strategy As Long, ByVal sVersion As String, ByVal stream_size As Long) As Long
    Dim vRet As Variant
    Dim baVer() As Byte
    baVer = StrConv(sVersion, vbFromUnicode)
    ReDim Preserve baVer(0 To UBound(baVer) + 1)
    baVer(UBound(baVer)) = 0
    Dim lVerPtr As Long: lVerPtr = VarPtr(baVer(0))
    Dim lStrmPtr As Long: lStrmPtr = VarPtr(strm)
    Dim vArgs(0 To 7) As Variant
    vArgs(0) = lStrmPtr: vArgs(1) = Level: vArgs(2) = Method: vArgs(3) = windowBits
    vArgs(4) = memLevel: vArgs(5) = strategy: vArgs(6) = lVerPtr: vArgs(7) = stream_size
    Dim vt(0 To 7) As Integer
    vt(0) = vbLong: vt(1) = vbLong: vt(2) = vbLong: vt(3) = vbLong
    vt(4) = vbLong: vt(5) = vbLong: vt(6) = vbLong: vt(7) = vbLong
    Dim pp(0 To 7) As Long
    pp(0) = VarPtr(vArgs(0)): pp(1) = VarPtr(vArgs(1)): pp(2) = VarPtr(vArgs(2)): pp(3) = VarPtr(vArgs(3))
    pp(4) = VarPtr(vArgs(4)): pp(5) = VarPtr(vArgs(5)): pp(6) = VarPtr(vArgs(6)): pp(7) = VarPtr(vArgs(7))
    DispCallFunc 0, m_pfnDeflateInit2, CC_CDECL, vbLong, 8, vt(0), pp(0), vRet
    zlib_deflateInit2 = CLng(vRet)
End Function

Private Function zlib_deflate(strm As z_stream, ByVal flush As Long) As Long
    Dim vRet As Variant
    Dim lStrmPtr As Long: lStrmPtr = VarPtr(strm)
    Dim vArgs(0 To 1) As Variant: vArgs(0) = lStrmPtr: vArgs(1) = flush
    Dim vt(0 To 1) As Integer: vt(0) = vbLong: vt(1) = vbLong
    Dim pp(0 To 1) As Long: pp(0) = VarPtr(vArgs(0)): pp(1) = VarPtr(vArgs(1))
    DispCallFunc 0, m_pfnDeflate, CC_CDECL, vbLong, 2, vt(0), pp(0), vRet
    zlib_deflate = CLng(vRet)
End Function

Private Function zlib_deflateEnd(strm As z_stream) As Long
    Dim vRet As Variant
    Dim lStrmPtr As Long: lStrmPtr = VarPtr(strm)
    Dim vArgs(0) As Variant: vArgs(0) = lStrmPtr
    Dim vt(0) As Integer: vt(0) = vbLong
    Dim pp(0) As Long: pp(0) = VarPtr(vArgs(0))
    DispCallFunc 0, m_pfnDeflateEnd, CC_CDECL, vbLong, 1, vt(0), pp(0), vRet
    zlib_deflateEnd = CLng(vRet)
End Function

Private Function zlib_inflateInit2(strm As z_stream, ByVal windowBits As Long, ByVal sVersion As String, ByVal stream_size As Long) As Long
    Dim vRet As Variant
    Dim baVer() As Byte
    baVer = StrConv(sVersion, vbFromUnicode)
    ReDim Preserve baVer(0 To UBound(baVer) + 1)
    baVer(UBound(baVer)) = 0
    Dim lVerPtr As Long: lVerPtr = VarPtr(baVer(0))
    Dim lStrmPtr As Long: lStrmPtr = VarPtr(strm)
    Dim vArgs(0 To 3) As Variant: vArgs(0) = lStrmPtr: vArgs(1) = windowBits: vArgs(2) = lVerPtr: vArgs(3) = stream_size
    Dim vt(0 To 3) As Integer: vt(0) = vbLong: vt(1) = vbLong: vt(2) = vbLong: vt(3) = vbLong
    Dim pp(0 To 3) As Long: pp(0) = VarPtr(vArgs(0)): pp(1) = VarPtr(vArgs(1)): pp(2) = VarPtr(vArgs(2)): pp(3) = VarPtr(vArgs(3))
    DispCallFunc 0, m_pfnInflateInit2, CC_CDECL, vbLong, 4, vt(0), pp(0), vRet
    zlib_inflateInit2 = CLng(vRet)
End Function

Private Function zlib_inflate(strm As z_stream, ByVal flush As Long) As Long
    Dim vRet As Variant
    Dim lStrmPtr As Long: lStrmPtr = VarPtr(strm)
    Dim vArgs(0 To 1) As Variant: vArgs(0) = lStrmPtr: vArgs(1) = flush
    Dim vt(0 To 1) As Integer: vt(0) = vbLong: vt(1) = vbLong
    Dim pp(0 To 1) As Long: pp(0) = VarPtr(vArgs(0)): pp(1) = VarPtr(vArgs(1))
    DispCallFunc 0, m_pfnInflate, CC_CDECL, vbLong, 2, vt(0), pp(0), vRet
    zlib_inflate = CLng(vRet)
End Function

Private Function zlib_inflateEnd(strm As z_stream) As Long
    Dim vRet As Variant
    Dim lStrmPtr As Long: lStrmPtr = VarPtr(strm)
    Dim vArgs(0) As Variant: vArgs(0) = lStrmPtr
    Dim vt(0) As Integer: vt(0) = vbLong
    Dim pp(0) As Long: pp(0) = VarPtr(vArgs(0))
    DispCallFunc 0, m_pfnInflateEnd, CC_CDECL, vbLong, 1, vt(0), pp(0), vRet
    zlib_inflateEnd = CLng(vRet)
End Function

Private Function zlib_crc32(ByVal crc As Long, ByVal buf As Long, ByVal nLen As Long) As Long
    Dim vRet As Variant
    Dim vArgs(0 To 2) As Variant: vArgs(0) = crc: vArgs(1) = buf: vArgs(2) = nLen
    Dim vt(0 To 2) As Integer: vt(0) = vbLong: vt(1) = vbLong: vt(2) = vbLong
    Dim pp(0 To 2) As Long: pp(0) = VarPtr(vArgs(0)): pp(1) = VarPtr(vArgs(1)): pp(2) = VarPtr(vArgs(2))
    DispCallFunc 0, m_pfnCrc32, CC_CDECL, vbLong, 3, vt(0), pp(0), vRet
    zlib_crc32 = CLng(vRet)
End Function

Private Function zlib_compress2(ByVal dest As Long, destLen As Long, ByVal Source As Long, ByVal sourceLen As Long, ByVal Level As Long) As Long
    Dim vRet As Variant
    Dim vArgs(0 To 4) As Variant: vArgs(0) = dest: vArgs(1) = VarPtr(destLen): vArgs(2) = Source: vArgs(3) = sourceLen: vArgs(4) = Level
    Dim vt(0 To 4) As Integer: vt(0) = vbLong: vt(1) = vbLong: vt(2) = vbLong: vt(3) = vbLong: vt(4) = vbLong
    Dim pp(0 To 4) As Long: pp(0) = VarPtr(vArgs(0)): pp(1) = VarPtr(vArgs(1)): pp(2) = VarPtr(vArgs(2)): pp(3) = VarPtr(vArgs(3)): pp(4) = VarPtr(vArgs(4))
    DispCallFunc 0, m_pfnCompress2, CC_CDECL, vbLong, 5, vt(0), pp(0), vRet
    zlib_compress2 = CLng(vRet)
End Function

Private Function zlib_uncompress(ByVal dest As Long, destLen As Long, ByVal Source As Long, ByVal sourceLen As Long) As Long
    Dim vRet As Variant
    Dim vArgs(0 To 3) As Variant: vArgs(0) = dest: vArgs(1) = VarPtr(destLen): vArgs(2) = Source: vArgs(3) = sourceLen
    Dim vt(0 To 3) As Integer: vt(0) = vbLong: vt(1) = vbLong: vt(2) = vbLong: vt(3) = vbLong
    Dim pp(0 To 3) As Long: pp(0) = VarPtr(vArgs(0)): pp(1) = VarPtr(vArgs(1)): pp(2) = VarPtr(vArgs(2)): pp(3) = VarPtr(vArgs(3))
    DispCallFunc 0, m_pfnUncompress, CC_CDECL, vbLong, 4, vt(0), pp(0), vRet
    zlib_uncompress = CLng(vRet)
End Function

Private Function zlib_compressBound(ByVal sourceLen As Long) As Long
    Dim vRet As Variant
    Dim vArgs(0) As Variant: vArgs(0) = sourceLen
    Dim vt(0) As Integer: vt(0) = vbLong
    Dim pp(0) As Long: pp(0) = VarPtr(vArgs(0))
    DispCallFunc 0, m_pfnCompressBound, CC_CDECL, vbLong, 1, vt(0), pp(0), vRet
    zlib_compressBound = CLng(vRet)
End Function

Private Function pvExtractEmbeddedZlib() As String
    '--- Extract embedded zlib1.dll ke %TEMP%\zlib1.dll
    '--- Jika file sudah ada dan ukurannya benar, skip extraction
    Const LNG_ZLIB_SIZE As Long = 75264
    Dim sDllPath As String
    Dim baData() As Byte
    Dim iFile As Integer
    
    On Error GoTo EH
    
    '--- Tentukan path
    sDllPath = Environ$("TEMP") & "\zlib1.dll"
    
    '--- Cek apakah sudah ada dan ukurannya benar
    If Dir$(sDllPath) <> "" Then
        If FileLen(sDllPath) = LNG_ZLIB_SIZE Then
            pvExtractEmbeddedZlib = sDllPath
            Exit Function
        End If
        On Error Resume Next
        Kill sDllPath
        On Error GoTo EH
    End If
    
    '--- Gabungkan Base64 dalam beberapa langkah
    Dim s1 As String, s2 As String, s3 As String, s4 As String
    s1 = STR_ZLIB_1 & STR_ZLIB_2 & STR_ZLIB_3 & STR_ZLIB_4 & STR_ZLIB_5
    s2 = STR_ZLIB_6 & STR_ZLIB_7 & STR_ZLIB_8 & STR_ZLIB_9 & STR_ZLIB_10
    s3 = STR_ZLIB_11 & STR_ZLIB_12 & STR_ZLIB_13 & STR_ZLIB_14 & STR_ZLIB_15
    s4 = STR_ZLIB_16 & STR_ZLIB_17 & STR_ZLIB_18 & STR_ZLIB_19 & STR_ZLIB_20
    
    Dim sBase64 As String
    sBase64 = s1 & s2 & s3 & s4
    s1 = "": s2 = "": s3 = "": s4 = ""
    
    '--- Decode Base64
    baData = FromBase64Array(sBase64)
    sBase64 = ""
    If UBound(baData) + 1 <> LNG_ZLIB_SIZE Then GoTo EH
    
    '--- Tulis ke file
    iFile = FreeFile
    Open sDllPath For Binary Access Write As #iFile
        Put #iFile, , baData
    Close #iFile
    
    pvExtractEmbeddedZlib = sDllPath
    Exit Function
EH:
    On Error Resume Next
    Close #iFile
    pvExtractEmbeddedZlib = ""
End Function

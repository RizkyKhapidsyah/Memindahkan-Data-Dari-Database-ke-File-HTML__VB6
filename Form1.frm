VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memindahkan Data dari Database ke File HTML"
   ClientHeight    =   1740
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   8895
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit   'Setiap variabel yang digunakan harus
                  'dideklarasikan!
                  
'Created by Rizky Khapidsyah
'Source Code Program dimulai dari sini

Private Sub Command1_Click()
Dim fnum As Integer
Dim db As Database
Dim rs As Recordset
Dim num_fields As Integer
Dim i As Integer
Dim num_processed As Integer
    On Error GoTo MiscError
    'Buka output file.
    fnum = FreeFile
    Open Text2.Text For Output As fnum
    'Tuliskan informasi header HTML
    Print #fnum, "<HTML>"
    Print #fnum, "<HEAD>"
    'Ganti "Ini Judul Atas" dengan judul file HTML yang
    'Anda inginkan tampil di bagian atas bar jendela
    Print #fnum, "<TITLE>Ini Judul Atas</TITLE>"
    Print #fnum, "</HEAD>"
    Print #fnum, ""
    Print #fnum, "<BODY TEXT=#000000 BGCOLOR=white>"
    'Ganti "Judul Tabel" dengan judul yang Anda
    'inginkan tampil di bagian atas halaman html
    'tersebut
    Print #fnum, "<H1>Judul Tabel</H1>"
    'Mulai buat tabel HTML
    Print #fnum, "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=2 BGCOLOR=#00C0FF BORDER=1>"
    'Buka database.
    Set db = OpenDatabase(Text1.Text)
    'Buka tabel di database.
    'Ganti "t_mhs" dengan nama tabel di database Anda
    'dan "NIM" dengan nama field yang Anda inginkan
    'disortir
    'Jika Anda tidak ingin menyortir tabel, Anda dapat
    'menghilangkan "ORDER BY NIM"
    Set rs = db.OpenRecordset("SELECT * FROM t_mhs ORDER BY NIM")
    'Gunakan nama field sebagai judul setiap
    'field/kolom
    Print #fnum, "    <TR>"     ' Mulai sebuah baris...
    num_fields = rs.Fields.Count
    For i = 0 To num_fields - 1
        Print #fnum, "        <TH>";
        Print #fnum, rs.Fields(i).Name;
        Print #fnum, "</TH>"
    Next i
    Print #fnum, "    </TR>"

    'Proses semua record...
    Do While Not rs.EOF
        num_processed = num_processed + 1
       'Mulai dengan sebuah baris baru untuk record ini
        Print #fnum, "    <TR>";
        For i = 0 To num_fields - 1
            Print #fnum, "        <TD>";
            Print #fnum, rs.Fields(i).Value;
            Print #fnum, "</TD>"
        Next i
        Print #fnum, "</TR>";
        rs.MoveNext   'Maju ke record berikutnya
    Loop
    'Akhir tabel di file HTML
    Print #fnum, "</TABLE>"
    Print #fnum, "<P>"
    Print #fnum, "<H3>" & _
        Format$(num_processed) & _
        " records ditampilkan...</H3>"
    Print #fnum, "</BODY>"
    Print #fnum, "</HTML>"
    rs.Close   'Tutup tabel di database
    db.Close   'Tutup database
    Close fnum 'Tutup file teks, lalu tampilkan pesan
    MsgBox "Berhasil memproses " _
           & Format$(num_processed) & " records.", _
           vbInformation, "Sukses"
    Exit Sub
MiscError: 'Jika terjadi error, tampilkan pesan nomor
           'dan deskripsi errornya...
    MsgBox "Error " & Err.Number & _
        vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
    'Tempatkan nama file database Anda di Text1
    'dan nama file HTML yang akan dibuat di Text2
    Text1.Text = App.Path + "\mahasiswa.mdb"
    Text2.Text = App.Path + "\DataHTML.html"
End Sub



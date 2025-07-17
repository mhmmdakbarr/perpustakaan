VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F_Data_Peminjaman 
   Caption         =   "F_Data_Peminjaman"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Tabel Data Peminjaman"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   15000
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_Data_Peminjaman.frx":0000
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "F_Data_Peminjaman.frx":0014
         TabIndex        =   27
         Top             =   240
         Width           =   14295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Proses"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   14415
      Begin VB.CommandButton Ckeluar 
         Caption         =   "Keluar"
         Height          =   495
         Left            =   12000
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Chapus 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   9720
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Cupdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   7560
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Cedit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Csimpan 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Ctambah 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   14415
      Begin VB.TextBox Tidpem 
         Height          =   495
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   2535
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Perpustakaan\Perpustakaan_A2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   8280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Data_peminjaman"
         Top             =   4560
         Width           =   2055
      End
      Begin VB.TextBox Tnamakar 
         Height          =   495
         Left            =   2520
         TabIndex        =   19
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox Tidkar 
         Height          =   525
         Left            =   2520
         TabIndex        =   18
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Ttangpeng 
         Height          =   525
         Left            =   2520
         TabIndex        =   17
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Ttangpem 
         Height          =   495
         Left            =   2520
         TabIndex        =   16
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Tnamab 
         Height          =   525
         Left            =   2520
         TabIndex        =   15
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Tkb 
         Height          =   525
         Left            =   2520
         TabIndex        =   14
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Tnamapem 
         Height          =   495
         Left            =   2520
         TabIndex        =   13
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Nama Karyawan"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "ID karyawan"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Tanggal Pengembalian"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Tanggal Peminjaman"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Nama Buku"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Kode Buku"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nama Peminjam"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "ID Peminjam"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Data Peminjaman"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "F_Data_Peminjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Perpustakaan As Database
Dim Data_peminjaman As Recordset

Private Sub Cedit_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan Id peminjam yang di cari", "Cari Data")
Data_peminjaman.Seek "=", pesan

If Data_peminjaman.NoMatch Then
X = MsgBox("Maaf Id Peminjam yang di cari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tidpem.Text = Data_peminjaman!Id_peminjam
Tidpem.Enabled = False
Tnamapem.Text = Data_peminjaman!Nama_peminjam
Tkb.Text = Data_peminjaman!kode_buku
Tnamab.Text = Data_peminjaman!Nama_buku
Ttangpem.Text = Data_peminjaman!Tanggal_peminjaman
Ttangpeng.Text = Data_peminjaman!Tanggal_pengembalian
Tidkar.Text = Data_peminjaman!Id_karyawan
Tnamakar.Text = Data_peminjaman!Nama_karyawan

End Sub

Private Sub Chapus_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan Id Peminjam yang dicari", "Cari Data")
Data_peminjaman.Seek "=", pesan

If Data_peminjaman.NoMatch Then
X = MsgBox("Maaf Id Peminjam yang dicari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tidpem.Text = Data_peminjaman!Id_peminjam
Tidpem.Enabled = False
Tnamapem.Text = Data_peminjaman!Nama_peminjam
Tkb.Text = Data_peminjaman!kode_buku
Tnamab.Text = Data_peminjaman!Nama_buku
Ttangpem.Text = Data_peminjaman!Tanggal_peminjaman
Ttangpeng.Text = Data_peminjaman!Tanggal_pengembalian
Tidkar.Text = Data_peminjaman!Id_karyawan
Tnamakar.Text = Data_peminjaman!Nama_karyawan

X = MsgBox("Yakin Data Akan dihapus ?", vbYesNo, "konfirmasi")
    If X = vbYes Then
    Data_peminjaman.Seek "=", Tidpem.Text
    Data_peminjaman.Delete
    kosong
    Data1.Refresh
    DBGrid1.Refresh
    End If

End Sub

Private Sub Ckeluar_Click()
End
End Sub

Private Sub Csimpan_Click()
Data_peminjaman.Seek "=", Tidpem.Text

If Data_peminjaman.NoMatch Then

Data_peminjaman.AddNew
Data_peminjaman!Id_peminjam = Tidpem.Text
Data_peminjaman!Nama_peminjam = Tnamapem.Text
Data_peminjaman!kode_buku = Tkb.Text
Data_peminjaman!Nama_buku = Tnamab.Text
Data_peminjaman!Tanggal_peminjaman = Ttangpem.Text
Data_peminjaman!Tanggal_pengembalian = Ttangpeng.Text
Data_peminjaman!Id_karyawan = Tidkar.Text
Data_peminjaman!Nama_karyawan = Tnamakar.Text
Data_peminjaman.Update

X = MsgBox("Data berhasil tersimpan", vbInformation, "pesan")
Data1.Refresh
DBGrid1.Refresh
Else
X = MsgBox("Maaf Id peminjaman ada yang sama", vbInformation, "pesan")
End If

End Sub

Private Sub Ctambah_Click()
Tidpem.SetFocus
kosong
End Sub
Private Sub kosong()
Tidpem.Text = ""
Tnamapem.Text = ""
Tkb.Text = ""
Tnamab.Text = ""
Ttangpem.Text = ""
Ttangpeng.Text = ""
Tidkar.Text = ""
Tnamakar.Text = ""

End Sub

Private Sub Cupdate_Click()
Data_peminjaman.Edit
Data_peminjaman!Nama_peminjam = Tnamapem.Text
Data_peminjaman!kode_buku = Tkb.Text
Data_peminjaman!Nama_buku = Tnamab.Text
Data_peminjaman!Tanggal_peminjaman = Ttangpem.Text
Data_peminjaman!Tanggal_pengembalian = Ttangpeng.Text
Data_peminjaman!Id_karyawan = Tidkar.Text
Data_peminjaman!Nama_karyawan = Tnamakar.Text
Data_peminjaman.Update

X = MsgBox("Data Berhasil diubah", vbInformation, "informasi")
Data1.Refresh
DBGrid1.Refresh
End Sub

Private Sub Form_Load()
Set Perpustakaan = OpenDatabase("D:\Perpustakaan\Perpustakaan_A2.mdb")
Set Data_peminjaman = Perpustakaan.OpenRecordset("Data_peminjaman")
Data_peminjaman.Index = "Kunci_peminjaman"
End Sub

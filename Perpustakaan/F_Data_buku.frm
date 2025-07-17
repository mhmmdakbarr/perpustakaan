VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F_Data_buku 
   BackColor       =   &H80000004&
   Caption         =   "Data_buku"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Tabel Data Buku "
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   22
      Top             =   6120
      Width           =   14415
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_Data_buku.frx":0000
         Height          =   3135
         Left            =   0
         OleObjectBlob   =   "F_Data_buku.frx":0014
         TabIndex        =   23
         Top             =   360
         Width           =   14295
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Proses "
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   14415
      Begin VB.CommandButton Ckeluar 
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Chapus 
         Caption         =   "Hapus"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cupdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Cedit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Csimpan 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Ctambah 
         Caption         =   "Tambah"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Input Data Buku "
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   14415
      Begin VB.Data Data1 
         BackColor       =   &H00000000&
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Perpustakaan\Perpustakaan_A2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   12720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Data_buku"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Tjum 
         Height          =   525
         Left            =   2040
         TabIndex        =   14
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox Tter 
         Height          =   525
         Left            =   2040
         TabIndex        =   13
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox Tpen 
         Height          =   525
         Left            =   2040
         TabIndex        =   12
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Tpeng 
         Height          =   525
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Tjb 
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Tkb 
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jumlah buku"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tahun Terbit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Penerbit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pengarang"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Judul Buku"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kode Buku"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Sans Serif Collection"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data Buku"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "F_Data_buku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Perpustakaan As Database
Dim Data_buku As Recordset


Private Sub Cedit_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan Kode Buku yang di cari", "Cari Data")
Data_buku.Seek "=", pesan

If Data_buku.NoMatch Then
X = MsgBox("Maaf Kode Buku yang di cari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tkb.Text = Data_buku!kode_buku
Tkb.Enabled = False
Tjb.Text = Data_buku!Judul_buku
Tpeng.Text = Data_buku!Pengarang
Tpen.Text = Data_buku!Penerbit
Tter.Text = Data_buku!Tahun_terbit
Tjum.Text = Data_buku!Jumlah_buku

End Sub

Private Sub Chapus_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan Kode Buku yang dicari", "Cari Data")
Data_buku.Seek "=", pesan

If Data_buku.NoMatch Then
X = MsgBox("Maaf Kode buku yang dicari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tkb.Text = Data_buku!kode_buku
Tkb.Enabled = False
Tjb.Text = Data_buku!Judul_buku
Tpeng.Text = Data_buku!Pengarang
Tpen.Text = Data_buku!Penerbit
Tter.Text = Data_buku!Tahun_terbit
Tjum.Text = Data_buku!Jumlah_buku

X = MsgBox("Yakin Data Akan dihapus ?", vbYesNo, "konfirmasi")
    If X = vbYes Then
    Data_buku.Seek "=", Tkb.Text
    Data_buku.Delete
    kosong
    Data1.Refresh
    DBGrid1.Refresh
    End If
End Sub

Private Sub Ckeluar_Click()
End
End Sub

Private Sub Csimpan_Click()
Data_buku.Seek "=", Tkb.Text

If Data_buku.NoMatch Then

Data_buku.AddNew
Data_buku!kode_buku = Tkb.Text
Data_buku!Judul_buku = Tjb.Text
Data_buku!Pengarang = Tpeng.Text
Data_buku!Penerbit = Tpen.Text
Data_buku!Tahun_terbit = Tter.Text
Data_buku!Jumlah_buku = Tjum.Text
Data_buku.Update

X = MsgBox("Data berhasil tersimpan", vbInformation, "pesan")
Data1.Refresh
DBGrid1.Refresh
Else
X = MsgBox("Maaf Id barang ada yang sama", vbInformation, "pesan")
End If
End Sub

Private Sub Ctambah_Click()
Tkb.SetFocus
kosong
End Sub
Private Sub kosong()
Tkb.Text = ""
Tjb.Text = ""
Tpeng.Text = ""
Tpen.Text = ""
Tter.Text = ""
Tjum.Text = ""

End Sub

Private Sub Cupdate_Click()
Data_buku.Edit
Data_buku!Judul_buku = Tjb.Text
Data_buku!Pengarang = Tpeng.Text
Data_buku!Penerbit = Tpen.Text
Data_buku!Tahun_terbit = Tter.Text
Data_buku!Jumlah_buku = Tjum.Text
Data_buku.Update

X = MsgBox("Data Berhasil diubah", vbInformation, "informasi")
Data1.Refresh
DBGrid1.Refresh

End Sub

Private Sub Form_Load()
Set Perpustakaan = OpenDatabase("D:\Perpustakaan\Perpustakaan_A2.mdb")
Set Data_buku = Perpustakaan.OpenRecordset("Data_buku")
Data_buku.Index = "Kunci_buku"
End Sub

Private Sub Text2_Change()

End Sub


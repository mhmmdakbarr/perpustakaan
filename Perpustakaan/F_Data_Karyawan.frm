VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F_Data_Karyawan 
   Caption         =   "F_Data_Karyawan"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Tabel Data Karyawan"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   14295
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_Data_Karyawan.frx":0000
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "F_Data_Karyawan.frx":0014
         TabIndex        =   23
         Top             =   360
         Width           =   14055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   14295
      Begin VB.CommandButton Csimpan 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   3120
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Ctambah 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   1080
         TabIndex        =   21
         Top             =   360
         Width           =   1515
      End
      Begin VB.CommandButton Ckeluar 
         Caption         =   "Keluar"
         Height          =   495
         Left            =   11880
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Chapus 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   9720
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Cupdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   7440
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Cedit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   5280
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   14295
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Perpustakaan\Perpustakaan_A2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Data_karyawan"
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox Talamat 
         Height          =   765
         Left            =   2040
         TabIndex        =   16
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Tno 
         Height          =   525
         Left            =   2040
         TabIndex        =   15
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Temail 
         Height          =   495
         Left            =   2040
         TabIndex        =   14
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox Cjenis 
         Height          =   315
         ItemData        =   "F_Data_Karyawan.frx":10B7
         Left            =   2040
         List            =   "F_Data_Karyawan.frx":10C1
         TabIndex        =   13
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Tnama 
         Height          =   525
         Left            =   2040
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Tid 
         Height          =   525
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "No Telpon"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Email"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Jenis Kelamin"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nama Karyawan"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "ID Karyawan"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.Label Label1 
         Caption         =   "Data Karyawan"
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
         Left            =   6240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "F_Data_Karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Perpustakaan As Database
Dim Data_karyawan As Recordset

Private Sub Cedit_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan Id karyawan yang di cari", "Cari Data")
Data_karyawan.Seek "=", pesan

If Data_karyawan.NoMatch Then
X = MsgBox("Maaf Id karyawan yang di cari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tid.Text = Data_karyawan!Id_karyawan
Tid.Enabled = False
Tnama.Text = Data_karyawan!Nama_karyawan
Cjenis.Text = Data_karyawan!Jenis_kelamin
Temail.Text = Data_karyawan!Email
Tno.Text = Data_karyawan!No_telpon
Talamat.Text = Data_karyawan!Alamat

End Sub

Private Sub Chapus_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan Id karyawan yang dicari", "Cari Data")
Data_karyawan.Seek "=", pesan

If Data_karyawan.NoMatch Then
X = MsgBox("Maaf Id karyawan yang dicari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tid.Text = Data_karyawan!Id_karyawan
Tid.Enabled = False
Tnama.Text = Data_karyawan!Nama_karyawan
Cjenis.Text = Data_karyawan!Jenis_kelamin
Temail.Text = Data_karyawan!Email
Tno.Text = Data_karyawan!No_telpon
Talamat.Text = Data_karyawan!Alamat

X = MsgBox("Yakin Data Akan dihapus ?", vbYesNo, "konfirmasi")
    If X = vbYes Then
    Data_karyawan.Seek "=", Tid.Text
    Data_karyawan.Delete
    kosong
    Data1.Refresh
    DBGrid1.Refresh
    End If

End Sub

Private Sub Ckeluar_Click()
End
End Sub

Private Sub Csimpan_Click()
Data_karyawan.Seek "=", Tid.Text

If Data_karyawan.NoMatch Then

Data_karyawan.AddNew
Data_karyawan!Id_karyawan = Tid.Text
Data_karyawan!Nama_karyawan = Tnama.Text
Data_karyawan!Jenis_kelamin = Cjenis.Text
Data_karyawan!Email = Temail.Text
Data_karyawan!No_telpon = Tno.Text
Data_karyawan!Alamat = Talamat.Text
Data_karyawan.Update

X = MsgBox("Data berhasil tersimpan", vbInformation, "pesan")
Data1.Refresh
DBGrid1.Refresh
Else
X = MsgBox("Maaf Id karyawan ada yang sama", vbInformation, "pesan")
End If
End Sub

Private Sub Ctambah_Click()
Tid.SetFocus
kosong
End Sub
Private Sub kosong()
Tid.Text = ""
Tnama.Text = ""
Cjenis.Text = ""
Temail.Text = ""
Tno.Text = ""
Talamat.Text = ""
End Sub

Private Sub Cupdate_Click()
Data_karyawan.Edit
Data_karyawan!Nama_karyawan = Tnama.Text
Data_karyawan!Jenis_kelamin = Cjenis.Text
Data_karyawan!Email = Temail.Text
Data_karyawan!No_telpon = Tno.Text
Data_karyawan!Alamat = Talamat.Text
Data_karyawan.Update

X = MsgBox("Data Berhasil diubah", vbInformation, "informasi")
Data1.Refresh
DBGrid1.Refresh

End Sub

Private Sub Form_Load()
Set Perpustakaan = OpenDatabase("D:\Perpustakaan\Perpustakaan_A2.mdb")
Set Data_karyawan = Perpustakaan.OpenRecordset("Data_karyawan")
Data_karyawan.Index = "Kunci_karyawan"
End Sub

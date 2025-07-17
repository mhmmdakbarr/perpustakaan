VERSION 5.00
Begin VB.Form F_Utama 
   Caption         =   "F_Utama"
   ClientHeight    =   9915
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10215
      Left            =   0
      Picture         =   "F_Utama.frx":0000
      ScaleHeight     =   10155
      ScaleWidth      =   18795
      TabIndex        =   0
      Top             =   0
      Width           =   18855
   End
   Begin VB.Menu MU 
      Caption         =   "Menu Utama"
      Index           =   1
      Begin VB.Menu DB 
         Caption         =   "Data Buku"
         Index           =   2
      End
      Begin VB.Menu DA 
         Caption         =   "Data Anggota"
         Index           =   3
      End
      Begin VB.Menu DK 
         Caption         =   "Data Karyawan"
         Index           =   4
      End
   End
   Begin VB.Menu P 
      Caption         =   "Pengelolaan"
      Index           =   5
      Begin VB.Menu DPem 
         Caption         =   "Data Peminjaman"
         Index           =   6
      End
      Begin VB.Menu DPeng 
         Caption         =   "Data Pengembalian"
         Index           =   7
      End
   End
   Begin VB.Menu L 
      Caption         =   "Laporan"
      Index           =   8
      Begin VB.Menu LDB 
         Caption         =   "Laporan Data Buku"
         Index           =   9
      End
      Begin VB.Menu LDA 
         Caption         =   "Laporan Data Anggota"
         Index           =   10
      End
      Begin VB.Menu LDK 
         Caption         =   "Laporan Data Karyawan"
         Index           =   11
      End
      Begin VB.Menu LDPem 
         Caption         =   "Laporan Data Peminjaman"
         Index           =   12
      End
      Begin VB.Menu LDpeng 
         Caption         =   "Laporan Data Pengembalian"
      End
   End
   Begin VB.Menu TS 
      Caption         =   "Tentang Saya"
   End
   Begin VB.Menu E 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "F_Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DA_Click(Index As Integer)
F_Data_anggota.Show
End Sub

Private Sub DB_Click(Index As Integer)
F_Data_buku.Show
End Sub

Private Sub DK_Click(Index As Integer)
F_Data_Karyawan.Show
End Sub

Private Sub DPem_Click(Index As Integer)
F_Data_Peminjaman.Show
End Sub

Private Sub DPeng_Click(Index As Integer)
F_Data_Pengembalian.Show
End Sub

Private Sub E_Click()
End
End Sub

Private Sub LDA_Click(Index As Integer)
F_Laporan_Data_Anggota.Show
End Sub

Private Sub LDB_Click(Index As Integer)
F_Laporan_Data_Buku.Show
End Sub

Private Sub LDK_Click(Index As Integer)
F_Laporan_Data_Karyawan.Show
End Sub

Private Sub LDPem_Click(Index As Integer)
F_Laporan_Data_Peminjaman.Show
End Sub


Private Sub LDpeng_Click()
F_Laporan_Data_Pengembalian.Show
End Sub

Private Sub TS_Click()
F_Tentang_saya.Show
End Sub

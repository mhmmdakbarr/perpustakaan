VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form F_Laporan_Data_Peminjaman 
   Caption         =   "F_Laporan_Data_Peminjaman"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   14490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   14415
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_Laporan_Data_Peminjaman.frx":0000
         Height          =   2775
         Left            =   120
         OleObjectBlob   =   "F_Laporan_Data_Peminjaman.frx":0014
         TabIndex        =   4
         Top             =   3360
         Width           =   14175
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Perpustakaan\Perpustakaan_A2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   615
         Left            =   11280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Data_peminjaman"
         Top             =   2640
         Width           =   2655
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9600
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Laporan Data Peminjaman"
         Height          =   615
         Left            =   1440
         TabIndex        =   3
         Top             =   1560
         Width           =   3015
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
         Caption         =   "Laporan Data Peminjaman"
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
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "F_Laporan_Data_Peminjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Perpustakaan As Database
Dim Data_peminjaman As Recordset

Private Sub Command1_Click()
CrystalReport1.ReportFileName = App.Path & "\Data Peminjaman.rpt."
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 0
End Sub


Private Sub Form_Load()
Set Perpustakaan = OpenDatabase("D:\Perpustakaan\Perpustakaan_A2.mdb")
Set Data_peminjaman = Perpustakaan.OpenRecordset("Data_peminjaman")
Data_peminjaman.Index = "Kunci_peminjaman"
End Sub

VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form F_Laporan_Data_Pengembalian 
   Caption         =   "F_Laporan_Data_Pengembalian"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   14295
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_Laporan_Data_Pengembalian.frx":0000
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "F_Laporan_Data_Pengembalian.frx":0014
         TabIndex        =   3
         Top             =   3480
         Width           =   13935
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9720
         Top             =   2880
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Perpustakaan\Perpustakaan_A2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   10680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Data_pengembalian"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Laporan Data Pengembalian"
         Height          =   615
         Left            =   1200
         TabIndex        =   2
         Top             =   1440
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "F_Laporan_Data_Pengembalian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Perpustakaan As Database
Dim Data_pengembalian As Recordset

Private Sub Command1_Click()
CrystalReport1.ReportFileName = App.Path & "\Data Pengembalian.rpt."
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 0
End Sub

Private Sub Form_Load()
Set Perpustakaan = OpenDatabase("D:\Perpustakaan\Perpustakaan_A2.mdb")
Set Data_pengembalian = Perpustakaan.OpenRecordset("Data_pengembalian")
Data_pengembalian.Index = "Kunci_pengembalian"
End Sub

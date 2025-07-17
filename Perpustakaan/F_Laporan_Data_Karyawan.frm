VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form F_Laporan_Data_Karyawan 
   Caption         =   "F_Laporan_Data_Karyawan"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   14295
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9360
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
         Left            =   10440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Data_karyawan"
         Top             =   2760
         Width           =   2775
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_Laporan_Data_Karyawan.frx":0000
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "F_Laporan_Data_Karyawan.frx":0014
         TabIndex        =   4
         Top             =   3480
         Width           =   13815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Laporan Data Karyawan"
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Laporan Data Karyawan"
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
         Width           =   3495
      End
   End
End
Attribute VB_Name = "F_Laporan_Data_Karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Perpustakaan As Database
Dim Data_karyawan As Recordset
Private Sub Command1_Click()
CrystalReport1.ReportFileName = App.Path & "\Data Karyawan.rpt."
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 0
End Sub
Private Sub Form_Load()
Set Perpustakaan = OpenDatabase("D:\Perpustakaan\Perpustakaan_A2.mdb")
Set Data_karyawan = Perpustakaan.OpenRecordset("Data_karyawan")
Data_karyawan.Index = "Kunci_karyawan"
End Sub

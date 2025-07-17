VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form F_Laporan_Data_Anggota 
   Caption         =   "F_Laporan_Data_Anggota"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      ForeColor       =   &H80000013&
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   14295
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   10080
         Top             =   2640
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
         Left            =   11040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Data_anggota"
         Top             =   2640
         Width           =   2655
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_Laporan_Data_Anggota.frx":0000
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "F_Laporan_Data_Anggota.frx":0014
         TabIndex        =   4
         Top             =   3240
         Width           =   13935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Laporan Data Anggota"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   3
         Top             =   1800
         Width           =   2775
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
         Caption         =   "Laporan Data Anggota"
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
         Left            =   5760
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "F_Laporan_Data_Anggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Perpustakaan As Database
Dim Data_anggota As Recordset
Private Sub Command1_Click()
CrystalReport1.ReportFileName = App.Path & "\Data Anggota.rpt."
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 0
End Sub
Private Sub Form_Load()
Set Perpustakaan = OpenDatabase("D:\Perpustakaan\Perpustakaan_A2.mdb")
Set Data_anggota = Perpustakaan.OpenRecordset("Data_anggota")
Data_anggota.Index = "Kunci_anggota"
End Sub


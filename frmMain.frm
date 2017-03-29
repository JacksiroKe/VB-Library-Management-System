VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Shirikisho Library Management Information System"
   ClientHeight    =   8070
   ClientLeft      =   -4275
   ClientTop       =   -1080
   ClientWidth     =   14130
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   14130
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9340
      _Version        =   393216
      Tab             =   2
      TabHeight       =   1411
      BackColor       =   0
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Members"
      TabPicture(0)   =   "frmMain.frx":A8C53
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Departments"
      TabPicture(1)   =   "frmMain.frx":A8C6F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Books"
      TabPicture(2)   =   "frmMain.frx":A8C8B
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   11895
         Begin VB.Frame fraHosteList 
            Caption         =   "Books List: "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   5500
            TabIndex        =   12
            Top             =   240
            Width           =   6240
            Begin VB.ListBox lstBooks 
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3300
               Left            =   135
               TabIndex        =   13
               Top             =   480
               Width           =   5985
            End
         End
         Begin VB.Label lblReturnBook 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Return a Book ->"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   360
            TabIndex        =   19
            Top             =   2160
            Width           =   4935
         End
         Begin VB.Label lblNewBook 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Register a Book ->"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   360
            TabIndex        =   9
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label lblBookBorrow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Borrow a Book ->"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   360
            TabIndex        =   8
            Top             =   1320
            Width           =   4935
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   -74760
         TabIndex        =   5
         Top             =   840
         Width           =   11895
         Begin VB.Frame fraDeptList 
            Caption         =   " Department List: "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   5500
            TabIndex        =   10
            Top             =   240
            Width           =   6240
            Begin VB.ListBox lstDepartment 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3270
               Left            =   255
               TabIndex        =   11
               Top             =   480
               Width           =   5865
            End
         End
         Begin VB.Label lblNewDepartment 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Add a Department ->"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   4455
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   11895
         Begin VB.Frame Frame2 
            Caption         =   "Members List:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   5500
            TabIndex        =   14
            Top             =   240
            Width           =   6240
            Begin VB.ListBox lstMembers 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3270
               Left            =   120
               TabIndex        =   15
               Top             =   480
               Width           =   5985
            End
         End
         Begin VB.Label lblAllMembers 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "All Members ->"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            TabIndex        =   18
            Top             =   2160
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.Label lblProfile 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Member Profile ->"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   5055
         End
         Begin VB.Label lblRegister 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Register a Member ->"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   5055
         End
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   3360
      Top             =   360
      Width           =   9015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shirikisho Library"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3480
      TabIndex        =   17
      Top             =   480
      Width           =   8775
   End
   Begin VB.Label lblLogout 
      Alignment       =   2  'Center
      Caption         =   "Logout?"
      Height          =   615
      Left            =   7200
      TabIndex        =   2
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label lblLoggedIn 
      Alignment       =   2  'Center
      Caption         =   "Logged In User"
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   7320
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim loggedinas As String
Dim loggedfirst As String
Dim loggedseco As String

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
    
    LoggedInMe
    Load_AllMembers
    Load_AllDepartments
    Load_AllBooks
End Sub

Public Sub Load_AllMembers()
lstMembers.Clear
Dim str As String
On Error GoTo errorhandler
 Set rs = New ADODB.Recordset
    rs.Open "Select * from members ORDER BY memberid ASC", con, adOpenKeyset, adLockOptimistic
    Do Until rs.EOF
        lstMembers.AddItem rs!m_fullname
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
errorhandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Public Sub Load_AllDepartments()
lstDepartment.Clear
Dim str As String
On Error GoTo errorhandler
 Set rs = New ADODB.Recordset
    rs.Open "Select * from departments ORDER BY departmentid ASC", con, adOpenKeyset, adLockOptimistic
    Do Until rs.EOF
        lstDepartment.AddItem rs!d_name
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
errorhandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub


Public Sub Load_AllBooks()
lstBooks.Clear
Dim str As String
On Error GoTo errorhandler
 Set rs = New ADODB.Recordset
    rs.Open "Select DISTINCT b_name from books", con, adOpenKeyset, adLockOptimistic
    Do Until rs.EOF
        lstBooks.AddItem rs!b_name
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
errorhandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub


Private Sub LoggedInMe()
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * from admnistrators WHERE a_username = '" & lblLoggedIn.Caption & "'", con, adOpenKeyset, adLockOptimistic
    loggedfirst = rs!a_firstname
    loggedseco = rs!a_seconame
    lblLoggedIn.Caption = "You are logged in as: " & loggedfirst & " " & loggedseco
    Exit Sub
errorhandler:     MsgBox Err.Description & " No. " & Err.Number
End Sub


Private Sub lblDepartment_Click()
    drpDepartment.Show
End Sub

Private Sub lblAllMembers_Click()
    frmMembers.Show
End Sub

Private Sub lblBookBorrow_Click()
    frmBookBorrow.Show
End Sub

Private Sub lblBorrowNow_Click()
    frmBookBorrow.Show
End Sub

Private Sub lblLogout_Click()
    frmAdmLogin.Show
    Unload Me
End Sub

Private Sub lblNewBook_Click()
    frmBookRegister.Show
End Sub

Private Sub lblNewDepartment_Click()
    frmDepartNew.Show
End Sub

Private Sub lblProfile_Click()
    frmMemProfile.Show
End Sub

Private Sub lblRegister_Click()
    frmMemRegister.Show
End Sub

Private Sub lblReturnBook_Click()
    frmBookReturn.Show
End Sub

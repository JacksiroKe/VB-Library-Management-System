VERSION 5.00
Begin VB.Form frmMemProfile 
   Caption         =   "Member Profile - Shirikisho Library"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5640
      TabIndex        =   11
      Top             =   120
      Width           =   5295
   End
   Begin VB.Frame fraStudent 
      Caption         =   "Member Profile Information:"
      Height          =   6615
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   14295
      Begin VB.CommandButton cmdDeallocation 
         Caption         =   "Clear from Room"
         Height          =   615
         Left            =   8640
         TabIndex        =   1
         Top             =   5280
         Width           =   3855
      End
      Begin VB.Label lblIdnumber 
         Height          =   495
         Left            =   2880
         TabIndex        =   25
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Label Label12 
         Caption         =   "ID. No:"
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblMobile 
         Height          =   735
         Left            =   2880
         TabIndex        =   23
         Top             =   4680
         Width           =   4455
      End
      Begin VB.Label lblEmail 
         Height          =   495
         Left            =   2880
         TabIndex        =   22
         Top             =   4080
         Width           =   4575
      End
      Begin VB.Label lblGender 
         Height          =   495
         Left            =   2880
         TabIndex        =   21
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Label Label11 
         Caption         =   "Mobile:"
         Height          =   735
         Left            =   360
         TabIndex        =   20
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Email:"
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Gender:"
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblRoom 
         Height          =   615
         Left            =   10800
         TabIndex        =   17
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Room:"
         Height          =   615
         Left            =   8400
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblStatus 
         Height          =   495
         Left            =   10800
         TabIndex        =   15
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Status:"
         Height          =   495
         Left            =   8400
         TabIndex        =   14
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Line Line2 
         X1              =   7800
         X2              =   7800
         Y1              =   360
         Y2              =   6360
      End
      Begin VB.Label lblStudent 
         Height          =   615
         Left            =   2760
         TabIndex        =   13
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Student Name:"
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Department:"
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Course:"
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Expected Fees:"
         Height          =   615
         Left            =   8400
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Fee Arears: "
         Height          =   495
         Left            =   8400
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   14040
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lblDepartment 
         Height          =   615
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label lblCourse 
         Height          =   735
         Left            =   2760
         TabIndex        =   4
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label lblExpected 
         Height          =   495
         Left            =   10800
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblFeeAreas 
         Height          =   495
         Left            =   10800
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Adm. No:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmMemProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


Private Sub cmdDeallocation_Click()
    hostelDeallocate
    studentDeallocate
    MsgBox lblStudent.Caption & " has been deallocated from Room: " & lblRoom.Caption, vbInformation, App.Title
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
 End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    If KeyAscii = vbKeyReturn Then
        Rs.Open "Select * from students WHERE s_admno ='" & txtSearch.Text & "'", con, adOpenKeyset, adLockOptimistic
        lblStudent.Caption = Rs!s_fullname
        lblIdnumber.Caption = Rs!s_idno
        lblGender.Caption = Rs!s_gender
        lblEmail.Caption = Rs!s_email
        lblMobile.Caption = Rs!s_mobile
        lblDepartment.Caption = Rs!s_department
        lblCourse.Caption = Rs!s_course
        lblRoom.Caption = Rs!s_room
        lblStatus.Caption = Rs!s_status
        lblExpected.Caption = Rs!s_expected
        lblFeeAreas.Caption = Rs!s_fees
        Rs.Update
        Rs.Close
    End If
    Exit Sub
ErrorHandler:
 MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub hostelDeallocate()
On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from hostels WHERE h_room = '" & lblRoom.Caption & "'", con, adOpenKeyset, adLockOptimistic
    Rs!h_students = ""
    Rs.Update
    Rs.Close
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub studentDeallocate()
On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from students WHERE s_fullname='" & lblStudent.Caption & "'", con, adOpenKeyset, adLockOptimistic
    Rs!s_room = ""
    Rs!s_status = "Commuter"
    Rs.Update
    Rs.Close
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub


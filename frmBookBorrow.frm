VERSION 5.00
Begin VB.Form frmBookBorrow 
   Caption         =   "Book Borrowing - Shirikisho Library"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   9405
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
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
   Begin VB.Frame m_idnumber 
      Caption         =   "Book Borrowing:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   8535
      Begin VB.TextBox txtDateReturning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   5160
         Width           =   4695
      End
      Begin VB.CheckBox chkCanBorrow 
         Caption         =   "Can Borrow a Book"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CheckBox chkAvailable 
         Caption         =   "Book is available"
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   5760
         Width           =   3015
      End
      Begin VB.TextBox txtBookSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   840
         TabIndex        =   7
         Top             =   1680
         Width           =   6615
      End
      Begin VB.CommandButton cmdBorrowBook 
         Appearance      =   0  'Flat
         Caption         =   "Borrow Book"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   1
         Top             =   6360
         Width           =   3855
      End
      Begin VB.Label lblDept 
         Caption         =   "."
         Height          =   255
         Left            =   7800
         TabIndex        =   22
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Returning Date:"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Borrowing Date:"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label lblDateBorrow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   17
         Top             =   4680
         Width           =   4695
      End
      Begin VB.Label lblDepartment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   3240
         Width           =   4695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Department:"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label lblLendingRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   4200
         Width           =   4695
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Lending Rate:"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblBookNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Number:"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblBookName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Name:"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblBkResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Search Book by name then press ENTER"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Member Name:"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   8280
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblMember 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Member Number"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type A Member's Number and Hit Enter to Search."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   7335
   End
End
Attribute VB_Name = "frmBookBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim member_id As Integer, book_id As Integer, ddate As Date

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
    
    ddate = DateValue(Now)
    lblDateBorrow.Caption = ddate
    cmdBorrowBook.Enabled = False
End Sub

Private Sub chkCanBorrow_Validate(Cancel As Boolean)
    If chkCanBorrow.Value = 1 Then
        txtBookSearch.Enabled = True
    Else
        txtBookSearch.Enabled = False
    End If
End Sub

Private Sub BookEdit()
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from books WHERE b_name='" & lblBookName.Caption & "'", con, adOpenKeyset, adLockOptimistic
    rs!b_available = "False"
    rs.Update
    rs.Close
errorhandler:
End Sub

Private Sub MemberEdit()
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from members WHERE m_fullname='" & lblMember.Caption & "'", con, adOpenKeyset, adLockOptimistic
    rs!m_freetoborrow = "False"
    rs.Update
    rs.Close
errorhandler:
End Sub

Private Sub issueBook()
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from borrowing", con, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!b_memberid = member_id
    rs!b_bookid = book_id
    rs!b_bordate = lblDateBorrow.Caption
    rs!b_retdate = txtDateReturning.Text
    rs!b_amount = lblLendingRate.Caption
    rs.Update
    rs.Close
    Exit Sub
errorhandler:
End Sub

Private Sub cmdBorrowBook_Click()
    If txtDateReturning.Text = "" Then
        MsgBox "Returning Date not Set!", vbCritical, "Shirikisho Library"
        txtDateReturning.SetFocus
    Else
        issueBook
        BookEdit
        MemberEdit
        MsgBox "Well Done! Borrowing a book wa successful.", vbInformation, App.Title
        Unload Me
    End If
End Sub

Private Sub lblDept_Change()

On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from departments WHERE departmentid=" & CInt(lblDept.Caption) & "", con, adOpenKeyset, adLockOptimistic
    lblDepartment.Caption = rs!d_name
    lblLendingRate.Caption = rs!d_lendingrate
    rs.Close
errorhandler:
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    If KeyAscii = vbKeyReturn Then
        rs.Open "Select * from members WHERE m_fullname LIKE '%" & txtSearch.Text & "%' ", con, adOpenKeyset, adLockOptimistic
        lblMember.Caption = rs!m_fullname
        member_id = rs!memberid
        If rs!m_freetoborrow = "True" Then
           chkCanBorrow.Value = 1
           txtBookSearch.Enabled = True
        Else
           chkCanBorrow.Value = 0
            txtBookSearch.Enabled = False
        End If
        rs.Close
        lblResult.BackColor = &H80FF80
        lblResult.Caption = "Type A Member's Number and Hit Enter to Search."
        
    End If
    Exit Sub
errorhandler:
    lblResult.BackColor = &HFF&
    lblResult.Caption = "Member not found, Please try to search again!"
End Sub

Private Sub txtBookSearch_KeyPress(KeyAscii As Integer)
On Error GoTo errorhandler
    If KeyAscii = vbKeyReturn Then
        SearchBook
    End If
    Exit Sub
errorhandler:
    lblBkResult.BackColor = &H80FFFF
    lblBkResult.Caption = "Book not found, please try to search again!"
    cmdBorrowBook.Enabled = False
    SearchBook
        
End Sub

Private Sub SearchBook()
    Set rs = New ADODB.Recordset
    rs.Open "Select * from books WHERE b_name LIKE '%" & txtBookSearch.Text & "%' ", con, adOpenKeyset, adLockOptimistic
    lblBookName.Caption = rs!b_name
    lblDept.Caption = rs!b_department
    lblBookNo.Caption = CStr(rs!b_number)
    book_id = rs!bookid
    If rs!b_available = "True" Then
       chkAvailable.Value = 1
    End If
    rs.Close
    
    If Not (lblBookName.Caption = "") Then
        lblBkResult.BackColor = &H80FFFF
        lblBkResult.Caption = "Search Book by name then press ENTER"
        cmdBorrowBook.Enabled = True
    End If
End Sub

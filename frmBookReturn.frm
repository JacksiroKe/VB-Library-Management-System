VERSION 5.00
Begin VB.Form frmBookReturn 
   Caption         =   "Book Borrowing - Shirikisho Library"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
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
      Height          =   6735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.TextBox txtBookSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   2
         Top             =   600
         Width           =   6615
      End
      Begin VB.CommandButton cmdClearBook 
         Appearance      =   0  'Flat
         Caption         =   "Clear Book"
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
         Top             =   5760
         Width           =   3855
      End
      Begin VB.Label lblDateRertuning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label lblBorrower 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   3480
         Width           =   4695
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   8160
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Borrowed by:"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label lblDept 
         Caption         =   "."
         Height          =   255
         Left            =   7800
         TabIndex        =   15
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Returned Date:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Borrowed Date:"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label lblDateBorrow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   4440
         Width           =   4695
      End
      Begin VB.Label lblDepartment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Department:"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblLendingRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount Paid:"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblBookNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Number:"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblBookName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Name:"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
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
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   8280
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   8520
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmBookReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim member_id As Integer, book_id As Integer, ddate As Date

Private Sub cmdClearBook_Click()
    BookEdit
    MemberEdit
    BookReturned
    MsgBox "Book returned successfully", vbInformation, "Shirikisho Library"
    Unload Me
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
    
    ddate = DateValue(Now)
End Sub

Private Sub BookReturned()
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from borrowing WHERE b_bookid='" & book_id & "'", con, adOpenKeyset, adLockOptimistic
    rs!b_returned = "True"
    rs.Update
    rs.Close
errorhandler:
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
    rs.Open "Select * from members WHERE m_fullname='" & lblBorrower.Caption & "'", con, adOpenKeyset, adLockOptimistic
    rs!m_freetoborrow = "True"
    rs.Update
    rs.Close
errorhandler:
End Sub

Private Sub issueBook()
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from borrowing", con, adOpenKeyset, adLockOptimistic
    rs!b_memberid = member_id
    rs!b_bookid = book_id
    rs!b_bordate = lblDateBorrow.Caption
    rs!b_retdate = lblDateRertuning.Caption
    rs!b_amount = lblLendingRate.Caption
    rs.Update
    rs.Close
    Exit Sub
errorhandler:
End Sub

Public Function GetBorrowingInfo()
On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from borrowing WHERE b_bookid=" & book_id & "", con, adOpenKeyset, adLockOptimistic
    member_id = rs!b_memberid
    lblDateBorrow.Caption = rs!b_bordate
    lblDateRertuning.Caption = rs!b_retdate
    lblLendingRate.Caption = rs!b_amount
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "Select * from members WHERE memberid=" & member_id & "", con, adOpenKeyset, adLockOptimistic
    lblBorrower.Caption = rs!m_fullname
    rs.Close
errorhandler:
End Function

Private Sub lblDept_Change()

On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from departments WHERE departmentid=" & CInt(lblDept.Caption) & "", con, adOpenKeyset, adLockOptimistic
    lblDepartment.Caption = rs!d_name
    lblLendingRate.Caption = rs!d_lendingrate
    rs.Close
errorhandler:
End Sub

Private Sub txtBookSearch_KeyPress(KeyAscii As Integer)
On Error GoTo errorhandler
    If KeyAscii = vbKeyReturn Then
        findBook
    End If
    Exit Sub
errorhandler:
    lblBkResult.BackColor = &H80FFFF
    lblBkResult.Caption = "Book not found, please try to search again!"
    cmdClearBook.Enabled = False
    findBook
End Sub

Private Sub findBook()
    Set rs = New ADODB.Recordset
    rs.Open "Select * from books WHERE b_name LIKE '%" & txtBookSearch.Text & "%' ", con, adOpenKeyset, adLockOptimistic
    lblBookName.Caption = rs!b_name
    lblDept.Caption = rs!b_department
    lblBookNo.Caption = rs!b_number
    book_id = rs!bookid
    rs.Close
    
    If Not (lblBookName.Caption = "") Then
        GetBorrowingInfo
        lblBkResult.BackColor = &H80FFFF
        lblBkResult.Caption = "Search Book by name then press ENTER"
        cmdClearBook.Enabled = True
    End If
End Sub

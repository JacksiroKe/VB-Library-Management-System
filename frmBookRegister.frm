VERSION 5.00
Begin VB.Form frmBookRegister 
   Caption         =   "Register a NewBook - Shirikisho Library"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegisterClose 
      Caption         =   "Register and Close"
      Height          =   800
      Left            =   1560
      TabIndex        =   9
      Top             =   3240
      Width           =   4000
   End
   Begin VB.CommandButton cmdRegisterAdd 
      Caption         =   "Register and Add Another"
      Height          =   800
      Left            =   1560
      TabIndex        =   8
      Top             =   4320
      Width           =   4000
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   450
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   5000
   End
   Begin VB.TextBox txtRegdate 
      Height          =   450
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   5000
   End
   Begin VB.TextBox txtBookno 
      Height          =   450
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   5000
   End
   Begin VB.TextBox txtBookName 
      Height          =   450
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   5000
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registration Date:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Book Number:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Book Name:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmBookRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim departmnt As Integer, ddate As Date

Private Sub cmdRegisterAdd_Click()
    checkForm
    NewBook
    frmMain.Load_AllBooks
    clearData
End Sub

Private Sub cmdRegisterClose_Click()
    checkForm
    NewBook
    frmMain.Load_AllBooks
    clearData
    Unload Me
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
    ddate = DateValue(Now)
    txtRegdate.Text = ddate
    departmnt = 0
    Load_AllDepartments
End Sub
 
Private Sub clearData()
    txtBookName.Text = ""
    txtBookno.Text = ""
    txtRegdate.Text = ""
    cmbDepartment.Text = ""
    ddate = DateValue(Now)
    txtRegdate.Text = ddate
End Sub

Private Sub checkForm()
    If txtBookName.Text = "" Then
        txtBookName.BackColor = &HFF&
        txtBookno.BackColor = &HFFFFFF
        txtRegdate.BackColor = &HFFFFFF
        cmbDepartment.BackColor = &HFFFFFF
        txtBookName.SetFocus
        Exit Sub
    ElseIf txtBookno.Text = "" Then
        txtBookName.BackColor = &HFFFFFF
        txtBookno.BackColor = &HFF&
        txtRegdate.BackColor = &HFFFFFF
        cmbDepartment.BackColor = &HFFFFFF
        txtBookno.SetFocus
        Exit Sub
    ElseIf txtRegdate.Text = "" Then
        txtBookName.BackColor = &HFFFFFF
        txtBookno.BackColor = &HFFFFFF
        txtRegdate.BackColor = &HFF&
        cmbDepartment.BackColor = &HFFFFFF
        txtRegdate.SetFocus
        Exit Sub
    ElseIf cmbDepartment.Text = "" Then
        txtBookName.BackColor = &HFFFFFF
        txtBookno.BackColor = &HFFFFFF
        txtRegdate.BackColor = &HFFFFFF
        cmbDepartment.BackColor = &HFF&
        cmbDepartment.SetFocus
        Exit Sub
    Else
        cmbDepartment.BackColor = &HFFFFFF
        Exit Sub
    End If
End Sub

Private Sub NewBook()
    'On Error GoTo ErrorHandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from books", con, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!b_name = txtBookName.Text
    rs!b_number = txtBookno.Text
    rs!b_department = departmnt
    rs.Update
    rs.Close
    MsgBox "Well Done! Book registered succesfully.", vbInformation, App.Title
    'Exit Sub
'ErrorHandler:
'MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Load_AllDepartments()
cmbDepartment.Clear
Dim str As String
On Error GoTo ErrorHandlerr
 Set rs = New ADODB.Recordset
    rs.Open "Select * from departments", con, adOpenKeyset, adLockOptimistic
    Do Until rs.EOF
        cmbDepartment.AddItem rs!d_name
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
ErrorHandlerr:
End Sub
Private Sub cmbDepartment_Click()
    Set rs = New ADODB.Recordset
    rs.Open "Select * from departments WHERE d_name='" & cmbDepartment.Text & "'", con, adOpenKeyset, adLockOptimistic
    departmnt = rs!departmentid
    rs.Close
End Sub


VERSION 5.00
Begin VB.Form frmDepartNew 
   Caption         =   "Add a Department - Ndhiwa Hostels"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7890
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
   ScaleHeight     =   5940
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOverdue 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2760
      Width           =   4575
   End
   Begin VB.CommandButton cmdSaveClose 
      Caption         =   "Save and Close"
      Height          =   700
      Left            =   2040
      TabIndex        =   7
      Top             =   4800
      Width           =   4000
   End
   Begin VB.CommandButton cmdSaveAdd 
      Caption         =   "Save and Add Another"
      Height          =   700
      Left            =   2040
      TabIndex        =   6
      Top             =   3840
      Width           =   4000
   End
   Begin VB.TextBox txtLending 
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Overdue Rate:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lending Rate:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department Number:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department Name:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmDepartNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub cmdSaveAdd_Click()
    checkForm
    frmMain.Load_AllDepartments
End Sub

Private Sub cmdSaveClose_Click()
    checkForm
    frmMain.Load_AllDepartments
    Unload Me
End Sub

Private Sub checkForm()
    If txtName.Text = "" Then
        txtName.BackColor = &HFF&
        txtNumber.BackColor = &HFFFFFF
        txtLending.BackColor = &HFFFFFF
        txtOverdue.BackColor = &HFFFFFF
        txtName.SetFocus
        Exit Sub
    ElseIf txtNumber.Text = "" Then
        txtName.BackColor = &HFFFFFF
        txtNumber.BackColor = &HFF&
        txtLending.BackColor = &HFFFFFF
        txtOverdue.BackColor = &HFFFFFF
        txtNumber.SetFocus
        Exit Sub
    ElseIf txtLending.Text = "" Then
        txtName.BackColor = &HFFFFFF
        txtNumber.BackColor = &HFFFFFF
        txtLending.BackColor = &HFF&
        txtOverdue.BackColor = &HFFFFFF
        txtLending.SetFocus
        Exit Sub
    ElseIf txtOverdue.Text = "" Then
        txtName.BackColor = &HFFFFFF
        txtNumber.BackColor = &HFFFFFF
        txtLending.BackColor = &HFFFFFF
        txtOverdue.BackColor = &HFF&
        txtOverdue.SetFocus
        Exit Sub
    Else
        txtOverdue.BackColor = &HFFFFFF
        SaveDepartment
    End If
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
End Sub
 
Private Sub clearData()
    txtName.Text = ""
    txtNumber.Text = ""
    txtLending.Text = ""
    txtOverdue.Text = ""
End Sub

Private Sub SaveDepartment()
    
    On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from departments", con, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!d_name = txtName.Text
    Rs!d_number = txtNumber.Text
    Rs!d_lendingrate = txtLending.Text
    Rs!d_overduefee = txtOverdue.Text
    Rs.Update
    clearData
    Exit Sub
ErrorHandler:
'MsgBox Err.Description & " No. " & Err.Number
End Sub

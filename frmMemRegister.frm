VERSION 5.00
Begin VB.Form frmMemRegister 
   Caption         =   "Register a New Member - Shirikisho Library"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8010
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
   ScaleHeight     =   6780
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegisterClose 
      Caption         =   "Register and Close"
      Height          =   800
      Left            =   2160
      TabIndex        =   14
      Top             =   5760
      Width           =   4000
   End
   Begin VB.CommandButton cmdRegisterAdd 
      Caption         =   "Register and Add Another"
      Height          =   800
      Left            =   2160
      TabIndex        =   13
      Top             =   4680
      Width           =   4000
   End
   Begin VB.TextBox txtLocation 
      Height          =   1050
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   2160
      Width           =   5000
   End
   Begin VB.TextBox txtRegdate 
      Height          =   450
      Left            =   2640
      TabIndex        =   10
      Top             =   3960
      Width           =   5000
   End
   Begin VB.OptionButton OptGender 
      Caption         =   "Female"
      Height          =   330
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.OptionButton OptGender 
      Caption         =   "Male"
      Height          =   330
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtOccupation 
      Height          =   450
      Left            =   2640
      TabIndex        =   5
      Top             =   3360
      Width           =   5000
   End
   Begin VB.TextBox txtIdno 
      Height          =   450
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   5000
   End
   Begin VB.TextBox txtFullName 
      Height          =   450
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   5000
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   2640
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Physical Address:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registration Date:"
      ForeColor       =   &H80000008&
      Height          =   430
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gender:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Occupation:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID. Number:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Full Name:"
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmMemRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim sex As String, expexted As String, feesares As String
Dim ddate As Date

Private Sub cmdRegisterAdd_Click()
    checkForm
    NewMember
    clearData
End Sub

Private Sub cmdRegisterClose_Click()
    checkForm
    NewMember
    Unload Me
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
    ddate = DateValue(Now)
    txtRegdate.Text = ddate
End Sub
 
Private Sub clearData()
    txtFullName.Text = ""
    txtIdno.Text = ""
    txtLocation.Text = ""
    txtRegdate.Text = ""
    txtOccupation.Text = ""
End Sub

Private Sub checkForm()
    If txtFullName.Text = "" Then
        txtFullName.BackColor = &HFF&
        txtLocation.BackColor = &HFFFFFF
        txtRegdate.BackColor = &HFFFFFF
        txtOccupation.BackColor = &HFFFFFF
        txtFullName.SetFocus
        Exit Sub
    ElseIf txtLocation.Text = "" Then
        txtFullName.BackColor = &HFFFFFF
        txtLocation.BackColor = &HFF&
        txtRegdate.BackColor = &HFFFFFF
        txtOccupation.BackColor = &HFFFFFF
        txtLocation.SetFocus
        Exit Sub
    ElseIf txtRegdate.Text = "" Then
        txtFullName.BackColor = &HFFFFFF
        txtLocation.BackColor = &HFFFFFF
        txtIdno.BackColor = &HFFFFFF
        txtRegdate.BackColor = &HFF&
        txtOccupation.BackColor = &HFFFFFF
        txtRegdate.SetFocus
        Exit Sub
    ElseIf txtOccupation.Text = "" Then
        txtFullName.BackColor = &HFFFFFF
        txtLocation.BackColor = &HFFFFFF
        txtIdno.BackColor = &HFFFFFF
        txtRegdate.BackColor = &HFFFFFF
        txtOccupation.BackColor = &HFF&
        txtOccupation.SetFocus
        Exit Sub
    End If
End Sub

Private Sub NewMember()
    
    On Error GoTo errorhandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from Members", con, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!m_fullname = txtFullName.Text
    rs!m_idnumber = txtIdno.Text
    rs!m_location = txtLocation.Text
    rs!m_gender = sex
    rs!m_regdate = txtRegdate.Text
    rs!m_occupation = txtOccupation.Text
    rs.Update
    MsgBox "Well Done! Member registered succesfully.", vbInformation, App.Title
    frmMain.Load_AllMembers
    Exit Sub
errorhandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub lblLogin_Click()
    frmAdmLogin.Show
    Unload Me
End Sub

Private Sub OptGender_Click(Index As Integer)
    If (Index = 1) Then
        sex = "F"
    Else
        sex = "M"
    End If
End Sub

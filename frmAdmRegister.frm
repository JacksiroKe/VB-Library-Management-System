VERSION 5.00
Begin VB.Form frmAdmRegister 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Register on Shirikisho Library"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11550
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8070
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Register Your Account: "
      Height          =   6255
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   10575
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register"
         Height          =   765
         Left            =   2400
         TabIndex        =   13
         Top             =   5160
         Width           =   3135
      End
      Begin VB.TextBox txtPasswordcon 
         Height          =   645
         IMEMode         =   3  'DISABLE
         Left            =   4320
         PasswordChar    =   "."
         TabIndex        =   12
         Top             =   4200
         Width           =   5535
      End
      Begin VB.TextBox txtPassword 
         Height          =   645
         IMEMode         =   3  'DISABLE
         Left            =   4320
         PasswordChar    =   "."
         TabIndex        =   10
         Top             =   3480
         Width           =   5535
      End
      Begin VB.TextBox txtEmail 
         Height          =   645
         Left            =   4320
         TabIndex        =   8
         Top             =   2760
         Width           =   5535
      End
      Begin VB.TextBox txtUsername 
         Height          =   555
         Left            =   4320
         TabIndex        =   6
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox txtSecondName 
         Height          =   645
         Left            =   4320
         TabIndex        =   4
         Top             =   1440
         Width           =   5535
      End
      Begin VB.TextBox txtFirstName 
         Height          =   645
         Left            =   4320
         TabIndex        =   2
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label lblLogin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "     or Login"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   6720
         TabIndex        =   14
         Top             =   5160
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Confirm Password:"
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preffered Password:"
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email Address:"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UserName:"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Second Name:"
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "First Name:"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      Height          =   1095
      Left            =   1440
      Top             =   240
      Width           =   9015
   End
   Begin VB.Label Label7 
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
      Left            =   1560
      TabIndex        =   15
      Top             =   360
      Width           =   8775
   End
End
Attribute VB_Name = "frmAdmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
End Sub
 
Private Sub clearData()
    txtFirstName.Text = ""
    txtSecondName.Text = ""
    txtUsername.Text = ""
    txtEmail.Text = ""
    txtPassword.Text = ""
    txtPasswordcon.Text = ""
End Sub

Private Sub cmdRegister_Click()
    If txtFirstName.Text = "" Then
        txtFirstName.BackColor = &HFF&
        txtSecondName.BackColor = &HFFFFFF
        txtUsername.BackColor = &HFFFFFF
        txtEmail.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFFFFFF
        txtPasswordcon.BackColor = &HFFFFFF
        txtFirstName.SetFocus
        Exit Sub
    ElseIf txtSecondName.Text = "" Then
        txtFirstName.BackColor = &HFFFFFF
        txtSecondName.BackColor = &HFF&
        txtUsername.BackColor = &HFFFFFF
        txtEmail.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFFFFFF
        txtPasswordcon.BackColor = &HFFFFFF
        txtSecondName.SetFocus
        Exit Sub
    ElseIf txtUsername.Text = "" Then
        txtFirstName.BackColor = &HFFFFFF
        txtSecondName.BackColor = &HFFFFFF
        txtUsername.BackColor = &HFF&
        txtEmail.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFFFFFF
        txtPasswordcon.BackColor = &HFFFFFF
        txtUsername.SetFocus
        Exit Sub
    ElseIf txtEmail.Text = "" Then
        txtFirstName.BackColor = &HFFFFFF
        txtSecondName.BackColor = &HFFFFFF
        txtUsername.BackColor = &HFFFFFF
        txtEmail.BackColor = &HFF&
        txtPassword.BackColor = &HFFFFFF
        txtPasswordcon.BackColor = &HFFFFFF
        txtEmail.SetFocus
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        txtFirstName.BackColor = &HFFFFFF
        txtSecondName.BackColor = &HFFFFFF
        txtUsername.BackColor = &HFFFFFF
        txtEmail.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFF&
        txtPasswordcon.BackColor = &HFFFFFF
        txtPassword.SetFocus
        Exit Sub
    ElseIf txtPasswordcon.Text = "" Then
        txtFirstName.BackColor = &HFFFFFF
        txtSecondName.BackColor = &HFFFFFF
        txtUsername.BackColor = &HFFFFFF
        txtEmail.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFFFFFF
        txtPasswordcon.BackColor = &HFF&
        txtPasswordcon.SetFocus
        Exit Sub
    Else
        txtPasswordcon.BackColor = &HFFFFFF
        
        If txtPassword.Text = txtPasswordcon.Text Then
           RegisterMe
        Else
           txtPasswordcon.BackColor = &HFF&
           txtPasswordcon.SetFocus
           MsgBox "Passwords do not match!", vbCritical, App.Title
           Exit Sub
        End If
    End If
    
End Sub

Private Sub RegisterMe()
    
    On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from admnistrators", con, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!a_firstname = txtFirstName.Text
    Rs!a_seconame = txtSecondName.Text
    Rs!a_username = txtUsername.Text
    Rs!a_email = txtEmail.Text
    Rs!a_password = txtPasswordcon.Text
    Rs.Update
    
    clearData
    MsgBox "Well Done " & txtFirstName.Text & " " & txtSecondName.Text & "! You have registered succesfully. You can Now Login to your Account.", vbInformation, App.Title
    frmAdmLogin.Show
    Unload Me
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub lblLogin_Click()
    frmAdmLogin.Show
    Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmAdmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login to Shirikisho Library"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.Frame fraLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Login to Continue: "
      Height          =   3495
      Left            =   840
      TabIndex        =   0
      Top             =   2880
      Width           =   8415
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   645
         Left            =   2160
         TabIndex        =   5
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         Height          =   555
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "."
         TabIndex        =   3
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox txtUserName 
         Height          =   555
         Left            =   3360
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblRegister 
         BackColor       =   &H00FFFFFF&
         Caption         =   "or Register"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5280
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password:"
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username:"
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   600
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
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   8775
   End
   Begin VB.Label lblLoggedin 
      Caption         =   "loggedin"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmAdmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
'Public OK As Boolean
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Dim loggedin As String, my_username As String, my_password As String


 Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
    
End Sub
   
Private Sub cmdLogin_Click()
    If txtUsername.Text = "" Then
        txtUsername.BackColor = &HFF&
        txtPassword.BackColor = &HFFFFFF
        txtUsername.SetFocus
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        txtUsername.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFF&
        txtPassword.SetFocus
        Exit Sub
    Else
        my_username = txtUsername.Text
        my_password = txtPassword.Text
        LoginMe
    End If
End Sub

Private Sub LoginMe()
    
On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT * from admnistrators WHERE a_username = '" & my_username & "' AND a_password = '" & my_password & "'", con, adOpenKeyset, adLockOptimistic
    lblLoggedin.Caption = Rs!a_username
    If Len(lblLoggedin.Caption) = 0 Then
        fraLogin.Caption = " Invalid password or Username! "
        txtUsername.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFF&
        txtPassword.Text = ""
        txtPassword.SetFocus
    Else
        frmMain.Show
        frmMain.lblLoggedin.Caption = lblLoggedin.Caption
        Unload Me
    End If
    Exit Sub
ErrorHandler:     MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub lblRegister_Click()
    frmAdmRegister.Show
    Unload Me
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        LoginMe
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        LoginMe
    End If
End Sub

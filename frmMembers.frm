VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMembers 
   Caption         =   "Shirikisho Library Members"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   18441
      _Version        =   393216
      Rows            =   7
      Cols            =   10
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorSel    =   12632064
      ForeColorSel    =   255
      BackColorBkg    =   4210752
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      TextStyle       =   4
      TextStyleFixed  =   3
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   3
      MousePointer    =   1
      FormatString    =   " |m_fullname                    |     m_location       |  m_idnumber  | m_gender   | m_occupation   |  m_regdate |  m_freetoborrow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim ddate As Date, dtime As Date

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\ShirikishoLibrary.mdb;"
    con.Open
     
    Load_Items
End Sub

Public Sub Load_Items()
    Set rs = New ADODB.Recordset
    rs.Open "Select * from members ORDER BY memberid ASC", con, adOpenKeyset, adLockOptimistic
    
    MSFlexGrid1.Rows = rs.RecordCount + 1
    MSFlexGrid1.Cols = rs.Fields.Count
    MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    MSFlexGrid1.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    MSFlexGrid1.Row = 1
End Sub

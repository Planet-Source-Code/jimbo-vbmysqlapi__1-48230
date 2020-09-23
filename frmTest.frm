VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB mySQL API (Test) by Jim Banasiak"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "SQL Operations"
      Height          =   3015
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdTables 
         Caption         =   "Tables"
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmdStat 
         Caption         =   "Status"
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "Shutdown"
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "Ping"
         Height          =   375
         Left            =   2760
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdListDBs 
         Caption         =   "Databases"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdListProcess 
         Caption         =   "Processes"
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdExec 
         Caption         =   "Execute"
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtQuery 
         Height          =   765
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmTest.frx":0000
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Query"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connect to MySQL Provider"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdCon 
         Caption         =   "Connect"
         Height          =   375
         Left            =   720
         TabIndex        =   26
         ToolTipText     =   "Performs a mysql_connect...nice and fast"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use Compression"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Text            =   "3306"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtDB 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Text            =   "mysql"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmdDiscon 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdRCon 
         Caption         =   "RealConnect"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         ToolTipText     =   "Performs a mysql_real_connect.  Takes longer..."
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "test"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "localhost"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Database"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Host Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   2  'Align Bottom
      Bindings        =   "frmTest.frx":0013
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   3300
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   495
      Left            =   3120
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  VBMySQL APi version .01                                                     #
'#  Copyright (C) 2000  Jim Banasiak           <itsjimbo@yahoo.com>                                 #
'#                                                                              #
'#  This program is free software; you can redistribute it and/or               #
'#  modify it under the terms of the GNU General Public License                 #
'#  as published by the Free Software Foundation; either version 2              #
'#  of the License, or (at your option) any later version.                      #
'#                                                                              #
'#  This program is distributed in the hope that it will be useful,             #
'#  but WITHOUT ANY WARRANTY; without even the implied warranty of              #
'#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               #
'#  GNU General Public License for more details.                                #
'#                                                                              #
'#  You should have received a copy of the GNU General Public License           #
'#  along with this program; if not, write to the Free Software                 #
'#  Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA. #
'#                                                                              #
'################################################################################
Option Explicit
Private mysql As cMysql
Private mrec As ADODB.Recordset

Private Sub cmdCon_Click()
On Error GoTo errhandler
    Set mysql = New cMysql                     'initialize everything
    mysql.connect txtHost, txtUser, txtPass
    EnableControls
    Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub EnableControls()
'everything went well..enable all controls
cmdDiscon.Enabled = True
cmdCon.Enabled = False
cmdRCon.Enabled = False
cmdExec.Enabled = True
cmdListProcess.Enabled = True
cmdListDBs.Enabled = True
cmdPing.Enabled = True
cmdShutdown.Enabled = True
cmdStat.Enabled = True
Populate_DBS
End Sub
Private Sub Populate_DBS()
'populate the databases into combo1
    Dim a As Recordset
    Dim i As Long
    Set a = mysql.list_dbs("")
    With a
    .MoveFirst
    Do Until .eof
      Combo1.AddItem .Fields(0).Value
      .MoveNext
    Loop
    End With
    Set a = Nothing
End Sub

Private Sub DisableControls()
cmdDiscon.Enabled = False
cmdCon.Enabled = True
cmdRCon.Enabled = True
cmdExec.Enabled = False
cmdListProcess.Enabled = False
cmdListDBs.Enabled = False
cmdPing.Enabled = False
cmdShutdown.Enabled = False
cmdStat.Enabled = False
cmdTables.Enabled = False
Combo1.Clear
End Sub

Private Sub cmdDiscon_Click()
    Set mysql = Nothing    'close up the connection
    If mrec Is Nothing Then Else mrec.Close
    Set mrec = Nothing
    DisableControls
End Sub

Private Sub cmdExec_Click()
On Error GoTo errhandler
    Set mrec = mysql.query(txtQuery.Text)
    'mrec.Open
    Set Data1.Recordset = mrec
    Exit Sub
errhandler:
 Select Case Err.Number
    Case 1046
     MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
            " not a mysql_real_connect you have to choose a " & vbCrLf & _
            " database, so Please Choose a database."
     Combo1.SetFocus
    Case Else
    MsgBox "[" & Err.Number & "] " & Err.Description
 End Select
End Sub

Private Sub cmdListDBs_Click()
  Set Data1.Recordset = mysql.list_dbs
End Sub

Private Sub cmdListProcess_Click()
  Set Data1.Recordset = mysql.list_processes
End Sub

Private Sub cmdPing_Click()
On Error GoTo errhandler
  mysql.ping
  MsgBox "Mysql Server is alive"
  Exit Sub
errhandler:
   MsgBox Err.Description
End Sub

Private Sub cmdRCon_Click()
On Error GoTo errhandler
    Dim mysql_options As Long
    Set mysql = New cMysql                     'initialize everything
    If Check1.Value Then mysql_options = 32 Else mysql_options = 0
    
    mysql.real_connect txtHost, txtUser, txtPass, txtDB, CLng(txtPort), , mysql_options
    EnableControls
    Debug.Print "Your thread_id is " & Str(mysql.thread_id)
    Debug.Print "Thread Safe is "; mysql.thread_safe
    Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub cmdShutdown_Click()
On Error GoTo errhandler
  mysql.shutdown
  Exit Sub
errhandler:
   MsgBox Err.Description
End Sub
Private Sub cmdStat_Click()
On Error GoTo errhandler
  Debug.Print "Client is "; mysql.get_client_info
  Debug.Print "Host is "; mysql.get_host_info
  Debug.Print "Proto is "; mysql.get_proto_info
  Debug.Print "Server is "; mysql.get_server_info
  MsgBox mysql.stat
  Exit Sub
errhandler:
   MsgBox Err.Description
End Sub

Private Sub cmdTables_Click()
  If Combo1.Text <> "" Then
    Set Data1.Recordset = mysql.list_tables
  End If
End Sub

Private Sub Combo1_Click()
  If Combo1.Text <> "" Then
   mysql.select_db Combo1.Text
   cmdTables.Enabled = True
  End If
End Sub

Private Sub Form_Load()
DisableControls
MsgBox "---------------------------------------------------------------------------------------" & vbCrLf & _
       "  VBMySQL APi version .01, Copyright (C) 2000-2001 Jim Banasiak         " & vbCrLf & _
       "                                           itsjimbo@yahoo.com                             " & vbCrLf & _
       "---------------------------------------------------------------------------------------" & vbCrLf & _
       "  VBMySQL APi comes with ABSOLUTELY NO WARRANTY; for details       " & vbCrLf & _
       "  see the source code.                                             " & vbCrLf & _
       "---------------------------------------------------------------------------------------", , "Welcome!"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdDiscon_Click
End Sub

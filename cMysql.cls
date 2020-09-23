VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMysql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################################'
'#  VBMySQL APi version .01                                                     #
'#  Copyright (C) 2000  Jim Banasiak                                            #
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
'/****************************************************************************/'
' Class Wrapper: cMySQL
'             We have to 'wrap' around pointers to various mysql API calls
'             to hide pointers and that other garbage for our Vb heros
'             This will use the API's from libmysql and translate
'/****************************************************************************/'
Option Explicit
Private myConHandle As Long        'valid connection pointer for this class
Private myRec As API_MYSQL         'valid record to hold ourselves
Private myRec_res As API_MYSQL_RES 'this is we we will store results for ourselves
Private myRec_field As API_MYSQL_FIELD
Private myRec_rows As API_MYSQL_ROWS
Private mUser As String
Private mPassword As String
Private mHost As String
Private mDb As String
Private mUnix_socket As String
Private mClientFlag As Long
Private mPort As Long
'mysql_init() --handled upon creation
Private Sub Class_Initialize()
    myConHandle = API_mysql_init(myRec)
    If myConHandle Then
      CopyMemory myRec, ByVal myConHandle, LenB(myRec) 'move the memory back from pointer
    Else
      Err.Raise 512, "cMysql::Initialize", "Cannot Initialize a mysql record::Fatal"
    End If
End Sub
'mysql_close() --handled upon destruction, unconditionally close
Private Sub Class_Terminate()
    If myConHandle <> 0 Then Call API_mysql_close(myRec)
End Sub
Public Sub real_connect(Optional host As String = vbNullString, Optional user As String = vbNullString, Optional Passwd As String = vbNullString, Optional DB As String = vbNullString, Optional Port As Long = MYSQL_PORT, Optional Unix_Socket As String = vbNullString, Optional clientflag As Long = 0)
   If Len(host) > 0 Then mHost = host
   If Len(user) > 0 Then mUser = user
   If Len(Passwd) > 0 Then mPassword = Passwd
   If Len(DB) > 0 Then mDb = DB
   If Len(Unix_Socket) > 0 Then mUnix_socket = Unix_Socket
   If Port <> MYSQL_PORT Then mPort = Port
   If clientflag <> 0 Then mClientFlag = clientflag
   
    myConHandle = API_mysql_real_connect(myRec, StrPtr(StrConv(mHost, vbFromUnicode)), StrPtr(StrConv(mUser, vbFromUnicode)), StrPtr(StrConv(mPassword, vbFromUnicode)), StrPtr(StrConv(mDb, vbFromUnicode)), Port, StrPtr(StrConv(mUnix_socket, vbFromUnicode)), mClientFlag)
    If myConHandle Then
      CopyMemory myRec, ByVal myConHandle, LenB(myRec)
    Else
      Err.Raise API_mysql_errno(myRec), "cMysql::real_connect", ptr2str(API_mysql_error(myRec))
    End If
End Sub
Public Sub connect(Optional host As String = vbNullString, Optional user As String = vbNullString, Optional Passwd As String = vbNullString)
   If Len(host) > 0 Then mHost = host
   If Len(user) > 0 Then mUser = user
   If Len(Passwd) > 0 Then mPassword = Passwd
    myConHandle = API_mysql_connect(myRec, StrPtr(StrConv(mHost, vbFromUnicode)), StrPtr(StrConv(mUser, vbFromUnicode)), StrPtr(StrConv(mPassword, vbFromUnicode)))
    If myConHandle Then
       CopyMemory myRec, ByVal myConHandle, LenB(myRec) 'again we have to mooo-ve the memory
    Else
        Err.Raise API_mysql_errno(myRec), "cMysql::connect", ptr2str(API_mysql_error(myRec))
    End If
End Sub
Public Function affected_rows() As Long
   affected_rows = convert642l(API_mysql_affected_rows(myRec))
End Function
Public Sub create_db(DB As String)
Dim ret As Long
   ret = API_mysql_create_db(myRec, StrPtr(StrConv(DB, vbFromUnicode)))
   If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::create_db", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub debug_push(DBUG_PUSH As String) ' does a DBUG_PUSH with the given string.
  Call API_mysql_debug(StrPtr(StrConv(DBUG_PUSH, vbFromUnicode)))
End Sub
Public Sub drop_db(DB As String)
Dim ret As Long
   ret = API_mysql_drop_db(myRec, StrPtr(StrConv(DB, vbFromUnicode)))
   If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::drop_db", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub dump_debug_info() 'forces server to dump debug info
   Dim ret As Long
   ret = API_mysql_dump_debug_info(myRec)
   If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::dump_debug_info", ptr2str(API_mysql_error(myRec))
End Sub
Public Function eof() As Boolean
   eof = API_mysql_eof(myRec_res)
End Function
Public Function err_no() As Long
  err_no = API_mysql_errno(myRec)
End Function
Public Function err_description() As String
  err_description = ptr2str(API_mysql_error(myRec))
End Function
Public Sub escape_string(to_ As String, from_ As String, from_length As Long)
Dim ret As Long
  ret = API_mysql_escape_string(myRec, StrPtr(StrConv(to_, vbFromUnicode)), StrPtr(StrConv(from_, vbFromUnicode)), from_length)
End Sub
Public Sub free_result()
  'free the memeory
  Dim ret As Long
  ret = API_mysql_free_result(myRec_res)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::free_result", ptr2str(API_mysql_error(myRec))
End Sub
Public Function get_client_info() As String
Dim ret As Long
 ret = API_mysql_get_client_info
 If ret <> 0 Then get_client_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_client_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function get_host_info() As String
Dim ret As Long
 ret = API_mysql_get_host_info(myRec)
 If ret <> 0 Then get_host_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_host_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function get_proto_info() As String
Dim ret As Long
 ret = API_mysql_get_proto_info(myRec)
 If ret <> 0 Then get_proto_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_proto_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function get_server_info()
Dim ret As Long
 ret = API_mysql_get_server_info(myRec)
 If ret <> 0 Then get_server_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_server_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function info() As String
Dim ret As Long
 ret = API_mysql_info(myRec)
 If ret <> 0 Then info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::info", ptr2str(API_mysql_error(myRec))
End Function
Public Function Insert_ID() As Long
  Insert_ID = convert642l(API_mysql_insert_id(myRec))
End Function
Public Sub Kill(PID As Long)
    Dim ret As Long
    ret = API_mysql_kill(myRec, PID)
    If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::kill", ptr2str(API_mysql_error(myRec))
End Sub
Public Function list_dbs(Optional wild As String = "%") As Recordset
Dim ret As Long
  ret = API_mysql_list_dbs(myRec, StrPtr(StrConv(wild, vbFromUnicode)))
  If ret Then
  CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
  Set list_dbs = ConvertResult(myRec_res)
  End If
End Function
Public Function list_fields(Optional table As String = vbNullString, Optional wild As String = "%") As Recordset
Dim ret As Long
  ret = API_mysql_list_fields(myRec, StrPtr(StrConv(table, vbFromUnicode)), StrPtr(StrConv(wild, vbFromUnicode)))
  If ret Then
  CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
  Set list_fields = ConvertResult(myRec_res)
  End If
End Function
Public Function list_processes() As Recordset
Dim ret As Long
  ret = API_mysql_list_processes(myRec)
  If ret Then
  CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
  Set list_processes = ConvertResult(myRec_res)
  End If
End Function
Public Function list_tables(Optional wild As String = "%") As Recordset
Dim ret As Long
  ret = API_mysql_list_tables(myRec, StrPtr(StrConv(wild, vbFromUnicode)))
  If ret Then
  CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
  Set list_tables = ConvertResult(myRec_res)
  End If
End Function
Public Function num_fields() As Long
  num_fields = API_mysql_num_fields(myRec_res)
End Function
Public Function num_rows() As Long
  num_rows = convert642l(API_mysql_num_rows(myRec_res))
End Function

'#############################################################
'     These are things I left out because they are mostly
'     used behind the scenes when converting to a recordset (i'm lazy...)
'#############################################################
'don't know if we need to set options
'Public Sub options(l As API_mysql_option, Optional arg As String = vbNullString)
'Dim ret As Long
'ret = API_mysql_options(myRec, , StrPtr(StrConv(arg, vbFromUnicode)))
'If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::options", ptr2str(API_mysql_error(myRec))
'End Sub
'Public Sub row_seek() 'Seeks to a row in a result set, using value returned from mysql_row_tell().
'  API_mysql_row_seek
'  myRec_rows.data
'End Sub
'Public Sub row_tell() 'current_cursor position
'Dim ret As Long
'  ret = API_mysql_row_tell(myRec_res) 'pointer to mysql_rows
'  CopyMemory myRec_rows, ret, LenB(myRec_rows)
'End Sub
'Public Sub fetch_field()
'Public Sub fetch_field_direct()
'Public Sub fetch_fields()
'Public Sub fetch_lengths()
'Public Sub fetch_row()
'Public Sub field_count()
'Public Sub field_seek()
'Public Sub field_tell()
Private Sub data_seek(row As Long)  'Seeks to an arbitrary row in a query result set.
  Dim a As Double
  CopyMemory a, convertl264(row).bytes(1), 8
   Call API_mysql_data_seek(myRec_res, a)
End Sub

Public Sub ping()
  Dim ret As Long
  ret = API_mysql_ping(myRec)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::ping", ptr2str(API_mysql_error(myRec))
End Sub
Public Function query(SQL As String) As Recordset
Dim ret As Long
 ret = API_mysql_query(myRec, StrPtr(StrConv(SQL, vbFromUnicode)))
 If ret = 0 Then 'query was good
  ret = API_mysql_store_result(myRec)
   If ret Then
     CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
     Set query = ConvertResult(myRec_res)
   End If
  Else
  Err.Raise API_mysql_errno(myRec), "cMysql::query", ptr2str(API_mysql_error(myRec))
  End If
End Function
Public Function real_query(query As String, length As Long) As Recordset
Dim ret As Long
Dim rs As Recordset
 ret = API_mysql_real_query(myRec, StrPtr(StrConv(query, vbFromUnicode)), length)
 If ret = 0 Then 'query was good
  ret = API_mysql_store_result(myRec)
   If ret Then
     CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
     Set real_query = ConvertResult(myRec_res)
   End If
  Else
  Err.Raise API_mysql_errno(myRec), "cMysql::real_query", ptr2str(API_mysql_error(myRec))
  End If
End Function
Public Sub Refresh(refresh_options As API_refresh_options)
Dim ret As Long
  ret = API_mysql_refresh(myRec, refresh_options)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::refresh", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub select_db(DB As String)
Dim ret As Long
  ret = API_mysql_select_db(myRec, StrPtr(StrConv(DB, vbFromUnicode)))
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::select_db", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub shutdown()
Dim ret As Long
  ret = API_mysql_shutdown(myRec)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::shutdown", ptr2str(API_mysql_error(myRec))
End Sub
Public Function stat() As String
Dim ret As Long
 ret = API_mysql_stat(myRec)
 If ret <> 0 Then stat = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::stat", ptr2str(API_mysql_error(myRec))
End Function
Public Function thread_id() As Long
Dim ret As Long
 ret = API_mysql_thread_id(myRec)
 If ret <> 0 Then thread_id = ret Else Err.Raise API_mysql_errno(myRec), "cMysql::thread_id", ptr2str(API_mysql_error(myRec))
End Function
Public Function thread_safe() As Long
thread_safe = API_mysql_thread_safe
End Function

'##############################################################################'
'right now i automatically use these to store and convert into recordsets
'  Public Sub store_result()
'  Public Sub use_result()
'##############################################################################'

'##############################################################################'
'   Functions I needed to do some conversions...
'##############################################################################'
'Translate MYSQL_RES's into Recordsets so our users can use recordsets
' instead of MySQL proprietary types...
' NOTE: if you use ADO,DAO,RDO..change code here
'       this is actually the core of the translation
Private Function ConvertResult(m As API_MYSQL_RES) As Recordset
Dim m_fieldcount As Long
Dim m_rowcount As Long
Dim i As Long
Dim j As Long
Dim ret As Long
Dim s As String
Dim rs As Recordset
Dim PickUp() As Long


  Set rs = New Recordset
  m_fieldcount = myRec_res.field_count
  m_rowcount = convert642l(myRec_res.row_count)
  ReDim PickUp(1 To m_fieldcount) 'if we don't have any..mysql should error before we get into trouble
  With rs
          For i = 1 To m_fieldcount  'append the fields to the recordset
            ret = API_mysql_fetch_field(myRec_res)
            If ret Then
             CopyMemory myRec_field, ByVal ret, LenB(myRec_field)
             '#################################################################
             ' you may need to alter the line below to meet your specifications
             ' if you use a different type of recordset. As of now everything
             ' is a bstr..i'll change this later
             .Fields.Append ptr2str(myRec_field.name), adBSTR, myRec_field.length
             '#################################################################
             End If
          Next
          .Open
            For j = 1 To m_rowcount  'append rows to the recordset
               ret = API_mysql_fetch_row(myRec_res) 'fetch a row
               If ret Then
                CopyMemory PickUp(1), ByVal ret, SIZE_OF_CHAR * m_fieldcount 'copy it into array so we can pick it up
                .AddNew
                    For i = 1 To m_fieldcount
                       s = ptr2str(PickUp(i))
                      '#################################################################
                      ' you may need to alter the line below to meet your specifications
                      ' if you use a different type of recordset
                      .Fields(i - 1) = s
                      '#################################################################
                   Next i
               End If
            Next j
         .UpdateBatch adAffectAllChapters
  End With
  Set ConvertResult = rs
End Function
'convert an lpstrsz to a visual basic valid bstr
Private Function ptr2str(ByVal ptr As Long) As String
On Error Resume Next
    If ptr = 0 Then Exit Function    'yeah..should never happen though
    Dim test As Byte
    Dim sout() As Byte
    Dim cchars As Long
     cchars = lstrlen(ptr)
     If cchars = 0 Then Exit Function
     ReDim sout(1 To cchars)                   'byte array
     sout = String$(cchars, " ")
     CopyMemory sout(1), ByVal (ptr), cchars
    ptr2str = StripNull(StrConv(sout, vbUnicode))  'convert byte array (lpstr) to unicode
End Function
Private Function StripNull(sName As String) As String
   Dim f As Long
   f = InStr(sName, vbNullChar)
   If f > 0 Then
    StripNull = Trim(left$(sName, f - 1))
   Else
    StripNull = sName
   End If
End Function
' We must call this sub to calculate a pseudo-64 bit number
' because vb6 as of now doesn't support 64 bit unsigned integers
' so we convert a 64 to 32 bit integer (long)
' it is an 8 byte structure, so the last 4 we just NEVER use
' note: the limit on this is 2,147,483,647 (you could also transform into single)
'       so don't try and retrieve more than 2 billion records! :)
Private Function convert642l(a As API_myulonglong) As Long
Dim res As Long
  CopyMemory res, a.bytes(1), 4
  convert642l = res  'return our result from computations
End Function
' opposite of the above function:
' converts a long to a 64 bit ulonglong type
Private Function convertl264(l2convert As Long) As API_myulonglong
Dim a As API_myulonglong
  CopyMemory a.bytes(1), l2convert, 4
  convertl264 = a
End Function
'##############################################################################'
'#          Property Declarations..                                           #'
'#                         Host,Username,Password,db,port,clientflag,         #'
'#                                        unix_socket                         #'
'##############################################################################'
Public Property Let user(ByVal sUserName As String)
    mUser = sUserName
End Property
Public Property Get user() As String
   user = mUser
End Property
Public Property Let Password(ByVal sPassword As String)
   mPassword = sPassword
End Property
Public Property Get Password() As String
  Password = mPassword
End Property
Public Property Let host(ByVal sHost As String)
   mHost = sHost
End Property
Public Property Get host() As String
  host = mHost
End Property
Public Property Let Port(ByVal sPort As Long)
  mPort = sPort
End Property
Public Property Get Port() As Long
  Port = mPort
End Property
Public Property Let DB(ByVal sDB As String)
  mDb = sDB
End Property
Public Property Get DB() As String
  DB = mDb
End Property
Public Property Let ClientFlags(ByVal sClientFlags As Long)
 mClientFlag = sClientFlags
End Property
Public Property Get ClientFlags() As Long
 ClientFlags = mClientFlag
End Property
Public Property Let Unix_Socket(ByVal sUnix_socket As String)
  mUnix_socket = sUnix_socket
End Property
Public Property Get Unix_Socket() As String
  Unix_Socket = mUnix_socket
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DATABASE_OP_CLASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim DB As New ADODB.Connection
Dim rs As New ADODB.Recordset
Public errstring As Variant
' OPEN CONNECTION STRING
Public Function OPEN_CONNECTION(CSTRING As String)
        DB.Open CSTRING
End Function

' RETURN CONNECTION STRING
Public Function GET_CONNECTION_STRING() As String
    GET_CONNECTION_STRING = DB.ConnectionString
End Function

' OPEN RECORD SET
Public Function OPEN_RECORD_SET(ByRef RECORDSET_QUERY As String, ACTIVE_CONNECTION As String, CURSOR_TYPE As ADODB.CursorTypeEnum, LOCKTYPE As ADODB.LockTypeEnum)
        rs.Open RECORDSET_QUERY, ACTIVE_CONNECTION, CURSOR_TYPE, LOCKTYPE
End Function

' RETURN RECORD COUNT
Public Function GET_RECORD_COUNT() As Integer
        GET_RECORD_COUNT = rs.RecordCount
End Function
' RETURN ACTIVE CONNECTION
Public Function GET_ACTIVE_CONNECTION() As ADODB.Connection
    Set GET_ACTIVE_CONNECTION = DB
End Function

' RETURN RECORD SET
Public Function GET_RECORDSET() As ADODB.Recordset
    Set GET_RECORDSET = rs
End Function

'RETURN CURRENT RECORD VALUES
Public Function GET_CURRENT_RECORD(ByVal FIELD_NUMBER As Integer) As Variant
        GET_CURRENT_RECORD = rs.Fields(FIELD_NUMBER).Value
End Function

Public Function ADDNEW()
    rs.ADDNEW
End Function
'SET CURRENT RECORD VALUES
Public Sub SET_CURRENT_RECORD(ByVal FIELDNO As Integer, ByVal FVALUE As String)
        On Error Resume Next
        rs.Fields(FIELDNO).Value = FVALUE
End Sub

Public Function UPDATE() As Boolean
    On Error GoTo returnfalse
    If rs.CursorLocation = adUseClient Then
        rs.UpdateBatch
    Else
        rs.UPDATE
    End If
    UPDATE = True
    Exit Function
returnfalse:
    errstring = Err.Description
    UPDATE = False
End Function

Public Function DELETE() As Boolean
        On Error GoTo dis
        If rs.EOF <> True Then
                rs.DELETE
                If rs.EOF <> True Then
                    rs.MoveNext
                End If
        ElseIf rs.BOF <> True Then
                rs.DELETE
                If rs.BOF <> True Then
                    rs.MovePrevious
                End If
        Else
                MsgBox "All Records Deleted ...", vbInformation
        End If
        
        
        DELETE = True
        Exit Function
dis:
            errstring = Err.Description
            DELETE = False
End Function




Public Function CANCEL_UPDATE()
    If rs.CursorLocation = adUseClient Then
        rs.CancelBatch
    Else
        rs.CancelUpdate
    End If
End Function

Attribute VB_Name = "Module1"
Public CN As New Connection

'Enumerator for form state
Public Enum FormState
    adStateAddMode = 0
    adStateEditMode = 1
    adStatePopupMode = 2
    adStateViewMode = 3
End Enum

Public Sub OpenDB()
    CN.CursorLocation = adUseClient
          
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False;Jet OLEDB:Database Password=jaypee"
End Sub

Public Function getIndex(ByVal srcTable As String) As Long
    On Error GoTo erR
    Dim rs As New Recordset
    Dim RI As Long
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM [Key Generator] WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
    
    RI = rs.Fields("NextNo")
    CN.BeginTrans
    rs.Fields("NextNo") = RI + 1
    rs.Update
    CN.CommitTrans
    getIndex = RI
    
    srcTable = ""
    RI = 0
    Set rs = Nothing
    Exit Function
erR:
        ''Error when incounter a null value
        If erR.Number = 94 Then
            getIndex = 1
            Resume Next
        Else
            MsgBox erR.Description
        End If
        CN.RollbackTrans
End Function

Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, CN, adOpenStatic, adLockOptimistic
    
    With srcDC
        .ListField = srcBindField
        .BoundColumn = srcColBound
        Set .RowSource = rs
        'Display the first record
        If ShowFirstRec = True Then
            If Not rs.RecordCount < 1 Then
                .BoundText = rs.Fields(srcColBound)
                .Tag = rs.RecordCount & "*~~~~~*" & rs.Fields(srcColBound)
            Else
                .Tag = "0*~~~~~*0"
            End If
        End If
    End With
    Set rs = Nothing
End Sub

Public Function getLynxGridPos(ByVal srcFlexGrd As LynxGrid, ByVal srcWhatCol As Integer, ByVal srcFindWhat As String) As Integer
    Dim R As Long, ret As Integer
    
    ret = -1 'Means not found
    For R = 0 To srcFlexGrd.Rows - 1
        If srcFlexGrd.CellText(R, srcWhatCol) = srcFindWhat Then ret = R: Exit For
    Next R
    
    getLynxGridPos = ret
    R = 0: ret = 0
End Function

Public Sub DelRecwSQL(ByVal sTable As String, ByVal sField As String, ByVal sString As String, ByVal isNumber As Boolean, ByVal snum As Long)
    If isNumber = True Then
        CN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " =" & snum
    Else
        CN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " ='" & sString & "'"
    End If
End Sub

Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double
    If srcCurrency = "" Then
        toNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        toNumber = retValue
        retValue = 0
    End If
End Function

Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(IIf(Trim(srcCurr) = "", 0, srcCurr), "#,##0.00")
End Function


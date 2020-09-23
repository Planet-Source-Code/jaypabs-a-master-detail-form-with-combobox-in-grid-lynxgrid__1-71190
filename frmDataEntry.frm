VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDataEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orders"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   345
      Left            =   5250
      TabIndex        =   11
      Top             =   1380
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   345
      Left            =   6180
      TabIndex        =   10
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   5580
      TabIndex        =   7
      Top             =   5340
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   405
      Left            =   4200
      TabIndex        =   6
      Top             =   5340
      Width           =   1335
   End
   Begin VB.TextBox txtNote 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1230
      Width           =   1965
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   435
      Left            =   1920
      TabIndex        =   2
      Top             =   750
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      _Version        =   393216
      Format          =   71761921
      CurrentDate     =   39728
   End
   Begin VB.TextBox txtInvoiceNo 
      Height          =   405
      Left            =   1920
      TabIndex        =   0
      Top             =   300
      Width           =   1905
   End
   Begin MSDataListLib.DataCombo dcProductName 
      Height          =   315
      Left            =   1590
      TabIndex        =   9
      Top             =   2940
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin Inventory.LynxGrid LynxGrid 
      Height          =   3165
      Left            =   330
      TabIndex        =   8
      Top             =   1860
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5583
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRectMode   =   2
      Appearance      =   0
      AllowColumnResizing=   -1  'True
      Editable        =   -1  'True
      AllowInsert     =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Note"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   780
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice No"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   330
      Width           =   1005
   End
End
Attribute VB_Name = "frmDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OrderID              As Long
Public State                As FormState
Dim rs                      As New Recordset
Dim lngCol                  As Long
Dim lngRow                  As Long

Private Sub CmdAdd_Click()
    LynxGrid.AddItem
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    LynxGrid.DeleteSelected
    
    HideButtons
End Sub

Private Sub CmdSave_Click()
    On Error GoTo erR

    If Trim(txtInvoiceNo.Text) = "" Then
        MsgBox "Please enter Invoice No.", vbInformation
        nsdName.SetFocus
        
        Exit Sub
    End If
    
    CN.BeginTrans

    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        
        rs.Fields("OrderID") = OrderID
    End If
    
    With rs
      .Fields("InvoiceNo") = txtInvoiceNo.Text
      .Fields("Date") = dtpDate.Value
      .Fields("Note") = txtNote.Text
      
      .Update
    End With

    Dim rsDetails As New Recordset

    rsDetails.CursorLocation = adUseClient
    rsDetails.Open "SELECT * FROM [qry_Order_Details] WHERE OrderID=" & OrderID, CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With LynxGrid
        'Save the details of the records
        If LynxGrid.ItemCount > 0 Then
            For c = 0 To LynxGrid.ItemCount - 1
                .Row = c
                If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                    rsDetails.AddNew
    
                    rsDetails![OrderID] = OrderID
                    rsDetails![ProductID] = .CellText(c, 2)
                    rsDetails![Qty] = .CellText(c, 4)
                    rsDetails![Price] = .CellText(c, 5)
    
                    rsDetails.Update
                ElseIf State = adStateEditMode Then
                    rsDetails.Filter = "OrderDetailID = " & toNumber(.CellText(c, 0))
                
                    If rsDetails.RecordCount = 0 Then GoTo AddNew
    
                    rsDetails![OrderID] = OrderID
                    rsDetails![ProductID] = .CellText(c, 2)
                    rsDetails![Qty] = .CellText(c, 4)
                    rsDetails![Price] = .CellText(c, 5)
    
                    rsDetails.Update
                End If
            Next c
        End If
    End With

    'Clear variables
    c = 0
    Set rsDetails = Nothing
    
    CN.CommitTrans

    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            
            OrderID = getIndex("Orders")
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub

erR:
    CN.RollbackTrans
    MsgBox erR.Description, vbExclamation
    Screen.MousePointer = vbDefault
End Sub

Private Sub dcProductName_Click(Area As Integer)
    'update the cell programmatically because if you leave the current row
    'without clicking the other row in lynxgrid the column is not updated
    LynxGrid.CellText(lngRow, 2) = dcProductName.BoundText
End Sub

Private Sub Form_Load()
    initGrid
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Orders WHERE OrderID = " & OrderID, CN, adOpenStatic, adLockOptimistic
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        
        dtpBirthDate.Value = Date
        
        cmdUsrHistory.Enabled = False
        PK = getIndex("Oders")
    Else
        Caption = "Edit Entry"
        EditForm
    End If
    
    bind_dc "SELECT * FROM [Products] ORDER BY ProductName ASC", "ProductName", dcProductName, "ProductID", True
End Sub

Private Sub initGrid()
   With LynxGrid
      'Setting Redraw to False stops the Grid redrawing when Items/Cells are
      'changed which makes adding data much faster (and stops application flickering)
      .Redraw = False

      'EditTrigger defines which actions toggle Cell Edits. You can use multiple
      'Triggers by using "Or" as below
      .EditTrigger = lgEnterKey Or lgF2Key Or lgMouseClick

      'The height used for each Row
      .MinRowHeight = 315

      'Create the Columns
      .AddColumn "OrderDetailID", 0
      .AddColumn "OrderID", 0
      .AddColumn "ProductID", 0
      .AddColumn "Product Name", 3000
      .AddColumn "Qty", 1000
      .AddColumn "Price", 1000, lgAlignRightCenter

      'Bind the external Controls to the Column
      .BindControl 3, dcProductName, lgBCLeft Or lgBCTop Or lgBCWidth

      'Tell the grid to Draw again!
      .Redraw = True
   End With
End Sub

Private Sub EditForm()
    On Error GoTo erR

    With rs
        txtInvoiceNo.Text = .Fields("InvoiceNo")
        dtpDate.Value = .Fields("Date")
        txtNote.Text = .Fields("Note")
    End With
    
    'Display the details
    Dim rsDetails As New Recordset

    Dim li As Long
    
    rsDetails.CursorLocation = adUseClient
    rsDetails.Open "SELECT * FROM [qry_Order_Details] WHERE OrderID=" & OrderID, CN, adOpenStatic, adLockOptimistic
    
    If rsDetails.RecordCount > 0 Then
        rsDetails.MoveFirst
        While Not rsDetails.EOF
          cIRowCount = cIRowCount + 1     'increment
            With LynxGrid
                li = .AddItem(rsDetails.Fields("OrderDetailID"))
                
                .CellText(li, 1) = rsDetails![OrderID]
                .CellText(li, 2) = rsDetails![ProductID]
                .CellText(li, 3) = rsDetails![ProductName]
                .CellText(li, 4) = rsDetails![Qty]
                .CellText(li, 5) = toMoney(rsDetails![Price])
            End With
            rsDetails.MoveNext
        Wend
        LynxGrid.Row = 1
    End If

    rsDetails.Close
    'Clear variables
    Set rsDetails = Nothing
    Exit Sub
erR:
        If erR.Number = 94 Then Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDataEntry = Nothing
End Sub

Private Sub LynxGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)
    Debug.Print "LynxGrid_RequestUpdate"
    
    Select Case Col
    Case 3
        NewValue = dcProductName.Text
        LynxGrid.CellText(Row, 2) = dcProductName.BoundText
    End Select
End Sub

Private Sub LynxGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    lngCol = Col
    lngRow = Row
    
    Debug.Print "LynxGrid_RequestEdit"
     
    Select Case Col
    Case 3
        dcProductName.Text = LynxGrid.CellText(Row, Col)
    End Select
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim rsDetails As New Recordset

    rsDetails.CursorLocation = adUseClient
    rsDetails.Open "SELECT * FROM [Order Details] WHERE OrderID=" & OrderID, CN, adOpenStatic, adLockOptimistic
    If rsDetails.RecordCount > 0 Then
        rsDetails.MoveFirst
        While Not rsDetails.EOF
            CurrRow = getLynxGridPos(LynxGrid, 0, rsDetails!OrderDetailID)
        
            'Add to grid
            With LynxGrid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "[Order Details]", "OrderDetailID", "", True, rsDetails!OrderDetailID
                End If
            End With
            rsDetails.MoveNext
        Wend
    End If
End Sub

Private Sub ResetFields()
    txtInvoiceNo.Text = ""
    txtNote.Text = ""
    
    Dim c As Integer

    With LynxGrid
        'Save the details of the records
        If LynxGrid.ItemCount > 0 Then
            For c = 0 To LynxGrid.ItemCount - 1
                .Row = c

                .DeleteSelected
            Next c
        End If
    End With

    'Clear variables
    c = 0
    
    LynxGrid.Redraw = True
    
    HideButtons
End Sub

Private Sub HideButtons()
    dcProductName.Visible = False
End Sub

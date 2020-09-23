VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orders"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7890
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView1 
      Height          =   4545
      Left            =   270
      TabIndex        =   0
      Top             =   810
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   8017
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Order ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Visit http://www.sourcecodester.com"
      Height          =   195
      Left            =   4350
      TabIndex        =   2
      Top             =   330
      Width           =   2865
   End
   Begin VB.Label Label 
      Caption         =   "Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   300
      TabIndex        =   1
      Top             =   180
      Width           =   2535
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    Dim rs As New Recordset
    Dim itmX As ListItem

    OpenDB
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Orders", CN, adOpenStatic, adLockOptimistic
    
    While Not rs.EOF
        Set itmX = ListView1.ListItems. _
            Add(, , CStr(rs!OrderID))   ' Author.
        
        itmX.SubItems(1) = CStr(rs!InvoiceNo)
        itmX.SubItems(2) = rs![Date]
        
        rs.MoveNext   ' Move to next record.
    Wend
    
    DisplayAsURL Label1
End Sub

Private Sub Label1_Click()
    BrowseTo "http://www.sourcecodester.com"
End Sub

Private Sub ListView1_DblClick()
    With frmDataEntry
        .OrderID = ListView1.SelectedItem
        .State = adStateEditMode
                        
        .Show vbModal
    End With
End Sub

Private Sub BrowseTo(ByRef pstrURL As String)
    ' Opens users default web browser and navigates to the selected URL
    Call ShellExecute(Me.hwnd, "Open", pstrURL, "", "", True)
End Sub

Private Sub DisplayAsURL(ByRef Link As VB.Label)
    ' Changes a link to look like a URL
    Link.Font.Underline = True
    Link.ForeColor = vbBlue
    Link.MousePointer = vbCustom
    Link.MouseIcon = LoadPicture(App.Path & "\Hand.cur")
End Sub


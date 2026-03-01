VERSION 5.00
Begin VB.Form frmInventory 
   Caption         =   "Inventory Manager"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtReorderLevel 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtQuantity 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox txtSKU 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox lstItems 
      Height          =   5985
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   4395
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Module-level globals — typical VB6 pattern
Dim gConn As ADODB.Connection
Dim gRS As ADODB.Recordset
Dim gCurrentID As Long
Dim gDirtyFlag As Boolean

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Set gConn = New ADODB.Connection
    gConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                             "Data Source=" & App.Path & "\inventory.mdb"
    gConn.Open
    
    Call LoadItemList
    gDirtyFlag = False
    Exit Sub
    
ErrHandler:
    MsgBox "Database error: " & Err.Description & " (" & Err.Number & ")", _
           vbCritical, "Connection Failed"
End Sub

Private Sub LoadItemList()
    On Error GoTo ErrHandler
    
    lstItems.Clear
    
    Set gRS = New ADODB.Recordset
    gRS.Open "SELECT ItemID, SKU, Description FROM Items ORDER BY SKU", _
             gConn, adOpenStatic, adLockReadOnly
    
    Do While Not gRS.EOF
        lstItems.AddItem gRS("SKU") & " - " & gRS("Description")
        lstItems.ItemData(lstItems.NewIndex) = gRS("ItemID")
        gRS.MoveNext
    Loop
    
    gRS.Close
    Set gRS = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error loading items: " & Err.Description, vbExclamation, "Load Error"
End Sub

Private Sub lstItems_Click()
    On Error GoTo ErrHandler
    
    If lstItems.ListIndex = -1 Then Exit Sub
    
    gCurrentID = lstItems.ItemData(lstItems.ListIndex)
    
    Set gRS = New ADODB.Recordset
    gRS.Open "SELECT * FROM Items WHERE ItemID = " & gCurrentID, _
             gConn, adOpenStatic, adLockReadOnly
    
    If Not gRS.EOF Then
        txtSKU.Text         = Nz(gRS("SKU"), "")
        txtDescription.Text = Nz(gRS("Description"), "")
        txtPrice.Text       = Format(Nz(gRS("Price"), 0), "0.00")
        txtQuantity.Text    = Nz(gRS("QuantityOnHand"), 0)
        txtReorderLevel.Text = Nz(gRS("ReorderLevel"), 0)
    End If
    
    gRS.Close
    Set gRS = Nothing
    gDirtyFlag = False
    Exit Sub
    
ErrHandler:
    MsgBox "Error loading item: " & Err.Description, vbExclamation, "Load Error"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    If txtSKU.Text = "" Then
        MsgBox "SKU is required.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    Dim sql As String
    
    If gCurrentID = 0 Then
        sql = "INSERT INTO Items (SKU, Description, Price, QuantityOnHand, ReorderLevel) " & _
              "VALUES ('" & EscapeStr(txtSKU.Text) & "', " & _
              "'" & EscapeStr(txtDescription.Text) & "', " & _
              CDbl(txtPrice.Text) & ", " & _
              CLng(txtQuantity.Text) & ", " & _
              CLng(txtReorderLevel.Text) & ")"
    Else
        sql = "UPDATE Items SET " & _
              "SKU = '" & EscapeStr(txtSKU.Text) & "', " & _
              "Description = '" & EscapeStr(txtDescription.Text) & "', " & _
              "Price = " & CDbl(txtPrice.Text) & ", " & _
              "QuantityOnHand = " & CLng(txtQuantity.Text) & ", " & _
              "ReorderLevel = " & CLng(txtReorderLevel.Text) & " " & _
              "WHERE ItemID = " & gCurrentID
    End If
    
    gConn.Execute sql
    gDirtyFlag = False
    Call LoadItemList
    MsgBox "Saved.", vbInformation, "Inventory"
    Exit Sub
    
ErrHandler:
    MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
End Sub

Private Sub cmdDelete_Click()
    If gCurrentID = 0 Then Exit Sub
    If MsgBox("Delete this item?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    On Error GoTo ErrHandler
    gConn.Execute "DELETE FROM Items WHERE ItemID = " & gCurrentID
    gCurrentID = 0
    Call ClearForm
    Call LoadItemList
    Exit Sub
    
ErrHandler:
    MsgBox "Delete failed: " & Err.Description, vbCritical, "Delete Error"
End Sub

Private Sub cmdNew_Click()
    gCurrentID = 0
    Call ClearForm
    txtSKU.SetFocus
End Sub

Private Sub ClearForm()
    txtSKU.Text = ""
    txtDescription.Text = ""
    txtPrice.Text = "0.00"
    txtQuantity.Text = "0"
    txtReorderLevel.Text = "0"
    gDirtyFlag = False
End Sub

Private Function EscapeStr(s As String) As String
    EscapeStr = Replace(s, "'", "''")
End Function

Private Sub Form_Unload(Cancel As Integer)
    If gDirtyFlag Then
        If MsgBox("You have unsaved changes. Exit anyway?", vbYesNo) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If Not gConn Is Nothing Then
        If gConn.State = adStateOpen Then gConn.Close
    End If
    Set gConn = Nothing
End Sub

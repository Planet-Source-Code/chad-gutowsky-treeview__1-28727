VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   2475
   ClientTop       =   1710
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6465
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   5280
      TabIndex        =   8
      Top             =   3600
      Width           =   2535
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load from Table"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save to Table"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Treeview"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Node"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CommandButton cmdChild 
         Caption         =   "Add Child"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Add Previous Sibling"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Add Next Sibling"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Add Last Sibling"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "Add First Sibling"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9763
      _Version        =   327680
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "bldtree.frx":0000
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "bldtree.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "bldtree.frx":0336
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example shows how to save and restore the nodes of a TreeView Control
'to an Access database. It will let you add nodes to the TreeView
'using the five different relationships of the Add Method (Nodes Collection).
'The numbering of the Nodes is seqental. So every time you add a Node the Node number
'is increameted. The text shows what button was used to add a Node. So you can read the
'Node text as follows: <button used> <Node count> example: Next 4

Option Explicit
Dim mDB As Database
Dim mRS As Recordset
Dim mnIndex As Integer  ' Holds the index of a Node
Dim mbIndrag As Boolean ' Flag that signals a Drag Drop operation.
Dim moDragNode As Object ' Item that is being dragged.






Private Sub cmdChild_Click()
'Add a node using tvwChild
    Dim oNodex As Node
    Dim skey As String
    Dim iIndex As Integer
        
    On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    iIndex = TreeView1.SelectedItem.Index 'Check to see if a Node is selected
    skey = GetNextKey() ' Get a key for the new Node
    Set oNodex = TreeView1.Nodes.Add(iIndex, tvwChild, skey, "Child " & skey, 1, 2)
    oNodex.EnsureVisible 'make sure the child node is visible
    Exit Sub
myerr:
    'Display a messge telling the user to select a node
    MsgBox ("You must select a Node to do an Add Child" & vbCrLf _
       & "If the TreeView is empty us Add Last to create the first node")
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    Cls
    TreeView1.Nodes.Clear
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
'Add a node using tvwFirst
    Dim skey As String
    Dim iIndex As Integer
    
    On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    iIndex = TreeView1.SelectedItem.Index 'Check to see if a Node is selected
    skey = GetNextKey() ' Get a key for the new Node
    TreeView1.Nodes.Add iIndex, tvwFirst, skey, "First " & skey, 1, 2
    Exit Sub
myerr:
    'Display a messge telling the user to select a node
    MsgBox ("You must select a Node to do an Add First" & vbCrLf _
       & "If the TreeView is empty us Add Last to create the first node")
    Exit Sub
    
End Sub

Private Sub cmdLast_Click()
'Add a node using tvwLast
    Dim skey As String
    skey = GetNextKey()
    On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    TreeView1.Nodes.Add TreeView1.SelectedItem.Index, tvwLast, skey, "Last " & skey, 1, 2
    Exit Sub
myerr:
    'add a root node in the last postion because no other node is selected.
    TreeView1.Nodes.Add , tvwLast, skey, "Last " & skey, 1, 2
    Exit Sub
End Sub

Private Sub cmdLoad_Click()
    LoadFromTable
End Sub

Private Sub GetFirstParent()
'Find the first parent node in the TreeView
    On Error GoTo myerr
    Dim i As Integer
    Dim nTmp As Integer
    
    For i = 1 To TreeView1.Nodes.Count
        'This will give an error if there is no parent
        nTmp = TreeView1.Nodes(i).Parent.Index
    Next
    Exit Sub
    
myerr:
    mnIndex = i
    Exit Sub
End Sub

Private Function GetNextKey() As String
'Returns a new key value for each Node being added to the TreeView
'This algorithm is very simple and will limit you to adding a total of 999 nodes
'Each node needs a unique key. If you allow users to remove Nodes you can't use
'the Nodes count +1 as the key for a new node.

    Dim sNewKey As String
    Dim iHold As Integer
    Dim i As Integer
    On Error GoTo myerr
    'The next line will return error #35600 if there are no Nodes in the TreeView
    iHold = Val(TreeView1.Nodes(1).Key)
    For i = 1 To TreeView1.Nodes.Count
        If Val(TreeView1.Nodes(i).Key) > iHold Then
            iHold = Val(TreeView1.Nodes(i).Key)
        End If
    Next
    iHold = iHold + 1
    sNewKey = CStr(iHold) & "_"
    GetNextKey = sNewKey 'Return a unique key
    Exit Function
myerr:
    'Because the TreeView is empty return a 1 for the key of the first Node
    GetNextKey = "1_"
    Exit Function
End Function

Private Sub LoadFromTable()
'asks the user to find the database and then restores the Nodes
'in the treeview control from the table
    Dim oNodex As Node
    Dim nImage As Integer
    Dim nSelectedImage As Integer
    Dim i As Integer
    Dim sTableNames As String
    Dim sNodeTable As String
    
    MsgBox "You must save a table before you can load one"
    nImage = 0
    nSelectedImage = 0
    CommonDialog1.Filter = "Access Database (*.MDB)|* .mdb"
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.filename) < 1 Then
        MsgBox ("No File was Choosen")
        Exit Sub
    End If
    
    Set mDB = DBEngine.Workspaces(0).OpenDatabase(CommonDialog1.filename)
    For i = 0 To mDB.TableDefs.Count - 1 'TableDefs is a 0 based collection
        If Left(mDB.TableDefs(i).Name, 4) <> "MSys" Then
            sTableNames = sTableNames & mDB.TableDefs(i).Name & ", "
        End If
    Next
    sNodeTable = InputBox(sTableNames, "Enter the Table that you saved the TreeView in")
    'Set mrs = db.OpenRecordset("loadnodes")
    'Debug.Print
    TreeView1.Nodes.Clear 'Clear the TreeView of any nodes
    Set mRS = mDB.OpenRecordset(sNodeTable)
    If mRS.RecordCount > 0 Then 'make sure there are records in the table
        mRS.MoveFirst
        Do While mRS.EOF = False
            nImage = mRS.Fields("image")
            nSelectedImage = mRS.Fields("selectedimage")
            If Trim(mRS.Fields("parent")) = "0_" Then 'All root nodes have 0_ in the parent field
                Set oNodex = TreeView1.Nodes.Add(, 1, Trim(mRS.Fields("key")), _
                  Trim(mRS.Fields("text")), nImage, nSelectedImage)
            Else 'All child nodes will have the parent key stored in the parent field
                Set oNodex = TreeView1.Nodes.Add(Trim(mRS.Fields("parent")), tvwChild, _
                   Trim(mRS.Fields("key")), Trim(mRS.Fields("text")), nImage, nSelectedImage)
                oNodex.EnsureVisible 'expend the TreeView so all nodes are visible
            End If
            mRS.MoveNext
        Loop
    End If
    mRS.Close 'Close the table
    mDB.Close 'Close the database
End Sub

Private Sub cmdNext_Click()
'Add a node using tvwNext
    Dim skey As String
    Dim iIndex As Integer
    
    On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    iIndex = TreeView1.SelectedItem.Index 'Check to see if a Node is selected
    skey = GetNextKey() ' Get a key for the new Node
    TreeView1.Nodes.Add iIndex, tvwNext, skey, "Next " & skey, 1, 2
    Exit Sub
myerr:
    'Display a messge telling the user to select a node
    MsgBox ("You must select a Node to do an Add Next" & vbCrLf _
       & "If the TreeView is empty us Add Last to create the first node")
    Exit Sub
End Sub

Private Sub cmdPrevious_Click()
'Add a node using tvwPrevious
    Dim skey As String
    Dim iIndex As Integer
    On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    iIndex = TreeView1.SelectedItem.Index 'Check to see if a Node is selected
    skey = GetNextKey() ' Get a key for the new Node
    TreeView1.Nodes.Add iIndex, tvwPrevious, skey, "Previous " & skey, 1, 2
    Exit Sub
myerr:
    'Display a messge telling the user to select a node
    MsgBox ("You must select a Node to do an Add Previous" & vbCrLf _
       & "If the TreeView is empty us Add Last to create the first node")
    Exit Sub
End Sub

Sub SaveToTable()
'Ask the user for the name of a mdb and table if it does not exist create it.
'Then store all of the nodes from the TreeView into the table.
    Dim sResponse As String
    Dim sMDBName As String
    Dim sTableName As String
    Dim i As Integer
           
    sResponse = MsgBox("Click YES to Create a new MDB", vbYesNo)
    If sResponse = vbYes Then
        CommonDialog1.Filter = "Access Database(*.MDB) |*.mdb"
        CommonDialog1.ShowSave
        If Len(CommonDialog1.filename) < 1 Then
            MsgBox ("You did not enter a file name")
            Exit Sub
        Else
            sTableName = InputBox("Enter a Table name", _
             "Table name to save Nodes to")

           If Len(sTableName) < 1 Then
                MsgBox ("You did not supply a table name")
                Exit Sub
           End If
            Set mDB = Workspaces(0).CreateDatabase(CommonDialog1.filename, _
              dbLangGeneral)
           CreateTable (sTableName) ' call the sub that creates a new table
           Set mRS = mDB.OpenRecordset(sTableName)
           WriteToTable  'Go to the sub that writes the nodes into the table
        End If
    Else
        CommonDialog1.Filter = "Access Database(*.MDB) |*.mdb"
        CommonDialog1.ShowOpen
        If Len(CommonDialog1.filename) < 1 Then
            MsgBox ("No File was Choosen")
            Exit Sub
        End If
    
        sTableName = InputBox("Enter a Table name", _
        "Table name to save Nodes to")
        
        If Len(sTableName) < 1 Then
            MsgBox ("You did not supply a table name")
            Exit Sub
        End If
        Set mDB = DBEngine.Workspaces(0).OpenDatabase(CommonDialog1.filename)
        For i = 0 To mDB.TableDefs.Count - 1 'TableDefs is a 0 based collection
            If mDB.TableDefs(i).Name = sTableName Then
            ' check to see if the table exist
                sResponse = MsgBox("All records in your table will be destroyed", vbOKCancel)
                If sResponse = vbOK Then
                    Set mRS = mDB.OpenRecordset(sTableName)
                    WriteToTable  'Go to the sub that writes the nodes into the table
                    mRS.Close 'close the recordset
                    mDB.Close 'close the database
                    Exit Sub
                Else
                    mDB.Close 'close the database
                    Exit Sub
                End If
            End If
        Next
        CreateTable (sTableName) ' call the sub that creates a new table
        Set mRS = mDB.OpenRecordset(sTableName)
        WriteToTable  'Go to the sub that writes the nodes into the table
    End If
mRS.Close 'close the recordset
mDB.Close 'close the database
End Sub

Private Sub cmdRemove_Click()
    'Remove the selected Node
    Dim iIndex As Integer
    
    On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    iIndex = TreeView1.SelectedItem.Index 'Check to see if a Node is selected
    TreeView1.Nodes.Remove iIndex 'Removes the Node and any children it has
    Exit Sub
myerr:
    'Display a messge telling the user to select a node
    MsgBox ("You must select a Node to do a Remove" & vbCrLf _
       & "If the TreeView is empty us Add Last to create the first node")
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    SaveToTable
End Sub

Function CreateTable(sTableName As String)
' Create a new table to store the Node information in
    Dim td As TableDef
    Dim fd As Field
    
    Set td = mDB.CreateTableDef(sTableName)
    Set fd = td.CreateField("key", dbText, 4)
    td.Fields.Append fd
    Set fd = td.CreateField("parent", dbText, 4)
    td.Fields.Append fd
    Set fd = td.CreateField("text", dbText, 20)
    td.Fields.Append fd
    Set fd = td.CreateField("image", dbInteger)
    td.Fields.Append fd
    Set fd = td.CreateField("selectedimage", dbInteger)
    td.Fields.Append fd
    mDB.TableDefs.Append td
           
End Function

Sub WriteToTable()
'Writes the Node information from the TreeView into a table
    Dim i As Integer
    Dim iTmp As Integer
    Dim iIndex As Integer
    
    If mRS.RecordCount > 0 Then
    ' Delete any records that may be in the table
        mRS.MoveFirst
        Do While mRS.EOF = False
            mRS.Delete
            mRS.MoveNext
        Loop
    End If
    
    GetFirstParent 'Find a root node in the treeview
    'get the index of the root node that is at the top of the treeview
    iIndex = TreeView1.Nodes(mnIndex).FirstSibling.Index
    iTmp = iIndex
    mRS.AddNew
    mRS("parent") = "0_" 'this is a root node
    mRS("key") = TreeView1.Nodes(iIndex).Key
    mRS("text") = TreeView1.Nodes(iIndex).Text
    mRS("image") = TreeView1.Nodes(iIndex).Image
    mRS("selectedimage") = TreeView1.Nodes(iIndex).SelectedImage
    mRS.Update
    'If the Node has Children call the sub that writes the children
    If TreeView1.Nodes(iIndex).Children > 0 Then
        WriteChild iIndex
    End If
    
    While iIndex <> TreeView1.Nodes(iTmp).LastSibling.Index
    'loop through all the root nodes
        mRS.AddNew
        mRS("parent") = "0_" 'this is a root node
        mRS("key") = TreeView1.Nodes(iIndex).Next.Key
        mRS("text") = TreeView1.Nodes(iIndex).Next.Text
        mRS("image") = TreeView1.Nodes(iIndex).Next.Image
        mRS("selectedimage") = TreeView1.Nodes(iIndex).Next.SelectedImage
        mRS.Update
         'If the Node has Children call the sub that writes the children
        If TreeView1.Nodes(iIndex).Next.Children > 0 Then
            WriteChild TreeView1.Nodes(iIndex).Next.Index
        End If
        ' Move to the Next root Node
        iIndex = TreeView1.Nodes(iIndex).Next.Index
    Wend
End Sub


Private Sub Form_Load()
    'Add some Nodes so the TreeView is not empty
    Set moDragNode = Nothing
    cmdLast_Click
    cmdLast_Click
    TreeView1.Nodes(1).Selected = True
    cmdChild_Click
End Sub

Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
   ' If user didn't move mouse or released it over an invalid area.

If TreeView1.DropHighlight Is Nothing Then
        mbIndrag = False
        Exit Sub
    Else
        ' Set dragged node's parent property to the target node.
        On Error GoTo checkerror ' To prevent circular errors.
        Set moDragNode.Parent = TreeView1.DropHighlight
        Cls
        Print TreeView1.DropHighlight.Text & _
        " is parent of " & moDragNode.Text
        ' Release the DropHighlight reference.
        Set TreeView1.DropHighlight = Nothing
        mbIndrag = False
        Set moDragNode = Nothing
        Exit Sub ' Exit if no errors occured.
    End If
 
checkerror:
    ' Define constants to represent Visual Basic errors code.
    Const CircularError = 35614
    If Err.Number = CircularError Then
        Dim msg As String
        msg = "A node can't be made a child of its own children."

' Display the message box with an exclamation mark icon
        ' and with OK and Cancel buttons.
        If MsgBox(msg, vbExclamation & vbOKCancel) = vbOK Then
            ' Release the DropHighlight reference.
            mbIndrag = False
            Set TreeView1.DropHighlight = Nothing
            Exit Sub
        End If
    End If


End Sub

Private Sub TreeView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If mbIndrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
        Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
    End If
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
    'Make sure we are over a Node
    If Not TreeView1.DropHighlight Is Nothing Then
        'Set the Node we are on to be the selected Node
        'if we don't do this it will not be the selected node
        'until we finish clicking on the Node
        TreeView1.SelectedItem = TreeView1.HitTest(x, y)
        Set moDragNode = TreeView1.SelectedItem ' Set the item being dragged.
    End If
    Set TreeView1.DropHighlight = Nothing
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then ' Signal a Drag operation.
        mbIndrag = True ' Set the flag to true.
        ' Set the drag icon with the CreateDragImage method.
        TreeView1.DragIcon = TreeView1.SelectedItem.CreateDragImage
        TreeView1.Drag vbBeginDrag ' Drag operation.
    End If

End Sub

Private Sub WriteChild(ByVal iNodeIndex As Integer)
' Write the child nodes to the table. This sub uses recursion
' to loop through the child nodes. It receives the Index of
' the node that has the children
    Dim i As Integer
    Dim iTempIndex As Integer
    iTempIndex = TreeView1.Nodes(iNodeIndex).Child.FirstSibling.Index
    'Loop through all a Parents Child Nodes
    For i = 1 To TreeView1.Nodes(iNodeIndex).Children
        mRS.AddNew
        mRS("parent") = TreeView1.Nodes(iTempIndex).Parent.Key
        mRS("key") = TreeView1.Nodes(iTempIndex).Key
        mRS("text") = TreeView1.Nodes(iTempIndex).Text
        mRS("image") = TreeView1.Nodes(iTempIndex).Image
        mRS("selectedimage") = TreeView1.Nodes(iTempIndex).SelectedImage
        mRS.Update
        ' If the Node we are on has a child call the Sub again
        If TreeView1.Nodes(iTempIndex).Children > 0 Then
            WriteChild (iTempIndex)
        End If
        ' If we are not on the last child move to the next child Node
        If i <> TreeView1.Nodes(iNodeIndex).Children Then
            iTempIndex = TreeView1.Nodes(iTempIndex).Next.Index
        End If
    Next i
End Sub



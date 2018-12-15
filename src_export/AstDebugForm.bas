Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private ProjectName As String

Public Sub DumpTokens(Tokens() As Token)
    Dim LB As ListBox
    
    Set LB = tokenList
    
    Dim I As Long
    LB.Clear
    For I = LBound(Tokens) To UBound(Tokens)
    LB.AddItem
    LB.Column(0, LB.ListCount - 1) = TokenToString(Tokens(I))
    LB.Column(1, LB.ListCount - 1) = Tokens(I).Text
    Next I
End Sub

Public Sub DumpTree(Root As AstNode)
    Dim TV As TreeView
    Dim I As Long
    Set TV = Ast
    TV.Nodes.Clear
    I = TV.Nodes.Add(, Text:="Root").Index
    DumpTreeIter Root, TV, I
End Sub

Public Sub DumpCode(C As CodeData)
    Dim S As Variable
    Dim F As Proc
    Dim L As Variant
    Dim Loc As String
    
    funcList.Clear
    
    For Each F In C.Functions
        Loc = ""
        For Each S In F.Variables
            Loc = Loc & " " & S.Name & "@" & S.Ref
        Next
        
        If F.IsUsed Then
            funcList.AddItem
            funcList.Column(0, funcList.ListCount - 1) = F.Name & "@" & CStr(F.Ref) & " (args: " & F.Args & ")"
            funcList.Column(1, funcList.ListCount - 1) = "FS=" & CStr(F.FrameSize) & " " & Loc
        End If
    Next
    
    globList.Clear
    
    For Each S In C.Globals
        globList.AddItem S.Name & "[" & S.Size & "]"
    Next
    
    asmListing.Text = C.HexDump
    labelSize.Caption = Str(C.HexSize) & "B (" & Format(C.HexSize / 1024, "0.00K") & ")"
    
    ProjectName = C.ProjectName
End Sub

Private Sub DumpTreeIter(N As AstNode, T As TreeView, ByVal I As Long)
    Dim Names
    Dim Ch As AstNode
    
    Select Case N.NodeType
    Case AnName, AnNumber
        Call T.Nodes.Add(I, 4, Text:=N.Value)
    Case AnString
        Call T.Nodes.Add(I, 4, Text:="""" & N.Value & """")
    Case AnChar
        Call T.Nodes.Add(I, 4, Text:="'" & N.Value & "'")
    Case AnSymbol
        Call T.Nodes.Add(I, 4, Text:="'" & N.Value)
    Case AnBlock
        I = T.Nodes.Add(I, 4, Text:="block").Index
        For Each Ch In N.Children
            DumpTreeIter Ch, T, I
        Next
    Case Else
        I = T.Nodes.Add(I, 4, Text:=N.NodeTypeName).Index
        For Each Ch In N.Children
            DumpTreeIter Ch, T, I
        Next
    End Select
End Sub

Private Sub buttonDebug_Click()
    Shell ThisDocument.Path & "\debug.cmd " & ThisDocument.Path & " " & ProjectName & " " & textArgs.Text
End Sub

Private Sub buttonSave_Click()
    Dim dialog
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.Show
    
    Dim txtDumpFs, txtDumpF
    
    Set txtDumpFs = CreateObject("Scripting.FileSystemObject")
    Set txtDumpF = txtDumpFs.CreateTextFile(dialog.SelectedItems.Item(1), True)
    
    txtDumpF.WriteLine asmListing.Text
    txtDumpF.Close
End Sub


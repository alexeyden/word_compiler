Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Public Sub MacroExpand(Root As AstNode)
    Dim Dict As Object
    
    Set Dict = CreateObject("Scripting.Dictionary")
    FindMacros Dict, Root
    
    WalkMacro Dict, Root
End Sub

Private Sub WalkMacro(Dict As Object, N As AstNode)
    Dim I As Integer
    Dim J As Integer
    I = 1
    While I <= N.Children.Count
        If N.Children(I).NodeType = AnBlock Then
            WalkMacro Dict, N.Children(I)
            
            If Dict.Exists(N.Children(I).BlockHead) Then
                InsertMacro N, I, Dict(N.Children(I).BlockHead)
            End If
        End If
        I = I + 1
    Wend
End Sub

Private Sub InsertMacro(N As AstNode, ByVal I As Integer, M As AstNode)
    Dim Ref As AstNode
    Dim Args As Object
    Dim J As Integer
    
    Set Args = CreateObject("Scripting.Dictionary")
    
    Set Ref = N.Children(I)
    
    Fail M.Children(3).Children.Count = Ref.Children.Count - 1, "Invalid number of arguments in macro call " & Ref.BlockHead, Ref
    
    For J = 1 To M.Children(3).Children.Count
        Args.Add M.Children(3).Children(J).Value, Ref.Children(J + 1)
    Next J
    
    N.Children.Add CopyMacroBody(Args, M.Children(4)), , , I
    N.Children.Remove I
End Sub

Private Function CopyMacroBody(Args As Object, N As AstNode) As AstNode
    Dim Clone As New AstNode
    Dim Ch As AstNode
    
    If N.NodeType = AnName And Args.Exists(N.Value) Then
        Set Clone = Args(N.Value)
    Else
        Clone.NodeType = N.NodeType
        Clone.Value = N.Value
        
        For Each Ch In N.Children
            Clone.Children.Add CopyMacroBody(Args, Ch)
        Next
    End If
    
    Set CopyMacroBody = Clone
End Function

Private Sub FindMacros(Dict As Variant, Root As AstNode)
    Dim N As AstNode
    Dim F As Proc
    Dim Name As String
    
    For Each N In Root.Children
        If N.BlockHead = "macro" Then
            Dict.Add N.Children(2).Value, N
        End If
    Next
End Sub


Private Function Fail(Cond As Boolean, Msg As String, N As AstNode)
    Utils.AssertError Cond, Msg & Chr(10) & "at line #" & N.LineNumber & Chr(10)
End Function



Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Public Function Parse(Tokens() As Token) As AstNode
Dim Root As New AstNode
Root.NodeType = AnBlock

Call ParseBlock(Tokens, 0, Root)

Set Parse = Root

End Function

Private Function ParseBlock(Tokens() As Token, I As Long, Parent As AstNode) As Boolean

Do While ParseExpr(Tokens, I, Parent)
Loop

End Function

Private Function ParseExpr(Tokens() As Token, I As Long, Parent As AstNode) As Boolean
Dim J As Long
Dim T As Token
Dim N As New AstNode

J = I

If I > UBound(Tokens) Then
    GoTo ExitError
End If

T = Tokens(I)
    
If Match(Tokens, I, TTNumber) Then
    N.NodeType = AnNumber
    N.Value = Val(T.Text)
ElseIf Match(Tokens, I, TTHexNumber) Then
    N.NodeType = AnNumber
    N.Value = Val("&H" & T.Text)
ElseIf Match(Tokens, I, TTBinNumber) Then
    N.NodeType = AnNumber
    N.Value = ParseBinary(T.Text)
ElseIf Match(Tokens, I, TTString) Then
    N.NodeType = AnString
    N.Value = T.Text
ElseIf Match(Tokens, I, TTChar) Then
    N.NodeType = AnChar
    N.Value = Asc(T.Text)
ElseIf Match(Tokens, I, TTSymbol) Then
    N.NodeType = AnSymbol
    N.Value = T.Text
ElseIf Match(Tokens, I, TTName) Then
    N.NodeType = AnName
    N.Value = T.Text
ElseIf Match(Tokens, I, TTPlus) Then
    N.NodeType = AnOpAdd
    N.Value = T.Text
ElseIf Match(Tokens, I, TTMinus) Then
    N.NodeType = AnOpSub
    N.Value = T.Text
ElseIf Match(Tokens, I, TTStar) Then
    N.NodeType = AnOpMul
    N.Value = T.Text
ElseIf Match(Tokens, I, TTSlash) Then
    N.NodeType = AnOpDiv
    N.Value = T.Text
ElseIf Match(Tokens, I, TTNot) Then
    N.NodeType = AnOpNot
    N.Value = T.Text
ElseIf Match(Tokens, I, TTAnd) Then
    N.NodeType = AnOpAnd
    N.Value = T.Text
ElseIf Match(Tokens, I, TTOr) Then
    N.NodeType = AnOpOr
    N.Value = T.Text
ElseIf Match(Tokens, I, TTLt) Then
    N.NodeType = AnOpLs
    N.Value = T.Text
ElseIf Match(Tokens, I, TTGt) Then
    N.NodeType = AnOpGt
    N.Value = T.Text
ElseIf Match(Tokens, I, TTEq) Then
    N.NodeType = AnOpEq
    N.Value = T.Text
ElseIf Match(Tokens, I, TTLe) Then
    N.NodeType = AnOpLe
    N.Value = T.Text
ElseIf Match(Tokens, I, TTGe) Then
    N.NodeType = AnOpGe
    N.Value = T.Text
ElseIf Match(Tokens, I, TTNeq) Then
    N.NodeType = AnOpNe
    N.Value = T.Text
ElseIf Match(Tokens, I, TTShl) Then
    N.NodeType = AnOpShl
    N.Value = T.Text
ElseIf Match(Tokens, I, TTShr) Then
    N.NodeType = AnOpShr
    N.Value = T.Text
ElseIf Match(Tokens, I, TTPercent) Then
    N.NodeType = AnOpRem
    N.Value = T.Text
ElseIf Match(Tokens, I, TTLPar) Then
    N.NodeType = AnBlock
    Do While ParseExpr(Tokens, I, N)
    Loop
    If Not Match(Tokens, I, TTRPar) Then
        AssertError False, "Cannot find matching closing bracked at line " & Tokens(I).PosB
    End If
ElseIf Match(Tokens, I, TTComment) Then
    ParseExpr = True
    Exit Function
ElseIf Match(Tokens, I, TTRPar) Then
    GoTo ExitError
Else
    AssertError False, "Unknown token " & Tokens(I).Text & " at line " & Tokens(I).PosB
End If

N.LineNumber = T.PosB

Parent.Children.Add N
ParseExpr = True

Exit Function

ExitError:

I = J
ParseExpr = False
End Function

Private Function ParseBinary(BinStr As String) As Integer
    Dim I As Integer
    Dim N As Integer
    Dim C As Integer
    
    N = 0
    
    For I = 1 To Len(BinStr)
        N = N + Val(Mid(BinStr, I, 1)) * 2 ^ (Len(BinStr) - I)
    Next I
    
    ParseBinary = N
End Function

Private Function Match(Tokens() As Token, I As Long, T As TokenType) As Boolean
If I > UBound(Tokens) Then
Match = False
ElseIf Tokens(I).Type = T Then
Match = True
I = I + 1
End If
End Function



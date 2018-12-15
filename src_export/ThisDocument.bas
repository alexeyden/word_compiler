Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub buttonCompile_Click()
    Dim Tokens() As Lexer.Token
    Dim Root As AstNode
    Dim Code As CodeData
    Dim Path As String
    Dim OutPath As String
    
    Path = "projects\" & ProjectName.Text & "\main.txt"
    OutPath = "projects\" & ProjectName.Text & "\" & ProjectName.Text & ".com"
    
    Path = ThisDocument.Path & "\" & Path
    OutPath = ThisDocument.Path & "\" & OutPath
    
    Tokens = Lexer.TokenizeFile(Path)
    
    Set Root = Parser.Parse(Tokens)
    Code = Compile.CompileAst(Root, ProjectName.Text, OutPath)
    
    ShowDebugView Tokens, Root, Code
End Sub

Private Sub ShowDebugView(Tokens() As Token, Root As AstNode, C As CodeData)
    Dim Form As New AstDebugForm
    
    Form.DumpTokens Tokens
    Form.DumpTree Root
    Form.DumpCode C
    
    Form.Show
End Sub


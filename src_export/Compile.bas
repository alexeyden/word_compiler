Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Public Type CodeData
    Constants As New Collection
    Functions As New Collection
    Globals As New Collection
    Imports As New Collection
    
    Main As Proc
    CurProc As Proc
    CurLoopContinue As Long
    CurLoopEnd As Long
    HexDump As String
    HexSize As Long
    ProjectName As String
End Type

Public Function CompileAst(Root As AstNode, ProjectName As String, FileName As String) As CodeData
    Dim C As CodeData
    Dim F As Proc
    Dim R As Long
    
    CompileImports C, Root
    MacroExpand Root
    
    AsmBegin (FileName)
    
    FindFunctions Root, C
    FindGlobals Root, C, ""
    
    C.ProjectName = ProjectName
    Set C.Main = FindProc(C, "main")
    
    FindGlobals C.Main.Node, C, "main@"
    
    CompileMain C, C.Main
    
    For Each F In C.Functions
        If F.Name = "main" Or Not F.IsUsed Then
            GoTo Continue
        End If
        
        R = AsmLabelLand(F.Label)
        F.Ref = R
        Set C.CurProc = F
        CompileFunction C, F
Continue:
    Next
    
    Set C.CurProc = Nothing
    
    AsmEnd
    
    Dim B As Byte, Col As Integer, I As Long
    
    Col = 1
    I = 0
    Open FileName For Binary As #1

    While Not EOF(1)
        Get #1, , B
        
        If Col = 1 Then
            C.HexDump = C.HexDump & Format(I, "0000 ")
        End If
        
        C.HexDump = C.HexDump & Replace(Format(Hex(B), ">@@"), " ", "0") & " "
        Col = Col + 1
        
        If Col > 16 Then
            Col = 1
            C.HexDump = C.HexDump & Chr(10)
        End If
        
        If Col Mod 9 = 0 Then
            C.HexDump = C.HexDump & " "
        End If
        I = I + 1
    Wend
    
    Close #1
    
    C.HexSize = FileLen(FileName)
    
    CompileAst = C
End Function

Public Sub CompileImports(Code As CodeData, Root As AstNode)
    Dim C As AstNode
    Dim I As Integer, J As Integer
    Dim N As AstNode, R As AstNode
    Dim ImportName As String
    Dim Tokens() As Token
    
    I = 1
    While I <= Root.Children.Count
        Set N = Root.Children(I)
        
        If N.BlockHead = "import" Then
            ImportName = N.Children(2).Value
            
            If Mid(ImportName, 1, 1) = "." Then
                Tokens = Lexer.TokenizeFile(ThisDocument.Path & "\projects\" & Code.ProjectName & "\" & Mid(ImportName, 2) & ".txt")
            Else
                Tokens = Lexer.TokenizeFile(ThisDocument.Path & "\lib\" & N.Children(2).Value & ".txt")
            End If
            
            Set R = Parser.Parse(Tokens)
            
            For J = 1 To Code.Imports.Count
                If Code.Imports(J) = N.Children(2).Value Then
                    GoTo NextStmt
                End If
            Next J
            
            Code.Imports.Add N.Children(2).Value
            
            CompileImports Code, R
            
            For Each C In R.Children
                Root.Children.Add C, , , I
            Next
            
            Root.Children.Remove I
            I = I + R.Children.Count - 1
        End If
NextStmt:
        I = I + 1
    Wend
End Sub

Public Sub CompileMain(Code As CodeData, Main As Proc)
    Dim LBypass As Long
    Dim Stmt As AstNode
    Dim I As Integer
    
    LBypass = AsmLabelNew()
    
    ' jmp <bypass>
    
    Asm AsmOpJmp, AsmWord, AsmConst, CInt(LBypass)

    CompileGlobals Code
    
    ' bypass:
    
    Call AsmLabelLand(LBypass)
    
    Set Code.CurProc = Main
    
    For I = 4 To Main.Node.Children.Count
        Set Stmt = Main.Node.Children.Item(I)
        CompileStmt Code, Stmt
        MarkUsedFunctions Code, Stmt
    Next I
    
    ' ret
    Asm AsmOpRet, AsmWord, AsmNone, AsmNone
End Sub

Private Sub CompileGlobals(C As CodeData)
    Dim G As Variable
    Dim T As AstNode
    Dim I As Integer
    Dim D As Integer
    
    For Each G In C.Globals
        G.Ref = AsmGetIp()
        
        For I = 1 To G.Data.Count
            D = G.Data(I)
            
            Fail Not (G.Size = 1 And D > &HFF), "Global variable " & G.Name & " value is too big", G.Node
            
            Asm IIf(G.Size = 1, AsmOpDb, AsmOpDw), IIf(G.Size = 1, AsmByte, AsmWord), AsmConst, D
        Next I
    Next
End Sub

Private Sub CompileLocals(C As CodeData, F As Proc)
    Dim L As Variable
    Dim T As AstNode
    Dim I As Integer
    Dim D
    
    ' stack L = local, A = arg
    ' RET_ADDR | L1 L2 L3 | A1 A2 A3
    
    For Each L In F.Variables
        I = 0
        For Each D In L.Data
            AsmGenSetValue CgTargetStack, IIf(L.Size = 1, AsmByte, AsmWord), CgFromConst, L.Ref + I, CgFromConst, CInt(D)
            I = I + L.Size
        Next
    Next
End Sub

Private Sub CompileFunction(C As CodeData, F As Proc)
    Dim Stmt As AstNode
    Dim I As Integer
    
    ' set locals
    CompileLocals C, F
    
    ' body
    For I = 4 To F.Node.Children.Count
        Set Stmt = F.Node.Children.Item(I)
        CompileStmt C, Stmt
    Next I
    
    ' mov ax,0
    Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, 0
    
    ' ret <F.FrameSize>
    Asm AsmOpRet, AsmWord, AsmConst, F.FrameSize
End Sub

Private Sub CompileRaw(C As CodeData, St As AstNode)
    Dim I As Integer
    Dim N As AstNode
    
    If St.BlockHead = "asm_offset" Then
        Dim Vr As Variable
        
        Set Vr = FindVar(C, St.Children(2).Value)
        
        Asm AsmOpDw, AsmWord, AsmConst, Vr.Ref
        
        Exit Sub
    End If
    
    If St.BlockHead = "asm_push" Then
        If St.Children(2).NodeType = AnNumber Or St.Children(2).NodeType = AnChar Then
            AsmGenPushConst CInt(St.Children(2).Value)
        Else
            CompileBlock C, St.Children(2)
        End If
        
        Exit Sub
    End If
    
    For I = 2 To St.Children.Count
        Set N = St.Children(I)
        
        Asm IIf(St.BlockHead = "asm_dw", AsmOpDw, AsmOpDb), _
            IIf(St.BlockHead = "asm_dw", AsmWord, AsmByte), AsmConst, N.Value
    Next I
    
End Sub

Private Sub CompileStmt(C As CodeData, St As AstNode)
    If St.NodeType <> AnBlock Then
        Exit Sub
    End If
    
    If St.BlockHead = "setw" Then
        CompileSetValue C, St, AsmWord
    ElseIf St.BlockHead = "setb" Then
        CompileSetValue C, St, AsmByte
    ElseIf St.BlockHead = "if" Then
        CompileIf C, St
    ElseIf St.BlockHead = "for" Then
        CompileFor C, St
    ElseIf St.BlockHead = "while" Then
        CompileWhile C, St
    ElseIf St.BlockHead = "return" Then
        CompileReturn C, St
    ElseIf St.BlockHead = "continue" Then
        Asm AsmOpJmp, AsmWord, AsmConst, C.CurLoopContinue
    ElseIf St.BlockHead = "break" Then
        Asm AsmOpJmp, AsmWord, AsmConst, C.CurLoopEnd
    ElseIf IsAsmKeyword(St.BlockHead) Then
        CompileRaw C, St
    ElseIf St.BlockHead = "dw" Then
    ElseIf St.BlockHead = "db" Then
    ElseIf St.BlockHead = "let" Then
    ElseIf St.Children.Count = 0 Then
    ElseIf St.Children.Count > 0 And St.Children(1).NodeType = AnBlock Then
        Dim I As Integer
        Dim Stmt As AstNode
        
        For I = 1 To St.Children.Count
            Set Stmt = St.Children.Item(I)
            CompileStmt C, Stmt
        Next I
    Else
        CompileCall C, St
    End If
End Sub

Private Sub CompileCall(C As CodeData, N As AstNode)
    Dim F As Proc
    Dim I As Long
    
    Set F = FindProc(C, N.Children(1).Value)
    
    AssertError Not (F Is Nothing), "Cannot find procedure or macros: " & N.Children(1).Value
    
    F.IsUsed = True
    
    Asm AsmOpPush, AsmWord, AsmReg, AsmCHBP  ' push bp
    
    ' compile args
    For I = 2 To N.Children.Count
        CompileBlock C, N.Children(I)
    Next I
    
    AsmGenCall F.Label, F.FrameSize - 2 * F.Args
End Sub

Private Sub CompileSetValue(C As CodeData, St As AstNode, Size As Integer)
    Dim Val As AstNode
    Dim Ref As Long
    Dim Var As Variable
    Dim Target As Integer
    Dim ValueSource As Integer
    Set Val = St.Children(3)

    ValueSource = IIf(St.Children(3).NodeType = AnNumber, CgFromConst, CgFromStack)
    
    If St.Children(2).NodeType = AnName Then
        If ValueSource = CgFromStack Then
            CompileBlock C, Val
        End If
        
        Set Var = FindVar(C, St.Children(2).Value)
        Target = IIf(Var.IsLocal, CgTargetStack, CgTargetGlobal)

        AsmGenSetValue Target, Size, CgFromConst, CInt(Var.Ref), ValueSource, IIf(ValueSource = CgFromConst, Val.Value, 0)
    Else
        If ValueSource = CgFromStack Then
            CompileBlock C, Val
        End If
        
        CompileBlock C, St.Children(2)
        
        AsmGenSetValue CgTargetGlobal, Size, CgFromStack, AsmNone, ValueSource, IIf(ValueSource = CgFromConst, CInt(Val.Value), 0)
    End If
End Sub

Private Sub CompileFor(C As CodeData, St As AstNode)
    ' Syntax: {for [.. init ..] [.. cond ..] [.. step ..] (.. body ..)}
    ' Example: {for [setw i 0] [< i 20] [inc! i] (... body ...)}
    ' Code generation scheme:
    ' <init code>
    ' start:
    ' if not <cond> goto <end>
    ' <body>
    ' <step code>
    ' goto start
    ' end:
    
    Dim I As Integer
    Dim Init As AstNode, Cond As AstNode, Step As AstNode
    Dim PrevLoopCont As Long, PrevLoopEnd As Long
    Dim LStart As Long, LEnd As Long, LContinue As Long
    
    Set Init = St.Children(2)
    Set Cond = St.Children(3)
    Set Step = St.Children(4)
    
    LContinue = AsmLabelNew()
    LEnd = AsmLabelNew()
    
    CompileStmt C, Init
    
    LStart = AsmLabelNew()
    AsmLabelLand LStart
    
    CompileBlock C, Cond
    AsmGenJumpOnZero LEnd
    
    PrevLoopCont = C.CurLoopContinue
    PrevLoopEnd = C.CurLoopEnd
    
    C.CurLoopContinue = LContinue
    C.CurLoopEnd = LEnd
    
    ' compile body
    For I = 5 To St.Children.Count
        CompileStmt C, St.Children(I)
    Next I
    
    C.CurLoopContinue = PrevLoopCont
    C.CurLoopEnd = PrevLoopEnd
    
    AsmLabelLand LContinue
    
    CompileStmt C, Step
    
    Asm AsmOpJmp, AsmWord, AsmConst, CInt(LStart)
    
    AsmLabelLand LEnd
End Sub

Private Sub CompileWhile(C As CodeData, St As AstNode)
    ' Syntax: {while [.. cond ..] (.. body ..)}
    ' Example: {while [< i 20] (... body ...)}
    ' Code generation scheme:
    ' start:
    ' if not <cond> goto <end>
    ' <body>
    ' goto start
    ' end:
    
    Dim I As Integer
    Dim Cond As AstNode
    Dim PrevLoopCont As Long, PrevLoopEnd As Long
    Dim LStart As Long, LEnd As Long, LContinue
    
    Set Cond = St.Children(2)
    
    LEnd = AsmLabelNew()
    LContinue = AsmLabelNew()
    
    LStart = AsmLabelNew()
    AsmLabelLand LStart
    
    CompileBlock C, Cond
    AsmGenJumpOnZero LEnd
    
    PrevLoopCont = C.CurLoopContinue
    PrevLoopEnd = C.CurLoopEnd
    
    C.CurLoopContinue = LContinue
    C.CurLoopEnd = LEnd
    
    ' compile body
    For I = 3 To St.Children.Count
        CompileStmt C, St.Children(I)
    Next I
    
    C.CurLoopContinue = PrevLoopCont
    C.CurLoopEnd = PrevLoopEnd
    
    AsmLabelLand (LContinue)
    
    Asm AsmOpJmp, AsmWord, AsmConst, CInt(LStart)
    
    AsmLabelLand LEnd
End Sub

Private Sub CompileIf(C As CodeData, St As AstNode)
    Dim I As Integer, J As Integer
    Dim Branch As AstNode
    Dim BrElse As AstNode
    Dim LExit As Long, LBypass As Long
    
    LExit = AsmLabelNew()
    
    ' Syntax:
    ' {if ([.. cond 1 ..] (.. branch 1 ..))
    '     ([.. cond 2 ..] (.. branch 2 ..))
    '     (else (.. branch else ..))}
    
    ' Gode generation scheme:
    
    ' if not <cond 1> goto <bypass 1>
    ' cond 1 is true code
    ' goto <exit>
    ' bypass 1:
    
    ' if not <cond 2> goto <bypass 2>
    ' cond 2 is true code
    ' goto <exit>
    ' bypass 2:
    
    ' <else code>
    ' exit:
    
    For I = 2 To St.Children.Count
        Set Branch = St.Children(I)
        
        If Branch.Children(1).Value <> "else" Then
            LBypass = AsmLabelNew()
            
            ' if not <condition> goto <bypass>
            
            CompileBlock C, Branch.Children(1)
            AsmGenJumpOnZero LBypass
            
            ' on true code
            For J = 2 To Branch.Children.Count
                CompileStmt C, Branch.Children(J)
            Next J
            
            Asm AsmOpJmp, AsmWord, AsmConst, CInt(LExit) ' jmp exit
            
            Call AsmLabelLand(LBypass) ' bypass:
        Else
            Set BrElse = Branch
        End If
    Next I
    
    If Not (BrElse Is Nothing) Then
        For J = 2 To BrElse.Children.Count
            CompileStmt C, BrElse.Children(J)
        Next J
    End If
    
    Call AsmLabelLand(LExit)
End Sub

Private Sub CompileReturn(C As CodeData, St As AstNode)
    If C.CurProc.Name <> "main" And St.Children.Count > 1 Then
        CompileBlock C, St.Children(2)
        Asm AsmOpPop, AsmWord, AsmReg, AsmALAX
        Asm AsmOpRet, AsmWord, AsmConst, C.CurProc.FrameSize
    Else
        Asm AsmOpRet, AsmWord, AsmNone, AsmNone
    End If
End Sub

Private Sub CompileBlock(C As CodeData, St As AstNode)
    Dim T As String
    Dim I As Long
    
    If St.NodeType = AnName Then
        If St.Value = "true" Then
            AsmGenPushConst 1
            Exit Sub
        ElseIf St.Value = "false" Then
            AsmGenPushConst 0
            Exit Sub
        End If
        
        Dim V As Variable
        Set V = FindVar(C, St.Value)
        
        If V.IsConst Then
            CompileBlock C, V.Node
            Exit Sub
        End If
        
        If V.IsLocal Then
            AsmGenGetValue CgTargetStack, IIf(V.Size = 2, AsmWord, AsmByte), CgFromConst, V.Ref
        Else
            AsmGenGetValue CgTargetGlobal, IIf(V.Size = 2, AsmWord, AsmByte), CgFromConst, V.Ref
        End If
        
        Exit Sub
    ElseIf St.NodeType = AnNumber Or St.NodeType = AnChar Then
        AsmGenPushConst CInt(St.Value)
        Exit Sub
    End If
    
    T = St.BlockHead
    
    If T = "+" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenMathOp CgMathAdd
        Next I
    ElseIf T = "-" Then
        CompileBlock C, St.Children(2)
        
        If St.Children.Count = 2 Then
            AsmGenLogOp CgLogNeg
            Exit Sub
        End If
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenMathOp CgMathSub
        Next I
    ElseIf T = "*" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenMathOp CgMathMul
        Next I
    ElseIf T = "imul" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenMathOp CgMathIMul
        Next I
    ElseIf T = "/" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenMathOp CgMathDiv
        Next
    ElseIf T = "idiv" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenMathOp CgMathIDiv
        Next I
    ElseIf T = "=" Then
        CompileBlock C, St.Children(2)
        CompileBlock C, St.Children(3)
        AsmGenLogOp CgLogEq
    ElseIf T = "<" Then
        CompileBlock C, St.Children(2)
        CompileBlock C, St.Children(3)
        AsmGenLogOp CgLogLt
    ElseIf T = ">" Then
        CompileBlock C, St.Children(2)
        CompileBlock C, St.Children(3)
        AsmGenLogOp CgLogGt
    ElseIf T = "<=" Then
        CompileBlock C, St.Children(2)
        CompileBlock C, St.Children(3)
        AsmGenLogOp CgLogLe
    ElseIf T = ">=" Then
        CompileBlock C, St.Children(2)
        CompileBlock C, St.Children(3)
        AsmGenLogOp CgLogGe
    ElseIf T = "!=" Then
        CompileBlock C, St.Children(2)
        CompileBlock C, St.Children(3)
        AsmGenLogOp CgLogNe
    ElseIf T = "!" Then
        CompileBlock C, St.Children(2)
        AsmGenLogOp CgLogNot
    ElseIf T = "<<" Then
        CompileBlock C, St.Children(2)
        AsmGenLogOp CgLogShl, St.Children(3).Value
    ElseIf T = ">>" Then
        CompileBlock C, St.Children(2)
        AsmGenLogOp CgLogShr, St.Children(3).Value
    ElseIf T = "|" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenLogOp CgLogOr
        Next I
    ElseIf T = "&" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenLogOp CgLogAnd
        Next I
    ElseIf T = "%" Then
        CompileBlock C, St.Children(2)
        
        For I = 3 To St.Children.Count
            CompileBlock C, St.Children(I)
            AsmGenMathOp CgMathRem
        Next
    ElseIf T = "ref" Then
        Dim Vr As Variable
        Set Vr = FindVar(C, St.Children(2).Value)
        
        If Vr.IsLocal Then
            AsmGenRefLocal Vr.Ref
        Else
            AsmGenPushConst Vr.Ref
        End If
    ElseIf T = "getb" Then
        CompileBlock C, St.Children(2)
        AsmGenGetValue CgTargetGlobal, AsmByte, CgFromStack, AsmNone
    ElseIf T = "getw" Then
        CompileBlock C, St.Children(2)
        AsmGenGetValue CgTargetGlobal, AsmWord, CgFromStack, AsmNone
    ElseIf T = "uref" Then
        CompileBlock C, St.Children(2)
    ElseIf T = "sizeof" Then
        Set Vr = FindVar(C, St.Children(2).Value)
        AsmGenPushConst Vr.Size * Vr.Data.Count
    ElseIf IsAsmKeyword(T) Then
        CompileRaw C, St
    ElseIf St.Children.Count > 0 And St.Children(1).NodeType = AnBlock Then
        Dim Stmt As AstNode
        
        For I = 1 To St.Children.Count
            Set Stmt = St.Children.Item(I)
            CompileBlock C, Stmt
        Next I
    Else
        CompileCall C, St
        
        ' push ax
        Asm AsmOpPush, AsmWord, AsmReg, AsmALAX
    End If
End Sub

Private Function FindProc(C As CodeData, Name As String) As Proc
    Dim F As Proc
    
    For Each F In C.Functions
        If F.Name = Name Then
            Set FindProc = F
            Exit Function
        End If
    Next
End Function

Private Function FindLocalVar(F As Proc, Name As String) As Variable
    Dim N As Variable
    
    For Each N In F.Variables
        If N.Name = Name Then
            Set FindLocalVar = N
            Exit Function
        End If
    Next
End Function

Private Function FindVar(Code As CodeData, Name As String) As Variable
    If Code.CurProc Is Nothing Or Code.CurProc Is Code.Main Then
        Set FindVar = FindGlobalVar(Code, Name)
        Exit Function
    End If
    
    Set FindVar = FindLocalVar(Code.CurProc, Name)
    
    If FindVar Is Nothing Then
        Set FindVar = FindGlobalVar(Code, Name)
    End If
End Function

Private Function FindGlobalVar(Code As CodeData, Name As String) As Variable
    Dim N As Variable
    
    For Each N In Code.Globals
        If N.Name = Name Or N.Name = "main@" & Name Then
            Set FindGlobalVar = N
            Exit Function
        End If
    Next
End Function

Private Sub MarkUsedFunctions(C As CodeData, N As AstNode)
    Dim P As Proc
    Dim Ch As AstNode
    Dim I As Integer
    
    If N.NodeType = AnBlock Then
        Set P = FindProc(C, N.BlockHead)
        
        If Not (P Is Nothing) Then
            P.IsUsed = True
            
            For I = 4 To P.Node.Children.Count
                Set Ch = P.Node.Children(I)
                MarkUsedFunctions C, Ch
            Next I
        End If
        
        For Each Ch In N.Children
            MarkUsedFunctions C, Ch
        Next
    End If
End Sub

Private Sub FindFunctions(Block As AstNode, C As CodeData)
    Dim N As AstNode
    Dim F As Proc
    Dim Name As String
    
    For Each N In Block.Children
        If N.BlockHead = "proc" Then
            Set F = New Proc
            
            F.Name = N.Children(2).Value
            Set F.Node = N
            
            F.FrameSize = 0
            F.IsUsed = False
            
            If F.Name <> "main" Then
                F.Label = AsmLabelNew()
                
                FindLocals F
                FindArgs F
            End If
            
            C.Functions.Add F
        End If
    Next
End Sub

Private Sub FindArgs(F As Proc)
    Dim N As AstNode
    Dim I As Integer
    Dim V As Variable
    Dim R As Long
    
    R = F.FrameSize
    F.Args = 0
    
    For I = F.Node.Children(3).Children.Count To 1 Step -1
        Set N = F.Node.Children(3).Children(I)
        
        Set V = New Variable
        
        V.Name = N.Value
        V.Size = 2
        V.Ref = R
        
        Set V.Node = N
        
        V.IsLocal = True
        
        F.Variables.Add V
        
        F.Args = F.Args + 1
        
        R = R + 2
    Next I
    
    F.FrameSize = R
End Sub

Private Sub FindLocals(F As Proc)
    Dim N As AstNode
    Dim V As Variable
    Dim R As Long
    
    R = 0
    
    For Each N In F.Node.Children
        If N.BlockHead = "dw" Or N.BlockHead = "db" Then
            Set V = CreateVar(N)
            V.Ref = R
            R = R + V.Size * V.Data.Count
            
            If R Mod 2 = 1 Then
                R = R + 1
            End If
            
            V.IsLocal = True
            
            F.Variables.Add V
        ElseIf N.BlockHead = "let" Then
            Set V = New Variable
            V.IsConst = True
            Set V.Node = N.Children(3)
            V.Name = N.Children(2).Value
            
            F.Variables.Add V
        End If
    Next
    
    F.FrameSize = R
End Sub

Private Function CreateVar(N As AstNode) As Variable
    Dim V As Variable
    Dim I As Integer
    Dim J As Integer
    
    Set V = New Variable
    
    V.Size = 2
    V.IsConst = False
    
    If N.Children(1).Value = "db" Then
        V.Size = 1
    End If
        
    '(db X (dup 10 0))
    If N.Children(3).BlockHead = "dup" Then
        For I = 1 To CInt(N.Children(3).Children(2).Value)
            V.Data.Add CInt(N.Children(3).Children(3).Value)
        Next I
    Else
        For I = 3 To CInt(N.Children.Count) ' (db X 0 1 2 3)
            If IsNumeric(N.Children(I).Value) Then
                V.Data.Add CInt(N.Children(I).Value)
            ElseIf VarType(N.Children(I).Value) = vbString Then ' (db X "string")
                For J = 1 To Len(N.Children(I).Value)
                    V.Data.Add Asc(Mid(N.Children(I).Value, J, 1))
                Next J
            End If
        Next I
    End If
    
    V.Name = N.Children(2).Value
    Set V.Node = N
    
    Set CreateVar = V
End Function

Private Sub FindGlobals(Block As AstNode, C As CodeData, Prefix As String)
    Dim N As AstNode
    Dim Name As String
    Dim V As Variable
    
    For Each N In Block.Children
        If N.BlockHead = "dw" Or N.BlockHead = "db" Then
            Set V = CreateVar(N)
            V.Name = Prefix & V.Name
            
            V.IsLocal = False
            
            C.Globals.Add V
        ElseIf N.BlockHead = "let" Then
            Set V = New Variable
            V.IsConst = True
            Set V.Node = N.Children(3)
            V.Name = Prefix & N.Children(2).Value
            C.Globals.Add V
        End If
    Next
End Sub

Private Function IsAsmKeyword(Kw As String)
    IsAsmKeyword = (Kw = "asm_db") Or _
        (Kw = "asm_dw") Or _
        (Kw = "asm_offset") Or _
        (Kw = "asm_push")
End Function

Private Function GetRegByName(Name As String) As Integer
Select Case Name
Case "ax"
GetRegByName = AsmALAX
Case "bx"
GetRegByName = AsmBLBX
Case "cx"
GetRegByName = AsmCLCX
Case "dx"
GetRegByName = AsmDLDX
End Select
End Function

Private Function Fail(Cond As Boolean, Msg As String, N As AstNode)
    Utils.AssertError Cond, Msg & Chr(10) & "at line #" & N.LineNumber & Chr(10)
End Function


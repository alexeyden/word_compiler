Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Public Type AsmData
    Org As Long
    Labels As New Collection
    Rels As New Collection
    PushList As New Collection
End Type

Public AsmDataGlobal As AsmData

Public Const OPTIMIZE_PUSH = True

Public Const AsmReg = 0
Public Const AsmMem = 1
Public Const AsmConst = 2

Public Const AsmMemNoDisp = 4
Public Const AsmMemBX = 8
Public Const AsmMemBPDI = 16
Public Const AsmMemBP = 32

Public Const AsmNone = 3

Public Const AsmALAX = &H0
Public Const AsmCLCX = &H1
Public Const AsmDLDX = &H2
Public Const AsmBLBX = &H3
Public Const AsmAHSP = &H4
Public Const AsmCHBP = &H5
Public Const AsmDHSI = &H6
Public Const AsmBHDI = &H7

Public Const AsmByte = 0
Public Const AsmWord = 1

Public Const CgMathAdd = 0
Public Const CgMathSub = 1
Public Const CgMathMul = 2
Public Const CgMathIMul = 3
Public Const CgMathDiv = 4
Public Const CgMathIDiv = 5
Public Const CgMathRem = 6
Public Const CgMathIRem = 7

Public Const CgLogNeg = 6
Public Const CgLogOr = 7
Public Const CgLogAnd = 8
Public Const CgLogNot = 9
Public Const CgLogLt = 10
Public Const CgLogGt = 11
Public Const CgLogEq = 12
Public Const CgLogShl = 13
Public Const CgLogShr = 14
Public Const CgLogLe = 15
Public Const CgLogGe = 16
Public Const CgLogNe = 17

Public Const CgTargetStack = 0
Public Const CgTargetGlobal = 1

Public Const CgFromConst = 0
Public Const CgFromStack = 1

Const AopNone = -1

Const AopWWide = 1
Const AopWShort = 0

Const AopDirDest = 2
Const AopDirSrc = 0

Const AopModAddr = 0
Const AopModDisp16 = &H80
Const AopModReg = &HC0

Const AopRmBPDI = 3
Const AopRmBX = 7
Const AopRmBP = 6

Const AopRmAddr = 6

Public Enum AsmOp
    AsmOpDb
    AsmOpDw
    AsmOpJmp
    AsmOpCall
    AsmOpJz
    AsmOpJnz
    AsmOpJl
    AsmOpJg
    AsmOpJge
    AsmOpJle
    AsmOpRet
    AsmOpInt
    AsmOpLea
    AsmOpMov
    AsmOpAdd
    AsmOpSub
    AsmOpXor
    AsmOpCmp
    AsmOpAnd
    AsmOpOr
    AsmOpPush
    AsmOpPop
    AsmOpShl
    AsmOpShr
    AsmOpNot
    AsmOpMul
    AsmOpIMul
    AsmOpDiv
    AsmOpIDiv
    AsmOpNeg
    AsmOpXchg
End Enum

Private Sub OutByte(ByVal B As Byte)
    Put #1, , B
End Sub

Private Sub OutWord(ByVal W As Integer)
    Put #1, , W
End Sub

Public Sub AsmBegin(fname As String)
    AsmDataGlobal.Org = &H100
    Set AsmDataGlobal.Labels = New Collection
    Set AsmDataGlobal.Rels = New Collection
    
    On Error Resume Next
    Kill fname
    On Error GoTo 0
    
    Open fname For Binary Access Read Write As #1
End Sub

Public Sub AsmEnd()
    Dim Pos As Variant
    Dim Label As Integer
    Dim Offset As Integer
    
    For Each Pos In AsmDataGlobal.Rels
        Get #1, Pos + 1, Label
        Label = AsmDataGlobal.Labels(Label) - (Pos + 2)
        Put #1, Pos + 1, Label
    Next
    
    Close #1
End Sub

Public Function AsmLabelNew() As Long
    AsmDataGlobal.Labels.Add 0
    AsmLabelNew = AsmDataGlobal.Labels.Count
End Function

Public Function AsmLabelLand(Label As Long)
    AsmDataGlobal.Labels.Add Loc(1), , , Label
    AsmDataGlobal.Labels.Remove Label
    AsmLabelLand = AsmGetIp()
End Function

Public Function AsmGetIp()
    AsmGetIp = AsmDataGlobal.Org + Loc(1)
End Function

Private Sub AsmLabelRef()
    AsmDataGlobal.Rels.Add Loc(1)
End Sub

Public Sub Asm(op As AsmOp, ByVal S As Integer, ByVal TA As Integer, ByVal A As Integer, Optional TB As Integer = 0, Optional B As Integer = 0)
    Const AopDB = &HFFF0
    Const AopDW = &HFFF1
    
    Const AopJmpShort = &HEB
    Const AopJmpNear = &HE9
    Const AopCall = &HE8
    Const AopJz = &H74
    Const AopJnz = &H75
    Const AopJl = &H7C
    Const AopJg = &H7F
    Const AopJle = &H7E
    Const AopJge = &H7D
    
    Const AopRet = &HC2
    Const AopRetConst = &HC3
    Const AopInt = &H23
    
    Const AopLea = &H8C
    Const AopMovReg = &H88
    Const AopMovMemConst = &HC4
    Const AopXor = &H30
    Const AopAddReg = &H0
    Const AopSubReg = &H28
    Const AopAddSubRegConst = &H81
    Const AopCmpReg = &H38
    Const AopCmpRegConst = &H81
    Const AopAnd = &H20
    Const AopOr = &H8
    
    Const AopRiGroupMath = &HF7
    Const AopRiGroupShift = &HD3
    Const AopRiGroupNone = &H0
    
    Const AopPushRi = &H50
    Const AopPopRi = &H58
    Const AopShlRi = &HE0
    Const AopShrRi = &HD8
    Const AopNotRi = &HD0
    Const AopMulRi = &HE0
    Const AopIMulRi = &HE8
    Const AopDivRi = &HF0
    Const AopIDivRi = &HF8
    Const AopNegRi = &HD8
    Const AopMovConstWordRi = &HB8
    Const AopMovConstByteRi = &HB1
    Const AopXchg = &H86
    
    If op <> AsmOpPop And op <> AsmOpPush Then
        AsmResetPush
    End If
    
    Select Case op
    Case AsmOpDb
        OutByte CByte(A)
    Case AsmOpDw
        OutWord (A)
    Case AsmOpJmp
        If S = AsmWord Then
            OutByte CByte(AopJmpNear)
            AsmLabelRef
            OutWord A
        ElseIf S = AsmByte Then
            OutByte CByte(AopJmpShort)
            OutByte CByte(A)
        End If
    Case AsmOpCall
        OutByte CByte(AopCall)
        AsmLabelRef
        OutWord CInt(A)
    Case AsmOpJz
        OutByte CByte(AopJz)
        OutByte CByte(A)
    Case AsmOpJnz
        OutByte CByte(AopJnz)
        OutByte CByte(A)
    Case AsmOpJl
        OutByte CByte(AopJl)
        OutByte CByte(A)
    Case AsmOpJg
        OutByte CByte(AopJg)
        OutByte CByte(A)
    Case AsmOpJge
        OutByte CByte(AopJge)
        OutByte CByte(A)
    Case AsmOpJle
        OutByte CByte(AopJle)
        OutByte CByte(A)
    Case AsmOpRet
        If TA <> AsmNone Then
            OutByte CByte(AopRet)
            OutWord A
        Else
            OutByte CByte(AopRetConst)
        End If
    Case AsmOpInt
        OutByte CByte(AopInt)
        OutWord A
    Case AsmOpLea
        AsmGeneric AopLea, S, TA, A, TB, B
    Case AsmOpMov
        If TA = AsmReg And TB = AsmConst Then
            OutByte IIf(S = AsmWord, CByte(AopMovConstWordRi), CByte(AopMovConstByteRi)) Or CByte(A)
            AsmWriteValue S, B
        ElseIf (TA And AsmMem) And TB = AsmConst Then
            AsmGeneric AopMovMemConst, S, TA, A, TB, B
        Else
            AsmGeneric AopMovReg, S, TA, A, TB, B
        End If
    Case AsmOpAdd
        Const Aop2Add = 0
        If TB = AsmConst Then
            AsmGeneric AopAddSubRegConst, S, TA, A, TB, B, Aop2Add
        Else
            AsmGeneric AopAddReg, S, TA, A, TB, B
        End If
    Case AsmOpSub
        Const Aop2Sub = 5
        If TB = AsmConst Then
            AsmGeneric AopAddSubRegConst, S, TA, A, TB, B, Aop2Sub
        Else
            AsmGeneric AopSubReg, S, TA, A, TB, B
        End If
    Case AsmOpXor
        AsmGeneric AopXor, S, TA, A, TB, B
    Case AsmOpCmp
        If TB = AsmReg Then
            AsmGeneric AopCmpReg, S, TA, A, TB, B
        ElseIf TB = AsmConst Then
            OutByte CByte(AopCmpRegConst)
            OutByte CByte(&HF8) Or TA
            AsmWriteValue S, B
        End If
    Case AsmOpAnd
        AsmGeneric AopAnd, S, TA, A, TB, B
    Case AsmOpOr
        AsmGeneric AopOr, S, TA, A, TB, B
    Case AsmOpPush
        AsmDataGlobal.PushList.Add A
        OutByte CByte(AopPushRi) Or A
    Case AsmOpPop
        If Not AsmTryRevertPush(A) Then
            AsmResetPush
            OutByte CByte(AopPopRi) Or A
        End If
    Case AsmOpShl
        OutByte CByte(AopRiGroupShift)
        OutByte CByte(AopShlRi) Or A
    Case AsmOpShr
        OutByte CByte(AopRiGroupShift)
        OutByte CByte(AopShrRi) Or A
    Case AsmOpNot
        OutByte CByte(AopRiGroupMath)
        OutByte CByte(AopNotRi) Or A
    Case AsmOpMul
        OutByte CByte(AopRiGroupMath)
        OutByte CByte(AopMulRi) Or A
    Case AsmOpIMul
        OutByte CByte(AopRiGroupMath)
        OutByte CByte(AopIMulRi) Or A
    Case AsmOpDiv
        OutByte CByte(AopRiGroupMath)
        OutByte CByte(AopDivRi) Or A
    Case AsmOpIDiv
        OutByte CByte(AopRiGroupMath)
        OutByte CByte(AopIDivRi) Or A
    Case AsmOpNeg
        OutByte CByte(AopRiGroupMath)
        OutByte CByte(AopNegRi) Or A
    Case AsmOpXchg
        AsmGeneric AopXchg, S, TA, A, TB, B
    End Select
End Sub

Private Sub AsmResetPush()
    While AsmDataGlobal.PushList.Count > 0
        AsmDataGlobal.PushList.Remove (1)
    Wend
End Sub

Private Function AsmTryRevertPush(Reg)
    Dim LastReg As Integer
    Dim I As Integer
    Dim dL As Long
    
    If Not OPTIMIZE_PUSH Then
        AsmTryRevertPush = False
        Exit Function
    End If
    
    If AsmDataGlobal.PushList.Count = 0 Then
        AsmTryRevertPush = False
        Exit Function
    End If
    
    LastReg = AsmDataGlobal.PushList(AsmDataGlobal.PushList.Count)
    
    If LastReg = Reg Then
        Seek #1, Loc(1)
        dL = 1
    Else
        AsmTryRevertPush = False
        Exit Function
    End If
    
    AsmDataGlobal.PushList.Remove (AsmDataGlobal.PushList.Count)
    
    'For I = 1 To AsmDataGlobal.Labels.Count
    '    If AsmDataGlobal.Labels(I) > Loc(1) Then
    '        AsmDataGlobal.Labels(I) = AsmDataGlobal.Labels(I) - dL
    '    End If
    'Next I
    
    AsmTryRevertPush = True
End Function

Private Sub AsmGeneric(op As Integer, S As Integer, TA As Integer, A As Integer, Optional TB As Integer = 0, Optional B As Integer = 0, Optional Op2 As Integer = 0)
    Dim D As Integer, Md As Integer, Reg As Integer, Rm As Integer
    Dim SA As Integer, SB As Integer
    
    SA = -1
    SB = -1
    
    If TA = AsmReg Then
        D = AopDirDest
        Reg = &H8 * A
            
        If TB = AsmReg Then ' %a, %b
            Md = AopModReg
            Rm = B
        ElseIf TB = AsmMem Then ' %a, [b]
            Md = AopModAddr
            Rm = AopRmAddr
            SB = AsmWord
        ElseIf TB And AsmMemBX Then ' %a, [%bx + b]
            Rm = AopRmBX
            Md = IIf(TB And AsmMemNoDisp, AopModAddr, AopModDisp16)
            SB = IIf(TB And AsmMemNoDisp, -1, AsmWord)
        ElseIf TB And AsmMemBPDI Then ' %a, [%bp + %di + b]
            Rm = AopRmBPDI
            Md = IIf(TB And AsmMemNoDisp, AopModAddr, AopModDisp16)
            SB = IIf(TB And AsmMemNoDisp, -1, AsmWord)
        ElseIf TB And AsmMemBP Then ' %a, [%bp + b]
            Rm = AopRmBP
            Md = IIf(TB And AsmMemNoDisp, AopModAddr, AopModDisp16)
            SB = IIf(TB And AsmMemNoDisp, -1, AsmWord)
        ElseIf TB = AsmConst Then ' %a, b
            D = AopDirSrc
            Reg = &H8 * Op2
            Md = AopModReg
            Rm = A
            SB = S
        End If
    ElseIf TA And AsmMem Then
        If TA = AsmMem Then ' [a]
            Rm = AopRmAddr
            Md = AopModAddr
            SA = AsmWord
        ElseIf TA And AsmMemBX Then ' [%bx + a]
            Rm = AopRmBX
            Md = IIf(TA And AsmMemNoDisp, AopModAddr, AopModDisp16)
            SA = IIf(TA And AsmMemNoDisp, -1, AsmWord)
        ElseIf TA And AsmMemBPDI Then ' [%bp + %di + a]
            Rm = AopRmBPDI
            Md = IIf(TA And AsmMemNoDisp, AopModAddr, AopModDisp16)
            SA = IIf(TA And AsmMemNoDisp, -1, AsmWord)
        ElseIf TA And AsmMemBP Then ' [%bp + a]
            Rm = AopRmBP
            Md = IIf(TA And AsmMemNoDisp, AopModAddr, AopModDisp16)
            SA = IIf(TA And AsmMemNoDisp, -1, AsmWord)
        End If
        
        D = AopDirSrc
         
        If TB = AsmReg Then ' [a], %b
            Reg = &H8 * B
        ElseIf TB = AsmConst Then ' [a], b
            D = AopDirDest
            Reg = 0
            SB = S
        End If
    End If
        
    OutByte (CByte(op) Or D Or S)
    OutByte (Md Or Reg Or Rm)
    
    If SA >= 0 Then
        AsmWriteValue SA, A
    End If
    
    If SB >= 0 Then
        AsmWriteValue SB, B
    End If
End Sub

Private Sub AsmWriteValue(S As Integer, V As Integer)
    If S = AsmWord Then
        OutWord V
    Else
        OutByte CByte(V)
    End If
End Sub

' Higher level interface

Public Sub AsmGenMathOp(op As Integer)
    Select Case op
        Case CgMathAdd
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpAdd, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' add ax,bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgMathSub
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpSub, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' sub ax,bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgMathMul
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpMul, AsmWord, AsmReg, AsmBLBX                      ' mul bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgMathIMul
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpIMul, AsmWord, AsmReg, AsmBLBX                     ' imul bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgMathDiv
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpXor, AsmWord, AsmReg, AsmDLDX, AsmReg, AsmDLDX     ' xor dx,dx
            Asm AsmOpDiv, AsmWord, AsmReg, AsmBLBX                      ' div bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgMathIDiv
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpXor, AsmWord, AsmReg, AsmDLDX, AsmReg, AsmDLDX     ' xor dx,dx
            Asm AsmOpNot, AsmWord, AsmReg, AsmDLDX                      ' not dx
            Asm AsmOpIDiv, AsmWord, AsmReg, AsmBLBX                     ' idiv bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgMathRem
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpXor, AsmWord, AsmReg, AsmDLDX, AsmReg, AsmDLDX     ' xor dx,dx
            Asm AsmOpDiv, AsmWord, AsmReg, AsmBLBX                      ' div bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmDLDX                     ' push dx
        Case CgMathIRem
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpXor, AsmWord, AsmReg, AsmDLDX, AsmReg, AsmDLDX     ' xor dx,dx
            Asm AsmOpNot, AsmWord, AsmReg, AsmDLDX                      ' not dx
            Asm AsmOpIDiv, AsmWord, AsmReg, AsmBLBX                     ' idiv bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmDLDX                     ' push dx
    End Select
End Sub

Public Sub AsmGenLogOp(op As Integer, Optional param As Integer = 0)
    Select Case op
        Case CgLogNeg
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpNeg, AsmWord, AsmReg, AsmALAX                      ' neg ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogOr
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpOr, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX      ' or ax, bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogAnd
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpAnd, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' and ax, bx
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogNot
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpNot, AsmWord, AsmReg, AsmALAX                      ' not ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogLt
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpCmp, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' cmp ax,bx
            Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, 1         ' mov ax,1
            Asm AsmOpJl, AsmWord, AsmConst, 2                           ' jl +2
            Asm AsmOpXor, AsmWord, AsmReg, AsmALAX, AsmReg, AsmALAX     ' xor ax,ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogGt
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpCmp, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' cmp ax,bx
            Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, 1         ' mov ax,1
            Asm AsmOpJg, AsmWord, AsmConst, 2                           ' jg +2
            Asm AsmOpXor, AsmWord, AsmReg, AsmALAX, AsmReg, AsmALAX     ' xor ax,ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogEq
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpCmp, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' cmp ax,bx
            Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, 1         ' mov ax,1
            Asm AsmOpJz, AsmWord, AsmConst, 2                           ' je +2
            Asm AsmOpXor, AsmWord, AsmReg, AsmALAX, AsmReg, AsmALAX     ' xor ax,ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogLe
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpCmp, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' cmp ax,bx
            Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, 1         ' mov ax,1
            Asm AsmOpJle, AsmWord, AsmConst, 2                          ' jle +2
            Asm AsmOpXor, AsmWord, AsmReg, AsmALAX, AsmReg, AsmALAX     ' xor ax,ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogGe
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpCmp, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' cmp ax,bx
            Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, 1         ' mov ax,1
            Asm AsmOpJge, AsmWord, AsmConst, 2                          ' jge +2
            Asm AsmOpXor, AsmWord, AsmReg, AsmALAX, AsmReg, AsmALAX     ' xor ax,ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogNe
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX                      ' pop bx
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
            Asm AsmOpCmp, AsmWord, AsmReg, AsmALAX, AsmReg, AsmBLBX     ' cmp ax,bx
            Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, 1         ' mov ax,1
            Asm AsmOpJnz, AsmWord, AsmConst, 2                          ' jne +2
            Asm AsmOpXor, AsmWord, AsmReg, AsmALAX, AsmReg, AsmALAX     ' xor ax,ax
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
        Case CgLogShl
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                          ' pop ax
            Asm AsmOpMov, AsmByte, AsmReg, AsmCLCX, AsmConst, CByte(param)  ' mov cl, <param>
            Asm AsmOpShl, AsmWord, AsmReg, AsmALAX, AsmReg, AsmCLCX         ' shl ax, cl
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                         ' push ax
        Case CgLogShr
            Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                          ' pop ax
            Asm AsmOpMov, AsmByte, AsmReg, AsmCLCX, AsmConst, CByte(param)  ' mov cl, <param>
            Asm AsmOpShr, AsmWord, AsmReg, AsmALAX, AsmReg, AsmCLCX         ' shr ax, cl
            Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                         ' push ax
    End Select
End Sub

Public Sub AsmGenSetValue(Target As Integer, Size As Integer, offsetSrc As Integer, Offset As Integer, valueSrc As Integer, Value As Integer)
    Dim TA As Integer, TB As Integer, A As Integer, B As Integer
    
    If Target = CgTargetGlobal Then
        If offsetSrc = CgFromStack Then ' mov [bx], <value>
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX ' pop bx
            TA = AsmMem Or AsmMemBX Or AsmMemNoDisp
        ElseIf offsetSrc = CgFromConst Then ' mov [offset], <value>
            TA = AsmMem
            A = Offset
        End If
    ElseIf Target = CgTargetStack Then
        If offsetSrc = CgFromStack Then ' mov [bp + di], <value>
            ' NOT USED
            Asm AsmOpPop, AsmWord, AsmReg, AsmBHDI ' pop di
            TA = AsmMem Or AsmMemBPDI Or AsmMemNoDisp
        ElseIf offsetSrc = CgFromConst Then ' mov [bp + offset], <value>
            TA = AsmMem Or AsmMemBP
            A = Offset
        End If
    End If
    
    If valueSrc = CgFromStack Then ' mov <offset>, ax
        Asm AsmOpPop, AsmWord, AsmReg, AsmALAX ' pop ax
        TB = AsmReg
        B = AsmALAX
    ElseIf valueSrc = CgFromConst Then ' mov <offset>, value
        TB = AsmConst
        B = Value
    End If
    
    Asm AsmOpMov, Size, TA, A, TB, B
End Sub

Public Sub AsmGenGetValue(Target As Integer, Size As Integer, offsetSrc As Integer, Offset As Integer)
    Dim TA As Integer, TB As Integer, A As Integer, B As Integer
    
    If Size = AsmByte Then                              ' xor ax,ax
        Asm AsmOpXor, AsmWord, AsmReg, AsmALAX, AsmReg, AsmALAX
    End If
    
    TA = AsmReg
    A = AsmALAX
    
    If Target = CgTargetGlobal Then
        If offsetSrc = CgFromStack Then
            Asm AsmOpPop, AsmWord, AsmReg, AsmBLBX      ' pop bx
            TB = AsmMem Or AsmMemBX Or AsmMemNoDisp     ' mov ax, [bx]
        ElseIf offsetSrc = CgFromConst Then
            TB = AsmMem
            B = Offset                                  ' mov ax, [offset]
        End If
    ElseIf Target = CgTargetStack Then
        If offsetSrc = CgFromStack Then
            Asm AsmOpPop, AsmWord, AsmReg, AsmBHDI      ' pop di
            TB = AsmMem Or AsmMemBPDI Or AsmMemNoDisp   ' mov ax, [bp + di]
        ElseIf offsetSrc = CgFromConst Then
            TB = AsmMem Or AsmMemBP
            B = Offset                                  ' mov ax, [bp + offset]
        End If
    End If
    
    Asm AsmOpMov, Size, TA, A, TB, B
    
    Asm AsmOpPush, AsmWord, TA, A, TB, B                ' push ax
End Sub

Public Sub AsmGenJumpOnZero(ByVal Label As Integer)
    Asm AsmOpPop, AsmWord, AsmReg, AsmALAX                      ' pop ax
    Asm AsmOpCmp, AsmWord, AsmReg, AsmALAX, AsmConst, 0         ' cmp ax,0
    Asm AsmOpJnz, AsmWord, AsmConst, 3                          ' jnz +3
    Asm AsmOpJmp, AsmWord, AsmConst, Label                      ' jmp <label>
End Sub

Public Sub AsmGenCall(Label As Integer, LocalVarsSize As Integer)
    Asm AsmOpSub, AsmWord, AsmReg, AsmAHSP, AsmConst, LocalVarsSize ' sub sp, <total local vars size>
    Asm AsmOpMov, AsmWord, AsmReg, AsmCHBP, AsmReg, AsmAHSP     ' mov bp, sp
    Asm AsmOpCall, AsmWord, AsmConst, Label                     ' call <label>
    Asm AsmOpPop, AsmWord, AsmReg, AsmCHBP                      ' pop bp
End Sub

Public Sub AsmGenRefLocal(Offset As Integer)
    Asm AsmOpLea, AsmWord, AsmMem Or AsmMemBP, Offset, AsmReg, AsmALAX  ' lea ax,[bp + <offset>]
    Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                             ' push ax
End Sub

Public Sub AsmGenPushConst(Value As Integer)
    Asm AsmOpMov, AsmWord, AsmReg, AsmALAX, AsmConst, Value     ' mov ax, <value>
    Asm AsmOpPush, AsmWord, AsmReg, AsmALAX                     ' push ax
End Sub


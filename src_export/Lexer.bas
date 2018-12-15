Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Public Enum TokenType
    TTNone = 0
    TTName = 1
    TTSymbol = 2
    TTChar = 3
    TTString = 4
    TTNumber = 5
    TTLPar = 6
    TTRPar = 7
    TTMinus = 8
    TTPlus = 9
    TTStar = 10
    TTSlash = 11
    TTAnd = 12
    TTLt = 13
    TTGt = 14
    TTEq = 15
    TTOr = 16
    TTNot = 17
    TTComment = 18
    TTShl = 19
    TTShr = 20
    TTHexNumber = 21
    TTPercent = 22
    TTBinNumber = 23
    TTLe = 24
    TTGe = 25
    TTNeq = 26
End Enum

Public Type Token
    Text As String
    Type As TokenType
    PosA As Long
    PosB As Long
End Type

Public Function TokenizeFile(Path As String) As Token()
    Dim Text As String
    Dim Line As String
    
    On Error GoTo ErrorHandler
    
    Open Path For Input As #2
    
    While Not EOF(2)
        Line Input #2, Line
        Text = Text & Line & Chr(13)
    Wend
    
    Close #2
    
    On Error GoTo 0
    
    TokenizeFile = Tokenize(Text, Path)
    
    Exit Function
ErrorHandler:
    AssertError False, "Source file not found: " & Path
End Function

Private Function Tokenize(Source As String, FileName As String) As Token()
Dim State As TokenType
Dim CurToken As Token
Dim Char As String
Dim I As Long
Dim Line As Long
Dim Tokens() As Token

ReDim Tokens(0)

State = TokenType.TTNone
I = 1
Line = 1

For I = 1 To Len(Source)
    Char = Mid(Source, I, 1)
    
    If Asc(Char) = 13 Then
        Line = Line + 1
    End If

    Select Case State
        Case TokenType.TTNone
            If Char = "'" Then          ' Char
            State = TTChar
            CurToken.Text = ""
            CurToken.Type = State
            CurToken.PosA = I
            ElseIf Char = """" Then     ' String
            State = TTString
            CurToken.Text = ""
            CurToken.Type = State
            CurToken.PosA = I
            ElseIf TryName(Char) And Not TryNumber(Char) Then 'Name
            State = TTName
            CurToken.Text = Char
            CurToken.Type = State
            CurToken.PosA = I
            ElseIf TryNumber(Char) Then  'Number
                If TryConst(Source, I, "0x") Then
                    State = TTHexNumber
                    CurToken.Text = ""
                    CurToken.Type = State
                    CurToken.PosA = I + 2
                    I = I + 1
                ElseIf TryConst(Source, I, "0b") Then
                    State = TTBinNumber
                    CurToken.Text = ""
                    CurToken.Type = State
                    CurToken.PosA = I + 2
                    I = I + 1
                Else
                    State = TTNumber
                    CurToken.Text = Char
                    CurToken.Type = State
                    CurToken.PosA = I
                End If
            ElseIf Char = "(" Or Char = "[" Or Char = "{" Then      ' Left Par
            State = TTNone
            CurToken.Text = Char
            CurToken.Type = TTLPar
            CurToken.PosA = I
            CurToken.PosB = Line
            AddToken Tokens, CurToken
            ElseIf Char = ")" Or Char = "]" Or Char = "}" Then      ' Right Par
            State = TTNone
            CurToken.Text = Char
            CurToken.Type = TTRPar
            CurToken.PosA = I
            CurToken.PosB = Line
            AddToken Tokens, CurToken
            ElseIf Char = "+" Then      ' Plus sign
            State = TTNone
            CurToken.Text = Char
            CurToken.Type = TTPlus
            CurToken.PosA = I
            CurToken.PosB = Line
            AddToken Tokens, CurToken
            ElseIf Char = "-" Then      ' Minus sign
                If I < Len(Source) And TryNumber(Mid(Source, I + 1, 1)) Then
                    State = TTNumber
                    CurToken.Text = Char
                    CurToken.Type = State
                    CurToken.PosA = I
                Else
                    State = TTNone
                    CurToken.Text = Char
                    CurToken.Type = TTMinus
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                End If
            ElseIf Char = "*" Then      ' Mul\Deref sign
            State = TTNone
            CurToken.Text = Char
            CurToken.Type = TTStar
            CurToken.PosA = I
            CurToken.PosB = Line
            AddToken Tokens, CurToken
            ElseIf Char = "<" Then      ' Less than
                If TryConst(Source, I, "<<") Then
                    State = TTNone
                    CurToken.Text = "<<"
                    CurToken.Type = TTShl
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                    I = I + 1
                ElseIf TryConst(Source, I, "<=") Then
                    State = TTNone
                    CurToken.Text = "<="
                    CurToken.Type = TTLe
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                    I = I + 1
                Else
                    State = TTNone
                    CurToken.Text = Char
                    CurToken.Type = TTLt
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                End If
            ElseIf Char = ">" Then      ' Greater than
                If TryConst(Source, I, ">>") Then
                    State = TTNone
                    CurToken.Text = ">>"
                    CurToken.Type = TTShr
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                    I = I + 1
                ElseIf TryConst(Source, I, ">=") Then
                    State = TTNone
                    CurToken.Text = ">="
                    CurToken.Type = TTGe
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                    I = I + 1
                Else
                    State = TTNone
                    CurToken.Text = Char
                    CurToken.Type = TTGt
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                End If
            ElseIf Char = "/" Then      ' Div sign
                If TryConst(Source, I, "/*") Then
                    State = TTComment
                    CurToken.Text = Empty
                    CurToken.Type = State
                    CurToken.PosA = I
                    I = I + 1
                Else
                    State = TTNone
                    CurToken.Text = Char
                    CurToken.Type = TTSlash
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                End If
            ElseIf Char = "&" Then      ' And\Ref sign
            State = TTNone
            CurToken.Text = Char
            CurToken.Type = TTAnd
            CurToken.PosA = I
            CurToken.PosB = Line
            AddToken Tokens, CurToken
            ElseIf Char = "|" Then      ' Or sign
            State = TTNone
            CurToken.Text = Char
            CurToken.Type = TTOr
            CurToken.PosA = I
            CurToken.PosB = Line
            AddToken Tokens, CurToken
            ElseIf Char = "=" Then      ' Comparsion\Equal sign
                State = TTNone
                CurToken.Text = "="
                CurToken.Type = TTEq
                CurToken.PosA = I
                CurToken.PosB = Line
                AddToken Tokens, CurToken
            ElseIf Char = "!" Then      ' Not sign
                If TryConst(Source, I, "!=") Then
                    State = TTNone
                    CurToken.Text = "!="
                    CurToken.Type = TTNeq
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                    I = I + 1
                Else
                    State = TTNone
                    CurToken.Text = Char
                    CurToken.Type = TTNot
                    CurToken.PosA = I
                    CurToken.PosB = Line
                    AddToken Tokens, CurToken
                End If
            ElseIf Char = "%" Then
                State = TTNone
                CurToken.Text = Char
                CurToken.Type = TTPercent
                CurToken.PosA = I
                CurToken.PosB = Line
                AddToken Tokens, CurToken
            ElseIf Asc(Char) <= Asc(" ") Then    ' Space or new line
            Else                        ' Unknown token
                Fail "Invalid Token: '" & Char & "' #" & Asc(Char), Line, FileName
            End If
        Case TokenType.TTChar
            If Char = "\" Then
                If TryConst(Source, I, "\n") Then
                    CurToken.Text = CurToken.Text & Chr(&HA)
                ElseIf TryConst(Source, I, "\r") Then
                    CurToken.Text = CurToken.Text & Chr(&HD)
                ElseIf TryConst(Source, I, "\'") Then
                    CurToken.Text = CurToken.Text & "'"
                Else
                    Fail "Unexpected escape char: " & Char, Line, FileName
                End If
                I = I + 1
            ElseIf Char = "'" Then
                State = TTNone
                CurToken.PosB = Line
                AddToken Tokens, CurToken
                CurToken.Text = Empty
                CurToken.Type = State
            Else
                CurToken.Text = CurToken.Text & Char
            End If
        Case TokenType.TTString
            If Char = "\" Then
                If TryConst(Source, I, "\n") Then
                    CurToken.Text = CurToken.Text & Chr(&HA)
                ElseIf TryConst(Source, I, "\r") Then
                    CurToken.Text = CurToken.Text & Chr(&HD)
                ElseIf TryConst(Source, I, "\""""") Then
                    CurToken.Text = CurToken.Text & """"
                Else
                    Fail "Unexpected escape char: " & Char, Line, FileName
                End If
                I = I + 1
            ElseIf Char = """" Then
                State = TTNone
                CurToken.PosB = Line
                AddToken Tokens, CurToken
                CurToken.Text = Empty
                CurToken.Type = State
            Else
                CurToken.Text = CurToken.Text & Char
            End If
        Case TokenType.TTNumber
            If TryNumber(Char) Then
                CurToken.Text = CurToken.Text & Char
            Else
                State = TTNone
                CurToken.PosB = Line
                AddToken Tokens, CurToken
                CurToken.Text = Empty
                CurToken.Type = State
                I = I - 1
            End If
        Case TokenType.TTHexNumber
            If TryNumber(Char) Or Asc(Char) >= Asc("a") And Asc(Char) <= Asc("f") Then
                CurToken.Text = CurToken.Text & Char
            Else
                State = TTNone
                CurToken.PosB = Line
                AddToken Tokens, CurToken
                CurToken.Text = Empty
                CurToken.Type = State
                I = I - 1
            End If
        Case TokenType.TTBinNumber
            If Char = "0" Or Char = "1" Then
                CurToken.Text = CurToken.Text & Char
            Else
                State = TTNone
                CurToken.PosB = Line
                AddToken Tokens, CurToken
                CurToken.Text = Empty
                CurToken.Type = State
                I = I - 1
            End If
        Case TokenType.TTName
            If TryName(Char) Or TryNumber(Char) Then
                CurToken.Text = CurToken.Text & Char
            Else
                State = TTNone
                CurToken.PosB = Line
                AddToken Tokens, CurToken
                CurToken.Text = Empty
                CurToken.Type = State
                I = I - 1
            End If
        Case TokenType.TTComment
            If TryConst(Source, I, "*/") Then
                State = TTNone
                CurToken.PosB = Line
                AddToken Tokens, CurToken
                CurToken.Text = Empty
                CurToken.Type = State
                I = I + 1
            Else
                CurToken.Text = CurToken.Text & Char
            End If
        Case Else
            Fail "Invalid state #" & State, Line, FileName
    End Select
Continue:
Next I

ReDim Preserve Tokens(UBound(Tokens) - 1)

Tokenize = Tokens
End Function

Public Function TokenToString(T As Token) As String
Dim Map(0 To 32)

Map(TTNone) = "none"
Map(TTName) = "name"
Map(TTSymbol) = "symbol"
Map(TTChar) = "char"
Map(TTString) = "string"
Map(TTNumber) = "number"
Map(TTLPar) = "left par"
Map(TTRPar) = "right par"
Map(TTMinus) = "minus"
Map(TTPlus) = "plus"
Map(TTStar) = "star"
Map(TTSlash) = "slash"
Map(TTAnd) = "and"
Map(TTLt) = "less than"
Map(TTGt) = "greater than"
Map(TTEq) = "equal"
Map(TTOr) = "or"
Map(TTComment) = "comment"
Map(TTShl) = "shift left"
Map(TTShr) = "shift right"
Map(TTHexNumber) = "hex num"
Map(TTPercent) = "percent"
Map(TTBinNumber) = "bin num"

TokenToString = Map(T.Type)

End Function

Private Sub AddToken(Tokens() As Token, T As Token)
ReDim Preserve Tokens(UBound(Tokens) + 1)
Tokens(UBound(Tokens) - 1) = T
End Sub

Private Function TryConst(Source As String, ByVal I As Long, SubStr As String) As Boolean
    If Mid(Source, I, Len(SubStr)) <> SubStr Then
        TryConst = False
        Exit Function
    End If
    
    TryConst = True
End Function

Private Function TryName(Char As String) As Boolean
TryName = Asc(Char) >= Asc("a") And Asc(Char) <= Asc("z") Or Asc(Char) >= Asc("A") And Asc(Char) <= Asc("Z") Or Char = "_" Or Char = "$"
End Function

Private Function TryNumber(Char As String) As Boolean
TryNumber = Asc(Char) >= Asc("0") And Asc(Char) <= Asc("9")
End Function

Private Sub Fail(Msg As String, ByVal Line As Integer, FileName As String)
    Utils.AssertError False, "Lexer Error:" & Chr(10) & Msg & Chr(10) & "at line #" & Line & " in " & Chr(10) & FileName
End Sub





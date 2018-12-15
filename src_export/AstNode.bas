Rem Attribute VBA_ModuleType=VBAClassModule
Option VBASupport 1
Option ClassModule
Option Explicit

Public Enum AstNodeType
    AnBlock = 0
    AnNumber = 1
    AnString = 2
    AnChar = 3
    AnSymbol = 4
    AnName = 5
    AnOpAdd = 6
    AnOpSub = 7
    AnOpMul = 8
    AnOpDiv = 9
    AnOpNot = 10
    AnOpAnd = 11
    AnOpOr = 12
    AnOpLs = 13
    AnOpGt = 14
    AnOpEq = 15
    AnOpShl = 16
    AnOpShr = 17
    AnOpRem = 18
    AnOpLe = 19
    AnOpGe = 20
    AnOpNe = 21
End Enum

Public NodeType As AstNodeType
Public Value As Variant
Public Children As New Collection
Public LineNumber As Integer

Public Function BlockHead() As Variant
    If NodeType = AnBlock And Children.Count > 0 Then
        BlockHead = Children(1).Value
    Else
        BlockHead = Empty
    End If
End Function

Public Function NodeTypeName()
Dim Names

Names = Array("block", _
                "number", _
                "string", _
                "char", _
                "symbol", _
                "name", _
                "+", "-", "*", "/", "!", "&", "|", "<", ">", "=", "<<", ">>", "%", "<=", ">=", "!=")
                
NodeTypeName = Names(NodeType)
End Function


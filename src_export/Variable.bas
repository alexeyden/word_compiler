Rem Attribute VBA_ModuleType=VBAClassModule
Option VBASupport 1
Option ClassModule
Option Explicit

Public Size As Integer
Public Name As String
Public Ref As Long
Public Node As AstNode
Public Data As New Collection
Public IsLocal As Boolean
Public IsConst As Boolean


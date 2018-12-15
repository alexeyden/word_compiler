Rem Attribute VBA_ModuleType=VBAClassModule
Option VBASupport 1
Option ClassModule
Option Explicit

Public Name As String
Public Ref As Long
Public Node As AstNode
Public FrameSize As Long
Public Args As Long
Public Label As Long
Public IsMacro As Boolean
Public IsUsed As Boolean

Public Variables As New Collection




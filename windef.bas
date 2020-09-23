Attribute VB_Name = "windef"
Option Explicit

    
    Public POINTAPI As POINTAPI
    Public Type POINTAPI
        X As Long
        Y As Long
    End Type

    Public RECT As RECT
    Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    

    Public Const MAX_PATH = 260

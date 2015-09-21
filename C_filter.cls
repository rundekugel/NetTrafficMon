VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private lastVal As Double
Public filterLen As Integer

Private Sub Class_Initialize()
    filterLen = 1
End Sub

Function filter(ByVal value As Double) As Double
    If filterLen < 1 Then filterLen = 1
    value = value - (value - lastVal) / filterLen
    lastVal = value
    filter = value
End Function

Public Property Get versiontext() As String
    versiontext = "$Id: ls 942 2015-09-03 16:39:08Z BEB $"
End Property
Public Property Get revision() As Integer
    revision = Val(Mid("$Rev:: 942  $", 6, 4))
End Property


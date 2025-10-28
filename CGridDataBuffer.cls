VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGridDataBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GridCell
    text As String
    foreColor As Long
    backColor As Long
End Type

Private m_cellBuffer() As GridCell
Private m_rowCount As Long
Private m_colCount As Long
Private m_initialized As Boolean

Public Sub InitializeBuffer(ByVal rows As Long, ByVal cols As Long)
    If m_initialized And m_rowCount = rows And m_colCount = cols Then
        Exit Sub
    End If
    
    m_rowCount = rows
    m_colCount = cols
    ReDim m_cellBuffer(1 To rows, 1 To cols) As GridCell
    m_initialized = True
End Sub

Public Function UpdateCellIfChanged(ByVal row As Long, ByVal col As Long, _
                                   ByVal newText As String, _
                                   ByVal newForeColor As Long, _
                                   ByVal newBackColor As Long) As Boolean
    If Not m_initialized Then
        UpdateCellIfChanged = True
        Exit Function
    End If
    
    If row < 1 Or row > m_rowCount Or col < 1 Or col > m_colCount Then
        UpdateCellIfChanged = True
        Exit Function
    End If
    
    With m_cellBuffer(row, col)
        If .text <> newText Or .foreColor <> newForeColor Or .backColor <> newBackColor Then
            .text = newText
            .foreColor = newForeColor
            .backColor = newBackColor
            UpdateCellIfChanged = True
        Else
            UpdateCellIfChanged = False
        End If
    End With
End Function

Public Sub Clear()
    Reset
End Sub

' 添加重置方法
Public Sub Reset()
    Erase m_cellBuffer
    m_initialized = False
    m_rowCount = 0
    m_colCount = 0
End Sub

' 添加完全重新初始化方法
Public Sub Reinitialize(ByVal rows As Long, ByVal cols As Long)
    Reset
    InitializeBuffer rows, cols
End Sub

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_initialized
End Property

Public Property Get RowCount() As Long
    RowCount = m_rowCount
End Property

Public Property Get ColCount() As Long
    ColCount = m_colCount
End Property


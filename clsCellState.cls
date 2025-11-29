VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCellState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author/Developer: Vadim Krifuks
'Collaborator: Man Ming Tse
'Last Updated: 2 November 2025

' -- Cell State Class Module: clsCellState --

Option Explicit

Private rng As Range

Private lineStyleLeftEdge As XlLineStyle
Private lineStyleRightEdge As XlLineStyle
Private lineStyleTopEdge As XlLineStyle
Private lineStyleBottomEdge As XlLineStyle

Private colorLeftEdge As Long
Private colorRightEdge As Long
Private colorTopEdge As Long
Private colorBottomEdge As Long

Private weightLeftEdge As XlBorderWeight
Private weightRightEdge As XlBorderWeight
Private weightTopEdge As XlBorderWeight
Private weightBottomEdge As XlBorderWeight

Public Property Set CellRng(cRng As Range)
'Set property for range
    Set rng = cRng
End Property

Public Property Get CellRng() As Range
'Get property for range
    Set CellRng = rng
End Property

Public Sub GetBorders()
'subrotine stores four cell border edges' line style, color, and weight into class variables

    With CellRng.borders(xlEdgeLeft)
        lineStyleLeftEdge = .LineStyle
        colorLeftEdge = .color
        weightLeftEdge = .Weight
    End With

    With CellRng.borders(xlEdgeRight)
        lineStyleRightEdge = .LineStyle
        colorRightEdge = .color
        weightRightEdge = .Weight
    End With

    With CellRng.borders(xlEdgeTop)
        lineStyleTopEdge = .LineStyle
        colorTopEdge = .color
        weightTopEdge = .Weight
    End With

    With CellRng.borders(xlEdgeBottom)
        lineStyleBottomEdge = .LineStyle
        colorBottomEdge = .color
        weightBottomEdge = .Weight
    End With
End Sub

Public Sub SetBorders()
'subrotine sets four cell border edges' line style, color, and weight from class variables

    With CellRng.borders(xlEdgeLeft)
        .LineStyle = lineStyleLeftEdge
        .color = colorLeftEdge
        .Weight = weightLeftEdge
    End With

    With CellRng.borders(xlEdgeRight)
        .LineStyle = lineStyleRightEdge
        .color = colorRightEdge
        .Weight = weightRightEdge
    End With

    With CellRng.borders(xlEdgeTop)
        .LineStyle = lineStyleTopEdge
        .color = colorTopEdge
        .Weight = weightTopEdge
    End With

    With CellRng.borders(xlEdgeBottom)
        .LineStyle = lineStyleBottomEdge
        .color = colorBottomEdge
        .Weight = weightBottomEdge
    End With
End Sub

Option Explicit

Sub VirticalFill()
'
  Dim LastRow, CurRow, OriginRow As Long
  Dim OriginColumn As Long
  Dim Origin As Range
  Dim CurValue
    
  Set Origin = ActiveCell
  OriginRow = Origin.Row
  OriginColumn = Origin.Column
  CurRow = OriginRow
  LastRow = ActiveCell.SpecialCells(xlLastCell).Row
  CurValue = Origin.Value
  
  CurRow = CurRow + 1
  
  While CurRow <= LastRow
    Cells(CurRow, OriginColumn).Activate
    If Cells(CurRow, OriginColumn).Value = "" Then
      Cells(CurRow, OriginColumn).Value = CurValue
    Else
      CurValue = Cells(CurRow, OriginColumn)
    End If
    CurRow = CurRow + 1
  Wend
  
MsgBox "done"

End Sub



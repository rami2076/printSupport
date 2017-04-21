Attribute VB_Name = "ExecutionApp"
Option Explicit
'this sub prosedure convert to function
Sub run()
  Dim dateUtility As DateUtil
  Dim item As Variant
  Dim dateCol As New Collection
  Set dateUtility = New DateUtil
  'valid file
  'open file
  With dateUtility
    If .isValid = True Then
     For Each item In .createList
       Debug.Print item
       'loop file
       'input item
       'printfile
       'next
     Next item
    End If
  End With
  'file close
End Sub

Sub lll()
Debug.Print IsDate(ThisWorkbook.Worksheets(1).Range("A2").Value)
End Sub

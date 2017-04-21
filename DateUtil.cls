VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private biginDateRange_ As Range
Private endDateRange_ As Range
'Private dateCollection_ As Collection
'Private dateFormat As String

Private Sub Class_initialize()
Dim sheetName As String
sheetName = ThisWorkbook.Worksheets(1).Name '=""
Set biginDateRange_ = ThisWorkbook.Worksheets(sheetName).Range("A2")
Set endDateRange_ = ThisWorkbook.Worksheets(sheetName).Range("B2")
'Set dateCollection_ = New Collection
'TODO::add dateFromat initialize
End Sub

'
'Create date list.reference for biginDateRange_ and endDateRange_.
'@return collection the collection between biginDate and endDate.
Function createList() As Collection
  Dim biginDate As Date
  Dim diff As Long
  Dim countDay As Long
  Dim tempList As New Collection
  biginDate = CDate(biginDateRange_.Value)
  diff = DateDiffer()
  For countDay = 0 To diff Step 1
    tempList.Add DateAdd("d", Round((CDbl(countDay))), biginDate)
  Next
  Set createList = tempList
End Function

'
'Validate the user input in the cell
'@return  true if  the input is valid
Function isValid() As Boolean
  Dim errorMsg As String
  errorMsg = ""
  'http://note.phyllo.net/?eid=1106086
  If Len(biginDateRange_.Value) = 0 Or Len(endDateRange_.Value) = 0 Then
   errorMsg = errorMsg & "データが入力されていません。ヒント《A2,B2が入力セル》"
  ElseIf Not IsDate(CStr(biginDateRange_.Value)) Or Not IsDate(CStr(endDateRange_.Value)) Then
   errorMsg = errorMsg & "日付の入力形式を正しくしてください。ヒント《yyyy/mm/dd》"
  ElseIf DateDiffer < 0 Then
   errorMsg = errorMsg & vbCrLf & "開始日と終了日の日時を反対にしてください。" & vbCrLf & vbCrLf & _
                                  "ヒント《開始日:(yyyy/mm/dd)》《終了日:(yyyy/mm/dd)+自然数》"
  End If
  
  If Len(errorMsg) = 0 Then
    isValid = True
    Exit Function
  Else
    'Show  error message.
    MsgBox prompt:=errorMsg, _
           Buttons:=vbCritical, _
           Title:="日付の不正な入力"
    isValid = False
  End If
End Function


'Date diffable
'@return  Long the Long is positive integer and negative integer.
Private Function DateDiffer() As Long
  Dim biginDate As Date
  Dim endDate As Date
  biginDate = CDate(biginDateRange_.Value)
  endDate = CDate(endDateRange_.Value)
  DateDiffer = Datediff("d", biginDate, endDate)
End Function

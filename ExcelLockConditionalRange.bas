Option Explicit
Option Base 1

'======= Settings =========
Const applyToRange = "$1:$1048576"
  'Must be entered absolute, similar to the formula entered in Rules Manager window.
Const changeSheetTabColor = True
Const promptRemoveDuplicates = False
'==========================

'0: Not set, receive from Worksheet
'1: Locked
'2: Not locked
Dim locked As Byte

Const sheetCondStateName = "SheetsConditionalLockState"
Const sheetCondStateNamePass = "3414"

Dim prevConditionalsCount As Integer

' stackoverflow.com/questions/6688131
Function sheetExist(sSheet As String) As Boolean
    On Error Resume Next
    sheetExist = (ThisWorkbook.Sheets(sSheet).Index > 0)
End Function

Private Sub updateSheetLockState()
    Dim wsh As Worksheet

    If locked = 0 Then 'Receive from Worksheet
        If sheetExist(sheetCondStateName) Then
            Dim match As Range, newRow As Integer
            Set wsh = ThisWorkbook.Sheets(sheetCondStateName)
            
            'Unlike Me.Name, Me.CodeName doesn't change by renaming
            Set match = wsh.UsedRange.Columns(1).Find(Me.CodeName)
            If match Is Nothing Then 'Add new row
                wsh.Unprotect (sheetCondStateNamePass)
                newRow = wsh.UsedRange.Row + wsh.UsedRange.Rows.Count
                wsh.Cells(newRow, 1) = Me.CodeName
                wsh.Cells(newRow, 2) = 2 'Default - Not locked
                wsh.Protect (sheetCondStateNamePass)
            Else
                locked = match.Cells(1, 2)
            End If
        Else
            Set wsh = ThisWorkbook.Sheets.Add
            wsh.Visible = xlSheetVeryHidden
            wsh.Name = sheetCondStateName
            wsh.Cells(1, 1) = Me.CodeName
            wsh.Cells(1, 2) = 2 'Default - Not locked
            locked = wsh.Cells(1, 2)
            wsh.Protect (sheetCondStateNamePass)
        End If
    End If
End Sub

Private Sub setSheetLockState(s As Byte)
    Dim wsh As Worksheet, match As Range
    Set wsh = ThisWorkbook.Sheets(sheetCondStateName)
    Set match = wsh.UsedRange.Columns(1).Find(Me.CodeName)
    
    locked = s
    wsh.Unprotect (sheetCondStateNamePass)
    match.Cells(1, 2) = s
    wsh.Protect (sheetCondStateNamePass)
End Sub

Sub Conditional_ToggleLock()
    updateSheetLockState

    If locked = 1 Then
        setSheetLockState (2)
        checkSheet 'To update sheet's color
        MsgBox "Unlocked.", vbInformation
    Else
        setSheetLockState (1)
        checkSheet True
        MsgBox "Locked.", vbInformation
    End If
End Sub

Private Sub checkSheet(Optional hardCheck As Boolean = False)
    With Me.Cells
        If locked = 1 Then
            If changeSheetTabColor Then
                If Me.Tab.Color <> RGB(0, 255, 0) Then
                    Me.Tab.Color = RGB(0, 255, 0)
                End If
            End If
            
            If .FormatConditions.Count > 0 Then
                Dim i As Integer
                
                '=== Set range of emaining rules to defined one
                For i = 1 To .FormatConditions.Count
                    If .FormatConditions(i).AppliesTo.Address <> applyToRange Then
                        .FormatConditions(i).ModifyAppliesToRange (ActiveSheet.Range(applyToRange))
                    End If
                Next
                
                '=== Remove duplicate rules. To check formulas correctly, ...
                '    ...their addresses became identical in previous step. ...
                '    ...The other way is check them relative to their first cell's position.
                
                '    It's not probable to remove and create a rule manually at the same time...
                '    ...(so count would be the same) and have a new rule duplicate of a previous rule!...
                '    ...Eventually, the existence of duplicate rules is not a bug.
                If prevConditionalsCount = 0 Or .FormatConditions.Count <> prevConditionalsCount Or _
                        hardCheck Then
                    Dim j As Integer, totalDupes As Integer
                    
                    totalDupes = 0
                    'For loop doesn't refresh its condition.
                    i = 1
                    Do While i < .FormatConditions.Count
                        j = i + 1
                        Do While j <= .FormatConditions.Count
                            If cmpConditions(.FormatConditions(i), .FormatConditions(j)) Then
                                totalDupes = totalDupes + 1
                                .FormatConditions(j).Delete
                                'Now j will target to new item and shouldn't be increased
                            Else
                                j = j + 1
                            End If
                        Loop
                        
                        i = i + 1
                    Loop
                    prevConditionalsCount = .FormatConditions.Count
                    
                    If promptRemoveDuplicates And totalDupes > 0 Then
                        MsgBox totalDupes & " duplicate conditions were removed.", vbInformation
                    End If
                End If
            End If
        Else
            If changeSheetTabColor Then
                If Me.Tab.Color <> vbRed Then
                    Me.Tab.Color = vbRed
                End If
            End If
        End If
    End With
End Sub

Private Function cmpConditions(ByRef c1 As Object, ByRef c2 As Object)
    Dim areSame As Boolean
    areSame = False
    
    'It's enough to check formulas here:
    If _
           (c1.Type = xlTextString Or c1.Type = xlExpression Or c1.Type = xlTimePeriod Or _
            c1.Type = xlErrorsCondition Or c1.Type = xlNoErrorsCondition Or _
            c1.Type = xlBlanksCondition Or c1.Type = xlNoBlanksCondition) _
            And _
           (c2.Type = xlTextString Or c2.Type = xlExpression Or c2.Type = xlTimePeriod Or _
            c2.Type = xlErrorsCondition Or c2.Type = xlNoErrorsCondition Or _
            c2.Type = xlBlanksCondition Or c2.Type = xlNoBlanksCondition) Then
        If c1.Formula1 = c2.Formula1 Then
            areSame = True
        End If


    'Otherwise, types must be identical:
    ElseIf c1.Type = c2.Type Then
        If c1.StopIfTrue = c2.StopIfTrue Then
            If c1.PTCondition = c2.PTCondition Then 'On PivotTable
                Select Case c2.Type
                
                Case xlCellValue
                    If c1.Operator = c2.Operator And c1.Formula1 = c2.Formula1 Then
                        If c1.Operator = xlBetween Or c1.Operator = xlNotBetween Then
                            If c1.Formula2 = c2.Formula2 Then
                                areSame = True
                            End If
                        Else
                            areSame = True
                        End If
                    End If
                
                Case xlUniqueValues
                    If c1.DupeUnique = c2.DupeUnique Then
                        areSame = True
                    End If
                    
                Case xlAboveAverageCondition
                    If c1.AboveBelow = c2.AboveBelow And c1.CalcFor = c2.CalcFor And _
                            c1.NumStdDev = c2.NumStdDev Then
                        areSame = True
                    End If
                
                Case xlTop10
                    If c1.TopBottom = c2.TopBottom And c1.CalcFor = c2.CalcFor And _
                            c1.Rank = c2.Rank And c1.Percent = c2.Percent Then
                        areSame = True
                    End If
                
                Case xlColorScale
                    If c1.ColorScaleCriteria.Count = c2.ColorScaleCriteria.Count Then
                        Dim hasInequality As Boolean, i As Integer
                        hasInequality = False
                        For i = 1 To c1.ColorScaleCriteria.Count
                            If c1.ColorScaleCriteria(i).Type <> c2.ColorScaleCriteria(i).Type Or _
                                    c1.ColorScaleCriteria(i).Value <> c2.ColorScaleCriteria(i).Value Then
                                hasInequality = True
                                Exit For
                            End If
                        Next
                        
                        If Not hasInequality Then
                            areSame = True
                        End If
                    End If
                    
                Case xlIconSets
                    If c1.ReverseOrder = c2.ReverseOrder Then
                        'Item 1 is same all the time, except its icon
                        If c1.IconCriteria(2).Operator & c1.IconCriteria(2).Type & c1.IconCriteria(2).Value = _
                                c2.IconCriteria(2).Operator & c2.IconCriteria(2).Type & c2.IconCriteria(2).Value And _
                                c1.IconCriteria(3).Operator & c1.IconCriteria(3).Type & c1.IconCriteria(3).Value = _
                                c2.IconCriteria(3).Operator & c2.IconCriteria(3).Type & c2.IconCriteria(3).Value Then
                            areSame = True
                        End If
                    End If
                    
                Case xlDatabar
                    If c1.MaxPoint.Type = c2.MaxPoint.Type And c1.MinPoint.Type = c2.MinPoint.Type Then
                        If c1.MaxPoint.Value = c2.MaxPoint.Value And c1.MinPoint.Value = c2.MinPoint.Value Then
                            areSame = True
                        End If
                    End If
                
                Case Else
                    'Error unknown type!
        
                End Select
                
            End If
        End If
    End If

    cmpConditions = areSame
End Function

Sub Conditional_Refresh()
    cRefresh True
End Sub

Private Sub cRefresh(Optional hardRefresh As Boolean = False)
    updateSheetLockState
    
    If hardRefresh And locked = 2 Then
        If MsgBox("Sheet is not locked! Lock sheet and refresh?", vbQuestion + vbYesNo) = vbYes Then
            Conditional_ToggleLock
        End If
        Exit Sub
    End If
    
    checkSheet hardRefresh
End Sub

Private Sub Worksheet_Activate()
    cRefresh
End Sub

Private Sub Worksheet_BeforeDelete()
    updateSheetLockState
    
    Dim wsh As Worksheet, match As Range
    Set wsh = ThisWorkbook.Sheets(sheetCondStateName)
    Set match = wsh.UsedRange.Columns(1).Find(Me.CodeName)
    
    wsh.Unprotect (sheetCondStateNamePass)
    match = ""
    match.Cells(1, 2) = ""
    wsh.Protect (sheetCondStateNamePass)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    cRefresh
End Sub

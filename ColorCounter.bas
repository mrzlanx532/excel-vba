Attribute VB_Name = "Module1"
Sub ПОДСЧЕТ_ЦВЕТОВ()
Attribute ПОДСЧЕТ_ЦВЕТОВ.VB_ProcData.VB_Invoke_Func = "e\n14"

    Dim GC As Long
    Dim RC As Long
    Dim BC As Long
    
    Dim Pstv As Long
    Dim Ngtv As Long
    Dim Ntrl As Long
    
    Dim TCount As Long
    
    GC = 12449213
    RC = 12500733
    BC = 16766902
    
    For Each CRow In Selection.Rows
        
        Pstv = 0
        Ngtv = 0
        Ntrl = 0
        
        TCount = CRow.Cells.Count
    
        For Each CCell In CRow.Cells
            If CCell.Interior.Color = GC Then
               Pstv = Pstv + 1
            ElseIf CCell.Interior.Color = RC Then
               Ngtv = Ngtv + 1
            ElseIf CCell.Interior.Color = BC Then
               Ntrl = Ntrl + 1
            End If
        Next
        
        If IsNumeric(CRow.Cells(TCount)) Or CRow.Cells(TCount) = "" Then
            CRow.Cells(TCount).Value = Ntrl
        Else
            MsgBox "Неправильно выделена область. Отмена"
            Exit Sub
        End If
        
        If IsNumeric(CRow.Cells(TCount - 1)) Or CRow.Cells(TCount - 1) = "" Then
            CRow.Cells(TCount - 1).Value = Pstv
        Else
            MsgBox "Неправильно выделена область. Отмена"
            Exit Sub
        End If
        
        If IsNumeric(CRow.Cells(TCount - 2)) Or CRow.Cells(TCount - 2) = "" Then
            CRow.Cells(TCount - 2).Value = Ngtv
        Else
            MsgBox "Неправильно выделена область. Отмена"
            Exit Sub
        End If
        
    Next
    
    MsgBox "Динамика подсчитана!"
    
End Sub


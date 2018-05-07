Attribute VB_Name = "Module1"
Function ConvertNumber(cell As Range) As String
    Dim addIn As COMAddIn
    Dim automationObject As Object
    Set addIn = Application.COMAddIns("ExcelImportData")
    Set automationObject = addIn.Object
    ConvertNumber = automationObject.ImportData(cell)
End Function
Sub AddToCellMenu()
    Dim ContextMenu As CommandBar

    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, before:=1

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "TranslateMacro"
        .FaceId = 59
        .Caption = "Translate"
        .Tag = "My_Cell_Control_Tag"
    End With

    ' Add a custom submenu with three buttons.
    

    ' Add a separator to the Cell context menu.
    ContextMenu.Controls(4).BeginGroup = True
End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub

Sub TranslateMacro()
    Dim CaseRange As Range
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub
    
    For Each cell In CaseRange.Cells
        Call CallTranslate(cell)
    Next cell
    
End Sub

Sub CallTranslate(cell As Range)
    Dim addIn As COMAddIn
    Dim automationObject As Object
    Set addIn = Application.COMAddIns("ExcelImportData")
    Set automationObject = addIn.Object
    Call automationObject.Translate(cell)
End Sub


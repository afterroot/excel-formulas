Attribute VB_Name = "slicerUitls"
Sub slicerItemSelect(slicerName As String, ByVal selection As String)
    ThisWorkbook.SlicerCaches(slicerName).SlicerItems(selection).Selected = True
    Dim item As SlicerItem
    For Each item In ThisWorkbook.SlicerCaches(slicerName).SlicerItems
        If Not item.Caption = selection Then
            item.Selected = False
        End If
    Next item
End Sub


' Snippets
Sub HeavyCalculationExample()
    ' Disable Automatic Calculation to Reduce CPU Usage for resource heavy procedures
    Application.Calculation = xlCalculationManual

    ' [WRITE CODE HERE]

    ' Enable Automatic Calculation after slicer operation finish
    Application.Calculation = xlCalculationAutomatic
End Sub
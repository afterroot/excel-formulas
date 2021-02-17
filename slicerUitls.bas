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

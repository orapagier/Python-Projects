Attribute VB_Name = "Module1"
Sub RemoveShadows()
    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                ' Check if the shape has a shadow and remove it
                If shp.Shadow.Visible = True Then
                    shp.Shadow.Visible = False
                End If
            End If
        Next shp
    Next sld
End Sub


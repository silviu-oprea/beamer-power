Sub AddProgressBar()
    On Error Resume Next
        With ActivePresentation
              sHeight = .PageSetup.SlideHeight - 12
              n = 0
              j = 0
              For i = 1 To .Slides.Count - 1
                If .Slides(i).SlideShowTransition.Hidden Then j = j + 1
              Next i:
              For i = 2 To .Slides.Count - 1
                .Slides(i).Shapes("progressBar").Delete
                If .Slides(i).SlideShowTransition.Hidden = msoFalse Then
                  Set slider = .Slides(i).Shapes.AddShape(msoShapeRectangle, 0, 0, (i - n - 1) * .PageSetup.SlideWidth / (.Slides.Count - j - 2), 12)
                  With slider
                      .Fill.ForeColor.RGB = ActivePresentation.SlideMaster.ColorScheme.Colors(ppFill).RGB
                      .Name = "progressBar"
                  End With
                Else
                   n = n + 1
                End If
              Next i:
        End With
End Sub

Sub RemoveProgressBar()
    On Error Resume Next
        With ActivePresentation
              For i = 1 To .Slides.Count
              .Slides(i).Shapes("progressBar").Delete
              Next i:
        End With
End Sub
Sub ContextDotsTop()
    On Error Resume Next
            With ActivePresentation
            
                SectionCount = .SectionProperties.Count
                
                For X = 1 To .Slides.Count
        
                    For i = 1 To 20
                        .Slides(X).Shapes("Bullet").Delete
                    Next i
                    .Slides(X).Shapes("Background").Delete
                    .Slides(X).Shapes("SectionTitleBox").Delete
                            
                    Set bg = .Slides(X).Shapes.AddShape(msoShapeRectangle, 0, 0, .PageSetup.SlideWidth, 25)
                    bg.Name = "Background"
                    ' Change the rectangle's colours here
                    bg.Fill.ForeColor.RGB = vbBlack
                    bg.Line.ForeColor.RGB = vbBlack
                    
                    ' Change the bullets' size, shape and spacing here
                    BulletSize = 6
                    BulletShape = msoShapeOval
                    BulletSpacing = 2
                    
                    Offset = 20
                    SlideNumber = 1
                    For Y = 2 To SectionCount - 1
                    
                        If Y <> 1 Then
                            Offset = (Y - 1) * .PageSetup.SlideWidth / (SectionCount)
                        End If
                        TextboxOffset = Offset - 7.5
                        
                        SectionSlidesCount = .SectionProperties.SlidesCount(Y)
                        sectionTitle = .SectionProperties.Name(Y)
                        Set textbox = .Slides(X).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                TextboxOffset, -2, 200, 50)
                        textbox.TextFrame.TextRange.Text = sectionTitle
                        textbox.Name = "SectionTitleBox"
                        ' Change the font colour, size and type here
                        textbox.TextFrame.TextRange.Font.Color = vbWhite
                        textbox.TextFrame.TextRange.Font.Size = 11
                        textbox.TextFrame.TextRange.Font.Name = "Calibri Light"
                        
                        For Z = 1 To SectionSlidesCount
                            Set Bullet = .Slides(X).Shapes.AddShape(BulletShape, _
                                Offset + (Z - 1) * (BulletSpacing + BulletSize), 16, BulletSize, BulletSize)
                            Bullet.Name = "Bullet"
                            Bullet.Line.Weight = 1
                            ' Change the bullets' fill and line colour here (case: Other slide)
                            Bullet.Fill.ForeColor.RGB = vbBlack
                            Bullet.Line.ForeColor.RGB = vbWhite
                            
                            
                            If X = SlideNumber + 1 Then
                                textbox.TextFrame.TextRange.Font.Bold = True
                                ' Change the bullets' fill and line colour here (case: active slide)
                                Bullet.Fill.ForeColor.RGB = vbWhite
                            End If
                            
                            SlideNumber = SlideNumber + 1
                        Next Z:
                    Next Y:
                    
                Next X:
            End With
End Sub
Sub ContextDotsBottom()
    On Error Resume Next
            With ActivePresentation
            
                SectionCount = .SectionProperties.Count
                
                For X = 1 To .Slides.Count
                    .Slides(X).Shapes("Background").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("SectionTitleBox").Delete
                    
                    TopOffset = .PageSetup.SlideHeight - 25
                    Set bg = .Slides(X).Shapes.AddShape(msoShapeRectangle, 0, TopOffset, .PageSetup.SlideWidth, 25)
                    bg.Name = "Background"
                    ' Change the rectangle's colours here
                    bg.Fill.ForeColor.RGB = vbBlack
                    bg.Line.ForeColor.RGB = vbBlack
                    
                    ' Change the bullets' size, shape and spacing here
                    BulletSize = 6
                    BulletShape = msoShapeOval
                    BulletSpacing = 2
                    
                    Offset = 20
                    SlideNumber = 1
                    For Y = 1 To SectionCount
                    
                        If Y <> 1 Then
                            Offset = (Y - 1) * .PageSetup.SlideWidth / (SectionCount)
                        End If
                        TextboxOffset = Offset - 7.5
                        
                        SectionSlidesCount = .SectionProperties.SlidesCount(Y)
                        sectionTitle = .SectionProperties.Name(Y)
                        Set textbox = .Slides(X).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                TextboxOffset, TopOffset - 2, 200, 50)
                        textbox.TextFrame.TextRange.Text = sectionTitle
                        textbox.Name = "SectionTitleBox"
                        ' Change the font colour, size and type here
                        textbox.TextFrame.TextRange.Font.Color = vbWhite
                        textbox.TextFrame.TextRange.Font.Size = 11
                        textbox.TextFrame.TextRange.Font.Name = "Calibri Light"
                        
                        For Z = 1 To SectionSlidesCount
                            Set Bullet = .Slides(X).Shapes.AddShape(BulletShape, _
                                Offset + (Z - 1) * (BulletSpacing + BulletSize), TopOffset + 16, BulletSize, BulletSize)
                            Bullet.Name = "Bullet"
                            Bullet.Line.Weight = 1
                            ' Change the bullets' fill and line colour here (case: Other slide)
                            Bullet.Fill.ForeColor.RGB = vbBlack
                            Bullet.Line.ForeColor.RGB = vbWhite
                            
                            
                            If X = SlideNumber Then
                                textbox.TextFrame.TextRange.Font.Bold = True
                                ' Change the bullets' fill and line colour here (case: active slide)
                                Bullet.Fill.ForeColor.RGB = vbWhite
                            End If
                            
                            SlideNumber = SlideNumber + 1
                        Next Z:
                    Next Y:
                    
                Next X:
            End With
End Sub
Sub RemoveContextDots()
    On Error Resume Next
    
            With ActivePresentation
            
                SectionCount = .SectionProperties.Count

                For X = 1 To .Slides.Count
                
                    For Each Shape In .Slides(X).Shapes
                        If InStr(1, Shape.Name, "Bullet") <> 0 Then Shape.Delete
                    Next Shape
                    .Slides(X).Shapes("Background").Delete
                    
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    
                    .Slides(X).Shapes("SectionTitleBox").Delete
                    
                Next X:
                
            End With
End Sub

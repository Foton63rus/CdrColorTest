Sub TestColor()
    Dim SizeX As Integer
    Dim XRec, YRec As Double
    'Set MM = 0.5374866
    Dim XNum As Integer
    XNum = 10
    Dim YNum As Integer
    YNum = 10
    Dim cC, cM, cY, cK As Integer
    XRec = 5.374866 * (XNum + 1) / 10
    YRec = 2 * (YNum + 1) / 10
    YTamplateOffset = YRec * YNum + 0.2
    
    Dim s1 As Shape
    Dim s2 As Shape
    XNum = XNum - 1
    YNum = YNum - 1
    
    For X = 0 To XNum
        For Y = 0 To YNum
            Set s1 = ActiveLayer.CreateRectangle(XRec * X, YRec * Y + 0.2, XRec * (X + 1), YRec * (Y + 1))
            s1.Fill.ApplyNoFill
            cC = Math.Round(100 * (XNum - X) / XNum * (YNum - Y) / YNum)
            cM = Math.Round(100 * X / XNum * (YNum - Y) / YNum)
            cY = Math.Round(100 * (XNum - X) / XNum * Y / YNum)
            cK = Math.Round(100 * X / XNum * Y / YNum)
            s1.Outline.SetProperties 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), False, False, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100
            s1.Fill.UniformColor.CMYKAssign cC, cM, cY, cK
            
            Set s2 = ActiveLayer.CreateArtisticText(XRec * X, YRec * Y + 0.1, "C=" & Str(cC) & "  M=" & Str(cM) & "  Y=" & Str(cY) & "  K=" & Str(cK) & "  ", cdrRussian, cdrCharSetRussian, "Arial", 8, cdrTrue, cdrFalse, cdrNoFontLine, cdrLeftAlignment)
            s2.Outline.SetProperties 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), False, False, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100
        Next
    Next
    
    For X = 0 To XNum
        For Y = 0 To YNum
            Set s1 = ActiveLayer.CreateRectangle(XRec * X, YRec * Y + 0.2 + YTamplateOffset, XRec * (X + 1), YRec * (Y + 1) + YTamplateOffset)
            s1.Fill.ApplyNoFill
            cC = Math.Round(100 * (XNum - X) / XNum * (YNum - Y) / YNum)
            cM = Math.Round(100 * X / XNum * (YNum - Y) / YNum)
            cY = 0
            cK = 0
            s1.Outline.SetProperties 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), False, False, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100
            s1.Fill.UniformColor.CMYKAssign cC, cM, cY, cK
            s1.ApplyEffectHSL 360 * (XNum - X), 1, 1
            
            Set s2 = ActiveLayer.CreateArtisticText(XRec * X, YRec * Y + 0.1 + YTamplateOffset, "C=" & Str(cC) & "  M=" & Str(cM) & "  Y=" & Str(cY) & "  K=" & Str(cK) & "  ", cdrRussian, cdrCharSetRussian, "Arial", 8, cdrTrue, cdrFalse, cdrNoFontLine, cdrLeftAlignment)
            s2.Outline.SetProperties 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), False, False, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100
        Next
    Next

End Sub

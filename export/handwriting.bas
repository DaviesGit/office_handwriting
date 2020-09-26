Attribute VB_Name = "handwriting"
Sub handwriting()
'
' handwriting Macro
'
'

    Dim font_count As Long
    font_count = 11
    
    Dim font_config() As FontConfig
    ReDim font_config(font_count)
    Dim total_probability As Double
    
    
    
    For i = 0 To (font_count - 1)
        Set font_config(i) = New FontConfig
    Next i
    
    
    font_config(0).InitializeWithValues "SJNMDJW", 18, -2, 20
    font_config(1).InitializeWithValues "MEIYUJW", 16, 0, 25
    font_config(2).InitializeWithValues "FZJingLeiS-R-GB", 14, 2, 15
    font_config(3).InitializeWithValues "AR GbxingkaiGB Extrabold", 14, 2, 20
    font_config(4).InitializeWithValues "HY JingBRJ", 17, -1, 15
    font_config(5).InitializeWithValues "?????", 15, 0, 0
    font_config(6).InitializeWithValues "BoLeZhuSunti", 15, 0, 15
    font_config(7).InitializeWithValues "DFPHanziPenW3-GB", 15, 1.2, 8
    font_config(8).InitializeWithValues "BoLeYaYati", 16, 0, 0
    font_config(9).InitializeWithValues "HYPPTiJ", 15, 1.2, 0
    font_config(10).InitializeWithValues "?????", 15, 0, 0
    
    
    total_probability = 0
    For i = 0 To (font_count - 1)
        total_probability = total_probability + font_config(i).probability
    Next i
    
    VBA.Randomize
    Dim random As Integer

    

    Dim last_font_ratio As Double
    Dim font_size As Double
    
    last_font_ratio = 0.2
    For Each R_Character In ActiveDocument.Characters
        random = Int(VBA.Rnd * total_probability)
        
        Dim current_font As FontConfig
        Dim current_count As Double
        current_count = 0
        For i = 0 To (font_count - 1)
            current_count = current_count + font_config(i).probability
            If random < current_count Then
                Set current_font = font_config(i)
                Exit For
            End If
        Next i
    
        font_ratio = last_font_ratio + (0.1 * VBA.Rnd - 0.05)
        If font_ratio > 0.25 Then
        font_ratio = 0.25
        End If
        If font_ratio < 0.15 Then
        font_ratio = 0.15
        End If
        last_font_ratio = font_ratio
        font_size = current_font.size * (1 + last_font_ratio)
        
        R_Character.Font.name = current_font.name
        R_Character.Font.size = font_size
        R_Character.Font.Position = -(VBA.Rnd * 0.2 + 0.1) * (font_size - 15)
        R_Character.Font.Spacing = current_font.expanded + VBA.Rnd * 2 + -2
    Next
        Application.ScreenUpdating = True



End Sub

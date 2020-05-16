Attribute VB_Name = "Module1"

Dim preShape As Shape
Dim cntEnter As Integer



'// ******************************************************** //
'// Tab�L�[�ɑΉ�                                            //
'// Input :                                                  //
'// Output:                                                  //
'// ******************************************************** //
Sub Macro_Tab()
Attribute Macro_Tab.VB_ProcData.VB_Invoke_Func = "t\n14"

    Set preShape = Selection.ShapeRange(1)
    cntEnter = 1
    Call AddShape
    
End Sub



'// ******************************************************** //
'// Enter�L�[�ɑΉ�(Macro_Tab���s��Ɏg�p�\)               //
'// Input :                                                  //
'// Output:                                                  //
'// ******************************************************** //
Sub Macro_Enter()
Attribute Macro_Enter.VB_ProcData.VB_Invoke_Func = "e\n14"

    cntEnter = cntEnter + 1
    Call AddShape

End Sub



'// ******************************************************** //
'// preShape���R�s�[����֐�                                 //
'// Input :                                                  //
'// Output:                                                  //
'// ******************************************************** //
Function AddShape()
    Dim nextShape As Shape
    Dim connector   As Shape

    preShape.Copy
    ActiveSheet.Paste
    
    Set nextShape = Selection.ShapeRange(1)
    With nextShape
        .Left = preShape.Left + preShape.Width * 1.5
        .Top = preShape.Top + preShape.Height * 1.5 * cntEnter
        
        .TextFrame2.TextRange.Characters().Text = ""
        '.TextFrame2.TextRange.Characters().Font.Size = glbFontSize
        .TextFrame2.TextRange.Characters().Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        '.TextFrame2.VerticalAnchor = msoAnchorMiddle
        '.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    Set connector = ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 1, 1, 1, 1)      '// �J�M���R�l�N�^
    With connector
        '// �R�l�N�^��ڑ�
        .ConnectorFormat.BeginConnect preShape, 6
        .ConnectorFormat.EndConnect nextShape, 3
        '.RerouteConnections    '// RerouteConnections����ƍŒZ�Őڑ�����̂ŃR�����g�A�E�g
    
        '// �R�l�N�^�̐F�Ƒ�����ݒ�
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1 'glbLineWeight
        .Line.EndArrowheadStyle = msoArrowheadTriangle
        
        '// �J�M���R�l�N�^�̒��_��ݒ�(.Height�͌Œ�l���������߁A.Top�̍����ō��������߂�)
        '.Adjustments.Item(1) = (nextShape.Top - preShape.Top - glbDiffY) / (nextShape.Top - preShape.Top)
    End With
End Function


'// ******************************************************** //
'// �ϐ��Z�b�g                                               //
'// Input :                                                  //
'// Output:                                                  //
'// ******************************************************** //
Sub Macro_setVal()
Attribute Macro_setVal.VB_ProcData.VB_Invoke_Func = "s\n14"

    Dim selectShape As Shape
    Dim valShape    As Shape
    Dim connector   As Shape
    
    Call AddValShape(selectShape, valShape, connector)
    With connector
        '// �R�l�N�^��ڑ�
        .ConnectorFormat.BeginConnect selectShape, 7
        .ConnectorFormat.EndConnect valShape, 2
        '.RerouteConnections    '// RerouteConnections����ƍŒZ�Őڑ�����̂ŃR�����g�A�E�g
        
        '// �R�l�N�^�̐F�Ƒ�����ݒ�
        .Line.ForeColor.RGB = RGB(255, 0, 0)
    End With
    
End Sub


'// ******************************************************** //
'// �ϐ��Q��                                                 //
'// Input :                                                  //
'// Output:                                                  //
'// ******************************************************** //
Sub Macro_readVal()
Attribute Macro_readVal.VB_ProcData.VB_Invoke_Func = "r\n14"

    Dim selectShape As Shape
    Dim valShape    As Shape
    Dim connector   As Shape
    
    Call AddValShape(selectShape, valShape, connector)
    With connector
        '// �R�l�N�^��ڑ�
        .ConnectorFormat.BeginConnect valShape, 2
        .ConnectorFormat.EndConnect selectShape, 7
        '.RerouteConnections    '// RerouteConnections����ƍŒZ�Őڑ�����̂ŃR�����g�A�E�g
        
        '// �R�l�N�^�̐F�Ƒ�����ݒ�
        .Line.ForeColor.RGB = RGB(0, 0, 255)
    End With

End Sub


'// ******************************************************** //
'// �ϐ��pshape���쐬����֐�                                 //
'// Input :                                                  //
'// Output:                                                  //
'// ******************************************************** //
Function AddValShape(selectShape As Shape, valShape As Shape, connector As Shape)
    
    Set selectShape = Selection.ShapeRange(1)
    
    Set valShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, selectShape.Left + selectShape.Width * 1.5, selectShape.Top, 200, 50)
    With valShape
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1
        
        .TextFrame2.TextRange.Characters().Text = ""
        '.TextFrame2.TextRange.Characters().Font.Size = glbFontSize
        .TextFrame2.TextRange.Characters().Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        '.TextFrame2.VerticalAnchor = msoAnchorMiddle
        '.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    Set connector = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, 1, 1, 1, 1)      '// �Ȑ��R�l�N�^
    With connector
        .Line.Weight = 1 'glbLineWeight
        .Line.EndArrowheadStyle = msoArrowheadTriangle
    End With
    
    '// valShape��I��
    valShape.Select
End Function



Attribute VB_Name = "Module1"
Sub Check_Density_For_One_Row()

    Dim wsT As Worksheet          ' Template sheet
    Dim targetRow As Long         ' Шалгах м?рийн дугаар
    Dim i As Long
    Dim v As Variant
    Dim hasValue As Boolean
    Dim allOK As Boolean
    Dim tryCount As Long
    Const MAX_TRIES As Long = 1000    ' Хэдэн удаа random-дож ?зэх вэ

    Set wsT = ThisWorkbook.Worksheets("Template")

    ' 1) М?рийн дугаарыг хэрэглэгчээс авна
    targetRow = CLng(Application.InputBox( _
                     Prompt:="Шалгах м?рийн дугаараа оруулна уу (Data sheet-ийн м?р):", _
                     Title:="М?р сонгох", _
                     Default:=2, _
                     Type:=1))

    If targetRow <= 0 Then Exit Sub   ' Cancel дарсан эсвэл буруу тоо оруулсан бол

    ' 2) Template дээр тухайн м?рийн дугаарыг бичнэ
    '    Хэрэв чиний загвар J1-г ашигладаг бол "L2"-г "J1" болгож солино.
    wsT.Range("L2").Value = targetRow

    ' 3) Random-ыг хэд хэд эрг??лж, нягтрал 95-аас дээш болтол шалгана
    For tryCount = 1 To MAX_TRIES

        wsT.Calculate   ' RANDBETWEEN гэх мэт томъёонуудыг шинэчилнэ

        hasValue = False
        allOK = True

        ' E27, F27, G27, H27, I27 н?дн??дийг шалгана
        For Each v In Array("E27", "F27", "G27", "H27", "I27")
            With wsT.Range(v)
                If Trim(.Value & "") <> "" Then      ' Хоосон биш байвал л шалгана
                    hasValue = True
                    If .Value < 95 Or .Value > 100 Then
                        allOK = False
                    End If
                End If
            End With
        Next v

        ' Хэрэв утгатай н?д байсан б?г??д б?гд 95–100 хооронд байвал OK
        If hasValue And allOK Then
            MsgBox "М?р #" & targetRow & " дээр нягтралын хувь 95-аас дээш (95–100) боллоо." & vbCrLf & _
                   "Оролдлого: " & tryCount, vbInformation
            Exit Sub
        End If

    Next tryCount

    ' Хэрэв энд х?рсэн бол MAX_TRIES х?рээд ч н?хц?л биелээг?й гэсэн ?г
    MsgBox "М?р #" & targetRow & _
           " дээр " & MAX_TRIES & " удаа random-дсон ч нягтрал 95–100% н?хц?лд х?рсэнг?й.", _
           vbExclamation

End Sub



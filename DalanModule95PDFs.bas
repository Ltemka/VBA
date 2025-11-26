Attribute VB_Name = "Module2"
Sub Randomize_And_SavePDF_PerRow()

    Dim wsT As Worksheet      ' Template
    Dim wsD As Worksheet      ' Data
    Dim i As Long             ' Data row index
    Dim j As Long             ' random loop
    Dim lastRow As Long
    Dim pdfPath As String
    Dim fileName As String
    
    Dim cellsArr As Variant
    Dim addr As Variant
    Dim nonEmptyCount As Long
    Dim ok As Boolean

    Set wsT = ThisWorkbook.Worksheets("Template")  ' Template sheet
    Set wsD = ThisWorkbook.Worksheets("Data")      ' Data sheet
    
    ' Data sheet дээрх сүүлийн мөр (A баганаар тооцож байна)
    lastRow = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    
    ' PDF хадгалах хавтас: файлынхаа хавтсанд "PDFs" гэдэг фолдер үүсгэнэ
    pdfPath = ThisWorkbook.Path & "\PDFs"
    If Dir(pdfPath, vbDirectory) = "" Then
        MkDir pdfPath
    End If

    cellsArr = Array("E27", "F27", "G27", "H27", "I27")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Мөр бүрээр гүйж PDF үүсгэнэ
    For i = 2 To lastRow          ' 2-р мөрөөс эхэлнэ (толгой мөрийг алгасна)
        
        ' 1) Template дээр J1-д энэ мөрийн дугаарыг өгнө
        '    Хэрвээ L2 ашигладаг бол J1-ийг L2 болгож солино.
        wsT.Range("J1").Value = i
        
        ' 2) Энэ мөр дээр нягтрал OK болох хүртэл 1000 удаа random-дож шалгана
        For j = 1 To 1000
            
            wsT.Calculate   ' RANDBETWEEN-үүд шинэчлэгдэнэ
            
            nonEmptyCount = 0
            ok = True
            
            ' E27, F27, G27, H27, I27 нүднүүдийг шалгана
            For Each addr In cellsArr
                With wsT.Range(addr)
                    If Trim(.Value & "") <> "" Then  ' утгатай нүд л шалгана
                        nonEmptyCount = nonEmptyCount + 1
                        If .Value < 95 Or .Value > 100 Then
                            ok = False
                            Exit For
                        End If
                    End If
                End With
            Next addr
            
            ' Хэрвээ бүгд хоосон байвал: энэ мөрийг зүгээр орхи, PDF үүсгэхгүй
            If nonEmptyCount = 0 Then
                Exit For    ' энэ мөр дууссан → дараагийн мөр рүү
            End If
            
            ' Хэрвээ утгатай бүх нүд 95–100 хооронд байвал → PDF үүсгэнэ
            If ok Then
                fileName = pdfPath & "\Row_" & i & "_" & _
                           Format(Now, "yyyymmdd_hhnnss") & ".pdf"
                
                wsT.ExportAsFixedFormat Type:=xlTypePDF, _
                    fileName:=fileName, Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False
                
                Exit For    ' энэ мөр OK боллоо → дараагийн мөр рүү
            End If
            
        Next j
        
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Data sheet дээрх бүх мөрийг боловсруулж дууслаа. 'PDFs' хавтсанд PDF-үүдээ шалгаарай.", vbInformation

End Sub




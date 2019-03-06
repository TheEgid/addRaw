Attribute VB_Name = "addRaw"
Option Explicit
Option Base 1

Sub add_Raw()

Dim fd As FileDialog
Dim wLetterWb As Workbook
Dim AWb As Workbook
Dim arrUKPr() As Variant
Dim arrUKP() As Variant
Dim arrResultChekedName
Dim count As Integer
Dim count2 As Integer
Dim i As Integer
Dim LAT As Boolean
Dim Finalrow As Integer
Dim fas_qty As Integer
Dim oFind
Dim tempname As Variant

Dim ChekedName As String
Dim ChekedRaw As Boolean
Dim filenameend As String
Dim Filename, FilePath, MyFilename As String
Dim spezObj As Object

Application.ScreenUpdating = False
Set wLetterWb = ActiveWorkbook

    'получаем полный список возможных кодов с листа ѕисьмо
    wLetterWb.Sheets("ѕисьмо").Activate
    arrUKPr = Range(Cells(5, 3), Cells(200, 3).Address) 'максимум 200
    
    For count = 1 To UBound(arrUKPr, 1)
        If Len(arrUKPr(count, 1)) = 7 Then
            count2 = count2 + 1
        End If
    Next count
    
    ReDim arrUKP(count2 + 1, 1)
    
    For count = 1 To UBound(arrUKPr, 1)
        If Len(arrUKPr(count, 1)) = 7 Then
            arrUKP(count, 1) = arrUKPr(count, 1)
        End If
    Next count

    'ActiveWorkbook.Sheets("ѕ”— ").Activate
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            If FolderExists(Environ("USERPROFILE") & "\Desktop\_“естирование") = True Then
                .InitialFileName = Environ("USERPROFILE") & "\Desktop\_“естирование"
            Else
                .InitialFileName = Environ("USERPROFILE") & "\Desktop"
            End If
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "‘айлы Excel", "*.xlsx;*.xls" ' .Filters.Add "‘айлы Excel", "*.xlsx;*.xls;*.xlsm;*.xlsb"
        End With
    fd.Show

    If (fd.SelectedItems.count = 0) Or (fd.SelectedItems.count < 1) Then
        Exit Sub
    End If
    
    
For i = 1 To fd.SelectedItems.count    'for count of files

        Workbooks.Open Filename:=fd.SelectedItems(i), ReadOnly:=True, UpdateLinks:=False  'open each file
        Range("A1").Activate
        FilePath = ActiveWorkbook.path
        
        If InStr(CStr(ActiveWorkbook.Name), "'") <> 0 Then
            MsgBox "Ќедопустимый символ ' в названии файла - " & ActiveWorkbook.Name
            ActiveWorkbook.Close False
            Exit Sub
        End If
        
        tempname = CheckTheName(ActiveWorkbook.Name)
        If tempname = "" Then
            i = i 'пропуск цикла
            ActiveWorkbook.Close savechanges:=False ' ||
        Else
            Set AWb = ActiveWorkbook
        
            If WorksheetIsExist("Ѕланк заказа") = True Then
                LAT = False
                AWb.Sheets("Ѕланк заказа").Activate
                Range("B15").Select
            ElseIf WorksheetIsExist("Blank Order") = True Then
                LAT = True
                AWb.Sheets("Blank Order").Activate
                Range("B15").Select
            Else:
                MsgBox (AWb.Name & " лист с бланком не найден!"): Exit Sub
            End If
        
            arrResultChekedName = Split(tempname, "|")
            ChekedName = arrResultChekedName(0)
            filenameend = arrResultChekedName(1)
            If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
            Finalrow = WorksheetFunction.CountA(Range("A:A")) + 6 'финальна¤ строка
            
            Range(Cells(1, 2), Cells(Finalrow, 2)).Select
            For Each oFind In Selection
                    
                    'услови¤ - значение найдено в массиве и кол-во не пустое
                    If (BoolisValueInArray(arrUKP, oFind.Value) = True) And (oFind.Offset(0, 5).Value > 0) Then
                        fas_qty = CInt(oFind.Offset(0, 5).Value)
                        Finalrow = Finalrow + 1
                        oFind.Interior.Color = RGB(127, 255, 212)
                        oFind.Offset(0, 5).Interior.Color = RGB(127, 255, 212)
                        Set spezObj = Get_SpezProduct(ChekedName, oFind.Value, wLetterWb)
                        If spezObj Is Nothing = False Then
                                ChekedRaw = True 'сырье присутствует
                                
                                Cells(Finalrow, 1).Value = spezObj.Get_long_kod()
                                Cells(Finalrow, 2).Value = spezObj.Kod_raw
                                Cells(Finalrow, 4).Value = spezObj.Fab_raw
                                Cells(Finalrow, 6).Value = spezObj.Name_raw
                                Cells(Finalrow, 7).Value = spezObj.Calculate_qty_raw(fas_qty) 'рассчет!
                                Cells(Finalrow, 7).Interior.Color = RGB(255, 69, 0)
                                If LAT = True Then
                                    Cells(Finalrow, 11).Value = spezObj.Weight_raw
                                    Cells(Finalrow, 12).Value = "=RC[-1]*RC[-5]" 'формула суммы веса
                                Else
                                    Cells(Finalrow, 18).Value = spezObj.Weight_raw
                                    Cells(Finalrow, 19).Value = "=RC[-1]*RC[-12]" 'формула суммы веса
                                End If
                                
                                If spezObj.Mix <> "" Then
                                    Finalrow = Finalrow + 1
                                    Processing_MIX_raw Cells(Finalrow, 1) 'добавл¤ем ћикс
                                End If
                                
                        End If
                    End If

            Next oFind
            
    
            Range(Cells(1, 7), Cells(Finalrow, 7)).Select
            For Each oFind In Selection
                    If (oFind.Value = "«а¤вка,кор") Or (oFind.Value = "Order, boxes") Then oFind.Interior.Color = RGB(255, 69, 0)
                    If oFind.Interior.Color <> RGB(255, 69, 0) Then oFind.ClearContents 'удал¤ем нераскрашенное

            Next oFind
            
        MyFilename = Left(AWb.Name, InStrRev(AWb.Name, ".") - 1)
        Application.DisplayAlerts = False
        MyFilename = MyFilename & " (" & CInt(Range("B8")) * 1 & ")" & filenameend
        MyFilename = DateMinis3days(MyFilename) 'правим дату минус 3 дн¤
        If (CInt(Range("B8")) <> 0) And (ChekedRaw = True) Then AWb.SaveAs Filename:=MyFilename
                    
       AWb.Close savechanges:=False ' ||
       End If
'    Stop
    
Next i


Application.ScreenUpdating = True
MsgBox ("Done!")
End Sub


Public Function Get_SpezProduct(myFeature As String, myKod As String, ishodWb As Workbook) As clsmSproduct
'функци¤ возвращает объект класса —пецпродукт
Dim myProduct As New clsmSproduct
Dim ADO As New ADO

Dim Zapros As String
Dim arrA() As Variant

ADO.DataSource = ishodWb.path & "\" & ishodWb.Name
ADO.Header = False
Zapros = "SELECT F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11 FROM [ѕисьмо$] WHERE F1 ='" & myFeature & "' AND F3 ='" & myKod & "';"
ADO.Query Trim(Zapros)

    If ADO.Recordset.EOF Then
        Set Get_SpezProduct = Nothing 'null
    Else
        arrA = ADO.ToArray

        myProduct.Feature = arrA(1, 1)              'особенность спецпродукта
        myProduct.Kod_fas = arrA(1, 3)              'код фасовки
        myProduct.Name_fas = arrA(1, 4)             'наименование фасовки
        myProduct.Weight_fas = arrA(1, 5) * 1       'вес фасовки
        myProduct.Kod_raw = arrA(1, 7)              'код сырь¤
        myProduct.Name_raw = arrA(1, 8)             'наименование сырь¤
        myProduct.Weight_raw = arrA(1, 9) * 1       'вес сырь¤
        myProduct.Fab_raw = arrA(1, 10)             'фабрика
        If IsNull(arrA(1, 11)) Then arrA(1, 11) = ""
        myProduct.Mix = arrA(1, 11)                 'микс
        Set Get_SpezProduct = myProduct 'return
        Set myProduct = Nothing
    End If
'Stop
ADO.Destroy
End Function

Function CheckTheName(Name As String) As String

Dim Checking1, Checking2, Checking3, Checking4, Checking5, Checking6 As Boolean
    'проверки -
Checking1 = InStr(1, CStr(Name), "—џ–№®", vbTextCompare) = 0 'не содержитс¤ обычное —џ–№®
Checking2 = (InStr(1, CStr(Name), "Id", vbTextCompare) > 0) Or (InStr(1, CStr(Name), "“ир", vbTextCompare) > 0) 'содержитс¤ “рансжир0
Checking3 = (InStr(1, CStr(Name), "–аи", vbTextCompare) > 0) Or (InStr(1, CStr(Name), "–авин", vbTextCompare) > 0) 'содержитс¤  ошер–аввин
Checking4 = (InStr(1, CStr(Name), " ор", vbTextCompare) > 0)  'содержитс¤  ошер
Checking5 = (InStr(1, CStr(Name), " ий", vbTextCompare) > 0)  'содержитс¤  итай365
Checking6 = (InStr(1, CStr(Name), "’ал", vbTextCompare) > 0) Or (InStr(1, CStr(Name), "’алaл", vbTextCompare) > 0) 'содержитс¤ ’ал¤ль

 If (Checking1 = True) And (Checking2 = False) And (Checking3 = False) And (Checking4 = False) And (Checking5 = False) And (Checking6 = False) Then
     CheckTheName = "ќб¤" & "| o—џ–№®.xlsx"
 End If
 
 If (Checking1 = True) And (Checking2 = True) Then
     CheckTheName = "“ры" & "| т—џ–№®.xlsx"
 End If
 
 If (Checking1 = True) And (Checking3 = True) Then
     CheckTheName = " ор–ан" & "| р—џ–№®.xlsx"
 End If
 
 If (Checking1 = True) And (Checking3 = False) And (Checking4 = True) Then
     CheckTheName = " ош" & "| k—џ–№®.xlsx"
 End If
 
 If (Checking1 = True) And (Checking5 = True) Then
     CheckTheName = " и" & "| ch—џ–№®.xlsx"
 End If
 
 If (Checking1 = True) And (Checking6 = True) Then
     CheckTheName = "’ал¤ль" & "| hl—џ–№®.xlsx"
 End If
    
End Function

Function Processing_MIX_raw(oRange As Range)

Dim MixFirst As Range, Mix_second As Range
Dim cutter As Variant

Set MixFirst = oRange.Offset(-1, 1)
Set Mix_second = oRange.Offset(0, 1)

    If Len(MixFirst) < 10 Then MsgBox "processing_raw_mix Error with: " & MixFirst.Value: Exit Function
    
    cutter = Split(Trim(MixFirst), " ") 'короткий код - absolutely need spacing
    If UBound(cutter) <> 1 Then MsgBox "processing_raw_mix Error! Perhaps space the space is not found!": Exit Function
    MixFirst = cutter(0)
    Mix_second = cutter(1)
    
    Mix_second.Offset(0, 2) = MixFirst.Offset(0, 2) 'фабрика
    
    Mix_second.Offset(0, 9) = MixFirst.Offset(0, 9) '!! вес - только одинаковый вес
    Mix_second.Offset(0, 10).FormulaR1C1 = MixFirst.Offset(0, 10).FormulaR1C1 '!! вес всего
    
    cutter = Split(Trim(MixFirst.Offset(0, 4)), " ") 'наименование - absolutely need spacing
    If UBound(cutter) <> 1 Then MsgBox "processing_raw_mix Error! Perhaps space the space is not found!": Exit Function
    MixFirst.Offset(0, 4) = cutter(0)
    Mix_second.Offset(0, 4) = cutter(1)
    
    cutter = WorksheetFunction.RoundUp((MixFirst.Offset(0, 5) / 2), 0) 'кол-во
    MixFirst.Offset(0, 5) = cutter      'фабрика
    Mix_second.Offset(0, 5) = cutter
    Mix_second.Offset(0, 5).Interior.Color = MixFirst.Offset(0, 5).Interior.Color
    
    MixFirst.Offset(0, -1) = MixFirst & "-0" & MixFirst.Offset(0, 9) * 1000 'длинный код
    Mix_second.Offset(0, -1) = Mix_second & "-0" & Mix_second.Offset(0, 9) * 1000

End Function

Function FolderExists(ByRef path As String) As Boolean
   On Error Resume Next
   FolderExists = GetAttr(path)
End Function

Function WorksheetIsExist(iName$) As Boolean
   On Error Resume Next
   WorksheetIsExist = (TypeName(ActiveWorkbook.Worksheets(iName$)) = "Worksheet")
End Function

'проверка на вхождение элемента в массив
Function BoolisValueInArray(ByVal myarray As Variant, ByVal element As Variant) As Boolean
    If IsNumeric(Application.Match(element, myarray, 0)) Then
        BoolisValueInArray = True
    Else
        BoolisValueInArray = False
    End If
End Function

Function DateMinis3days(in_str As String) As String
Dim re As Object
Dim s As Variant
Dim out_date As Date
Dim datSTR As String
Set re = CreateObject("vbscript.regexp")
    re.Pattern = "\d+[,._/\|-]\d+[,._/\|-][\d]+"
    re.Global = True
    If re.test(in_str) Then
        datSTR = re.Execute(in_str)(0)
        s = Split(in_str, datSTR, 2) 'раздел¤ем
        datSTR = Replace(datSTR, "-", ".")
        datSTR = Replace(datSTR, "/", ".")
        datSTR = Replace(datSTR, "\", ".")
        datSTR = Replace(datSTR, ",", ".")
        datSTR = Replace(datSTR, "_", ".")
        out_date = DateAdd("d", -3, CDate(datSTR))
        DateMinis3days = s(0) & CStr(out_date) & s(1) 'соедин¤ем
    Else
       'pass
    End If
Set re = Nothing
End Function

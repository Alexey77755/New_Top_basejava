package ru.javawebinar.basejava;

Option Explicit
' номера столбцов по порядку в массиве rng
Dim cActionType     As Integer
Dim cProductCode    As Integer
Dim cProductName    As Integer
Dim cSupplierCode   As Integer
Dim cSupplierName   As Integer
Dim cItemCode       As Integer
Dim cItemName       As Integer
Dim cOldOptions     As Integer
Dim cNewOptions     As Integer
Dim cAddOptions     As Integer
Dim cComm           As Integer
Dim wsForm          As Worksheet
Dim wsTerra         As Worksheet
Dim firstRow        As Integer
Dim lastRow         As Long
Public i            As Long
Dim Lc              As Long
Dim strItemCode     As String
'переменные для подключения к БД
Dim rsOra
Dim cnORA
Dim strQry          As String
'словари
Dim DictTer         As Object
Dim DictMatrRc      As Object
Dim DictFO          As Object
Dim DictMLT         As Object
Dim DictMLTKK       As Object
Dim DictKK          As Object
Dim DictRuchKK      As Object
Dim dictFUZ         As Object
Public errcount
Public DictVirtUzel As Object
Dim strUzelSH As String: Dim irow As Long
Dim zaendrow
Dim wsResult        As Worksheet
Public zV           As Integer
Public zH           As Integer
Dim MaxZH           As Long
Dim TypePost        As String
Dim str
Dim strGolduzel
Dim varItem
Dim uz_item         As String
Dim errstr          As String
Dim finALC: Dim finPIV: Dim finPIVBA


Sub Connect_DB()
    Set cnORA = CreateObject("ADODB.Connection")
    On Error Resume Next
    cnORA.Open = "Provider=OraOLEDB.Oracle;Password=rdr2009;User ID=reader2;Data Source=PITER506;Persist Security Info=True"
    If Err.Number = -2147467259 Or Err.Number = 3706 Then
        MsgBox "Не удается подключиться к БД."
    End If
    On Error GoTo 0
    Set rsOra = CreateObject("ADODB.Recordset")

End Sub

Sub Binding_Form()
    Set wsForm = Worksheets("Форма_A02ZA")
    With wsForm.Cells
        On Error Resume Next
        cActionType = .Find(What:="Тип действия", LookAt:=xlWhole).Column
        On Error GoTo 0
        If cActionType = 0 Then
            MsgBox "Форма A02ZA не найдена"
            Exit Sub
        End If
        cProductCode = .Find(What:="Код GOLD товара", LookAt:=xlWhole).Column
        cProductName = .Find(What:="Наименование товара", LookAt:=xlWhole).Column
        cSupplierCode = .Find(What:="Код GOLD поставщика", LookAt:=xlWhole).Column
        cSupplierName = .Find(What:="Наименование поставщика", LookAt:=xlWhole).Column
        cItemCode = .Find(What:="Код узла", LookAt:=xlWhole).Column
        cItemName = .Find(What:="Описание узла", LookAt:=xlWhole).Column
        cOldOptions = .Find(What:="СТАРЫЕ НАСТРОЙКИ", LookAt:=xlWhole).Column
        cNewOptions = .Find(What:="НОВЫЕ НАСТРОЙКИ", LookAt:=xlWhole).Column
        cAddOptions = .Find(What:="ТВЗ", LookAt:=xlWhole).Column
        cComm = .Find(What:="Комментарий", LookAt:=xlWhole).Column

        firstRow = .Find(What:="Тип действия", LookAt:=xlWhole).Row + 3
        lastRow = WorksheetFunction.Max(.Cells(.Rows.Count, cActionType).End(xlUp).Row, .Cells(.Rows.Count, cProductCode).End(xlUp).Row, .Cells(.Rows.Count, cSupplierCode).End(xlUp).Row, _
                .Cells(.Rows.Count, cItemCode).End(xlUp).Row, .Cells(.Rows.Count, cOldOptions).End(xlUp).Row, .Cells(.Rows.Count, cOldOptions + 1).End(xlUp).Row, _
                .Cells(.Rows.Count, cOldOptions + 2).End(xlUp).Row, .Cells(.Rows.Count, cOldOptions + 3).End(xlUp).Row, .Cells(.Rows.Count, cOldOptions + 4).End(xlUp).Row, .Cells(.Rows.Count, cOldOptions + 5).End(xlUp).Row)
        Do While Len(.Cells(lastRow + 1, 4).Value2) > 0 Or Len(.Cells(lastRow + 1, 6).Value2) > 0
            lastRow = lastRow + 1
        Loop
    End With

End Sub

Sub Unloading()
    Dim wsUnload    As Worksheet
    Dim NewList     As Object
    Dim TempRow     As Long
    Dim newTempRow  As Long
    Dim sh          As Worksheet
    Dim subQry      As String
    Dim str
    Dim NotlastLevelItem As Boolean
    Dim error       As String
    Dim BenchMark   As Double
    Dim dTime       As Double
    Dim startTime   As Date
    Dim endTime     As Date
    Dim Sheet       As Worksheet
    Dim k           As Long
    Dim l           As Long
    BenchMark = Timer
    startTime = Now()    ' at the start of the analysis
    ' at the end of the analysis


    Application.ScreenUpdating = False
    Application.StatusBar = ""
    Call Binding_Form

    If lastRow < firstRow Then
        MsgBox "Не найдены данные в форме"
        Exit Sub
    End If

    Call Connect_DB

    'проверка на фильтры или скрытые строки
    If wsForm.FilterMode Then
        MsgBox ("Снимите фильтры с листа ""Форма_A02ZA""")
        End
    End If


    For i = firstRow To lastRow
        If wsForm.Rows(i).EntireRow.Hidden Or wsForm.Rows(i).RowHeight < 8 Then
            MsgBox ("На листе ""Форма_A02ZA"" есть скрытые строки. Нужно их раскрыть")
            End
        End If
    Next


    On Error Resume Next
    Set wsUnload = Sheets("Выгрузка")
    On Error GoTo 0
    If wsUnload Is Nothing Then
        MsgBox "Лист ""Выгрузка"" отсутствует"
        Exit Sub
    End If

    wsUnload.AutoFilterMode = False
    TempRow = wsUnload.Cells(wsUnload.Rows.Count, 1).End(xlUp).Row
    'If wsUnload.Rows.Count = 1048576 Then TempRow = 1048576
    If TempRow > 2 Then wsUnload.Range("A3:R" & TempRow).ClearContents

    TempRow = 2
    k = 0
    l = 1
    Application.DisplayAlerts = False
    ActiveWorkbook.Unprotect (1347)
    For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.Name Like "Выгрузка_*" Then

            Sheet.Delete
        End If
    Next
    wsUnload.Range("A2:R2").AutoFilter
    For i = firstRow To lastRow
        'Application.Wait (TimeValue("00:00:05"))
        'wsForm.Cells(i, 10).NumberFormat = "0"
        wsForm.Cells(i, 3).NumberFormat = "@"
        wsForm.Cells(i, 5).NumberFormat = "@"
        wsForm.Range(wsForm.Cells(i, 10), wsForm.Cells(i, 17)).NumberFormat = "@"
        wsForm.Range(wsForm.Cells(i, 20), wsForm.Cells(i, 28)).NumberFormat = "@"

        error = ""
        If cnORA.State = 1 Then
            wsUnload.Activate
            'wsUnload.AutoFilterMode = True
            strItemCode = wsForm.Cells(i, cItemCode).Value
            If Len(strItemCode) = 0 Then strItemCode = "&"
            For Each str In Split(Application.Trim(Replace(Replace(Replace(strItemCode, ";", ","), ",", ", "), ",", " ,")), ",")    'узел
                strQry = "Select CS.ARTUC.ARACEXR as ""Код товара"", cs.pkstrucobj.get_desc(1, CS.ARTUC.ARACINR, 'RU') as ""Наим.товара"", CS.ARTUC.aracexvl as ""ЛВ"", (select CS.ARTUL.ARUTYPUL from CS.ARTUL where CS.ARTUC.ARACINL = CS.ARTUL.ARUCINL) as ""ЛЕ"", CS.FOUDGENE.FOUCNUF as ""Код пост."", " & _
                        "CS.FOUDGENE.FOULIBL as ""Наим.пост."", CS.ARTUC.ARANFILF as ""Адр.цеп."", CS.FOUCCOM.FCCNUM as ""КК"",CS.FOUCCOM.fcclib as ""Наименование КК"", CS.ARTUC.ARASITE as ""Код узла"", CS.ARTUC.araddeb as ""Нач.дата"", CS.ARTUC.ARADFIN as ""Кон.дата"", CS.ARTUC.aracean as ""Код EAN"", CS.ARTUC.ARATFOU as ""Приз.осн.пост"", " & _
                        "CS.ARTUC.ARAMUA as ""КратностьЗЕ"", CS.ARTUC.aramincde as ""Мин.заказЗЕ"", CS.ARTUC.aramaxcde as ""Макс.заказЗЕ"", CASE CS.ARTUC.ARATCDE WHEN 1 THEN 'Заказываемый' WHEN 2 THEN 'Разреш. только прием товара' WHEN 3 THEN 'Запрещ. заказ и прием' END AS ""Повед. ассорт."" " & _
                        "From CS.ARTUC inner join CS.FOUDGENE on CS.ARTUC.ARACFIN  = CS.FOUDGENE.FOUCFIN left join CS.FOUCCOM on CS.FOUCCOM.FCCCCIN = CS.ARTUC.ARACCIN where CS.ARTUC.ARADFIN > sysdate"
                '-------------------------------------------------------------------------------------------------------------------------------------

                '-------------------------------------------------------------------------------------------------------------------------------------
                If Len(wsForm.Cells(i, cProductCode).Value) > 0 Then    'товар
                    strQry = strQry + " and CS.ARTUC.ARACEXR = '" & wsForm.Cells(i, cProductCode).Value & "' "
                    If Len(wsForm.Cells(i, cOldOptions + 1).Value) > 0 Then    'ЛВ
                        If IsNumeric(wsForm.Cells(i, cOldOptions + 1).Value2) Then
                            If Fix(wsForm.Cells(i, cOldOptions + 1).Value2) = CLng(wsForm.Cells(i, cOldOptions + 1).Value2) Then
                                strQry = strQry + " and CS.ARTUC.aracexvl = '" & wsForm.Cells(i, cOldOptions + 1).Value & "' "
                            End If
                        End If
                    End If
                    '        Else
                    '             error = "Для выгрузки необходимо дополнить фильтр: код товара"
                    If Len(wsForm.Cells(i, cOldOptions + 2).Value) > 0 Then    'ЛЕ
                        If IsNumeric(wsForm.Cells(i, cOldOptions + 2).Value2) Then
                            If Fix(wsForm.Cells(i, cOldOptions + 2).Value2) = CLng(wsForm.Cells(i, cOldOptions + 2).Value2) Then
                                strQry = strQry + " and CS.ARTUC.ARACEXTA = '" & wsForm.Cells(i, cOldOptions + 2).Value & "' "
                            End If
                        End If
                    End If
                End If
                '-------------------------------------------------------------------------------------------------------------------------------------

                If Len(wsForm.Cells(i, cSupplierCode).Value) > 0 Then    'поставщик
                    subQry = "select CS.FOUDGENE.foutype from CS.FOUDGENE where CS.FOUDGENE.FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value & "'"
                    rsOra.Open subQry, cnORA
                    If rsOra.bof Then
                        error = "Код поставщика не найден в системе"
                        rsOra.Close
                    Else
                        strQry = strQry + " and CS.FOUDGENE.FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value & "'"
                        If rsOra!FOUTYPE <> "1" Then    'не внешний поставщик
                            If Len(wsForm.Cells(i, cProductCode).Value) = 0 Then
                                error = "Для выгрузки необходимо дополнить фильтр: код товара"
                            End If
                        Else
                            If Len(str) > 0 And str <> "&" And Len(wsForm.Cells(i, cProductCode).Value) = 0 Then error = "Для выгрузки необходимо дополнить фильтр: код товара"
                        End If
                        rsOra.Close
                        If Len(wsForm.Cells(i, cOldOptions + 3).Value) > 0 Then    'АЦ
                            strQry = strQry + " and CS.ARTUC.ARANFILF = '" & wsForm.Cells(i, cOldOptions + 3).Value & "'"
                        End If
                    End If
                Else
                    If Len(wsForm.Cells(i, cProductCode).Value) = 0 Then error = "Для выгрузки необходимо дополнить фильтр: код товара"
                End If

                If error = "" Then
                    '-------------------------------------------------------------------------------------------------------------------------------------
                    If Len(wsForm.Cells(i, cOldOptions + 4).Value) > 0 Then    'КК
                        strQry = strQry + " and CS.FOUCCOM.FCCNUM = '" & wsForm.Cells(i, cOldOptions + 4).Value & "'"
                    End If
                    '-------------------------------------------------------------------------------------------------------------------------------------
                    If Len(wsForm.Cells(i, cProductCode).Value) > 0 Then
                        If Len(str) > 0 And str <> "&" Then
                            strQry = strQry + " and CS.ARTUC.ARASITE in ( SELECT relid  FROM cs.resrel where reldfin >= sysdate START WITH relpere  = '" & str & "' CONNECT BY PRIOR relid = relpere union all SELECT relpere FROM cs.resrel where reldfin >= sysdate and relpere not in ('94001','7') START WITH relid  = '" & str & "' CONNECT BY PRIOR relpere = relid union all select " & str & " from dual) "
                        End If
                    End If

                    rsOra.Open strQry, cnORA, 3
                    k = k + rsOra.RecordCount
                    If k >= 1035000 Then
                        ActiveWorkbook.Unprotect (1347)
                        Set NewList = Worksheets.Add(After:=wsUnload)
                        NewList.Name = "Выгрузка_" & l
                        wsUnload.Range("A1:R2").Copy
                        NewList.Range("A1:R2").PasteSpecial
                        ActiveWindow.Zoom = 80
                        Set wsUnload = Sheets("Выгрузка_" & l)
                        wsUnload.Range("A2:R2").AutoFilter
                        wsUnload.Columns("A:R").ColumnWidth = 15
                        l = l + 1
                        k = 0
                        'ActiveSheet = wsUnload
                    End If
                    TempRow = wsUnload.Cells(wsUnload.Rows.Count, 1).End(xlUp).Row
                    wsUnload.Cells(TempRow + 1, 1).CopyFromRecordset rsOra
                    rsOra.Close
                    TempRow = wsUnload.Cells(wsUnload.Rows.Count, 1).End(xlUp).Row
                Else
                    wsUnload.Cells(TempRow + 1, 1).Value = error
                    TempRow = TempRow + 1
                End If
            Next str
        End If
        endTime = Now()
        dTime = endTime - startTime
        'dTime = dTime / (60 * 60 * 24) 'convert seconds into days
        DoEvents
        Application.StatusBar = "Обработано: " & Round((i - 9) * 100 / (lastRow - 9), 0) & "% (" & i - 9 & " строк из " & lastRow - 9 & ") " & Format(dTime, "hh:mm:ss") & ""
        'curNumberFile = curNumberFile + 1
    Next i



    'Время работы
    wsForm.Cells(5, 4) = "Время работы = " & Round(Timer - BenchMark, 2)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ActiveWorkbook.Protect (1347)
End Sub

Sub Check()
    Dim TypeKK      As String
    Dim flagProd As Boolean: Dim flagSup As Boolean: Dim flagOF As Boolean: Dim flagND As Boolean
    Dim flagOD As Boolean: Dim flagItem As Boolean: Dim flagKK As Boolean
    Dim flagAll As Boolean: Dim flagNewLV As Boolean: Dim flagOldLV As Boolean
    Dim flagNewLE As Boolean: Dim flagTVZ As Boolean: Dim flagOsnPostRC As Boolean
    Dim flagOsnPostMag As Boolean: Dim flagAM As Boolean: Dim flagExtData As String
    Dim flagTMC As Boolean: Dim MassItem() As String: Dim strArray() As String
    Dim OsnYzel As String: Dim j As Integer
    Dim k As Integer: Dim z As Long: Dim zRow As Long
    Dim DictSmena As Object: Dim SmflagOld As Boolean: Dim SmflagNew As Boolean
    Dim SmReversOld As Boolean: Dim SmReversNew As Boolean: Dim flagFO As Boolean
    Dim falgYesRevers As Long: Dim Item: Dim SmYzel
    Dim SmStr: Dim Tr: Dim SmTr: Dim StrRow
    Dim strTer As String: Dim strTerMags As String: Dim YzelMags As String
    Dim NTGOLD As String: Dim BenchMark As Double
    Dim SmenaItem: Dim SmenaItemRev: Dim a As Long: Dim OtherSmFlag As Boolean
    Dim OtherYzel As Object: Dim f As Long: Dim FlagPostaDates As Boolean
    MaxZH = 4: z = 3: BenchMark = Timer
    Application.ScreenUpdating = False
    Set wsResult = Worksheets("Отчет")

    Call Binding_Form

    If lastRow < firstRow Then
        MsgBox "Не найдены данные в форме"
        Application.EnableEvents = True
        Exit Sub
    Else
        'меняем буквы в контракте на английские
        toContracts ws:=wsForm, colNumber:=13, lLastRow:=lastRow    'Старые настройки
        toContracts ws:=wsForm, colNumber:=23, lLastRow:=lastRow    'Новые настройки
    End If

    'проверка на фильтры или скрытые строки
    If wsForm.FilterMode Then MsgBox ("Снимите фильтры с листа ""Форма_A02ZA"""): End

    For i = firstRow To lastRow
        If wsForm.Rows(i).EntireRow.Hidden Or wsForm.Rows(i).RowHeight < 8 Then
            MsgBox ("На листе ""Форма_A02ZA"" есть скрытые строки. Нужно их раскрыть")
            Application.EnableEvents = True
            End
        End If
    Next

    Call Connect_DB: Call DesignResult(wsResult)
    '-----------------------------------проверка актуальности шаблона формы-----------------------------------------------
    If cnORA.State = 1 Then
        strQry = "SELECT ATTVALDAT as Date_Form FROM CS.ATTRIVAL WHERE ATTCODE = 'A02ZA'"
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            If rsOra!Date_Form = CDate(wsForm.Cells(4, 3).Value) Then
                wsResult.Cells(1, 7).Value = "Шаблон формы актуальный": wsResult.Cells(1, 7).Interior.Color = RGB(146, 208, 80)
            Else
                wsResult.Cells(1, 7).Value = "Используется старый шаблон формы": wsResult.Cells(1, 7).Interior.Color = 255
            End If
        End If
        rsOra.Close
    End If
    '----------------------------------------------------------------------------------------------------------------------
    zV = 2: Call Dict_Proc: ActiveWorkbook.Unprotect (1347)
    With Worksheets("ТВЗ")
        .Protect Password:="1347", UserInterfaceOnly:=True
        .AutoFilterMode = False
        zRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If zRow > 2 Then .Range("A3:O" & zRow).ClearContents
        If .Visible = True Then .Visible = False
    End With
    ActiveWorkbook.Protect (1347)

    Set DictSmena = CreateObject("Scripting.Dictionary")

    Set dictFUZ = CreateObject("Scripting.Dictionary")
    strUzelSH = ""
    If cnORA.State = 1 Then
        'strQry = "select distinct  yzel1 from ok__nodes where yzel3 = 94003 and yzel1 is not null Union all select distinct yzel2 from ok__nodes where yzel3 = 94003 and yzel2 is not null"

        strQry = "select distinct  yzel1,nvl(yzel2,1) yzel2 from ok__nodes where yzel3 = 94003 and yzel1 is not null Union " & _
                " select '94009','1' from dual union select '94010','1' from dual union select '94011','1' from dual union" & _
                " select distinct yzel1,yzel2 from ok__nodes where yzel3 = 94003 and yzel2 is not null"
        'Debug.Print strQry
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            Do While Not rsOra.EOF
                strUzelSH = strUzelSH & rsOra!yzel1 & ","
                If dictFUZ.exists(CStr(rsOra!yzel2)) = False Then
                    uz_item = ""
                    For f = 2 To wsTerra.Cells(wsTerra.Rows.Count, 7).End(xlUp).Row

                        If CStr(wsTerra.Cells(f, 4).Value2) = CStr(rsOra!yzel2) Then
                            uz_item = uz_item & wsTerra.Cells(f, 2).Value2 & ","
                        End If
                    Next f
                    If Right(uz_item, 1) = "," Then uz_item = Left(uz_item, Len(uz_item) - 1)
                    dictFUZ.Add (CStr(rsOra!yzel2)), CStr(uz_item)
                End If
                rsOra.movenext
            Loop
        End If
        rsOra.Close
    End If


    For i = firstRow To lastRow
        errcount = 0
        'wsForm.Cells(i, 10).NumberFormat = "0"
        wsForm.Cells(i, 3).NumberFormat = "@": wsForm.Cells(i, 5).NumberFormat = "@"
        wsForm.Range(wsForm.Cells(i, 10), wsForm.Cells(i, 17)).NumberFormat = "@"
        wsForm.Range(wsForm.Cells(i, 20), wsForm.Cells(i, 28)).NumberFormat = "@"
        wsForm.Cells(i, 9).NumberFormat = "dd.mm.yyyy;@"
        wsForm.Range(wsForm.Cells(i, 18), wsForm.Cells(i, 19)).NumberFormat = "dd.mm.yyyy;@"
        If IsDate(wsForm.Cells(i, 9).Value) Then wsForm.Cells(i, 9).Value = CDate(wsForm.Cells(i, 9))
        If IsDate(wsForm.Cells(i, 18).Value) Then wsForm.Cells(i, 18) = CDate(wsForm.Cells(i, 18))
        If IsDate(wsForm.Cells(i, 19).Value) Then wsForm.Cells(i, 19) = CDate(wsForm.Cells(i, 19))
        With wsForm.Range(wsForm.Cells(i, 7), wsForm.Cells(i, 8))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
            .RowHeight = 15
        End With

        'Убираем лишние пробелы, нечитаемые символы
        Application.EnableEvents = False
        With wsForm
            For j = 2 To .Cells(9, .Columns.Count).End(xlToLeft).Column + 1
                .Cells(i, j).Value = Application.Trim(Replace(Replace(Replace(Replace(Replace(.Cells(i, j).Value, vbLf, " "), vbTab, " "), Chr(13), " "), Chr(10), " "), Chr(160), " "))
            Next j
        End With
        Application.EnableEvents = True


        zH = 4
        flagProd = True: flagSup = True: flagND = True: flagOD = True: flagItem = True: flagKK = True: flagAll = True: flagNewLV = True: flagOldLV = True: flagNewLE = True: flagExtData = ""
        flagOF = True: flagTVZ = True: flagOsnPostRC = False: flagOsnPostMag = False: flagAM = True: flagTMC = False
        TypePost = "": TypeKK = "": NTGOLD = "": OsnYzel = ""

        If wsForm.Cells(i, cProductCode).Value2 = "" Then
            zH = errFunction(zH, zV, i, "Код товара не заполнен"): flagProd = False
        Else
            If cnORA.State = 1 Then
                strQry = "select ARTETAT, cs.pkstrucobj.get_desc(1, CS.ARTRAC.ARTCINR, 'RU') as Name from CS.ARTRAC where ARTCEXR = '" & wsForm.Cells(i, cProductCode).Value2 & "'"
                rsOra.Open strQry, cnORA
                If rsOra.bof Then
                    zH = errFunction(zH, zV, i, "Код товара " & wsForm.Cells(i, cProductCode).Value2 & " не найден в базе данных")
                    flagProd = False
                Else
                    If rsOra!ARTETAT <> 1 Then zH = errFunction(zH, zV, i, "Код товара " & wsForm.Cells(i, cProductCode).Value2 & " заморожен"): flagProd = False
                    If wsForm.Cells(i, cProductName).Value2 = "" Then
                        zH = errFunction(zH, zV, i, "Описание товара не указано. Заполнено из системы"): wsForm.Cells(i, cProductName).Value2 = rsOra!Name
                    Else
                        NTGOLD = Application.Trim(Replace(Replace(Replace(rsOra!Name, vbLf, " "), vbTab, " "), Chr(160), " "))
                        If wsForm.Cells(i, cProductName).Value2 <> NTGOLD Then zH = errFunction(zH, zV, i, "Указанному коду товара " & wsForm.Cells(i, cProductCode).Value2 & " принадлежит другое наименование " & rsOra!Name & ". Необходимо скорректировать.")
                    End If
                End If
                rsOra.Close
            End If
        End If
        If wsForm.Cells(i, cSupplierCode).Value2 = "" Then
            zH = errFunction(zH, zV, i, "Код поставщика не заполнен"): flagSup = False
        Else
            If cnORA.State = 1 Then
                strQry = "select FOULIBL, FOUTYPE, FOUPAYS, FOUETAT from CS.FOUDGENE where FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value2 & "'"
                ' Debug.Print strQry
                rsOra.Open strQry, cnORA
                If rsOra.bof Then
                    zH = errFunction(zH, zV, i, "Код поставщика " & wsForm.Cells(i, cSupplierCode).Value2 & " не найден в базе данных")
                    flagSup = False
                ElseIf rsOra!FOUETAT <> 1 Then
                    zH = errFunction(zH, zV, i, "Код поставщика " & wsForm.Cells(i, cSupplierCode).Value2 & " заморожен"): flagSup = False
                ElseIf rsOra!FOULIBL <> wsForm.Cells(i, cSupplierName).Value2 Then
                    If wsForm.Cells(i, cSupplierName).Value2 = "" Then
                        wsForm.Cells(i, cSupplierName).Value2 = rsOra!FOULIBL
                        zH = errFunction(zH, zV, i, "Описание поставщика не указано. Заполнено из системы")
                        TypePost = rsOra!FOUTYPE & rsOra!FOUPAYS
                    Else
                        zH = errFunction(zH, zV, i, "Указанному коду поставщика " & wsForm.Cells(i, cSupplierCode).Value2 & " принадлежит другое наименование " & rsOra!FOULIBL & ". Необходимо скорректировать.")
                        TypePost = rsOra!FOUTYPE & rsOra!FOUPAYS
                    End If
                Else
                    TypePost = rsOra!FOUTYPE & rsOra!FOUPAYS
                End If
                rsOra.Close
            End If
        End If

        If flagSup Then
            If (wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Or _
                    wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика") And Left(TypePost, 1) <> "1" Then
                zH = errFunction(zH, zV, i, "Должен быть указан внешний поставщик")
            End If
        End If

        If wsForm.Cells(i, cItemCode).Value2 = "" Then
            zH = errFunction(zH, zV, i, "Код узла не заполнен"): flagItem = False
        Else
            MassItem = Split(Application.Trim(Replace(Replace(Replace(wsForm.Cells(i, cItemCode).Value2, ";", ","), ",", ", "), ",", " ,")), ",")    'узел
            For k = LBound(MassItem) To UBound(MassItem)
                MassItem(k) = Application.Trim(MassItem(k))
            Next
            For k = LBound(MassItem) To UBound(MassItem)
                If Not IsNumeric(MassItem(k)) Then
                    zH = errFunction(zH, zV, i, "Код узла " & MassItem(k) & " указан в текстовом формате"): flagItem = False
                ElseIf MatchDuplicatesArray(MassItem, k) = True Then
                    zH = errFunction(zH, zV, i, "Код узла " & MassItem(k) & " дублируется. Удалите дубли"): flagItem = False
                ElseIf cnORA.State = 1 Then
                    'strQry = "select TROBID from CS.TRA_RESOBJ where TROBID in (" & str & ")"
                    strQry = "select RELID from CS.RESREL where RELID in (" & MassItem(k) & ")"
                    rsOra.Open strQry, cnORA
                    If rsOra.bof Then
                        zH = errFunction(zH, zV, i, "Код узла " & MassItem(k) & " не найден в базе данных"): flagItem = False
                    ElseIf Len(TypePost) > 0 Then
                        If rsOra!RELID = "94035" Then
                            If Left(TypePost, 1) <> "3" Then zH = errFunction(zH, zV, i, "Для ОПТа указанный поставщик некорректен")
                            If Len(wsForm.Cells(i, cNewOptions + 5).Value2) > 0 Then
                                If Not wsForm.Cells(i, cNewOptions + 5).Value2 Like "CL*" Then zH = errFunction(zH, zV, i, "Для ОПТа указанный контракт некорректен")
                            End If
                        ElseIf Left(TypePost, 1) = "2" And IsError(Application.Match(CStr(MassItem(k)), Array("94072", "94015", "94009", "94010", "94011", "94076"), 0)) Then
                            zH = errFunction(zH, zV, i, "Для транзитного поставщика узел указан неверно")
                            flagItem = False
                        ElseIf Left(TypePost, 1) = "1" Then
                            If Right(TypePost, Len(TypePost) - 1) = "643" And CStr(MassItem(k)) = "94015" Then
                                zH = errFunction(zH, zV, i, "Некорректно указан узел для российского поставщика")
                                flagItem = False
                            ElseIf Right(TypePost, Len(TypePost) - 1) <> "643" And CStr(MassItem(k)) <> "94015" Then
                                zH = errFunction(zH, zV, i, "Для импортного поставщика должен быть указан узел таможенного склада")
                                flagItem = False
                            End If

                            If wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Then
                                rsOra.Close
                                strQry = "with parents as (select relpere, level As l from cs.resrel where reldfin >= trunc(sysdate) start with relid  = '" & MassItem(k) & "' connect by prior relpere = relid union all select " & MassItem(k) & ", 0 from dual) select relpere from parents where l = (select max(l) from parents)-1"
                                rsOra.Open strQry, cnORA
                                If Not rsOra.bof Then
                                    If rsOra!relpere <> "94001" And DictTer.exists(CStr(MassItem(k))) = False Then
                                        zH = errFunction(zH, zV, i, "Код узла " & CStr(MassItem(k)) & " указан некорректно")
                                        flagItem = False
                                    Else
                                        If rsOra!relpere = "94001" Then flagOsnPostMag = True
                                    End If
                                End If
                                rsOra.Close
                            End If
                            'Проверка ЛЕ если поставщик внешний и узел магазины
                            If wsForm.Cells(i, cNewOptions + 3).Value2 <> "" And wsForm.Cells(i, cNewOptions + 3).Value2 <> "41" And wsForm.Cells(i, cNewOptions + 3).Value2 <> "21" Then
                                If rsOra.State = 1 Then rsOra.Close
                                strQry = "with parents as (select relpere, level As l from cs.resrel where reldfin >= trunc(sysdate) start with relid  = '" & MassItem(k) & "' connect by prior relpere = relid union all select " & MassItem(k) & ", 0 from dual) select relpere from parents where l = (select max(l) from parents)-1"
                                rsOra.Open strQry, cnORA
                                If Not rsOra.bof Then
                                    If rsOra!relpere = "94001" Then
                                        zH = errFunction(zH, zV, i, "Значение ЛЕ должно быть коробка - ""41"" или вн.упаковка - ""21""")
                                        flagNewLE = False
                                    End If
                                End If
                                rsOra.Close
                            End If
                        ElseIf Left(TypePost, 1) = "3" Then
                            If CStr(MassItem(k)) = "70995" And wsForm.Cells(i, cNewOptions + 5).Value2 <> "" And wsForm.Cells(i, cSupplierCode).Value2 <> (Right(wsForm.Cells(i, cNewOptions + 5).Value2, 5)) Then
                                zH = errFunction(zH, zV, i, "Для виртуального сайта номер контракта должен быть равен номеру поставщика")
                                flagItem = False
                            ElseIf DictTer.exists(CStr(wsForm.Cells(i, cSupplierCode).Value2)) Then
                                If CStr(wsForm.Cells(i, cSupplierCode).Value2) = "70007" And CStr(MassItem(k)) = "94007" And DictRuchKK.exists(CStr(wsForm.Cells(i, cNewOptions + 5).Value2)) Then

                                ElseIf (IsError(Application.Match(wsForm.Cells(i, cSupplierCode).Value2, Array("70007", "70011", "70081", "70035", "70005"), 0)) = False And (CStr(MassItem(k)) = "94006" Or CStr(MassItem(k)) = "94007")) Or _
                                        (CStr(wsForm.Cells(i, cSupplierCode).Value2) = "70003" And (CStr(MassItem(k)) = "94005" Or CStr(MassItem(k)) = "94007")) Or _
                                        (CStr(wsForm.Cells(i, cSupplierCode).Value2) = "70034" And (CStr(MassItem(k)) = "94006" Or CStr(MassItem(k)) = "94005")) Then
                                    zH = errFunction(zH, zV, i, "Поставщик и узел ФО не соответствуют друг другу")
                                    flagItem = False
                                End If

                                If IsError(Application.Match(CStr(MassItem(k)), Array("70007", "70011", "70081", "70035", "70005", "70003", "70034", "94003", "94009", "94010", "94011"), 0)) = False Then
                                    zH = errFunction(zH, zV, i, "Код узла " & CStr(MassItem(k)) & " не может быть указан для внутреннего поставщика")
                                    flagItem = False
                                End If
                            End If
                        End If
                    End If
                    If rsOra.State = 1 Then rsOra.Close
                    If cnORA.State = 1 Then
                        strQry = "select ARTCEXR as tovar from CS.ARTRAC where ARTCEXR = '" & wsForm.Cells(i, cProductCode).Value2 & "' and artcinr in (select objcint from cs.strucrel where objdfin >= trunc(sysdate) start with objpere = '261640' connect by prior objcint = objpere)"
                        rsOra.Open strQry, cnORA
                        'ActiveCell = strQry
                        If Not rsOra.bof Then
                            flagTMC = True
                            If IsError(Application.Match(CStr(MassItem(k)), Array("94007", "94006", "94005", "94004", "70995", "94035"), 0)) And Left(TypePost, 1) = "3" Then
                                zH = errFunction(zH, zV, i, "Для товара ТМЦ узел " & MassItem(k) & " указан неверно")
                                flagItem = False: flagOF = False
                            End If
                        End If
                    End If
                    rsOra.Close
                End If
                'rsOra.Close
            Next k
        End If

        a = 0
        If wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки" Then
            DictSmena.RemoveAll
            SmenaItem = Replace(Replace(Replace(Trim(wsForm.Cells(i, cItemCode).Value2), ";", ","), ", ", ","), " ,", ",")          'узел
            SmenaItemRev = ""
            For j = firstRow To lastRow
                If wsForm.Cells(j, cProductCode).Value2 <> "" And wsForm.Cells(j, cItemCode).Value2 <> "" Then
                    'If j <> i Then
                    If wsForm.Cells(i, cProductCode).Value2 = wsForm.Cells(j, cProductCode).Value2 And _
                            wsForm.Cells(i, cActionType).Value2 = wsForm.Cells(j, cActionType).Value2 Then
                        SmenaItemRev = Replace(Replace(Replace(Trim(wsForm.Cells(j, cItemCode).Value2), ";", ","), ", ", ","), " ,", ",")             'узел
                        If SmenaItem <> SmenaItemRev Then
                            If DictSmena.exists(CStr(wsForm.Cells(j, cProductCode).Value2)) = False Then
                                DictSmena.Add CStr(wsForm.Cells(j, cProductCode).Value2), CreateCollection
                                DictSmena(CStr(wsForm.Cells(j, cProductCode).Value2)).Add SmenaItemRev & "_" & j
                            Else
                                DictSmena(CStr(wsForm.Cells(j, cProductCode).Value2)).Add SmenaItemRev & "_" & j
                            End If
                        Else
                            If DictSmena.exists(CStr(wsForm.Cells(j, cProductCode).Value2)) = False Then
                                DictSmena.Add CStr(wsForm.Cells(j, cProductCode).Value2), CreateCollection
                                DictSmena(CStr(wsForm.Cells(j, cProductCode).Value2)).Add SmenaItemRev & "_" & j
                            Else
                                DictSmena(CStr(wsForm.Cells(j, cProductCode).Value2)).Add SmenaItemRev & "_" & j
                            End If
                            a = a + 1
                        End If
                    End If
                    'End If
                End If
            Next
        End If

        If wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки" Then
            SmflagOld = False: SmflagNew = False: SmReversOld = False: SmReversNew = False
            SmYzel = "": SmStr = "": StrRow = ""
            For j = 0 To 8
                If Len(wsForm.Cells(i, cOldOptions + j)) > 0 Then SmflagOld = True
            Next

            For j = 0 To 9
                If Len(wsForm.Cells(i, cNewOptions + j)) > 0 Then SmflagNew = True
            Next

            If (SmflagOld = False And SmflagNew = False) Or (SmflagOld And SmflagNew) Then
                zH = errFunction(zH, zV, i, "Для указанного типа действия должен быть заполнен либо блок для закрытия, либо блок для открытия")
            ElseIf SmflagOld Or SmflagNew Then
                If DictSmena.exists(CStr(wsForm.Cells(i, cProductCode).Value2)) Then
                    If a >= 2 Then
                        SmenaItem = Replace(Replace(Replace(Trim(wsForm.Cells(i, cItemCode).Value2), ";", ","), ", ", ","), " ,", ",")     'узел
                        OtherSmFlag = False
                        For Each Item In DictSmena(CStr(wsForm.Cells(i, cProductCode).Value2))
                            On Error Resume Next
                            SmYzel = "" & Split(Trim(Item), "_")(0)
                            SmStr = "" & Split(Trim(Item), "_")(1)
                            On Error GoTo 0
                            If SmStr <> i Then
                                If SmYzel = SmenaItem Then
                                    OtherSmFlag = True
                                    If SmflagOld Then
                                        For j = 0 To 9
                                            If Len(wsForm.Cells(SmStr, cNewOptions + j)) > 0 Then SmReversNew = True
                                        Next
                                    ElseIf SmflagNew Then
                                        For j = 0 To 8
                                            If Len(wsForm.Cells(SmStr, cOldOptions + j)) > 0 Then SmReversOld = True
                                        Next
                                    End If
                                End If
                            End If
                        Next

                        falgYesRevers = 0
                        If SmReversNew = False And SmReversOld = False Then
                            For Each SmTr In Split(SmenaItem, ",")
                                flagFO = False
                                For Each Item In DictSmena(CStr(wsForm.Cells(i, cProductCode).Value2))
                                    On Error Resume Next
                                    SmYzel = "" & Split(Trim(Item), "_")(0)
                                    SmStr = "" & Split(Trim(Item), "_")(1)
                                    On Error GoTo 0
                                    If SmStr <> i Then
                                        For Each Tr In Split(SmYzel, ",")
                                            If DictFO.exists(Tr) And DictFO.exists(SmTr) Then
                                                If DictFO(Tr) = DictFO(SmTr) Then
                                                    flagFO = True
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                    If flagFO Then Exit For
                                Next

                                If flagFO = False Then StrRow = StrRow & SmStr & ",": falgYesRevers = 1
                            Next

                            If falgYesRevers = 1 Then
                                StrRow = Left(StrRow, Len(StrRow) - 1)
                                zH = errFunction(zH, zV, i, "Проверьте узлы комм.сети. Обнаружены разные ФО по строкам: " & StrRow)
                            End If
                        End If

                        If SmReversNew = False And SmReversOld = False And OtherSmFlag Then
                            If SmflagOld Then
                                zH = errFunction(zH, zV, i, "Для указанных товара и территорий не найдена строка на открытие")
                            ElseIf SmflagNew Then
                                zH = errFunction(zH, zV, i, "Для указанных товара и территорий не найдена строка на закрытие")
                            End If
                        End If
                    Else     'если по узлу только одна строка, то надо убедиться, что есть две строки с другими узлами
                        Set OtherYzel = CreateObject("Scripting.Dictionary"): f = 1
                        SmenaItem = Replace(Replace(Replace(Trim(wsForm.Cells(i, cItemCode).Value2), ";", ","), ", ", ","), " ,", ",")     'узел
                        For Each Item In DictSmena(CStr(wsForm.Cells(i, cProductCode).Value2))
                            On Error Resume Next
                            SmYzel = "" & Split(Trim(Item), "_")(0)
                            SmStr = "" & Split(Trim(Item), "_")(1)
                            On Error GoTo 0
                            If OtherYzel.exists(SmYzel) = False Then
                                OtherYzel.Add SmYzel, 1
                            Else
                                f = f + 1
                            End If
                        Next
                        OtherYzel.RemoveAll
                        If f > 1 Then     'если есть какие-то строки с одинаковыми узлами, то проверяем, что наш узел относится к тому же ФО
                            falgYesRevers = 0
                            For Each SmTr In Split(SmenaItem, ",")
                                flagFO = False
                                For Each Item In DictSmena(CStr(wsForm.Cells(i, cProductCode).Value2))
                                    On Error Resume Next
                                    SmYzel = "" & Split(Trim(Item), "_")(0)
                                    SmStr = "" & Split(Trim(Item), "_")(1)
                                    On Error GoTo 0
                                    If SmStr <> i Then
                                        For Each Tr In Split(SmYzel, ",")
                                            If DictFO.exists(Tr) And DictFO.exists(SmTr) Then
                                                If DictFO(Tr) = DictFO(SmTr) Then
                                                    flagFO = True
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                    If flagFO Then Exit For
                                Next

                                If flagFO = False Then StrRow = StrRow & SmStr & ",": falgYesRevers = 1
                            Next
                            If falgYesRevers = 1 Then
                                StrRow = Left(StrRow, Len(StrRow) - 1)
                                zH = errFunction(zH, zV, i, "Проверьте узлы комм.сети. Обнаружены разные ФО по строкам: " & StrRow)
                            End If
                        Else
                            If SmflagOld Then
                                zH = errFunction(zH, zV, i, "Для указанных товара и территорий не найдена строка на открытие")
                            ElseIf SmflagNew Then
                                zH = errFunction(zH, zV, i, "Для указанных товара и территорий не найдена строка на закрытие")
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If (wsForm.Cells(i, cNewOptions + 2).Value2 <> "") Then
            If Not IsNumeric(wsForm.Cells(i, cNewOptions + 2).Value2) Then
                zH = errFunction(zH, zV, i, "ЛВ новой настройки указан некорректно")
                flagNewLV = False
            ElseIf CStr(Fix(wsForm.Cells(i, cNewOptions + 2).Value2)) <> wsForm.Cells(i, cNewOptions + 2).Value2 Then
                zH = errFunction(zH, zV, i, "ЛВ новой настройки указан некорректно")
                flagNewLV = False
            End If
        Else
            flagNewLV = False
        End If
        If (wsForm.Cells(i, cOldOptions + 1).Value2 <> "") And wsForm.Cells(i, cActionType).Value2 <> "Открытие" And wsForm.Cells(i, cActionType).Value2 <> "Опт: смена поставщика для авторасценки" Then
            If Not IsNumeric(wsForm.Cells(i, cOldOptions + 1).Value2) Then
                zH = errFunction(zH, zV, i, "ЛВ старой настройки указан некорректно")
                flagOldLV = False
            ElseIf CStr(Fix(wsForm.Cells(i, cOldOptions + 1).Value2)) <> wsForm.Cells(i, cOldOptions + 1).Value2 Then
                zH = errFunction(zH, zV, i, "ЛВ старой настройки указан некорректно")
                flagOldLV = False
            End If
        Else
            flagOldLV = False
        End If

        If wsForm.Cells(i, cNewOptions + 2).Value2 = "" And (wsForm.Cells(i, cActionType).Value2 = "Открытие" Or wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Or _
                ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew)) Then zH = errFunction(zH, zV, i, "ЛВ новой настройки не заполнен")

        If flagNewLV And cnORA.State = 1 And flagProd Then
            strQry = "SELECT ARLETAT FROM CS.ARTVL WHERE CS.ARTVL.ARLCEXR = '" & wsForm.Cells(i, cProductCode).Value2 & "' and CS.ARTVL.ARLCEXVL = '" & wsForm.Cells(i, cNewOptions + 2).Value2 & "'"
            rsOra.Open strQry, cnORA
            If rsOra.bof Then
                zH = errFunction(zH, zV, i, "Указанный номер ЛВ не существует в карточке товара " & wsForm.Cells(i, cProductCode).Value2)
                flagAll = False
            Else
                If rsOra!ARLETAT = 2 Then
                    zH = errFunction(zH, zV, i, "Указанный номер ЛВ в карточке товара " & wsForm.Cells(i, cProductCode).Value2 & " заморожен")
                    flagAll = False
                End If
            End If
            rsOra.Close
        End If

        If (wsForm.Cells(i, cNewOptions + 3).Value2 <> "") Then
            If Not IsNumeric(wsForm.Cells(i, cNewOptions + 3).Value2) Then
                zH = errFunction(zH, zV, i, "Код ЛЕ новой настройки указан некорректно")
                flagNewLE = False
            ElseIf CStr(Fix(wsForm.Cells(i, cNewOptions + 3).Value2)) <> wsForm.Cells(i, cNewOptions + 3).Value2 Then
                zH = errFunction(zH, zV, i, "Код ЛЕ новой настройки указан некорректно")
                flagNewLE = False
            ElseIf flagProd And wsForm.Cells(i, cNewOptions + 3).Value2 <> "41" And wsForm.Cells(i, cNewOptions + 3).Value2 <> "1" And wsForm.Cells(i, cNewOptions + 3).Value2 <> "81" And wsForm.Cells(i, cNewOptions + 3).Value2 <> "21" Then              'And cnORA.State = 1
                zH = errFunction(zH, zV, i, "Код ЛЕ новой настройки указан не из выпадающего списка")
                flagNewLE = False
            ElseIf Len(wsForm.Cells(i, cNewOptions + 2).Value2) > 0 And flagAll Then
                strQry = "SELECT ARUCINL  FROM CS.ARTUL inner join cs.artvl on ARUSEQVL = ARLSEQVL WHERE arlcexvl = '" & CStr(wsForm.Cells(i, cNewOptions + 2).Value2) & "' and ARUTYPUL = '" & CStr(wsForm.Cells(i, cNewOptions + 3).Value2) & "' AND ARUCINR = (SELECT ARTCINR FROM CS.ARTRAC WHERE ARTCEXR = '" & CStr(wsForm.Cells(i, cProductCode).Value2) & "')"
                rsOra.Open strQry, cnORA
                If rsOra.bof And CStr(wsForm.Cells(i, cNewOptions + 2).Value2) <> "0" And CStr(wsForm.Cells(i, cNewOptions + 3).Value2) <> "1" Then
                    zH = errFunction(zH, zV, i, "В карточке нового ЛВ нет указанной ЛЕ"): flagNewLE = False
                End If
                rsOra.Close
            End If
        End If

        If wsForm.Cells(i, cOldOptions).Value2 <> "" And wsForm.Cells(i, cActionType).Value2 <> "Открытие" And wsForm.Cells(i, cActionType).Value2 <> "Опт: смена поставщика для авторасценки" Then
            If IsDate(wsForm.Cells(i, cOldOptions).Value) Then
                If wsForm.Cells(i, cActionType).Value2 = "Изменение" Then
                    If wsForm.Cells(i, cItemCode).Value2 <> "" Then
                        For k = LBound(MassItem) To UBound(MassItem)
                            If CDate(wsForm.Cells(i, cOldOptions).Value2) <> Date - 1 And Left(TypePost, 1) = "1" And (IsError(Application.Match(CStr(MassItem(k)), Array("94015", "94009", "94010", "94011", "94003", "70081", "70035", "70005", "70003", "70007", "70011", "70034"), 0)) Or (wsForm.Cells(i, cNewOptions + 6).Value2 <> "" Or wsForm.Cells(i, cNewOptions + 7).Value2 <> "" Or wsForm.Cells(i, cNewOptions + 8).Value2 <> "" Or wsForm.Cells(i, cNewOptions + 9).Value2 <> "")) Then
                                If CDate(wsForm.Cells(i, cOldOptions).Value2) < Date Then
                                    zH = errFunction(zH, zV, i, "Кон.дата старой настройки должна быть больше текущей")
                                    flagOD = False
                                ElseIf CDate(wsForm.Cells(i, cOldOptions).Value2) = Date And CDate(wsForm.Cells(i, cNewOptions).Value2) < Date Then
                                    wsForm.Cells(i, cOldOptions).Value2 = Date + 1
                                    zH = errFunction(zH, zV, i, "Кон.дата старой настройки текущая, заменена на завтрашнюю")
                                End If
                            ElseIf CDate(wsForm.Cells(i, cOldOptions).Value2) = Date - 1 Then
                                wsForm.Cells(i, cOldOptions).Value2 = Date
                                zH = errFunction(zH, zV, i, "Кон.дата старой настройки вчерашняя, заменена на текущую")
                            ElseIf CDate(wsForm.Cells(i, cOldOptions).Value2) < Date - 1 Then
                                zH = errFunction(zH, zV, i, "Кон.дата старой настройки должна быть не меньше текущей")
                                flagOD = False
                            End If
                        Next
                    End If
                ElseIf wsForm.Cells(i, cActionType).Value2 = "Закрытие/Удаление" Or _
                        ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagOld) Then
                    If CDate(wsForm.Cells(i, cOldOptions).Value2) < Date - 1 Then
                        zH = errFunction(zH, zV, i, "Кон.дата старой настройки должна быть не меньше текущей")
                        flagOD = False
                    ElseIf CDate(wsForm.Cells(i, cOldOptions).Value2) = Date - 1 Then
                        wsForm.Cells(i, cOldOptions).Value2 = Date
                        zH = errFunction(zH, zV, i, "Кон.дата старой настройки вчерашняя, заменена на текущую")
                    End If
                End If
            Else
                zH = errFunction(zH, zV, i, "Кон.дата старой настройки указана в некорректном формате"): flagOD = False
            End If
        End If

        If flagSup And flagProd And cnORA.State = 1 And Left(TypePost, 1) = "1" Then
            strQry = "select ARTNATU from cs.artrac where ARTCEXR = '" & wsForm.Cells(i, cProductCode).Value2 & "' and ARTNATU = '6'"
            rsOra.Open strQry, cnORA
            If Not rsOra.bof Then
                rsOra.Close
                strQry = "SELECT FCLCFIN FROM CS.FOUCATALOG WHERE FCLDFIN > SYSDATE AND CS.FOUCATALOG.FCLCFIN = (SELECT FOUCFIN FROM CS.FOUDGENE WHERE FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value2 & "') AND CS.FOUCATALOG.FCLSEQVL in (SELECT ARLSEQVL FROM CS.ARTVL WHERE ARLCEXR = '" & wsForm.Cells(i, cProductCode).Value2 & "')"
                rsOra.Open strQry, cnORA
                If rsOra.bof Then zH = errFunction(zH, zV, i, "Товар комиссионный. Указанный поставщик не пристыкован в карточку товара")
                rsOra.Close
            Else
                rsOra.Close
            End If
        End If
        '-----------------------------------------------------новые настройки-----------------------------------------------------

        If wsForm.Cells(i, cActionType).Value2 = "Открытие" Or wsForm.Cells(i, cActionType).Value2 = "Изменение" Or wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Or _
                ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew) Then
            If wsForm.Cells(i, cNewOptions).Value2 = "" Then
                zH = errFunction(zH, zV, i, "Нач.дата новой настройки не заполнена"): flagND = False
            ElseIf IsDate(wsForm.Cells(i, cNewOptions).Value) Then
                If wsForm.Cells(i, cItemCode).Value2 <> "" Then
                    For k = LBound(MassItem) To UBound(MassItem)
                        If flagND Then
                            If CDate(wsForm.Cells(i, cNewOptions).Value2) <> Date - 1 And Left(TypePost, 1) = "1" And (IsError(Application.Match(CStr(MassItem(k)), Array("94015", "94009", "94010", "94011", "94003", "70003", "70081", "70035", "70005", "70007", "70011", "70034"), 0)) Or (wsForm.Cells(i, cNewOptions + 6).Value2 <> "" Or wsForm.Cells(i, cNewOptions + 7).Value2 <> "" Or wsForm.Cells(i, cNewOptions + 8).Value2 <> "" Or wsForm.Cells(i, cNewOptions + 9).Value2 <> "")) Then
                                If CDate(wsForm.Cells(i, cNewOptions).Value2) < Date Then
                                    zH = errFunction(zH, zV, i, "Нач.дата новой настройки должна быть больше текущей")
                                    flagND = False
                                ElseIf CDate(wsForm.Cells(i, cNewOptions).Value2) = Date And wsForm.Cells(i, cActionType).Value2 <> "Открытие" Then    'And wsForm.Cells(i, cActionType).Value2 <> "Опт: смена поставщика для авторасценки"
                                    wsForm.Cells(i, cNewOptions).Value2 = Date + 1
                                    zH = errFunction(zH, zV, i, "Нач.дата новой настройки текущая, заменена на завтрашнюю")
                                End If
                            ElseIf CDate(wsForm.Cells(i, cNewOptions).Value2) = Date - 1 Then
                                If wsForm.Cells(i, cNewOptions + 1).Value <> "" And IsDate(wsForm.Cells(i, cNewOptions + 1).Value) Then
                                    If CDate(wsForm.Cells(i, cNewOptions).Value2) = CDate(wsForm.Cells(i, cNewOptions + 1).Value2) Then
                                        wsForm.Cells(i, cNewOptions + 1).Value2 = Date
                                        zH = errFunction(zH, zV, i, "Кон.дата новой настройки вчерашняя, заменена на текущую")
                                    End If
                                End If
                                wsForm.Cells(i, cNewOptions).Value2 = Date
                                zH = errFunction(zH, zV, i, "Нач.дата новой настройки вчерашняя, заменена на текущую")
                            ElseIf CDate(wsForm.Cells(i, cNewOptions).Value2) < Date - 1 Then
                                zH = errFunction(zH, zV, i, "Нач.дата новой настройки должна быть не меньше текущей")
                                flagND = False
                            End If
                        End If
                    Next
                End If
                If wsForm.Cells(i, cNewOptions + 1).Value = "" Then
                    If wsForm.Cells(i, cActionType).Value2 = "Открытие" Or wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Or _
                            ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew) Then
                        zH = errFunction(zH, zV, i, "Кон.дата новой настройки не заполнена")
                        flagND = False
                    End If
                Else
                    If IsDate(wsForm.Cells(i, cNewOptions + 1).Value) Then
                        If CDate(wsForm.Cells(i, cNewOptions + 1).Value2) < CDate(wsForm.Cells(i, cNewOptions).Value2) And flagND Then
                            zH = errFunction(zH, zV, i, "Кон.дата новой настройки не может быть меньше начальной"): flagND = False
                        End If
                    Else
                        zH = errFunction(zH, zV, i, "Кон.дата новой настройки указана в некорректном формате"): flagND = False
                    End If
                End If
                If wsForm.Cells(i, cOldOptions).Value2 <> "" And IsDate(wsForm.Cells(i, cOldOptions).Value) And wsForm.Cells(i, cAddOptions).Value2 = "" And flagND And flagOD Then
                    If CDate(wsForm.Cells(i, cOldOptions).Value2) >= Date And CDate(wsForm.Cells(i, cOldOptions).Value2) < CDate(wsForm.Cells(i, cNewOptions).Value2) And CDate(wsForm.Cells(i, cOldOptions).Value2) <> CDate(wsForm.Cells(i, cNewOptions).Value2) - 1 And CDate(wsForm.Cells(i, cOldOptions).Value2) <> CDate(wsForm.Cells(i, cNewOptions).Value2) Then
                        zH = errFunction(zH, zV, i, "Между кон.датой старой настройки и нач.датой новой настройки не должен быть разрыв дат больше, чем на один день")
                    End If
                    If CDate(wsForm.Cells(i, cOldOptions).Value2) > CDate(wsForm.Cells(i, cNewOptions).Value) And wsForm.Cells(i, cAddOptions).Value2 = "" Then
                        zH = errFunction(zH, zV, i, "Кон.дата старой настройки должна быть меньше нач.даты новой настройки")
                        flagOD = False
                    End If
                End If
            Else
                zH = errFunction(zH, zV, i, "Нач.дата новой настройки указана в некорректном формате"): flagND = False
            End If

            If flagND And flagOD Then
                If wsForm.Cells(i, cOldOptions).Value2 <> "" Then
                    If IsDate(wsForm.Cells(i, cOldOptions).Value2) Then
                        If CDate(wsForm.Cells(i, cOldOptions).Value2) >= Date + 1 And CDate(wsForm.Cells(i, cOldOptions).Value2) = CDate(wsForm.Cells(i, cNewOptions).Value) Then
                            zH = errFunction(zH, zV, i, "Кон.дата старой настройки равна нач.дате новой настройки. Необходимо скорректировать какую-либо из этих дат")
                            flagOD = False: flagND = False
                        End If
                    End If
                End If
            End If
            'проверка матричности
            strTer = "": strTerMags = "": YzelMags = ""
            If Len(wsForm.Cells(i, cItemCode).Value2) > 0 And (wsForm.Cells(i, cActionType).Value2 = "Открытие" Or _
                    ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew)) And flagProd And flagSup And flagND And flagTMC = False Then
                If CDate(wsForm.Cells(i, cNewOptions + 1).Value2) <> CDate(wsForm.Cells(i, cNewOptions).Value2) Then
                    'если поставщик внешний
                    If Left(TypePost, 1) = "1" Then
                        For k = LBound(MassItem) To UBound(MassItem)
                            If CStr(MassItem(k)) <> "70004" Then
                                If flagAM Then
                                    If MassItem(k) = "94015" Then
                                        strQry = "select AATCINR from CS.ARTATTRI inner join cs.artrac on AATCINR=artcinr where AATDFIN>trunc(sysdate) and AATCCLA='AS' and artcexr='" & wsForm.Cells(i, cProductCode).Value2 & "' "
                                        rsOra.Open strQry, cnORA
                                        If rsOra.bof Then zH = errFunction(zH, zV, i, "Товар не матричный"): flagAM = False
                                        rsOra.Close
                                    Else
                                        strQry = "with parents as (select relpere, level As l from cs.resrel where reldfin >= trunc(sysdate) start with relid  = '" & MassItem(k) & "' connect by prior relpere = relid union all select " & MassItem(k) & ", 0 from dual) select relpere from parents where l = (select max(l) from parents)-1"
                                        rsOra.Open strQry, cnORA
                                        If Not rsOra.bof Then YzelMags = rsOra!relpere
                                        rsOra.Close
                                        If YzelMags = "94001" Then     'проверка матричности от внешнего на маг
                                            strQry = "select AATCINR from CS.ARTATTRI inner join cs.artrac on AATCINR=artcinr where AATDFIN>trunc(sysdate) and artcexr='" & wsForm.Cells(i, cProductCode).Value2 & "' " & _
                                                    "and AATCCLA='AS' and substr(aatcatt, 1, 5) in (select relid from cs.resrel where reldfin >= sysdate start with relpere  = '" & Application.Trim(MassItem(k)) & "' connect by prior relid = relpere Union all select TO_NUMBER('" & Application.Trim(MassItem(k)) & "') from dual)"
                                            rsOra.Open strQry, cnORA
                                            If rsOra.bof Then strTerMags = strTerMags & CStr(MassItem(k)) & ","
                                            rsOra.Close
                                        ElseIf DictTer.exists(CStr(MassItem(k))) And CStr(MassItem(k)) Like "70*" Then      'если узел=РЦ
                                            'проверка для РЦ
                                            strQry = "select AATCINR from CS.ARTATTRI inner join cs.artrac on AATCINR=artcinr where AATDFIN>trunc(sysdate) and artcexr='" & wsForm.Cells(i, cProductCode).Value2 & "' " & _
                                                    "and AATCCLA='AS' and substr(aatcatt, 1, 5) in (select satsite from cs.sitattri where satcla='WH' and satdfin>trunc(sysdate) and '70'||SUBSTR(satatt, 1, 3)=" & CStr(MassItem(k)) & " and SUBSTR(satatt, 4, 2) = '_1') and rownum=1"
                                            rsOra.Open strQry, cnORA
                                            If rsOra.bof Then strTer = strTer & CStr(MassItem(k)) & ","
                                            rsOra.Close
                                        Else     'если указаны узлы в которые входят РЦ
                                            If DictMatrRc.exists(CStr(MassItem(k))) Then
                                                For Each Item In DictMatrRc(CStr(MassItem(k)))
                                                    strQry = "select AATCINR from CS.ARTATTRI inner join cs.artrac on AATCINR=artcinr where AATDFIN>trunc(sysdate) and artcexr='" & wsForm.Cells(i, cProductCode).Value2 & "' " & _
                                                            "and AATCCLA='AS' and substr(aatcatt, 1, 5) in (select satsite from cs.sitattri where satcla='WH' and satdfin>trunc(sysdate) and '70'||SUBSTR(satatt, 1, 3)=" & Item & " and SUBSTR(satatt, 4, 2) = '_1') and rownum=1"
                                                    rsOra.Open strQry, cnORA
                                                    If rsOra.bof Then strTer = strTer & Item & ","
                                                    rsOra.Close
                                                Next
                                            End If
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            End If
                        Next

                        If strTer <> "" Then
                            strTer = Left(strTer, Len(strTer) - 1)
                            zH = errFunction(zH, zV, i, " Для РЦ " & strTer & " по товару нет магазинов в матрице")
                        ElseIf strTerMags <> "" Then
                            strTerMags = Left(strTerMags, Len(strTerMags) - 1)
                            zH = errFunction(zH, zV, i, " Для узлов " & strTerMags & " по товару нет магазинов в матрице")
                        End If
                    End If
                ElseIf CDate(wsForm.Cells(i, cNewOptions + 1).Value2) = CDate(wsForm.Cells(i, cNewOptions).Value2) Then
                    'если поставщик внешний
                    If Left(TypePost, 1) = "1" Then zH = errFunction(zH, zV, i, "Период открытия равен одному дню. Заказ товара будет невозможен")
                End If
            End If
            '-------------------------Проверки для типа Опт: смена поставщика для авторасценки---------------------------------------------------
            If flagProd And flagSup And flagND And flagItem And flagAll And flagNewLV And wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Then
                If cnORA.State = 1 Then
                    MassItem = Split(Application.Trim(Replace(Replace(Replace(wsForm.Cells(i, cItemCode).Value2, ";", ","), ",", ", "), ",", " ,")), ",")    'узел
                    For k = LBound(MassItem) To UBound(MassItem)
                        If Application.Trim(MassItem(k)) Like "700*" Then
                            strQry = "SELECT ARACEXR,ARACEXVL,foucnuf,ARASITE FROM cs.ARTUC inner join cs.foudgene on aracfin=FOUCFIN where ARACEXR = '" & wsForm.Cells(i, cProductCode).Value2 & "' and ARACEXVL = '" & wsForm.Cells(i, cNewOptions + 2).Value2 & "' and foucnuf = '" & wsForm.Cells(i, cSupplierCode).Value2 & "' and ARASITE = '" & Application.Trim(MassItem(k)) & "' and aradfin > trunc(sysdate)"
                            rsOra.Open strQry, cnORA
                            If rsOra.bof Then
                                flagOsnPostRC = True
                                rsOra.Close
                                Exit For
                            End If
                            rsOra.Close
                        End If
                    Next
                End If
            End If
            If wsForm.Cells(i, cOldOptions + 4).Value2 = "" And wsForm.Cells(i, cAddOptions).Value2 <> "да" And ((wsForm.Cells(i, cActionType).Value2 = "Изменение" And Left(TypePost, 1) = "3") Or _
                    (wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" And flagOsnPostMag)) Then
                zH = errFunction(zH, zV, i, "Контракт старой настройки не заполнен")
            End If
            If wsForm.Cells(i, cNewOptions + 5).Value2 = "" Then
                If wsForm.Cells(i, cActionType).Value2 = "Открытие" Or flagOsnPostRC Or flagOsnPostMag Or _
                        ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew) Then
                    zH = errFunction(zH, zV, i, "Контракт новой настройки не заполнен"): flagKK = False
                End If
            Else
                If cnORA.State = 1 Then
                    strQry = "SELECT CS.FOUCCOM.FCCNUM FROM CS.FOUCCOM WHERE CS.FOUCCOM.FCCNUM = '" & wsForm.Cells(i, cNewOptions + 5).Value2 & "'"
                    rsOra.Open strQry, cnORA
                    If rsOra.bof Then
                        zH = errFunction(zH, zV, i, "Комм.контракт " & wsForm.Cells(i, cNewOptions + 5).Value2 & "  не существует в базе данных")
                        flagKK = False
                    End If
                    rsOra.Close
                    If flagSup And flagKK Then
                        strQry = "SELECT CS.LIENCOM.LICCFIN FROM CS.LIENCOM INNER JOIN CS.FOUDGENE ON CS.LIENCOM.LICCFIN = CS.FOUDGENE.FOUCFIN AND CS.FOUDGENE.FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value2 & "'" & _
                                "INNER JOIN CS.FOUCCOM  ON CS.LIENCOM.LICCCIN = CS.FOUCCOM.FCCCCIN AND CS.FOUCCOM.FCCNUM = '" & wsForm.Cells(i, cNewOptions + 5).Value2 & "' AND CS.FOUCCOM.FCCDFIN = TO_DATE('31.12.2049','DD.MM.YYYY') WHERE CS.LIENCOM.LICDFIN > SYSDATE"
                        rsOra.Open strQry, cnORA
                        If rsOra.bof Then
                            zH = errFunction(zH, zV, i, "Указанный КК " & wsForm.Cells(i, cNewOptions + 5).Value2 & " не принадлежит поставщику " & wsForm.Cells(i, cSupplierCode).Value2)
                            flagKK = False
                        End If
                        rsOra.Close
                        If Left(TypePost, 1) = "1" And flagKK Then
                            strQry = "SELECT DISTINCT CS.LIENREGL.LIRDFIN from CS.LIENREGL INNER JOIN CS.FOUCCOM ON CS.LIENREGL.LIRCCIN = CS.FOUCCOM.FCCCCIN AND CS.FOUCCOM.FCCNUM = '" & wsForm.Cells(i, cNewOptions + 5).Value2 & "'" & _
                                    "INNER JOIN CS.FOUDGENE ON CS.LIENREGL.LIRCFIN = CS.FOUDGENE.FOUCFIN AND CS.FOUDGENE.FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value2 & "' WHERE CS.LIENREGL.LIRDFIN > SYSDATE"
                            rsOra.Open strQry, cnORA
                            If rsOra.bof Then
                                zH = errFunction(zH, zV, i, "Указанный комм.контракт " & wsForm.Cells(i, cNewOptions + 5).Value2 & " закрыт")
                                flagKK = False
                            End If
                            rsOra.Close
                        End If
                    End If
                    If Left(TypePost, 1) <> "2" And flagKK Then
                        strQry = "SELECT CS.FOUCCOM.FCCNATC FROM CS.FOUCCOM WHERE CS.FOUCCOM.FCCNUM = '" & wsForm.Cells(i, cNewOptions + 5) & "'"
                        rsOra.Open strQry, cnORA
                        TypeKK = rsOra!FCCNATC
                        If TypeKK = "8" And flagOF Then
                            For k = LBound(MassItem) To UBound(MassItem)
                                If IsError(Application.Match(CStr(MassItem(k)), Array("94007", "94006", "94005", "94004"), 0)) Then
                                    zH = errFunction(zH, zV, i, "Для офисного контракта узел " & MassItem(k) & " указан неверно")
                                    flagItem = False
                                End If
                            Next k
                        End If
                        rsOra.Close
                    End If
                End If
            End If
            'Проверка наличия ЛЕ в карточке товара и в ЗА
            If cnORA.State = 1 Then
                If wsForm.Cells(i, cActionType).Value2 = "Изменение" Then
                    If flagNewLV And flagNewLE And Len(wsForm.Cells(i, cNewOptions + 2).Value2) > 0 Then
                        Dim DictLe As Object
                        Set DictLe = CreateObject("Scripting.Dictionary")
                        strQry = "SELECT ARUTYPUL le FROM CS.ARTUL inner join cs.artvl on ARUSEQVL = ARLSEQVL WHERE arlcexvl = '" & CStr(wsForm.Cells(i, cNewOptions + 2).Value2) & "' AND ARUCINR = (SELECT ARTCINR FROM CS.ARTRAC WHERE ARTCEXR = '" & CStr(wsForm.Cells(i, cProductCode).Value2) & "')"
                        rsOra.Open strQry, cnORA
                        If Not rsOra.bof Then
                            Do While Not rsOra.EOF
                                If DictLe.exists(rsOra.Fields("le").Value) = False Then DictLe.Add CStr(rsOra.Fields("le").Value), 1
                                rsOra.movenext
                            Loop
                        End If
                        rsOra.Close
                        If DictLe.exists("21") = False And Len(wsForm.Cells(i, cNewOptions + 3).Value2) = 0 Then
                            strQry = "SELECT distinct ARUTYPUL ZaLe FROM cs.ARTUC inner join cs.foudgene on aracfin=FOUCFIN INNER JOIN CS.ARTUL ON ARUCINL = ARACINL where ARACEXR = '" & CStr(wsForm.Cells(i, cProductCode).Value2) & "' and foucnuf = '" & CStr(wsForm.Cells(i, cSupplierCode).Value2) & "'  and aradfin > trunc(sysdate)"
                            rsOra.Open strQry, cnORA
                            If Not rsOra.bof Then
                                Do While Not rsOra.EOF
                                    If CStr(rsOra!ZaLe) = "21" Then
                                        zH = errFunction(zH, zV, i, "Обратите внимание! В указанном ЛВ нет ЛЕ 21, но в системе есть настройки на ЛЕ 21. Они не смогут быть обработанными.")
                                        Exit Do
                                    End If
                                    rsOra.movenext
                                Loop
                            End If
                            rsOra.Close
                        End If

                        If Len(wsForm.Cells(i, cNewOptions + 3).Value2) > 0 Then
                            If DictLe.exists(wsForm.Cells(i, cNewOptions + 3).Value2) = False Then
                                zH = errFunction(zH, zV, i, "ЛЕ " & wsForm.Cells(i, cNewOptions + 3).Value2 & " из блока новых настроек отсутствует в ЛВ " & wsForm.Cells(i, cNewOptions + 2).Value2 & " карточки товара")
                            End If
                        End If
                    End If
                End If
            End If
            If Left(TypePost, 1) = "3" And (wsForm.Cells(i, cActionType).Value2 = "Открытие" Or _
                    (wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки" And SmflagNew)) Then
                wsForm.Cells(i, cNewOptions + 4).Value2 = "1"
            ElseIf wsForm.Cells(i, cNewOptions + 4).Value2 = "" And Left(TypePost, 1) = "1" Then
                If wsForm.Cells(i, cActionType).Value2 = "Открытие" Or flagOsnPostRC Or _
                        ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew) Then zH = errFunction(zH, zV, i, "АЦ не заполнена")

            ElseIf wsForm.Cells(i, cNewOptions + 4).Value2 <> "" Then
                If Not IsNumeric(wsForm.Cells(i, cNewOptions + 4).Value2) Then
                    zH = errFunction(zH, zV, i, "АЦ новой настройки указана некорректно")
                ElseIf CStr(Fix(wsForm.Cells(i, cNewOptions + 4).Value2)) <> wsForm.Cells(i, cNewOptions + 4).Value2 Then
                    zH = errFunction(zH, zV, i, "АЦ новой настройки указана некорректно")
                ElseIf cnORA.State = 1 And flagSup Then
                    strQry = "select LICNFILF from CS.LIENCOM where LICCFIN = (select FOUCFIN from Cs.FOUDGENE where FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value2 & "') and LICNFILF = '" & wsForm.Cells(i, cNewOptions + 4).Value2 & "'"
                    rsOra.Open strQry, cnORA
                    If rsOra.bof Then
                        zH = errFunction(zH, zV, i, "Указанной Адр.цепочки не существует в карточке поставщика " & wsForm.Cells(i, cSupplierCode).Value2)
                        rsOra.Close
                    Else
                        rsOra.Close
                        If Left(TypePost, 1) = "1" And flagKK Then
                            strQry = "select FFINFILF from CS.FOUFILIE where FFICFIN = (select FOUCFIN from CS.FOUDGENE where FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value2 & "') and FFINFILF = '" & wsForm.Cells(i, cNewOptions + 4).Value2 & "' and FFIDFIN > trunc(sysdate)"
                            rsOra.Open strQry, cnORA
                            If rsOra.bof Then zH = errFunction(zH, zV, i, "Указанная адр.цепочка закрыта в карточке поставщика")
                            rsOra.Close
                        ElseIf Left(TypePost, 1) = "2" And flagKK Then
                            If CStr(Mid(wsForm.Cells(i, cNewOptions + 5).Value2, 6, 1)) = "0" Then
                                If wsForm.Cells(i, cNewOptions + 4).Value2 <> CInt(Right(wsForm.Cells(i, cNewOptions + 5).Value2, 2)) Then zH = errFunction(zH, zV, i, "Для транзитного поставщика АЦ должна быть равна последним двум цифрам комм.контракта")
                            ElseIf CStr(Mid(wsForm.Cells(i, cNewOptions + 5).Value2, 6, 1)) <> "0" Then
                                If wsForm.Cells(i, cNewOptions + 4).Value2 <> CInt(Right(wsForm.Cells(i, cNewOptions + 5).Value2, 3)) Then zH = errFunction(zH, zV, i, "Для транзитного поставщика АЦ должна быть равна последним трем цифрам комм.контракта")
                            End If
                        End If
                    End If
                End If
            ElseIf wsForm.Cells(i, cSupplierCode).Value2 = "DT00000" And flagKK Then
                If CStr(Mid(wsForm.Cells(i, cNewOptions + 5).Value2, 6, 1)) = "0" Then
                    wsForm.Cells(i, cNewOptions + 4).Value2 = CInt(Right(wsForm.Cells(i, cNewOptions + 5).Value2, 2))
                ElseIf CStr(Mid(wsForm.Cells(i, cNewOptions + 5).Value2, 6, 1)) <> "0" Then
                    wsForm.Cells(i, cNewOptions + 4).Value2 = CInt(Right(wsForm.Cells(i, cNewOptions + 5).Value2, 3))
                End If
            End If
            If wsForm.Cells(i, cNewOptions + 3).Value2 = "" And (wsForm.Cells(i, cActionType).Value2 = "Открытие" Or flagOsnPostRC Or _
                    ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew)) Then
                zH = errFunction(zH, zV, i, "Код ЛЕ новой настройки не заполнен")
                flagAll = False
            End If
            '----------------------------------------------
            Call AC_ALCKK_
            '-----------------------------------------------------
            '------------Опт: смена поставщика для авторасценки. Проверка если есть магазины------------------
            'ключ Товар+Поставщик+ЛВ+КК+Узел
            If flagOsnPostMag And flagProd And flagSup And flagKK And flagND And flagItem And flagAll And flagNewLV Then
                Dim NoItem: NoItem = ""
                If cnORA.State = 1 Then
                    MassItem = Split(Application.Trim(Replace(Replace(Replace(wsForm.Cells(i, cItemCode).Value2, ";", ","), ",", ", "), ",", " ,")), ",")    'узел
                    For k = LBound(MassItem) To UBound(MassItem)
                        If Not Application.Trim(MassItem(k)) Like "700*" Then
                            strQry = "SELECT ARACEXR FROM cs.ARTUC inner join cs.foudgene on aracfin=FOUCFIN INNER JOIN CS.FOUCCOM ON CS.FOUCCOM.FCCCCIN = CS.ARTUC.ARACCIN " & _
                                    "where ARACEXR = '" & wsForm.Cells(i, cProductCode).Value2 & "' and ARACEXVL = '" & wsForm.Cells(i, cNewOptions + 2).Value2 & "' and foucnuf = '" & wsForm.Cells(i, cSupplierCode).Value2 & "' and FCCNUM = '" & wsForm.Cells(i, cNewOptions + 5).Value & "' " & _
                                    " and ARASITE in (select relid from cs.resrel where reldfin >= sysdate start with relpere  = '" & Application.Trim(MassItem(k)) & "' connect by prior relid = relpere Union all select TO_NUMBER('" & Application.Trim(MassItem(k)) & "') from dual) and aradfin > to_date('" & wsForm.Cells(i, cNewOptions).Value & "','DD.MM.YYYY') "
                            rsOra.Open strQry, cnORA
                            If rsOra.bof Then
                                NoItem = NoItem & Application.Trim(MassItem(k)) & ","
                                rsOra.Close
                            End If
                            If rsOra.State = 1 Then rsOra.Close
                        End If
                    Next

                    If Len(NoItem) > 0 Then
                        NoItem = Left(NoItem, Len(NoItem) - 1)
                        zH = errFunction(zH, zV, i, "По узлам: " & NoItem & " нет данных для смены основного поставщика в указанном КК")
                    End If
                End If
            End If

            'EAN
            If wsForm.Cells(i, cNewOptions + 6).Value2 <> "" And CStr(wsForm.Cells(i, cItemCode).Value2) <> "94035" And CStr(wsForm.Cells(i, cNewOptions + 5).Value2) <> "CL_VCT" Then
                wsForm.Cells(i, cNewOptions + 6).Value2 = Trim(wsForm.Cells(i, cNewOptions + 6).Value2)
                wsForm.Cells(i, cNewOptions + 6).Value2 = Replace(wsForm.Cells(i, cNewOptions + 6).Value2, " ", "")
                If Len(wsForm.Cells(i, cNewOptions + 6).Value2) <> Len(check_str(wsForm.Cells(i, cNewOptions + 6).Value2, "#")) Or Len(wsForm.Cells(i, cNewOptions + 6).Value2) > 14 Then
                    zH = errFunction(zH, zV, i, "Указанный код EAN некорректный")
                ElseIf flagProd Then
                    If cnORA.State = 1 Then
                        strQry = "select cs.artrac.artcexr as Code from cs.artrac where cs.artrac.artcinr = (select cs.ARTCOCA.ARCCINR from cs.ARTCOCA where cs.ARTCOCA.ARCCODE = '" & wsForm.Cells(i, cNewOptions + 6).Value2 & "'  and cs.ARTCOCA.ARCDFIN > sysdate)"
                        rsOra.Open strQry, cnORA
                        If Not rsOra.bof Then
                            If wsForm.Cells(i, cProductCode).Value2 <> rsOra!code Then zH = errFunction(zH, zV, i, "Код EAN принадлежит другому товару - " & rsOra!code)
                        Else
                            zH = errFunction(zH, zV, i, "Кода EAN нет в системе")
                        End If
                        rsOra.Close
                    End If
                End If
            Else
                If IsDate(wsForm.Cells(i, cNewOptions + 1)) Then
                    If Left(TypePost, 1) = "1" And (wsForm.Cells(i, cActionType).Value2 = "Открытие" Or flagOsnPostRC Or _
                            ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagNew)) And CDate(wsForm.Cells(i, cNewOptions + 1)) > Date Then zH = errFunction(zH, zV, i, "EAN не указан")
                End If
            End If
            'ЗЕ
            If wsForm.Cells(i, cNewOptions + 7).Value2 <> "" Then
                If Not check_natur(wsForm.Cells(i, cNewOptions + 7).Value2) Then
                    zH = errFunction(zH, zV, i, "Кратность ЗЕ указана некорректно")
                ElseIf Len(wsForm.Cells(i, cNewOptions + 7).Value2) <> Len(check_str(wsForm.Cells(i, cNewOptions + 7).Value2, "#")) Then
                    zH = errFunction(zH, zV, i, "Кратность ЗЕ должна иметь числовой формат")
                End If
            End If
            If wsForm.Cells(i, cNewOptions + 8).Value2 <> "" Then
                If Not check_natur(wsForm.Cells(i, cNewOptions + 8).Value2) Then
                    zH = errFunction(zH, zV, i, "Мин.заказ в ЗЕ указан некорректно")
                ElseIf Len(wsForm.Cells(i, cNewOptions + 8).Value2) <> Len(check_str(wsForm.Cells(i, cNewOptions + 8).Value2, "#")) Then
                    zH = errFunction(zH, zV, i, "Мин.заказ в ЗЕ должен иметь числовой формат")
                End If
            End If
            If wsForm.Cells(i, cNewOptions + 9).Value2 <> "" Then
                If Not check_natur(wsForm.Cells(i, cNewOptions + 9).Value2) Then
                    zH = errFunction(zH, zV, i, "Макс.заказ в ЗЕ указан некорректно")
                ElseIf Len(wsForm.Cells(i, cNewOptions + 9).Value2) <> Len(check_str(wsForm.Cells(i, cNewOptions + 9).Value2, "#")) Then
                    zH = errFunction(zH, zV, i, "Макс.заказ в ЗЕ должен иметь числовой формат")
                End If
            End If
            '-------------------------------------------------------------------изменение конечной даты / удаление/закрытие---------------------------------------------------------------------------
        ElseIf wsForm.Cells(i, cActionType).Value2 = "Изменение кон.даты" Or wsForm.Cells(i, cActionType).Value2 = "Закрытие/Удаление" Or _
                ((wsForm.Cells(i, cActionType).Value2 = "Смена внешнего поставщика" Or wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки") And SmflagOld) Then
            If wsForm.Cells(i, cOldOptions).Value2 = "" Then
                zH = errFunction(zH, zV, i, "Для типа действия " & wsForm.Cells(i, cActionType).Value2 & " следует заполнить кон.дату старой настройки")
            End If
            For j = 0 To 9
                If wsForm.Cells(i, cNewOptions + j).Value2 <> "" Then flagExtData = "<Новые настройки>"
            Next j
        Else
            zH = errFunction(zH, zV, i, "Тип действия не корректен")
        End If
        '-----------------------------------------------------старые & новые настройки (необходимость ТВЗ)-----------------------------------------------------
        If wsForm.Cells(i, cAddOptions).Value2 <> "" Then
            If wsForm.Cells(i, cActionType).Value2 <> "Изменение" And wsForm.Cells(i, cActionType).Value2 <> "Открытие" Then
                zH = errFunction(zH, zV, i, "При необходимости ТВЗ тип действия указан некорректно")
                flagTVZ = False
            Else
                If wsForm.Cells(i, cAddOptions).Value2 = "да" Then
                    If wsForm.Cells(i, cOldOptions).Value2 = "" And wsForm.Cells(i, cActionType).Value2 = "Изменение" Then
                        zH = errFunction(zH, zV, i, "При ТВЗ должна быть заполнена кон. дата старой настройки")
                        flagOD = False: flagTVZ = False
                    End If
                    If Left(TypePost, 1) <> "3" Then
                        zH = errFunction(zH, zV, i, "При необходимости ТВЗ должен быть указан внутренний поставщик")
                        flagTVZ = False
                    End If
                    If wsForm.Cells(i, cActionType).Value2 = "Открытие" And wsForm.Cells(i, cOldOptions + 1).Value2 = "" Then
                        zH = errFunction(zH, zV, i, "При необходимости ТВЗ должен быть заполнен ЛВ старой настройки")
                        flagTVZ = False
                    End If
                    If flagND Then
                        If wsForm.Cells(i, cActionType).Value2 = "Открытие" And CDate(wsForm.Cells(i, cNewOptions + 1).Value2) = Date And CDate(wsForm.Cells(i, cNewOptions + 1).Value2) = CDate(wsForm.Cells(i, cNewOptions).Value2) Then
                            zH = errFunction(zH, zV, i, "Период ТВЗ не должен быть текущими датами"): flagTVZ = False
                        End If
                    End If
                    If wsForm.Cells(i, cNewOptions).Value2 = "" Then
                        zH = errFunction(zH, zV, i, "При ТВЗ должна быть заполнена нач. дата новой настройки"): flagTVZ = False
                    End If
                    If wsForm.Cells(i, cOldOptions + 1).Value2 = "" And wsForm.Cells(i, cActionType).Value2 = "Изменение" Then
                        zH = errFunction(zH, zV, i, "При необходимости ТВЗ должен быть заполнен ЛВ старой настройки"): flagTVZ = False
                    End If
                    If wsForm.Cells(i, cNewOptions + 2).Value2 = "" And wsForm.Cells(i, cActionType).Value2 = "Изменение" Then
                        zH = errFunction(zH, zV, i, "При необходимости ТВЗ должен быть заполнен ЛВ новой настройки"): flagTVZ = False
                    End If
                    If wsForm.Cells(i, cOldOptions + 1).Value2 = wsForm.Cells(i, cNewOptions + 2).Value2 Then    'And wsForm.Cells(i, cActionType).Value2 = "Изменение" Then
                        zH = errFunction(zH, zV, i, "При ТВЗ должны отличаться ЛВ новой и старой настройки"): flagTVZ = False
                    End If
                    If Not IsDate(wsForm.Cells(i, cOldOptions).Value) And wsForm.Cells(i, cOldOptions).Value <> "" And wsForm.Cells(i, cActionType).Value2 <> "Открытие" And wsForm.Cells(i, cActionType).Value2 <> "Опт: смена поставщика для авторасценки" Then
                        zH = errFunction(zH, zV, i, "Кон.дата старой настройки указана в некорректном формате"): flagOD = False
                    End If
                    If flagND Then
                        If Not IsDate(wsForm.Cells(i, cNewOptions).Value) Then
                            zH = errFunction(zH, zV, i, "Нач.дата новой настройки указана в некорректном формате")
                        ElseIf flagOD And wsForm.Cells(i, cActionType).Value2 = "Изменение" Then
                            If CDate(wsForm.Cells(i, cNewOptions).Value2) >= CDate(wsForm.Cells(i, cOldOptions).Value2) Then
                                zH = errFunction(zH, zV, i, "При ТВЗ необходимо пересечение дат старой и новой настроек")
                                flagTVZ = False
                            End If
                        End If
                    End If
                    'проверка контрактов
                    If Len(wsForm.Cells(i, cNewOptions + 5).Value2) > 0 And Len(wsForm.Cells(i, cOldOptions + 4).Value2) > 0 Then
                        If wsForm.Cells(i, cNewOptions + 5).Value2 <> wsForm.Cells(i, cOldOptions + 4).Value2 Then
                            zH = errFunction(zH, zV, i, "Указаны разные контракты"): flagTVZ = False
                        End If
                    End If
                Else
                    zH = errFunction(zH, zV, i, "При необходимости ТВЗ значение должно быть ""да"", необходимо скорректировать")
                    flagTVZ = False
                End If
            End If
            If wsForm.Cells(i, cActionType).Value2 = "Открытие" Or wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Then
                If wsForm.Cells(i, cOldOptions).Value2 <> "" Then flagExtData = "<Старые настройки>"
                For j = 2 To 8
                    If (wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" And flagOsnPostMag And wsForm.Cells(i, cOldOptions + 4).Value2 <> "") Or _
                            (wsForm.Cells(i, cActionType).Value2 = "Открытие" And wsForm.Cells(i, cAddOptions).Value2 = "да" And (wsForm.Cells(i, cOldOptions + 1).Value2 <> "" Or flagTVZ = False)) Then
                    Else
                        flagExtData = "<Старые настройки>"
                    End If
                Next j
            End If
        ElseIf wsForm.Cells(i, cActionType).Value2 = "Открытие" Or wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" Then
            For j = 0 To 8
                If wsForm.Cells(i, cOldOptions + j).Value2 <> "" Then
                    If (wsForm.Cells(i, cActionType).Value2 = "Опт: смена поставщика для авторасценки" And flagOsnPostMag And wsForm.Cells(i, cOldOptions + 4).Value2 <> "") Or _
                            (wsForm.Cells(i, cActionType).Value2 = "Открытие" And wsForm.Cells(i, cAddOptions).Value2 = "да" And (wsForm.Cells(i, cOldOptions + 1).Value2 <> "" Or flagTVZ = False)) Then
                    Else
                        flagExtData = "<Старые настройки>"
                    End If
                End If
            Next j
        End If
        '-------------проверки по контрактам ПУШ/Постакция-------------------------
        Dim flagPostPush As Boolean: flagPostPush = False
        If flagKK And Left(TypePost, 1) = "3" Then
            If Len(wsForm.Cells(i, cNewOptions + 5).Value2) > 0 Or Len(wsForm.Cells(i, cOldOptions + 4).Value2) > 0 Then
                If (wsForm.Cells(i, cActionType).Value2 = "Открытие" Or (wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки" And SmflagNew)) And DictKK.exists(wsForm.Cells(i, cNewOptions + 5).Value2) Then
                    If DictKK(wsForm.Cells(i, cNewOptions + 5).Value2) = "0" And wsForm.Cells(i, cAddOptions).Value2 <> "да" Then zH = errFunction(zH, zV, i, "Указанный " & wsForm.Cells(i, cNewOptions + 5).Value2 & " обработан не будет, для корректировки кк_постакции необходимо обратиться в ОСЗМ")
                ElseIf (wsForm.Cells(i, cActionType).Value2 = "Изменение" And wsForm.Cells(i, cAddOptions).Value2 <> "да") Or wsForm.Cells(i, cActionType).Value2 = "Закрытие/Удаление" Or _
                        (wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки" And SmflagOld) Then
                    If DictKK.exists(wsForm.Cells(i, cNewOptions + 5).Value2) Or DictKK.exists(wsForm.Cells(i, cOldOptions + 4).Value2) Then
                        If DictKK.exists(wsForm.Cells(i, cNewOptions + 5).Value2) Then zH = errFunction(zH, zV, i, "Указанный " & wsForm.Cells(i, cNewOptions + 5).Value2 & " обработан не будет, для корректировки кк_пуш/кк_постакции необходимо обратиться в ОСЗМ")
                        If DictKK.exists(wsForm.Cells(i, cOldOptions + 4).Value2) Then zH = errFunction(zH, zV, i, "Указанный " & wsForm.Cells(i, cOldOptions + 4).Value2 & " обработан не будет, для корректировки кк_пуш/кк_постакции необходимо обратиться в ОСЗМ")
                    End If
                End If
            Else
                If cnORA.State = 1 Then
                    strQry = "select ARACEXR from cs.ARTUC inner join cs.FOUCCOM on ARACCIN = FCCCCIN and (FCCNATC='0' or (FCCNATC='8' and FCCLIB like '%ПУШ%')) " & _
                            "where ARADFIN>trunc(sysdate) and ARACEXR in ('" & wsForm.Cells(i, cProductCode).Value2 & "') and cs.PKFOUDGENE.GET_CNUF(1, aracfin) in ('" & wsForm.Cells(i, cSupplierCode).Value2 & "') and rownum=1"
                    'Debug.Print strQry
                    rsOra.Open strQry, cnORA
                    If Not rsOra.bof Then flagPostPush = True
                    rsOra.Close
                End If
                If flagPostPush Then
                    If (wsForm.Cells(i, cActionType).Value2 = "Закрытие/Удаление" Or wsForm.Cells(i, cActionType).Value2 = "Изменение" Or _
                            (wsForm.Cells(i, cActionType).Value2 = "Смена условий поставки" And SmflagOld)) And wsForm.Cells(i, cAddOptions).Value2 <> "да" Then
                        zH = errFunction(zH, zV, i, "Настройки в кк_пуш/кк_постакция обработаны не будут, необходимо обратиться в ОСЗМ")
                    End If
                End If
            End If
        End If
        If wsForm.Cells(i, cAddOptions + 1).Value2 = "да" And Left(TypePost, 1) <> "3" Then
            zH = errFunction(zH, zV, i, "Для настроек PBL поставщик указан некорректно")
        ElseIf wsForm.Cells(i, cAddOptions + 1).Value2 <> "да" And Len(wsForm.Cells(i, cAddOptions + 1).Value2) > 0 Then
            zH = errFunction(zH, zV, i, "Для настроек PBL значение должно быть ""да"", необходимо скорректировать")
        End If
        '----------------------------проверка по МЛТ------------------
        If flagKK And Len(wsForm.Cells(i, cNewOptions + 5).Value2) > 0 Then    ' если в блоке новых настроек указан контракт и он корректен
            If flagSup And flagProd And Left(TypePost, 1) = "1" Then    'если нет ошибок по товару и поставщику и поставщик внешний
                If DictMLT.exists(wsForm.Cells(i, cProductCode).Value2) Then    'если товар является карточкой мультипака
                    If DictMLTKK.exists(wsForm.Cells(i, cNewOptions + 5).Value2) = False Then
                        zH = errFunction(zH, zV, i, "Для товара МЛТ коммерческий контракт " & wsForm.Cells(i, cNewOptions + 5).Value2 & " указан неверно")
                    End If
                End If
            End If
        End If
        If flagExtData <> "" Then zH = errFunction(zH, zV, i, "Для типа действия " & wsForm.Cells(i, cActionType).Value2 & ", указаны лишние данные. Необходимо затереть все значения в блоке " & flagExtData)
        '-----------------------проверка ЛЕ в системе-------------------------------------------
        Dim flagLETrue As Boolean
        If wsForm.Cells(i, cActionType).Value2 = "Изменение" And flagProd And flagSup And flagItem And flagNewLE And _
                Len(wsForm.Cells(i, cNewOptions + 3).Value2) > 0 And Len(wsForm.Cells(i, cOldOptions + 2).Value2) = 0 Then
            flagLETrue = True
            For k = LBound(MassItem) To UBound(MassItem)
                strQry = "Select distinct ARUTYPUL FROM CS.ARTUC INNER JOIN CS.ARTUL ON CS.ARTUL.ARUCINL = CS.ARTUC.ARACINL INNER JOIN CS.FOUDGENE ON CS.ARTUC.ARACFIN = CS.FOUDGENE.FOUCFIN INNER JOIN CS.FOUCCOM ON CS.FOUCCOM.FCCCCIN = CS.ARTUC.ARACCIN WHERE CS.ARTUC.ARACEXR = '" & wsForm.Cells(i, cProductCode).Value & "' and CS.FOUDGENE.FOUCNUF = '" & wsForm.Cells(i, cSupplierCode).Value & "' and ARADFIN>trunc(sysdate) "
                If wsForm.Cells(i, cItemCode).Value2 <> "" Then strQry = strQry + " and CS.ARTUC.ARASITE in (select relid from cs.resrel where reldfin >= trunc(sysdate) start with relpere  = '" & MassItem(k) & "' connect by prior relid = relpere Union all select TO_NUMBER('" & MassItem(k) & "') from dual) "
                If wsForm.Cells(i, cOldOptions).Value2 <> "" Then strQry = strQry + " and CS.ARTUC.ARADFIN >= to_date('" & wsForm.Cells(i, cOldOptions).Value & "','DD.MM.YYYY')"
                If wsForm.Cells(i, cOldOptions + 1).Value2 <> "" Then strQry = strQry + " and ARACEXVL = '" & wsForm.Cells(i, cOldOptions + 1).Value2 & "'"
                If wsForm.Cells(i, cOldOptions + 3).Value2 <> "" Then strQry = strQry + " and ARANFILF = '" & wsForm.Cells(i, cOldOptions + 3).Value2 & "'"
                If wsForm.Cells(i, cOldOptions + 4).Value2 <> "" Then strQry = strQry + " and FCCNUM = '" & wsForm.Cells(i, cOldOptions + 4).Value2 & "'"
                If wsForm.Cells(i, cOldOptions + 5).Value2 <> "" Then strQry = strQry + " and ARACEAN = '" & wsForm.Cells(i, cOldOptions + 5).Value2 & "'"
                If wsForm.Cells(i, cOldOptions + 6).Value2 <> "" Then strQry = strQry + " and ARAMUA = '" & wsForm.Cells(i, cOldOptions + 6).Value2 & "'"
                If wsForm.Cells(i, cOldOptions + 7).Value2 <> "" Then strQry = strQry + " and ARAMINCDE = '" & wsForm.Cells(i, cOldOptions + 7).Value2 & "'"
                If wsForm.Cells(i, cOldOptions + 8).Value2 <> "" Then strQry = strQry + " and ARAMAXCDE = '" & wsForm.Cells(i, cOldOptions + 8).Value2 & "'"
                'Debug.Print strQry
                rsOra.Open strQry, cnORA
                If Not rsOra.bof Then
                    Do While Not rsOra.EOF
                        If rsOra!ARUTYPUL <> CStr(wsForm.Cells(i, cNewOptions + 3).Value2) Then flagLETrue = False
                        rsOra.movenext
                    Loop
                End If
                rsOra.Close
            Next
            If flagLETrue = False Then zH = errFunction(zH, zV, i, "Обратите внимание, ЛЕ в блоке новые настройки отличается от ЛЕ в системе")
        End If
        '------------------------------формируем данные на листе ТВЗ только для просмотра-------------------------------------------
        Dim LEOld: LEOld = ""
        If wsForm.Cells(i, cAddOptions).Value2 = "да" And flagTVZ Then
            If Worksheets("ТВЗ").Visible = False Then
                ActiveWorkbook.Unprotect (1347)
                Worksheets("ТВЗ").Visible = True
                ActiveWorkbook.Protect (1347)
            End If
            Worksheets("ТВЗ").Protect Password:="1347", UserInterfaceOnly:=True
            With Worksheets("ТВЗ")
                .Cells(z, 1).Value = wsForm.Cells(i, cSupplierCode).Value
                .Cells(z, 2).Value = wsForm.Cells(i, cProductCode).Value
                .Cells(z, 3).Value = wsForm.Cells(i, cProductName).Value
                .Cells(z, 6).Value = wsForm.Cells(i, cProductCode).Value
                .Cells(z, 7).Value = wsForm.Cells(i, cProductName).Value
                .Cells(z, 10).Value = 1
                .Cells(z, 11).Value = 1
                .Cells(z, 12).Value = "2 - Замена при дефиците"
                If Len(wsForm.Cells(i, cOldOptions + 2).Value) = 0 And Len(wsForm.Cells(i, cOldOptions + 1).Value) > 0 And cnORA.State = 1 Then
                    strQry = "SELECT distinct ARUTYPUL ZaLe FROM cs.ARTUC inner join cs.foudgene on aracfin=FOUCFIN INNER JOIN CS.ARTUL ON ARUCINL = ARACINL INNER JOIN CS.FOUCCOM ON CS.FOUCCOM.FCCCCIN = CS.ARTUC.ARACCIN where ARACEXR = '" & CStr(wsForm.Cells(i, cProductCode).Value2) & "' and foucnuf = '" & CStr(wsForm.Cells(i, cSupplierCode).Value2) & "'  "
                    If wsForm.Cells(i, cOldOptions).Value <> "" Then
                        strQry = strQry + " and aradfin >= to_date('" & wsForm.Cells(i, cOldOptions).Value & "','DD.MM.YYYY')"
                    Else
                        strQry = strQry + " and aradfin >= trunc(sysdate) "
                    End If
                    If wsForm.Cells(i, cOldOptions + 1).Value <> "" Then strQry = strQry + " and CS.ARTUC.ARACEXVL = '" & wsForm.Cells(i, cOldOptions + 1).Value & "' "
                    If wsForm.Cells(i, cOldOptions + 2).Value <> "" Then strQry = strQry + " AND CS.ARTUL.ARUTYPUL = '" & wsForm.Cells(i, cOldOptions + 2).Value & "' "
                    If wsForm.Cells(i, cOldOptions + 3).Value <> "" Then strQry = strQry + " and CS.ARTUC.ARANFILF = '" & wsForm.Cells(i, cOldOptions + 3).Value & "' "
                    If wsForm.Cells(i, cOldOptions + 4).Value <> "" Then strQry = strQry + " and CS.FOUCCOM.FCCNUM = '" & wsForm.Cells(i, cOldOptions + 4).Value & "' "
                    If wsForm.Cells(i, cOldOptions + 5).Value <> "" Then strQry = strQry + " and CS.ARTUC.ARACEAN = '" & wsForm.Cells(i, cOldOptions + 5).Value & "' "
                    If wsForm.Cells(i, cOldOptions + 6).Value <> "" Then strQry = strQry + " AND CS.ARTUC.ARAMUA = '" & wsForm.Cells(i, cOldOptions + 6).Value & "' "
                    If wsForm.Cells(i, cOldOptions + 7).Value <> "" Then strQry = strQry + " AND CS.ARTUC.ARAMINCDE = '" & wsForm.Cells(i, cOldOptions + 7).Value & "' "
                    If wsForm.Cells(i, cOldOptions + 8).Value <> "" Then strQry = strQry + " AND CS.ARTUC.ARAMAXCDE = '" & wsForm.Cells(i, cOldOptions + 8).Value & "' "
                    strQry = strQry + "and rownum=1"
                    'Debug.Print strQry
                    rsOra.Open strQry, cnORA
                    If Not rsOra.bof Then LEOld = rsOra!ZaLe
                    rsOra.Close
                End If
                If wsForm.Cells(i, cNewOptions).Value <> "" Then .Cells(z, 13).Value = CDate(wsForm.Cells(i, cNewOptions).Value)
                If wsForm.Cells(i, cActionType).Value = "Открытие" Then
                    .Cells(z, 14).Value = CDate(wsForm.Cells(i, cNewOptions + 1).Value)
                    .Cells(z, 4).Value = wsForm.Cells(i, cNewOptions + 2).Value
                    If Len(wsForm.Cells(i, cNewOptions + 3).Value) > 0 Then
                        .Cells(z, 5).Value = wsForm.Cells(i, cNewOptions + 3).Value
                    ElseIf Len(wsForm.Cells(i, cOldOptions + 2).Value) > 0 Then
                        .Cells(z, 5).Value = wsForm.Cells(i, cOldOptions + 2).Value
                    Else
                        .Cells(z, 5).Value = LEOld
                    End If
                    .Cells(z, 8).Value = wsForm.Cells(i, cOldOptions + 1).Value
                    If Len(wsForm.Cells(i, cOldOptions + 2).Value) > 0 Then
                        .Cells(z, 9).Value = wsForm.Cells(i, cOldOptions + 2).Value
                    Else
                        .Cells(z, 9).Value = LEOld
                    End If
                ElseIf wsForm.Cells(i, cActionType).Value = "Изменение" And wsForm.Cells(i, cOldOptions).Value <> "" Then
                    .Cells(z, 14).Value = CDate(wsForm.Cells(i, cOldOptions).Value)
                    .Cells(z, 4).Value = wsForm.Cells(i, cOldOptions + 1).Value
                    If Len(wsForm.Cells(i, cOldOptions + 2).Value) > 0 Then
                        .Cells(z, 5).Value = wsForm.Cells(i, cOldOptions + 2).Value
                    Else
                        .Cells(z, 5).Value = LEOld
                    End If
                    .Cells(z, 8).Value = wsForm.Cells(i, cNewOptions + 2).Value
                    If Len(wsForm.Cells(i, cNewOptions + 3).Value) > 0 Then
                        .Cells(z, 9).Value = wsForm.Cells(i, cNewOptions + 3).Value
                    ElseIf Len(wsForm.Cells(i, cOldOptions + 2).Value) > 0 Then
                        .Cells(z, 9).Value = wsForm.Cells(i, cOldOptions + 2).Value
                    Else
                        .Cells(z, 9).Value = LEOld
                    End If
                End If
                z = z + 1
            End With
            zH = errFunction(zH, zV, i, "Определен заменитель и заменяемый товар, информация на листе ""ТВЗ""")
        End If
        '----------------------------------проверка по постакции для типа действия Открытие-----------------------------------------
        If wsForm.Cells(i, cAddOptions).Value2 = "да" And flagTVZ And wsForm.Cells(i, cActionType).Value = "Открытие" And _
                Len(wsForm.Cells(i, cOldOptions + 1)) > 0 And DictKK.exists(wsForm.Cells(i, cNewOptions + 5).Value2) And flagND Then
            If DictKK(wsForm.Cells(i, cNewOptions + 5).Value2) = "0" Then
                If cnORA.State = 1 Then
                    strQry = "select distinct ARADDEB, ARADFIN from cs.ARTUC inner join cs.FOUCCOM on ARACCIN = FCCCCIN " & _
                            "inner join cs.artvl on ARASEQVL = ARLSEQVL where ARADFIN>trunc(sysdate) and ARACEXR in ('" & wsForm.Cells(i, cProductCode).Value2 & "') and " & _
                            "arlcexvl = '" & CStr(wsForm.Cells(i, cOldOptions + 1).Value2) & "' and FCCNUM='" & wsForm.Cells(i, cNewOptions + 5).Value2 & "'"
                    'Debug.Print strQry
                    rsOra.Open strQry, cnORA
                    If rsOra.bof Then
                        zH = errFunction(zH, zV, i, "В системе нет настроек в указанном КК_постакции на ЛВ из блока старые настройки")
                    Else
                        FlagPostaDates = False
                        Do While Not rsOra.EOF
                            If CDate(wsForm.Cells(i, cNewOptions).Value) > Date Then
                                If rsOra!ARADDEB > Date Then
                                    If CDate(wsForm.Cells(i, cNewOptions).Value) = rsOra!ARADDEB And _
                                            CDate(wsForm.Cells(i, cNewOptions + 1).Value) = rsOra!ARADFIN Then
                                        FlagPostaDates = True: Exit Do
                                    End If
                                Else
                                    If CDate(wsForm.Cells(i, cNewOptions + 1).Value) = rsOra!ARADFIN Then
                                        FlagPostaDates = True: Exit Do
                                    End If
                                End If
                            ElseIf CDate(wsForm.Cells(i, cNewOptions).Value) <= Date Then
                                If CDate(wsForm.Cells(i, cNewOptions + 1).Value) = rsOra!ARADFIN Then
                                    FlagPostaDates = True: Exit Do
                                End If
                            End If
                            rsOra.movenext
                        Loop
                        If FlagPostaDates = False Then
                            zH = errFunction(zH, zV, i, "Даты в блоке Новые настройки по КК_постакция не совпадают с системой, необходимо скорректировать")
                        End If
                    End If
                    rsOra.Close
                End If
            End If
        End If
        Call LVSTOK
        ''''-------------------------------------- ЛВ в сток

        '---------------------------------------------------------------------------------------------------------------------------
        If zH > 4 Then zV = zV + 1
        If zH > MaxZH Then MaxZH = zH
    Next i
    '-----------------------------------------------------дубликаты--------------------------------------------------------------
    'wsForm.Activate
    'wsForm.Range(wsForm.Cells(lastRow + 1, 2), wsForm.Cells(Rows.Count, 30)).Interior.Color = xlNone
    Dim arrStr()    As String
    If lastRow >= 10 Then
        ReDim Preserve arrStr(lastRow - 10)
        For i = firstRow To lastRow
            For j = 2 To 30
                If wsForm.Cells(i, j).Interior.Color = RGB(218, 238, 243) Then wsForm.Cells(i, j).Interior.Color = xlNone
                arrStr(i - 10) = arrStr(i - 10) & wsForm.Cells(i, j).Value
            Next j
            If wsForm.Cells(i, 2).Interior.Color <> RGB(253, 233, 217) Then wsForm.Range(wsForm.Cells(i, 1), wsForm.Cells(i, 30)).Locked = False
        Next i
        For i = 0 To UBound(arrStr)
            For j = i + 1 To UBound(arrStr)
                If arrStr(i) = arrStr(j) Then
                    wsForm.Range(wsForm.Cells(i + 10, 2), wsForm.Cells(i + 10, 29)).Interior.Color = RGB(218, 238, 243)
                    wsForm.Range(wsForm.Cells(j + 10, 2), wsForm.Cells(j + 10, 29)).Interior.Color = RGB(218, 238, 243)
                End If
            Next j
        Next i
    End If
    '----------------------------------------------------------------------------------------------------------------------------
    Lc = wsResult.Rows(1).CurrentRegion.Columns.Count
    wsResult.Range(wsResult.Columns(1), wsResult.Columns(Lc)).AutoFilter
    If MaxZH > 4 Then wsResult.Columns("A:" & Col_Letter(MaxZH - 1)).EntireColumn.AutoFit
    wsResult.Activate
    Application.ScreenUpdating = True
    wsForm.Cells(5, 4) = "Время работы = " & Round(Timer - BenchMark, 2)    'Время работы

    Worksheets("Форма_A02ZA").Protect Password:="1347", UserInterfaceOnly:=True

End Sub

Private Function errFunction(zH As Integer, zV As Integer, i As Long, errText As String) As Integer
    If errcount = 1 Then errFunction = zH: GoTo lol
    Worksheets("Отчет").Cells(zV, zH) = errText
    If zH = 4 Then
        Worksheets("Отчет").Cells(zV, 1).Value2 = i
        Worksheets("Отчет").Cells(zV, 2).Value2 = Worksheets("Форма_A02ZA").Cells(i, 3).Value2
        Worksheets("Отчет").Cells(zV, 3).Value2 = Worksheets("Форма_A02ZA").Cells(i, 4).Value2
    End If
    errFunction = zH + 1
lol:

End Function
Private Sub DesignResult(wsResult As Worksheet)
    'With wsResult
    'If .Cells(.Rows.Count, 1).End(xlUp).Row > 1 Then
    '.Rows("2:" & .Cells(.Rows.Count, 1).End(xlUp).Row).ClearContents
    'End If
    'With wsResult
    With wsResult.Columns
        .Clear
        .Borders.Weight = xlThin
        .ColumnWidth = 13.57
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 10
    End With
    wsResult.Columns("A:C").HorizontalAlignment = xlCenter
    wsResult.Columns("A:C").VerticalAlignment = xlCenter
    With wsResult
        .Cells(1, 1) = "№ строки"
        .Cells(1, 2) = "Код товара GOLD"
        .Cells(1, 3) = "Наименование товарной позиции"
        .Cells(1, 4) = "Ошибки"
        .Columns(1).ColumnWidth = 6.43
        .Columns(2).ColumnWidth = 16
        .Columns(3).ColumnWidth = 32
        .Rows.RowHeight = 45
        .Rows(1).RowHeight = 35
        .Rows(1).Interior.Color = RGB(141, 180, 226)
    End With
    Application.GoTo wsResult.Range("A1"), True
    ActiveWindow.Zoom = 80
    With wsResult.Range(Cells(1, 1), Cells(1, 4))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    With wsResult.Range(Cells(1, 4), Cells(1, wsResult.Columns.Count))
        .Borders.LineStyle = xlNone
    End With
    With wsResult.Rows(1)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    wsResult.Cells.FormatConditions.Delete
    With wsResult.Rows("2:" & wsResult.Rows.Count)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ОСТАТ(СТРОКА();2)"
        .FormatConditions(1).Interior.Color = RGB(197, 217, 241)
        .FormatConditions(1).StopIfTrue = False
    End With
    wsResult.Rows(1).RowHeight = 65
    wsResult.Range("G1").Merge
    With wsResult.Range("G1")
        .Font.Size = 14
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Borders.Weight = xlThin
        .ColumnWidth = 20.86
    End With

End Sub
Function Col_Letter(lngCol As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
Sub ReturnToMainSheet()
    Worksheets("Форма_A02ZA").Activate
End Sub
Function check_str(str As String, mask As String) As String
    Dim i As Integer, total As String
    total = ""
    For i = 1 To Len(str)
        If Mid(str, i, 1) Like mask Then
            total = total + Mid(str, i, 1)
        End If
    Next i
    check_str = total
End Function
Function check_natur(str As String) As Boolean
    Dim i           As Integer
    For i = 1 To Len(str)
        If Mid(str, i, 1) = "-" Or Mid(str, i, 1) = "," Then
            check_natur = False
            Exit Function
        End If
    Next i
    check_natur = True
End Function
Function MatchDuplicatesArray(strArray() As String, k As Integer) As Boolean
    Dim l           As Integer
    For l = k To UBound(strArray)
        If l <> k Then
            If strArray(k) = strArray(l) Then MatchDuplicatesArray = True
        End If
    Next
End Function

Function CreateCollection() As Collection
    Dim coll        As New Collection

    Set CreateCollection = coll
End Function

Function CreateCollectionRC() As Collection
    Dim collRC      As New Collection

    Set CreateCollectionRC = collRC
End Function

Sub Dict_Proc()
    Dim lastTer     As Long

    Set wsTerra = Worksheets("Территории")
    If cnORA.State = 1 Then
        'выгрузка магазинов
        strQry = "select distinct d1.relid Sait,(select trobdesc from cs.tra_resobj where trobid = d1.relid and langue = 'RU') Opis,d4.relpere as FO " & _
                "from cs.resrel d1 left join cs.resrel d2 on (d2.relid=d1.relpere and trunc(sysdate) between d2.relddeb and d2.reldfin) " & _
                "left join cs.resrel d3 on (d3.relid=d2.relpere and trunc(sysdate) between d3.relddeb and d3.reldfin) " & _
                "left join cs.resrel d4 on (d4.relid=d3.relpere and trunc(sysdate) between d4.relddeb and d4.reldfin) " & _
                "left join cs.resrel d5 on (d5.relid=d4.relpere and trunc(sysdate) between d5.relddeb and d5.reldfin) " & _
                "left join cs.resobj on d1.relpere = robid " & _
                "inner join cs.sitattri on satsite = d1.relid and satcla='WH' and  satdfin>=trunc(sysdate) order by 1"
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            lastTer = wsTerra.Cells(wsTerra.Rows.Count, 8).End(xlUp).Row
            wsTerra.Range(wsTerra.Cells(2, 8), wsTerra.Cells(lastTer, 10)).ClearContents
            wsTerra.Cells(2, 8).CopyFromRecordset rsOra
        End If
        rsOra.Close
        'выгрузка узлов
        strQry = "select YZEL1,OPIS1,YZEL2,OPIS2,YZEL3,OPIS3 from (select distinct to_char(d2.relpere) YZEL1,(select trobdesc from cs.tra_resobj where trobid = d2.relpere and langue = 'RU')OPIS1, " & _
                " to_char(d3.relpere) YZEL2,(select trobdesc from cs.tra_resobj where trobid = d3.relpere and langue = 'RU')OPIS2," & _
                " to_char(d4.relpere) YZEL3,case when (select trobdesc from cs.tra_resobj where trobid = d4.relpere and langue = 'RU')='ЦЕНТРАЛЬНЫЙ' then 'ЦФО МАГАЗИНЫ'" & _
                " when (select trobdesc from cs.tra_resobj where trobid = d4.relpere and langue = 'RU')='СЕВЕРО-ЗАПАДНЫЙ' then 'СЗФО МАГАЗИНЫ'" & _
                " when (select trobdesc from cs.tra_resobj where trobid = d4.relpere and langue = 'RU')='УРАЛЬСКИЙ' then 'УФО МАГАЗИНЫ' end OPIS3" & _
                " from cs.resrel d1" & _
                " left join cs.resrel d2 on (d2.relid=d1.relpere and trunc(sysdate) between d2.relddeb and d2.reldfin)" & _
                " left join cs.resrel d3 on (d3.relid=d2.relpere and trunc(sysdate) between d3.relddeb and d3.reldfin)" & _
                " left join cs.resrel d4 on (d4.relid=d3.relpere and trunc(sysdate) between d4.relddeb and d4.reldfin)" & _
                " left join cs.resrel d5 on (d5.relid=d4.relpere and trunc(sysdate) between d5.relddeb and d5.reldfin)" & _
                " left join cs.resobj on d1.relpere = robid" & _
                " where trunc(sysdate) between d1.relddeb and d1.reldfin" & _
                " and robresid != 4" & _
                " and d1.relid in (select relid from cs.resrel d1, cs.resobj where relid = robid and robprof = -1)" & _
                " and d4.relpere in ('94005','94006','94007') order by 5,3,1) Union all" & _
                " select YZEL1,OPIS1,YZEL2,OPIS2,YZEL3,OPIS3 from OK__NODES order by 5,3,1"
        ' Debug.Print strQry
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            lastTer = wsTerra.Cells(wsTerra.Rows.Count, 7).End(xlUp).Row
            wsTerra.Range(wsTerra.Cells(2, 2), wsTerra.Cells(lastTer, 7)).ClearContents
            wsTerra.Cells(2, 2).CopyFromRecordset rsOra
        End If
        rsOra.Close
    End If


    Set DictTer = CreateObject("Scripting.Dictionary")
    For i = 2 To wsTerra.Cells(wsTerra.Rows.Count, 7).End(xlUp).Row
        If CStr(wsTerra.Cells(i, 6)) = "94003" Then
            If DictTer.exists(CStr(wsTerra.Cells(i, 2).Value2)) = False Then DictTer.Add CStr(wsTerra.Cells(i, 2).Value2), CStr(wsTerra.Cells(i, 6).Value2)
        End If
    Next


    Set DictMatrRc = CreateObject("Scripting.Dictionary")
    For i = 2 To wsTerra.Cells(wsTerra.Rows.Count, 7).End(xlUp).Row
        If CStr(wsTerra.Cells(i, 6)) = "94003" And Len(wsTerra.Cells(i, 4).Value2) > 0 Then
            If DictMatrRc.exists(wsTerra.Cells(i, 6).Value2) = False Then
                DictMatrRc.Add wsTerra.Cells(i, 6).Value2, CreateCollectionRC
                DictMatrRc(wsTerra.Cells(i, 6).Value2).Add wsTerra.Cells(i, 2).Value2
            Else
                DictMatrRc(wsTerra.Cells(i, 6).Value2).Add wsTerra.Cells(i, 2).Value2
            End If

            If DictMatrRc.exists(wsTerra.Cells(i, 4).Value2) = False Then
                DictMatrRc.Add wsTerra.Cells(i, 4).Value2, CreateCollectionRC
                DictMatrRc(wsTerra.Cells(i, 4).Value2).Add wsTerra.Cells(i, 2).Value2
            Else
                DictMatrRc(wsTerra.Cells(i, 4).Value2).Add wsTerra.Cells(i, 2).Value2
            End If
        End If
    Next



    Set DictFO = CreateObject("Scripting.Dictionary")
    For i = 2 To wsTerra.Cells(wsTerra.Rows.Count, 7).End(xlUp).Row
        If CStr(wsTerra.Cells(i, 6)) = "94003" Then
            If CStr(wsTerra.Cells(i, 2).Value2) = "94005" Or CStr(wsTerra.Cells(i, 2).Value2) = "94006" Or CStr(wsTerra.Cells(i, 2).Value2) = "94007" Then
                If DictFO.exists(CStr(wsTerra.Cells(i, 2).Value2)) = False Then DictFO.Add CStr(wsTerra.Cells(i, 2).Value2), CStr(wsTerra.Cells(i, 2).Value2)
            Else
                If DictFO.exists(CStr(wsTerra.Cells(i, 2).Value2)) = False Then
                    If CStr(wsTerra.Cells(i, 4).Value2) = "94009" Then
                        DictFO.Add CStr(wsTerra.Cells(i, 2).Value2), "94005"
                        If DictFO.exists(CStr(wsTerra.Cells(i, 4).Value2)) = False Then DictFO.Add CStr(wsTerra.Cells(i, 4).Value2), "94005"
                    ElseIf CStr(wsTerra.Cells(i, 4).Value2) = "94010" Then
                        DictFO.Add CStr(wsTerra.Cells(i, 2).Value2), "94006"
                        If DictFO.exists(CStr(wsTerra.Cells(i, 4).Value2)) = False Then DictFO.Add CStr(wsTerra.Cells(i, 4).Value2), "94006"
                    ElseIf CStr(wsTerra.Cells(i, 4).Value2) = "94011" Then
                        DictFO.Add CStr(wsTerra.Cells(i, 2).Value2), "94007"
                        If DictFO.exists(CStr(wsTerra.Cells(i, 4).Value2)) = False Then DictFO.Add CStr(wsTerra.Cells(i, 4).Value2), "94007"
                    End If
                End If
            End If
        ElseIf CStr(wsTerra.Cells(i, 6).Value2) = "94005" Or CStr(wsTerra.Cells(i, 6).Value2) = "94006" Or CStr(wsTerra.Cells(i, 6).Value2) = "94007" Then
            If DictFO.exists(CStr(wsTerra.Cells(i, 2).Value2)) = False Then DictFO.Add CStr(wsTerra.Cells(i, 2).Value2), CStr(wsTerra.Cells(i, 6).Value2)
            If DictFO.exists(CStr(wsTerra.Cells(i, 4).Value2)) = False Then DictFO.Add CStr(wsTerra.Cells(i, 4).Value2), CStr(wsTerra.Cells(i, 6).Value2)
        End If
    Next

    For i = 2 To wsTerra.Cells(wsTerra.Rows.Count, 8).End(xlUp).Row
        If DictFO.exists(CStr(wsTerra.Cells(i, 8).Value2)) = False Then DictFO.Add CStr(wsTerra.Cells(i, 8).Value2), CStr(wsTerra.Cells(i, 10).Value2)
    Next

    '---------словарь мультипаков-------------
    'товары

    Set DictMLT = CreateObject("Scripting.Dictionary")
    'контракты

    Set DictMLTKK = CreateObject("Scripting.Dictionary")
    If cnORA.State = 1 Then
        strQry = "select artcexr tovar from cs.ARTATTRI, cs.artrac where artcinr = AATCINR and AATCCLA = 'MLT' and AATDFIN > trunc(sysdate)"
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            Do While Not rsOra.EOF
                If DictMLT.exists(rsOra.Fields("tovar").Value) = False Then DictMLT.Add rsOra.Fields("tovar").Value, 1
                rsOra.movenext
            Loop
        End If
        rsOra.Close

        strQry = "select FCCNUM KK from CS.FOUCCOM where FCCLIB LIKE '%МЛТ%' and FCCDFIN>trunc(sysdate)"
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            Do While Not rsOra.EOF
                If DictMLTKK.exists(rsOra.Fields("KK").Value) = False Then DictMLTKK.Add rsOra.Fields("KK").Value, 1
                rsOra.movenext
            Loop
        End If
        rsOra.Close
    End If


    Set DictKK = CreateObject("Scripting.Dictionary")

    If cnORA.State = 1 Then
        'словарь КК ПУШ и Постакции
        strQry = "select fccnum KK,FCCNATC TypeKK from cs.fouccom where (FCCNATC='0' or (FCCNATC='8' and FCCLIB like '%ПУШ%')) and FCCDFIN>trunc(sysdate)"
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            Do While Not rsOra.EOF
                If DictKK.exists(rsOra.Fields("KK").Value) = False Then DictKK.Add rsOra.Fields("KK").Value, rsOra.Fields("TypeKK").Value
                rsOra.movenext
            Loop
        End If
        rsOra.Close
    End If


    Set DictRuchKK = CreateObject("Scripting.Dictionary")
    If cnORA.State = 1 Then
        'словарь КК Ручные
        strQry = "select fccnum KK from cs.fouccom where FCCNATC='8' and FCCDFIN>trunc(sysdate)"
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            Do While Not rsOra.EOF
                If DictRuchKK.exists(rsOra.Fields("KK").Value) = False Then DictRuchKK.Add rsOra.Fields("KK").Value, 1
                rsOra.movenext
            Loop
        End If
        rsOra.Close
    End If

End Sub

Public Sub LVSTOK()


    Dim f
    Dim OraUzel
    '-------------------------------------- ЛВ в сток
    'собираем строки на открытие настроек на узел 70995 по товару+Поставщику(РЦ)+ЛВ
    Set dictFUZ = CreateObject("Scripting.Dictionary")
    Set DictVirtUzel = CreateObject("Scripting.Dictionary")
    For irow = firstRow To lastRow
        'If IsError(Application.Match(CLng(70995), CLng(wsForm.Cells(irow, cItemCode)), 0)) = False Then
        If InStr(1, "70995", CStr(wsForm.Cells(irow, cItemCode))) <> 0 Then
            If DictVirtUzel.exists(CStr(wsForm.Cells(irow, cProductCode) & wsForm.Cells(irow, cSupplierCode) & wsForm.Cells(irow, cNewOptions + 2))) = False Then
                DictVirtUzel.Add CStr(wsForm.Cells(irow, cProductCode) & wsForm.Cells(irow, cSupplierCode) & wsForm.Cells(irow, cNewOptions + 2)), ""
            End If
        End If
    Next irow
    If cnORA.State = 1 Then
        'strQry = "select distinct  yzel1 from ok__nodes where yzel3 = 94003 and yzel1 is not null Union all select distinct yzel2 from ok__nodes where yzel3 = 94003 and yzel2 is not null"

        strQry = "select distinct  yzel1,nvl(yzel2,1) yzel2 from ok__nodes where yzel3 = 94003 and yzel1 is not null Union " & _
                " select '94009','1' from dual union select '94010','1' from dual union select '94011','1' from dual union" & _
                " select distinct yzel1,yzel2 from ok__nodes where yzel3 = 94003 and yzel2 is not null"
        'Debug.Print strQry
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            Do While Not rsOra.EOF
                strUzelSH = strUzelSH & rsOra!yzel1 & ","
                If dictFUZ.exists(CStr(rsOra!yzel2)) = False Then
                    uz_item = ""
                    For f = 2 To wsTerra.Cells(wsTerra.Rows.Count, 7).End(xlUp).Row

                        If CStr(wsTerra.Cells(f, 4).Value2) = CStr(rsOra!yzel2) Then
                            uz_item = uz_item & wsTerra.Cells(f, 2).Value2 & ","
                        End If
                    Next f
                    If Right(uz_item, 1) = "," Then uz_item = Left(uz_item, Len(uz_item) - 1)
                    dictFUZ.Add (CStr(rsOra!yzel2)), CStr(uz_item)
                End If
                rsOra.movenext
            Loop
        End If
        rsOra.Close
    End If

    strItemCode = wsForm.Cells(i, cItemCode).Value

    For Each str In Split(Application.Trim(Replace(Replace(Replace(strItemCode, ";", ","), ",", ", "), ",", " ,")), ",")    'узел
        If Left(TypePost, 1) = "1" And Len(wsForm.Cells(i, cProductCode).Value) > 0 And InStr(1, strUzelSH, Trim(str)) <> 0 And Len(CStr(wsForm.Cells(i, cNewOptions + 2).Value2)) > 0 And wsForm.Cells(i, cItemName).Value2 Like "*РЦ*" Then
            If cnORA.State = 1 Then
                strQry = "select ar_cproin, RTRIM(xmlagg(xmlelement(E,ar_donord,',')).extract('//text()'),',') uzel from (select ar_cproin, ar_donord, ar_ilogis, ar_libpro, SLG.PI_LI2POS" & _
                        " from tb_art@stock7 , tb_parl@stock7 stt, tb_parl@stock7 slg where ar_donord = 70007" & _
                        " and (stt.pi_tablex = 'STT' and stt.pi_postex = ar_stprod and stt.pi_lang = 'RU')" & _
                        " and (slg.pi_tablex = 'SLG' and slg.pi_postex = ar_donlog and slg.pi_lang = 'RU')" & _
                        " Union all select ar_cproin, ar_donord, ar_ilogis, ar_libpro, SLG.PI_LI2POS" & _
                        " from tb_art@stock11 , tb_parl@stock11 stt, tb_parl@stock11 slg where ar_donord = 70011" & _
                        " and (stt.pi_tablex = 'STT' and stt.pi_postex = ar_stprod and stt.pi_lang = 'RU')" & _
                        " and (slg.pi_tablex = 'SLG' and slg.pi_postex = ar_donlog and slg.pi_lang = 'RU')" & _
                        " Union all select ar_cproin, ar_donord, ar_ilogis, ar_libpro, SLG.PI_LI2POS" & _
                        " from tb_art@stock35 , tb_parl@stock35 stt, tb_parl@stock35 slg where ar_donord = 70004" & _
                        " and (stt.pi_tablex = 'STT' and stt.pi_postex = ar_stprod and stt.pi_lang = 'RU')" & _
                        " and (slg.pi_tablex = 'SLG' and slg.pi_postex = ar_donlog and slg.pi_lang = 'RU')" & _
                        " Union all select ar_cproin, ar_donord, ar_ilogis, ar_libpro, SLG.PI_LI2POS" & _
                        " from tb_art@stock  , tb_parl@stock stt, tb_parl@stock slg where ar_donord = 70003" & _
                        " and (stt.pi_tablex = 'STT' and stt.pi_postex = ar_stprod and stt.pi_lang = 'RU')" & _
                        " and (slg.pi_tablex = 'SLG' and slg.pi_postex = ar_donlog and slg.pi_lang = 'RU')" & _
                        " Union all select ar_cproin, ar_donord, ar_ilogis, ar_libpro, SLG.PI_LI2POS" & _
                        " from tb_art@stock34 , tb_parl@stock34 stt, tb_parl@stock34 slg where ar_donord = 70034" & _
                        " and (stt.pi_tablex = 'STT' and stt.pi_postex = ar_stprod and stt.pi_lang = 'RU')" & _
                        " and (slg.pi_tablex = 'SLG' and slg.pi_postex = ar_donlog and slg.pi_lang = 'RU')" & _
                        " Union all select ar_cproin, ar_donord, ar_ilogis, ar_libpro, SLG.PI_LI2POS" & _
                        " from tb_art@stock35 , tb_parl@stock35 stt, tb_parl@stock35 slg where ar_donord = 70035" & _
                        " and (stt.pi_tablex = 'STT' and stt.pi_postex = ar_stprod and stt.pi_lang = 'RU')" & _
                        " and (slg.pi_tablex = 'SLG' and slg.pi_postex = ar_donlog and slg.pi_lang = 'RU'))" & _
                        " where ar_cproin = '" & wsForm.Cells(i, cProductCode) & "' and to_number(ar_ilogis) = " & CStr(wsForm.Cells(i, cNewOptions + 2)) & " group by ar_cproin"
                ' Debug.Print strQry
                rsOra.Open strQry, cnORA
                strGolduzel = ""
                If Not rsOra.bof Then
                    Do While Not rsOra.EOF
                        'If rsOra!uzel <> Trim(str) Then
                        strGolduzel = strGolduzel & rsOra!uzel & " "
                        'End If
                        rsOra.movenext
                    Loop

                End If
                rsOra.Close
            End If
            If IsError(Application.Match(CStr(Trim(str)), Array("94010", "94009", "94011"), 0)) Then
                '''            Do While Not rsOra.EOF
                '''            'If rsOra!uzel <> Trim(str) Then
                '''               strGolduzel = strGolduzel & rsOra!uzel & " "
                '''            'End If
                '''            rsOra.movenext
                '''            Loop
                ' For Each uzelform In Split(Trim(str), ",")
                If InStr(1, Trim(strGolduzel), Trim(str)) = 0 Then
                    ' If DictVirtUzel.exists(CStr(wsForm.Cells(i, cProductCode) & wsForm.Cells(i, cSupplierCode) & wsForm.Cells(i, cNewOptions + 2))) = False Then
                    If find70995 = False Then
                        zaendrow = wsForm.Cells(wsForm.Rows.Count, 2).End(xlUp).Row + 1
                        wsForm.Range(wsForm.Cells(i, 1), wsForm.Cells(i, 30)).Copy
                        wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).PasteSpecial
                        wsForm.Cells(zaendrow, 2) = "Открытие"
                        wsForm.Cells(zaendrow, 5) = Trim(str)
                        strQry = "select FOULIBL from CS.FOUDGENE where FOUCNUF = '" & Trim(str) & "'"
                        rsOra.Open strQry, cnORA
                        If Not rsOra.bof Then
                            wsForm.Cells(zaendrow, 6) = rsOra!FOULIBL
                        End If
                        rsOra.Close
                        wsForm.Cells(zaendrow, 7) = "70995"
                        'wsForm.Cells(ZAEndRow, 8) = "Виртуальный сайт"
                        wsForm.Range(wsForm.Cells(zaendrow, 9), wsForm.Cells(zaendrow, 17)).ClearContents
                        wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
                        wsForm.Cells(zaendrow, 19) = CDate(wsForm.Cells(zaendrow, 18)) + 90
                        wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
                        wsForm.Cells(zaendrow, 22) = "1"
                        wsForm.Cells(zaendrow, 23) = "CC_" & Trim(wsForm.Cells(zaendrow, 5))
                        wsForm.Range(wsForm.Cells(zaendrow, 24), wsForm.Cells(zaendrow, 30)).ClearContents
                        ''' rsOra.Close
                        strQry = "select nvl((select distinct '1' from cs.artul u where u.arucinr = artcinr and ARUTYPUL = 21),'0') LE from cs.artrac where artcexr = '" & wsForm.Cells(i, cProductCode) & "'"
                        rsOra.Open strQry, cnORA
                        If Not rsOra.bof Then
                            If rsOra!LE = "0" Then
                                wsForm.Cells(zaendrow, 21) = "41"
                                wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 255, 0)
                                zH = errFunction(zH, zV, i, "Заполнена строка для прогрузки Коробки в сток. При необходмости поменяйте на ""Штуку"" значение ЛЕ")
                                errcount = 1
                            ElseIf rsOra!LE = "1" Then
                                wsForm.Cells(zaendrow, 21) = ""
                                wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 0, 0)
                                zH = errFunction(zH, zV, i, "У ЛВ есть внутр.упаковка. Добавлена строка для прогрузки ЛВ в сток.  Заполните ЛЕ")
                                errcount = 1
                            End If
                        End If
                        rsOra.Close
                        wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 20)).Locked = True
                        wsForm.Range(wsForm.Cells(zaendrow, 22), wsForm.Cells(zaendrow, 30)).Locked = True
                        wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).Interior.Color = RGB(253, 233, 217)

                    End If
                End If
                'Next uzelform
            Else
                If dictFUZ.exists(Trim(str)) Then
                    OraUzel = strGolduzel    'rsOra!uzel
                    For Each varItem In Split(dictFUZ.Item(Trim(str)), ",")
                        If InStr(1, OraUzel, Trim(varItem)) = 0 Then
                            'If DictVirtUzel.exists(CStr(wsForm.Cells(i, cProductCode) & wsForm.Cells(i, cSupplierCode) & wsForm.Cells(i, cNewOptions + 2))) = False Then
                            If find70995 = False Then
                                zaendrow = wsForm.Cells(wsForm.Rows.Count, 2).End(xlUp).Row + 1
                                wsForm.Range(wsForm.Cells(i, 1), wsForm.Cells(i, 30)).Copy
                                wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).PasteSpecial
                                wsForm.Cells(zaendrow, 2) = "Открытие"
                                wsForm.Cells(zaendrow, 5) = Trim(varItem)
                                strQry = "select FOULIBL from CS.FOUDGENE where FOUCNUF = '" & Trim(varItem) & "'"
                                rsOra.Open strQry, cnORA
                                If Not rsOra.bof Then
                                    wsForm.Cells(zaendrow, 6) = rsOra!FOULIBL
                                End If
                                rsOra.Close
                                wsForm.Cells(zaendrow, 7) = "70995"
                                'wsForm.Cells(ZAEndRow, 8) = "Виртуальный сайт"
                                wsForm.Range(wsForm.Cells(zaendrow, 9), wsForm.Cells(zaendrow, 17)).ClearContents
                                wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
                                wsForm.Cells(zaendrow, 19) = CDate(wsForm.Cells(zaendrow, 18)) + 90
                                wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
                                wsForm.Cells(zaendrow, 22) = "1"
                                wsForm.Cells(zaendrow, 23) = "CC_" & Trim(wsForm.Cells(zaendrow, 5))
                                wsForm.Range(wsForm.Cells(zaendrow, 24), wsForm.Cells(zaendrow, 30)).ClearContents
                                ''' rsOra.Close
                                strQry = "select nvl((select distinct '1' from cs.artul u where u.arucinr = artcinr and ARUTYPUL = 21),'0') LE from cs.artrac where artcexr = '" & wsForm.Cells(i, cProductCode) & "'"
                                rsOra.Open strQry, cnORA
                                If Not rsOra.bof Then
                                    If rsOra!LE = "0" Then
                                        wsForm.Cells(zaendrow, 21) = "41"
                                        wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 255, 0)
                                        zH = errFunction(zH, zV, i, "Заполнена строка для прогрузки Коробки в сток. При необходмости поменяйте на ""Штуку"" значение ЛЕ")
                                        errcount = 1
                                    ElseIf rsOra!LE = "1" Then
                                        wsForm.Cells(zaendrow, 21) = ""
                                        wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 0, 0)
                                        zH = errFunction(zH, zV, i, "У ЛВ есть внутр.упаковка. Добавлена строка для прогрузки ЛВ в сток.  Заполните ЛЕ")
                                        errcount = 1
                                    End If
                                End If

                                wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 20)).Locked = True
                                wsForm.Range(wsForm.Cells(zaendrow, 22), wsForm.Cells(zaendrow, 30)).Locked = True
                                wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).Interior.Color = RGB(253, 233, 217)
                                rsOra.Close
                            End If
                        End If

                    Next
                End If
            End If
            '           If dictFUZ.exists(Trim(str)) Then
            '           OraUzel = rsOra!uzel
            '              For Each varItem In Split(dictFUZ.Item(Trim(str)), ",")
            '                               If InStr(1, OraUzel, Trim(varItem)) = 0 Then
            '                 If DictVirtUzel.exists(CStr(wsForm.Cells(i, cProductCode) & wsForm.Cells(i, cSupplierCode) & wsForm.Cells(i, cNewOptions + 2))) = False Then
            '                                zaendrow = wsForm.Cells(wsForm.Rows.Count, 2).End(xlUp).Row + 1
            '                                 wsForm.Range(wsForm.Cells(zaendrow - 1, 1), wsForm.Cells(zaendrow - 1, 30)).Copy
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).PasteSpecial
            '                                 wsForm.Cells(zaendrow, 2) = "Открытие"
            '                                 wsForm.Cells(zaendrow, 5) = Trim(varItem)
            '                                 wsForm.Cells(zaendrow, 7) = "70995"
            '                                 'wsForm.Cells(ZAEndRow, 8) = "Виртуальный сайт"
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 9), wsForm.Cells(zaendrow, 17)).ClearContents
            '                                 wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
            '                                 wsForm.Cells(zaendrow, 19) = CDate(wsForm.Cells(zaendrow, 18)) + 90
            '                                 wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
            '                                 wsForm.Cells(zaendrow, 22) = "1"
            '                                 wsForm.Cells(zaendrow, 23) = "CC_" & Trim(wsForm.Cells(zaendrow, 5))
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 24), wsForm.Cells(zaendrow, 30)).ClearContents
            '                                 rsOra.Close
            '                                 strQry = "select nvl((select distinct '1' from cs.artul u where u.arucinr = artcinr and ARUTYPUL = 21),'0') LE from cs.artrac where artcexr = '" & wsForm.Cells(i, cProductCode) & "'"
            '                                 rsOra.Open strQry, cnORA
            '                                 If Not rsOra.bof Then
            '                                   If rsOra!LE = "0" Then
            '                                      wsForm.Cells(zaendrow, 21) = "41"
            '                                      wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 255, 0)
            '                                      zH = errFunction(zH, zV, i, "Заполнена строка для прогрузки Коробки в сток. При необходмости поменяйте на ""Штуку"" значение ЛЕ")
            '                                      errcount = 1
            '                                   ElseIf rsOra!LE = "1" Then
            '                                      wsForm.Cells(zaendrow, 21) = ""
            '                                      wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 0, 0)
            '                                      zH = errFunction(zH, zV, i, "У ЛВ есть внутр.упаковка. Добавлена строка для прогрузки ЛВ в сток.  Заполните ЛЕ")
            '                                      errcount = 1
            '                                   End If
            '                                 End If
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).Locked = True
            '                                ' rsOra.Close
            '                 End If
            '              End If
            '
            '              Next
            '           End If
            '           Else
            '             If dictFUZ.exists(Trim(str)) Then
            '             ' OraUzel = rsOra!uzel
            '              For Each varItem In Split(dictFUZ.Item(Trim(str)), ",")
            '                              ' If InStr(1, OraUzel, Trim(varItem)) = 0 Then
            '                 If DictVirtUzel.exists(CStr(wsForm.Cells(i, cProductCode) & wsForm.Cells(i, cSupplierCode) & wsForm.Cells(i, cNewOptions + 2))) = False Then
            '                                zaendrow = wsForm.Cells(wsForm.Rows.Count, 2).End(xlUp).Row + 1
            '                                 wsForm.Range(wsForm.Cells(zaendrow - 1, 1), wsForm.Cells(zaendrow - 1, 30)).Copy
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).PasteSpecial
            '                                 wsForm.Cells(zaendrow, 2) = "Открытие"
            '                                 wsForm.Cells(zaendrow, 5) = Trim(varItem)
            '                                 wsForm.Cells(zaendrow, 7) = "70995"
            '                                 'wsForm.Cells(ZAEndRow, 8) = "Виртуальный сайт"
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 9), wsForm.Cells(zaendrow, 17)).ClearContents
            '                                 wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
            '                                 wsForm.Cells(zaendrow, 19) = CDate(wsForm.Cells(zaendrow, 18)) + 90
            '                                 wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
            '                                 wsForm.Cells(zaendrow, 22) = "1"
            '                                 wsForm.Cells(zaendrow, 23) = "CC_" & Trim(wsForm.Cells(zaendrow, 5))
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 24), wsForm.Cells(zaendrow, 30)).ClearContents
            '                                 rsOra.Close
            '                                 strQry = "select nvl((select distinct '1' from cs.artul u where u.arucinr = artcinr and ARUTYPUL = 21),'0') LE from cs.artrac where artcexr = '" & wsForm.Cells(i, cProductCode) & "'"
            '                                 rsOra.Open strQry, cnORA
            '                                 If Not rsOra.bof Then
            '                                   If rsOra!LE = "0" Then
            '                                      wsForm.Cells(zaendrow, 21) = "41"
            '                                      wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 255, 0)
            '                                      zH = errFunction(zH, zV, i, "Заполнена строка для прогрузки Коробки в сток. При необходмости поменяйте на ""Штуку"" значение ЛЕ")
            '                                      errcount = 1
            '                                   ElseIf rsOra!LE = "1" Then
            '                                      wsForm.Cells(zaendrow, 21) = ""
            '                                      wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 0, 0)
            '                                      zH = errFunction(zH, zV, i, "У ЛВ есть внутр.упаковка. Добавлена строка для прогрузки ЛВ в сток.  Заполните ЛЕ")
            '                                      errcount = 1
            '                                   End If
            '                                 End If
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).Locked = True
            '                                ' rsOra.Close
            '                 'End If
            '                End If
            '               Next
            '            Else
            '
            '               'If InStr(1, OraUzel, Trim(str)) = 0 Then
            '                 If DictVirtUzel.exists(CStr(wsForm.Cells(i, cProductCode) & wsForm.Cells(i, cSupplierCode) & wsForm.Cells(i, cNewOptions + 2))) = False Then
            '                                zaendrow = wsForm.Cells(wsForm.Rows.Count, 2).End(xlUp).Row + 1
            '                                 wsForm.Range(wsForm.Cells(zaendrow - 1, 1), wsForm.Cells(zaendrow - 1, 30)).Copy
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).PasteSpecial
            '                                 wsForm.Cells(zaendrow, 2) = "Открытие"
            '                                 wsForm.Cells(zaendrow, 5) = Trim(str)
            '                                 wsForm.Cells(zaendrow, 7) = "70995"
            '                                 'wsForm.Cells(ZAEndRow, 8) = "Виртуальный сайт"
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 9), wsForm.Cells(zaendrow, 17)).ClearContents
            '                                 wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
            '                                 wsForm.Cells(zaendrow, 19) = CDate(wsForm.Cells(zaendrow, 18)) + 90
            '                                 wsForm.Cells(zaendrow, 19).NumberFormat = "dd.mm.yyyy;@"
            '                                 wsForm.Cells(zaendrow, 22) = "1"
            '                                 wsForm.Cells(zaendrow, 23) = "CC_" & Trim(wsForm.Cells(zaendrow, 5))
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 24), wsForm.Cells(zaendrow, 30)).ClearContents
            '                                 rsOra.Close
            '                                 strQry = "select nvl((select distinct '1' from cs.artul u where u.arucinr = artcinr and ARUTYPUL = 21),'0') LE from cs.artrac where artcexr = '" & wsForm.Cells(i, cProductCode) & "'"
            '                                 rsOra.Open strQry, cnORA
            '                                 If Not rsOra.bof Then
            '                                   If rsOra!LE = "0" Then
            '                                      wsForm.Cells(zaendrow, 21) = "41"
            '                                      wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 255, 0)
            '                                      zH = errFunction(zH, zV, i, "Заполнена строка для прогрузки Коробки в сток. При необходмости поменяйте на ""Штуку"" значение ЛЕ")
            '                                      errcount = 1
            '                                   ElseIf rsOra!LE = "1" Then
            '                                      wsForm.Cells(zaendrow, 21) = ""
            '                                      wsForm.Cells(zaendrow, 21).Interior.Color = RGB(255, 0, 0)
            '                                      zH = errFunction(zH, zV, i, "У ЛВ есть внутр.упаковка. Добавлена строка для прогрузки ЛВ в сток.  Заполните ЛЕ")
            '                                      errcount = 1
            '                                   End If
            '                                 End If
            '                                 wsForm.Range(wsForm.Cells(zaendrow, 1), wsForm.Cells(zaendrow, 30)).Locked = True
            '                                ' rsOra.Close
            '                 End If
            '              'End If
            '            End If
            '''          End If
            '''
            '''
            '''        rsOra.Close
            '''   End If
        End If
    Next str
End Sub

Sub clearsheet()
    Set wsForm = Worksheets("Форма_A02ZA")
    'zaendrow = wsForm.Cells(wsForm.Rows.Count, 2).End(xlUp).Row + 1
    zaendrow = WorksheetFunction.Max(wsForm.Cells(wsForm.Rows.Count, 2).End(xlUp).Row, wsForm.Cells(wsForm.Rows.Count, 3).End(xlUp).Row, wsForm.Cells(wsForm.Rows.Count, 5).End(xlUp).Row, _
            wsForm.Cells(wsForm.Rows.Count, 7).End(xlUp).Row)
    wsForm.Unprotect Password:="1347"
    wsForm.Range("B10:AD" & zaendrow).ClearContents
    wsForm.Range("B10:AD" & zaendrow).Interior.Color = xlNone
    wsForm.Range("B10:AD" & zaendrow).Borders.LineStyle = False
    wsForm.Range("B10:AD" & zaendrow).Locked = False
End Sub

Public Function find70995()
    Dim varItemh
    find70995 = False
    If IsError(Application.Match(CStr(Trim(str)), Array("94010", "94009", "94011"), 0)) Then
        If DictVirtUzel.exists(CStr(wsForm.Cells(i, cProductCode) & Trim(str) & wsForm.Cells(i, cNewOptions + 2))) = True Then find70995 = True

    Else
        For Each varItemh In Split(dictFUZ.Item(Trim(str)), ",")
            If DictVirtUzel.exists(CStr(wsForm.Cells(i, cProductCode) & varItemh & wsForm.Cells(i, cNewOptions + 2))) = True Then find70995 = True
        Next
    End If
End Function

Public Function findnaim(str)
    Dim cRow
    Dim cColumn
    findnaim = ""
    'wsTerra.Cells(wsTerra.Rows.Count, 7).End(xlUp).Row
    With wsTerra.Cells
        cRow = .Find(What:=Trim(str), LookAt:=xlWhole).Row
        cColumn = .Find(What:=Trim(str), LookAt:=xlWhole).Column
        findnaim = .Cells(cRow, cColumn + 1)
    End With
End Function

Public Sub AC_ALCKK_()
    '------------проверка АЦ внешних поствщиков при условии поставки на Склад------------------
    errstr = ""
    If Left(TypePost, 1) = "1" And wsForm.Cells(i, cItemName).Value2 Like "*РЦ*" Then

        strQry = " select count(adrcep) countadrcep from (select distinct foucnuf as kodpost,(pc1.lirnfilf) as adrcep" & _
                " from  cs.lienregl pc1 inner join cs.foudgene on pc1.lircfin=foucfin and foutype =1 and foucnuf not in ('DI99999','8888888')" & _
                " left join cs.lienserv on liscfin=pc1.lircfin and pc1.lirccin=lisccin and pc1.lirsite=lissite" & _
                " and pc1.lirnfilf=lisnfilf and lisdfin>trunc(sysdate) left join cs.foufilie on ffinfilf = pc1.lirnfilf and FFICFIN = pc1.lircfin" & _
                " where (substr(pc1.lirsite,1,3)='700') and pc1.lirdfin > trunc(sysdate) and foucnuf = '" & wsForm.Cells(i, cSupplierCode).Value2 & "')"
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            If rsOra!countadrcep > 1 Then
                rsOra.Close
                strQry = "select distinct foucnuf as kodpost, RTRIM(xmlagg(xmlelement(E,lirsite,',')).extract('//text()'),',') as sait from  cs.lienregl pc1 inner join cs.fouccom on pc1.lirccin = fccccin" & _
                        " inner join cs.foudgene on pc1.lircfin=foucfin and foutype =1 and foucnuf not in ('DI99999','8888888')" & _
                        " left join cs.lienserv on liscfin=pc1.lircfin and pc1.lirccin=lisccin and pc1.lirsite=lissite and pc1.lirnfilf=lisnfilf and lisdfin>trunc(sysdate)" & _
                        " left join cs.foufilie on ffinfilf = pc1.lirnfilf and FFICFIN = pc1.lircfin where (substr(pc1.lirsite,1,3)='700')" & _
                        " and pc1.lirsite not in (select relid from cs.resrel where level=2 and trunc(sysdate) between RELDDEB and RELDFIN" & _
                        " start with relpere in ('94004') connect by relpere=prior relid) and pc1.lirdfin > trunc(sysdate)" & _
                        " and ((FFILIBL is null or (FFILIBL not like '%УФО%' and FFILIBL not like '%ЦФО%' and FFILIBL not like '%СЗФО%')) or (pc1.lirsite in ('70034') and FFILIBL like '%УФО%') or (pc1.lirsite = '70003' and FFILIBL like '%СЗФО%') or (pc1.lirsite in ('70004','70007','70011','70035') and FFILIBL like '%ЦФО%'))" & _
                        " and foucnuf = '" & wsForm.Cells(i, cSupplierCode).Value2 & "' and FCCNUM = '" & wsForm.Cells(i, cNewOptions + 5) & "' and pc1.lirnfilf = '" & wsForm.Cells(i, cNewOptions + 4).Value2 & "' group by foucnuf"
                rsOra.Open strQry, cnORA
                If Not rsOra.bof Then
                    strGolduzel = rsOra!sait
                    For Each str In Split(Application.Trim(Replace(Replace(Replace(wsForm.Cells(i, cItemCode).Value, ";", ","), ",", ", "), ",", " ,")), ",")    'узел
                        If IsError(Application.Match(CStr(Trim(str)), Array("94010", "94009", "94011"), 0)) Then
                            If InStr(1, Trim(strGolduzel), Trim(str)) = 0 Then
                                errstr = errstr & Trim(str) & ","
                            End If
                        Else
                            If dictFUZ.exists(Trim(str)) Then
                                For Each varItem In Split(dictFUZ.Item(Trim(str)), ",")
                                    If InStr(1, strGolduzel, Trim(varItem)) = 0 Then
                                        errstr = errstr & Trim(varItem) & ","
                                    End If
                                Next varItem
                            End If
                        End If
                    Next str
                End If

            End If
        End If
        rsOra.Close
        If errstr <> "" Then zH = errFunction(zH, zV, i, "У поставщика несколько адр.цепочек. Для РЦ " & Left(errstr, Len(errstr) - 1) & " цепочка выбрана некорректно. Создать заказ будет невозможно")

    End If

    '-------------------- Если по товару + лв + РЦ в старой или новой настройке уже есть настроенная твз с к.д. больше сегодня
    If (wsForm.Cells(i, cActionType).Value2 = "Изменение" Or wsForm.Cells(i, cActionType).Value2 = "Открытие") And Left(TypePost, 1) = "3" Then    'And wsForm.Cells(i, cItemName).Value2 Like "*РЦ*" Then
        ' For Each str In Split(Application.Trim(Replace(Replace(Replace(wsForm.Cells(i, cItemCode).Value, ";", ","), ",", ", "), ",", " ,")), ",") 'узел
        'If IsError(Application.Match(CStr(Trim(str)), Array("94010", "94009", "94011"), 0)) Then
        If Len(wsForm.Cells(i, cOldOptions + 1)) > 0 Then
            strQry = "select b1.arlcexr tovar1, cs.pkstrucobj.get_desc(1, b1.arlcinr, 'RU') opistovar1, b1.arlcexvl as lv1, b2.arlcexr as tovar2, cs.pkstrucobj.get_desc(1, b2.arlcinr, 'RU') opistovar2," & _
                    " b2.arlcexvl as lv2, arrsite As uzel from cs.artrempl inner join cs.artul a1 on arrcinlo = a1.arucinl inner join cs.artvl b1 on arrseqvlo = b1.arlseqvl" & _
                    " inner join cs.artul a2 on arrcinlr = a2.arucinl inner join cs.artvl b2 on arrseqvlr = b2.arlseqvl" & _
                    " where arrsite = '" & wsForm.Cells(i, cSupplierCode) & "' and (b1.arlcexr = '" & wsForm.Cells(i, cProductCode) & "' or b2.arlcexr = '" & wsForm.Cells(i, cProductCode) & "')" & _
                    " and (b2.arlcexvl = '" & wsForm.Cells(i, cOldOptions + 1) & "' or b1.arlcexvl = '" & wsForm.Cells(i, cOldOptions + 1) & "') and arrdfin >= trunc(sysdate)"
            'Debug.Print strQry
            rsOra.Open strQry, cnORA
            If Not rsOra.bof Then
                zH = errFunction(zH, zV, i, "По товару+ЛВ+РЦ (" & wsForm.Cells(i, cProductCode) & "+" & wsForm.Cells(i, cOldOptions + 1) & "+" & Trim(wsForm.Cells(i, cSupplierCode)) & ") в блоке закрытие есть настроенная ТВЗ")
            End If
            rsOra.Close
        End If
        If Len(wsForm.Cells(i, cNewOptions + 2)) > 0 Then
            strQry = "select b1.arlcexr tovar1, cs.pkstrucobj.get_desc(1, b1.arlcinr, 'RU') opistovar1, b1.arlcexvl as lv1, b2.arlcexr as tovar2, cs.pkstrucobj.get_desc(1, b2.arlcinr, 'RU') opistovar2," & _
                    " b2.arlcexvl as lv2, arrsite As uzel from cs.artrempl inner join cs.artul a1 on arrcinlo = a1.arucinl inner join cs.artvl b1 on arrseqvlo = b1.arlseqvl" & _
                    " inner join cs.artul a2 on arrcinlr = a2.arucinl inner join cs.artvl b2 on arrseqvlr = b2.arlseqvl" & _
                    " where arrsite = '" & wsForm.Cells(i, cSupplierCode) & "' and (b1.arlcexr = '" & wsForm.Cells(i, cProductCode) & "' or b2.arlcexr = '" & wsForm.Cells(i, cProductCode) & "')" & _
                    " and (b2.arlcexvl = '" & wsForm.Cells(i, cNewOptions + 2) & "' or b1.arlcexvl = '" & wsForm.Cells(i, cNewOptions + 2) & "') and arrdfin >= trunc(sysdate)"
            rsOra.Open strQry, cnORA
            If Not rsOra.bof Then
                zH = errFunction(zH, zV, i, "По товару+ЛВ+РЦ (" & wsForm.Cells(i, cProductCode) & "+" & wsForm.Cells(i, cNewOptions + 2) & "+" & Trim(wsForm.Cells(i, cSupplierCode)) & ") в блоке открытие есть настроенная ТВЗ")
            End If
            rsOra.Close
        End If
        '''  Else
        '''     If dictFUZ.exists(Trim(str)) Then
        '''         For Each varItem In Split(dictFUZ.Item(Trim(str)), ",")
        '''     If Len(wsForm.Cells(i, cOldOptions + 1)) > 0 Then
        '''        strQry = "select b1.arlcexr tovar1, cs.pkstrucobj.get_desc(1, b1.arlcinr, 'RU') opistovar1, b1.arlcexvl as lv1, b2.arlcexr as tovar2, cs.pkstrucobj.get_desc(1, b2.arlcinr, 'RU') opistovar2," & _
         '''        " b2.arlcexvl as lv2, arrsite As uzel from cs.artrempl inner join cs.artul a1 on arrcinlo = a1.arucinl inner join cs.artvl b1 on arrseqvlo = b1.arlseqvl" & _
         '''        " inner join cs.artul a2 on arrcinlr = a2.arucinl inner join cs.artvl b2 on arrseqvlr = b2.arlseqvl" & _
         '''        " left join cs.foucatalog on FCLSEQVL = b1.arlseqvl or FCLSEQVL = b2.arlseqvl where arrsite = '" & Trim(varItem) & "' and (b1.arlcexr = '" & wsForm.Cells(i, cProductCode) & "' or b2.arlcexr = '" & wsForm.Cells(i, cProductCode) & "')" & _
         '''        " and (b2.arlcexvl = '" & wsForm.Cells(i, cOldOptions + 1) & "' or b1.arlcexvl = '" & wsForm.Cells(i, cOldOptions + 1) & "') and arrdfin >= trunc(sysdate)"
        '''         rsOra.Open strQry, cnORA
        '''             If Not rsOra.bof Then
        '''                zH = errFunction(zH, zV, i, "По товару+ЛВ+РЦ (" & wsForm.Cells(i, cProductCode) & "+" & wsForm.Cells(i, cOldOptions + 1) & "+" & Trim(varItem) & ") в блоке закрытие есть настроенная ТВЗ")
        '''             End If
        '''         rsOra.Close
        '''     End If
        '''     If Len(wsForm.Cells(i, cNewOptions + 2)) > 0 Then
        '''        strQry = "select b1.arlcexr tovar1, cs.pkstrucobj.get_desc(1, b1.arlcinr, 'RU') opistovar1, b1.arlcexvl as lv1, b2.arlcexr as tovar2, cs.pkstrucobj.get_desc(1, b2.arlcinr, 'RU') opistovar2," & _
         '''        " b2.arlcexvl as lv2, arrsite As uzel from cs.artrempl inner join cs.artul a1 on arrcinlo = a1.arucinl inner join cs.artvl b1 on arrseqvlo = b1.arlseqvl" & _
         '''        " inner join cs.artul a2 on arrcinlr = a2.arucinl inner join cs.artvl b2 on arrseqvlr = b2.arlseqvl" & _
         '''        " left join cs.foucatalog on FCLSEQVL = b1.arlseqvl or FCLSEQVL = b2.arlseqvl where arrsite = '" & Trim(varItem) & "' and (b1.arlcexr = '" & wsForm.Cells(i, cProductCode) & "' or b2.arlcexr = '" & wsForm.Cells(i, cProductCode) & "')" & _
         '''        " and (b2.arlcexvl = '" & wsForm.Cells(i, cNewOptions + 2) & "' or b1.arlcexvl = '" & wsForm.Cells(i, cNewOptions + 2) & "') and arrdfin >= trunc(sysdate)"
        '''         rsOra.Open strQry, cnORA
        '''             If Not rsOra.bof Then
        '''                zH = errFunction(zH, zV, i, "По товару+ЛВ+РЦ (" & wsForm.Cells(i, cProductCode) & "+" & wsForm.Cells(i, cNewOptions + 2) & "+" & Trim(varItem) & ") в блоке открытие есть настроенная ТВЗ")
        '''             End If
        '''         rsOra.Close
        '''     End If
        '''         Next varItem
        '''     End If
        ''' End If
        ' Next str
    End If

    '-------------проверка корректности КК по алкоголю
    finALC = 0: finPIV = 0: finPIVBA = 0
    If Left(TypePost, 1) = "3" And (wsForm.Cells(i, cItemName).Value2 Like "*МАГАЗИНЫ*" Or wsForm.Cells(i, cItemName).Value2 Like "*ДИКСИ*") And Len(wsForm.Cells(i, cNewOptions + 5)) > 0 Then
        strQry = "select FCCLIB, nvl(Priznak,'zero') Priznak from OK__TOV_KLASSIFIKATOR, CS.FOUCCOM where kodyr5 in (select nvl((select SOBCEXT from cs.strucrel  left join cs.strucobj s1 on OBJPERE=SOBCINT" & _
                " left join cs.tra_strucobj on tsobcint=SOBCINT and langue = 'RU' where trunc(sysdate) between objddeb and objdfin and objcint=artcinr),'Frozen') TK from cs.artrac where artcexr = '" & wsForm.Cells(i, cProductCode) & "') and FCCNUM = '" & wsForm.Cells(i, cNewOptions + 5).Value & "'"
        ' Debug.Print strQry
        rsOra.Open strQry, cnORA
        If Not rsOra.bof Then
            If rsOra!Priznak Like "*АЛК*" Then
                ' If InStr(1, Replace(rsOra!FCCLIB, "БЕЗАЛК", ""), "АЛК") = 0 Or InStr(1, Replace(rsOra!FCCLIB, "Б/А", ""), "АЛК") = 0 Then finALC = 1
                If InStr(1, rsOra!FCCLIB, "БЕЗАЛК") = 0 And InStr(1, rsOra!FCCLIB, "Б/А") = 0 Then
                    If InStr(1, rsOra!FCCLIB, "АЛК") = 0 Then finALC = 1
                    If finALC = 1 Then zH = errFunction(zH, zV, i, "Для лицензируемого алкоголя указан некорректный КК. Необходимо исправить")
                Else
                    zH = errFunction(zH, zV, i, "Для лицензируемого алкоголя указан некорректный КК. Необходимо исправить")
                End If
            End If
            If rsOra!Priznak Like "*ПИВО*" Then
                If InStr(1, rsOra!FCCLIB, "БЕЗАЛК") = 0 And InStr(1, rsOra!FCCLIB, "Б/А") = 0 Then
                    If InStr(1, rsOra!FCCLIB, "ПИВ") = 0 Then finPIV = 1
                    If finPIV = 1 Then zH = errFunction(zH, zV, i, "Для Пива указан некорректный КК. Необходимо исправить")
                Else
                    zH = errFunction(zH, zV, i, "Для Пива указан некорректный КК. Необходимо исправить")
                End If
            End If
            If InStr(1, CStr(rsOra!Priznak), "ПИВО") = 0 And InStr(1, CStr(rsOra!Priznak), "АЛК") = 0 Then
                If (rsOra!FCCLIB Like "*ПИВ*" Or rsOra!FCCLIB Like "*АЛК*") And (Not rsOra!FCCLIB Like "*БЕЗАЛК*" And Not rsOra!FCCLIB Like "*Б/А*") Then zH = errFunction(zH, zV, i, "Товар не является Алкоголем. Указан алкогольный КК. Необходимо исправить")
            End If
        End If
        rsOra.Close
    End If
End Sub


}

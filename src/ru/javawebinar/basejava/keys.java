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

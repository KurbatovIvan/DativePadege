' DativeCaseInCell - дательный падеж от ФИО, записанных

'                    в текущей ячейке MS Excel

'

'   Текущая ячейка должна содержать следующую информацию:

'   фамилия, имя и отчество (именно в таком порядке)



Public Sub DativeCaseInCell()

    Dim s1 As String, s2 As String, s3 As String, cel As Range

' Цикл по выделенным
    For Each cel In Selection
     s1 = ChooseWord(cel.Value, 1)
     s2 = ChooseWord(cel.Value, 2)
     s3 = ChooseWord(cel.Value, 3)

'    If Len(s1) = 0 Or Len(s2) = 0 Or Len(s3) = 0 Then Exit Sub

    Cells(cel.Row, cel.Column) = DativeCase(s1, s2, s3)
Next cel

End Sub



' DativeCase - формирование дательного падежа от ФИО

'

' Параметры: sSurname    - фамилия

'            sName       - имя

'            sPatronymic - отчество

'

' Результат: ФИО в дательном падеже



Private Function DativeCase(sSurname As String, sName As String, sPatronymic As String) As String

  Dim bMaleSex As Boolean

    

  bMaleSex = (Right(sPatronymic, 1) = "ч")

        

'   Фамилия



  If Len(sSurname) > 0 Then

    If bMaleSex Then

        Select Case Right(sSurname, 1)

            Case "о", "и", "я", "а"

                DativeCase = sSurname

            Case "й"

                DativeCase = Mid(sSurname, 1, Len(sSurname) - 2) + "ому"

            Case Else

                DativeCase = sSurname + "у"

        End Select

    Else

        Select Case Right(sSurname, 1)

            Case "о", "и", "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "х", "ц", "ч", "ш", "щ", "ь"

                DativeCase = sSurname

            Case "я"

                DativeCase = Mid(sSurname, 1, Len(sSurname) - 2) & "ой"

            Case Else

                DativeCase = Mid(sSurname, 1, Len(sSurname) - 1) & "ой"

        End Select

    End If

    DativeCase = DativeCase & " "

  End If



'   Имя



  If Len(sName) > 0 Then

    If bMaleSex Then

        Select Case Right(sName, 1)

            Case "й", "ь"

                DativeCase = DativeCase & Mid(sName, 1, Len(sName) - 1) & "ю"

            Case Else

                DativeCase = DativeCase & sName & "у"

        End Select

    Else

        Select Case Right(sName, 1)

            Case "а", "я"

                If Mid(sName, Len(sName) - 1, 1) = "и" Then

                    DativeCase = DativeCase & Mid(sName, 1, Len(sName) - 1) & "и"

                Else

                    DativeCase = DativeCase & Mid(sName, 1, Len(sName) - 1) & "е"

                End If

            Case "ь"

                DativeCase = DativeCase & Mid(sName, 1, Len(sName) - 1) & "и"

            Case Else

                DativeCase = DativeCase & sName

        End Select

    End If

    DativeCase = DativeCase & " "

  End If



'   Отчество



  If Len(sPatronymic) > 0 Then

    If bMaleSex Then

            DativeCase = DativeCase & sPatronymic & "у"

    Else

            DativeCase = DativeCase & Mid(sPatronymic, 1, Len(sPatronymic) - 1) & "е"

    End If

  End If

End Function



' ChooseWord - выделение i-го слова из строки

'

' Параметры: sString - строка

'            iNum    - номер слова (1, 2,...)

'

' Результат: i-е слово или пустая строка, если такого слова нет





Private Function ChooseWord(sString As String, iNum As Integer)

    Dim sTemp As String

    Dim i As Integer, iPos As Integer

    

    sTemp = Trim(sString)

    For i = 1 To iNum - 1

        iPos = InStr(sTemp, " ")

        If iPos = 0 Then iPos = Len(sTemp)

        sTemp = Trim(Right(sTemp, Len(sTemp) - iPos))

    Next i

    iPos = InStr(sTemp, " ")

    If iPos = 0 Then iPos = Len(sTemp)

    ChooseWord = Trim(Left(sTemp, iPos))

    

End Function


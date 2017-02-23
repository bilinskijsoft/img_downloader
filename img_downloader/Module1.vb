Imports System.Net
Imports System.Text.RegularExpressions

Module Module1
    Const TARGET_URL As String = "http://kowo.me/"

    Const QUOTE As Char = Chr(34)
    Const DOWNLOAD_FOLDER As String = "downloaded\"

    Dim visited As String() = {" "} 'Список посещенных ссылок

    Sub Main()
        iterateUrls(TARGET_URL) 'Начнем работу
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Загрузка файла
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sub downloadFile(url As String, targetPath As String)

        Dim download As WebClient = New WebClient
        'Пусть думают что мы не бот
        download.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36")

        Try
            Console.Write("Found img: " & targetPath & " - Downloading...")
            'Если файл уже существует то не будем его загружать
            If IO.File.Exists(DOWNLOAD_FOLDER & targetPath) Then
                Console.WriteLine(" File exists!")
                Exit Sub
            End If
            download.DownloadFile(New Uri(url), DOWNLOAD_FOLDER & targetPath)
            Console.WriteLine(" OK!")
        Catch ex As Exception
            Console.WriteLine(" ERROR:" & ex.Message)
        End Try
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Загрузка HTML кода страницы
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Function downloadHTML(url As String) As String
        Dim HTML As String
        Dim download As WebClient = New WebClient
        'Пусть думают что мы не бот
        download.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36")
        Try
        Console.Write(url & " - Requesting...")
            'Если нашли в строке расширение картинки
            If url.Split(".")(url.Split(".").Count - 1) = "png" Or url.Split(".")(url.Split(".").Count - 1) = "jpg" Then
                Console.WriteLine(" Is image! Downloading...")
                Dim fileName As String()
                fileName = Split(url, "/")
                downloadFile(url, fileName(fileName.Count - 1))
                HTML = ""
                'Чтобы не обрабатывать css и js
            ElseIf url.Split(".")(url.Split(".").Count - 1) = "css" Or url.Split(".")(url.Split(".").Count - 1) = "js" Then
                Console.WriteLine(" not a web page!")
                HTML = ""
            Else
                HTML = download.DownloadString(New Uri(url))
                Console.WriteLine(" OK!")
            End If
        Catch ex As Exception
            Console.WriteLine(" ERROR: " & ex.Message.ToString)
            Return ""
        End Try
        Return HTML
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Рекурсивная функция загрузки изображений по ссылкам
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sub iterateUrls(url As String)
        For Each Str As String In visited
            If Str.Contains(url) Then
                Exit Sub
            End If
        Next
        ReDim Preserve visited(0 To visited.Count)
        visited(visited.Count - 1) = url    'Добавляем ссылку в посещенные чтобы нбольше с ней не работать
        Console.Write("Download images from: ")
        downloadImgsFromUrl(getImgUrl(downloadHTML(url)), url) 'Загружаем картинки
        'Ищем ссылки на странице
        Console.Write("Search urls from: ")
        Dim match As MatchCollection = getUrls(downloadHTML(url))
        Dim i
        'Обходим все ссылки
        For i = 0 To match.Count - 1
            'Если нашли в списке посещенных текущюю ссылку то просто выходим
            Dim skip As Boolean = False
            For Each Str As String In visited
                If Str.Contains(Split(match.Item(i).ToString, "=")(1).Trim(QUOTE)) Then
                    skip = True
                End If
            Next
            If (skip) Then
                i = i + 1
            Else
                iterateUrls(Split(match.Item(i).ToString, "=")(1).Trim(QUOTE))
            End If
        Next
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Загружаем все изображения по ссылке
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sub downloadImgsFromUrl(match1 As MatchCollection, url As String)
        Dim j
        'Обходим все ссылки
        If match1.Count > 0 Then
            For j = 0 To match1.Count - 1
                Dim fileName As String()
                'Разбиваме строка по "/"
                fileName = Split(match1.Item(j).ToString.Trim(QUOTE), "/")
                'Если ссылка полная то просто загружаем
                If Left(match1.Item(j).ToString.Trim(QUOTE), 2) = "ht" Then
                    downloadFile(match1.Item(j).ToString.Trim(QUOTE), fileName(fileName.Count - 1))
                    'Если нехватает http добавляем и загружаем
                ElseIf Left(match1.Item(j).ToString.Trim(QUOTE), 2) = "//" Then
                    downloadFile("http:" & match1.Item(j).ToString.Trim(QUOTE), fileName(fileName.Count - 1))
                    'Если ссілка вобще без адреса хоста - добавляем и загружаем
                Else
                    downloadFile("http://" & Split(url, "/")(2) & match1.Item(j).ToString.Trim(QUOTE), fileName(fileName.Count - 1))
                End If
            Next j
        End If

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''
    ' Полачаем ссылки из HTML кода
    ''''''''''''''''''''''''''''''''''''''''''
    Function getUrls(HTML As String) As MatchCollection
        Dim reHref As Regex
        Dim match As MatchCollection
        reHref = New Regex("href\s*=\s*""http(\s*|s)://(?:[""'](?<1>[^""']*)[""']|(?<1>\S+))""")
        match = reHref.Matches(HTML)
        Return match
    End Function
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Получаем ссылки на изображение из HTML кода
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Function getImgUrl(HTML As String) As MatchCollection
        Dim reHref As Regex
        Dim match As MatchCollection
        reHref = New Regex("""(http(\s*|s)://|\s*/)(?:[""'](?<1>[^""']*)[""']|(?<1>\S+))(.jpg|.png)""")
        match = reHref.Matches(HTML)
        Return match
    End Function

End Module
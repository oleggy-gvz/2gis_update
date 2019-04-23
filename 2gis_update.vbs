Option Explicit

Const prog = "2gis_update"

Const name_prog = "Обновление базы 2GIS v1.33"
Const url_update = "http://info.2gis.ru/novosibirsk/products/download"
' пример текста ссылки - ' http://download.2gis.com/arhives/2GISData_Novosibirsk-245.0.6.msi
Const pattern_begin_url = "http://download.2gis.com/arhives/2GISData_Novosibirsk-"
Const pattern_end_url = ".msi"

Const pattern_begin_ver = "база данных новосибирска</a></h2><div class=""downloads__itemdescription"">"
Const pattern_end_ver = ","

Dim file_ver_gdat, file_prog_log, is_err
file_ver_gdat = prog & ".ver"
file_prog_log = prog & ".log"
is_err = False

Const name_DB_in_msi = "Novosibirsk_DGDAT"
Const name_DB_base="Data_Novosibirsk.dgdat"

Const vbForReading = 1
Const vbForWriting = 2
Const vbForAppending = 8

Const exe_2gis="grym.exe"

Const MSG_FOND_NEW_VER_2GIS = "На сервере обнаружена другая версия базы данных 2ГИС: "
Const MSG_LOCAL_VER_2GIS = "Локальная версия базы 2ГИС: "
Const MSG_UPDATE_NOW = "Обновить сейчас перед запуском?"

Const MSG_CANT_2GIS = "Скрипт находится не в каталоге программы, отсутствует файл запуска 2GIS"
Const MSG_CONNECT = "Соедение с сервером"
Const MSG_CANT_CONNECT = "Нет соедения с сервером"
Const MSG_FIND_DB = "Поиск базы данных на сервере"
Const MSG_CANT_FIND_DB = "Невозможно найти базу данных на сервере"
Const MSG_FIND_VER_DB = "Ищем информацию о месяце базы данных"
Const MSG_CANT_FIND_VER_DB = "Не возможно найти информацию о месяце базы данных"
Const MSG_READE_OLD_VER_DB = "Читаем информацию о месяце локальной базы данных"
Const MSG_CANT_READE_OLD_VER_DB = "Невозможно прочитать информацию о месяце локальной базы данных"
Const MSG_DOWNLOAD = "Скачивание базы данных с сервера"
Const MSG_CANT_DOWNLOAD = "Невозможно скачать базу данных с сервера"
Const MSG_EXTRACT = "Извлечение базы данных из скаченного файла"
Const MSG_CANT_EXTRACT = "Невозможно извлеч базу данных из скаченного файла"
Const MSG_DETETE_MSI_DB = "Удаление файа установки новой базы данных"
Const MSG_CANT_DETETE_MSI_DB = "Невозможно удалить файл установки новой базы данных"
Const MSG_DETETE_OLD_DB = "Удаление старой базы данных"
Const MSG_CANT_DETETE_OLD_DB = "Невозможно удалить старую базу данных"
Const MSG_RENAME = "Переименовывание новой версии базы данных"
Const MSG_CANT_RENAME = "Невозможно переименовать новую версию базы данных"
Const MSG_WRITE_VER = "Сохранение информации об обновленной базе данных"
Const MSG_CANT_WRITE_VER = "Невозможно сохранить информацию об обновленной базе данных"
Const MSG_CANT_UPDATE_DB = "Не удалось обновить базу данных (подробности в файле лога). Будет запущен 2GIS со старой базой данных."

Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(file_prog_log) Then deleteFile(file_prog_log) ' если файл лога запуска существует - удалить

If Not update() Then: msg_time(MSG_CANT_UPDATE_DB)

If Not FSO.FileExists(exe_2gis) Then msg_err(MSG_CANT_2GIS): WScript.Quit ' если файла запуска 2GIS нет, выходим

Dim WshShell: Set WshShell = WScript.CreateObject("WScript.Shell")
Dim res: res = WshShell.Run(exe_2gis, 1, False)

'------------------------------------------------------------------------------------------------------------------------------------------------

Function update()
	update = False
	' соединение с интернетом
	msg_log(MSG_CONNECT)
	Dim html_text: html_text = getHtmlCodeWebPageFromUrl(url_update)
	If html_text ="" Then msg_log(MSG_CANT_CONNECT): Exit Function
	
	' проверяем наличие ссылки на БД
	msg_log(MSG_FIND_DB)
	Dim down_link: down_link = getTextBetween2Patt(html_text, pattern_begin_url, pattern_end_url) ' ищем ссылку на скачиание MSI файла
	If down_link ="" Then msg_log(MSG_CANT_FIND_DB): Exit Function
 	down_link = pattern_begin_url & down_link & pattern_end_url ' ссылка на скачивание файла MSI
 	
	' читаем версию БД
	msg_log(MSG_FIND_VER_DB)
	Dim ver_gdat: ver_gdat = getTextBetween2Patt(html_text, pattern_begin_ver, pattern_end_ver) ' ищем версию базы
	If ver_gdat ="" Then msg_log(MSG_CANT_FIND_VER_DB): Exit Function
	
	' читаем локальную версию БД
	msg_log(MSG_READE_OLD_VER_DB)
	Dim ver_gdat_local: ver_gdat_local = readTxtFromFile(file_ver_gdat)
	If ver_gdat_local = -1 Then ver_gdat_local = "<неизвестная>": msg_log(MSG_CANT_READE_OLD_VER_DB)
	'ver_gdat_local = Replace(ver_gdat_local, " ", "") ' удаляем все пробелы
	If ver_gdat_local = "" Then ver_gdat_local = "<неизвестная>": msg_log(MSG_CANT_READE_OLD_VER_DB)
	
	' выводим сообщение
	Dim key_press
	If StrComp(ver_gdat, ver_gdat_local, vbTextCompare) <> 0 Then ' если версии не совпадает, уведомляем
		key_press = MsgBox(MSG_FOND_NEW_VER_2GIS & ver_gdat & vbCrLf & MSG_LOCAL_VER_2GIS & ver_gdat_local & vbCrLf & MSG_UPDATE_NOW, _
				vbYesNo + vbInformation + vbDefaultButton2, name_prog) ' по умолчанию выбрано 'Нет'
	Else ' если версии собвадают, выходим
		update = True
	 	Exit Function
	End If
	If key_press = vbNo Then update = True: Exit Function ' если было нажато 'Нет' для обновления
	
	Dim name_msi: name_msi = name_DB_in_msi + ".msi" ' имя под которым будет сохранен файл MSI
	' скачиваем файл MSI по ссылке
	msg_log(MSG_DOWNLOAD & ": " & down_link & "," & name_msi)
 	If Not downloadFileByUrl(down_link, name_msi) Then msg_log(MSG_CANT_DOWNLOAD): Exit Function
 	
	' извлекаем из MSI файл БД
	msg_log(MSG_EXTRACT)
	If Not extractFileFromMsi(name_msi, "", name_DB_in_msi) Then msg_log(MSG_CANT_EXTRACT): Exit Function
	
	' удаляем MSI файл БД
	msg_log(MSG_DETETE_MSI_DB)
	If Not deleteFile(name_msi) Then msg_log(MSG_CANT_DETETE_MSI_DB): Exit Function
	
	' удаляем старую БД
	msg_log(MSG_DETETE_OLD_DB)
	If Not deleteFile(name_DB_base) Then msg_log(MSG_CANT_DETETE_OLD_DB): Exit Function
	
	' переименовываем файл новой БД
	msg_log(MSG_RENAME)
	If Not renameFile(name_DB_in_msi, name_DB_base) Then msg_log(MSG_CANT_RENAME): Exit Function
	
	' сохраняем информацию о новой версии БД в файл
	msg_log(MSG_WRITE_VER)
	If Not writeTxtToFile(file_ver_gdat, ver_gdat, vbForWriting) Then msg_log(MSG_CANT_WRITE_VER)
	update = True
End Function	

'------------------------------------------------------------------------------------------------------------------------------------------------

Sub msg_err(ByVal msg)
	MsgBox msg, vbCritical + vbOkOnly, name_prog & " - ошибка"
	msg_log(msg)
End Sub

Sub msg_time(ByVal msg)
	Dim WshShell: Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.Popup msg, 4
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------

Sub msg_log(ByVal msg)
	Dim res
	res = writeTxtToFile(file_prog_log, msg & vbCrLf, vbForAppending)
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------

Function getHtmlCodeWebPageFromUrl(ByVal url)
	Dim html, i_1, i_2
	getHtmlCodeWebPageFromUrl = ""
	' если параметр валюта не передавался в функцию, то возвращаем 0
	If url = "" Then Exit Function
	' формируем http запрос
	Dim oHttp
	On Error Resume Next
		Set oHttp = CreateObject("MSXML2.XMLHTTP")
		If Err.Number <> 0 Then Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
	On Error GoTo 0
	If oHttp Is Nothing Then Exit Function ' если объект запроса не создан, то выходим
	oHttp.Open "GET", url, False
	oHttp.Send ' отправляем запрос
	html = UCase(oHttp.responseText) ' сохраняем ответ на запрос (HTML текст сайта)
	Set oHttp = Nothing ' освобождаем объект запроса
	getHtmlCodeWebPageFromUrl = html
End Function

' получаем текст из строки между двух шаблонов
' str - искомая строка, pat1 - шаблон передний, pat2 - шаблон задний
' возвращаемое значение - текст находящийся между двух шаблонов, если нет, возвращает ""
Function getTextBetween2Patt(str, pat1, pat2)
	getTextBetween2Patt = ""
	Dim i_1: i_1 = InStr(1, str, pat1, vbTextCompare) ' ищем 1-е вхождение строки шаблона 1 без учета регистра
	If i_1 = 0 Then Exit Function ' если не найдено, выходим
	i_1 = i_1 + Len(pat1) ' сдвигаем на 1-ый символ после шаблона 1 
	Dim i_2: i_2 = InStr(i_1, str, pat2, vbTextCompare) ' ищем 1-ое вхождение шаблона 2 без учета регистра после того как найден шаблон 1
	If i_2 = 0 Then Exit Function ' если не найдено, выходим
	getTextBetween2Patt = Mid(str, i_1, i_2 - i_1) ' текст вырезки
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------

Function writeTxtToFile(ByVal filename, ByVal text, ByVal iomode)
	writeTxtToFile = False
	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")	
	On Error Resume Next
	Dim txtFile: Set txtFile = FSO.OpenTextFile(filename, iomode, True) ' True - создание файла если нет
	txtFile.Write text
	txtFile.Close
	If Err.Number <> 0 Then Exit Function
	On Error GoTo 0
	writeTxtToFile = True
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------

Function readTxtFromFile(ByVal filename)
 	readTxtFromFile = -1
 	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
 	On Error Resume Next
	Dim txtFile: Set txtFile = FSO.OpenTextFile(filename, vbForReading) ' 3-ий параметер False по умолчанию
	readTxtFromFile = txtFile.ReadAll
	txtFile.Close
	If Err.Number <> 0 Then Exit Function
	On Error GoTo 0
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------

Function downloadFileByUrl(url, save_file)
	'Const exe_downloader = "UrlFileDownloader.exe"
	Const exe_downloader = "D:\Programs\UrlFileDownloader\UrlFileDownloader.exe"
	downloadFileByUrl = False
	Dim res, cmd
	cmd = exe_downloader & " url=" & url & " file=" & save_file
	Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
	res = WshShell.Run(cmd, 1, True) ' True - ждем окончания выполнения 
	WScript.Sleep 1000
 	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
 	if Not FSO.FileExists(save_file) Then Exit Function
	downloadFileByUrl = True
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------

Function extractFileFromMsi(file_msi, path_dest, extract_file)
	Const exe_7zip = "%COMMANDER_PATH%\Utilites\7-Zip\7z.exe"
	extractFileFromMsi = False
	Dim res, cmd
 	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
 	if Not FSO.FileExists(file_msi) Then Exit Function ' если архиватора 7zip не найдено, выходим
	cmd = exe_7zip & " e " & file_msi & " " & extract_file
	' если указана директория извлечения, добавляем ее в параметер
	If Not (path_dest = "") Then cmd = cmd & " -o" & Chr(34) & path_dest & Chr(34) ' кавычки Chr(34), что бы учитывать пробелы
	If FSO.FileExists(extract_file) Then cmd = cmd & " -ao"
	Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
	res = WshShell.Run(cmd, 1, True) ' True - ждем окончания выполнения 
 	if Not FSO.FileExists(extract_file) Then Exit Function
	extractFileFromMsi = True
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------

Function deleteFile(ByVal path_file)
	deleteFile = False
	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
	if Not FSO.FileExists(path_file) Then: deleteFile = True: Exit Function
	On Error Resume Next
	FSO.DeleteFile (path_file)
	If Err.Number <> 0 Then Exit Function ' ошибка доступа к файлу
	On Error GoTo 0
	deleteFile = True
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------

Function renameFile(ByVal path_file, ByVal file_new)
	renameFile = False
	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
	Dim File: Set File = FSO.GetFile(path_file)
	On Error Resume Next
	File.Name = file_new
	If Err.Number <> 0 Then Exit Function ' ошибка доступа к файлу
	On Error GoTo 0
	renameFile = True
End Function
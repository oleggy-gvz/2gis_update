Option Explicit

Const prog = "2gis_update"

Const name_prog = "���������� ���� 2GIS v1.33"
Const url_update = "http://info.2gis.ru/novosibirsk/products/download"
' ������ ������ ������ - ' http://download.2gis.com/arhives/2GISData_Novosibirsk-245.0.6.msi
Const pattern_begin_url = "http://download.2gis.com/arhives/2GISData_Novosibirsk-"
Const pattern_end_url = ".msi"

Const pattern_begin_ver = "���� ������ ������������</a></h2><div class=""downloads__itemdescription"">"
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

Const MSG_FOND_NEW_VER_2GIS = "�� ������� ���������� ������ ������ ���� ������ 2���: "
Const MSG_LOCAL_VER_2GIS = "��������� ������ ���� 2���: "
Const MSG_UPDATE_NOW = "�������� ������ ����� ��������?"

Const MSG_CANT_2GIS = "������ ��������� �� � �������� ���������, ����������� ���� ������� 2GIS"
Const MSG_CONNECT = "�������� � ��������"
Const MSG_CANT_CONNECT = "��� �������� � ��������"
Const MSG_FIND_DB = "����� ���� ������ �� �������"
Const MSG_CANT_FIND_DB = "���������� ����� ���� ������ �� �������"
Const MSG_FIND_VER_DB = "���� ���������� � ������ ���� ������"
Const MSG_CANT_FIND_VER_DB = "�� �������� ����� ���������� � ������ ���� ������"
Const MSG_READE_OLD_VER_DB = "������ ���������� � ������ ��������� ���� ������"
Const MSG_CANT_READE_OLD_VER_DB = "���������� ��������� ���������� � ������ ��������� ���� ������"
Const MSG_DOWNLOAD = "���������� ���� ������ � �������"
Const MSG_CANT_DOWNLOAD = "���������� ������� ���� ������ � �������"
Const MSG_EXTRACT = "���������� ���� ������ �� ���������� �����"
Const MSG_CANT_EXTRACT = "���������� ������ ���� ������ �� ���������� �����"
Const MSG_DETETE_MSI_DB = "�������� ���� ��������� ����� ���� ������"
Const MSG_CANT_DETETE_MSI_DB = "���������� ������� ���� ��������� ����� ���� ������"
Const MSG_DETETE_OLD_DB = "�������� ������ ���� ������"
Const MSG_CANT_DETETE_OLD_DB = "���������� ������� ������ ���� ������"
Const MSG_RENAME = "���������������� ����� ������ ���� ������"
Const MSG_CANT_RENAME = "���������� ������������� ����� ������ ���� ������"
Const MSG_WRITE_VER = "���������� ���������� �� ����������� ���� ������"
Const MSG_CANT_WRITE_VER = "���������� ��������� ���������� �� ����������� ���� ������"
Const MSG_CANT_UPDATE_DB = "�� ������� �������� ���� ������ (����������� � ����� ����). ����� ������� 2GIS �� ������ ����� ������."

Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(file_prog_log) Then deleteFile(file_prog_log) ' ���� ���� ���� ������� ���������� - �������

If Not update() Then: msg_time(MSG_CANT_UPDATE_DB)

If Not FSO.FileExists(exe_2gis) Then msg_err(MSG_CANT_2GIS): WScript.Quit ' ���� ����� ������� 2GIS ���, �������

Dim WshShell: Set WshShell = WScript.CreateObject("WScript.Shell")
Dim res: res = WshShell.Run(exe_2gis, 1, False)

'------------------------------------------------------------------------------------------------------------------------------------------------

Function update()
	update = False
	' ���������� � ����������
	msg_log(MSG_CONNECT)
	Dim html_text: html_text = getHtmlCodeWebPageFromUrl(url_update)
	If html_text ="" Then msg_log(MSG_CANT_CONNECT): Exit Function
	
	' ��������� ������� ������ �� ��
	msg_log(MSG_FIND_DB)
	Dim down_link: down_link = getTextBetween2Patt(html_text, pattern_begin_url, pattern_end_url) ' ���� ������ �� ��������� MSI �����
	If down_link ="" Then msg_log(MSG_CANT_FIND_DB): Exit Function
 	down_link = pattern_begin_url & down_link & pattern_end_url ' ������ �� ���������� ����� MSI
 	
	' ������ ������ ��
	msg_log(MSG_FIND_VER_DB)
	Dim ver_gdat: ver_gdat = getTextBetween2Patt(html_text, pattern_begin_ver, pattern_end_ver) ' ���� ������ ����
	If ver_gdat ="" Then msg_log(MSG_CANT_FIND_VER_DB): Exit Function
	
	' ������ ��������� ������ ��
	msg_log(MSG_READE_OLD_VER_DB)
	Dim ver_gdat_local: ver_gdat_local = readTxtFromFile(file_ver_gdat)
	If ver_gdat_local = -1 Then ver_gdat_local = "<�����������>": msg_log(MSG_CANT_READE_OLD_VER_DB)
	'ver_gdat_local = Replace(ver_gdat_local, " ", "") ' ������� ��� �������
	If ver_gdat_local = "" Then ver_gdat_local = "<�����������>": msg_log(MSG_CANT_READE_OLD_VER_DB)
	
	' ������� ���������
	Dim key_press
	If StrComp(ver_gdat, ver_gdat_local, vbTextCompare) <> 0 Then ' ���� ������ �� ���������, ����������
		key_press = MsgBox(MSG_FOND_NEW_VER_2GIS & ver_gdat & vbCrLf & MSG_LOCAL_VER_2GIS & ver_gdat_local & vbCrLf & MSG_UPDATE_NOW, _
				vbYesNo + vbInformation + vbDefaultButton2, name_prog) ' �� ��������� ������� '���'
	Else ' ���� ������ ���������, �������
		update = True
	 	Exit Function
	End If
	If key_press = vbNo Then update = True: Exit Function ' ���� ���� ������ '���' ��� ����������
	
	Dim name_msi: name_msi = name_DB_in_msi + ".msi" ' ��� ��� ������� ����� �������� ���� MSI
	' ��������� ���� MSI �� ������
	msg_log(MSG_DOWNLOAD & ": " & down_link & "," & name_msi)
 	If Not downloadFileByUrl(down_link, name_msi) Then msg_log(MSG_CANT_DOWNLOAD): Exit Function
 	
	' ��������� �� MSI ���� ��
	msg_log(MSG_EXTRACT)
	If Not extractFileFromMsi(name_msi, "", name_DB_in_msi) Then msg_log(MSG_CANT_EXTRACT): Exit Function
	
	' ������� MSI ���� ��
	msg_log(MSG_DETETE_MSI_DB)
	If Not deleteFile(name_msi) Then msg_log(MSG_CANT_DETETE_MSI_DB): Exit Function
	
	' ������� ������ ��
	msg_log(MSG_DETETE_OLD_DB)
	If Not deleteFile(name_DB_base) Then msg_log(MSG_CANT_DETETE_OLD_DB): Exit Function
	
	' ��������������� ���� ����� ��
	msg_log(MSG_RENAME)
	If Not renameFile(name_DB_in_msi, name_DB_base) Then msg_log(MSG_CANT_RENAME): Exit Function
	
	' ��������� ���������� � ����� ������ �� � ����
	msg_log(MSG_WRITE_VER)
	If Not writeTxtToFile(file_ver_gdat, ver_gdat, vbForWriting) Then msg_log(MSG_CANT_WRITE_VER)
	update = True
End Function	

'------------------------------------------------------------------------------------------------------------------------------------------------

Sub msg_err(ByVal msg)
	MsgBox msg, vbCritical + vbOkOnly, name_prog & " - ������"
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
	' ���� �������� ������ �� ����������� � �������, �� ���������� 0
	If url = "" Then Exit Function
	' ��������� http ������
	Dim oHttp
	On Error Resume Next
		Set oHttp = CreateObject("MSXML2.XMLHTTP")
		If Err.Number <> 0 Then Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
	On Error GoTo 0
	If oHttp Is Nothing Then Exit Function ' ���� ������ ������� �� ������, �� �������
	oHttp.Open "GET", url, False
	oHttp.Send ' ���������� ������
	html = UCase(oHttp.responseText) ' ��������� ����� �� ������ (HTML ����� �����)
	Set oHttp = Nothing ' ����������� ������ �������
	getHtmlCodeWebPageFromUrl = html
End Function

' �������� ����� �� ������ ����� ���� ��������
' str - ������� ������, pat1 - ������ ��������, pat2 - ������ ������
' ������������ �������� - ����� ����������� ����� ���� ��������, ���� ���, ���������� ""
Function getTextBetween2Patt(str, pat1, pat2)
	getTextBetween2Patt = ""
	Dim i_1: i_1 = InStr(1, str, pat1, vbTextCompare) ' ���� 1-� ��������� ������ ������� 1 ��� ����� ��������
	If i_1 = 0 Then Exit Function ' ���� �� �������, �������
	i_1 = i_1 + Len(pat1) ' �������� �� 1-�� ������ ����� ������� 1 
	Dim i_2: i_2 = InStr(i_1, str, pat2, vbTextCompare) ' ���� 1-�� ��������� ������� 2 ��� ����� �������� ����� ���� ��� ������ ������ 1
	If i_2 = 0 Then Exit Function ' ���� �� �������, �������
	getTextBetween2Patt = Mid(str, i_1, i_2 - i_1) ' ����� �������
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------

Function writeTxtToFile(ByVal filename, ByVal text, ByVal iomode)
	writeTxtToFile = False
	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")	
	On Error Resume Next
	Dim txtFile: Set txtFile = FSO.OpenTextFile(filename, iomode, True) ' True - �������� ����� ���� ���
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
	Dim txtFile: Set txtFile = FSO.OpenTextFile(filename, vbForReading) ' 3-�� ��������� False �� ���������
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
	res = WshShell.Run(cmd, 1, True) ' True - ���� ��������� ���������� 
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
 	if Not FSO.FileExists(file_msi) Then Exit Function ' ���� ���������� 7zip �� �������, �������
	cmd = exe_7zip & " e " & file_msi & " " & extract_file
	' ���� ������� ���������� ����������, ��������� �� � ���������
	If Not (path_dest = "") Then cmd = cmd & " -o" & Chr(34) & path_dest & Chr(34) ' ������� Chr(34), ��� �� ��������� �������
	If FSO.FileExists(extract_file) Then cmd = cmd & " -ao"
	Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
	res = WshShell.Run(cmd, 1, True) ' True - ���� ��������� ���������� 
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
	If Err.Number <> 0 Then Exit Function ' ������ ������� � �����
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
	If Err.Number <> 0 Then Exit Function ' ������ ������� � �����
	On Error GoTo 0
	renameFile = True
End Function
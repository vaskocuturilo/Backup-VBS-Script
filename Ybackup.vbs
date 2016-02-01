Rem ===========================================================
Rem ===========================================================
Const Timeout = 3600, _	
	WorkFolder  = "D:\DATABASE", _
	L1ArhPath   = "F:\ARCHIVE", _
	L1ArhPrefix = "DBARH", _
	L1ArhDepth  = 7, _
	L2ArhPath   = "\\BACKSRV\DBARH", _
	L2ArhPrefix = "DBARH", _
	SMTPAddr = "127.0.0.1", _
	SMTPPort = 25, _
	MailFrom = "nobody@local.domain", _
	MailTo = "admin@local.domain"
Rem --------------------
Rem --------------------
Const ForReading = 1, ForWriting = 2
Const WindowsFolder = 0, SystemFolder = 1, TemporaryFolder = 2
Const MoveMode = True, CopyMode = False
Rem ---------------------------------
Rem ---------------------------------
Dim oFSO, oWSNet
Dim res, TmpName, L1ArhName, L2ArhName, FSize, log, descr
Set oFSO = CreateObject ("Scripting.FileSystemObject")
Set oWSNet = CreateObject ("WScript.Network")
res = True
log = ""
Rem -------------------------------------
Rem -------------------------------------
TmpName = TempFileName
L1ArhName = L1ArhPath & "\" & L1FileName (L1ArhPrefix, 1)
L2ArhName = L2ArhPath & "\" & L2FileName (L2ArhPrefix)
res = ZipFolder (WorkFolder, TmpName)
Rem Проверка результата операции.
descr = " упаковка """ & WorkFolder & """ -> """ & TmpName & """ (размер: " & FormatNumber (FileSize (TmpName), 0) & " байт)."
If res Then
	log = "1 - успешно:" & descr
Else
	log = "1 - ошибка:" & descr
End If
Rem -----------------------------------------
Rem -----------------------------------------
If res Then
	RotateFiles L1ArhPath, L1ArhPrefix, L1ArhDepth
	Rem Проверка результата операции.
	descr = " ротация оперативного архива."
	If res Then
		log = log & vbCrLf & "2 - успешно:" & descr
	Else
		log = log & vbCrLf & "2 - ошибка:" & descr
	End If
End If
Rem -----------------------------------
Rem 3. Перемещение в архив 1-го уровня.
Rem -----------------------------------
If res Then
	res = DragFile (TmpName, L1ArhName, MoveMode)
	Rem Проверка результата операции.
	descr = " помещение в оперативный архив """ & TmpName & """ => """ & L1ArhName & """."
	If res Then
		log = log & vbCrLf & "3 - успешно:" & descr
	Else
		log = log & vbCrLf & "3 - ошибка: " & descr
	End If
End If
Rem -----------------------------------
Rem -----------------------------------
If res Then
	res = DragFile (L1ArhName, L2ArhName, CopyMode)
	Rem Проверка результата операции.
	descr = " помещение в долговременный архив """ & L1ArhName & """ -> """ & L2ArhName & """."
	If res Then
		log = log & vbCrLf & "4 - успешно:" & descr
	Else
		log = log & vbCrLf & "4 - ошибка:" & descr
	End If
End If
Rem ----------------------------------------------------------
Rem ----------------------------------------------------------
descr = oWSNet.UserName & "@" & oWSNet.ComputerName & " сообщает: "
Err.Clear
On Error Resume Next
If res Then
	SendEmail descr & "успешное резервное копирование.", log
Else
	SendEmail descr & "ошибка резервного копирования.", log
End If
On Error GoTo 0
Rem -----------------------------------------------------------
Rem -----------------------------------------------------------
If res And Err.Number = 0 Then
	WScript.Quit (0)	' успешное завершение
Else
	WScript.Quit (1)	' завершение с ошибкой
End If
Rem ============================================
Rem Ротация файлов в папке.
Rem Вызов: Path - путь к папке,
Rem        Prefix - префикс имени файла,
Rem        Count - количество ротируемых файлов.
Rem Возврат: True - успешное завершение,
Rem          False - ошибка ротации.
Rem ============================================
Function RotateFiles (Path, Prefix, Count)
	Dim res, n, FName1, i, FName2
	res = True
	Rem --------------------------------------------------
	Rem Поиск свободного слота в списке ротируемых файлов.
	Rem --------------------------------------------------
	n = 0
	Do
		n = n + 1
		FName1 = Path & "\" & L1FileName (Prefix, n)
	Loop While oFSO.FileExists (FName1) And n < Count
	Rem -------------------
	Rem Выполнение ротации.
	Rem -------------------
	If n = Count Then Kill FName1
	For i = n To 2 Step -1
		FName2 = Path & "\" & L1FileName (Prefix, i - 1)
		res = res And DragFile (FName2, FName1, MoveMode)
		FName1 = FName2
	Next
	RotateFiles = res
End Function
Rem ====================================================
Rem Копирование папки со всем её содержимым в Zip-архив.
Rem Вызов: SrcFld - исходная папка,
Rem        TgtZip - архив назначения.
Rem Возврат: True - папка записана в Zip-архив,
Rem          False - ошибка копирования.
Rem ====================================================
Function ZipFolder (SrcFld, TgtZip)
	Dim oApp, oNS	' объекты
	Dim res, EmptyZip, f, t
	res = False
	EmptyZip = "PK" + Chr (5) + Chr (6) + String (18, Chr (0))
	On Error Resume Next
	Rem --------------------------------
	Rem Создание пустого ZIP-контейнера.
	Rem --------------------------------
	If Not oFSO.FileExists (TgtZip) Then
		Set f = oFSO.OpenTextFile (TgtZip, ForWriting, True)
		f.Write (EmptyZip)
		f.Close
		Set f = Nothing
	End If
	Rem --------------------------------------------
	Rem Проверка результата создания ZIP-контейнера.
	Rem --------------------------------------------
	If Err.Number <> 0 Then
		Kill TgtZip
	Else
		Rem ----------------------------------
		Rem Копирование папки в ZIP-контейнер.
		Rem ----------------------------------
		If Right (SrcFld, 1) <> "" Then	SrcFld = SrcFld + "\"
		Set oApp = CreateObject ("Shell.Application")
		Set oNS = oApp.NameSpace (TgtZip)
		n = oNS.Items.Count + 1
		oNS.CopyHere SrcFld
		Rem --------------------------------
		Rem Ожидание завершения копирования.
		Rem --------------------------------
		t = 0
		Do
			WScript.Sleep (1000)	' задержка на 1 секунду
			t = t + 1
		Loop While oNS.Items.Count < n And t < Timeout And Err.Number = 0
		Set oNS = Nothing
		Set oApp = Nothing
		Rem --------------------------------
		Rem Проверка результата копирования.
		Rem --------------------------------
		If t < TimeOut And Err.Number = 0 Then
			res = oFSO.FileExists (TgtZip)
		Else
			Kill TgtZip
		End If
	End If
	On Error Goto 0
	ZipFolder = res
End Function
Rem =================================
Rem Получение имени временного файла.
Rem Возврат: имя файла.
Rem =================================
Function TempFileName
	TempFileName = oFSO.GetSpecialFolder (TemporaryFolder) & "\" & oFSO.GetTempName & ".zip"
End Function
Rem ====================================
Rem Удаление файла.
Rem Вызов: FName - имя удаляемого файла.
Rem ====================================
Sub Kill (FName)
	On Error Resume Next
	If oFSO.FileExists (FName) Then oFSO.DeleteFile FName
	On Error Goto 0
	Err.Clear
End Sub
Rem =================================
Rem Получение размера файла в байтах.
Rem Вызов: FName - имя файла.
Rem Возврат: размер файла или
Rem          -1 в случае ошибки.
Rem =================================
Function FileSize (FName)
	Dim f, res
	If oFSO.FileExists (FName) Then
		Set f = oFSO.GetFile (FName)
		res = f.Size
		Set f = Nothing
	Else
		res = -1
	End If
	FileSize = res
End Function
Rem =========================================================
Rem Перемещение или копирование файла с контролем результата.
Rem Вызов: SrcFName - имя файла-источника,
Rem        TgtFName - имя файла-приёмника,
Rem        Mode - перемещение, если True,
Rem               и копирование, если False.
Rem Возврат: True - файл перемещён (копирован) успешно,
Rem          False - ошибка перемещения (копирования) файла.
Rem =========================================================
Function DragFile (SrcFName, TgtFName, Mode)
	Dim res, sz
	res = False
	On Error Resume Next
	sz = FileSize (SrcFName)
	If Mode Then oFSO.MoveFile SrcFName, TgtFName Else oFSO.CopyFile SrcFName, TgtFName, True
	If Err.Number = 0 Then res = sz = FileSize (TgtFName) Else Kill TgtFName
	On Error GoTo 0
	Err.Clear
	DragFile = res
End Function
Rem ========================================
Rem Отправка сообщения по электронной почте.
Rem Вызов: Subject - тема сообщения,
Rem        Message - текст сообщения.
Rem ========================================
Sub SendEmail (Subject, Message)
	Const cdoSendUsingPort = 2
	Dim email
	Set email = CreateObject ("CDO.Message")
	With email
		.From = MailFrom
		.To = MailTo
		.Subject = Subject
		.BodyPart.Charset = "koi8-r"
		.TextBody = Message
		With .Configuration.Fields
			.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
			.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPAddr
			.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort
			.Update
		End With
		.Send
	End With
	Set email = Nothing
End Sub
Rem ================================================
Rem Формирование имени файла для архива 1-го уровня.
Rem Вызов: Prefix - префикс имени файла,
Rem        n - номер версии файла.
Rem Возврат: имя файла в формате "Prefix#.zip".
Rem ================================================
Function L1FileName (Prefix, n)
	L1FileName = Prefix & n & ".zip"
End Function
Rem ================================================
Rem Формирование имени файла для архива 2-го уровня.
Rem Вызов: Prefix - префикс имени файла.
Rem Возврат: имя файла в формате "PrefixYYYYMM.zip".
Rem ================================================
Function L2FileName (Prefix)
	Dim Today
	d = Date
	L2FileName = Prefix & Year (d) & Right ("0" & Month (d), 2) & ".zip"
End Function

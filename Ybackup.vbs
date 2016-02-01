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
Rem �������� ���������� ��������.
descr = " �������� """ & WorkFolder & """ -> """ & TmpName & """ (������: " & FormatNumber (FileSize (TmpName), 0) & " ����)."
If res Then
	log = "1 - �������:" & descr
Else
	log = "1 - ������:" & descr
End If
Rem -----------------------------------------
Rem -----------------------------------------
If res Then
	RotateFiles L1ArhPath, L1ArhPrefix, L1ArhDepth
	Rem �������� ���������� ��������.
	descr = " ������� ������������ ������."
	If res Then
		log = log & vbCrLf & "2 - �������:" & descr
	Else
		log = log & vbCrLf & "2 - ������:" & descr
	End If
End If
Rem -----------------------------------
Rem 3. ����������� � ����� 1-�� ������.
Rem -----------------------------------
If res Then
	res = DragFile (TmpName, L1ArhName, MoveMode)
	Rem �������� ���������� ��������.
	descr = " ��������� � ����������� ����� """ & TmpName & """ => """ & L1ArhName & """."
	If res Then
		log = log & vbCrLf & "3 - �������:" & descr
	Else
		log = log & vbCrLf & "3 - ������: " & descr
	End If
End If
Rem -----------------------------------
Rem -----------------------------------
If res Then
	res = DragFile (L1ArhName, L2ArhName, CopyMode)
	Rem �������� ���������� ��������.
	descr = " ��������� � �������������� ����� """ & L1ArhName & """ -> """ & L2ArhName & """."
	If res Then
		log = log & vbCrLf & "4 - �������:" & descr
	Else
		log = log & vbCrLf & "4 - ������:" & descr
	End If
End If
Rem ----------------------------------------------------------
Rem ----------------------------------------------------------
descr = oWSNet.UserName & "@" & oWSNet.ComputerName & " ��������: "
Err.Clear
On Error Resume Next
If res Then
	SendEmail descr & "�������� ��������� �����������.", log
Else
	SendEmail descr & "������ ���������� �����������.", log
End If
On Error GoTo 0
Rem -----------------------------------------------------------
Rem -----------------------------------------------------------
If res And Err.Number = 0 Then
	WScript.Quit (0)	' �������� ����������
Else
	WScript.Quit (1)	' ���������� � �������
End If
Rem ============================================
Rem ������� ������ � �����.
Rem �����: Path - ���� � �����,
Rem        Prefix - ������� ����� �����,
Rem        Count - ���������� ���������� ������.
Rem �������: True - �������� ����������,
Rem          False - ������ �������.
Rem ============================================
Function RotateFiles (Path, Prefix, Count)
	Dim res, n, FName1, i, FName2
	res = True
	Rem --------------------------------------------------
	Rem ����� ���������� ����� � ������ ���������� ������.
	Rem --------------------------------------------------
	n = 0
	Do
		n = n + 1
		FName1 = Path & "\" & L1FileName (Prefix, n)
	Loop While oFSO.FileExists (FName1) And n < Count
	Rem -------------------
	Rem ���������� �������.
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
Rem ����������� ����� �� ���� � ���������� � Zip-�����.
Rem �����: SrcFld - �������� �����,
Rem        TgtZip - ����� ����������.
Rem �������: True - ����� �������� � Zip-�����,
Rem          False - ������ �����������.
Rem ====================================================
Function ZipFolder (SrcFld, TgtZip)
	Dim oApp, oNS	' �������
	Dim res, EmptyZip, f, t
	res = False
	EmptyZip = "PK" + Chr (5) + Chr (6) + String (18, Chr (0))
	On Error Resume Next
	Rem --------------------------------
	Rem �������� ������� ZIP-����������.
	Rem --------------------------------
	If Not oFSO.FileExists (TgtZip) Then
		Set f = oFSO.OpenTextFile (TgtZip, ForWriting, True)
		f.Write (EmptyZip)
		f.Close
		Set f = Nothing
	End If
	Rem --------------------------------------------
	Rem �������� ���������� �������� ZIP-����������.
	Rem --------------------------------------------
	If Err.Number <> 0 Then
		Kill TgtZip
	Else
		Rem ----------------------------------
		Rem ����������� ����� � ZIP-���������.
		Rem ----------------------------------
		If Right (SrcFld, 1) <> "" Then	SrcFld = SrcFld + "\"
		Set oApp = CreateObject ("Shell.Application")
		Set oNS = oApp.NameSpace (TgtZip)
		n = oNS.Items.Count + 1
		oNS.CopyHere SrcFld
		Rem --------------------------------
		Rem �������� ���������� �����������.
		Rem --------------------------------
		t = 0
		Do
			WScript.Sleep (1000)	' �������� �� 1 �������
			t = t + 1
		Loop While oNS.Items.Count < n And t < Timeout And Err.Number = 0
		Set oNS = Nothing
		Set oApp = Nothing
		Rem --------------------------------
		Rem �������� ���������� �����������.
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
Rem ��������� ����� ���������� �����.
Rem �������: ��� �����.
Rem =================================
Function TempFileName
	TempFileName = oFSO.GetSpecialFolder (TemporaryFolder) & "\" & oFSO.GetTempName & ".zip"
End Function
Rem ====================================
Rem �������� �����.
Rem �����: FName - ��� ���������� �����.
Rem ====================================
Sub Kill (FName)
	On Error Resume Next
	If oFSO.FileExists (FName) Then oFSO.DeleteFile FName
	On Error Goto 0
	Err.Clear
End Sub
Rem =================================
Rem ��������� ������� ����� � ������.
Rem �����: FName - ��� �����.
Rem �������: ������ ����� ���
Rem          -1 � ������ ������.
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
Rem ����������� ��� ����������� ����� � ��������� ����������.
Rem �����: SrcFName - ��� �����-���������,
Rem        TgtFName - ��� �����-��������,
Rem        Mode - �����������, ���� True,
Rem               � �����������, ���� False.
Rem �������: True - ���� ��������� (���������) �������,
Rem          False - ������ ����������� (�����������) �����.
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
Rem �������� ��������� �� ����������� �����.
Rem �����: Subject - ���� ���������,
Rem        Message - ����� ���������.
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
Rem ������������ ����� ����� ��� ������ 1-�� ������.
Rem �����: Prefix - ������� ����� �����,
Rem        n - ����� ������ �����.
Rem �������: ��� ����� � ������� "Prefix#.zip".
Rem ================================================
Function L1FileName (Prefix, n)
	L1FileName = Prefix & n & ".zip"
End Function
Rem ================================================
Rem ������������ ����� ����� ��� ������ 2-�� ������.
Rem �����: Prefix - ������� ����� �����.
Rem �������: ��� ����� � ������� "PrefixYYYYMM.zip".
Rem ================================================
Function L2FileName (Prefix)
	Dim Today
	d = Date
	L2FileName = Prefix & Year (d) & Right ("0" & Month (d), 2) & ".zip"
End Function

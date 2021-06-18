'===============================================================================
'==                                                                           ==
'==�@���m�c�[��                                                               ==
'==                                                                           ==
'==    �E����1��"0"�A��������[�Ȃ�]�́A���m�`�F�b�N�����s                     ==
'==    �E����1��"1"�ȏ�̐��l�́A���m�`�F�b�N�̌��ƂȂ�XML���X�g���쐬        ==
'==                                                                           ==
'==                                                                           ==
'==                                                                           ==
'==�@�쐬 : 2016/07/**�@��  �V�K�쐬                                          ==
'==�@�X�V : ****/**/**                                                        ==
'==                                                                           ==
'===============================================================================

	'Option Explicit

	' �ϐ��錾
	CONST CFG_FILE_NAME = "���m�Ώۃ��X�g.ini"
	CONST CHKSUM_FILE_NAME = "���m�Ώۃf�[�^.xml"

	Dim strSearchList		' ���m�Ώۃt�H���_���X�g�̔z��i�t�@�C���Ǎ���̓��ꕨ�j
	Dim arg1				' VBS�t�@�C���̈���1�Ԗ�

	Call OutputLog("-- �v���O�������s�J�n -----------------------------------------------------")

	' ���I�z��I�u�W�F�N�g���p
	Set strSearchList = CreateObject("System.Collections.ArrayList")

	'------------------------------------------------------------------------
	'                              �����擾
	'------------------------------------------------------------------------
	If WScript.Arguments.Count > 0 Then
		arg1 = WScript.Arguments(0)
		If IsNumeric(arg1) = False Then
			Call OutputLog("�yERROR�z�����̒l�������ł��iarg1=" & arg1 & "�j�B")
			WScript.Quit
		End If

		Call OutputLog("����������܂��iarg1=" & arg1 & "�j�B")
		If arg1 = 0 Then
			Call OutputLog("�u���m�`�F�b�N�v�����s���܂��B")
		Else
			Call OutputLog("�u���m�Ώۃ��X�g�쐬�v�����s���܂��B")
		End If
	Else
		arg1 = 0
		Call OutputLog("�����Ȃ��ł��B�u���m�`�F�b�N�v�����s���܂��B")
	End If



	'------------------------------------------------------------------------
	'                           2�d�N���`�F�b�N
	'------------------------------------------------------------------------
	' 2�d�N���`�F�b�N�i���O�o�́j
	If ChkDouble(arg1) = False Then WScript.Quit


	'------------------------------------------------------------------------
	'                     ���m�Ώۃt�H���_���X�g��Ǎ�
	'------------------------------------------------------------------------
	If ReadListFile(strSearchList) = False Then WScript.Quit


	'------------------------------------------------------------------------
	'                 �`�F�b�N�T���f�[�^�t�@�C���̑��݃`�F�b�N
	'------------------------------------------------------------------------
	If arg1 > 0 Then
	ElseIf ExistCheckSumFile() = False Then
		' ���m�Ώۃf�[�^�iXML�j�̑��݃`�F�b�N
		WScript.Quit
	End If


	'------------------------------------------------------------------------
	'                    ���m�Ώۃt�H���_�̑��݃`�F�b�N
	'------------------------------------------------------------------------
	If arg1 = 0 Then
		If ChkListFolder(strSearchList) = False Then WScript.Quit
	End If


	'------------------------------------------------------------------------
	'                    ���m�Ώۃt�H���_��XML���X�g�쐬
	'------------------------------------------------------------------------
	If arg1 <> 0 Then
		If MakeXMLList(strSearchList) = False Then WScript.Quit
	End If

	'------------------------------------------------------------------------
	'                        ��@�Ɓ@�I�@���@�`�I�I
	'------------------------------------------------------------------------
	Set strSearchList = Nothing

	' ����I��
	Call OutputLog("����ɏI�����܂����B")
	WScript.Quit







'======================================================
'=
'=	�� ��d�N���̋֎~
'=	    �E���X�g�`�F�b�N���������ꍇ�ɂ́A�����ɒ��~
'=	    �E���X�g�쐬�ł́A10�b��ɍēx�v���Z�X�`�F�b�N
'=
'=�@����1 : ����VBS�t�@�C���ւ̈���1�̂���
'=�@�߂�l : �p��(True)/���~(False)
'=
'======================================================
Function ChkDouble(arg1)
	' �u���m�`�F�b�N�v��������~�B�u���X�g�쐬�v��10�b�ҋ@�B
	Dim wmiLocator
	Dim wmiService
	Dim objEnumerator
	Dim strQuery		' SQL��
	Dim i 				' �J�E���^�[

	strQuery = "Select * FROM Win32_Process WHERE (Caption = 'wscript.exe' OR " & _
		"Caption = 'cscript.exe') AND CommandLine LIKE '%" & WScript.ScriptName & "%'"

	Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
	Set wmiService = wmiLocator.ConnectServer
	Set objEnumerator = wmiService.ExecQuery(strQuery)

	i = 0
	ChkDouble = False	' �߂�l������

	Do While (i < 100)
		If objEnumerator.Count = 1 then
			Call OutputLog("2�d�N���`�F�b�N ����")
			ChkDouble = True	' ����
			Exit Do
		ElseIf objEnumerator.Count > 1 and arg1 <> 0 Then
			' ����1��1�ȏ�i�����X�g�쐬�����j�́A3�b�ҋ@
			WScript.Sleep 3000
		ElseIf  objEnumerator.Count > 1 and arg1 = 0 Then
			' ����1��0�i�����m�`�F�b�N�j�́u������~�v�̖߂�l
			Call OutputLog("2�d�N���̂��߈ׁA�I�����܂�")
			ChkDouble = False	' ������~
			Exit Do
		End If
		i = i + 1
	Loop

	' ���̍s�܂ŗ�����A300�b�҂��ł��I����Ă��Ȃ�
	If i > 100 Then Call OutputLog("�yERROR�z�����v���Z�X���I�����Ȃ��ׁA�����I��")

	' �I������
	Set wmiLocator = Nothing
	Set wmiService = Nothing
	Set objEnumerator = Nothing

End Function


'======================================================
'=
'=	�� ���m�Ώۃ��X�g��Ǎ�
'=	   �E�z��ɁA�ǂݍ��񂾃t�H���_�ꗗ��}������
'=
'=�@����1 : �i�Q�Ɠn���j��̃t�@�C�����X�g�z��
'=�@�߂�l : �p��(True)/���~(False)
'=
'======================================================
Function ReadListFile(ByRef strSearchList)

	Dim objFileSys
	Dim objReadLine
	Dim strScriptPath		' ���s�p�X
	Dim buffer				' ���m�Ώۃ��X�g�t�@�C���s�ǂݍ��ݓ��e
	Dim str					' �ꎞ�g�p������

	ReadListFile = True		' �߂�l������


	' ���m�Ώۃ��X�g�t�@�C���̃p�X
	strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CFG_FILE_NAME

	' ���m�Ώۃ��X�g�t�@�C���̑��݃`�F�b�N
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	If objFileSys.FileExists(strScriptPath) = False Then
		Call OutputLog("�yERROR�z���m�Ώۃ��X�g�t�@�C��������܂���B�i" & strScriptPath & "�j")
		ReadListFile = False
		Exit Function
	End If

	' �ݒ�t�@�C�����J��
	Set objReadLine = objFileSys.OpenTextFile(strScriptPath , 1 )
	If Err.Number <> 0 Then
		' �t�@�C���A�N�Z�X�G���[�̏ꍇ�͏I��
		Call OutputLog("�yERROR�z�ݒ�t�@�C���̓ǂݍ��݃G���[")
		ReadListFile = False
		Exit Function
	End if

	'  �Ǎ��݊J�n�A����сA�z��ɑ}������B
	Do While not objReadLine.AtEndOfStream
		buffer = objReadLine.ReadLine
		If Trim(buffer) <> "" And Left(Trim(buffer),1) <> "#" Then
			strSearchList.Add buffer
		End If
	Loop

	' �I������
	Set objFileSys = Nothing
	Set objReadLine = Nothing

End Function


'======================================================
'=
'=�@�� ���O�t�@�C�������o��
'=�@�@�E���O�t�@�C�����́A�u[���sVBS��]_[yyyymmdd].log�v�ƂȂ�
'=
'=�@����1 : �������ݓ��e
'=�@�߂�l : �Ȃ�
'=
'======================================================
Sub OutputLog(strMsg)
	Dim objFSO		' FileSystemObject
	Dim objFile		' �t�@�C���������ݗp
	Dim strDate1	' ���ݓ��t
	Dim strDate2	' yyyymmdd�����镶����

	strDate1 = Now()

	' ���t��yyyymmdd�ɂ���
	strDate2 = Now()
	strDate2 = Left(strDate1, 10)
	strDate2 = Replace(strDate2, "/", "")

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If Err.Number = 0 Then
		' ���O�t�@�C�����I�[�v���i�ǋL�����݃��[�h�j
		Set objFile = objFSO.OpenTextFile(Replace(WScript.ScriptFullName,WScript.ScriptName, "") & _
			Left(WScript.ScriptName, Len(WScript.ScriptName) - 4) & "_" & strDate2 & ".log", 8, True)
		If Err.Number = 0 Then
			objFile.WriteLine(strDate1 & " " & strMsg)
			objFile.Close
		End If

	End If

	' �I������
	Set objFile = Nothing
	Set objFSO = Nothing

End Sub


'======================================================
'=
'=	�� ���m�Ώۃt�H���_���̌��m���s��
'=	    �E���m�Ώۃt�H���_�z������[�v�ő��݃`�F�b�N
'=
'=�@����1 : ���m�Ώۃt�H���_�z��
'=�@�߂�l : �p��(True)/���~(False)
'=
'======================================================
Function ChkListFolder(strSearchList)
	Dim objFileSys		' �I�u�W�F�N�g�N���X�i�t�@�C���E�V�X�e���j
	Dim objFolder		' �I�u�W�F�N�g�N���X�i�t�H���_�j
	Dim objSubFolder	' �I�u�W�F�N�g�N���X�i�T�u�t�H���_�j
	Dim objFile			' �I�u�W�F�N�g�N���X�i�t�@�C���j

	Dim f				' �z�����1�v�f
	Dim strList1		' �����t�@�C���i�T�u�t�H���_�j�z��
	Dim strListXML		' XML�t�@�C���̑Ώۃf�[�^�i�t�@�C�����A�`�F�b�N�T���A�X�V���j
	Dim strSHA1			' �Ώۃt�@�C���̃`�F�b�N�T���l
	Dim flg0			' XML�ƌ����Ƃ́u���́v��r����
	Dim flg1			' XML�ƌ����Ƃ́u�`�F�b�N�T���v��r����
	Dim flg2			' XML�ƌ����Ƃ́u�X�V���v��r����



	ChkListFolder = True		' �߂�l�̏�����
	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	' ------------------------------------------------------------------------
	' 1�v�f���@���݃`�F�b�N�@���@�A�t�H���_�̃`�F�b�N�@���@�B�t�@�C���̃`�F�b�N
	' ------------------------------------------------------------------------

	' �`�F�b�N�T��XML�Ǎ�
	strXML = ReadXMLFile

	For Each f In strSearchList
		' �@���݃`�F�b�N
		If objFileSys.FolderExists(f) = False Then
			' �G���[�I�I
			Call OutputLog("�yWARN�z���m�Ώۃt�H���_��[" & f & "]������܂���B")
			' �������������C�x���g�֓o�^������������
			ChkListFolder = False
			
		Else

			' �ϐ�f�̃t�H���_���L�[�ɂ��āAXML����Ώۃf�[�^��z��ŕԂ�
			strListXML = MakeXMLArray(strXML, f)

			' �Ώۃt�H���_���̃t�H���_�i�����T�u�t�H���_�j���擾
			Set objFolder = objFileSys.GetFolder(f)
			 
			'Folder�I�u�W�F�N�g��SubFolders�v���p�e�B����Folder�I�u�W�F�N�g���擾
			For Each objSubFolder In objFolder.SubFolders

				' objSubFolder.Name�v���p�e�B�́A����
				' objSubFolder.DateLastModified�v���p�e�B�́A�ŏI�X�V��

				' �t���O������
				flg0 = False
				flg2 = False

				' �A�t�H���_��r���{
				For i = 0 to UBound(strListXML)
					' ���̂̔�r
					If strListXML(i, 0) = objSubFolder.Name Then
						flg0 = True
						' �X�V���̔�r
						If strListXML(i, 2) = objSubFolder.DateLastModified Then flg2 = True
					End If
				Next

				If flg0 = False Then
					' �G���[�I�I
					Call OutputLog("�yError�z���m�Ώۃt�H���_�ȉ���[" & f & "\" & objSubFolder.Name & "]�t�H���_���ǉ�����Ă��܂��B")
					' �������������C�x���g�֓o�^������������
					ChkListFolder = False
				ElseIf flg2 = False Then
					' �G���[�I�I
					Call OutputLog("�yError�z���m�Ώۃt�H���_�ȉ���[" & f & "\" & objSubFolder.Name & "]�t�H���_�̍X�V�����قȂ��Ă��܂��B")
					' �������������C�x���g�֓o�^������������
					ChkListFolder = False
				End If
			Next



			'Folder�I�u�W�F�N�g��Files�v���p�e�B����File�I�u�W�F�N�g���擾
			For Each objFile In objFolder.Files
				' objFile.Name�v���p�e�B�́A����
				' objFile.DateLastModified�v���p�e�B�́A�ŏI�X�V��

				' �`�F�b�N�T�����擾����
				strSHA1 = CreateSHA1(f & "\" & objFile.Name)

				' �t���O������
				flg0 = False
				flg1 = False
				flg2 = False

				' �B�t�@�C����r���{
				For i = 0 to UBound(strListXML)
					' ���̂̔�r
					If strListXML(i, 0) = objFile.Name Then
						flg0 = True
						' �`�F�b�N�T���̔�r
						If strListXML(i, 1) = strSHA1 Then flg1 = True
						' �X�V���̔�r
						If strListXML(i, 2) = objFile.DateLastModified Then flg2 = True
					End If
				Next

				If flg0 = False Then
					' �G���[�I�I
					Call OutputLog("�yError�z���m�Ώۃt�H���_�ȉ���[" & f & "\" & objFile.Name & "]�t�@�C�����ǉ�����Ă��܂��B")
					' �������������C�x���g�֓o�^������������
					ChkListFolder = False
				ElseIf flg1 = False Then
					' �G���[�I�I
					Call OutputLog("�yError�z���m�Ώۃt�H���_�ȉ���[" & f & "\" & objFile.Name & "]�t�@�C���̃`�F�b�N�T�����قȂ��Ă��܂��B")
					' �������������C�x���g�֓o�^������������
					ChkListFolder = False
				ElseIf flg2 = False Then
					' �G���[�I�I
					Call OutputLog("�yError�z���m�Ώۃt�H���_�ȉ���[" & f & "\" & objFile.Name & "]�t�@�C���̍X�V�����قȂ��Ă��܂��B")
					' �������������C�x���g�֓o�^������������
					ChkListFolder = False
				End If

			Next


		End If

	Next

	If ChkListFolder = False Then Call OutputLog("�yError�z�ُ�����m���ďI�����܂����B")

	' �I������
	Set objFileSys = Nothing
	Set objFolder = Nothing

End Function




'======================================================
'=
'=  �� �`�F�b�N�T���E�f�[�^�t�@�C���̑��݃`�F�b�N
'=       �E���m�`�F�b�N�����̂ݎ��s���܂��B
'=
'=�@����1 : �Ȃ�
'=�@�߂�l : ����(True)/���݂Ȃ�(False)
'=
'======================================================
Function ExistCheckSumFile()

	Dim objFileSys
	Dim strXMLPath		' ���s�p�X

	ExistCheckSumFile = True		' �߂�l������

	' �`�F�b�N�T���E�t�@�C���̃p�X
	strXMLPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CHKSUM_FILE_NAME

	' �`�F�b�N�T���E�t�@�C���̑��݃`�F�b�N
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	If objFileSys.FileExists(strXMLPath) = False Then
		Call OutputLog("�yERROR�z�`�F�b�N�T���E�t�@�C��������܂���B�i" & strXMLPath & "�j")
		ExistCheckSumFile = False
	End If

	
End Function




'======================================================
'=
'=	�� �`�F�b�N�T���E�f�[�^��Ǎ�
'=	   �E���e�����̂܂ܓǍ����܂��B
'=
'=�@����1 : �Ȃ�
'=�@�߂�l : �Ǎ������t�@�C���̓��e
'=
'======================================================
Function ReadXMLFile()

	Dim objFileSys
	Dim objReadLine
	Dim strXMLPath			' XML�p�X
	Dim buffer				' ���m�Ώۃ��X�g�t�@�C���s�ǂݍ��ݓ��e
	Dim str					' �ꎞ�g�p������

	ReadXMLFile = ""		' �߂�l������

	' �`�F�b�N�T���E�f�[�^�t�@�C���̑��݃`�F�b�N
	If ExistCheckSumFile() = False Then
		Exit Function
	End If

	' �ݒ�t�@�C�����J��
	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	' ���m�Ώۃ��X�g�t�@�C���̃p�X
	strXMLPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CHKSUM_FILE_NAME

	Set objReadLine = objFileSys.OpenTextFile(strXMLPath , 1 )
	If Err.Number <> 0 Then
		' �t�@�C���A�N�Z�X�G���[�̏ꍇ�͏I��
		Call OutputLog("�yERROR�z�`�F�b�N�T���E�f�[�^�t�@�C���̓ǂݍ��݃G���[")
		Exit Function
	End if

	' �Ǎ��݊J�n�A����сA�z��ɑ}������B
	Do While not objReadLine.AtEndOfStream
		ReadXMLFile = ReadXMLFile & objReadLine.ReadLine
	Loop

	' �f�[�^���̃^�u�A���s�R�[�h���폜
	ReadXMLFile = Replace(ReadXMLFile, vbTab,"")
	ReadXMLFile = Replace(ReadXMLFile, vbCrLf,"")

	' �I������
	Set objFileSys = Nothing
	Set objReadLine = Nothing

End Function


'======================================================
'=
'=	�� �`�F�b�N�T���E�f�[�^�iXML�j����A�Ώۃt�H���_�̓��e��z��ɑg����
'=
'=  ����1 : strXML �`�F�b�N�T���f�[�^
'=  ����2 : strFName  �Ώۃt�H���_����
'=�@�߂�l : 2�����z��i1������:�f�[�^����\���A2������:0->���́A1->�`�F�b�N�T���A2->�X�V���A3->���ݗL���j
'=
'======================================================
Function MakeXMLArray(strXML, strFName)
	Dim objRegExp		' �I�u�W�F�N�g�N���X�i���K�\���j
	Dim objMatches		' �I�u�W�F�N�g�N���X�i���K�\���̌��ʁj

	Dim aryXML(0, 3)
	Dim MaxRow			' aryXML��1�����ڂ̃f�[�^����ϐ��ŕێ�
	Dim i				' �J�E���^�[

	' ---------------------------------------------------------------
	' --               �T�u�t�H���_�p�f�[�^�擾
	' ---------------------------------------------------------------
	objRegExp.Pattern = "<CheckList>.*?<ListName>" & strFName & "</ListName>.*?<folder>.*?<name>(.+?)" _
		& "</name>.*?<date>(.+?)</date>.*?<ListNameEnd>" & strFName & "</ListNameEnd></CheckList>"
	objRegExp.IgnoreCase = True						' �啶���Ə���������ʂ��Ȃ��悤�ɐݒ肵�܂��B
	objRegExp.Global = True							' ������S�̂���������悤�ɐݒ肵�܂��B

	Set objMatches = objRegExp.Execute(strXML)

	' ���ʊm�F
	If objMatches.Count > 0 Then

		For i = 0 to objMatches.Count -1

			If i > 0 Then ReDim Preserve aryXML(i, 3)	' ReDim���s
			aryXML(i, 0) = Match.Item(0)	' ����
			aryXML(i, 1) = ""				' SHA1�̃`�F�b�N�T���i�t�H���_�͖����j
			aryXML(i, 2) = Match.Item(1)	' �X�V��
			aryXML(i, 3) = 0				' �f�[�^���������݃`�F�b�N�̌��ʂ̓��ꕨ�i�u0:�Ȃ�/1:����v�Ƃ���j

		Next
	End If

	' ---------------------------------------------------------------
	' --                  �t�@�C���p�f�[�^�擾
	' ---------------------------------------------------------------
	objRegExp.Pattern = "<CheckList>.*?<ListName>" & strFName & "</ListName>.*?<file>.*?<name>(.+?)" _
		& "</name>.*?<sha1>(.+?)</sha1>.*?<date>(.+?)</date>.*?<ListNameEnd>" & strFName & "</ListNameEnd></CheckList>"
	objRegExp.IgnoreCase = True						' �啶���Ə���������ʂ��Ȃ��悤�ɐݒ肵�܂��B
	objRegExp.Global = True							' ������S�̂���������悤�ɐݒ肵�܂��B
	Set objMatches = objRegExp.Execute(strXML)		' XML�f�[�^��

	' ���ʊm�F
	If objMatches.Count > 0 Then

		MaxRow = UBound(aryXML)

		For i = 0 to objMatches.Count -1

			If i > 0 And MaxRow > 0 Then ReDim Preserve aryXML(i + MaxRow, 3)	' ReDim���s
			aryXML(i+ MaxRow, 0) = Match.Item(0)	' ����
			aryXML(i+ MaxRow, 1) = Match.Item(1)	' SHA1�̃`�F�b�N�T��
			aryXML(i+ MaxRow, 2) = Match.Item(2)	' �X�V��
			aryXML(i+ MaxRow, 3) = 0				' XML�f�[�^���������݃`�F�b�N���ʂ̓��ꕨ�i�u0:�Ȃ�/1:����v�Ƃ���j

		Next
	End If

	' �߂�l
	MakeXMLArray = aryXML

End Function


'======================================================
'=
'=  �� �`�F�b�N�T���E�f�[�^�̎擾
'=       �E�����̃t�@�C����SHA1�̌��ʂ�߂��܂��B
'=
'=�@����1 : �Ώۃt�@�C���̐�΃p�X
'=�@�߂�l : �`�F�b�N�T���̕�����
'=
'======================================================
Function CreateSHA1(strFilePath)

	' �d�g�݂͕����������B
	' http://d.hatena.ne.jp/papaking_ken/20110224/1298564016
	' ���Q�l�ɂ��܂����B

	Dim SHA1		' �I�u�W�F�N�g�N���X
	Dim MSXML		' �I�u�W�F�N�g�N���X
	Dim EL			' �I�u�W�F�N�g�N���X
	Dim binaryData	' �o�C�i���`���̃f�[�^

	' �o�C�i���`���œǍ�
WScript.echo "@1 " & strFilePath
	binaryData = ReadBinaryFile(strFilePath)

	' 0 byte�̃t�@�C���́A�`�F�b�N�T�����쐬�s��
	IF IsNull(binaryData) = False And IsEmpty(binaryData) = False Then

		Set SHA1 = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
		SHA1.ComputeHash_2(binaryData)

		Set MSXML = CreateObject("MSXML2.DOMDocument")
		Set EL = MSXML.CreateElement("tmp")
		EL.DataType = "bin.hex"
		EL.NodeTypedValue = SHA1.Hash

		' �߂�l
		CreateSHA1 = EL.Text

		' �I������
		Set SHA1 = Nothing
		Set MSXML = Nothing
		Set EL = Nothing
	Else
		CreateSHA1 = "[null]"
	End If

End Function

'======================================================
'=
'=  �� Binary�`���Ńt�@�C����ǂݍ���
'=
'=
'=�@����1 : �t�@�C���̃t���p�X
'=�@�߂�l : �o�C�i���`���œǂݍ��񂾃t�@�C�����e
'=
'=  ���l : �`�F�b�N�T���̎擾�Ŏg�p���Ă��܂�
'=
'======================================================
Function ReadBinaryFile(FileName)
WScript.echo "@2 " & FileName
	Const adTypeBinary = 1
	Dim objStream

	Set objStream = CreateObject("ADODB.Stream")
	objStream.Type = 1
	objStream.Open
	objStream.LoadFromFile(FileName)
	ReadBinaryFile = objStream.Read(-1)
	objStream.Close

	' �I������
	Set objStream = Nothing

End Function

'======================================================
'=
'=  �� ���m�t�H���_�ꗗ����AXML���X�g���쐬����
'=       �E
'=
'=�@����1 : ���m�t�H���_���X�g�E�E�E�Ƃ����Ȃ���z��
'=�@�߂�l : ����I���iTrue�j/ �ُ�iFalse�j
'=
'======================================================
Function MakeXMLList(strSearchList)

	Dim objFileSys			' �I�u�W�F�N�g�N���X
	Dim objFolder			' �I�u�W�F�N�g�N���X
	Dim objSubFolder		' �I�u�W�F�N�g�N���X
	Dim objFile				' �I�u�W�F�N�g�N���X

	Dim strXMLPath			' XML�t�@�C���̐�΃p�X
	Dim f					' �z���1�v�f
	Dim strSHA1				' �`�F�b�N�T���̒l

	MakeXMLList = False		' �߂�l�̏�����

	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	' XML�w�b�_�[������
	Call OutputXML(0, "")

	For Each f In strSearchList

		' ���݃`�F�b�N
		If objFileSys.FolderExists(f) = False Then
			' �x���I�I
			Call OutputLog("�yWARN�z���m�Ώۃt�H���_��[" & f & "]������܂���B")
			
		Else

			' XML��<CheckList>�^�O�A<ListName>�^�O
			Call OutputXML(1, f)

			' �Ώۃt�H���_���̃t�H���_�i�����T�u�t�H���_�j���擾
			Set objFolder = objFileSys.GetFolder(f)


			'Folder�I�u�W�F�N�g��SubFolders�v���p�e�B����Folder�I�u�W�F�N�g���擾
			For Each objSubFolder In objFolder.SubFolders

				' objSubFolder.Name�v���p�e�B�́A����
				' objSubFolder.DateLastModified�v���p�e�B�́A�ŏI�X�V��

				' XML��<folder>�^�O�A<name>�^�O
				Call OutputXML(2, objSubFolder.Name)

				' XML��<date>�^�O�A</folder>�^�O
				Call OutputXML(3, objSubFolder.DateLastModified)
			Next



			'Folder�I�u�W�F�N�g��Files�v���p�e�B����File�I�u�W�F�N�g���擾
			For Each objFile In objFolder.Files
				' objFile.Name�v���p�e�B�́A����
				' objFile.DateLastModified�v���p�e�B�́A�ŏI�X�V��

				' �`�F�b�N�T�����擾����
				strSHA1 = CreateSHA1(f & "\" & objFile.Name)

				' XML��<file>�^�O�A<name>�^�O
				Call OutputXML(4, objFile.Name)

				' XML��<SHA1>�^�O
				Call OutputXML(5, strSHA1)

				' XML��<date>�^�O�A</file>�^�O
				Call OutputXML(6, objFile.DateLastModified)




			Next

		End If

		' XML��<ListNameEnd>�^�O�A</CheckList>�^�O
		Call OutputXML(8, f)

	Next

	' XML�t�b�^�[��</root>�^�O������
	Call OutputXML(9, "")

	' �I������
	Set objFileSys = Nothing
	Set objFolder = Nothing

	MakeXMLList = True	' �߂�l

End Function



'======================================================
'=
'=  �� XML���X�g�ɏ������݂�����
'=
'=�@����1 : �����݃L�[�i���Q�Ɓj
'=�@����2 : �^�O�ɖ��ߍ��ޓ��e
'=�@�߂�l : �Ȃ�
'=�@
'=�@���l�F �� �����݃L�[�ɂ��Đ���
'=�@�@�@�@�@�@0 -> <xml>�錾�^�O�A<root>�J�n�^�O
'=�@�@�@�@�@�@1 -> <CheckList>�J�n�^�O�A<ListName>�J�n�I���^�O
'=�@�@�@�@�@�@2 -> <folder>�J�n�^�O�A<name>�J�n�I���^�O
'=�@�@�@�@�@�@3 -> <date>�J�n�I���^�O�A</folder>�I���^�O
'=�@�@�@�@�@�@4 -> <file>�J�n�^�O�A<name>�J�n�I���^�O
'=�@�@�@�@�@�@5 -> <sha1>�J�n�I���^�O
'=�@�@�@�@�@�@6 -> <date>�J�n�I���^�O�A</file>�I���^�O
'=�@�@�@�@�@�@7 -> [�󂫔�]
'=�@�@�@�@�@�@8 -> <ListNameEnd>�J�n�I���^�O�A</CheckList>�I���^�O
'=�@�@�@�@�@�@9 -> <root>�I���^�O
'=
'======================================================
Sub OutputXML(intKey, str)

	Dim objFSO		' FileSystemObject
	Dim objFile		' �t�@�C���������ݗp
	Dim strXMLPath	' XML�t�@�C���̐�΃p�X

	' XML�t�@�C���̃p�X
	strXMLPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CHKSUM_FILE_NAME

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If intKey = 0 Then
		' XML�t�@�C�����I�[�v���i�㏑�����[�h�j
		Set objFile = objFSO.OpenTextFile(strXMLPath, 2, True)
	Else
		' XML�t�@�C�����I�[�v���i�Ǐ����݃��[�h�j
		Set objFile = objFSO.OpenTextFile(strXMLPath, 8, True)
	End If

	If Err.Number <> 0 Then
		Call OutputLog("XML�t�@�C���̃I�[�v���G���[�BErr.Number=" & Err.Number )
		Exit Sub
	End If


	Select Case intKey
		Case 0
			' <xml>�錾�^�O�A<root>�J�n�^�O
			objFile.WriteLine("<?xml version=""1.0"" encoding=""Shift-JIS"" standalone=""yes""?>")
			objFile.WriteLine("<root>")
		Case 1
			' <CheckList>�J�n�^�O�A<ListName>�J�n�I���^�O
			objFile.WriteLine(vbTab & "<CheckList>")
			objFile.WriteLine(vbTab & vbTab & "<ListName>" & str & "</ListName>")
		Case 2
			' <folder>�J�n�^�O�A<name>�J�n�I���^�O
			objFile.WriteLine(vbTab & vbTab & "<folder>")
			objFile.WriteLine(vbTab & vbTab & vbTab & "<name>" & str & "</name>")
		Case 3
			' <date>�J�n�I���^�O�A</folder>�I���^�O
			objFile.WriteLine(vbTab & vbTab & vbTab & "<date>" & str & "</date>")
			objFile.WriteLine(vbTab & vbTab & "</folder>")
		Case 4
			' <file>�J�n�^�O�A<name>�J�n�I���^�O
			objFile.WriteLine(vbTab & vbTab & "<file>")
			objFile.WriteLine(vbTab & vbTab & vbTab & "<name>" & str & "</name>")
		Case 5
			' <sha1>�J�n�I���^�O
			objFile.WriteLine(vbTab & vbTab & vbTab & "<sha1>" & str & "</sha1>")
		Case 6
			' <date>�J�n�I���^�O�A</file>�I���^�O
			objFile.WriteLine(vbTab & vbTab & vbTab & "<date>" & str & "</date>")
			objFile.WriteLine(vbTab & vbTab & "</file>")
		Case 7
		Case 8
			' <ListNameEnd>�J�n�I���^�O�A</CheckList>�I���^�O
			objFile.WriteLine(vbTab & vbTab & "<ListNameEnd>" & str & "</ListNameEnd>")
			objFile.WriteLine(vbTab & "</CheckList>")
		Case 9
			' <root>�I���^�O
			objFile.WriteLine("</root>")
	End Select

	objFile.Close


	' �I������
	Set objFile = Nothing
	Set objFSO = Nothing


End Sub
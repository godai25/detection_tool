'===============================================================================
'==                                                                           ==
'==　検知ツール                                                               ==
'==                                                                           ==
'==    ・引数1に"0"、もしくは[なし]は、検知チェックを実行                     ==
'==    ・引数1に"1"以上の数値は、検知チェックの元となるXMLリストを作成        ==
'==                                                                           ==
'==                                                                           ==
'==                                                                           ==
'==　作成 : 2016/07/**　林  新規作成                                          ==
'==　更新 : ****/**/**                                                        ==
'==                                                                           ==
'===============================================================================

	'Option Explicit

	' 変数宣言
	CONST CFG_FILE_NAME = "検知対象リスト.ini"
	CONST CHKSUM_FILE_NAME = "検知対象データ.xml"

	Dim strSearchList		' 検知対象フォルダリストの配列（ファイル読込後の入れ物）
	Dim arg1				' VBSファイルの引数1番目

	Call OutputLog("-- プログラム実行開始 -----------------------------------------------------")

	' 動的配列オブジェクト利用
	Set strSearchList = CreateObject("System.Collections.ArrayList")

	'------------------------------------------------------------------------
	'                              引数取得
	'------------------------------------------------------------------------
	If WScript.Arguments.Count > 0 Then
		arg1 = WScript.Arguments(0)
		If IsNumeric(arg1) = False Then
			Call OutputLog("【ERROR】引数の値が無効です（arg1=" & arg1 & "）。")
			WScript.Quit
		End If

		Call OutputLog("引数があります（arg1=" & arg1 & "）。")
		If arg1 = 0 Then
			Call OutputLog("「検知チェック」を実行します。")
		Else
			Call OutputLog("「検知対象リスト作成」を実行します。")
		End If
	Else
		arg1 = 0
		Call OutputLog("引数なしです。「検知チェック」を実行します。")
	End If



	'------------------------------------------------------------------------
	'                           2重起動チェック
	'------------------------------------------------------------------------
	' 2重起動チェック（ログ出力）
	If ChkDouble(arg1) = False Then WScript.Quit


	'------------------------------------------------------------------------
	'                     検知対象フォルダリストを読込
	'------------------------------------------------------------------------
	If ReadListFile(strSearchList) = False Then WScript.Quit


	'------------------------------------------------------------------------
	'                 チェックサムデータファイルの存在チェック
	'------------------------------------------------------------------------
	If arg1 > 0 Then
	ElseIf ExistCheckSumFile() = False Then
		' 検知対象データ（XML）の存在チェック
		WScript.Quit
	End If


	'------------------------------------------------------------------------
	'                    検知対象フォルダの存在チェック
	'------------------------------------------------------------------------
	If arg1 = 0 Then
		If ChkListFolder(strSearchList) = False Then WScript.Quit
	End If


	'------------------------------------------------------------------------
	'                    検知対象フォルダのXMLリスト作成
	'------------------------------------------------------------------------
	If arg1 <> 0 Then
		If MakeXMLList(strSearchList) = False Then WScript.Quit
	End If

	'------------------------------------------------------------------------
	'                        作　業　終　了　～！！
	'------------------------------------------------------------------------
	Set strSearchList = Nothing

	' 正常終了
	Call OutputLog("正常に終了しました。")
	WScript.Quit







'======================================================
'=
'=	■ 二重起動の禁止
'=	    ・リストチェックをかけた場合には、すぐに中止
'=	    ・リスト作成では、10秒後に再度プロセスチェック
'=
'=　引数1 : このVBSファイルへの引数1のもの
'=　戻り値 : 継続(True)/中止(False)
'=
'======================================================
Function ChkDouble(arg1)
	' 「検知チェック」→処理停止。「リスト作成」→10秒待機。
	Dim wmiLocator
	Dim wmiService
	Dim objEnumerator
	Dim strQuery		' SQL文
	Dim i 				' カウンター

	strQuery = "Select * FROM Win32_Process WHERE (Caption = 'wscript.exe' OR " & _
		"Caption = 'cscript.exe') AND CommandLine LIKE '%" & WScript.ScriptName & "%'"

	Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
	Set wmiService = wmiLocator.ConnectServer
	Set objEnumerator = wmiService.ExecQuery(strQuery)

	i = 0
	ChkDouble = False	' 戻り値初期化

	Do While (i < 100)
		If objEnumerator.Count = 1 then
			Call OutputLog("2重起動チェック 正常")
			ChkDouble = True	' 正常
			Exit Do
		ElseIf objEnumerator.Count > 1 and arg1 <> 0 Then
			' 引数1が1以上（＝リスト作成処理）は、3秒待機
			WScript.Sleep 3000
		ElseIf  objEnumerator.Count > 1 and arg1 = 0 Then
			' 引数1が0（＝検知チェック）は「処理停止」の戻り値
			Call OutputLog("2重起動のため為、終了します")
			ChkDouble = False	' 処理停止
			Exit Do
		End If
		i = i + 1
	Loop

	' この行まで来たら、300秒待ちでも終わっていない
	If i > 100 Then Call OutputLog("【ERROR】既存プロセスが終了しない為、強制終了")

	' 終了処理
	Set wmiLocator = Nothing
	Set wmiService = Nothing
	Set objEnumerator = Nothing

End Function


'======================================================
'=
'=	■ 検知対象リストを読込
'=	   ・配列に、読み込んだフォルダ一覧を挿入する
'=
'=　引数1 : （参照渡し）空のファイルリスト配列
'=　戻り値 : 継続(True)/中止(False)
'=
'======================================================
Function ReadListFile(ByRef strSearchList)

	Dim objFileSys
	Dim objReadLine
	Dim strScriptPath		' 実行パス
	Dim buffer				' 検知対象リストファイル行読み込み内容
	Dim str					' 一時使用文字列

	ReadListFile = True		' 戻り値初期化


	' 検知対象リストファイルのパス
	strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CFG_FILE_NAME

	' 検知対象リストファイルの存在チェック
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	If objFileSys.FileExists(strScriptPath) = False Then
		Call OutputLog("【ERROR】検知対象リストファイルがありません。（" & strScriptPath & "）")
		ReadListFile = False
		Exit Function
	End If

	' 設定ファイルを開く
	Set objReadLine = objFileSys.OpenTextFile(strScriptPath , 1 )
	If Err.Number <> 0 Then
		' ファイルアクセスエラーの場合は終了
		Call OutputLog("【ERROR】設定ファイルの読み込みエラー")
		ReadListFile = False
		Exit Function
	End if

	'  読込み開始、および、配列に挿入する。
	Do While not objReadLine.AtEndOfStream
		buffer = objReadLine.ReadLine
		If Trim(buffer) <> "" And Left(Trim(buffer),1) <> "#" Then
			strSearchList.Add buffer
		End If
	Loop

	' 終了処理
	Set objFileSys = Nothing
	Set objReadLine = Nothing

End Function


'======================================================
'=
'=　■ ログファイル書き出し
'=　　・ログファイル名は、「[実行VBS名]_[yyyymmdd].log」となる
'=
'=　引数1 : 書き込み内容
'=　戻り値 : なし
'=
'======================================================
Sub OutputLog(strMsg)
	Dim objFSO		' FileSystemObject
	Dim objFile		' ファイル書き込み用
	Dim strDate1	' 現在日付
	Dim strDate2	' yyyymmddを入れる文字列

	strDate1 = Now()

	' 日付をyyyymmddにする
	strDate2 = Now()
	strDate2 = Left(strDate1, 10)
	strDate2 = Replace(strDate2, "/", "")

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If Err.Number = 0 Then
		' ログファイルをオープン（追記書込みモード）
		Set objFile = objFSO.OpenTextFile(Replace(WScript.ScriptFullName,WScript.ScriptName, "") & _
			Left(WScript.ScriptName, Len(WScript.ScriptName) - 4) & "_" & strDate2 & ".log", 8, True)
		If Err.Number = 0 Then
			objFile.WriteLine(strDate1 & " " & strMsg)
			objFile.Close
		End If

	End If

	' 終了処理
	Set objFile = Nothing
	Set objFSO = Nothing

End Sub


'======================================================
'=
'=	■ 検知対象フォルダ内の検知を行う
'=	    ・検知対象フォルダ配列をループで存在チェック
'=
'=　引数1 : 検知対象フォルダ配列
'=　戻り値 : 継続(True)/中止(False)
'=
'======================================================
Function ChkListFolder(strSearchList)
	Dim objFileSys		' オブジェクトクラス（ファイル・システム）
	Dim objFolder		' オブジェクトクラス（フォルダ）
	Dim objSubFolder	' オブジェクトクラス（サブフォルダ）
	Dim objFile			' オブジェクトクラス（ファイル）

	Dim f				' 配列内の1要素
	Dim strList1		' 現存ファイル（サブフォルダ）配列
	Dim strListXML		' XMLファイルの対象データ（ファイル名、チェックサム、更新日）
	Dim strSHA1			' 対象ファイルのチェックサム値
	Dim flg0			' XMLと現存との「名称」比較結果
	Dim flg1			' XMLと現存との「チェックサム」比較結果
	Dim flg2			' XMLと現存との「更新日」比較結果



	ChkListFolder = True		' 戻り値の初期化
	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	' ------------------------------------------------------------------------
	' 1要素ずつ①存在チェック　→　②フォルダのチェック　→　③ファイルのチェック
	' ------------------------------------------------------------------------

	' チェックサムXML読込
	strXML = ReadXMLFile

	For Each f In strSearchList
		' ①存在チェック
		If objFileSys.FolderExists(f) = False Then
			' エラー！！
			Call OutputLog("【WARN】検知対象フォルダの[" & f & "]がありません。")
			' ●●●●●●イベントへ登録●●●●●●
			ChkListFolder = False
			
		Else

			' 変数fのフォルダをキーにして、XMLから対象データを配列で返す
			strListXML = MakeXMLArray(strXML, f)

			' 対象フォルダ内のフォルダ（所謂サブフォルダ）を取得
			Set objFolder = objFileSys.GetFolder(f)
			 
			'FolderオブジェクトのSubFoldersプロパティからFolderオブジェクトを取得
			For Each objSubFolder In objFolder.SubFolders

				' objSubFolder.Nameプロパティは、名称
				' objSubFolder.DateLastModifiedプロパティは、最終更新日

				' フラグ初期化
				flg0 = False
				flg2 = False

				' ②フォルダ比較実施
				For i = 0 to UBound(strListXML)
					' 名称の比較
					If strListXML(i, 0) = objSubFolder.Name Then
						flg0 = True
						' 更新日の比較
						If strListXML(i, 2) = objSubFolder.DateLastModified Then flg2 = True
					End If
				Next

				If flg0 = False Then
					' エラー！！
					Call OutputLog("【Error】検知対象フォルダ以下の[" & f & "\" & objSubFolder.Name & "]フォルダが追加されています。")
					' ●●●●●●イベントへ登録●●●●●●
					ChkListFolder = False
				ElseIf flg2 = False Then
					' エラー！！
					Call OutputLog("【Error】検知対象フォルダ以下の[" & f & "\" & objSubFolder.Name & "]フォルダの更新日が異なっています。")
					' ●●●●●●イベントへ登録●●●●●●
					ChkListFolder = False
				End If
			Next



			'FolderオブジェクトのFilesプロパティからFileオブジェクトを取得
			For Each objFile In objFolder.Files
				' objFile.Nameプロパティは、名称
				' objFile.DateLastModifiedプロパティは、最終更新日

				' チェックサムを取得する
				strSHA1 = CreateSHA1(f & "\" & objFile.Name)

				' フラグ初期化
				flg0 = False
				flg1 = False
				flg2 = False

				' ③ファイル比較実施
				For i = 0 to UBound(strListXML)
					' 名称の比較
					If strListXML(i, 0) = objFile.Name Then
						flg0 = True
						' チェックサムの比較
						If strListXML(i, 1) = strSHA1 Then flg1 = True
						' 更新日の比較
						If strListXML(i, 2) = objFile.DateLastModified Then flg2 = True
					End If
				Next

				If flg0 = False Then
					' エラー！！
					Call OutputLog("【Error】検知対象フォルダ以下の[" & f & "\" & objFile.Name & "]ファイルが追加されています。")
					' ●●●●●●イベントへ登録●●●●●●
					ChkListFolder = False
				ElseIf flg1 = False Then
					' エラー！！
					Call OutputLog("【Error】検知対象フォルダ以下の[" & f & "\" & objFile.Name & "]ファイルのチェックサムが異なっています。")
					' ●●●●●●イベントへ登録●●●●●●
					ChkListFolder = False
				ElseIf flg2 = False Then
					' エラー！！
					Call OutputLog("【Error】検知対象フォルダ以下の[" & f & "\" & objFile.Name & "]ファイルの更新日が異なっています。")
					' ●●●●●●イベントへ登録●●●●●●
					ChkListFolder = False
				End If

			Next


		End If

	Next

	If ChkListFolder = False Then Call OutputLog("【Error】異常を検知して終了しました。")

	' 終了処理
	Set objFileSys = Nothing
	Set objFolder = Nothing

End Function




'======================================================
'=
'=  ■ チェックサム・データファイルの存在チェック
'=       ・検知チェック処理のみ実行します。
'=
'=　引数1 : なし
'=　戻り値 : 正常(True)/存在なし(False)
'=
'======================================================
Function ExistCheckSumFile()

	Dim objFileSys
	Dim strXMLPath		' 実行パス

	ExistCheckSumFile = True		' 戻り値初期化

	' チェックサム・ファイルのパス
	strXMLPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CHKSUM_FILE_NAME

	' チェックサム・ファイルの存在チェック
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	If objFileSys.FileExists(strXMLPath) = False Then
		Call OutputLog("【ERROR】チェックサム・ファイルがありません。（" & strXMLPath & "）")
		ExistCheckSumFile = False
	End If

	
End Function




'======================================================
'=
'=	■ チェックサム・データを読込
'=	   ・内容をそのまま読込します。
'=
'=　引数1 : なし
'=　戻り値 : 読込したファイルの内容
'=
'======================================================
Function ReadXMLFile()

	Dim objFileSys
	Dim objReadLine
	Dim strXMLPath			' XMLパス
	Dim buffer				' 検知対象リストファイル行読み込み内容
	Dim str					' 一時使用文字列

	ReadXMLFile = ""		' 戻り値初期化

	' チェックサム・データファイルの存在チェック
	If ExistCheckSumFile() = False Then
		Exit Function
	End If

	' 設定ファイルを開く
	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	' 検知対象リストファイルのパス
	strXMLPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CHKSUM_FILE_NAME

	Set objReadLine = objFileSys.OpenTextFile(strXMLPath , 1 )
	If Err.Number <> 0 Then
		' ファイルアクセスエラーの場合は終了
		Call OutputLog("【ERROR】チェックサム・データファイルの読み込みエラー")
		Exit Function
	End if

	' 読込み開始、および、配列に挿入する。
	Do While not objReadLine.AtEndOfStream
		ReadXMLFile = ReadXMLFile & objReadLine.ReadLine
	Loop

	' データ内のタブ、改行コードを削除
	ReadXMLFile = Replace(ReadXMLFile, vbTab,"")
	ReadXMLFile = Replace(ReadXMLFile, vbCrLf,"")

	' 終了処理
	Set objFileSys = Nothing
	Set objReadLine = Nothing

End Function


'======================================================
'=
'=	■ チェックサム・データ（XML）から、対象フォルダの内容を配列に組込み
'=
'=  引数1 : strXML チェックサムデータ
'=  引数2 : strFName  対象フォルダ名称
'=　戻り値 : 2次元配列（1次元目:データ数を表す、2次元目:0->名称、1->チェックサム、2->更新日、3->存在有無）
'=
'======================================================
Function MakeXMLArray(strXML, strFName)
	Dim objRegExp		' オブジェクトクラス（正規表現）
	Dim objMatches		' オブジェクトクラス（正規表現の結果）

	Dim aryXML(0, 3)
	Dim MaxRow			' aryXMLの1次元目のデータ数を変数で保持
	Dim i				' カウンター

	' ---------------------------------------------------------------
	' --               サブフォルダ用データ取得
	' ---------------------------------------------------------------
	objRegExp.Pattern = "<CheckList>.*?<ListName>" & strFName & "</ListName>.*?<folder>.*?<name>(.+?)" _
		& "</name>.*?<date>(.+?)</date>.*?<ListNameEnd>" & strFName & "</ListNameEnd></CheckList>"
	objRegExp.IgnoreCase = True						' 大文字と小文字を区別しないように設定します。
	objRegExp.Global = True							' 文字列全体を検索するように設定します。

	Set objMatches = objRegExp.Execute(strXML)

	' 結果確認
	If objMatches.Count > 0 Then

		For i = 0 to objMatches.Count -1

			If i > 0 Then ReDim Preserve aryXML(i, 3)	' ReDim実行
			aryXML(i, 0) = Match.Item(0)	' 名称
			aryXML(i, 1) = ""				' SHA1のチェックサム（フォルダは無し）
			aryXML(i, 2) = Match.Item(1)	' 更新日
			aryXML(i, 3) = 0				' データ側見た存在チェックの結果の入れ物（「0:なし/1:あり」とする）

		Next
	End If

	' ---------------------------------------------------------------
	' --                  ファイル用データ取得
	' ---------------------------------------------------------------
	objRegExp.Pattern = "<CheckList>.*?<ListName>" & strFName & "</ListName>.*?<file>.*?<name>(.+?)" _
		& "</name>.*?<sha1>(.+?)</sha1>.*?<date>(.+?)</date>.*?<ListNameEnd>" & strFName & "</ListNameEnd></CheckList>"
	objRegExp.IgnoreCase = True						' 大文字と小文字を区別しないように設定します。
	objRegExp.Global = True							' 文字列全体を検索するように設定します。
	Set objMatches = objRegExp.Execute(strXML)		' XMLデータの

	' 結果確認
	If objMatches.Count > 0 Then

		MaxRow = UBound(aryXML)

		For i = 0 to objMatches.Count -1

			If i > 0 And MaxRow > 0 Then ReDim Preserve aryXML(i + MaxRow, 3)	' ReDim実行
			aryXML(i+ MaxRow, 0) = Match.Item(0)	' 名称
			aryXML(i+ MaxRow, 1) = Match.Item(1)	' SHA1のチェックサム
			aryXML(i+ MaxRow, 2) = Match.Item(2)	' 更新日
			aryXML(i+ MaxRow, 3) = 0				' XMLデータ側見た存在チェック結果の入れ物（「0:なし/1:あり」とする）

		Next
	End If

	' 戻り値
	MakeXMLArray = aryXML

End Function


'======================================================
'=
'=  ■ チェックサム・データの取得
'=       ・引数のファイルのSHA1の結果を戻します。
'=
'=　引数1 : 対象ファイルの絶対パス
'=　戻り値 : チェックサムの文字列
'=
'======================================================
Function CreateSHA1(strFilePath)

	' 仕組みは分からんっす。
	' http://d.hatena.ne.jp/papaking_ken/20110224/1298564016
	' を参考にしました。

	Dim SHA1		' オブジェクトクラス
	Dim MSXML		' オブジェクトクラス
	Dim EL			' オブジェクトクラス
	Dim binaryData	' バイナリ形式のデータ

	' バイナリ形式で読込
WScript.echo "@1 " & strFilePath
	binaryData = ReadBinaryFile(strFilePath)

	' 0 byteのファイルは、チェックサムが作成不可
	IF IsNull(binaryData) = False And IsEmpty(binaryData) = False Then

		Set SHA1 = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
		SHA1.ComputeHash_2(binaryData)

		Set MSXML = CreateObject("MSXML2.DOMDocument")
		Set EL = MSXML.CreateElement("tmp")
		EL.DataType = "bin.hex"
		EL.NodeTypedValue = SHA1.Hash

		' 戻り値
		CreateSHA1 = EL.Text

		' 終了処理
		Set SHA1 = Nothing
		Set MSXML = Nothing
		Set EL = Nothing
	Else
		CreateSHA1 = "[null]"
	End If

End Function

'======================================================
'=
'=  ■ Binary形式でファイルを読み込み
'=
'=
'=　引数1 : ファイルのフルパス
'=　戻り値 : バイナリ形式で読み込んだファイル内容
'=
'=  備考 : チェックサムの取得で使用しています
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

	' 終了処理
	Set objStream = Nothing

End Function

'======================================================
'=
'=  ■ 検知フォルダ一覧から、XMLリストを作成する
'=       ・
'=
'=　引数1 : 検知フォルダリスト・・・といいながら配列
'=　戻り値 : 正常終了（True）/ 異常（False）
'=
'======================================================
Function MakeXMLList(strSearchList)

	Dim objFileSys			' オブジェクトクラス
	Dim objFolder			' オブジェクトクラス
	Dim objSubFolder		' オブジェクトクラス
	Dim objFile				' オブジェクトクラス

	Dim strXMLPath			' XMLファイルの絶対パス
	Dim f					' 配列の1要素
	Dim strSHA1				' チェックサムの値

	MakeXMLList = False		' 戻り値の初期化

	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	' XMLヘッダー書込み
	Call OutputXML(0, "")

	For Each f In strSearchList

		' 存在チェック
		If objFileSys.FolderExists(f) = False Then
			' 警告！！
			Call OutputLog("【WARN】検知対象フォルダの[" & f & "]がありません。")
			
		Else

			' XMLの<CheckList>タグ、<ListName>タグ
			Call OutputXML(1, f)

			' 対象フォルダ内のフォルダ（所謂サブフォルダ）を取得
			Set objFolder = objFileSys.GetFolder(f)


			'FolderオブジェクトのSubFoldersプロパティからFolderオブジェクトを取得
			For Each objSubFolder In objFolder.SubFolders

				' objSubFolder.Nameプロパティは、名称
				' objSubFolder.DateLastModifiedプロパティは、最終更新日

				' XMLの<folder>タグ、<name>タグ
				Call OutputXML(2, objSubFolder.Name)

				' XMLの<date>タグ、</folder>タグ
				Call OutputXML(3, objSubFolder.DateLastModified)
			Next



			'FolderオブジェクトのFilesプロパティからFileオブジェクトを取得
			For Each objFile In objFolder.Files
				' objFile.Nameプロパティは、名称
				' objFile.DateLastModifiedプロパティは、最終更新日

				' チェックサムを取得する
				strSHA1 = CreateSHA1(f & "\" & objFile.Name)

				' XMLの<file>タグ、<name>タグ
				Call OutputXML(4, objFile.Name)

				' XMLの<SHA1>タグ
				Call OutputXML(5, strSHA1)

				' XMLの<date>タグ、</file>タグ
				Call OutputXML(6, objFile.DateLastModified)




			Next

		End If

		' XMLの<ListNameEnd>タグ、</CheckList>タグ
		Call OutputXML(8, f)

	Next

	' XMLフッターの</root>タグ書込み
	Call OutputXML(9, "")

	' 終了処理
	Set objFileSys = Nothing
	Set objFolder = Nothing

	MakeXMLList = True	' 戻り値

End Function



'======================================================
'=
'=  ■ XMLリストに書き込みをする
'=
'=　引数1 : 書込みキー（※参照）
'=　引数2 : タグに埋め込む内容
'=　戻り値 : なし
'=　
'=　備考： ※ 書込みキーについて説明
'=　　　　　　0 -> <xml>宣言タグ、<root>開始タグ
'=　　　　　　1 -> <CheckList>開始タグ、<ListName>開始終了タグ
'=　　　　　　2 -> <folder>開始タグ、<name>開始終了タグ
'=　　　　　　3 -> <date>開始終了タグ、</folder>終了タグ
'=　　　　　　4 -> <file>開始タグ、<name>開始終了タグ
'=　　　　　　5 -> <sha1>開始終了タグ
'=　　　　　　6 -> <date>開始終了タグ、</file>終了タグ
'=　　　　　　7 -> [空き番]
'=　　　　　　8 -> <ListNameEnd>開始終了タグ、</CheckList>終了タグ
'=　　　　　　9 -> <root>終了タグ
'=
'======================================================
Sub OutputXML(intKey, str)

	Dim objFSO		' FileSystemObject
	Dim objFile		' ファイル書き込み用
	Dim strXMLPath	' XMLファイルの絶対パス

	' XMLファイルのパス
	strXMLPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & CHKSUM_FILE_NAME

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If intKey = 0 Then
		' XMLファイルをオープン（上書きモード）
		Set objFile = objFSO.OpenTextFile(strXMLPath, 2, True)
	Else
		' XMLファイルをオープン（追書込みモード）
		Set objFile = objFSO.OpenTextFile(strXMLPath, 8, True)
	End If

	If Err.Number <> 0 Then
		Call OutputLog("XMLファイルのオープンエラー。Err.Number=" & Err.Number )
		Exit Sub
	End If


	Select Case intKey
		Case 0
			' <xml>宣言タグ、<root>開始タグ
			objFile.WriteLine("<?xml version=""1.0"" encoding=""Shift-JIS"" standalone=""yes""?>")
			objFile.WriteLine("<root>")
		Case 1
			' <CheckList>開始タグ、<ListName>開始終了タグ
			objFile.WriteLine(vbTab & "<CheckList>")
			objFile.WriteLine(vbTab & vbTab & "<ListName>" & str & "</ListName>")
		Case 2
			' <folder>開始タグ、<name>開始終了タグ
			objFile.WriteLine(vbTab & vbTab & "<folder>")
			objFile.WriteLine(vbTab & vbTab & vbTab & "<name>" & str & "</name>")
		Case 3
			' <date>開始終了タグ、</folder>終了タグ
			objFile.WriteLine(vbTab & vbTab & vbTab & "<date>" & str & "</date>")
			objFile.WriteLine(vbTab & vbTab & "</folder>")
		Case 4
			' <file>開始タグ、<name>開始終了タグ
			objFile.WriteLine(vbTab & vbTab & "<file>")
			objFile.WriteLine(vbTab & vbTab & vbTab & "<name>" & str & "</name>")
		Case 5
			' <sha1>開始終了タグ
			objFile.WriteLine(vbTab & vbTab & vbTab & "<sha1>" & str & "</sha1>")
		Case 6
			' <date>開始終了タグ、</file>終了タグ
			objFile.WriteLine(vbTab & vbTab & vbTab & "<date>" & str & "</date>")
			objFile.WriteLine(vbTab & vbTab & "</file>")
		Case 7
		Case 8
			' <ListNameEnd>開始終了タグ、</CheckList>終了タグ
			objFile.WriteLine(vbTab & vbTab & "<ListNameEnd>" & str & "</ListNameEnd>")
			objFile.WriteLine(vbTab & "</CheckList>")
		Case 9
			' <root>終了タグ
			objFile.WriteLine("</root>")
	End Select

	objFile.Close


	' 終了処理
	Set objFile = Nothing
	Set objFSO = Nothing


End Sub
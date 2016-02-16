'Dir関数使用（VBA用）
Sub ExeDir()

	Dim strFileName
	Dim Counter
	
	'引数に指定したファイルがあればファイル名を取得
	strFileName = Dir("E:\TEMP\*.exe") 
	Counter = 0
	
	Do Until strFileName = ""
		Counter = Counter + 1
		strFileName = Dir() '引数省略でまだ返していないファイル名取得
	Loop
	
	MsgBox "ファイルの数は" & Counter & "個です。"
	
End Sub

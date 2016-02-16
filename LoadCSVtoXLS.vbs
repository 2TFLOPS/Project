	Dim objFSO, objXLS
	Dim strDir, strFile, strName, strExt, strPath
	
	strPath = "E:\TEMP\"
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set strDir = objFSO.GetFolder(strPath) 'フォルダオブジェクト
 
	Set objXLS = CreateObject("Excel.Application")
	objXLS.Workbooks.Add
 
	'CSVがあるだけ処理
	For Each strFile In strDir.Files
		strName = objFSO.GetBaseName(strFile)		'ファイル名を取得
		strEXT = objFSO.GetExtensionName(strFile)	'拡張子名を取得
		If LCase(strEXT) = "csv" Then	'拡張子を大文字→小文字がcsvなら
			WScript.Echo strFile
			objXLS.Workbooks(1).Worksheets.Add.Name = strName 'ファイル名のシート名
			CSVSheet objXLS.Workbooks(1).Worksheets(strName), strFile
		End If
	Next

	'デフォルトのシートは削除して保存
	objXLS.DisplayAlerts = False
	objXLS.Workbooks(1).Worksheets("Sheet1").Delete
	objXLS.Workbooks(1).Worksheets("Sheet2").Delete
	objXLS.Workbooks(1).Worksheets("Sheet3").Delete
	objXLS.DisplayAlerts = True

	objXLS.Workbooks(1).SaveAs(strPath & "hoge.xlsx")
	objXLS.Workbooks(1).Close

	MsgBox strPath & "に保存しました。"
	
	Set objXLS = Nothing

'CSVファイルをシートに読み込む
Sub CSVSheet(desWS, strFile)

	Dim sobjXLS, wrkBook, wrkSheet, strUSE
	
	Set sobjXLS = CreateObject("Excel.Application")
	Set wrkBook = sobjXLS.Workbooks.Open(strFile, False, True, 2) '2=区切り文字(,)
	
	If Not (wrkBook Is Nothing) Then
		Set wrkSheet = wrkBook.Worksheets(1)
		strUSE = wrkSheet.UsedRange.Address			 '使用したセル範囲を取得
		desWS.Range(strUSE).Value = wrkSheet.Range(strUSE).Value 'そのセル範囲の値を取得
		wrkBook.Saved = True
		wrkBook.Close
	End If

	Set wrkBook = Nothing
	Set wrkSheet = Nothing
	Set sobjXLS = Nothing
	
End Sub

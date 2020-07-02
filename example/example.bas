Attribute VB_Name = "example"
'<License>------------------------------------------------------------
'
' Copyright (c) 2020 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

Public Sub saveAsCSV()
    
    Call saveAs("CSV")
    
End Sub

Public Sub saveAsJSON_KeyAndObj()
    
    Call saveAs("JSON_KeyAndObj")
    
End Sub

Public Sub saveAsJSON_Array()
    
    Call saveAs("JSON_ObjArray")
    
End Sub

Private Function saveAs(ByVal str_type As String)
    
    ' Excel Book 選択ダイアログ
    var_path = Application.GetOpenFilename( _
        FileFilter:="Microsoft Excelブック,*.xls*", _
        MultiSelect:=True _
    ) 'todo 自分のファイルを選択しておく
    
    If VarType(var_path) = vbBoolean Then 'Boolean 型の場合
        If Not var_path Then '1つも選択されたなかった場合
            Exit Function '終了
        End If
    End If
    
    Set obj_fs = CreateObject("Scripting.FileSystemObject")
    Dim str_builder() As String
    ReDim str_builder(0)
    bl_arrFirst = True
    
    ' 選択ファイル網羅ループ
    int_minIdx = LBound(var_path)
    int_maxIdx = UBound(var_path)
    For int_idx = int_minIdx To int_maxIdx
        
        Debug.Print var_path(int_idx)
        
        'すでに開いているかどうか確認
        Set obj_ret = func_isAlreadyOpenPath(var_path(int_idx))
        If (TypeName(obj_ret) = "Workbook") Then 'すでに開いている Workbook の場合
            Set obj_book = obj_ret
            
        Else 'すでに開いている Workbook ではない場合
            
            ' Workbook を開く
            Set obj_book = Workbooks.Open( _
                fileName:=var_path(int_idx), _
                ReadOnly:=True _
            )
            
        End If
        
        'ファイル名と同名のディレクトリを生成
        str_outDirPath = func_makeNonExsitenceDirectoryPath(Left(var_path(int_idx), InStrRev(var_path(int_idx), ".") - 1))
        MkDir str_outDirPath
        If bl_arrFirst Then
            bl_arrFirst = False
        Else
            ReDim Preserve str_builder(UBound(str_builder) + 1)
        End If
        str_builder(UBound(str_builder)) = str_outDirPath
        
        ' WorkSheet 毎 CSV 生成ループ
        For Each obj_sheet In obj_book.Worksheets
            
            Debug.Print Space(4) & obj_sheet.Name
            
            Select Case str_type
            
                Case "CSV"
                    str_outFilePath = str_outDirPath & "\" & obj_sheet.Name & ".csv"
                    ret = SaveAsAny.saveSheetAsCSV(obj_sheet, str_outFilePath)
                
                Case "JSON_KeyAndObj"
                    str_outFilePath = str_outDirPath & "\" & obj_sheet.Name & ".json"
                    ret = SaveAsAny.saveSheetAsJSON( _
                        obj_sheet, _
                        str_outFilePath, _
                        arrayStyle:=False, _
                        typeGuessing:=0 _
                    )
                    
                Case "JSON_ObjArray"
                    str_outFilePath = str_outDirPath & "\" & obj_sheet.Name & ".json"
                    ret = SaveAsAny.saveSheetAsJSON( _
                        obj_sheet, _
                        str_outFilePath, _
                        arrayStyle:=True, _
                        typeGuessing:=2 _
                    )
                
                Case Else
                    'nothing todo
                    
            End Select
        
            
            
        Next obj_sheet
        
        If (TypeName(obj_ret) = "Nothing") Then 'CSVを作るために Workbook を開いた場合
            obj_book.Close SaveChanges:=False '開いたファイルを閉じる
            
        End If
        
    Next int_idx
    
    str_toClipPaths = Join(str_builder, vbCrLf) '保存先ディレクトリリストを文字列化
    
    int_ans = MsgBox( _
        "Done!" & vbCrLf & _
        vbCrLf & _
        "Files were generated in following directory(s)." & vbCrLf & _
        "Select `Yes` to copy" & vbCrLf & _
        vbCrLf & _
        str_toClipPaths, _
        vbYesNo + vbInformation _
    )
    
    If int_ans = vbYes Then ' OK が押された場合
        SetCB (str_toClipPaths) '保存先ディレクトリリストをコピー
    End If
    
End Function

'
' 指定したフルパスがすでに開いている Workbook であれば、その WorkBook オブジェクトを返す
' 開いていない場合は、Nothing を返す
'
Private Function func_isAlreadyOpenPath(ByVal str_path As String) As Variant

    Dim obj_book As Variant
    
    Set obj_book = Nothing
    
    For Each obj_tmpBook In Workbooks
        str_tmppath = obj_tmpBook.Path & "\" & obj_tmpBook.Name
        If str_path = str_tmppath Then  ' すでに開いている場合
            Set obj_book = obj_tmpBook
            Exit For
        End If
    Next obj_tmpBook
    
    Set func_isAlreadyOpenPath = obj_book
    
End Function

'
' 指定したファイル名がすでに開いている Workbook であれば、その WorkBook オブジェクトを返す
' 開いていない場合は、Nothing を返す
'
Private Function func_isAlreadyOpenFile(ByVal str_fileName As String) As Variant

    Dim obj_book As Variant
    
    Set obj_book = Nothing
    
    For Each obj_tmpBook In Workbooks
        If str_path = obj_tmpBook.Name Then  ' すでに開いている場合
            Set obj_book = obj_tmpBook
            Exit For
        End If
    Next obj_tmpBook
    
    Set func_isAlreadyOpenFile = obj_book
    
End Function

'
' 以下の条件を満たすファイルパス名を生成して返す
' 存在しないファイルパス名であること & 開いていないファイル名であること
'
Private Function func_makeNonExsitenceFilePath(ByVal str_candidate As String) As String
    
    Dim str_nonExistenceFilePath As String
    Dim lng_suffix As Long
    
    Set obj_fs = CreateObject("Scripting.FileSystemObject")
    
    '<親ディレクトリ・ファイル名・拡張子名の抽出>-------------------
    
    int_inStrRevOfDot = InStrRev(str_candidate, ".")
    int_inStrRevOfBackSlash = InStrRev(str_candidate, "\")
    
    str_parentPath = Left(str_candidate, int_inStrRevOfBackSlash - 1)
    str_fileName = Right(str_candidate, Len(str_candidate) - int_inStrRevOfBackSlash)
    
    int_inStrRevOfDot = InStrRev(str_fileName, ".")
    
    ' `.` が存在しない場合
    ' `.` が先頭の場合 e.g. `.gitignore`
    ' `.` が末尾の場合(<-ファイルシステム的にありえないかも。)
    If _
        (int_inStrRevOfDot = 0) Or _
        (int_inStrRevOfDot = 1) Or _
        (int_inStrRevOfDot = Len(str_fileName)) _
    Then
        str_noExitFileName = str_fileName
        str_extName = ""
        
    Else ' 有効な拡張子が存在する場合
    
        str_noExitFileName = Left(str_fileName, int_inStrRevOfDot - 1)
        str_extName = Right(str_fileName, Len(str_fileName) - (int_inStrRevOfDot - 1))
    
    End If
    
    '------------------</親ディレクトリ・ファイル名・拡張子名の抽出>
    
    str_nonExistenceFilePath = str_candidate
    str_nonExistenceFileName = str_fileName
    lng_suffix = 0
    
    '存在するファイルパス名であること or
    '開いているファイル名である間は、無限ループ
    Do While _
        (obj_fs.FileExists(str_nonExistenceFilePath)) Or _
        (TypeName(func_isAlreadyOpenFile(str_nonExistenceFileName)) = "Workbook")
        
        lng_suffix = lng_suffix + 1
        str_nonExistenceFileName = str_noExitFileName & "_" & Format(lng_suffix, "0") & str_extName
        str_nonExistenceFilePath = str_parentPath & "\" & str_nonExistenceFileName
        
    Loop
    
    func_makeNonExsitenceFilePath = str_nonExistenceFilePath
    
End Function

'
' 存在しないディレクトリパス名を生成して返す
'
Private Function func_makeNonExsitenceDirectoryPath(ByVal str_candidate As String) As String
    
    Dim str_nonExistenceDirectoryPath As String
    Dim lng_suffix As Long
    
    Set obj_fs = CreateObject("Scripting.FileSystemObject")
    
    str_nonExistenceDirectoryPath = str_candidate
    lng_suffix = 0
    
    '存在するパス名である間は、無限ループ
    Do While (obj_fs.FolderExists(str_nonExistenceDirectoryPath))
        
        lng_suffix = lng_suffix + 1
        str_nonExistenceDirectoryPath = str_candidate & "_" & Format(lng_suffix, "0")
        
    Loop
    
    func_makeNonExsitenceDirectoryPath = str_nonExistenceDirectoryPath
    
End Function

'<クリップボード操作>-------------------------------------------

'クリップボードに文字列を格納
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'クリップボードから文字列を取得
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</クリップボード操作>



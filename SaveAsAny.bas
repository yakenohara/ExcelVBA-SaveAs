Attribute VB_Name = "SaveAsAny"
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

'
' ワークシートを CSV ファイルとして保存する
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   CSV として保存する Worksheet
'
' outPath : String
'   保存先ファイルパス
'
Public Function saveSheetAsCSV( _
    ByVal byThisSheet As Worksheet, _
    ByVal outPath As String _
)

    Application.ScreenUpdating = False
    
    byThisSheet.Copy 'note .Copy しないで CSV 保存すると、渡された Worksheet 自体を CSV に変換してしまう
    
    Set obj_newBook = ActiveWorkbook
    obj_newBook.Sheets(1).saveAs _
        fileName:=outPath, _
        FileFormat:=xlCSV
    
    obj_newBook.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
End Function

'
' セル範囲を CSV ファイルとして保存する
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   CSV として保存する Worksheet
'
' outPath : String
'   保存先ファイルパス
'
Public Function saveRangeAsCSV( _
    ByVal byThisRange As Range, _
    ByVal outPath As String _
)

    Application.ScreenUpdating = False
    
    Set obj_newBook = Workbooks.Add
    byThisRange.Copy
    obj_newBook.Sheets(1).Range("A1").PasteSpecial
    obj_newBook.Sheets(1).saveAs _
        fileName:=outPath, _
        FileFormat:=xlCSV
    
    obj_newBook.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
End Function


'
' ワークシートを JSON ファイルとして保存する
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   JSON として保存する Worksheet
'
' outPath : String
'   保存先ファイルパス
'
' arrayStyle : Boolean default True
'   出力形式。
'   True (as default) は 1 データを 1 Object として、それが配列として連なった形式。
'   False は 1 データの最左列の値を Key 名、
'
'    e.g. 以下ののようなテーブルは、
'
'    | a   | bl_b  | b     | dbl_c | c   | d    | e      |
'    | --- | ----- | ----- | ----- | --- | ---- | ------ |
'    | 1   | TRUE  | TRUE  | 29    | 29  | stst | 1月1日 |
'
'    ↓ arrayStyle:=True だと、このようになる ↓
'
'    ```json
'    [
'        {
'            "a":1,
'            "bl_b":true,
'            "b":true,
'            "dbl_c":29,
'            "c":29,
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
'    ↓ arrayStyle:=False だと、このようになる ↓
'    ```json
'    {
'        "1":{
'            "bl_b":true,
'            "b":true,
'            "dbl_c":29,
'            "c":29,
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    }
'    ```
'
' typeGuessing : Integer default 0
'   出力先 JSON ファイル内でのデータ型判定方法
'   0 : 自動判定する。セルに入力されたデータの型に応じて決定する。
'   e.g. arrayStyle の説明で使用したテーブル例を typeGuessing:=0 で実行すると以下のようになる
'    ```json
'    [
'        {
'            "a":1,
'            "bl_b":true,
'            "b":true,
'            "dbl_c":29,
'            "c":29,
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
'   1 : Key 名につけた prefix で明示的に指定する。prefix の種類は以下の通り。
'       bl_      : Boolean 型とする
'       dbl_     : Double 型とする
'       上記以外 : String 型とする
'   e.g. arrayStyle の説明で使用したテーブル例を typeGuessing:=1 で実行すると以下のようになる
'    ```json
'    [
'        {
'            "a":"1",
'            "bl_b":true,
'            "b":"True",
'            "dbl_c":29,
'            "c":"29",
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
'   2 : 型判定しない。すべてのセルの値は String 型として扱う。
'   e.g. arrayStyle の説明で使用したテーブル例を typeGuessing:=2 で実行すると以下のようになる
'    ```json
'    [
'        {
'            "a":"1",
'            "bl_b":"True",
'            "b":"True",
'            "dbl_c":"29",
'            "c":"29",
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
' rowNumOfTitle : Long default 1
'   Key 名定義の存在する行 No
'
' colNumOfLeft : Long default 1
'   テーブルの最も左の列 No
'
' rowNumOfDataStart : Long default 2
'   データが開始される行 No
'
' indent : Integer default 4
'   出力先 JSON ファイル内での indent 幅
'
Public Function saveSheetAsJSON( _
    ByVal byThisSheet As Worksheet, _
    ByVal outPath As String, _
    Optional ByVal arrayStyle As Boolean = True, _
    Optional ByVal typeGuessing As Integer = 0, _
    Optional ByVal rowNumOfTitle As Long = 1, _
    Optional ByVal colNumOfLeft As Long = 1, _
    Optional ByVal rowNumOfDataStart As Long = 2, _
    Optional ByVal indent As Integer = 4 _
)
    
    '最終行取得
    lng_maxRow = byThisSheet.Cells(Rows.Count, colNumOfLeft).End(xlUp).Row
    
    '最終列取得
    lng_maxCol = byThisSheet.Cells(rowNumOfTitle, Columns.Count).End(xlToLeft).Column
    
    Set obj_toSaveRange = byThisSheet.Range(byThisSheet.Cells(rowNumOfTitle, colNumOfLeft), byThisSheet.Cells(lng_maxRow, lng_maxCol))
    
    x = saveRangeAsJSON( _
        obj_toSaveRange, _
        outPath, _
        arrayStyle, _
        typeGuessing, _
        rowNumOfDataStart, _
        indent _
    )
    
End Function

'
' セル範囲を JSON ファイルとして保存する
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   JSON として保存する Worksheet
'
' outPath : String
'   保存先ファイルパス
'
' arrayStyle : Boolean default True
'   出力形式。
'   True (as default) は 1 データを 1 Object として、それが配列として連なった形式。
'   False は 1 データの最左列の値を Key 名、
'
'    e.g. 以下ののようなテーブルは、
'
'    | a   | bl_b  | b     | dbl_c | c   | d    | e      |
'    | --- | ----- | ----- | ----- | --- | ---- | ------ |
'    | 1   | TRUE  | TRUE  | 29    | 29  | stst | 1月1日 |
'
'    ↓ arrayStyle:=True だと、このようになる ↓
'
'    ```json
'    [
'        {
'            "a":1,
'            "bl_b":true,
'            "b":true,
'            "dbl_c":29,
'            "c":29,
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
'    ↓ arrayStyle:=False だと、このようになる ↓
'    ```json
'    {
'        "1":{
'            "bl_b":true,
'            "b":true,
'            "dbl_c":29,
'            "c":29,
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    }
'    ```
'
' typeGuessing : Integer default 0
'   出力先 JSON ファイル内でのデータ型判定方法
'   0 : 自動判定する。セルに入力されたデータの型に応じて決定する。
'   e.g. arrayStyle の説明で使用したテーブル例を typeGuessing:=0 で実行すると以下のようになる
'    ```json
'    [
'        {
'            "a":1,
'            "bl_b":true,
'            "b":true,
'            "dbl_c":29,
'            "c":29,
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
'   1 : Key 名につけた prefix で明示的に指定する。prefix の種類は以下の通り。
'       bl_      : Boolean 型とする
'       dbl_     : Double 型とする
'       上記以外 : String 型とする
'   e.g. arrayStyle の説明で使用したテーブル例を typeGuessing:=1 で実行すると以下のようになる
'    ```json
'    [
'        {
'            "a":"1",
'            "bl_b":true,
'            "b":"True",
'            "dbl_c":29,
'            "c":"29",
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
'   2 : 型判定しない。すべてのセルの値は String 型として扱う。
'   e.g. arrayStyle の説明で使用したテーブル例を typeGuessing:=2 で実行すると以下のようになる
'    ```json
'    [
'        {
'            "a":"1",
'            "bl_b":"True",
'            "b":"True",
'            "dbl_c":"29",
'            "c":"29",
'            "d":"stst",
'            "e":"2020/01/01"
'        }
'    ]
'    ```
'
' rowNumOfDataStart : Long default 2
'   データが開始される行 No
'
' indent : Integer default 4
'   出力先 JSON ファイル内での indent 幅
'
Public Function saveRangeAsJSON( _
    ByVal byThisRange As Range, _
    ByVal outPath As String, _
    Optional ByVal arrayStyle As Boolean = True, _
    Optional ByVal typeGuessing As Integer = 0, _
    Optional ByVal rowNumOfDataStart As Long = 2, _
    Optional ByVal indent As Integer = 4 _
)

    '変数定義
    Dim fileName, fileFolder, fileFile As String
    Dim u As Long
    Dim strarr_typeDefs() As String
    Dim strarr_builder() As String
    Dim vararr_table As Variant
    
    vararr_table = byThisRange.Value
    
    lng_lIdx_1d = LBound(vararr_table, 1)
    lng_uIdx_1d = UBound(vararr_table, 1)
    lng_lIdx_2d = LBound(vararr_table, 2)
    lng_uIdx_2d = UBound(vararr_table, 2)
    
    ' 主キーに対応した object を生成する時に必要なデータ型定義リストの生成
    ReDim strarr_typeDefs(lng_lIdx_2d To lng_uIdx_2d)
    Select Case typeGuessing
        
        Case 0 '自動判定の場合
            
            lng_startIdxOfdata = lng_lIdx_1d + (rowNumOfDataStart - 1)
            For lng_colIdx = (lng_lIdx_2d) To lng_uIdx_2d
                
                var_tmp = vararr_table(lng_startIdxOfdata, lng_colIdx)
                
                Select Case TypeName(var_tmp)
                    Case "Boolean"
                        strarr_typeDefs(lng_colIdx) = "Boolean"
                        
                    Case "Double"
                        strarr_typeDefs(lng_colIdx) = "Double"
                        
                    Case Else
                        strarr_typeDefs(lng_colIdx) = "String"
                        
                End Select
                
            Next
            
        Case 1 'Key 名の prefix による明示的型指定の場合
            
            For lng_colIdx = (lng_lIdx_2d) To lng_uIdx_2d
                
                str_tmp = vararr_table(lng_lIdx_2d, lng_colIdx)
                str_prefix = Left(str_tmp, InStr(str_tmp, "_"))
                
                Select Case str_prefix
                    Case "bl_"
                        strarr_typeDefs(lng_colIdx) = "Boolean"
                        
                    Case "dbl_"
                        strarr_typeDefs(lng_colIdx) = "Double"
                        
                    Case Else
                        strarr_typeDefs(lng_colIdx) = "String"
                        
                End Select
                
            Next
            
        Case 2 '型判定しない場合
            
            For lng_colIdx = (lng_lIdx_2d + 1) To lng_uIdx_2d
                
                strarr_typeDefs(lng_colIdx) = "String"
                
            Next
            
    End Select
    
    ' JSON 開始文字 `{`
    ReDim strarr_builder(0 To 0)
    If arrayStyle Then ' 配列形式出力の場合
        strarr_builder(UBound(strarr_builder)) = "["
    
    Else ' Key and Object 形式出力の場合
        strarr_builder(UBound(strarr_builder)) = "{"
    
    End If
    
    'リストをオブジェクトに書き込む
    lng_startIdxOfdata = lng_lIdx_1d + (rowNumOfDataStart - 1)
    For lng_rowIdx = lng_startIdxOfdata To lng_uIdx_1d
        
        If lng_startIdxOfdata < lng_rowIdx Then '2つ目以降の場合
            strarr_builder(UBound(strarr_builder)) = strarr_builder(UBound(strarr_builder)) & "," '行頭に","を挿入
            
        End If
        
        '主キー
        ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
        
        If arrayStyle Then ' 配列形式出力の場合
            strarr_builder(UBound(strarr_builder)) = String(indent * 1, " ") & "{"
        
        Else ' Key and Object 形式出力の場合
            strarr_builder(UBound(strarr_builder)) = String(indent * 1, " ") & """" & CStr(vararr_table(lng_rowIdx, lng_lIdx_1d)) & """" & ":" & "{"
        
        End If
        
        
        ' 主キーに対応した object を生成
        For lng_colIdx = (lng_lIdx_2d + IIf(arrayStyle, 0, 1)) To lng_uIdx_2d  ' Key and Object 形式出力の場合は 2 列目以降を Object にする
            
            ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
            
            ' Property の定義
            str_tmpStr1 = _
                String(indent * 2, " ") & _
                """" & vararr_table(lng_lIdx_2d, lng_colIdx) & """" & _
                ":"
            
            ' Value の定義
            Select Case strarr_typeDefs(lng_colIdx)
                Case "Boolean"
                    str_tmpStr2 = LCase(CStr(vararr_table(lng_rowIdx, lng_colIdx)))
                    
                Case "Double"
                    str_tmpStr2 = CStr(vararr_table(lng_rowIdx, lng_colIdx))
                    
                Case Else
                    str_tmpStr2 = """" & CStr(vararr_table(lng_rowIdx, lng_colIdx)) & """"
                
            End Select
            
            If lng_colIdx <> lng_uIdx_2d Then
                str_tmpStr3 = ","
            Else
                str_tmpStr3 = ""
            End If
            
            strarr_builder(UBound(strarr_builder)) = str_tmpStr1 & str_tmpStr2 & str_tmpStr3
            
        Next
        
        '行の閉じタグを挿入
        ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
        strarr_builder(UBound(strarr_builder)) = String(indent * 1, " ") & "}"
    Next

    'JSON終了タグ
    ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
    
    If arrayStyle Then ' 配列形式出力の場合
        strarr_builder(UBound(strarr_builder)) = "]"
    
    Else ' Key and Object 形式出力の場合
        strarr_builder(UBound(strarr_builder)) = "}"
    
    End If
    
    ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
    strarr_builder(UBound(strarr_builder)) = ""
    
    'UTF-8 で保存
    ret = func_saveAsUTF8(Join(strarr_builder, vbCrLf), outPath)
    
End Function

'
' BOM なし UTF-8 でテキスト保存する
'
Private Function func_saveAsUTF8(ByVal str_content As String, ByVal str_outPath As String)

    '同名のJSONファイルが既にある場合は削除する
    If Dir(str_outPath) <> "" Then
        Kill str_outPath
    End If

    'JSON作成
    'オブジェクトを用意する
    Dim txt As Object
    Set txt = CreateObject("ADODB.Stream")
    txt.Charset = "UTF-8"
    txt.Open

    '内容記載
    txt.WriteText str_content
    
    'BOM なしにする
    txt.Position = 0 'ストリームの位置を0にする
    txt.Type = 1 'adTypeBinary 'データの種類をバイナリデータに変更
    txt.Position = 3 'ストリームの位置を3にする
    Dim byteData() As Byte '一時格納用
    byteData = txt.Read 'ストリームの内容を一時格納用変数に保存
    txt.Close '一旦ストリームを閉じる（リセット）
    txt.Open 'ストリームを開く
    txt.Write byteData 'ストリームに一時格納したデータを流し込む
    
    'オブジェクトの内容をファイルに保存
    txt.SaveToFile str_outPath

    'オブジェクトを閉じる
    txt.Close

End Function



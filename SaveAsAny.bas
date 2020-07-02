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
' ���[�N�V�[�g�� CSV �t�@�C���Ƃ��ĕۑ�����
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   CSV �Ƃ��ĕۑ����� Worksheet
'
' outPath : String
'   �ۑ���t�@�C���p�X
'
Public Function saveSheetAsCSV( _
    ByVal byThisSheet As Worksheet, _
    ByVal outPath As String _
)

    Application.ScreenUpdating = False
    
    byThisSheet.Copy 'note .Copy ���Ȃ��� CSV �ۑ�����ƁA�n���ꂽ Worksheet ���̂� CSV �ɕϊ����Ă��܂�
    
    Set obj_newBook = ActiveWorkbook
    obj_newBook.Sheets(1).saveAs _
        fileName:=outPath, _
        FileFormat:=xlCSV
    
    obj_newBook.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
End Function

'
' �Z���͈͂� CSV �t�@�C���Ƃ��ĕۑ�����
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   CSV �Ƃ��ĕۑ����� Worksheet
'
' outPath : String
'   �ۑ���t�@�C���p�X
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
' ���[�N�V�[�g�� JSON �t�@�C���Ƃ��ĕۑ�����
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   JSON �Ƃ��ĕۑ����� Worksheet
'
' outPath : String
'   �ۑ���t�@�C���p�X
'
' arrayStyle : Boolean default True
'   �o�͌`���B
'   True (as default) �� 1 �f�[�^�� 1 Object �Ƃ��āA���ꂪ�z��Ƃ��ĘA�Ȃ����`���B
'   False �� 1 �f�[�^�̍ō���̒l�� Key ���A
'
'    e.g. �ȉ��̂̂悤�ȃe�[�u���́A
'
'    | a   | bl_b  | b     | dbl_c | c   | d    | e      |
'    | --- | ----- | ----- | ----- | --- | ---- | ------ |
'    | 1   | TRUE  | TRUE  | 29    | 29  | stst | 1��1�� |
'
'    �� arrayStyle:=True ���ƁA���̂悤�ɂȂ� ��
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
'    �� arrayStyle:=False ���ƁA���̂悤�ɂȂ� ��
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
'   �o�͐� JSON �t�@�C�����ł̃f�[�^�^������@
'   0 : �������肷��B�Z���ɓ��͂��ꂽ�f�[�^�̌^�ɉ����Č��肷��B
'   e.g. arrayStyle �̐����Ŏg�p�����e�[�u����� typeGuessing:=0 �Ŏ��s����ƈȉ��̂悤�ɂȂ�
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
'   1 : Key ���ɂ��� prefix �Ŗ����I�Ɏw�肷��Bprefix �̎�ނ͈ȉ��̒ʂ�B
'       bl_      : Boolean �^�Ƃ���
'       dbl_     : Double �^�Ƃ���
'       ��L�ȊO : String �^�Ƃ���
'   e.g. arrayStyle �̐����Ŏg�p�����e�[�u����� typeGuessing:=1 �Ŏ��s����ƈȉ��̂悤�ɂȂ�
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
'   2 : �^���肵�Ȃ��B���ׂẴZ���̒l�� String �^�Ƃ��Ĉ����B
'   e.g. arrayStyle �̐����Ŏg�p�����e�[�u����� typeGuessing:=2 �Ŏ��s����ƈȉ��̂悤�ɂȂ�
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
'   Key ����`�̑��݂���s No
'
' colNumOfLeft : Long default 1
'   �e�[�u���̍ł����̗� No
'
' rowNumOfDataStart : Long default 2
'   �f�[�^���J�n�����s No
'
' indent : Integer default 4
'   �o�͐� JSON �t�@�C�����ł� indent ��
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
    
    '�ŏI�s�擾
    lng_maxRow = byThisSheet.Cells(Rows.Count, colNumOfLeft).End(xlUp).Row
    
    '�ŏI��擾
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
' �Z���͈͂� JSON �t�@�C���Ƃ��ĕۑ�����
'
' Parameters
' ----------
' byThisSheet : Worksheet
'   JSON �Ƃ��ĕۑ����� Worksheet
'
' outPath : String
'   �ۑ���t�@�C���p�X
'
' arrayStyle : Boolean default True
'   �o�͌`���B
'   True (as default) �� 1 �f�[�^�� 1 Object �Ƃ��āA���ꂪ�z��Ƃ��ĘA�Ȃ����`���B
'   False �� 1 �f�[�^�̍ō���̒l�� Key ���A
'
'    e.g. �ȉ��̂̂悤�ȃe�[�u���́A
'
'    | a   | bl_b  | b     | dbl_c | c   | d    | e      |
'    | --- | ----- | ----- | ----- | --- | ---- | ------ |
'    | 1   | TRUE  | TRUE  | 29    | 29  | stst | 1��1�� |
'
'    �� arrayStyle:=True ���ƁA���̂悤�ɂȂ� ��
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
'    �� arrayStyle:=False ���ƁA���̂悤�ɂȂ� ��
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
'   �o�͐� JSON �t�@�C�����ł̃f�[�^�^������@
'   0 : �������肷��B�Z���ɓ��͂��ꂽ�f�[�^�̌^�ɉ����Č��肷��B
'   e.g. arrayStyle �̐����Ŏg�p�����e�[�u����� typeGuessing:=0 �Ŏ��s����ƈȉ��̂悤�ɂȂ�
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
'   1 : Key ���ɂ��� prefix �Ŗ����I�Ɏw�肷��Bprefix �̎�ނ͈ȉ��̒ʂ�B
'       bl_      : Boolean �^�Ƃ���
'       dbl_     : Double �^�Ƃ���
'       ��L�ȊO : String �^�Ƃ���
'   e.g. arrayStyle �̐����Ŏg�p�����e�[�u����� typeGuessing:=1 �Ŏ��s����ƈȉ��̂悤�ɂȂ�
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
'   2 : �^���肵�Ȃ��B���ׂẴZ���̒l�� String �^�Ƃ��Ĉ����B
'   e.g. arrayStyle �̐����Ŏg�p�����e�[�u����� typeGuessing:=2 �Ŏ��s����ƈȉ��̂悤�ɂȂ�
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
'   �f�[�^���J�n�����s No
'
' indent : Integer default 4
'   �o�͐� JSON �t�@�C�����ł� indent ��
'
Public Function saveRangeAsJSON( _
    ByVal byThisRange As Range, _
    ByVal outPath As String, _
    Optional ByVal arrayStyle As Boolean = True, _
    Optional ByVal typeGuessing As Integer = 0, _
    Optional ByVal rowNumOfDataStart As Long = 2, _
    Optional ByVal indent As Integer = 4 _
)

    '�ϐ���`
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
    
    ' ��L�[�ɑΉ����� object �𐶐����鎞�ɕK�v�ȃf�[�^�^��`���X�g�̐���
    ReDim strarr_typeDefs(lng_lIdx_2d To lng_uIdx_2d)
    Select Case typeGuessing
        
        Case 0 '��������̏ꍇ
            
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
            
        Case 1 'Key ���� prefix �ɂ�閾���I�^�w��̏ꍇ
            
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
            
        Case 2 '�^���肵�Ȃ��ꍇ
            
            For lng_colIdx = (lng_lIdx_2d + 1) To lng_uIdx_2d
                
                strarr_typeDefs(lng_colIdx) = "String"
                
            Next
            
    End Select
    
    ' JSON �J�n���� `{`
    ReDim strarr_builder(0 To 0)
    If arrayStyle Then ' �z��`���o�͂̏ꍇ
        strarr_builder(UBound(strarr_builder)) = "["
    
    Else ' Key and Object �`���o�͂̏ꍇ
        strarr_builder(UBound(strarr_builder)) = "{"
    
    End If
    
    '���X�g���I�u�W�F�N�g�ɏ�������
    lng_startIdxOfdata = lng_lIdx_1d + (rowNumOfDataStart - 1)
    For lng_rowIdx = lng_startIdxOfdata To lng_uIdx_1d
        
        If lng_startIdxOfdata < lng_rowIdx Then '2�ڈȍ~�̏ꍇ
            strarr_builder(UBound(strarr_builder)) = strarr_builder(UBound(strarr_builder)) & "," '�s����","��}��
            
        End If
        
        '��L�[
        ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
        
        If arrayStyle Then ' �z��`���o�͂̏ꍇ
            strarr_builder(UBound(strarr_builder)) = String(indent * 1, " ") & "{"
        
        Else ' Key and Object �`���o�͂̏ꍇ
            strarr_builder(UBound(strarr_builder)) = String(indent * 1, " ") & """" & CStr(vararr_table(lng_rowIdx, lng_lIdx_1d)) & """" & ":" & "{"
        
        End If
        
        
        ' ��L�[�ɑΉ����� object �𐶐�
        For lng_colIdx = (lng_lIdx_2d + IIf(arrayStyle, 0, 1)) To lng_uIdx_2d  ' Key and Object �`���o�͂̏ꍇ�� 2 ��ڈȍ~�� Object �ɂ���
            
            ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
            
            ' Property �̒�`
            str_tmpStr1 = _
                String(indent * 2, " ") & _
                """" & vararr_table(lng_lIdx_2d, lng_colIdx) & """" & _
                ":"
            
            ' Value �̒�`
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
        
        '�s�̕��^�O��}��
        ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
        strarr_builder(UBound(strarr_builder)) = String(indent * 1, " ") & "}"
    Next

    'JSON�I���^�O
    ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
    
    If arrayStyle Then ' �z��`���o�͂̏ꍇ
        strarr_builder(UBound(strarr_builder)) = "]"
    
    Else ' Key and Object �`���o�͂̏ꍇ
        strarr_builder(UBound(strarr_builder)) = "}"
    
    End If
    
    ReDim Preserve strarr_builder(0 To UBound(strarr_builder) + 1)
    strarr_builder(UBound(strarr_builder)) = ""
    
    'UTF-8 �ŕۑ�
    ret = func_saveAsUTF8(Join(strarr_builder, vbCrLf), outPath)
    
End Function

'
' BOM �Ȃ� UTF-8 �Ńe�L�X�g�ۑ�����
'
Private Function func_saveAsUTF8(ByVal str_content As String, ByVal str_outPath As String)

    '������JSON�t�@�C�������ɂ���ꍇ�͍폜����
    If Dir(str_outPath) <> "" Then
        Kill str_outPath
    End If

    'JSON�쐬
    '�I�u�W�F�N�g��p�ӂ���
    Dim txt As Object
    Set txt = CreateObject("ADODB.Stream")
    txt.Charset = "UTF-8"
    txt.Open

    '���e�L��
    txt.WriteText str_content
    
    'BOM �Ȃ��ɂ���
    txt.Position = 0 '�X�g���[���̈ʒu��0�ɂ���
    txt.Type = 1 'adTypeBinary '�f�[�^�̎�ނ��o�C�i���f�[�^�ɕύX
    txt.Position = 3 '�X�g���[���̈ʒu��3�ɂ���
    Dim byteData() As Byte '�ꎞ�i�[�p
    byteData = txt.Read '�X�g���[���̓��e���ꎞ�i�[�p�ϐ��ɕۑ�
    txt.Close '��U�X�g���[�������i���Z�b�g�j
    txt.Open '�X�g���[�����J��
    txt.Write byteData '�X�g���[���Ɉꎞ�i�[�����f�[�^�𗬂�����
    
    '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
    txt.SaveToFile str_outPath

    '�I�u�W�F�N�g�����
    txt.Close

End Function



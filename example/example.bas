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
    
    ' Excel Book �I���_�C�A���O
    var_path = Application.GetOpenFilename( _
        FileFilter:="Microsoft Excel�u�b�N,*.xls*", _
        MultiSelect:=True _
    ) 'todo �����̃t�@�C����I�����Ă���
    
    If VarType(var_path) = vbBoolean Then 'Boolean �^�̏ꍇ
        If Not var_path Then '1���I�����ꂽ�Ȃ������ꍇ
            Exit Function '�I��
        End If
    End If
    
    Set obj_fs = CreateObject("Scripting.FileSystemObject")
    Dim str_builder() As String
    ReDim str_builder(0)
    bl_arrFirst = True
    
    ' �I���t�@�C���ԗ����[�v
    int_minIdx = LBound(var_path)
    int_maxIdx = UBound(var_path)
    For int_idx = int_minIdx To int_maxIdx
        
        Debug.Print var_path(int_idx)
        
        '���łɊJ���Ă��邩�ǂ����m�F
        Set obj_ret = func_isAlreadyOpenPath(var_path(int_idx))
        If (TypeName(obj_ret) = "Workbook") Then '���łɊJ���Ă��� Workbook �̏ꍇ
            Set obj_book = obj_ret
            
        Else '���łɊJ���Ă��� Workbook �ł͂Ȃ��ꍇ
            
            ' Workbook ���J��
            Set obj_book = Workbooks.Open( _
                fileName:=var_path(int_idx), _
                ReadOnly:=True _
            )
            
        End If
        
        '�t�@�C�����Ɠ����̃f�B���N�g���𐶐�
        str_outDirPath = func_makeNonExsitenceDirectoryPath(Left(var_path(int_idx), InStrRev(var_path(int_idx), ".") - 1))
        MkDir str_outDirPath
        If bl_arrFirst Then
            bl_arrFirst = False
        Else
            ReDim Preserve str_builder(UBound(str_builder) + 1)
        End If
        str_builder(UBound(str_builder)) = str_outDirPath
        
        ' WorkSheet �� CSV �������[�v
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
        
        If (TypeName(obj_ret) = "Nothing") Then 'CSV����邽�߂� Workbook ���J�����ꍇ
            obj_book.Close SaveChanges:=False '�J�����t�@�C�������
            
        End If
        
    Next int_idx
    
    str_toClipPaths = Join(str_builder, vbCrLf) '�ۑ���f�B���N�g�����X�g�𕶎���
    
    int_ans = MsgBox( _
        "Done!" & vbCrLf & _
        vbCrLf & _
        "Files were generated in following directory(s)." & vbCrLf & _
        "Select `Yes` to copy" & vbCrLf & _
        vbCrLf & _
        str_toClipPaths, _
        vbYesNo + vbInformation _
    )
    
    If int_ans = vbYes Then ' OK �������ꂽ�ꍇ
        SetCB (str_toClipPaths) '�ۑ���f�B���N�g�����X�g���R�s�[
    End If
    
End Function

'
' �w�肵���t���p�X�����łɊJ���Ă��� Workbook �ł���΁A���� WorkBook �I�u�W�F�N�g��Ԃ�
' �J���Ă��Ȃ��ꍇ�́ANothing ��Ԃ�
'
Private Function func_isAlreadyOpenPath(ByVal str_path As String) As Variant

    Dim obj_book As Variant
    
    Set obj_book = Nothing
    
    For Each obj_tmpBook In Workbooks
        str_tmppath = obj_tmpBook.Path & "\" & obj_tmpBook.Name
        If str_path = str_tmppath Then  ' ���łɊJ���Ă���ꍇ
            Set obj_book = obj_tmpBook
            Exit For
        End If
    Next obj_tmpBook
    
    Set func_isAlreadyOpenPath = obj_book
    
End Function

'
' �w�肵���t�@�C���������łɊJ���Ă��� Workbook �ł���΁A���� WorkBook �I�u�W�F�N�g��Ԃ�
' �J���Ă��Ȃ��ꍇ�́ANothing ��Ԃ�
'
Private Function func_isAlreadyOpenFile(ByVal str_fileName As String) As Variant

    Dim obj_book As Variant
    
    Set obj_book = Nothing
    
    For Each obj_tmpBook In Workbooks
        If str_path = obj_tmpBook.Name Then  ' ���łɊJ���Ă���ꍇ
            Set obj_book = obj_tmpBook
            Exit For
        End If
    Next obj_tmpBook
    
    Set func_isAlreadyOpenFile = obj_book
    
End Function

'
' �ȉ��̏����𖞂����t�@�C���p�X���𐶐����ĕԂ�
' ���݂��Ȃ��t�@�C���p�X���ł��邱�� & �J���Ă��Ȃ��t�@�C�����ł��邱��
'
Private Function func_makeNonExsitenceFilePath(ByVal str_candidate As String) As String
    
    Dim str_nonExistenceFilePath As String
    Dim lng_suffix As Long
    
    Set obj_fs = CreateObject("Scripting.FileSystemObject")
    
    '<�e�f�B���N�g���E�t�@�C�����E�g���q���̒��o>-------------------
    
    int_inStrRevOfDot = InStrRev(str_candidate, ".")
    int_inStrRevOfBackSlash = InStrRev(str_candidate, "\")
    
    str_parentPath = Left(str_candidate, int_inStrRevOfBackSlash - 1)
    str_fileName = Right(str_candidate, Len(str_candidate) - int_inStrRevOfBackSlash)
    
    int_inStrRevOfDot = InStrRev(str_fileName, ".")
    
    ' `.` �����݂��Ȃ��ꍇ
    ' `.` ���擪�̏ꍇ e.g. `.gitignore`
    ' `.` �������̏ꍇ(<-�t�@�C���V�X�e���I�ɂ��肦�Ȃ������B)
    If _
        (int_inStrRevOfDot = 0) Or _
        (int_inStrRevOfDot = 1) Or _
        (int_inStrRevOfDot = Len(str_fileName)) _
    Then
        str_noExitFileName = str_fileName
        str_extName = ""
        
    Else ' �L���Ȋg���q�����݂���ꍇ
    
        str_noExitFileName = Left(str_fileName, int_inStrRevOfDot - 1)
        str_extName = Right(str_fileName, Len(str_fileName) - (int_inStrRevOfDot - 1))
    
    End If
    
    '------------------</�e�f�B���N�g���E�t�@�C�����E�g���q���̒��o>
    
    str_nonExistenceFilePath = str_candidate
    str_nonExistenceFileName = str_fileName
    lng_suffix = 0
    
    '���݂���t�@�C���p�X���ł��邱�� or
    '�J���Ă���t�@�C�����ł���Ԃ́A�������[�v
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
' ���݂��Ȃ��f�B���N�g���p�X���𐶐����ĕԂ�
'
Private Function func_makeNonExsitenceDirectoryPath(ByVal str_candidate As String) As String
    
    Dim str_nonExistenceDirectoryPath As String
    Dim lng_suffix As Long
    
    Set obj_fs = CreateObject("Scripting.FileSystemObject")
    
    str_nonExistenceDirectoryPath = str_candidate
    lng_suffix = 0
    
    '���݂���p�X���ł���Ԃ́A�������[�v
    Do While (obj_fs.FolderExists(str_nonExistenceDirectoryPath))
        
        lng_suffix = lng_suffix + 1
        str_nonExistenceDirectoryPath = str_candidate & "_" & Format(lng_suffix, "0")
        
    Loop
    
    func_makeNonExsitenceDirectoryPath = str_nonExistenceDirectoryPath
    
End Function

'<�N���b�v�{�[�h����>-------------------------------------------

'�N���b�v�{�[�h�ɕ�������i�[
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'�N���b�v�{�[�h���當������擾
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</�N���b�v�{�[�h����>



Attribute VB_Name = "CodeManager"

Option Explicit

' ========================================================
'  [메인] 현재경로\IMPORT 폴더의 .bas, .cls, .frm 파일로 모듈 업데이트
' ========================================================
Sub ImportVBA_From_ImportFolder()
    Dim fso As Object, folder As Object, file As Object
    Dim importPath As String
    Dim fileName As String, codeContent As String
    Dim modName As String, modType As String
    Dim ext As String
    
    ' 1. IMPORT 폴더 경로 설정
    importPath = ThisWorkbook.Path & "\IMPORT\"
    
    ' 2. 폴더 존재 여부 확인
    If Dir(importPath, vbDirectory) = "" Then
        MsgBox "현재 파일 위치에 'IMPORT' 폴더가 없습니다." & vbCrLf & _
               "경로: " & importPath, vbExclamation
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(importPath)
    
    ' 3. IMPORT 폴더 내의 파일 순회
    For Each file In folder.Files
        ' 확장자 추출 (소문자로 변환)
        ext = LCase(fso.GetExtensionName(file.Path))
        
        ' 지원하는 확장자: bas(표준), cls(클래스), frm(폼)
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            
            ' 파일명에서 확장자를 뺀 순수 이름 추출 (예: Module1.bas -> Module1)
            modName = fso.GetBaseName(file.Path)
            modType = ext
            
            ' 4. 파일 내용 읽기 (UTF-8 가정)
            ' ※ ReadTextFileUTF8 함수는 기존에 가지고 계신 것을 사용한다고 가정합니다.
            codeContent = ReadTextFileUTF8(file.Path)
            
            ' 5. 모듈 업데이트 또는 생성
            ' 폼(frm) 파일의 경우 바이너리(.frx)가 없으면 UI가 깨질 수 있으니 주의하세요.
            UpdateModuleCode modName, modType, codeContent
            
            Debug.Print "Updated: " & modName & " (" & modType & ")"
        End If
    Next file
    
    MsgBox "IMPORT 폴더의 파일(.bas/.cls/.frm)로 코드가 업데이트 되었습니다.", vbInformation
End Sub

' ========================================================
'  [핵심 로직] 모듈을 찾아 코드를 교체 (없으면 생성)
'  (이 부분은 기존 코드와 동일하지만, 혹시 몰라 함께 드립니다)
' ========================================================
Sub UpdateModuleCode(modName As String, modType As String, content As String)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim compType As Integer
    Dim exists As Boolean
    
    Set vbProj = ThisWorkbook.VBProject
    
    ' 1. 기존 모듈 존재 여부 확인
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(modName)
    If Err.Number = 0 Then exists = True Else exists = False
    On Error GoTo 0
    
    ' 2. 모듈이 없으면 새로 생성 (시트 모듈 제외)
    If Not exists Then
        ' 모듈 타입 결정
        Select Case modType
            Case "bas": compType = 1 ' 표준 모듈
            Case "cls": compType = 2 ' 클래스 모듈
            Case "frm": compType = 3 ' 폼
            Case Else: compType = 0  ' 기타
        End Select
        
        If compType > 0 Then
            Set vbComp = vbProj.VBComponents.Add(compType)
            vbComp.Name = modName
        Else
            Debug.Print "Skipped (Invalid Type): " & modName
            Exit Sub
        End If
    End If
    
    ' 3. 코드 교체 (기존 내용 지우고 덮어쓰기)
    If Not vbComp Is Nothing Then
        With vbComp.CodeModule
            ' 내용 전체 삭제 후 새로 쓰기
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromString content
        End With
    End If
End Sub

' ========================================================
'  [공통 함수] UTF-8 텍스트 파일 읽기
' ========================================================
Function ReadTextFileUTF8(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2              ' 텍스트 모드
        .Charset = "UTF-8"     ' UTF-8 설정
        .Open
        .LoadFromFile filePath
        ReadTextFileUTF8 = .ReadText
        .Close
    End With
    
    Set stream = Nothing
End Function

' ========================================================
'  공통 함수: 텍스트 내용을 UTF-8 파일로 저장
' ========================================================
Sub SaveTextAsUTF8(filePath As String, content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2 ' 덮어쓰기
        .Close
    End With
End Sub

' ========================================================
'  1. 원래 확장자(.bas, .cls 등)를 유지하며 UTF-8로 내보내기
' ========================================================
Sub ExportAllVBAElements()
    Dim vbComp As Object
    Dim savePath As String
    Dim folderName As String
    Dim codeContent As String
    Dim fileExt As String
    
    ' 저장할 폴더 설정
    folderName = "VBA_Export_" & Format(Now, "yyyymmdd_hhmmss")
    savePath = ThisWorkbook.Path & "\" & folderName & "\"
    
    If Dir(savePath, vbDirectory) = "" Then MkDir savePath
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' 코드가 있는 경우만 처리
        If vbComp.CodeModule.CountOfLines > 0 Then
            ' 확장자 결정
            Select Case vbComp.Type
                Case 1: fileExt = ".bas" ' 표준 모듈
                Case 2: fileExt = ".cls" ' 클래스 모듈
                Case 3: fileExt = ".frm" ' 폼 (주의: 폼의 디자인/바이너리 정보는 제외되고 코드만 저장됨)
                Case 100: fileExt = ".cls" ' 시트/워크북
                Case Else: fileExt = ".txt"
            End Select
            
            ' 코드 내용 가져오기
            codeContent = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            
            ' UTF-8로 저장 (Attribute 정보 등은 제외되고 순수 코드만 저장됩니다)
            SaveTextAsUTF8 savePath & vbComp.Name & fileExt, codeContent
        End If
    Next vbComp
    
    MsgBox "모든 코드가 UTF-8 형식으로 다음 위치에 저장되었습니다:" & vbCrLf & savePath, vbInformation
End Sub

' ========================================================
'  2. 모든 코드를 개별 텍스트 파일(.txt)로 UTF-8 내보내기
' ========================================================
Sub ExportAllToTextFiles()
Dim vbComp As Object
    Dim savePath As String
    Dim folderName As String
    Dim codeContent As String
    Dim typeExt As String   ' 원래의 타입 확장자 (.bas, .cls 등)
    
    ' 저장할 폴더 설정
    folderName = "VBA_Export_" & Format(Now, "yyyymmdd_hhmmss")
    savePath = ThisWorkbook.Path & "\" & folderName & "\"
    
    ' 폴더 생성
    If Dir(savePath, vbDirectory) = "" Then MkDir savePath
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' 코드가 있는 경우만 처리
        If vbComp.CodeModule.CountOfLines > 0 Then
            
            ' 1. 원래 확장자(타입) 결정
            Select Case vbComp.Type
                Case 1: typeExt = ".bas"   ' 표준 모듈
                Case 2: typeExt = ".cls"   ' 클래스 모듈
                Case 3: typeExt = ".frm"   ' 폼 (코드만 저장됨)
                Case 100: typeExt = ".cls" ' 시트/ThisWorkbook (이것들도 일종의 클래스입니다)
                Case Else: typeExt = ""    ' 기타
            End Select
            
            ' 코드 내용 가져오기
            codeContent = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            
            ' 2. UTF-8로 저장 (파일명 규칙: 모듈명 + 타입확장자 + .txt)
            ' 예: MyMacro.bas.txt, Sheet1.cls.txt, Class1.cls.txt
            SaveTextAsUTF8 savePath & vbComp.Name & typeExt & ".txt", codeContent
            
        End If
    Next vbComp
    
    MsgBox "모든 코드가 [.타입.txt] 형식의 UTF-8 파일로 저장되었습니다:" & vbCrLf & savePath, vbInformation
End Sub
' ========================================================
'  3. 모든 코드를 하나의 텍스트 파일로 통합하여 UTF-8 내보내기
' ========================================================
Sub ExportAllToOneTextFile()
    Dim vbComp As Object
    Dim savePath As String
    Dim stream As Object
    Dim codeText As String
    
    savePath = ThisWorkbook.Path & "\All_VBA_Code.txt"
    
    ' ADODB 스트림 생성 (하나의 파일에 계속 쓰기 위해 루프 밖에서 생성)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    
    ' 헤더 작성
    stream.WriteText "================================================" & vbCrLf
    stream.WriteText " VBA PROJECT EXPORT - " & Now & vbCrLf
    stream.WriteText " FILE: " & ThisWorkbook.Name & vbCrLf
    stream.WriteText "================================================" & vbCrLf & vbCrLf
    
    ' 모든 구성 요소 순회
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.CodeModule.CountOfLines > 0 Then
            ' 구분선 및 모듈 정보
            stream.WriteText ">>> MODULE: " & vbComp.Name & " (Type: " & vbComp.Type & ")" & vbCrLf
            stream.WriteText "------------------------------------------------" & vbCrLf
            
            ' 모듈 내 코드 가져오기
            codeText = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            stream.WriteText codeText & vbCrLf
            
            ' 푸터
            stream.WriteText vbCrLf & "' --- End of " & vbComp.Name & " ---" & vbCrLf & vbCrLf
        End If
    Next vbComp
    
    ' 파일 저장
    stream.SaveToFile savePath, 2 ' 2 = 덮어쓰기
    stream.Close
    Set stream = Nothing
    
    MsgBox "모든 코드가 하나의 UTF-8 파일로 합쳐졌습니다!" & vbCrLf & savePath, vbInformation
End Sub

Attribute VB_Name = "modInputParse"
' @desc 입력 원문 토큰화와 값 후보 해석 담당
Option Explicit

' @desc 멀티라인 원문 받아 고유 토큰 배열 반환, 빈 입력 시 빈 배열 반환
Public Function Tokenize(ByVal sRaw As String) As Variant
    Dim s As String
    s = Replace(sRaw, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbTab, vbLf)

    Dim parts() As String
    parts = Split(s, vbLf)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim out As New Collection
    Dim i As Long, t As String
    For i = LBound(parts) To UBound(parts)
        t = Trim$(parts(i))
        If Len(t) > 0 Then
            If Not seen.exists(t) Then
                seen(t) = True
                out.Add t
            End If
        End If
    Next i

    If out.Count = 0 Then
        Tokenize = Array()
    Else
        Dim arr() As Variant
        ReDim arr(0 To out.Count - 1)
        For i = 1 To out.Count
            arr(i - 1) = out(i)
        Next i
        Tokenize = arr
    End If
End Function

' @desc 토큰 받아 값 후보 Double 배열 반환, 해석 불가 시 빈 배열 반환
' 후보는 최대 5개(직접/통화/%/날짜 x2)라 값당 Dictionary 대신 고정 배열 + 수동 dedup.
' CreateObject("Scripting.Dictionary")가 호출마다 late-bind COM 생성(~0.1-0.3ms)이라
' 미스 토큰 전부에 걸려 대량 붙여넣기 미스 경로를 지배했다(modPerfProbe.RunValueLaneProbe 실측).
Public Function TryParseValues(ByVal sToken As String) As Variant
    Dim t As String
    t = Trim$(sToken)
    If Len(t) = 0 Then TryParseValues = Array(): Exit Function

    Dim vals(0 To 7) As Double
    Dim cnt As Long
    Dim d As Double

    ' 1. 직접 숫자 ("1000", "1,234", "3.14")
    If TryCDbl(t, d) Then PushVal vals, cnt, d

    ' 2. 통화기호/공백/괄호음수 제거 후 숫자 (해당 문자가 있을 때만 - Replace 6회 회피)
    Dim c As String
    c = t
    If HasCurrencyShape(t) Then c = StripCurrency(t)
    If c <> t Then
        If TryCDbl(c, d) Then PushVal vals, cnt, d
    End If

    ' 3. 백분율 ("50%" -> 0.5)
    If Right$(t, 1) = "%" Then
        Dim inner As String
        inner = StripCurrency(Left$(t, Len(t) - 1))
        If TryCDbl(inner, d) Then PushVal vals, cnt, d / 100#
    End If

    ' 4. 날짜/시간 (구분자 형태가 있는 토큰만 시도)
    If HasDateShape(t) Then
        Dim dt As Date
        If TryCDate(t, dt) Then PushVal vals, cnt, CDbl(dt)
        Dim nrm As String
        nrm = NormalizeHangulDate(t)
        If nrm <> t Then
            If TryCDate(nrm, dt) Then PushVal vals, cnt, CDbl(dt)
        End If
    End If

    If cnt = 0 Then
        TryParseValues = Array()
    Else
        Dim out() As Variant
        ReDim out(0 To cnt - 1)
        Dim i As Long
        For i = 0 To cnt - 1
            out(i) = vals(i)
        Next i
        TryParseValues = out
    End If
End Function

' 후보 값을 중복 없이 고정 배열에 담는다(cnt<=5라 선형 dedup, Dictionary 회피).
Private Sub PushVal(ByRef vals() As Double, ByRef cnt As Long, ByVal d As Double)
    Dim i As Long
    For i = 0 To cnt - 1
        If vals(i) = d Then Exit Sub          ' 이미 존재(수치 동일 = CStr 동일 dedup과 등가)
    Next i
    If cnt > UBound(vals) Then Exit Sub        ' 상한 방어(도달 불가)
    vals(cnt) = d
    cnt = cnt + 1
End Sub

' CDbl 변환을 오류 없이 시도한다. IsNumeric 비던짐 선검사로 텍스트 토큰의 예외 비용을
' 소멸시킨다(미스율 높은 대량 붙여넣기 핫스팟 - modPerfProbe.RunValueLaneProbe 실측:
' text_miss 2000토큰 ~1.1s가 던지는 CDbl에 지배됨). CDbl의 On Error는 IsNumeric True인데
' 오버플로 등으로 실패하는 잔여 경우의 백스톱으로 유지 - 파싱되는 값 집합은 불변.
Private Function TryCDbl(ByVal s As String, ByRef d As Double) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    On Error GoTo Fail
    d = CDbl(t)
    TryCDbl = True
Fail:
End Function

' CDate 변환을 오류 없이 시도한다. 순수 숫자 문자열은 배제(CDate의 수치 해석 방지).
Private Function TryCDate(ByVal s As String, ByRef dt As Date) As Boolean
    On Error GoTo Fail
    Dim dummy As Double
    If TryCDbl(s, dummy) Then Exit Function
    dt = CDate(s)
    TryCDate = True
Fail:
End Function

' StripCurrency가 손댈 문자($/공백/괄호/원화/NBSP)가 있는지 값싸게 확인한다
' (없으면 Replace 6회를 건너뜀 - 대량 붙여넣기 미스 경로 고정비 감량).
Private Function HasCurrencyShape(ByVal s As String) As Boolean
    HasCurrencyShape = (InStr(s, "$") > 0) Or (InStr(s, " ") > 0) _
        Or (InStr(s, "(") > 0) Or (InStr(s, ChrW(8361)) > 0) _
        Or (InStr(s, ChrW(65510)) > 0) Or (InStr(s, ChrW(160)) > 0)
End Function

' 통화기호(원/달러)/공백/NBSP를 제거하고 괄호 음수를 마이너스로 바꾼다.
Private Function StripCurrency(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)

    ' 괄호 음수: (500) -> -500
    If Len(t) >= 2 Then
        If Left$(t, 1) = "(" And Right$(t, 1) = ")" Then
            t = "-" & Mid$(t, 2, Len(t) - 2)
        End If
    End If

    t = Replace(t, ChrW(8361), "")   ' 원화 기호
    t = Replace(t, ChrW(65510), "")  ' 전각 원화 기호
    t = Replace(t, "$", "")
    t = Replace(t, ChrW(160), "")    ' NBSP
    t = Replace(t, " ", "")

    ' "--500" 같은 이중 음수 방지
    If Left$(t, 2) = "--" Then t = Mid$(t, 2)

    StripCurrency = t
End Function

' 날짜/시간으로 해석해볼 만한 구분자가 있는지 확인한다.
Private Function HasDateShape(ByVal s As String) As Boolean
    HasDateShape = (InStr(s, "-") > 0) Or (InStr(s, "/") > 0) Or (InStr(s, ":") > 0) _
                   Or (InStr(s, ChrW(&HB144)) > 0)
End Function

' 한글 날짜 표기를 CDate가 읽는 형태로 바꾼다 ("2024년 1월 3일" -> "2024-1-3").
Private Function NormalizeHangulDate(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, ChrW(&HB144), "-")   ' 년
    t = Replace(t, ChrW(&HC6D4), "-")   ' 월
    t = Replace(t, ChrW(&HC77C), "")    ' 일
    t = Replace(t, " ", "")
    NormalizeHangulDate = t
End Function

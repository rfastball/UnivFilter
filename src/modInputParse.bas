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
Public Function TryParseValues(ByVal sToken As String) As Variant
    Dim found As Object
    Set found = CreateObject("Scripting.Dictionary")

    Dim t As String
    t = Trim$(sToken)
    If Len(t) = 0 Then TryParseValues = Array(): Exit Function

    Dim d As Double

    ' 1. 직접 숫자 ("1000", "1,234", "3.14")
    If TryCDbl(t, d) Then AddVal found, d

    ' 2. 통화기호/공백/괄호음수 제거 후 숫자
    Dim c As String
    c = StripCurrency(t)
    If c <> t Then
        If TryCDbl(c, d) Then AddVal found, d
    End If

    ' 3. 백분율 ("50%" -> 0.5)
    If Right$(t, 1) = "%" Then
        Dim inner As String
        inner = StripCurrency(Left$(t, Len(t) - 1))
        If TryCDbl(inner, d) Then AddVal found, d / 100#
    End If

    ' 4. 날짜/시간 (구분자 형태가 있는 토큰만 시도)
    If HasDateShape(t) Then
        Dim dt As Date
        If TryCDate(t, dt) Then AddVal found, CDbl(dt)
        Dim n As String
        n = NormalizeHangulDate(t)
        If n <> t Then
            If TryCDate(n, dt) Then AddVal found, CDbl(dt)
        End If
    End If

    If found.Count = 0 Then
        TryParseValues = Array()
    Else
        TryParseValues = found.Items
    End If
End Function

' 후보 값을 중복 없이 등록한다.
Private Sub AddVal(ByVal found As Object, ByVal d As Double)
    Dim k As String
    k = CStr(d)
    If Not found.exists(k) Then found(k) = d
End Sub

' CDbl 변환을 오류 없이 시도한다.
Private Function TryCDbl(ByVal s As String, ByRef d As Double) As Boolean
    On Error GoTo Fail
    If Len(Trim$(s)) = 0 Then Exit Function
    d = CDbl(s)
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

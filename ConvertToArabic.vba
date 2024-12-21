Function ConvertToArabic(inputText As String) As String
    Dim charMap As Object
    Dim result As String
    Dim i As Integer
    Dim currentChar As String
    Dim arabicPart As String
    Dim numberPart As String
    
    ' إنشاء خريطة الحروف
    Set charMap = CreateObject("Scripting.Dictionary")
    charMap.Add "A", ChrW(1571) ' أ
    charMap.Add "B", ChrW(1576) ' ب
    charMap.Add "J", ChrW(1581) ' ح
    charMap.Add "D", ChrW(1583) ' د
    charMap.Add "R", ChrW(1585) ' ر
    charMap.Add "S", ChrW(1587) ' س
    charMap.Add "X", ChrW(1589) ' ص
    charMap.Add "T", ChrW(1591) ' ط
    charMap.Add "E", ChrW(1593) ' ع
    charMap.Add "G", ChrW(1602) ' ق
    charMap.Add "K", ChrW(1603) ' ك
    charMap.Add "L", ChrW(1604) ' ل
    charMap.Add "Z", ChrW(1605) ' م
    charMap.Add "N", ChrW(1606) ' ن
    charMap.Add "H", ChrW(1607) ' ه
    charMap.Add "U", ChrW(1608) ' و
    charMap.Add "V", ChrW(1609) ' ى
    
    ' تهيئة النتائج
    result = ""
    arabicPart = ""
    numberPart = ""
    
    ' قراءة النص
    For i = 1 To Len(inputText)
        currentChar = Mid(inputText, i, 1)
        If IsNumeric(currentChar) Then
            numberPart = numberPart & currentChar
        ElseIf charMap.exists(currentChar) Then
            arabicPart = arabicPart & charMap(currentChar) & " " ' إضافة مسافة بعد كل حرف
        End If
    Next i
    
    ' عكس ترتيب الحروف العربية
    arabicPart = Trim(arabicPart)
    arabicPart = StrReverse(arabicPart) ' عكس الحروف

    ' دمج النصوص: الحروف العربية (عكس الترتيب) أولاً ثم الأرقام
    ConvertToArabic = arabicPart & " " & numberPart
End Function

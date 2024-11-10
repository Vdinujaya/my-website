<%
' Array with names
Dim a(30)
a(0) = "Anna"
a(1) = "Brittany"
a(2) = "Cinderella"
a(3) = "Diana"
a(4) = "Eva"
a(5) = "Fiona"
a(6) = "Gunda"
a(7) = "Hege"
a(8) = "Inga"
a(9) = "Johanna"
a(10) = "Kitty"
a(11) = "Linda"
a(12) = "Nina"
a(13) = "Ophelia"
a(14) = "Petunia"
a(15) = "Amanda"
a(16) = "Raquel"
a(17) = "Cindy"
a(18) = "Doris"
a(19) = "Eve"
a(20) = "Evita"
a(21) = "Sunniva"
a(22) = "Tove"
a(23) = "Unni"
a(24) = "Violet"
a(25) = "Liza"
a(26) = "Elizabeth"
a(27) = "Ellen"
a(28) = "Wenche"
a(29) = "Vicky"

' Get the 'q' parameter from the URL
Dim q, hint, i, len
q = Request.QueryString("q")
hint = ""

' If q is not empty, perform the lookup
If Len(q) > 0 Then
    q = LCase(q)
    len = Len(q)
    
    ' Loop through the array of names
    For i = 0 To UBound(a)
        ' If name starts with the string in q, add to the hint
        If InStr(1, LCase(a(i)), Left(q, len)) > 0 Then
            If hint = "" Then
                hint = a(i)
            Else
                hint = hint & ", " & a(i)
            End If
        End If
    Next
End If

' Output the hint or "no suggestion" if no matches found
If hint = "" Then
    Response.Write "no suggestion"
Else
    Response.Write hint
End If
%>

<!DOCTYPE html>
<html>
    <head>
        <title>SECURITY CLASSES</title>
    </head>
<body>
<% 
'Coded by Muhammad Aizat
Class Security
    Private function ereg_replace(strOriginalString, strPattern, strReplacement)
        dim objRegExp : set objRegExp = new RegExp
        objRegExp.Pattern = strPattern
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
        set objRegExp = nothing
    End function

    Private function URLDecode(sText)
        sDecoded = sText
        Set oRegExpr = Server.CreateObject("VBScript.RegExp")
        oRegExpr.Pattern = "%[0-9,A-F]{2}"
        oRegExpr.Global = True
        Set oMatchCollection = oRegExpr.Execute(sText)
        For Each oMatch In oMatchCollection
            sDecoded = Replace(sDecoded,oMatch.value,Chr(CInt("&H" & Right(oMatch.Value,2))))
        Next
        URLDecode = sDecoded
    End function

    Private function ereg_sqli(strOriginalString)
        Dim quotes
        quotes =  array("%27","%22")
        For i=0 to (uBound(quotes))
            strOriginalString = ereg_replace(Server.URLEncode(strOriginalString),quotes(i),"/"&URLDecode(quotes(i)))
            strOriginalString = URLDecode(strOriginalString)
        Next
        ereg_sqli = URLDecode(strOriginalString)
    End function


    Public function injection(strText,strType,lvl)
        Dim point
        Select Case strType
            Case "sqli"
                if lvl=1 Then
                    strText = ereg_replace(strText,"[\'\-\#\""]", "")
                ElseIf lvl=2 Then
                    strText = ereg_sqli(strText)
                    strText = ereg_replace(strText,"[\+]"," ")
                End if
            Case "xss"
                strText = Server.HTMLEncode(strText)
            Case "spcial"
                strText = ereg_replace(strText,"[^A-Za-z0-9\s]", "")
        End Select
        injection = strText
    End function

End Class
%>
<% 
Dim sec
Set sec = new Security

response.write("Escape only html character<br>")

'XSS Example
Dim xss
xss = sec.injection(request.querystring("str"),"xss",null)
response.write(xss)

response.write("<br>Remove All Special Character<br>")

'Remove All Special Character
Dim spc
spc = sec.injection(request.querystring("str"),"spcial",null)
response.write(spc)

response.write("<br>Remove only single&double quote , minus and hashtag symbol<br>")

'SQL Injection lvl 1
Dim sql1
sql1 = sec.injection(request.querystring("str"),"sqli",1)
response.write(sql1)

response.write("<br>Single and double quote escape string<br>")

'SQL Injection lvl 2
Dim sql2
sql2 = sec.injection(request.querystring("str"),"sqli",2)
response.write(sql2)

%>
</body>
</html>

<!DOCTYPE html>
<html>
    <head>
        <title>EagleEye ASP Sanitizing</title>
    </head>
<body>
    <p>
       Click <a href="?str=!@%23$<code>abc</code>'%22 union select 1,2-- -%23..//%252e%252e//*&file=lol.jpg">Here</a>
    </p>
<!--#include virtual="security.inc"-->
<% 
response.write("Original String : "&request.querystring("str")&"<br><br>")
Dim sec
Set sec = new Security

response.write("Escape only html character<br>")
'XSS Example
Dim xss
xss = sec.htmlEsc(request.querystring("str"))
response.write(xss&"<br>")
response.write("<img src='"&xss&"'>")

response.write("<br><br>Remove All Special Character<br>")
'Remove All Special Character
Dim spc
spc = sec.noSpecialChar(request.querystring("str"))
response.write(spc)

response.write("<br><br>Remove only single&double quote , minus and hashtag symbol<br>")
'SQL Injection lvl 1
Dim sql1
sql1 = sec.rmSqliChar(request.querystring("str"))
response.write(sql1)

response.write("<br><br>Single and double quote escape string<br>")
'SQL Injection lvl 2
Dim sql2
sql2 = sec.sqlEscStr(request.querystring("str"))
response.write(sql2)

response.write("<br><br>Allowed Extension Image&Docs<br>")
'File extension filter
Dim ftx
ftx = sec.extAllowed(request.querystring("file"),null)
response.write("Allowed : "&ftx)

response.write("<br><br>Basename : ")
'File extension filter
response.write(sec.basename(request.querystring("file")))

response.write("<br><br>Is Email valid? :")
'Email filter
response.write("<br>mail's@email.com : ")
response.write(sec.isEmail("mail's@email.com"))
response.write("<br>eagle@email.com : ")
response.write(sec.isEmail("eagle@email.com"))

%>
</body>
</html>

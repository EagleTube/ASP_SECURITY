<!DOCTYPE html>
<html>
    <head>
        <title>EagleEye ASP Sanitizing</title>
    </head>
<body>
    <p>
       Click <a href="?str=!@%23$<code>abc</code>'%22 union select 1,2-- -%23..//%252e%252e//*&img=lol.txt.html">Here</a>
    </p>
<!--#include virtual="security.inc"-->
<% 
response.write("Original String : "&request.querystring("str")&"<br><br>")
Dim sec
Set sec = new Security

response.write("Escape only html character<br>")
'XSS Example
Dim xss
xss = sec.sanitize(request.querystring("str"),"htmlesc",null)
response.write(xss)

response.write("<br><br>Remove All Special Character<br>")
'Remove All Special Character
Dim spc
spc = sec.sanitize(request.querystring("str"),"rmspc",null)
response.write(spc)

response.write("<br><br>Remove only single&double quote , minus and hashtag symbol<br>")
'SQL Injection lvl 1
Dim sql1
sql1 = sec.sanitize(request.querystring("str"),"sqlesc",1)
response.write(sql1)

response.write("<br><br>Single and double quote escape string<br>")
'SQL Injection lvl 2
Dim sql2
sql2 = sec.sanitize(request.querystring("str"),"sqlesc",2)
response.write(sql2)

response.write("<br><br>Dot and Slashes remover<br>")
'Path traversal
Dim trv
trv = sec.sanitize(request.querystring("str"),"trav",null)
response.write(trv)

response.write("<br><br>Allowed Extension Image/Docs<br>")
'File extension filter
Dim ftx, ftx1
ftx = sec.sanitize(request.querystring("img"),"fext",1)
ftx1 = sec.sanitize(request.querystring("img"),"fext",2)
response.write("Image ext : "&ftx&"<br>")
response.write("Document ext : "&ftx1)
%>
</body>
</html>

<%
   
'  Coded by Eagle Eye
'  Github : https://github.com/EagleTube
'  Youtube : https://www.youtube.com/c/EagleTube1337
'  If you found a bypass please comment
   
Class Security

    private defaultAllowedArray

'Initialize for default allowed extension

	Private Sub Class_Initialize() 
        defaultAllowedArray = array("jpg","gif","bmp","jpeg","png","tiff","tif","eps","doc","docx","pdf","csv","xls","xlsx","ppt","pptx","odt")
	End Sub

' Checking user input email validation
' Are the user really put the real email address?
' This function help you to validate the string that been inserted by the users
' Example : isEmail("myemail.com") -> False
' Example : isEmail("myemail@email.com") -> True

    Public function isEmail(strAddr) 
        dim isBool
        dim regEx
        isBool = True
        set regEx = New RegExp
        regEx.IgnoreCase = False
        regEx.Pattern = "^[-+.\w]{1,64}@[-.\w]{1,64}\.[-.\w]{2,6}$"
        isBool = regEx.Test(strAddr)
        isEmail = isBool
    End function

' Check existing of file/folder
' This function to ease us detect our local file or folder exist in our webserver
' We have to specific a folder/path with file name
' Example: isExists("../../flag.txt") -> True (when it found) | False (When not found)

    Public function isExists(fname) 
        dim fsys
        set fsys=Server.CreateObject("Scripting.FileSystemObject")
        if fsys.FileExists(fname)=true then
            file_exists=true
        else
            file_exists=false
        end if
        set fsys=nothing
    End function

' Used for match pattern and replace with new valie
' This function must provide a pattern , or just specific character/string
' After providing a pattern , originalstring will taken for regex and match
' If match the pattern , the matched char/string from originalstring will replaced with what we want
' Example : ereg_replace("Hello Im Eagle","Eagle","Security Researcher") -> Hello Im Security Researcher

    Public function ereg_replace(strOriginalString, strPattern, strReplacement) 
        dim objRegExp : set objRegExp = new RegExp
        objRegExp.Pattern = strPattern
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
        set objRegExp = nothing
    End function

' Having trouble decoding urlencoded string?
' This function are the solution to decode urlencoded string
' Example : URLDecode("%2F%3D") -> /=

    Private function URLDecode(sText) 
        sDecoded = sText
        Set oRegExpr = Server.CreateObject("VBScript.RegExp")
        oRegExpr.Pattern = "%[0-9,A-F,a-f]{2}"
        oRegExpr.Global = True
        Set oMatchCollection = oRegExpr.Execute(sText)
        For Each oMatch In oMatchCollection
            sDecoded = Replace(sDecoded,oMatch.value,Chr(CInt("&H" & Right(oMatch.Value,2))))
        Next
        URLDecode = sDecoded
    End function

' This function about to same as real_escape_string in php
' Private members only be used internal class

    Private function ereg_sqli(strOriginalString) 
        Dim quotes, slashes
        quotes =  array("%27","%22","%00","%0a","%0d")
        backslash = "%5c" 'escaper
        strOriginalString = URLDecode(ereg_replace(Server.URLEncode(strOriginalString),backslash,"%5c"&backslash))
        For i=0 to (uBound(quotes))
            strOriginalString = ereg_replace(Server.URLEncode(strOriginalString),quotes(i),"\"&URLDecode(quotes(i)))
            strOriginalString = URLDecode(strOriginalString)
        Next
        ereg_sqli = strOriginalString
    End function

' Same like basename() function in php where it takes the only filename 
' At strText where we have to pass a value which point to a filename
' This function help to avoid path traversal attack
' The slash and dotdot will be ignored and the value will return only filename
' Example : basename('../../../../Hello.asp') -> Hello.asp
' Example : basename('Folder/1/2/3/file.asp') -> file.asp

    Public function basename(strText) 
        Dim esplit,out
        Dim strArr : strArr = array("%2F","%5C","%25")
        For Each path in strArr
            esplit = split(server.URLEncode(strText),path)
            If uBound(esplit) > 0 Then
                out = esplit(uBound(esplit))
                out = ereg_replace(out,"%25","%")
                basename = URLDecode(out)
                exit function
            End If
        Next
        basename = URLDecode(strText)
    End function

' This function can help both sql and xss injection (but still can be bypassed)
' Its main objective to remove the character used for sql or xss injection
' It will remove *,',-,#," chars

    Public function rmSqliChar(strText) 
        strText = ereg_replace(strText,"[\*\'\-\#\""]", "")
        rmSqliChar = strText
    End function

' Escape character that been used for sql injection, just like real_escape_string
' This function will escape specific char like \r,\n,',",\
' After escape those string , there wont happen sql injection

    Public function sqlEscStr(strText)
        strText = ereg_sqli(strText)
        strText = ereg_replace(strText,"[\+]"," ")
        sqlEscStr = strText
    End function

' In php there are htmlentities and htmlspecialchars functions
' So this function are about to same like that
' Escape all html character to protect from html&xss injection
' It also work against xss from html attributes where it escape quotes

    Public function htmlEsc(strText) 
        quoteArr = array("%22","%27")
        For i=0 to (uBound(quoteArr))
            strText = ereg_replace(server.URLEncode(strText),quoteArr(i),"&quot;")
            strText = URLDecode(strText)
        Next
        strText = Server.HTMLEncode(strText)
        htmlEsc = ereg_replace(strText,"&amp;","&")
    End function

' Accept only words and number , remove all special characters
' This function are suitable to used when any user string input only allowed words and number
' The function might help to reduce any type of attackss

    Public function noSpecialChar(strText)
        strText = ereg_replace(strText,"[^A-Za-z0-9\s]", "")
        noSpecialChar = strText
    End function

' Used to filter and get only allowed extension to return true, else will return false
' This function are suitable to be used when uploading a file
' The file name will be placed to strText
' you can provide your own allowed extension to strAllowedArray
' example : myAllowedExt = array("jpg","gif","doc","docx")
' call the Security class , example : myObj.extAllowed("filename.txt",myAllowedExt)

    Public function extAllowed(strText,strAllowedArray) 
        Dim extArr,extName
        extArr = split(strText,".")
        extName = extArr(uBound(extArr))
        if strAllowedArray <> null Then
            For i=0 to (uBound(strAllowedArray))
                if strAllowedArray(i)=extName Then
                    extAllowed = true
                    exit function
                End if
            Next
        else
            For i=0 to (uBound(defaultAllowedArray))
                if defaultAllowedArray(i)=extName Then
                    extAllowed = true
                    exit function
                End if
            Next
        End if
        extAllowed = false
    End function


End Class
%>

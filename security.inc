<%
'  Coded by Eagle Eye
'  Github : https://github.com/EagleTube
'  Youtube : https://www.youtube.com/c/EagleTube1337
Class Security

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
        oRegExpr.Pattern = "%[0-9,A-F,a-f]{2}"
        oRegExpr.Global = True
        Set oMatchCollection = oRegExpr.Execute(sText)
        For Each oMatch In oMatchCollection
            sDecoded = Replace(sDecoded,oMatch.value,Chr(CInt("&H" & Right(oMatch.Value,2))))
        Next
        URLDecode = sDecoded
    End function

    Private function ereg_sqli(strOriginalString)
        Dim quotes, rep
        quotes =  array("%2f","%5c","%27","%22","%25")
        For i=0 to (uBound(quotes))
            strOriginalString = ereg_replace(Server.URLEncode(strOriginalString),quotes(i),"/"&URLDecode(quotes(i)))
            strOriginalString = URLDecode(strOriginalString)
        Next
        ereg_sqli = (strOriginalString)
    End function


    Public function sanitize(strText,strType,lvl)
        Dim imgExt
        Dim docExt
        Dim extArr
        Dim extName

        imgExt = array("jpg","gif","bmp","jpeg","png","tiff","tif","eps")
        docExt = array("doc","docx","pdf","csv","xls","xlsx","ppt","pptx","odt")

        Select Case strType
            Case "sqlesc"
                if lvl=1 Then
                    strText = ereg_replace(strText,"[\*\'\-\#\""]", "")
                ElseIf lvl=2 Then
                    strText = ereg_sqli(strText)
                    strText = ereg_replace(strText,"[\+]"," ")
                End if
            Case "htmlesc"
                strText = Server.HTMLEncode(strText)
            Case "rmspc"
                strText = ereg_replace(strText,"[^A-Za-z0-9\s]", "")
            Case "trav"
                strText = ereg_replace(strText,"[\%\.\/\\]", "")
            Case "fext"
                extArr = split(strText,".")
                extName = extArr(uBound(extArr))
                if lvl=1 Then
                    For i=0 to (uBound(imgExt))
                        if imgExt(i)=extName Then
                            sanitize = true
                            exit function
                        End if
                    Next
                    strText = false
                ElseIf lvl=2 Then
                    For i=0 to (uBound(docExt))
                        if docExt(i)=extName Then
                            sanitize = true
                            exit function
                        End if
                    Next
                    strText = false
                End if
        End Select
        sanitize = strText
    End function

End Class
%>

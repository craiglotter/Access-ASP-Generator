Imports System.IO

Public Class Template_Generator

    Function Get_Template_Folder() As String
        Dim f As FileInfo
        f = New FileInfo(Application.ExecutablePath())
        Dim dir As Directory
        If dir.Exists((f.DirectoryName() & "\ASP_Templates")) = False Then
            dir.CreateDirectory((f.DirectoryName() & "\ASP_Templates"))
        End If
        Return (f.DirectoryName() & "\ASP_Templates")
    End Function


    Function Generate_Simple_Create_Page() As Boolean
        Try
            Dim Template_Folder As String = Get_Template_Folder()
            Dim filewriter As New StreamWriter(Template_Folder & "\Template_Simple_Create_Page.txt", False, System.Text.Encoding.UTF8)

            filewriter.WriteLine("<%")
            filewriter.WriteLine("Dim filename, conn")
            filewriter.WriteLine("filename = server.mappath(""fpdb/#insertdatabase#"")")
            filewriter.WriteLine("Set conn = server.createobject(""ADODB.Connection"")")
            filewriter.WriteLine("conn.mode = 3")
            filewriter.WriteLine("Dim strCon")
            filewriter.WriteLine("strCon = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="" & filename & "";""")
            filewriter.WriteLine("Dim createcomplete")
            filewriter.WriteLine("createcomplete = """"")
            filewriter.WriteLine("createcomplete = request.querystring(""createcomplete"")")
            filewriter.WriteLine("Dim Page_Action")
            filewriter.WriteLine("Page_Action = """"")
            filewriter.WriteLine("Page_Action = Trim(Replace(request.querystring(""Page_Action""), ""'"", """"))")
            filewriter.WriteLine("If Page_Action = ""create"" Then")
            filewriter.WriteLine("#retrievevariables#")
            filewriter.WriteLine("conn.open strCon")
            filewriter.WriteLine("Set objRec = server.createobject(""ADODB.Recordset"")")
            filewriter.WriteLine("objRec.ActiveConnection = conn")
            filewriter.WriteLine("strSQL = ""select count(#insertfieldkey#) as numberofresults from #inserttable# where #insertfieldkey# = '"" & #insertfieldkey# & ""'""")
            filewriter.WriteLine("objRec.open(strSQL)")
            filewriter.WriteLine("insertflag = ""true""")
            filewriter.WriteLine("#validateentries#")
            filewriter.WriteLine("If objRec.fields(""numberofresults"").value > 0 Then")
            filewriter.WriteLine("objRec.close()")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("insertflag = ""false""")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<script Anime_Language=""javascript"">")
            filewriter.WriteLine("alert(""The #insertobjectdescription# cannot be added as the record key you specified is already included in the database"");")
            filewriter.WriteLine("window.location = ""#insertfilename#""")
            filewriter.WriteLine("</script>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("If insertflag = ""true"" Then")
            filewriter.WriteLine("#generateinsertSQL#")
            filewriter.WriteLine("conn.execute countersql")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("response.redirect(""#insertfilename#?createcomplete=true"")")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<html>")
            filewriter.WriteLine("<head>")
            filewriter.WriteLine("<Title>#insertpagetitle#</Title>")
            filewriter.WriteLine("</head>")
            filewriter.WriteLine("<body>")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<table cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse"" bordercolor=""111111"">")
            filewriter.WriteLine("<tr>")
            filewriter.WriteLine("<td align=""center"">")
            filewriter.WriteLine("<hr color=""000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("If Not createcomplete = ""true"" Then")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<h3>Create #insertobjectdescription#</h3>")
            filewriter.WriteLine("Fill in the input fields below  to create a new #insertobjectdescription#:")
            filewriter.WriteLine("<br>")
            filewriter.WriteLine("<form name=""create"" method=""POST"" action=""#insertfilename#?Page_Action=create"">")
            filewriter.WriteLine("<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""111111"" width=""62%"">")
            filewriter.WriteLine("#createtextboxes#")
            filewriter.WriteLine("</table>")
            filewriter.WriteLine("<p align=""center"">")
            filewriter.WriteLine("<input type=""button"" value=""Create #insertobjectdescription#"" name=""b1"" onclick=""submit();""></p>")
            filewriter.WriteLine("</form>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("Else")
            filewriter.WriteLine("response.write(""#insertobjectdescription# Successfully Created"")")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("<hr color=""000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("</td>")
            filewriter.WriteLine("</tr>")
            filewriter.WriteLine("</table>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("</body>")
            filewriter.WriteLine("</html>")

            filewriter.Close()
            Return True
        Catch e As Exception
            Return False
        End Try
    End Function

    Function Generate_Simple_Remove_Page() As Boolean
        Try
            Dim Template_Folder As String = Get_Template_Folder()
            Dim filewriter As New StreamWriter(Template_Folder & "\Template_Simple_Remove_Page.txt", False, System.Text.Encoding.UTF8)

            filewriter.WriteLine("<%")
            filewriter.WriteLine("Dim filename, conn")
            filewriter.WriteLine("filename=server.mappath(""fpdb/#insertdatabase#"")")
            filewriter.WriteLine("Set conn = server.createobject(""ADODB.Connection"")")
            filewriter.WriteLine("conn.mode = 3")
            filewriter.WriteLine("Dim strCon")
            filewriter.WriteLine("strCon = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="" & filename & "";""")
            filewriter.WriteLine("Dim removecomplete")
            filewriter.WriteLine("removecomplete = """"")
            filewriter.WriteLine("removecomplete = request.querystring(""removecomplete"")")
            filewriter.WriteLine("Page_Action = """"")
            filewriter.WriteLine("Page_Action = trim(replace(request.querystring(""Page_Action""),""'"",""""))")
            filewriter.WriteLine("#insertfieldkey# = trim(replace(request.form(""#insertfieldkey#""),""'"",""""))")
            filewriter.WriteLine("if #insertfieldkey# = """" then")
            filewriter.WriteLine("#insertfieldkey# = Session(""S_#insertfieldkey#"")")
            filewriter.WriteLine("end if")
            filewriter.WriteLine("Session(""S_#insertfieldkey#"") = #insertfieldkey#")
            filewriter.WriteLine("Confirm = trim(replace(request.querystring(""Confirm""),""'"",""""))")
            filewriter.WriteLine("if Page_Action = ""delete"" then")
            filewriter.WriteLine("if Confirm = ""true"" then")
            filewriter.WriteLine("conn.open(strCon)")
            filewriter.WriteLine("countersql = ""delete from #inserttable# where  #insertfieldkey# ='""& #insertfieldkey# &""'"" ")
            filewriter.WriteLine("conn.execute(countersql)")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("response.redirect ""#insertfilename#?removecomplete=true""")
            filewriter.WriteLine("Else")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<script language=""javascript"">")
            filewriter.WriteLine("var bMyVar = confirm(""Are you sure you wish to delete this #insertobjectdescription#?\n(You cannot undo this operation)"");")
            filewriter.WriteLine("if (bMyVar == true)")
            filewriter.WriteLine("{")
            filewriter.WriteLine("window.location = ""#insertfilename#?Confirm=true&Page_Action=delete""")
            filewriter.WriteLine("}")
            filewriter.WriteLine("Else")
            filewriter.WriteLine("{")
            filewriter.WriteLine("window.location = ""#insertfilename#""")
            filewriter.WriteLine("}")
            filewriter.WriteLine("</script> ")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<html>")
            filewriter.WriteLine("<head>")
            filewriter.WriteLine("<Title>#insertpagetitle#</Title>")
            filewriter.WriteLine("</head>")
            filewriter.WriteLine("<body>")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<table cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"">")
            filewriter.WriteLine("<tr>")
            filewriter.WriteLine("<td align=""center"">")
            filewriter.WriteLine("<hr color=""#000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<% ")
            filewriter.WriteLine("if not removecomplete = ""true"" then")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<h3>Remove #insertobjectdescription#</h3>")
            filewriter.WriteLine("Select from the list, the #insertobjectdescription# to remove:")
            filewriter.WriteLine("<form name=""delete"" method=""POST"" action=""#insertfilename#?Page_Action=delete"">")
            filewriter.WriteLine("<p align=""center"">")
            filewriter.WriteLine("<select size=""1"" name=""#insertfieldkey#"">")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("conn.open strCon")
            filewriter.WriteLine("Set objRec = server.createobject(""ADODB.Recordset"")")
            filewriter.WriteLine("strSQL = ""select #insertfieldkey# from #inserttable# order by #insertfieldkey#""")
            filewriter.WriteLine("objRec.ActiveConnection = conn")
            filewriter.WriteLine("objRec.open strSQL")
            filewriter.WriteLine("do while not objRec.eof")
            filewriter.WriteLine("response.write ""<option>""& objRec.fields(""#insertfieldkey#"").value &""</option>""")
            filewriter.WriteLine("objRec.moveNext")
            filewriter.WriteLine("loop")
            filewriter.WriteLine("objRec.close")
            filewriter.WriteLine("conn.close")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("</select>")
            filewriter.WriteLine("</p>")
            filewriter.WriteLine("<p align=""center"">")
            filewriter.WriteLine("<input type=""submit"" value=""Remove #insertobjectdescription#"" name=""B1""></p>")
            filewriter.WriteLine("</form>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("else")
            filewriter.WriteLine("response.write ""#insertobjectdescription# Successfully Removed""")
            filewriter.WriteLine("end if")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("<hr color=""#000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("</td>")
            filewriter.WriteLine("</tr>")
            filewriter.WriteLine("</table>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("</body>")
            filewriter.WriteLine("</html>")

            filewriter.Close()
            Return True
        Catch e As Exception
            Return False
        End Try
    End Function

    Function Generate_Simple_Edit_Page() As Boolean
        Try
            Dim Template_Folder As String = Get_Template_Folder()
            Dim filewriter As New StreamWriter(Template_Folder & "\Template_Simple_Edit_Page.txt", False, System.Text.Encoding.UTF8)

            filewriter.WriteLine("<%")
            filewriter.WriteLine("Dim filename, conn")
            filewriter.WriteLine("filename=server.mappath(""fpdb/#insertdatabase#"")")
            filewriter.WriteLine("Set conn = server.createobject(""ADODB.Connection"")")
            filewriter.WriteLine("conn.mode = 3")
            filewriter.WriteLine("Dim strCon")
            filewriter.WriteLine("strCon = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="" & filename & "";""")
            filewriter.WriteLine("Dim editcomplete")
            filewriter.WriteLine("editcomplete = """"")
            filewriter.WriteLine("editcomplete = request.querystring(""editcomplete"")")
            filewriter.WriteLine("Page_Action = """"")
            filewriter.WriteLine("Page_Action = trim(replace(request.querystring(""Page_Action""),""'"",""""))")
            filewriter.WriteLine("#insertfieldkey# = trim(replace(request.form(""#insertfieldkey#""),""'"",""""))")
            filewriter.WriteLine("if #insertfieldkey# = """" then")
            filewriter.WriteLine("#insertfieldkey# = Session(""S_#insertfieldkey#"")")
            filewriter.WriteLine("end if")
            filewriter.WriteLine("Session(""S_#insertfieldkey#"") = #insertfieldkey#")
            filewriter.WriteLine("If Page_Action = ""update_details"" Then")
            filewriter.WriteLine("#retrievevariables#")
            filewriter.WriteLine("old#insertfieldkey# = trim(replace(request.form(""old#insertfieldkey#""),""'"",""`""))")
            filewriter.WriteLine("insertflag = ""true""")
            filewriter.WriteLine("if #insertfieldkey# = """" then")
            filewriter.WriteLine("insertflag = ""false""")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<script Language=""javascript"">")
            filewriter.WriteLine("alert(""The #insertobjectdescription# cannot be updated as not all of the required text fields have been filled in"");")
            filewriter.WriteLine("window.location = ""#insertfilename#""")
            filewriter.WriteLine("</script>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("if not #insertfieldkey# = old#insertfieldkey# then")
            filewriter.WriteLine("conn.open(strCon)")
            filewriter.WriteLine("Set objRec = server.createobject(""ADODB.Recordset"")")
            filewriter.WriteLine("objRec.ActiveConnection = conn")
            filewriter.WriteLine("strSQL = ""select count(#insertfieldkey#) as numberofresults from #inserttable# where #insertfieldkey# = '"" & #insertfieldkey# & ""'""")
            filewriter.WriteLine("objRec.open(strSQL)")
            filewriter.WriteLine("If objRec.fields(""numberofresults"").value > 0 Then")
            filewriter.WriteLine("objRec.close()")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("insertflag = ""false""")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<script Anime_Language=""javascript"">")
            filewriter.WriteLine("alert(""The #insertobjectdescription# cannot be updated as the record key you specified is already included in the database"");")
            filewriter.WriteLine("window.location = ""#insertfilename#""")
            filewriter.WriteLine("</script>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("Else")
            filewriter.WriteLine("objRec.close()")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("If insertflag = ""true"" Then")
            filewriter.WriteLine("#generateupdateSQL#")
            filewriter.WriteLine("conn.open strCon")
            filewriter.WriteLine("conn.execute(countersql)")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("response.redirect(""#insertfilename#?editcomplete=true"")")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("End If")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<html>")
            filewriter.WriteLine("<head>")
            filewriter.WriteLine("<Title>#insertpagetitle#</Title>")
            filewriter.WriteLine("</head>")
            filewriter.WriteLine("<body>")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<table cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"">")
            filewriter.WriteLine("<tr>")
            filewriter.WriteLine("<td align=""center"">")
            filewriter.WriteLine("<hr color=""#000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("if not editcomplete = ""true"" then")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<h3>Edit #insertobjectdescription#</h3>")
            filewriter.WriteLine("<% if not Page_Action = ""edit_details"" then %>")
            filewriter.WriteLine("Select from the list, the #insertobjectdescription# to edit:")
            filewriter.WriteLine("<form name=""delete"" method=""POST"" action=""#insertfilename#?Page_Action=edit_details"">")
            filewriter.WriteLine("<p align=""center"">")
            filewriter.WriteLine("<select size=""1"" name=""#insertfieldkey#"">")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("conn.open(strCon)")
            filewriter.WriteLine("Set objRec = server.createobject(""ADODB.Recordset"")")
            filewriter.WriteLine("strSQL = ""select #insertfieldkey# from #inserttable# order by #insertfieldkey#""")
            filewriter.WriteLine("objRec.ActiveConnection = conn")
            filewriter.WriteLine("objRec.open(strSQL)")
            filewriter.WriteLine("Do While Not objRec.eof")
            filewriter.WriteLine("response.write ""<option>""& objRec.fields(""#insertfieldkey#"").value &""</option>""")
            filewriter.WriteLine("objRec.moveNext()")
            filewriter.WriteLine("Loop")
            filewriter.WriteLine("objRec.close()")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("</select>")
            filewriter.WriteLine("</p>")
            filewriter.WriteLine("<p align=""center"">")
            filewriter.WriteLine("<input type=""submit"" value=""Edit #insertobjectdescription#"" name=""Edit_Button""></p>")
            filewriter.WriteLine("</form>")
            filewriter.WriteLine("<% else %>")
            filewriter.WriteLine("<form name=""create"" method=""POST"" action=""#insertfilename#?Page_Action=update_details"">")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("conn.open(strCon)")
            filewriter.WriteLine("Set objRec = server.createobject(""ADODB.Recordset"")")
            filewriter.WriteLine("strSQL = ""select * from #inserttable# where #insertfieldkey# = '"" & #insertfieldkey# & ""'""")
            filewriter.WriteLine("objRec.ActiveConnection = conn")
            filewriter.WriteLine("objRec.open(strSQL)")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""111111"" width=""62%"">")
            filewriter.WriteLine("#createfilledtextboxes#")
            filewriter.WriteLine("<input type=""hidden"" name=""old#insertfieldkey#"" value=""<% response.write objRec.fields(""#insertfieldkey#"").value %>"">")
            filewriter.WriteLine("</table>")
            filewriter.WriteLine("<p align=""center"">")
            filewriter.WriteLine("<input type=""button"" value=""Update #insertobjectdescription#"" name=""b1"" onclick=""submit();""></p>")
            filewriter.WriteLine("</form>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("objRec.close()")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<% end if %>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("else")
            filewriter.WriteLine("response.write ""#insertobjectdescription# Successfully Edited""")
            filewriter.WriteLine("end if")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("<hr color=""#000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("</td>")
            filewriter.WriteLine("</tr>")
            filewriter.WriteLine("</table>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("</body>")
            filewriter.WriteLine("</html>")
            filewriter.Close()
            Return True
        Catch e As Exception
            Return False
        End Try
    End Function

    Function Generate_Simple_Display_Page() As Boolean
        Try
            Dim Template_Folder As String = Get_Template_Folder()
            Dim filewriter As New StreamWriter(Template_Folder & "\Template_Simple_Display_Page.txt", False, System.Text.Encoding.UTF8)

            filewriter.WriteLine("<%")
            filewriter.WriteLine("Dim filename, conn")
            filewriter.WriteLine("filename=server.mappath(""fpdb/#insertdatabase#"")")
            filewriter.WriteLine("Set conn = server.createobject(""ADODB.Connection"")")
            filewriter.WriteLine("conn.mode = 3")
            filewriter.WriteLine("Dim strCon")
            filewriter.WriteLine("strCon = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="" & filename & "";""")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<html>")
            filewriter.WriteLine("<head>")
            filewriter.WriteLine("<Title>#insertpagetitle#</Title>")
            filewriter.WriteLine("</head>")
            filewriter.WriteLine("<body>")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<table cellspacing=""0"" cellpadding=""0"" border=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"">")
            filewriter.WriteLine("<tr>")
            filewriter.WriteLine("<td align=""center"">")
            filewriter.WriteLine("<hr color=""#000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("<center>")
            filewriter.WriteLine("<h3>#insertobjectdescription#</h3>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("conn.open(strCon)")
            filewriter.WriteLine("Set objRec = server.createobject(""ADODB.Recordset"")")
            filewriter.WriteLine("strSQL = ""select * from #inserttable# order by #insertfieldkey#""")
            filewriter.WriteLine("objRec.ActiveConnection = conn")
            filewriter.WriteLine("objRec.open(strSQL)")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("<table border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#C0C0C0"" width=""62%"">")
            filewriter.WriteLine("#createdisplaytable#")
            filewriter.WriteLine("</table>")
            filewriter.WriteLine("<%")
            filewriter.WriteLine("objRec.close()")
            filewriter.WriteLine("conn.close()")
            filewriter.WriteLine("%>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("<br>")
            filewriter.WriteLine("<hr color=""#000000"" width=""50%"" size=""1"">")
            filewriter.WriteLine("</td>")
            filewriter.WriteLine("</tr>")
            filewriter.WriteLine("</table>")
            filewriter.WriteLine("</center>")
            filewriter.WriteLine("</body>")
            filewriter.WriteLine("</html>")
            filewriter.Close()
            Return True
        Catch e As Exception
            Return False
        End Try
    End Function
End Class
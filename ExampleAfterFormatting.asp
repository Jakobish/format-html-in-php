<%@  language="VBScript" %>
<%Response.Buffer=true%>
<meta http-equiv="Content-Type" content="text/html; charset=Visual">
<!-- #include file=CheckSession.asp -->
<!-- #include file=include/ViewFunctions.asp -->
<!-- #include file=include/BidiFunctions.asp -->
<!-- #include file=include/const.asp -->
<!-- #INCLUDE FILE="Rose_ADO.asp" -->
<!--#include file="siteinclude/HtmlFunctions.asp"-->
<!--#include file=include/AddToDb.asp-->
<!--#include file=LockFunctions.asp-->
<!--#include file=siteinclude/stringFunctions.asp-->
<%

if  session("UserId")="" then

	Response.Redirect "writeErorr.asp?iErrorId=1"

end if

If not( request("entityvalidate")="1" or request("validate")="1") Then
	CheckLock
end if

spe="#%$^"
dim sErrorStr1
dim sErrorStr2
dim sLinkShow
dim iLinkPeople
dim afterSubmit
dim iDocId
dim iSubjectId

iDocId=request("DocId")
iSubjectId=request("SubjectId")
afterSubmit=request("afterSubmit")
sLinkShow=""
sErrorStr1=""
sErrorStr2=""
iLinkPeople=request("links")
docOpener=cstr(request("docOpener"))



Set errorList = com.ErrorList()


	If request("entityvalidate")="1" or request("validate")="1" Then

		If check(errorList) Then
			'Response.Write "ok"
			'Response.Write request("area")
			'Response.Redirect "writeErorr.asp?iErrorId=4"
			'Response.Redirect "CurrentPages.asp"
		end if
	end if

%>
<html>

<head>
    <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
    <meta name="ProgId" content="FrontPage.Editor.Document">
    <meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-8 Visual">
    <title>עדכון אישיות</title>

    <script language="javascript">
    <!--
    function UploadeImage() {
        var sUploade = window.open('', sUploade,
            'toolbar=no,status=no,resizable=yes,width=300,height=400,scrollbars=yes');
        sUploade.location.href = 'UPLOADsend.asp';
    }

    function GetImagePath(sPath, sPhName) {
        eval('document.people.image_url.value =  sPath;');
        if (sPhName != "") {
            eval('document.people.fImgText.value =  sPhName;');
        }
    }

    function UpdateLink(ID, SubjectId, DocId)

    {
        window.open("UpdateLink.asp?linkId=" + ID + "&subjectId=" + SubjectId + "&DocId=" + DocId + "&Action=1", "",
            "width=600,height=300")

    }

    function setOrder(SubjectId, DocId) {
        window.open("linksByOrder.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "",
            "width=600,height=500,scrollbars=yes");
    }

    function AddALink(SubjectId, DocId) {
        var sTitle = document.people.fHead1.value
        if (sTitle == "") {
            alert("please enter a heading!")
        } else {
            window.open("AddLink.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "", "width=600,height=300")
        }
    }

    function DeletLink(ID, SubjectId, DocId) {
        window.open("UpdateLink.asp?linkId=" + ID + "&subjectId=" + SubjectId + "&DocId=" + DocId + "&Action=2", "",
            "width=600,height=300")
    }

    function GetLinks() {

        document.people.subjectId.value = <%=request("SubjectId")%>;
        document.people.docid.value = <%=request("docid")%>;
        LinksPoeple = document.people.links.value
        document.people.links.value = parseInt(LinksPoeple) + 1;
        document.people.submit()
    }

    function ChangeHidden(status) {
        document.people.RecAction.value = status;
        if (document.people.secondSubject.checked == true) {

            document.people.secondSubject.value = "1";
        }

        document.people.submit();

    }
    //////////////////////////////////////////////////////////////////////
    function ChangeHiddenDel(status) {
        if (confirm("?למחוק רשומה זו") == true) {
            document.people.RecAction.value = status;
            if (document.people.secondSubject.checked == true) {

                document.people.secondSubject.value = "1";
            }

            document.people.submit();
        }
    }
    /////////////////////////////////////////////////////////////////////
    //
    -->
    function
    LogOut()
    {
    if
    (event.clientY
    < 0) { var objHTTPGame=new ActiveXObject("Microsoft.XMLHTTP"); url="LogOutPopUp.asp?userID=<%=Session("LockedByID")%>"
        objHTTPGame.Open("get", url, false); objHTTPGame.send(); //
        ww5=window.open('LogOutPopUp.asp?userID=<%=Session("LockedByID")%>','LogOutPopUp','width=100,height=100,left=0,top=0,scrollbars=no');
        } } function AddNewLinkR() { if (document.people.TextReplace.value=='' ) { alert('נא מלא את המילה להחלפה'); }
        else { if (document.people.linkReplace.value=='' ) { alert('נא מלא את הקישור להחלפה'); } else {
        window.open('txtLinkReplace.asp?stage=addNew&docid=<%=idocid%>&subjectid=<%=isubjectid%>&txt='+escape(document.people.TextReplace.value)+'
        &link='+escape(document.people.linkReplace.value),'
        txtLinkReplace','width=450,height=450,scrollbars=yes,resizable=yes'); document.news.TextReplace.value='' ;
        document.news.linkReplace.value='' ; } } } function NewLinkR() {
        window.open('txtLinkReplace.asp?docid=<%=idocid%>&subjectid=<%=isubjectid%>','txtLinkReplace','width=450,height=450,scrollbars=yes,resizable=yes');
        } </script>

        <style>
        <!--
        body {
            font-family: Arial;
        }

        font {
            font-family: Arial, Helvetica, sans-serif;
        }
        -->
        </style>
</head>

<script language="JavaScript">
<!--
function newwinGallery() { //v2.0
    ww22 = open('XuploadPopUp.asp', 'gallerywin', 'width=650 height=400 top=0 left=0 scrollbars=yes');
}


//
-->
</script>

<body topmargin="0" onbeforeunload="LogOut()">
    <div align="center">
        <table cellspacing="0" cellpadding="0" border="0" width="751">
            <tr>
                <!--------------------------- TOP TABLES ------------------------------------>
                <td colspan="3" valign="top" valign="bottom">
                </td>
                <!--------------------------- END TOP TABLES ------------------------------------>
            </tr>
            <tr>
                <!-------------------------- LEFT TABLES --------------------------------------->
                <td bgcolor="#F5F5F5" width="142" valign="top" align="right" style="border-left: 1px solid #C0C0C0;
                    border-right: 1px solid #C0C0C0; padding-right: 4px">
                    <br>
                    <br>
                    <br>
                    <br>
                    <br>
                    <br>


                    <script>
                    function AddNewRLinkR() {
                        if (document.people.NewreferenceText.value == '') {
                            alert('נא מלא את התאור מראה מקום');
                        } else {
                            if (document.people.NewreferenceURL.value == '') {
                                alert('נא מלא את הקישור למראה מקום');
                            } else {
                                window.open(
                                    'txtLinkRefText.asp?stage=addNew&docid=<%=idocid%>&subjectid=<%=isubjectid%>&txt=' +
                                    escape(document.people.NewreferenceText.value) + '&link=' + escape(document
                                        .people.NewreferenceURL.value) + '&NewOrderNum=1', 'AddNewRLinkRFFF',
                                    'width=700,height=500,scrollbars=yes,resizable=yes');
                                document.people.TextReplace.value = '';
                                document.people.linkReplace.value = '';
                            }
                        }


                    }

                    function NewRLinkR() {
                        window.open('txtLinkRefText.asp?docid=<%=idocid%>&subjectid=<%=isubjectid%>',
                            'txtLinkReplace', 'width=700,height=500,scrollbars=yes,resizable=yes');
                    }

                    function openWindow(url, name) {
                        popupWin = window.open(url, name, 'resizable,scrollbars=yes,width=600,height=400,top=0,left=0')
                    }
                    </script>

                    <a
                        href="javascript:openWindow('search/frmAddsearch.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','searh')">
                        ימניד רושיק</a><br>
                    <br>
                    <a
                        href="javascript:openWindow('links/frmAddLink.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','links')">
                        ךרע יפל רושיק</a><br>
                    <br>
                    <a
                        href="javascript:openWindow('AddFeedback.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','rates')">
                        בושמ לוהינ</a><br>
                    <!--br>
                    <a href="javascript:openWindow('replyes/adminResponse.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','reply')">
                        םיישיא םיבתכמ</a><br-->
                    <br>
                    <a
                        href="javascript:openWindow('search_By_word.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','search_By_word')">
                        רושיק תולימ</a>
                    <br>
                    <br>
                    <a
                        href="javascript:openWindow('indexOrder.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','indexOrder',180)">
                        םוסרפ לוהינ</a>
                </td>
                <!-------------------------- END LEFT TABLES --------------------------------------->
                <!-------------------------- MAIN TABLES ------------------------------------------->
                <td width="491" valign="top" align="center">
                    <form action="updatePeople.asp?subjectId=<%=request("SubjectId")%>" method="post" name="people">
                        <input type="hidden" name="validate" value="1">
                        <input type="hidden" name="subjectId" value="<%=request("subjectId")%>">
                        <input type="hidden" name="docid" value="<%=request("docid")%>">
                        <input type="hidden" name="status" value="<%=request("status")%>">
                        <input type="hidden" name="afterSubmit" value="<%=afterSubmit%>">
                        <input type="hidden" name="RecAction" value="<%=request("RecAction")%>">
                        <input type="hidden" name="parentdocid" value="<%=request("parentdocid")%>">
                        <input type="hidden" name="docOpener" value="<%=request("docOpener")%>">
                        <input type="hidden" name="strSearch" value="<%=request("strSearch")%>">
                        <input type="hidden" name="txtSearch" value="<%=request("txtSearch")%>">
                        <input type="hidden" name="opSearch" value="<%=request("opSearch")%>">
                        <input type="hidden" name="CurPage" value="<%=request("CurPage")%>">
<%
if request("links")="" then
                      %>
<input type="hidden" name="links" value="0">
<%else%>
<input type="hidden" name="links" value="<%=request("links")%>">
<%end if%>
<table border="0" cellspacing="0" cellpadding="0">
<%

if request("subjectId")=6 then

strTitle="<h2><font color='navy' face='Arial'>תוישיא ןוכדע</font></h2>"
else
strTitle="<h2><font color='navy' face='Arial'>החמומ ןוכדע</font></h2>"
end if
                            %>
<tr>
                                <th colspan="2" height="15" align="center">
                                    <br>
<%=strTitle%>
<br>
                                    <br>
                                </th>
                            </tr>
                            <tr align="right">
                                <td align="right" nowrap>
                                    <font size="3" color="red">הבוח תודש םניה * ב םינמוסמה תודשה</font>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="right">
                                    <font color="red" size="2">
<%set objDButils=server.CreateObject("yoav.Com_logic")


			strSqlDoc="select * from T_people where docId=" & iDocId & " and subjectId= " & iSubjectId & " and status=" & request("status")

			set subjectrs=objDButils.GetFastRecoredset(strSqlDoc)
			if subjectrs.bof then
				Response.Write "No Documents found."

			else

			if subjectrs("writername")<>session("UpdatedByUserName") and session("userid")=3 then
				response.redirect "currentpages.asp?do=noPermissions"
			end if

				while not subjectrs.eof
				iUserId=subjectrs("userId")
				sHead1=subjectrs("head1")
				sHead2=subjectrs("head2")
				sDuty=subjectrs("duty")
				sInstitute=subjectrs("institute")
				sImage=subjectrs("image")
				sImageText=subjectrs("imgText")

				sText=subjectrs("Mtext")
                'sText=findReplace(subjectrs("Mtext"))
				ssubject=subjectrs("secondsubjectid")
				ShowDetails=subjectrs("ShowDetails")

				 personalDetails=subjectrs("personalDetails")
				 generalInfo=subjectrs("generalInfo")
				 educationDetails=subjectrs("educationDetails")
				 jobDetails=subjectrs("jobDetails")
				 experience=subjectrs("experience")
				 memberships=subjectrs("memberships")
				 publications=subjectrs("publications")
				 customCaption1=subjectrs("customCaption1")
				 customText1=subjectrs("customText1")
				 customCaption2=subjectrs("customCaption2")
				 customText2=subjectrs("customText2")
				subjectrs.movenext
				wend
				subjectrs.close
			end if
		if sErrorStr1 <> "" or sErrorStr2 <> "" then
		Response.Write translatelogical2visual(sErrorStr1,60) & "&nbsp;<img src='..\img\bull.gif'><br>"
		Response.Write translatelogical2visual(sErrorStr2,20) & "&nbsp;<img src='..\img\bull.gif'><br>"
		end if
                                        %>
                                    </font>
                                </td>
                            </tr>
                            <tr align="right">
                                <%sHead1 = replace(sHead1,"""","&quot;")%>
                                <td>
                                    <div dir="rtl">
                                        <input onblur="document.people.hebrewvalue.value=document.people.fHead1.value" type="text"
                                            name="fHead1" size="80" value="<%=sHead1%>" tabindex="1"></div>
                                </td>
                                <td valign="top" nowrap>
                                    <font size="2">:(תירבעב) םש</font><font color="red">*</font></td>
                            </tr>
                            <tr align="right">
                                <%sHead2 = replace(sHead2,"""","&quot;")%>
                                <td>
                                    <div dir="LTR">
                                        <input type="text" name="fHead2" size="80" value="<%=sHead2%>" tabindex="2"></div>
                                </td>
                                <td valign="top">
                                    <font size="2">:(תילגנאב) םש</font></td>
                            </tr>
                            <tr align="right">
                                <%sDuty = replace(sDuty,"""","&quot;")%>
                                <td>
                                    <div dir="rtl">
                                        <input type="text" name="fDuty" size="80" value="<%=sDuty%>" tabindex="3"></div>
                                </td>
                                <td valign="top">
                                    <font size="2">:קוסיע/דיקפת</font></td>
                            </tr>
                            <tr align="right">
                                <%sInstitute = replace(sInstitute,"""","&quot;")%>
                                <td valign="top" dir="RTL">
                                    <input type="text" name="fInstitute" value="<%=sInstitute%>" size="80" tabindex="4"></td>
                                <td valign="top">
                                    <font size="2">:דסומ/המריפ</font></td>
                            </tr>
                            <tr align="right">
                                <%sImage = replace(sImage,"""","&quot;")%>
                                <td>
                                    <a onclick='newwinGallery();' href="javascript:void(0);">
                                        <img border="0" src="EditorIcons/icon_ins_image.gif"></a><input type="text" name="image_url"
                                            value="<%=sImage%>" size="30" tabindex="5"></td>
                                <td>
                                    <font size="2">:הנומת</font></td>
                            </tr>
                            <tr align="right">
                                <td colspan="2">
                                    <input type="button" name="InsImage" value="הכנס תמונה" size="35" onclick="UploadeImage();"
                                        tabindex="6">
                                </td>
                            </tr>
                            <tr align="right">
                                <%fDuty = replace(request("fDuty"),"""","&quot;")%>
                                <td dir="RTL">
                                    <textarea rows="2" cols="50" name="fImgText" tabindex="7"><%=sImageText%></textarea></td>
                                <td>
                                    <font size="2">:הנומת בותיכ</font></td>
                            </tr>
                            <tr>
                                <td align="right" valign="top">
                                    <input type="text" name="fDate" value="<%=date()%>" size="10" tabindex="8"></td>
                                <td align="RIGHT">
                                    <font size="2">:ךיראת</font></td>
                            </tr>
                            <tr align="right">
                                <td colspan="2" valign="top">
                                    <br>
                                    <font size="2">:יוצרה טסקטה תא בותכל אנ</font></td>
                            </tr>
                            <tr align="right">
                                <td colspan="2" valign="top">
                                    <div dir="RTL">
                                        <br>
                                        <textarea rows="20" cols="60" name="fText" tabindex="9"><%=findReplace(sText)%></textarea></div>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">םיישיא םיטרפ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="PersonalDetails" cols="63" dir="rtl"><%=PersonalDetails%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">יללכ ישיא עדימ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="generalInfo" cols="63" dir="rtl"><%=generalInfo%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">הלכשה יטרפ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="educationDetails" cols="63" dir="rtl"><%=educationDetails%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">הקוסעת/הדובע יטרפ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="jobDetails" cols="63" dir="rtl"><%=jobDetails%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">הדובעב ןויסינ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="experience" cols="63" dir="rtl"><%=experience%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">תודסומ/םידוגיאב תירבח</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="memberships" cols="63" dir="rtl"><%=memberships%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">םייעוצקמ םימוסריפ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="publications" cols="63" dir="rtl"><%=publications%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">&nbsp;1 תרתוכ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <input class="inputText" type="text" name="customCaption1" size="80" value="<%=(customCaption1)%>"></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">1 טסקט</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="customText1" cols="63" dir="rtl"><%=(customText1)%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">&nbsp; 2 תרתוכ</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <input class="inputText" type="text" name="customCaption2" size="80" value="<%=(customCaption2)%>"></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <font size="2">2 טסקט</font></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    <textarea rows="5" name="customText2" cols="63" dir="rtl"><%=(customText2)%></textarea></td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top" align="right">
                                    &nbsp;</td>
                            </tr>
                             <tr align="right">
                            <td colspan="2" valign="top" dir="rtl" align="right">
                                <b><font size="2" color="#000080">מראה מקום</font></b></td>
                        </tr>
                        <tr align="right">
                            <td colspan="2" valign="top">
                                <table border="0" id="table5" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td align="right">
                                            <input type="button" value="עדכן" onclick="javascript:NewRLinkR();" style="font-size: 10px;
                                                font-family: Arial;">&nbsp;&nbsp;
                                        </td>
                                        <td align="right">
                                            <input type="button" value="הוסף חדש" onclick="javascript:AddNewRLinkR();" style="font-size: 10px;
                                                font-family: Arial;">&nbsp;&nbsp;
                                        </td>
                                        <td align="right">
                                            <font size="2">
                                                <input type="text" name="NewreferenceURL" size="20" dir="ltr" value="" style="font-size: 8pt;
                                                    font-family: Arial"></font></td>
                                        <td width="10%" align="right" dir="rtl">
                                            <font size="2">&nbsp;קישור:</font></td>
                                        <td align="right">
                                            <font size="2">
                                                <input type="text" name="NewreferenceText" size="30" value="" dir="rtl" style="font-size: 8pt;
                                                    font-family: Arial"></font></td>
                                        <td width="10%" align="right" dir="rtl">
                                            <font size="2">&nbsp;תאור:</font></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>

                             <tr align="right">
                            <td colspan="2" valign="top" dir="rtl" align="right">
                                <b><font size="2" color="#000080">החלפת מילה בקישור</font></b></td>
                        </tr>
                        <tr align="right">
                            <td colspan="2" valign="top">
                                <table border="0" id="table1" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <td align="right">
                                            <input type="button" value="עדכן" onclick="javascript:NewLinkR();" style="font-size: 10px;
                                                font-family: Arial;">&nbsp;&nbsp;
                                        </td>
                                        <td align="right">
                                            <input type="button" value="הוסף חדש" onclick="javascript:AddNewLinkR();" style="font-size: 10px;
                                                font-family: Arial;">&nbsp;&nbsp;
                                        </td>
                                        <td align="right">
                                            <font size="2">
                                                <input type="text" name="linkReplace" size="30" dir="ltr" value="" style="font-size: 8pt;
                                                    font-family: Arial"></font></td>
                                        <td width="10%" align="right" dir="rtl">
                                            <font size="2">&nbsp;קישור:</font></td>
                                        <td align="right">
                                            <font size="2">
                                                <input type="text" name="TextReplace" size="20" value="" dir="rtl" style="font-size: 8pt;
                                                    font-family: Arial"></font></td>
                                        <td width="10%" align="right" dir="rtl">
                                            <font size="2">&nbsp;מילה:</font></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>

                            <tr align="center">
                                <td colspan="2" align="center">
                                    <%
	dim sSql

	 sSql=" SELECT * FROM T_links WHERE " & _
	      " docId=" & iDocId & " AND " & _
	      " subjectId="& iSubjectId & _
	      " ORDER BY orderByNum,linkName "

	 set objDButils=server.CreateObject("yoav.com_logic")

	 set rs=objDButils.GetFastRecoredset(sSql)
		LinkCount=1
		sLinkShow="<tr><td colspan=2 align=right><br><font size=""-1"" color=""#0051BD""><u>:הכ דע ופסונש םירושיקה ןלהל</font></u></td></tr>"

		while not rs.eof
		sLinkShow=sLinkShow &  "<tr><td colspan = '2' align=right><font size=""-1"" color=""#0051BD"">"
		sLinkShow=sLinkShow &  "<a href='javascript:UpdateLink("& rs("id")&","& rs("subjectId")&","& rs("docId") &")'>"
		sLinkShow=sLinkShow &  "ןכדע</a>"
		sLinkShow=sLinkShow &  "&nbsp;<a href='javascript:DeletLink("& rs("id")&","& rs("subjectId")&","& rs("docId") &")'>"
		sLinkShow=sLinkShow &  "קחמ </a></font>"
		sLinkShow=sLinkShow &  "<font size=""-1"" color=""#0051BD"">"
		sLinkShow=sLinkShow &  translatelogical2visual(rs("linkName"),50)
		sLinkShow=sLinkShow &  "</font>&nbsp;<font color=black size=""-1"">"& LinkCount &"</font></td></tr>"
		LinkCount=LinkCount+1
		rs.movenext
		wend


                                    %>
                                    <%
	Response.Write  sLinkShow

                                    %>
                                </td>
                            </tr>
                            <tr align="right">
                                <td colspan="2" valign="top" align="center">
                                    <br>
                                    <br>
                                    <a href="javascript:setOrder(<%=isubjectId%>,<%=iDocId%>)" tabindex="10">רדס עבק</a>&nbsp;&nbsp;
                                    <a href="javascript:AddALink(<%=isubjectId%>,<%=iDocId%>)" tabindex="10">רושיק ףסוה</a><br>
                                </td>
                            </tr>
                        </table>
                        <input type="hidden" name="entityvalidate" value="">
                        <table border="0" cellspacing="2" cellpadding="2" width="480">
                            <%


	a = errorList.Items

	 For fcount = 0 To errorList.Count -1


	 response.Write "<tr><td align='right' colspan=4><FONT COLOR='red' size='-1'><STRONG>"
      Select Case a(fcount)

	  Case "fname"

			Response.Write translatelogical2visual("נא הכנס שם פרטי ",20)& "&nbsp;<img src='..\img\bull.gif'>"
	  Case "lname"

			Response.Write translatelogical2visual("נא הכנס שם משפחה",20)& "&nbsp;<img src='..\img\bull.gif'>"
	  Case "Address"

			Response.Write translatelogical2visual("נא הכנס כתובת",20)& "&nbsp;<img src='..\img\bull.gif'>"
     Case "phone"

			Response.Write translatelogical2visual("נא הכנס מספר טלפון",20)& "&nbsp;<img src='..\img\bull.gif'>"
	  Case "Email"

   			Response.Write translatelogical2visual("נא הכנס דואר אלקטרוני",30)& "&nbsp;<img src='..\img\bull.gif'>"

		case "hebrewvalue"
			Response.Write translatelogical2visual("נא הכנס ערך בעברית",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "englishvalue"
			Response.Write translatelogical2visual("נא הכנס ערך באנגלית",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "area"
			Response.Write translatelogical2visual("נא בחר תחום",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "zipcode"
			Response.Write translatelogical2visual("נא הכנס מיקוד נכון",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "classcification"
			Response.Write translatelogical2visual("נא בחר סיווג",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "duty"
			Response.Write translatelogical2visual("נא בחר תפקיד",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "settling"
			Response.Write translatelogical2visual("נא הכנס יישוב",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "birthyear"
			Response.Write translatelogical2visual("נא הכנס שנת לידה",20)& "&nbsp;<img src='..\img\bull.gif'>"
      end select
      response.Write "</STRONG></font></td></tr>"
    next

                            %>
                            <%


		strSqlDoc="select * from T_entity where doc_Id=" & iDocId & " and subjectId= " & iSubjectId

			set NTTrs=objDButils.GetFastRecoredset(strSqlDoc)
			if NTTrs.bof then
				Response.Write "No Documents found."

			else

				while not NTTrs.eof
				sHebValue=NTTrs("hebrew_dscr")
				sEngVlaue=NTTrs("english_dscr")
				sFname=NTTrs("first_name")
				sLname=NTTrs("last_name")
				sAddress=NTTrs("address")
				iZipCode=NTTrs("zipcode")
				sPhone=NTTrs("phone")
				sSettling=NTTrs("settling")
				sFax=NTTrs("fax")
				sEmail=NTTrs("email")
				iArea=NTTrs("area1_id")
				Affair_id=NTTrs("Affair_id")&""
				Lang=NTTrs("Lang")&""
				skill_id=NTTrs("skill_id")&""
				skill_Parent_id=NTTrs("skill_Parent_id")&""
				iClasscification=NTTrs("classcification")
				category_id=NTTrs("category_id")&""
				Province_id=NTTrs("Province_id")&""

				iDuty=NTTrs("duty_id")
				sFirm=NTTrs("firm_name")
				iYear=NTTrs("year")
				sBirthday=NTTrs("birthday")
				iSex=NTTrs("sex")
				iNation1=NTTrs("nation1")
				iNation2=NTTrs("nation2")
				site_url=NTTrs("site_url")
				sSalary=NTTrs("person_total_salary")
				iSalaryDate=trim(NTTrs("person_salary_date"))
				education=NTTrs("education")
				iplaceToBeBorn=NTTrs("placeToBeBorn")
				idecease_date=NTTrs("decease_date")
				idate=NTTrs("entry_date")
			    iClasscification2=NTTrs("new_classcification")
			    iduty2=NTTrs("new_duty")
                OnlyMembers = NTTrs("OnlyMembers")
				NTTrs.movenext

				wend
			end if


            strSqlDoc="select * from T_MailingList where email='" & sEmail & "'"
			set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)
                            %>
                            <tr align="right">
                                <%sHebValue = replace(sHebValue,"""","&quot;")%>
                                <td dir="RTL" colspan="3">
                                    <textarea rows="6" name="hebrewvalue" cols="40" tabindex="11"><%=sHebValue%></textarea>
                                    <!--INPUT type="text"  name="hebrewvalue" value="<%=sHebValue%>" size="65" tabindex=11-->
                                </td>
                                <td>
                                    <font size="2">:ךרע םש</font><font size="2" color="red">*</font></td>
                            </tr>
                            <tr align="right">
                                <%sEngVlaue = replace(sEngVlaue,"""","&quot;")%>
                                <td dir="LTR" colspan="3">
                                    <input type="text" name="englishvalue" value="<%=sEngVlaue%>" size="65" tabindex="12"></td>
                                <td nowrap>
                                    <font size="2">:(תילגנאב) ךרע םש</font></td>
                            </tr>
                            <tr align="right">
                                <%sFname = replace(sFname,"""","&quot;")%>
                                <%sLname = replace(sLname,"""","&quot;")%>
                                <td dir="LTR">
                                    <input type="fname" name="fname" value="<%=sFname%>" size="15" tabindex="14"></td>
                                <td>
                                    <font size="2">:יטרפ םש</font><font size="2" color="red">*</font></td>
                                <td dir="LTR">
                                    <input type="lname" name="lname" value="<%=sLname%>" size="25" tabindex="13"></td>
                                <td nowrap>
                                    <font size="2">:החפשמ םש</font><font size="2" color="red">*</font></td>
                            </tr>
                            <tr align="right">
                                <%if not request("placetobeborn")="" then%>
                                <td dir="RTL">
                                    <%
		call ShowCountryCombo(request("placetobeborn"),"placeToBeBorn","tabindex=16")%>
                                </td>
                                <td>
                                    <font size="2">:הדיל ץרא</font></td>
                            </tr>
                            <caption>
                    &nbsp;</td>
                <%else%>
                </caption>
                <tr>
                    <td dir="RTL">
                        <%if not isnull(iplaceToBeBorn)then

		call ShowCountryCombo(iplaceToBeBorn,"placeToBeBorn","tabindex=16")
		else
		call ShowCountryCombo("בחר","placeToBeBorn","tabindex=16")
		end if
                        %>
                    </td>
                    <td>
                        <font size="2">:הדיל ץרא</font></td>
                </tr>
                <caption>
                    &nbsp;</td>
                    <%end if%>
                </caption>
                <tr>
                    <td dir="RTL" valign="top">
                        <% fsex=request("fsex")
	if request("fsex")="" and not iSex="" then

	 	select case cint(iSex)

		case 1
		  sSelected1="selected"
		  sSelected2=""
		case 2
		  sSelected2="selected"
		  sSelected1=""
		end select
                        %>
                        <select name="fsex" tabindex="15">
                            <option value="0">בחר</option>
                            <option value="1" <%=sSelected1%>>זכר</option>
                            <option value="2" <%=sSelected2%>>נקבה</option>
                        </select>
                        <%
	else
	    select case fsex
	    case 1
		  sSelected1="selected"
		  sSelected2=""
		case 2
		  sSelected2="selected"
		  sSelected1=""
		end select

                        %>
                        <select name="fsex" tabindex="15">
                            <option value="0">בחר</option>
                            <option value="1" <%=sSelected1%>>זכר</option>
                            <option value="2" <%=sSelected2%>>נקבה</option>
                        </select>
                        <%
	end if
                        %>
                    </td>
                    <td valign="top" align="right">
                        <font size="2">:ןימ</font></td>
                </tr>
                <tr>
                    <%sBirthdayCheck=1
                    if instr(1, sBirthday & "", "/")>0 then
                    sBirthdayArr = split(sBirthday & "", "/")
                    else
                    sBirthdayCheck=0
                    end if
                    if  trim(sBirthday & "") = "" or not isdate(sBirthday) then sBirthday="0/0/0"
		 If sBirthday="0/0/0" then
'==============================================================================	%>
<td dir="RTL" nowrap>
                                    <select name="day" tabindex="20">
                                        <option selected value="">יום</option>
<% if request("day")="" then
if sBirthdayCheck=0 then
else%>
<option value="<%=sBirthdayArr(0)%>" selected>
<%=sBirthdayArr(0)%>
</option>
                                        <% end if
                                        %>

                                        <%
                                        For iDayCount=1 to 31
                                            Response.Write"<OPTION value="& iDayCount &">"& iDayCount &"</OPTION>"
                                           Next

                                         else
                                                          %>
                                        <option value="<%=request("day")%>" selected>
<%=request("day")%>
</option>
                                        <%
                                        For iDayCount=1 to 31
                                            Response.Write"<OPTION value="& iDayCount &">"& iDayCount &"</OPTION>"
                                           Next%>
                                        <%end if%>
                                    </select>
<% '==============================================================================


                       %>
                       <select name="month" tabindex="19">
                           <option selected value="">חודש</option>
                           <%if request("month")="" then
                           if sBirthdayCheck=0 then
                           else%>
                           <option value="<%=sBirthdayArr(1)%>" selected>
                               <%=sBirthdayArr(1)%>
                           </option>
                          <% end if%>
                           <%
for iMonCount=1 to 12
Response.Write"<OPTION value="& iMonCount &">"& iMonCount &"</OPTION>"
next
else%>
                           <option value="<%=request("month")%>" selected>
                               <%=request("month")%>
                           </option>
                           <%
for iMonCount=1 to 12
Response.Write"<OPTION value="& iMonCount &">"& iMonCount &"</OPTION>"
next
end if
                           %>
                       </select>
                       <%'==========================================================================================%>
<select name="year" tabindex="18">
                                        <option value="" selected>שנה</option>
<%if request("year")="" then
if sBirthdayCheck=0 then
else%>
<option value="<%=sBirthdayArr(2)%>" selected>
<%=sBirthdayArr(2)%>
</option>
                                        <%end if%>
                                        <%
                                        for iYearCount=1900 to year(now)
                                        Response.Write"<OPTION value="& iYearCount&">"& iYearCount &"</OPTION>"
                                        next
                                                                       %>
                                        <option value="1800" selected>ללא תאריך</option>
<%else%>
<option value="<%=request("year")%>" selected>
<%=request("year")%>
</option>
<%
for iYearCount=1900 to year(now)
Response.Write"<OPTION value="& iYearCount&">"& iYearCount &"</OPTION>"
next

end if%>
</select>
                                </td>
<%
'=============================================================================================
'=============================================================================================
'=============================================================================================



	Else



	today=Cdate(sBirthday)
                    %>
                    <td dir="RTL" nowrap>
                        <select name="day" tabindex="20">
                            <option value="" selected>יום</option>
                            <option value="<%=day(today)%>" selected>
                                <%=day(today)%>
                            </option>
                            <%
	         For iDayCount=1 to 31
	             Response.Write"<OPTION value="& iDayCount &">"& iDayCount &"</OPTION>"
             Next

                            %>
                        </select>
                        <% '==============================================================================


                        %>
<select name="month" tabindex="19">
                                    <option value="" selected>חודש</option>
<%%>
<option value="<%=month(today)%>" selected>
<%=month(today)%>
</option>
<%
for iMonCount=1 to 12
Response.Write"<OPTION value="& iMonCount &">"& iMonCount &"</OPTION>"
next

                           %>
</select>
<%'==========================================================================================%>
                        <select name="year" tabindex="18">
                            <option value="" selected>שנה</option>

                            <%%>
                            <option value="<%=year(today)%>" selected>
                            <%if(year(today)<>1800) then %>    <%=year(today)%>

                                <%else %>
                               ללא תאריך
                                <%end if %>

                            </option>
                            <%
	for iYearCount=1930 to year(now)
	Response.Write"<OPTION value="& iYearCount&">"& iYearCount &"</OPTION>"
	next

                            %>
                               <option value="1800" >ללא תאריך</option>
                        </select>
                    </td>
                    <%

	End If
'================================================================================================
'================================================================================================
'================================================================================================
                    %>
<td align="right">
                                    <font size="2">:הדיל ךיראת</font>
                                </td>
                                <td dir="RTL" valign="top">
                                    <select name="fBirthYear" tabindex="17">
<%
IF iYear=1 and request("fBirthYear")="" then

      sSelected="selected"

Response.Write "<OPTION value='1' selected>-</OPTION>"

for yearCount=year(now) to  1810 step -1

Response.Write "<OPTION value='"& yearCount &"'>"& yearCount &"</OPTION>"

next


ELSE

if request("fBirthYear")="" then  'cint(iYear)
sSelected="selected"
Response.Write "<OPTION value='"& iYear &"'"& sSelected &" >"& iYear &"</OPTION>"
Response.Write "<OPTION value=1>-</OPTION>"
for yearCount=year(now) to  1810 step -1

Response.Write "<OPTION value='"& yearCount &"'>"& yearCount &"</OPTION>"
next
else
sSelected="selected"
Response.Write "<OPTION value='"& request("fBirthYear") &"'"& sSelected &" >"& request("fBirthYear") &"</OPTION>"
Response.Write "<OPTION value=1>-</OPTION>"
for yearCount=year(now) to 1810 step -1

Response.Write "<OPTION value='"& yearCount &"'>"& yearCount &"</OPTION>"
next
end if

 END IF
                          %>
                          <option value="-1">-</option>
                      </select>
                  </td>
                  <td valign="top" align="right">
                      <font size="2">:הדיל תנש</font></td>
              </tr>
              <%'=====================+++++++++++++++++++++++++++++==============================================%>
<tr align="right">
                                <td dir="RTL">
                                    <font face="Arial (hebrew)">
<%if request("fEducation")="" then
   call ShowSpecialCombo(education,"T_education","fEducation","tabindex='22'")
else
   call ShowSpecialCombo(request("fEducation"),"T_education","fEducation","tabindex='22'")
end if
                        %>
</font>
                                </td>
                                <td>
                                    <font size="2">:הלכשה</font>
                                </td>
                                <td dir="RTL">
                                    <font face="Arial (hebrew)">
<%if request("fNation1")="" then
   call ShowSpecialCombo(iNation1,"T_nation","fNation1","tabindex='21'")
else
   call ShowSpecialCombo(request("fNation1"),"T_nation","fNation1","tabindex='21'")
end if
                        %>
</font>
                                </td>
                                <td nowrap>
                                    <font size="2">: תוחרזא</font>
                                </td>
                            </tr>
                            <tr align="right">
                                <%decease_date = replace(decease_date,"""","&quot;")%>
                                <%if request(decease_date)="" then%>
                                <td dir="RTL">
                                    <input type="text" name="decease_date" value="<%=idecease_date%>" tabindex="24"
                                        size="20">
                                </td>
                                <td>
                                    <font size="2">:<span lang="he">יתחפשמ בצמ</span></font>
                                </td>
<%else%>
<td dir="RTL">
                                    <input type="text" name="decease_date" value="<%=request("decease_date")%>" tabindex="24"
                                        size="20">
                                </td>
                                <td>
                                    <font size="2">:<span lang="he">יתחפשמ בצמ</span></font>
                                </td>
<%
end if%>
<td dir="LTR" valign="top">
                                    <font face="Arial (hebrew)">
                                        <%if request("fNation2")="" then
                                        		     call ShowSpecialCombo(iNation2,"T_nation","fNation2","tabindex='23'")
                                        		  else
                                        		     call ShowSpecialCombo(request("fNation2"),"T_nation","fNation2","tabindex='23'")
                                        		  end if
                                        		'call ShowSpecialCombo(iNation2,"T_nation","fNation2","")%>
                                                                </font>
                                                            </td>
                                                            <td valign="top" nowrap>
                                                                <font size="2">:תפסונ תוחרזא</font></td>
                                                        </tr>
                                                        <tr align="right">
                                                            <td dir="LTR">
                                                                <font face="Arial (hebrew)">
                                                                    <%'if request("classcification")="" then
                                        		   '  call ShowSpecialCombo(iClasscification,"T_classification","classcification","tabindex='24'")
                                        	      'else
                                        	       '  call ShowSpecialCombo(request("classcification"),"T_classification","classcification","tabindex='24'")
                                        	      'end if
                                        	    '===========New classcification================================================
                                        	    if cstr(request("classcification2"))="" and isnull(iclasscification2)then

                                        	         call  ShowClassCombo("","classcification2","size=3 MULTIPLE tabindex=26")

                                        	      else

                                        	       if cstr(request("classcification2"))="" then

                                        		      'call ShowAreaCombo (iArea,"area","size=3 MULTIPLE")
                                        		      'Response.Write iArea
                                        		      showAreaid=split(iclasscification2,",")
                                        			  call ShowClassUpdate(showAreaid,26)

                                        		   else

                                        		     'call ShowAreaCombo (request("area"),"area","size=3 MULTIPLE")
                                        		     'test=cstr(iarea))

                                        		     iArea=cstr(request("classcification2"))
                                        			 showAreaid=split(iArea,",")

                                        			 call ShowClassUpdate(showAreaid,26)

                                        		   end if
                                        		 end if
                                                                    %>
                                                                </font>
                                                            </td>
                                                            <td align="right">
                                                                <font size="2">:עוצקמ/ראות</font></td>
                                                            <td dir="LTR">
                                                                <font face="Arial (hebrew)">
                                                                    <%'if request("duty")="" then
                                        		    ' call ShowSpecialCombo(iDuty,"T_duty","duty","tabindex='23'")
                                        		  'else
                                        		   '  call ShowSpecialCombo(request("duty"),"T_duty","duty","tabindex='23'")
                                        		  'end if
                                        		 '==============New duty====================================================
                                        		  If 	cstr(request("duty2"))="" and isnull(iDuty2) then

                                        	        call  ShowDutyCombo("","duty2","size=3 MULTIPLE tabindex=25")
                                        		  Else

                                        		  if cstr(request("duty2"))="" then

                                        		     'call ShowAreaCombo (iArea,"area","size=3 MULTIPLE")
                                        		     'Response.Write iArea
                                        		     showAreaid=split(iDuty2,",")
                                        			 call ShowDutyUpdate(showAreaid,25)

                                        		  else

                                        		     'call ShowAreaCombo (request("area"),"area","size=3 MULTIPLE")
                                        		     'test=cstr(iarea))

                                        		    iArea=cstr(request("duty2"))
                                        			showAreaid=split(iarea,",")

                                        			call ShowDutyUpdate(showAreaid,25)

                                        		  end if
                                        		End if
                                                                    %>
                                                                </font>
                                                            </td>
                                                            <td align="right">
                                                                <font size="2">:קוסיע/דיקפת</font></td>
                                                        </tr>
                                                        <tr align="right">
                                                            <td dir="LTR" valign="top">
                                                                <font face="Arial (hebrew)">
                                                                    <%
                                        		if request(area)="" and isempty(iarea) then
                                        		call  ShowAreaCombo("","area","size=3 MULTIPLE tabindex=28")
                                        		else
                                        		if request("area")="" then

                                        		     'call ShowAreaCombo (iArea,"area","size=3 MULTIPLE")
                                        		     'Response.Write iArea
                                        		     showAreaid=split(iArea,",")
                                        			 call ShowAreaupdate(showAreaid,28)

                                        		  else

                                        		     'call ShowAreaCombo (request("area"),"area","size=3 MULTIPLE")
                                        		     'test=cstr(iarea))

                                        		    iArea=request("area")
                                        			showAreaid=split(iArea,",")

                                        			call ShowAreaupdate(showAreaid,28)

                                        		  end if
                                        		end if
                                                                    %>
                                                                </font>
                                                            </td>
                                                            <td valign="top">
                                                                <font size="2">:םוחת</font></td>
                                                            <%sFirm = replace(sFirm,"""","&quot;")%>
                                                            <td dir="RTL">
                                                                <input type="text" name="firm_name" value="<%=sFirm%>" size="25" tabindex="27"></td>
                                                            <td>
                                                                <font size="2">:דסומ/המריפ</font></td>
                                                        </tr>
                                                        <tr align="right">
                                                            <td dir="LTR" valign="top">
                                                                <%
                                        			if cstr(request("Province_id"))="" then
                                        		    	showProvinceID=split(Province_id,",")
                                        		  	else
                                        		    	Province_id=cstr(request("Province_id"))
                                        				showProvinceID=split(Province_id,",")
                                        		  	end if
                                        			call ShowProvinceUpdate(showProvinceID,18)


                                                                %>
                                                            </td>
                                                            <td valign="top">
                                                                <font size="2">:זוחמ</font></td>
                                                            <td dir="LTR">
                                                                <%
                                        			if cstr(request("category_id"))="" then
                                        		    	showcategoryID=split(category_id,",")
                                        		  	else
                                        		    	category_id=cstr(request("category_id"))
                                        				showcategoryID=split(category_id,",")
                                        		  	end if
                                        			call ShowCategoryUpdate(showcategoryID,18)


                                                                %>
                                                            </td>
                                                            <td align="right">
                                                                <font size="2">:הירוגטק</font></td>
                                                        </tr>
                                                        <tr align="right">
                                                            <td dir="LTR" valign="top">
                                                                <%
                                        			if cstr(request("Affair_id"))="" then
                                        		    	showAffairID=split(Affair_id,",")
                                        		  	else
                                        		    	Affair_id=cstr(request("Affair_id"))
                                        				showAffairID=split(Affair_id,",")
                                        		  	end if
                                        			call ShowAffairUpdate(showAffairID,18)


                                                                %>
                                                            </td>
                                                            <td valign="top">
                                                                <font size="2">:השרפ</font></td>
                                                            <td dir="LTR">
                                                                <%
                                        			if cstr(request("LAng"))="" then
                                        		    	showLang=split(Lang,",")
                                        		  	else
                                        		    	Lang=cstr(request("Lang"))
                                        				showLang=split(Lang,",")
                                        		  	end if
                                        			call ShowLangUpdate(showLang,18)
                                                                %>
                                                            </td>
                                                            <td>
                                                                <font size="2">:הפש</font></td>
                                                        </tr>
                                                        <%
                                        	' קטגוריות תחום
                                        set Hiconn = Server.CreateObject("ADODB.CONNECTION")
                                        Hiconn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;" 'strconn
                                        set librs=server.createobject("adodb.recordset")


                                                        %>

                                                        <script language="JavaScript">
                                        function manuselected1(elem){
                                        for (var i = document.people.Skill_ID.options.length; i >= 0; i--){
                                        document.people.Skill_ID.options[i] = null;
                                        }
                                        <%
                                        set rsBig=Hiconn.execute("select * from T_AreaCategories where parentID=0 order by catname")
                                        do until rsBig.EOF
                                        %>
                                        if (elem.options[elem.selectedIndex].value==<%=rsBig("catID")%>){
                                        			document.people.Skill_ID.options[document.people.Skill_ID.options.length] = new Option('בחר','0');
                                        		<%
                                        			set rsChiled=Hiconn.execute("select * from T_AreaCategories where parentID="&rsBig("catID")&" order by catname")
                                        			do until rsChiled.EOF
                                        			%>
                                        				document.people.Skill_ID.options[document.people.Skill_ID.options.length] = new Option('<%=rsChiled("catNaME")%>','<%=rsChiled("catNaME")%>');
                                        <%
                                        	rsChiled.moveNext
                                        	loop
                                        %>
                                        }
                                        <%
                                        rsBig.MoveNext
                                        loop
                                        %>
                                        }
                                        </script>

                            <tr align="right">
                                <td>
                                    <select style="width: 120" size="3" name="Skill_ID" dir="rtl" multiple>
<%

  if not isnumeric(skill_Parent_id) or skill_Parent_id & "" = "" then
        skill_Parent_id = 0
  end if
  if request("Skill_Parent_ID") & "" <> "" OR skill_Parent_id > 0  then

  	strtmpSK=skill_Parent_id

  	if request("Skill_Parent_ID")<>"" Then strtmpSK=request("Skill_Parent_ID")


     libsql="select * from T_AreaCategories where parentId="& strtmpSK &" order by catname"
     librs.open libsql,Hiconn

     do until librs.eof
%>
<option value="<%=Replace(librs("catname")&"",chr(34),"&quot;")%>" <%
								      if Request("Skill_ID")<>"" Then
								      	arrSkills=Split(Request("Skill_ID"),",")
								      Else
								      	arrSkills=Split(Skill_ID,",")
								      End if
								      for skIndex=0 to Ubound(arrSkills)
								      	if Cstr(Trim(librs("catname")))=Cstr(Trim(arrSkills(skIndex))) then response.write "  selected "
								      Next
								      %>>
<%=librs("catname")%>
</option>
<%
     librs.movenext

    loop
   librs.close
  end if
%>
</select>
                                </td>
                                <td dir="LTR">
                                    &nbsp;</td>
                                <td dir="LTR">
<%
             libsql="select * from T_AreaCategories where parentId=0 order by catname"
             librs.open libsql,Hiconn
%>
<select size="3" name="Skill_Parent_ID" dir="rtl" style="width: 120"
                                        onchange="manuselected1(this);">
<%
           do until librs.eof
%>
<option value="<%=librs("catid")%>" <%if ( Cstr(request("Skill_Parent_ID"))=Cstr(librs("catid")) OR Cstr(skill_Parent_id)=Cstr(librs("catid"))  ) then response.write "  selected"%>>
<%=librs("catname")%>
</option>
<%
        librs.movenext
           loop
        librs.close
%>
</select>
                                </td>
                                <td>
                                    &nbsp;</td>
                            </tr>
                            <tr align="right">
                                <td dir="LTR">
                                    <input type="text" name="fax" value="<%=sFax%>" size="15" tabindex="30">
                                </td>
                                <td>
                                    <font size="2">:סקפ</font>
                                </td>
                                <td dir="LTR">
                                    <input type="text" name="phone" value="<%=sPhone%>" size="15" tabindex="29">
                                </td>
                                <td>
                                    <font size="2">:ןופלט</font>
                                </td>
                            </tr>
                            <tr align="right">
                                <%site_url = replace(site_url,"""","&quot;")%>
                                <%email = replace(email,"""","&quot;")%>
                                <%if site_url="0" then%>
                                <td dir="LTR" align="right">
                                    <input type="text" name="site_url" value="http://" size="20" tabindex="32">
                                </td>
<%else
if trim(request("site_url"))="http://" then
                %>
<td dir="LTR" align="right">
                                    <input type="text" name="site_url" value="http://" size="20" tabindex="32">
                                </td>
<%else%>
<td dir="LTR" align="right">
                                    <input type="text" name="site_url" value="<%=site_url%>" size="20"
                                        tabindex="32">
                                </td>
<%end if
end if%>
<td nowrap>
                                    <font size="2">:טנרטניא רתא</font>
                                </td>
                                <td dir="RTL">
                                    <input type="text" name="email" value="<%=sEmail%>" size="15" tabindex="31">
                                </td>
                                <td nowrap>
                                    <font size="2">:ינורטקלא ראוד</font>
                                </td>
                            </tr>
                            <tr align="right">
                                <%sSettling = replace(sSettling,"""","&quot;")%>
                                <%sAddress = replace(sAddress,"""","&quot;")%>
                                <td dir="LTR">
                                    <input type name="settling" value="<%=sSettling%>" size="15" tabindex="34">
                                </td>
                                <td>
                                    <font size="2">:בושי</font>
                                </td>
                                <td dir="LTR">
                                    <input type name="address" value="<%=sAddress%>" size="25" tabindex="33">
                                </td>
                                <td>
                                    <font size="2">:תבותכ</font>
                                </td>
                            </tr>
                            <tr>
                                <td dir="LTR" valign="top" colspan="3" align="RIGHT">
                                    <input type name="zipcode" value="<%=iZipCode%>" size="15" tabindex="35">
                                </td>
                                <td valign="top" align="RIGHT">
                                    <font size="2">:דוקימ</font>
                                </td>
                            </tr>
<%
if cint(iSubjectId)=con_people then
'    if iYcount=1 then %>
               <!--TR align="right">
	 <TD dir=LTR valign="top">
	 <SELECT  name="salary_date">
	 <OPTION value="0" selected>בחר</OPTION>
	 <% for iYcount="1810" to year(now)
	    Response.Write "<OPTION value="& iYcount &">" & iYcount & "</OPTION>"
	    next
	 %>
	 </SELECT>
	 </TD>
	<TD valign="top" nowrap><font size="2">:תנשל</font></TD-->
               <%   'else

               %>
<tr align="right">
                                <td dir="LTR" valign="top">
                                    <select name="salary_date" tabindex="37&quot;">
<%if iSalaryDate = "" or isnull(iSalaryDate) then iSalaryDate="0"
if  iSalaryDate="0" or iSalaryDate = "" then	%>
<option value="0" selected>בחר</option>
                                        <option value="1">-</option>
<%else
if isalarydate="1" then%>
<option value="0">בחר</option>
                                        <option value="1" selected>-</option>
<%else%>
<option value="<%=iSalaryDate%>" selected>
<%=iSalaryDate%>
</option>
                                        <option value="0">בחר</option>
                                        <option value="1">-</option>
                                        <%end if%>
                                        <%end if%>
                                        <% for iYcount=year(now) to 1810 step -1
                                        Response.Write "<OPTION value="& iYcount &">" & iYcount & "</OPTION>"
                                        next
                                                              %>
                                    </select>
                                </td>
                                <td valign="top" nowrap>
                                    <font size="2">:תנשל רכש </font>
                                </td>
                                <td dir="LTR" valign="top">
                                    <input type="text" name="person_salary" value="<%=sSalary%>" size="20"
                                        tabindex="36">
                                </td>
                                <td valign="top" nowrap>
                                    <font size="2">:(םירלודב) יתנש רכש</font>
                                </td>
                            </tr>
<%'end if
else
              %>
              <tr align="right">
                  <td dir="LTR" valign="top">
                      <%'Response.Write isalarydate%>
<select name="salary_date" tabindex="37&quot;">
<%
if  iSalaryDate="0" or iSalaryDate="" or isnull(iSalaryDate) then%>
<option value="0" selected>בחר</option>
                                <option value="1">-</option>
<%else
if isalarydate="1" then %>
<option value="0">בחר</option>
                                <option value="1" selected>-</option>
<%else%>
<option value="<%=iSalaryDate%>" selected>
<%=iSalaryDate%>
</option>
                                <option value="0">בחר</option>
                                <option value="1">-</option>
                                <%end if%>
                                <%end if%>
                                <% for iYcount=year(now) to 1810 step -1
                                Response.Write "<OPTION value="& iYcount &">" & iYcount & "</OPTION>"
                                next
                                                          %>
                            </select>
                </td>
                <td valign="top" nowrap>
                    <font size="2">:תנשל</font>
                </td>
                <td dir="LTR" valign="top">
                    <input type="text" name="person_salary" value="<%=sSalary%>" size="36">
                </td>
                <td valign="top" nowrap>
                    <font size="2">:(םירלודב) יתנש רכש</font>
                </td>
            </tr>
<%end if%>
<tr align="right">
                <%
                		  'today=month(date) & "/" & day(date) & "/" & year(date)
                                    %>
                                    <td dir="RTL">
                                        <input type="text" name="updatedate" value="<%=date()%>" size="15" tabindex="39"
                                            readonly></td>
                                    <td>
                                        <font size="2">:ןוכדע ךיראת</font></td>
                                    <td dir="RTL">
                                        <input type="text" name="Cdate" value="<%=idate%>" size="15" tabindex="38"></td>
                                    <td>
                                        <font size="2">:הריצי ךיראת</font></td>
                                </tr>
                                <tr>
                                    <td align="RIGHT">
                                        <input type="button" value="העבר" name="updaterec1" onclick="javascript:ChangeHidden(3)"
                                            tabindex="60"><select name="permission" tabindex="27" dir="rtl">
                                                <option value="2" <%if request("status")="2" then response.write "selected"%>>מסמכים
                                                    מאושרים</option>
                                                <option value="1" <%if request("status")="1" then response.write "selected"%>>לא מוצגים</option>
                                                <%
                      sqlWebMasters="select * from t_webmaster"
                      set rsWeb=objDButils.GetFastRecoredset(sqlWebMasters)
                      do until rsWeb.EOF
                                                %>
                                                <option value="<%=rsWeb("id")%>" <%if Cstr(rsWeb("id"))=request("status") then response.write "selected"%>>
                                                    <%=Trim(rsWeb("user_name"))%>
                                                </option>
                                                <%
                      rsWeb.MoveNext
                      Loop
                      set rsWeb=nothing
                                                %>
                                            </select>
                                    </td>
                                    <td align="RIGHT">
                                        <font size="2">היקיתל הרבעה </font>
                                    </td>
                                    <%
                      if cint(iSubjectId)=con_people then
                      strSecondSubject=":םירשגמל ףרצ"
                      else
                      strSecondSubject=":םישיאל ףרצ"
                      end if

                                    %>
                                    <td align="right">
                                        <input type="checkbox" name="secondSubject" <%if ssubject = 1 then Response.Write "checked"%>
                                            tabindex="41" value="ON"></td>
                                    <td align="RIGHT">
                                        <font size="2">
                                            <%= strSecondSubject%>
                                        </font>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3" align="right">
                                        <font size="2">
                                            <input type="checkbox" name="AddToMailingList" <%if not MailListrs.eof and not trim(sEmail)="na"  then Response.Write "checked"%>
                                                tabindex="43" value="ON"></font></td>
                                    <td align="RIGHT">
                                        <font size="2">רוויד תמישרל ףרצ</font>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3" align="right">
                                        <input type="checkbox" name="ShowDetails" <%if Cstr(ShowDetails)="0" then Response.Write "checked"%>
                                            value="0" tabindex="150">
                                    </td>
                                    <td align="RIGHT">
                                        <font size="2">םיטרפ גיצהל אל</font></td>
                                </tr>
                                <tr>
                                            <td align="right" colspan="4" style="direction:rtl;">
                                                <font size="2">מסמך בתשלום:</font>&nbsp;<select name="OnlyMembers">
                                                <option value="0" <%if cstr(OnlyMembers & "") = "0" then Response.Write "selected"%>>פתוח</option>
                                                <option value="1" <%if cstr(OnlyMembers & "") = "1" then Response.Write "selected"%>>מסמך בתשלום</option>
                                                <option value="2" <%if cstr(OnlyMembers & "") = "2" then Response.Write "selected"%>>ללא הגבלה</option>
                                                </select>
                                            </td>
                                        </tr>
                                <tr align="center">
                                    <td colspan="6">
                                        <br>
                                        <br>
                                        <% if cint(request("status")) = 3 then %>
                                        <input type="button" value="לא מאושר" name="notgood" onclick="javascript:ChangeHiddenDel(5)"
                                            tabindex="44">
                                        <% else %>
                                        <input type="button" value="לא מאושר" name="notgood" onclick="javascript:ChangeHiddenDel(5)"
                                            tabindex="45">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <input type="button" value="עדכון" name="updaterec" onclick="javascript:ChangeHidden(3)"
                                            tabindex="46">
                                        <input type="button" value="מאושר לא להצגה" name="updaterec" onclick="javascript:ChangeHidden(4)"
                                            tabindex="47">
                                        <input type="button" value="אישור" name="allright" onclick="javascript:ChangeHidden(1)"
                                            tabindex="48">
                                        <%end if%>
                                    </td>
                                </tr>
                        </table>
                        </form>
                        <center>
                            <p>
                                <a href="main.asp">ישארה טירפתל הרזח</a><br>
                            </p>
                        </center>
                        </td>
                        <!-------------------- END MAIN TABLES ------------------------------------------->
                        <!-------------------------------------- RIGHT TABLES ------------------------------------------->
                        <td width="118" background="img/right_bg1.gif" bgcolor="#F5F5F5" valign="top" align="center"
                            style="border-left: 1px solid #C0C0C0; border-right: 1px solid #C0C0C0">
                            <table cellspacing="0" cellpadding="0" border="0">
                                <tr>
                                    <!------------- nfc search ---------------->
                                    <td align="center" valign="top">
                                    </td>
                                    <!------------- end nfc search ---------------->
                                </tr>
                                <tr>
                                    <!------------- links table ------------------->
                                    <td align="right" bgcolor="#F0F0F0">
                                    </td>
                                    <!-------------------------------------- END RIGHT TABLES ------------------------------------------->
                                </tr>
                            </table>
                        </td>
                        </tr> </table>
                    </div>
                </body>
                </html>
                <%
                function check(errorList)

                '====================check for news template=====goes into T_news====
                	userId= session("UserId")
                	status=cint(request("status"))	 'when the document first goes in the db the status is 0
                	permission=cstr(request("permission")) 'permission to update the doc comes from the managers only
                	priority =0  'only managers give priority for homepage show
                	parentdocid=request("parentdocid")
                	userdate=now()

                	'** head 1**
                	head1=replace(request("fHead1"),"'","''")
                	head1=replace(head1,"&#8211;","-")
                	'Response.Write "pppppp"
                	if head1="" then
                		sErrorStr1="יש להכניס כותרת ראשית "
                	else
                		sErrorStr1=""
                	end if
                	'** head2 **
                	'head2=replace(replace(request("fHead2"),"'",""),"-","&macaf")
                	head2=replace(request("fHead2"),"'","''")
                	head2=replace(head2,"&#8211;","-")
                	'** writername**
                	writerName="0"'replace(request("fWriter"),"'","")
                	'** image**
                	'upload
                	image=replace(request("image_url"),"'","''")
                	if image="" then
                	image="0"
                	end if

                	'** imgText **
                	imgText=replace(request("fImgText"),"'","''")

                	'** duty **
                	Pduty=replace(request("fDuty"),"'","''")
                	if request("fduty")= "" then

                	Pduty = 0

                	end if

                	'** institute **
                	institute=replace(request("fInstitute"),"'","''")
                	'** text**
                	'text=replace(replace(request("fText"),"'","''"),"-","&macaf")
                	text=replace(request("fText"),"'","''")
                	text=replace(text,"&#8211;","-")
                	if text="" then
                		'sErrorStr2="יש להכניס טקסט"
                	else
                		sErrorStr2=""
                	end if


                	'** linkname**
                	linkname=replace(request("fLinkName"),"'","''")
                	'** linkurl **
                	linkurl=replace(request("fLinkUrl"),"'","''")

                    '** PERMISSION **
                	'permission=cstr(request("permission"))
                '============= check for entity===== goes into T_entity===============
                	'into db :id,subject_id

                	'** hebrew dscr **
                		hebrewvalue=replace(request("hebrewvalue"),"'","''")
                		if hebrewvalue="" then
                		errorList.Add "6","hebrewvalue"
                		end if
                	'**english dscr**
                		englishvalue=replace(request("englishvalue"),"'","''")
                		if englishvalue="" then
                		  englishvalue = "-"
                		'	errorList.Add "7","englishvalue"
                		end if
                	'**first name **
                	fname=replace(request("fname"),"'","''")
                	if fname="" then
                		errorList.Add "1","fname"
                	end if

                	'** last name **
                	lname=replace(request("lname"),"'","''")
                	if lname="" then
                		errorList.Add "2","lname"
                	end if
                	'** duty_id**
                		'duty="null"
                		duty=cint(2)
                		duty2=request("duty2")
                		if duty2="" then
                			'errorList.Add "12","duty"
                			duty2="-"
                		end if
                	'** firm name**
                		'firm_name="null"
                		firm_name=replace(request("firm_name"),"'","''")
                	'** year of birth**
                		'birthyear="null" '
                		birthyear=request("fBirthyear")
                		if birthyear=0 then
                			'errorList.Add "14","birthyear"
                			birthyear="1"
                		end if
                		if birthyear="" then
                			birthyear = "1"
                		end if
                	'** decease_date**
                	    if request("decease_date")="" then
                		   decease_date="0" 'replace(request("decease"),"'","")
                        else
                           decease_date=replace(request("decease_date"),"'","")
                        end if
                	'** birthday **
                		'birthday="null"
                		if not (request("day")="" or request("month")="" or request("year")="") then
                		birthday=request("day")&"/"& request("month")&"/"& request("year")
                		else
                		birthday="0/0/0"
                		end if
                	'** sex**
                		sex=request("fsex")
                		if sex = "" then sex = "0"

                	'**nation1**
                		nation1=request("fNation1")

                	'**nation2**
                		nation2=request("fNation2")

                	'** Address **
                	address=replace(request("address"),"'","''")
                	'if Address="" then
                		'errorList.Add "3","address"
                	'end if
                	'**settling**
                		settling=replace(request("settling"),"'","''")
                		if settling="" then
                			'errorList.Add "13","settling"
                			settling="-"
                	   end if
                	'**zipcode**
                		zipcode=request("zipcode")
                		if not IsNumeric(zipcode) then
                			'errorList.Add "10","zipcode"
                			zipcode="0"
                		end if
                	'** Telephone **
                	phone=replace(request("phone"),"'","''")
                	if phone="" then
                		'errorList.Add "4","phone"
                		phone="-"
                	end if
                	'**fax**
                	fax=replace(request("fax"),"'","''")

                	'** E-mail **
                	email=replace(request("email"),"'","''")
                		if email = "" or not email="na" then
                			if not instr(1,email,"@")<> 0 Then
                			'errorList.Add "5","Email"
                			email="na"
                			end if
                		end if
                	'**site_url**
                		site_url=trim(request("site_url"))
                		if site_url="http://" then
                		   site_url="0"
                		end if
                	'entry date comes automatic

                	'**area**
                		area=request("area")
                		if area="" then
                			'errorList.Add "8","area"
                			area="כללי"
                		end if

                		Affair_id=request("Affair_id")
                		Lang=Request("Lang")
                		skill_id=Request("skill_id")
                		Category_id=Request("Category_id")
                		Province_id=Request("Province_id")


                	'** classcification**
                		classcification=cint(3)
                		classcification2=request("classcification2")
                		if classcification2="" then
                			'errorList.Add "11","classcification"
                			classcification2="-"
                		end if
                	'**deal type**

                		dealtype= 0    'request("dael_type")

                	'**person_total_salary**
                		person_total_salary=request("person_salary")
                		if person_total_salary="" then
                		    person_total_salary=0
                		end if

                	'**person_total_date**
                		person_salary_date=request("salary_date")

                	'**firm_total_balance**
                		firm_total_balance="0" 'request("firm_balance")

                	'**firm_balance_date**
                		firm_balance_date="null"  'request("firm_balance")

                	'**firm_total_income**
                		firm_total_income=0'"null" 'request("firm_income")

                	'**firm_income_date**
                		firm_income_date="0" 'request("firm_income")

                	'**firm_managers_total**
                		firm_managers_total=0'"null"'replace(request("five_managers"),"'","")

                	'update_date,doc_id
                	'**eduction**
                		fEducation=request("fEducation")
                		placeToBeBorn=request("placeToBeBorn")

                			'education="null"

                		strDate=month(date) & "/" & day(date) & "/" & year(date)
                		strFullDate=strDate & " " & time()

                		if isdate(request("cdate")) then
                		CreateDate = day(request("cdate")) & "/" & month(request("cDate")) & "/" & mid(year(request("cDate")),3,2)
                		else
                		CreateDate=request("cDate")
                		end if
                		if CreateDate = "" then CreateDate = date()'day(date) & "/" & month(date) & "/" & year(date)
                	'=======end validation============
                	If errorList.Count >0 or sErrorStr<>"" Then

                        check = False

                        Exit Function

                	else

                		Call UpdateHtmlStatus(request("docid"),request("subjectid"),0)
                		UnLockDoc request("docid"),request("subjectid")
                	Call WriteHistoryLog(request("docid"),request("subjectid"))


                	'** if there are no mistakes then we insert the user into the db
                   check=true

                       	getPage = "https://www.news1.co.il/CloudFlareAPI/DeleteCache.aspx?url=https://www.news1.co.il/Archive/" & GenerateName(request("docid"),request("subjectid"),"")
                        GetFastRemoteUrl(getPage)
                       	getPage = "https://www.news1.co.il/CloudFlareAPI/DeleteCache.aspx?url=https://m.news1.co.il/Archive/" & GenerateName(request("docid"),request("subjectid"),"")
                		GetFastRemoteUrl(getPage)


                	 select case Request("RecAction")
                		case "1"
                		   'אושר
                		   secondsubjectId=cint(request("secondSubject"))
                		   if request("secondSubject")="" then
                		   Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),2,request("parentDocId"),"T_people",2
                		   else
                		   Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),2,request("parentDocId"),"T_people",1
                		   end if
                		   com.sendEmailForConform email
                		   Response.Redirect "currentpages.asp"
                		case "2"
                		 ' לא אושר


                		  Com.DeleteT_peoplePermissionByDocAndSubject request("docid"),request("subjectid")

                		  Response.Redirect "main.asp"

                '============================================================================================================
                		case "3"

                	    	If status=0 Then
                			    status=0
                			Else
                		       if permission="1" then
                			      status=1
                			   Else
                			   	  Status=Cint(permission)
                			   end if
                		    End if



                  	     StrPeople= status & spe & head1 & spe & head2
                		 StrPeople= StrPeople & spe & Pduty & spe & institute & spe &  text & spe & image
                		 StrPeople= StrPeople & spe & imgText & spe & userId & spe & permission & spe & userdate & spe & writerName & spe & parentDocId

                		 Com.UpdateT_PeoplePermission request("docid"),request("subjectid"),StrPeople
                		 UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),status,request("parentDocId"),"T_people"


                	    '===str for insert to entity table====
                		 StrEntity= request("subjectid") & spe & hebrewvalue & spe & englishvalue & spe & fname & spe & lname
                		 	StrEntity=StrEntity+   spe & duty & spe & firm_name & spe & birthyear  & spe & decease_date & spe & birthday & spe & sex & spe & nation1 & spe & nation2
                		 StrEntity=StrEntity+   spe & Address & spe & settling & spe & zipcode & spe &  Phone & spe & fax & spe & Email
                		 StrEntity=StrEntity+   spe & site_url & spe & CreateDate & spe & area & spe & classcification & spe & dealtype
                		 StrEntity=StrEntity+   spe & person_total_salary & spe & person_salary_date & spe & firm_total_balance & spe & firm_balance_date
                		 	StrEntity=StrEntity+   spe & firm_total_income & spe & firm_income_date & spe & firm_managers_total & spe & strFullDate & spe & request("docid") & spe & fEducation & spe & placeToBeBorn
                		 	StrEntity=StrEntity+   spe & duty2 & spe & classcification2 & spe & replace(Affair_id&"","'","''") & spe & Request("skill_id")  & spe & Request("skill_Parent_id")  & spe & Lang & spe &category_id & spe & province_id  & spe & SearchWords
                		 StrEntity=StrEntity+   spe & moreInfo & spe & authorText
                '		com.UpdateEntity StrEntity
                		Call UpdateEntity(StrEntity)


                		 strsql=strsql & "UPDATE T_people SET "
                		 strsql=strsql & "personalDetails='" & Replace(request("personalDetails"),"'","''") &"',"
                		 strsql=strsql & "generalInfo='" & Replace(request("generalInfo"),"'","''") &"',"
                		 strsql=strsql & "educationDetails='" & Replace(request("educationDetails"),"'","''") &"',"
                		 strsql=strsql & "jobDetails='" & Replace(request("jobDetails"),"'","''") &"',"
                		 strsql=strsql & "experience='" & Replace(request("experience"),"'","''") &"',"
                		 strsql=strsql & "memberships='" & Replace(request("memberships"),"'","''") &"',"
                		 strsql=strsql & "publications='" & Replace(request("publications"),"'","''") &"',"
                		 strsql=strsql & "customCaption1='" & Replace(request("customCaption1"),"'","''") &"',"
                		 strsql=strsql & "customText1='" & Replace(request("customText1"),"'","''") &"',"
                		 strsql=strsql & "customCaption2='" & Replace(request("customCaption2"),"'","''") &"',"
                		 strsql=strsql & "customText2='" & Replace(request("customText2"),"'","''") &"'"
                		 strsql=strsql & " WHERE docid="& request("docid") &" AND subjectid="&request("subjectid")

                		 set UpdateConn=createobject("adodb.connection")
                		 UpdateConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms"
                		 UpdateConn.execute(strsql)


                         strsql="update t_entity set "
                		    strsql = strsql & "OnlyMembers=" & request("OnlyMembers")
                		 strsql=strsql & " WHERE doc_id="& request("docid") &" AND subjectid="&request("subjectid")
                		 UpdateConn.execute(strsql)
                    if request("status")="2" then
                		 txtToTags =  replace(head1 & "","''", "'") & chr(10) & replace(head2 & "","''", "'") & chr(10) & replace(text & "","''", "'")
                         UpdateArticleTagsGlobal request("docid"),request("subjectid"),txtToTags
                		end if



                		  sql = "delete from T_docsArea where docid = " & request("docid") & " and subjectid = " & request("subjectid")
                			UpdateConn.execute(sql)
                			sql = "delete from T_docsAffair where docid = " & request("docid") & " and subjectid = " & request("subjectid")
                			UpdateConn.execute(sql)


                			if request("area") <> "" then
                				areaId = getAreaIds(request("area"))
                				areaIdArr = split(areaId, ",")
                				for i = 0 to Ubound(areaIdArr)
                					if areaIdArr(i) <> "" and isnumeric(areaIdArr(i)) then
                						Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & "," & request("subjectid") & "," & areaIdArr(i) & ")"
                						UpdateConn.execute(Isql)
                					end if
                				next
                			end if

                			if request("Affair_id") <> "" then
                				AffairId = getAffairIds(request("Affair_id"))
                				AffairIdArr = split(AffairId, ",")
                				for i = 0 to Ubound(AffairIdArr)
                					if AffairIdArr(i) <> "" and isnumeric(AffairIdArr(i)) then
                						Isql = "insert into T_docsAffair (docid,subjectid,Affairid) values (" & request("docid") & "," & request("subjectid") & "," & AffairIdArr(i) & ")"
                						UpdateConn.execute(Isql)
                					end if
                				next
                			end if





                		 set UpdateConn=nothing


                					 sql="select T_Firms.docid,T_Firms.subjectid,head1 from T_Linked_Documents inner join T_Firms on T_Linked_Documents.docid=T_Firms.docid and T_Linked_Documents.subjectid=T_Firms.subjectid where T_Linked_Documents.LinkedDocID="&request("docId")&" AND T_Linked_Documents.LinkedSubjectID="&request("SubjectId")&" order by orderbyid"

                set objDbUtils=CreateObject("yoav.DButils")
                set rsLinkedPepole=objDbUtils.GetFastRecoredset(sql)

                do while not rsLinkedPepole.EOF
                 Call UpdateHtmlStatus(rsLinkedPepole("docid"),rsLinkedPepole("subjectid"),0)
                    rsLinkedPepole.movenext

                                    loop


                		if request("showDetails")="" Then showDetails=1 else showDetails=0

                	     set logoConn=createObject("adodb.connection")
                			logoConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
                				logoSql="update T_People set logo='"&session("imgLogo")&"', UpdatedByUserName='"&session("UpdatedByUserName")&"',showDetails='"&showDetails&"'  where docid="& request("docid") &" AND subjectid="&request("subjectid")

                			logoConn.execute logoSql
                			logoConn.close
                			set logoConn=nothing
                         if request("secondsubject")="" then
                		    Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),status,request("parentDocId"),"T_people",2
                	     else
                	        Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),status,request("parentDocId"),"T_people",1
                	     end if
                      '---------------------------------------------------------------------------------------------------'

                           if request("AddToMailingList") = "on"  then

                		      set objDButils=server.CreateObject("yoav.Com_Logic")
                		      strSqlDoc="select * from T_MailingList where doc_Id=" & request("docid") & " and subjectId= " & request("subjectid")

                		   set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)

                		      if MailListrs.eof then
                		         Com.InsertMailingList request("subjectid"),request("docid"),Email
                		      end if

                		   else

                		       set objDButils=server.CreateObject("yoav.Com_Logic")
                		       strSqlDoc="delete T_MailingList where email='" & email & "'"
                		       set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)

                		   end if

                		  'update Affairs if selected
                			'if trim(request("Affair_id")&"") <> "" and len(request("Affair_id")&"") > 1 then

                			'	set objDButils = server.createobject("yoav.dbutils")
                			'	arrAffair = split(request("Affair_id")&"",",")
                		'		for i = 0 to Ubound(arrAffair)
                			'		if trim(arrAffair(i)&"") <> "" then

                			'			set rs=objDButils.GetFastRecoredset("SELECT id,dscr FROM T_area WHERE status=2 and dscr = '"&trim(arrAffair(i)&"")&"'")
                			'			if not rs.EOF then
                			'			    call CreateOneHTMLPage("https://www.news1.co.il/createHtml/HtmlCreateAreaByID.aspx?areaid=" & rs("id"))
                			'			end if
                			'			set rs = nothing
                		'			end if
                			'	next
                		'		set objDButils = nothing

                		'	end if


                	  '---------------------------------------------------------------------------------------------------'
                		 if request("status")=0 then
                		     Response.Redirect "newdetails.asp"
                		 end if
                		 if request("status")=1 then
                		      Response.Redirect "Main.asp" '"waitinglistpage.asp"
                		 end if
                		 if request("status")=2 then
                		    Response.Redirect "Main.asp" '"currentpages.asp"
                		 elseif request("status")>2 then
                		 	Response.Redirect "Main.asp" '"currentpages.asp"
                		 end if

                '============================================================================================================

                		 if request("secondsubject")="" then
                		    Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),1,request("parentDocId"),"T_people",2
                	     else
                	        Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),1,request("parentDocId"),"T_people",1
                	     end if

                	       Response.Redirect "Main.asp" '"WaitingListPage.asp"

                        Case "4"
                						UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),1,request("parentDocId"),"T_people"
                					Response.Redirect "Main.asp" '"WaitingListPage.asp"

                	     Case "5"
                			Call UpdateT_newsPermissionByDocAndSubject(request("docid"),request("subjectid"),5,0,"T_people")
                			Response.Redirect "Main.asp" '"currentpages.asp"
                		End select


                		   if request("AddToMailingList") = "on"  then
                		    set objDButils=server.CreateObject("yoav.Com_Logic")
                		    strSqlDoc="select * from T_MailingList where doc_Id=" & request("docid") & " and subjectId= " & request("subjectid")
                		   set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)

                		   if MailListrs.eof then
                		        Com.InsertMailingList request("subjectid"),request("docid"),Email
                		   end if
                		  else
                		       set objDButils=server.CreateObject("yoav.Com_Logic")
                		       strSqlDoc="delete T_MailingList where email='" & email & "'"
                		       set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)


                		 end if

                    end if
                end function




                %>

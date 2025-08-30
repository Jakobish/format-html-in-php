<%@  language="VBScript" %>
<%Response.Buffer=true
Server.ScriptTimeout = 2147483647
session.Timeout = 40
starttime=timer

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1
%> <%session.LCID = 1037 %>
<!-- #include file=CheckSession.asp -->
<!-- #include file=include/ViewFunctions.asp -->
<!-- #include file=include/BidiFunctions.asp -->
<!-- #include file=include/const.asp -->
<!-- #include file=rose_ado.asp -->
<!--#include file="siteinclude/HtmlFunctions.asp"-->
<!--#include file=include/AddToDb.asp-->
<!--#include file=siteinclude/stringFunctions.asp-->
<!--#include file=LockFunctions.asp-->
<%
if session("UserId")="" then
	Response.Redirect "writeErorr.asp?iErrorId=1"
end if



If not( request("entityvalidate")="1" or request("validate")="1") Then
	CheckLock
end if


spe="#%$^"
dim iDocId
dim iSubjectId

iDocId=request("DocId")
iSubjectId=request("SubjectId")
dim sErrorStr1
dim sErrorStr2
sErrorStr1=""
sErrorStr2=""
docOpener=cstr(request("docOpener"))
subjectId=request("subjectId")
iLinkPeople=request("links")






Set errorList = com.ErrorList()
	

	If request("entityvalidate")="1" or request("newsvalidate")="1" Then
		If check(errorList) Then
			'Response.Redirect "CurrentPages.asp"
		end if
	end if

%>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-8 Visual">
  <title>New News</title>
<!--#include file=include/autofillDetails.asp-->
<SCRIPT LANGUAGE="JavaScript" SRC="CalendarPopup.js"></SCRIPT>
  <SCRIPT LANGUAGE="JavaScript">
  document.write(CalendarPopup_getStyles());
  </SCRIPT>
  <script language="javascript">
  function AddNewSource() {
    if (document.news.txtSourceName.value == '' && document.news.txtSourceMedia.value == '') {
      alert('נא מלא את שם המקור או כלי תקשורת');
    } else {
      window.open('editSources.asp?stage=addNew&docid=<%=idocid%>&subjectid=<%=isubjectid%>&txtSourceName=' +
        escape(document.news.txtSourceName.value) + '&txtSourceMedia=' + escape(document.news.txtSourceMedia.value) +
        '&txtSourceDate=' + escape(document.news.txtSourceDate.value) + '&txtSourceLink=' + escape(document.news
          .txtSourceLink.value), 'txtSources', 'width=600,height=450,scrollbars=yes,resizable=yes');
      document.news.txtSourceName.value = '';
      document.news.txtSourceMedia.value = '';
      document.news.txtSourceDate.value = '<%=date() %>';
      document.news.txtSourceLink.value = '';
    }

  }

  function NewSource() {
    window.open('editSources.asp?docid=<%=idocid%>&subjectid=<%=isubjectid%>', 'txtSources',
      'width=600,height=450,scrollbars=yes,resizable=yes');
  }

  function AddNewRLinkR() {
    if (document.news.NewreferenceText.value == '') {
      alert('נא מלא את התאור מראה מקום');
    } else {
      if (document.news.NewreferenceURL.value == '') {
        alert('נא מלא את הקישור למראה מקום');
      } else {
        window.open('txtLinkRefText.asp?stage=addNew&docid=<%=idocid%>&subjectid=<%=isubjectid%>&txt=' + escape(
            document.news.NewreferenceText.value) + '&link=' + escape(document.news.NewreferenceURL.value) +
          '&NewOrderNum=1', 'AddNewRLinkRFFF', 'width=700,height=500,scrollbars=yes,resizable=yes');
        document.news.TextReplace.value = '';
        document.news.linkReplace.value = '';
      }
    }

  }

  function NewRLinkR() {
    window.open('txtLinkRefText.asp?docid=<%=idocid%>&subjectid=<%=isubjectid%>', 'txtLinkReplace',
      'width=700,height=500,scrollbars=yes,resizable=yes');
  }

  function AddNewLinkR() {
    if (document.news.TextReplace.value == '') {
      alert('נא מלא את המילה להחלפה');
    } else {
      if (document.news.linkReplace.value == '') {
        alert('נא מלא את הקישור להחלפה');
      } else {
        window.open('txtLinkReplace.asp?stage=addNew&docid=<%=idocid%>&subjectid=<%=isubjectid%>&txt=' + escape(
            document.news.TextReplace.value) + '&link=' + escape(document.news.linkReplace.value), 'txtLinkReplace',
          'width=450,height=450,scrollbars=yes,resizable=yes');
        document.news.TextReplace.value = '';
        document.news.linkReplace.value = '';
      }
    }

  }

  function NewLinkR() {
    window.open('txtLinkReplace.asp?docid=<%=idocid%>&subjectid=<%=isubjectid%>', 'txtLinkReplace',
      'width=450,height=450,scrollbars=yes,resizable=yes');
  }

  function openMainLinks() {
    window.open('addLinkToMainText.asp?docid=<%=idocid%>&subjectid=<%=isubjectid%>', 'txtLinkReplace',
      'width=600,height=500,scrollbars=yes,resizable=yes');
  }

  function UploadeImage() {
    var sUploade = window.open('', sUploade, 'toolbar=no,status=no,resizable=yes,width=300,height=400,scrollbars=yes');
    sUploade.location.href = 'UPLOADsend.asp';
  }

  function UploadeVideo() {
    var sUploade = window.open('', sUploade,
      'toolbar=no,status=yes,statusbar=yes,personalbar=yes,menubar=yes,resizable=yes,width=300,height=400,scrollbars=yes'
      );

    sUploade.location.href = 'UPLOADsendVideo.asp';

  }

  function GetVideoPath(sPath, sPhName) {
    eval('document.news.SVidioFile.value =  sPath;');
    if (sPhName != "") {
      eval('document.news.videoText.value =  sPhName;');
    }
  }

  function GetImagePath(sPath, sPhName) {
    eval('document.news.image_url.value =  sPath;');
    if (sPhName != "") {
      eval('document.news.SImageText.value =  sPhName;');
    }
  }

  function submitNews() {
    //alert("submitNews")
    document.news.submit()

  }

  function setOrder(SubjectId, DocId) {
    window.open("linksByOrder.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "",
      "width=600,height=500,scrollbars=yes");
  }

  function submitFrmAfterOpenWin() {
    document.news.fromOpenWin.value = 'YES';
    document.news.RecAction.value = 3;
    document.news.submit();
  }

  function ChangeHidden(status) {
    document.news.RecAction.value = status;
    document.news.submit();

  }

  function ChangeHiddenDel(status) {
    if (confirm("?למחוק רשומה זו") == true) {
      document.news.RecAction.value = status;
      document.news.submit();
    }
  }

  function Showlegal(isubject, idoc)

  {
    VisualWindow = window.open('', 'VisualWindow',
      'toolbar=no,status=no,resizable=yes,width=450,height=560,scrollbars=no');
    VisualWindow.location.href = 'HomePageEdit.asp?idocid=' + idoc + '&isubjectid=' + isubject
  }

  function UpdateLink(ID, SubjectId, DocId)

  {
    window.open("UpdateLink.asp?linkId=" + ID + "&subjectId=" + SubjectId + "&DocId=" + DocId + "&Action=1", "",
      "width=600,height=300")

  }

  function AddALink(SubjectId, DocId) {
    var sTitle = document.news.fHead1.value
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

  function setOrder(SubjectId, DocId) {
    window.open("linksByOrder.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "",
      "width=600,height=500,scrollbars=yes");
  }

  function GetLinks() {
    LinksPoeple = document.news.links.value
    document.news.links.value = parseInt(LinksPoeple) + 1;
    //alert(document.people.links.value);
    document.news.submit()
  }

  function LogOut() {
    if (event.clientY < 0) {
      var objHTTPGame = new ActiveXObject("Microsoft.XMLHTTP");
      url = "LogOutPopUp.asp?userID=<%=Session("LockedByID")%>"
      objHTTPGame.Open("get", url, false);
      objHTTPGame.send();

      //ww5=window.open('LogOutPopUp.asp?userID=<%=Session("LockedByID")%>','LogOutPopUp','width=100,height=100,left=0,top=0,scrollbars=no');
    }
  }

  function GetPartners() {
    document.news.gotArticle.value = 1;
    document.news.afterSubmit.value = 1;
    document.news.submit()
  }

  function AddNewsText(SubjectId, DocId) {
    window.open("AddNewsText.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "",
    "width=600,height=500,scrollbars=yes");
  }

  function AddNewsTextLinks(SubjectId, DocId, recordId) {
    window.open("addNewsLinkToText.asp?subjectId=" + SubjectId + "&DocId=" + DocId + "&textID=" + recordId, "textLinks",
      "top=0,left=0,width=600,height=500,scrollbars=yes");
  }

  function setOrder2(SubjectId, DocId) {
    window.open("newsTextByOrder.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "",
      "width=600,height=500,scrollbars=yes");
  }

  function AddLeadsText(SubjectId, DocId) {
    window.open("AddLeadText.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "",
    "width=600,height=500,scrollbars=yes");
  }

  function setOrderLeads(SubjectId, DocId) {
    window.open("LeadsByOrder.asp?subjectId=" + SubjectId + "&DocId=" + DocId, "",
      "width=600,height=500,scrollbars=yes");
  }

  function openWindow(url, name, h) {
    popupWin = window.open(url, name, 'resizable,scrollbars=yes,width=650,height=' + h + ',top=0,left=0')
  }

  function InsertLinkedTags() {
    strSearch2 = document.FrmLinkedTags.strSearch.value;
    if (strSearch2 != '') {
      openWindow('AddLinkedTag.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>&strSearch=' + strSearch2,
        'LinkedTags', 500)
    } else {
      alert('יש להכניס טקסט לחיפוש');
    }
  }

  function InsertLink() {
    var URL = window.prompt("Enter the URL you want link to!", "http://");
    if (URL == "" || URL == null || URL == "http://") {
      alert("לא הוקלד קישור");
    } else {
      var LinkName = window.prompt("הוסף שם או תיאור לקישור", "");
      if (LinkName == "" || LinkName == null) {
        alert("יש להקליד שם או תיאור לקישור");
      } else {
        document.news.fText.value = document.news.fText.value + " <a class=links  href='" + URL + "' target='_blank'>" +
          LinkName + "</a>";
      }
    }
  }

  function InsertLinkedDocument() {
    strSearch = document.FrmLinkedDocs.strSearch.value;
    if (strSearch != '') {
      openWindow('AddLinkedDocument.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>&strSearch=' + strSearch,
        'LinkedDocs', 500)
    } else {
      alert('יש להכניס טקסט לחיפוש');
    }
  }

  function InsertLinkedImages() {
    strFile = document.FrmLinkedImages.imgName.value;
    imgDescription = document.FrmLinkedImages.imgDescription.value
    if (strFile != '') {
      document.FrmLinkedImages.imgName.value = '';
      document.FrmLinkedImages.imgDescription.value = '';
      openWindow('AddLinkedImages.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>&imgName=' + strFile +
        '&imgDescription=' + imgDescription, 'LinkedDocs', 500)

    } else {
      alert('יש לבחור תמונה');
    }
  }

  function newwinGallery3() { //v2.0
    ww22 = open('XuploadPopUp.asp?FormName=FrmLinkedBotImages&FieldName=imgName', 'gallerywin',
      'width=650 height=400 top=0 left=0 scrollbars=yes');
  }

  function InsertLinkedBotImages() {
    strFile = document.FrmLinkedBotImages.imgName.value;
    imgDescription = document.FrmLinkedBotImages.imgDescription.value
    if (strFile != '') {
      document.FrmLinkedBotImages.imgName.value = '';
      document.FrmLinkedBotImages.imgDescription.value = '';
      openWindow('AddLinkedBotImages.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>&imgName=' + strFile +
        '&imgDescription=' + imgDescription, 'LinkedDocs', 500)

    } else {
      alert('יש לבחור תמונה');
    }
  }

  function newwinLogo() { //v2.0
    ww = open('frmAddLogo.asp', 'logowin', 'width=250 height=250 top=100 left=100 scrollbars=no');
  }

  function newwinLogo2(fldName, frmName) { //v2.0
    ww2 = open('GeneralUPLOAD.asp?fieldName=' + fldName + '&formName=' + frmName, 'pic2',
      'width=250 height=250 top=100 left=100 scrollbars=no');
  }

  function newwinGallery() { //v2.0
    ww22 = open('XuploadPopUp.asp', 'gallerywin', 'width=650 height=400 top=0 left=0 scrollbars=yes');
  }

  function newwinGalleryVideo() { //v2.0
    wwvv = open('XuploadVideoPopUp.asp', 'videowin', 'width=650 height=400 top=0 left=0 scrollbars=yes');
  }

  function newwinGallery2(FormName, FieldName) { //v2.0
    ww22 = open('XuploadPopUp.asp?FormName=' + FormName + '&FieldName=' + FieldName, 'gallerywin',
      'width=650 height=400 top=0 left=0 scrollbars=yes');
  }

  function newwinGallery2WithDesc(FormName, FieldName, DescFieldName) { //v2.0
    ww22 = open('XuploadPopUp.asp?FormName=' + FormName + '&FieldName=' + FieldName + '&descFieldName=' + DescFieldName,
      'gallerywin', 'width=650 height=400 top=0 left=0 scrollbars=yes');
  }
  </script>

  <style type="text/css">
  input {
    font-size: 12px;
    font-family: Arial;
  }

  select {
    font-size: 12px;
    font-family: Arial;
  }
  </style>

</head>
<body topmargin="0" style="font-family: Arial;" onbeforeunload="LogOut()">
<%		YoavTime = now
                ShowFastNews=1
                CountFastNews=50
                StartFastNews=8
                EndFastNews=14
		set objDButils = server.CreateObject("Yoav.Com_logic")
		strSqlDoc="select * from T_news where docId=" & iDocId & " and subjectId= " & iSubjectId & " and status="& request("status")
			
			
			'Response.Write strSqlDoc
			'Response.End
			set subjectrs=objDButils.GetFastRecoredset(strSqlDoc)
			


			if subjectrs.bof then
				Response.Write "No Documents found."
				
			else
			
			if subjectrs("writername")<>session("UpdatedByUserName") and session("userid")=3 then 
				response.redirect "currentpages.asp?do=noPermissions" 
			end if

			
				while not subjectrs.eof
				ShortNews=subjectrs("ShortNews")
				ShowInShortNews=subjectrs("ShowInShortNews")
				ShowInBursaNames=subjectrs("ShowInBursaNames")
				iUserId=subjectrs("userId")
				sHead1=subjectrs("head1")
				sHead2=subjectrs("head2")
				sWriter=subjectrs("writername")
				sImage=subjectrs("image")
				sImage2=subjectrs("image2")
				sImage3=subjectrs("image3")
				sImage4=subjectrs("image4")
				sImage5=subjectrs("image5")
				sImage6=subjectrs("image6")
				
				YoavTime=subjectrs("YoavTime")
				sDate=day(YoavTime) & "/" & month(YoavTime) & "/" & year(YoavTime)
				
				sTime=subjectrs("userhour")
				sText=subjectrs("Mtext")
				slinkName=subjectrs("linkName")
				slinkurl=subjectrs("linkUrl")
				status=subjectrs("status")
				SVidioFile=subjectrs("VidioFile")
				VideoYouTube=subjectrs("VideoYouTube")
				SImageText=subjectrs("ImageText")
				SImageText2=subjectrs("ImageText2")
				SImageText3=subjectrs("ImageText3")
				SImageText4=subjectrs("ImageText4")
				SImageText5=subjectrs("ImageText5")
				SImageText6=subjectrs("ImageText6")
				session("imgLogo")=subjectrs("logo")
				TopTitle=Replace(subjectrs("TopTitle")&"",chr(34),"&quot;")
				bottomIMG=subjectRS("bottomIMG")
				Special=subjectRS("Special")
				SpecialAffair=subjectRS("SpecialAffair")
				RelatedSpecial=subjectRS("RelatedSpecial")

                    ShowFastNews=subjectRS("ShowFastNews")
                    CountFastNews=subjectRS("CountFastNews")
                    StartFastNews=subjectRS("StartFastNews")
                EndFastNews=subjectRS("EndFastNews")
				
				
				linkReplace=subjectRS("linkReplace")&""
				TextReplace=subjectRS("TextReplace")&""
				
				if TextReplace <> "" then
					call updateTxtLinksReplaceTable(TextReplace,linkReplace,iDocId,iSubjectId)
				end if
				
				referenceURL=subjectRS("referenceURL")
				referenceTEXT=subjectRS("referenceTEXT")
				referenceURL2=subjectRS("referenceURL2")
				referenceTEXT2=subjectRS("referenceTEXT2")
				referenceURL3=subjectRS("referenceURL3")
				referenceTEXT3=subjectRS("referenceTEXT3")
				videoText=subjectrs("videoText")
				LargeVideo=subjectrs("LargeVideo")
				leads=subjectrs("leads")
				adverStatus=subjectRS("adverStatus")
				
           'new ddl
            invoiceStatus=subjectRS("invoiceStatus")
				NoFindReplace=subjectRS("NoFindReplace")

                TopTitleColor=subjectRS("TopTitleColor") & ""
                TopTitleFontSize=subjectRS("TopTitleFontSize") & ""
                Head1Color=subjectRS("Head1Color") & ""
                Head1FontSize=subjectRS("Head1FontSize") & ""
				adverStatus12Home=subjectRS("adverStatus12Home")
				subjectrs.movenext
				wend
				subjectrs.close
			end if
			
			
			
			
set objDButils=server.CreateObject("Yoav.Com_logic")
		'set objConvert=server.CreateObject("yoav.convert2visual")
		
'			strSqlDoc="select * from T_entity where doc_Id=" & iDocId & " and subject_Id= " & iSubjectId 
		strSqlDoc="select * from T_entity where doc_Id=" & iDocId & " and subjectId= " & iSubjectId 
			'Response.Write strSql
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
				category_id=NTTrs("category_id")&""
				Province_id=NTTrs("Province_id")&""
				iClasscification=NTTrs("classcification")
				iDuty=NTTrs("duty_id")
				site_url=NTTrs("site_url")
				iplaceToBeBorn=NTTrs("placeToBeBorn")
				ifirm_name=NTTrs("firm_name")
				JobTitle=NTTrs("JobTitle")
				iDuty2=NTTrs("new_duty")
				iclasscification2=NTTrs("new_classcification")
				'eDate=NTTrs("entry_date")
				sdate=day(YoavTime) & "/" & month(YoavTime) & "/" & year(YoavTime)
				uDate=NTTrs("update_date")
				moreinfo=NTTrs("moreinfo")
				adminComments=NTTrs("adminComments")
				autoUserInfo = NTTrs("autoUserInfo")
				ShowChat = NTTrs("ShowChat")
				ShowStarsOnLeft = NTTrs("ShowStarsOnLeft")
				OnlyMembers = NTTrs("OnlyMembers")
                AutoAffairDocs = NTTrs("AutoAffairDocs")
				authorText=NTTrs("authorText")

                     EventBoard = NTTrs("EventBoard")
                    EventBoardStartDate = NTTrs("EventBoardStartDate")
                    EventBoardEndDate = NTTrs("EventBoardEndDate")

                    if EventBoard = 1 then
                        if isdate(EventBoardStartDate) and cstr(EventBoardStartDate&"") <> "" then
                            if day(EventBoardStartDate) < 10 then
                            EventStartDate =  "0" & day(EventBoardStartDate)
                            else
                            EventStartDate =  day(EventBoardStartDate)
                            end if
                            if month(EventBoardStartDate) < 10 then
                            EventStartDate =  EventStartDate & "/0" & month(EventBoardStartDate)
                            else
                            EventStartDate =  EventStartDate & "/" & month(EventBoardStartDate)
                            end if
                            EventStartDate =  EventStartDate & "/" & year(EventBoardStartDate)
                            EventStartTime = formatdatetime(EventBoardStartDate, 4)
                        else
                            EventStartDate = ""
                            EventStartTime = ""
                        end if
                        if isdate(EventBoardEndDate) and cstr(EventBoardEndDate&"") <> "" then
                            if day(EventBoardEndDate) < 10 then
                            EventEndDate =  "0" & day(EventBoardEndDate)
                            else
                            EventEndDate =  day(EventBoardEndDate)
                            end if
                            if month(EventBoardEndDate) < 10 then
                            EventEndDate =  EventEndDate & "/0" & month(EventBoardEndDate)
                            else
                            EventEndDate =  EventEndDate & "/" & month(EventBoardEndDate)
                            end if
                            EventEndDate =  EventEndDate & "/" & year(EventBoardEndDate)

                            EventEndTime = formatdatetime(EventBoardEndDate, 4)
                        else
                            EventEndDate = ""
                            EventEndTime = ""
                        end if
                    else
                        EventStartDate = ""
                        EventStartTime = ""
                        EventEndDate = ""
                        EventEndTime = ""
                    end if


				NTTrs.movenext
				wend 

	        strSqlDoc="select * from T_MailingList where doc_Id=" & idocId & " and subjectId= " & iSubjectId
			'Response.Write strSqlDoc 
			set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)
			end if

			
			
    %>
<div align="center">
    <table cellspacing="0" cellpadding="0" border="0" width="751" style=" border: 1px solid #C0C0C0">
      <tr>
        <!-------------------------- LEFT TABLES --------------------------------------->
        <td bgcolor="#F5F5F5" width="200" valign="top" align="right"
          style="padding-top: 10px; border-right: 1px solid #C0C0C0;">
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;"><a
                  href="javascript:openWindow('linkedFiles.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','Files',400)">
                  <font dir="rtl">מסמכי מקור</font>
                </a><br>
                <br>
                <a
                  href="javascript:openWindow('search/frmAddsearch.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','searh',400)">
                  ימניד רושיק</a><br>
                <br>
                <a
                  href="javascript:openWindow('links/frmAddLink.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','links',400)">
                  ךרע יפל רושיק</a><br>
                <br>
                <a
                  href="javascript:openWindow('AddFeedback.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','rates',400)">
                  בושמ לוהינ</a><br>
                <br>
                <a
                  href="javascript:openWindow('replyes/adminResponse.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','reply',400)">
                  תובוגת לוהינ</a>
<%set aConn=createObject("adodb.connection")
		                    aConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
		                        sql = "select count(Aid) as c from T_MemberUploads where docid=" & request("docId") & " and subjectid=" & request("SubjectId") 
		                    set rsA = aConn.Execute(sql)
                            if rsA("c") > 0 then %>
<br><br>
                <a href="javascript:openWindow('MemberUploads.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','reply',500)"
                  dir="ltr">
                  מסמכים מבימה חופשית</a>
<%end if
                        set rsA = nothing
                        set aConn = nothing
                         %>
<br /><br /><a
                  href="javascript:openWindow('ManageArticleTables.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','ManageArticleTables',600)"
                  dir="ltr">
                  ניהול טבלאות</a><br>
                <br>
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;">
                <form name="FrmLinkedImages" method="get">
                  <span align="center"><u><b>
                        <font face="Arial" dir="rtl" size="2">עריכת תמונות</font>
                      </b></u><br>
                    <input type="button"
                      onclick="window.open('frmAddLinkedImage.asp','LinkedImageUpload','resizable,scrollbars=yes,width=200,height=200,top=200,left=200');"
                      name="b13" value="בחר">
                    <input type="text" name="imgName" size="6" dir="rtl">
                    <a onclick="newwinGallery2('FrmLinkedImages','imgName');" href="javascript:void(0);">
                      <img border="0" src="EditorIcons/icon_ins_image.gif"></a><br>
                    <textarea name="imgDescription" dir="rtl" rows="5" style="font-family: Arial; font-size: 11px;
                                width: 100%;"></textarea><br>
                    <input type="Button" onclick="InsertLinkedImages();" name="b12" value="הוסף">
                    <input type="button"
                      onclick="openWindow('EditLinkedImages.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','editLinkedImages',300);"
                      name="b12" value="ערוך">
                  </span>
                </form>
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;"><span align="center" dir="rtl"><b><u>
                      <font face="Arial" dir="rtl" size="2">עריכת לידים</font>
                    </u></b><br>
                  <a href="javascript:AddLeadsText(<%=iSubjectId%>,<%=iDocId%>)" style="font-family: Arial;"
                    dir="rtl">הוסף/עדכן ליד</a><br>
                  <br>
                  <a href="javascript:setOrderLeads(<%=iSubjectId%>,<%=iDocId%>)" style="font-family: Arial;"
                    dir="rtl">קבע סדר</a> </span></td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;">
                <form name="FrmLinkedDocs" method="get">
                  <span align="center"><b><u>
                        <font face="Arial" dir="rtl" size="2">עריכת מסמכים מקושרים</font>
                      </u></b><br>
                    <input type="text" name="strSearch" dir="rtl" style="width: 100%;"><br>
                    <input type="button" onclick="InsertLinkedDocument();" name="b11" value="הוסף">
                    <input type="button"
                      onclick="openWindow('EditLinkedDocument.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','editLinkedDocs',300);"
                      name="b12" value="ערוך">
                  </span>
                </form>
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;">
                <form name="FrmLinkedBotImages" method="get">
                  <span align="center"><u><b>
                        <font face="Arial" dir="rtl" size="2">עריכת תמונות בתחתית</font>
                      </b></u><br>
                    <input type="button"
                      onclick="window.open('frmAddLinkedBotImage.asp','LinkedBotImageUpload','resizable,scrollbars=yes,width=200,height=200,top=200,left=200');"
                      name="b13" value="בחר"><input type="text" name="imgName" size="6" dir="rtl"><a
                      onclick='newwinGallery3();' href="javascript:void(0);"><img border="0"
                        src="EditorIcons/icon_ins_image.gif"></a><br>
                    <textarea name="imgDescription" dir="rtl" rows="5" style="font-family: Arial; font-size: 11px;
                                width: 100%;"></textarea><br>
                    <input type="Button" onclick="InsertLinkedBotImages();" name="b12" value="הוסף">
                    <input type="button"
                      onclick="openWindow('EditLinkedBotImages.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','editLinkedImages',300);"
                      name="b12" value="ערוך">
                  </span>
                </form>
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;">
                <form name="FrmLinkedTags" id="FrmLinkedTags" method="get">
                  <span align="center"><b><u>
                        <font face="Arial" dir="rtl" size="2">עריכת תגיות</font>
                      </u></b><br>
                    <input type="text" name="strSearch" style="width: 100%" dir="rtl"><br>
                    <input type="button" onclick="InsertLinkedTags();" name="b11" value="הוסף">
                    <input type="button"
                      onclick="openWindow('EditLinkedTags.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','editLinkedTags',300);"
                      name="b12" value="ערוך">
                  </span>
                </form>
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;"><span align="center"><b><u>
                      <font face="Arial" dir="rtl" size="2">ניהול קישורים מכלי תקשורת אחרים</font>
                    </u></b><br />
                  <input type="button"
                    onclick="openWindow('AddRssLinksToArticle.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','AddRssLinksToArticle',300);"
                    value="הוסף">
                  &nbsp;&nbsp;&nbsp;<input type="button"
                    onclick="openWindow('EditRssLinksInArticle.asp?docId=<%=request("docId")%>&subjectId=<%=request("SubjectId")%>','EditRssLinksInArticle',300);"
                    value="ערוך">
                </span></td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; background-color:#E6F0FB;">
            <tr>
              <td align="right" style="padding:10px;">
                <table border="0" cellpadding="0" cellspacing="0">
                  <tr align="center">
                    <td colspan="2" align="center">
                      <%
	
	 dim sSql
	 
	 sSql= " SELECT * FROM T_links WHERE " &_
	       " docId="& iDocId &" AND " &_
	       " subjectId="& iSubjectId & _
	       " ORDER BY orderBynum,linkName "
	 
	 set objDButils=server.CreateObject("yoav.com_logic")
	
	 set rs=objDButils.GetFastRecoredset(sSql)
		LinkCount=1
		sLinkShow="<tr><td colspan=2 align='right'><font size=""2"" color=""#000000""><b><u>םירושיק</font></b></u></td></tr>"
	
		while not rs.eof 
	
		sLinkShow=sLinkShow &  "<tr><td colspan = '2' align=right dir=rtl><font size=""-1"" color=""#000000"">"
		sLinkShow=sLinkShow &  "<font size=""-1"" color=""#000000"">&nbsp;<font color=black size=""-1"">"& LinkCount &".</font>"
		sLinkShow=sLinkShow &  rs("linkName")
		sLinkShow=sLinkShow &  "</font><br>"
		sLinkShow=sLinkShow &  "<a href='javascript:UpdateLink("& rs("id")&","& rs("subjectId")&","& rs("docId") &")'>"
		sLinkShow=sLinkShow &  "עדכן</a>"
		sLinkShow=sLinkShow &  "&nbsp;<a href='javascript:DeletLink("& rs("id")&","& rs("subjectId")&","& rs("docId") &")'>"
		sLinkShow=sLinkShow &  "מחק </a></font></td></tr>"
		LinkCount=LinkCount+1
		rs.movenext
	
		wend
	

                                %>
                      <%
	Response.Write  sLinkShow
	
                                %>
                  <tr align="right">
                    <td colspan="2" valign="top" align="right">
                      <br>
                      <a href="javascript:setOrder(<%=subjectId%>,<%=iDocId%>)" tabindex="9"
                        style="font-size: 12px;">
                        רדס עבק</a>&nbsp;&nbsp; <a href="javascript:AddALink(<%=isubjectId%>,<%=iDocId%>)"
                        tabindex="8" style="font-size: 12px;">רושיק ףסוה</a>
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%; height:30px; background-color:#FFFFFF;">
            <tr>
              <td align="center">
                <hr size="3" width="100%" color="#333333" />
              </td>
            </tr>
          </table>
          <form action="UPDATENEWS.asp" method="post" name="news">
            <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="rtl">
              <tr>
                <td align="right">
                  <input type="checkbox" name="adverStatus" value="1" <%if Trim(adverStatus)="1" Then response.write "checked"%> tabindex="10">
                  <font size="2">מדור פרסומי</font>
                </td>
              </tr>
              <tr>
                <td align="right" style="padding-right:10px; padding-left:10px;">

                  <select id="ddlInvoice" name="invoiceStatus">
                    <option value="0">
                      מצב חשבונית
                    </option>
                    <option value="1" <%if Trim(invoiceStatus)="1" Then response.write "selected"%>>
                      פטור
                    </option>
                    <option value="2" <%if Trim(invoiceStatus)="2" Then response.write "selected"%>>
                      טרם
                    </option>
                    <option value="3" <%if Trim(invoiceStatus)="3" Then response.write "selected"%>>
                      חשבונית
                    </option>
                    <option value="4" <%if Trim(invoiceStatus)="4" Then response.write "selected"%>>
                      שולם
                    </option>
                  </select>

                </td>
              </tr>
              <tr>
                <td align="right">
                  <input type="checkbox" name="adverStatus12Home" value="1" <%if Trim(adverStatus12Home)&""<>"0" Then response.write "checked" %>>
                  <font size="2">הצג כתבה מקודמות בדף הבית</font>
                </td>
              </tr>
            </table>
        </td>
<%
    If cint(subjectId)=1 Then
       caption_str="תושדח ןכדע"
    Else
       caption_str="תועדוה ןכדע"
    End If
                %>
<td width="491" valign="top" align="center">
          <table border="0" cellspacing="0" cellpadding="0" width="480">
            <input type="hidden" name="newsvalidate" value="1">
            <input type="Hidden" name="RecAction" value="">
            <input type="Hidden" name="docid" value='<%=request("docid")%>'>
            <input type="Hidden" name="status" value='<%=request("status")%>'>
            <input type="Hidden" name="subjectid" value='<%=request("subjectid")%>'>
            <input type="Hidden" name="ParentDocId" value='<%=request("ParentDocId")%>'>
            <input type="hidden" name="docOpener" value="<%=request("docOpener")%>">
            <input type="hidden" name="strSearch" value="<%=request("strSearch")%>">
            <input type="hidden" name="txtSearch" value="<%=request("txtSearch")%>">
            <input type="hidden" name="opSearch" value="<%=request("opSearch")%>">
            <input type="hidden" name="gotArticle" value="<%=request("gotArticle")%>">
            <input type="hidden" name="afterSubmit" value="<%=afterSubmit%>">
            <input type="hidden" name="fromOpenWin" value="">
<%
		if request("links")="" then
                        %>
<input type="hidden" name="links" value="0">
<%else%>
<input type="hidden" name="links" value='<%=request("links")%>'>
<%end if%>
<tr>
              <td colspan="2" align="center">
                <b>
                  <font color='navy' face='Arial'>
<%=caption_str%>
</font>
                </b>
                <br>
              </td>
            </tr>
            <tr align="right">
              <td align="right" nowrap>
                <font size="3" color="red">הבוח תודש םניה * ב םינמוסמה תודשה</font>
              </td>
            </tr>
            <tr>
              <td colspan="2" align="right">
                <font color="red" size="2">
<%
		if sErrorStr1 <> "" or sErrorStr2 <> "" then
		Response.Write translatelogical2visual(sErrorStr1,60) & "<br>"
		Response.Write translatelogical2visual(sErrorStr2,20) 
		end if
		
                                    %>
</font>
              </td>
            </tr>
      </tr>
      <tr>
        <td align="right" colspan="2" style="font-family: Arial; font-size: 12px; direction:rtl;"><input type="radio"
            name="ShortNews" value="0" <%if cstr(ShortNews & "") = "0" and cstr(ShowInShortNews & "") = 0 and cstr(ShowInBursaNames & "") = 0 then response.write "checked"%> />&nbsp;חדשות&nbsp;&nbsp;&nbsp;<input type="radio"
            name="ShortNews" value="1" <%if cstr(ShortNews&"") = "1" and cstr(ShowInShortNews&"") = 0 then response.write "checked"%> />&nbsp;מבזק חדשות&nbsp;&nbsp;&nbsp;<input type="radio"
            name="ShortNews" value="2" <%if cstr(ShowInShortNews&"") = "1" then response.write "checked"%> />&nbsp;חדשות ומבזק חדשות&nbsp;&nbsp;&nbsp;<input type="radio"
            name="ShortNews" value="3" <%if cstr(ShortNews&"") = "2" and cstr(ShowInBursaNames&"") = 0 then response.write "checked"%> />&nbsp;פַּנְדוֹרָה&nbsp;&nbsp;&nbsp;<input type="radio"
            name="ShortNews" value="4" <%if cstr(ShowInBursaNames&"") = "1" then response.write "checked"%> />&nbsp;חדשות ובורסת השמות</td>
      </tr>
      <tr>
        <td align="right" colspan="2">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="rtl">
            <tr>
              <td align="right" dir="rtl">
                <font size="2">כותרת גג</font>
              </td>
              <td align="right" dir="rtl" width="70">
                <font size="2">צבע</font>
              </td>
              <td align="right" dir="rtl" width="60">
                <font size="2">גודל</font>
              </td>
            </tr>
            <tr>
              <td align="right" dir="rtl"><input type="text" name="TopTitle" value="<%=replace(TopTitle & "","""","&quot;") %>" dir="rtl"
                  tabindex="1" style="width: 100%;"></td>
              <td align="right" dir="rtl" width="70"><select name="TopTitleColor" style="width: 60px; font-size: 10px;
                                font-family: Arial;">
                  <option value="" style="background-color: #003A6A; color: #FFFFFF;" <%if TopTitleColor="" then response.write "selected" %>>
                    כחול</option>
                  <option disabled="true" style="background-color:#FFFFFF; height:1px; font-size:1px;" Value="1">&nbsp;
                  </option>
                  <option value="#000000" style="background-color: #000000; color: #FFFFFF;" <%if TopTitleColor="#000000" then response.write "selected" %>>
                    שחור</option>
                  <option disabled="true" style="background-color:#FFFFFF; height:1px; font-size:1px;" Value="2">&nbsp;
                  </option>
                  <option value="#3F3F3F" style="background-color: #3F3F3F; color: #FFFFFF;" <%if TopTitleColor="#3F3F3F" then response.write "selected" %>>
                    אפור</option>
                  <option disabled="true" style="background-color:#FFFFFF; height:1px; font-size:1px;" Value="3">&nbsp;
                  </option>
                  <option value="#CC0000" style="background-color: #CC0000; color: #FFFFFF;" <%if TopTitleColor="#CC0000" then response.write "selected" %>>
                    אדום</option>
                </select></td>
              <td align="right" dir="ltr" width="60" style="font-size:11px;"><select name="TopTitleFontSize" style="width: 50px; height: 18px; font-size: 10px;
                                font-family: Arial;" dir="rtl">
                  <option value="" <%if TopTitleFontSize="" then response.write "selected" %>>13.5</option>
<%for i = 14 to 72 %>
<option value="<%=cstr(i) %>" <%if TopTitleFontSize=cstr(i) then response.write "selected" %>>
<%=cstr(i) %>
</option>
<%next %>
</select>PT</td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td align="right">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="rtl">
            <tr>
              <td align="right" dir="rtl">
                <font size="2">כותרת ראשית</font>
              </td>
              <td align="right" dir="rtl" width="70">
                <font size="2">צבע</font>
              </td>
              <td align="right" dir="rtl" width="60">
                <font size="2">גודל</font>
              </td>
            </tr>
            <tr>
              <td align="right" dir="rtl">
                <div dir="rtl">
<%
	    if session("imgLogo")<>"" then
                                    %>
<img name="imgLogo" src='/uploadimages/<%=session("imgLogo")%>' height="50"><br>
<%
	    end if
                                    %>
<textarea rows="3" name="fHead1" tabindex="2" style="width: 100%;">
<%=sHead1%>
</textarea>
                </div>
              </td>
              <td align="right" dir="rtl" width="70" valign="top"><select name="Head1Color" style="width: 60px; font-size: 10px;
                                font-family: Arial;">
                  <option value="" style="background-color: #000000; color: #FFFFFF;" <%if Head1Color="" then response.write "selected" %>>
                    שחור</option>
                  <option disabled="true" style="background-color:#FFFFFF; height:1px; font-size:1px;" Value="1">&nbsp;
                  </option>
                  <option value="#003A6A" style="background-color: #003A6A; color: #FFFFFF;" <%if Head1Color="#003A6A" then response.write "selected" %>>
                    כחול</option>
                  <option disabled="true" style="background-color:#FFFFFF; height:1px; font-size:1px;" Value="2">&nbsp;
                  </option>
                  <option value="#3F3F3F" style="background-color: #3F3F3F; color: #FFFFFF;" <%if Head1Color="#3F3F3F" then response.write "selected" %>>
                    אפור</option>
                  <option disabled="true" style="background-color:#FFFFFF; height:1px; font-size:1px;" Value="3">&nbsp;
                  </option>
                  <option value="#CC0000" style="background-color: #CC0000; color: #FFFFFF;" <%if Head1Color="#CC0000" then response.write "selected" %>>
                    אדום</option>
                </select></td>
              <td align="right" dir="ltr" valign="top" width="60" style="font-size:11px;"><select name="Head1FontSize"
                  style="width: 50px; height: 18px; font-size: 10px;
                                font-family: Arial;" dir="rtl">
                  <option value="" <%if Head1FontSize="" then response.write "selected" %>>24</option>
<%for i = 21 to 72 %>
<option value="<%=cstr(i) %>" <%if Head1FontSize=cstr(i) then response.write "selected" %>>
<%=cstr(i) %>
</option>
<%next %>
</select>PT</td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td valign="top" align="Right" colspan="2">
          <font size="2">:הנשמ תרתוכ</font>
        </td>
      </tr>
      <tr>
        <td align="Right" colspan="2">
          <div dir="rtl">
            <textarea rows="3" name="fHead2" tabindex="3" style="width: 100%;">
<%=sHead2%>
</textarea>
          </div>
        </td>
      </tr>
      <tr>
        <td align="right" colspan="2">
          <table border="0" cellpadding="0" cellspacing="0" dir="rtl"
            style="font-family:Arial; font-size:12px; margin-top:5px; margin-bottom:5px;">
            <tr>
              <td align="right" valign="top"><input type="checkbox" name="EventBoard" value="1" <%if EventBoard=1 then response.write "checked" %>
                  onclick="javascript:if (this.checked==1) { ShowEventBoardEdit.style.display = 'block'; } else { ShowEventBoardEdit.style.display = 'none'; } void(0);" />לוח
                אירועים</td>
              <td align="right" style="padding-right:10px;">
                <table border="0" cellpadding="2" cellspacing="2" id="ShowEventBoardEdit"
                  style="font-family:Arial; font-size:12px;display:<%if EventBoard=1 then response.write "block" else response.write "none"%>;">
                  <tr>
                    <td>תאריך התחלה</td>
                    <td>שעת התחלה</td>
                    <td>תאריך סיום משוער</td>
                    <td>שעת סיום משוערת</td>
                  </tr>
                  <tr>
                    <td style="white-space:nowrap;"><input type="text" name="EventStartDate" value="<%=EventStartDate %>"
                        readonly style="width:75px;" />
                      <SCRIPT LANGUAGE="JavaScript">
                      var cal3 = new CalendarPopup();
                      cal3.showYearNavigation();
                      </SCRIPT>
                      <A HREF="#"
                        onClick="cal3.select(document.news.EventStartDate,'anchor2','dd/MM/yyyy'); return false;"
                        NAME="anchor2" ID="anchor2">
                        <img src="images/calendar.gif" border=0 align="top">
                      </A>
                    </td>
                    <td><select name="EventStartM" style="font-family:Arial; font-size:12px;">
<%for i = 0 to 59 
                                    if i < 10 then%>
<option value="<%="0" & i %>" <%if right(EventStartTime & "", 2) = "0" & i then response.write "selected" %>>
<%="0" & i  %>
</option>
<%else %>
<option value="<%=i %>" <%if right(EventStartTime & "", 2) = i & "" then response.write "selected" %>>
<%=i  %>
</option>
                        <%end if %>
                        <%next %>
                      </select>:<select name="EventStartH" style="font-family:Arial; font-size:12px;">
<%for i = 0 to 23 
                                    if i < 10 then%>
<option value="<%="0" & i %>" <%if left(EventStartTime & "", 2) = "0" & i then response.write "selected" %>>
<%="0" & i  %>
</option>
<%else %>
<option value="<%=i %>" <%if left(EventStartTime & "", 2) = i & "" then response.write "selected" %>>
<%=i  %>
</option>
                        <%end if %>
                        <%next %>
                      </select></td>
                    <td style="white-space:nowrap;"><input type="text" name="EventEndDate" value="<%=EventEndDate %>"
                        readonly style="width:75px;" />
                      <SCRIPT LANGUAGE="JavaScript">
                      var cal4 = new CalendarPopup();
                      cal4.showYearNavigation();
                      </SCRIPT>
                      <A HREF="#"
                        onClick="cal4.select(document.news.EventEndDate,'anchor3','dd/MM/yyyy'); return false;"
                        NAME="anchor3" ID="anchor3">
                        <img src="images/calendar.gif" border=0 align="top">
                      </A>
                    </td>
                    <td><select name="EventEndM" style="font-family:Arial; font-size:12px;">
<%for i = 0 to 59 
                                    if i < 10 then%>
<option value="<%="0" & i %>" <%if right(EventEndTime & "", 2) = "0" & i then response.write "selected" %>>
<%="0" & i  %>
</option>
<%else %>
<option value="<%=i %>" <%if right(EventEndTime & "", 2) = i & "" then response.write "selected" %>>
<%=i  %>
</option>
                        <%end if %>
                        <%next %>
                      </select>:<select name="EventEndH" style="font-family:Arial; font-size:12px;">
<%for i = 0 to 23 
                                    if i < 10 then%>
<option value="<%="0" & i %>" <%if left(EventEndTime & "", 2) = "0" & i then response.write "selected" %>>
<%="0" & i  %>
</option>
<%else %>
<option value="<%=i %>" <%if left(EventEndTime & "", 2) = i & "" then response.write "selected" %>>
<%=i  %>
</option>
                        <%end if %>
                        <%next %>
                      </select></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr align="right">
        <td colspan="2" align="right">
          <table border="0" id="table2" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td align="right" width="20%">
                <font size="2">רתא</font>
              </td>
              <td align="right" width="20%">
                <font size="2">ל&quot;אוד</font>
              </td>
              <td align="right" width="20%">
                <font size="2">דסומ/המריפ</font>
              </td>
              <td align="right" width="20%" dir="rtl">
                <font size="2">תפקיד</font>
              </td>
              <td align="right" width="20%">
                <font size="2">:תאמ</font>
              </td>
            </tr>
            <tr>
              <td align="right" style="padding-right: 2px;">
<%if site_url="0" then%>
<input type="text" name="site_url" value="http://" style="width: 100%;" tabindex="8">
<%else%>
<input type="text" name="site_url" value="<%=site_url%>" style="width: 100%;" tabindex="7">
<%end if%>
</td>
              <td align="right" style="padding-right: 2px;">
                <input type="text" name="email" value="<%=sEmail%>" style="width: 100%;" tabindex="6">
              </td>
              <td align="right" style="padding-right: 2px;">
<% ifirm_name= replace(ifirm_name&"","""","&quot;")%>
<input type="text" name="firm_name" value="<%=ifirm_name%>" style="width: 100%;" tabindex="5"
                  dir="rtl">
              </td>
              <td align="right" style="padding-right: 2px;">
<% JobTitle= replace(JobTitle&"","""","&quot;")%>
<input type="text" name="JobTitle" value="<%=JobTitle%>" style="width: 100%;" tabindex="5"
                  dir="rtl">
              </td>
              <td align="right">
<% FWriter= replace(sWriter&"","""","&quot;")%>
<input type="text" name="fWriter" value="<%=FWriter%>" style="width: 100%;" tabindex="4" dir="rtl">
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td align="right" colspan="4">
          <table border="0" cellpadding="0" cellspacing="0">
<%if request("docid") <> "" and isnumeric(request("docid")) and request("subjectid") <> "" and isnumeric(request("subjectid")) then %>
<tr align="right">
              <td valign="top">
                <table border="0" cellspacing="1" cellpadding="1" dir="rtl">
                  <tr>
                    <td align="right" dir="rtl">
                      <b>
                        <font size="2" color="#000080">מקורות</font>
                      </b>
                      <font size="2">/ שם:</font>
                    </td>
                    <td align="right" dir="rtl">
                      <font size="2">&nbsp;כלי תקשורת:</font>
                    </td>
                    <td align="right" dir="rtl">
                      <font size="2">&nbsp;תאריך:</font>
                    </td>
                    <td align="right" dir="rtl">
                      <font size="2">&nbsp;קישור:</font>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                  </tr>
                  <tr>
                    <td align="right">
                      <font size="2">
                        <input type="text" name="txtSourceName" size="17" value="" dir="rtl" style="font-size: 8pt;
                                                                font-family: Arial">
                      </font>
                    </td>
                    <td align="right">
                      <font size="2">
                        <input type="text" name="txtSourceMedia" size="17" dir="rtl" value="" style="font-size: 8pt;
                                                                font-family: Arial">
                      </font>
                    </td>
                    <td align="right">
                      <font size="2">
                        <input type="text" name="txtSourceDate" size="7" dir="ltr" style="font-size: 8pt;
                                                                font-family: Arial" value="<%=date() %>">
                      </font>
                    </td>
                    <td align="right">
                      <font size="2">
                        <input type="text" name="txtSourceLink" size="12" dir="ltr" value="" style="font-size: 8pt;
                                                                font-family: Arial">
                      </font>
                    </td>
                    <td align="right">
                      <input type="button" value="הוסף חדש" onclick="javascript:AddNewSource();" style="font-size: 10px;
                                                            font-family: Arial;">&nbsp;
                    </td>
                    <td align="right">
                      <input type="button" value="עדכן" onclick="javascript:NewSource();" style="font-size: 10px;
                                                            font-family: Arial;">
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
<%end if %>
</table>
        </td>
      </tr>
      <tr align="right">
<% 
                                sImage= replace(sImage & "","""","&quot;")
                                if sImage="0" then
                                    sImage= ""
                                end if
                            %>
<td colspan="2">
          <table border="0" id="table4" dir="rtl" cellspacing="0" cellpadding="0" width="100%">
            <tr>
              <td>
                <a onclick="newwinGallery();" href="javascript:void(0);">
                  <font size="2" style="color:Blue;">תמונה ראשית</font>
                </a>
              </td>
              <td>
              </td>
              <td>
                <a onclick="newwinGallery2WithDesc('news','image_url2','SImageText2');" href="javascript:void(0);">
                  <font size="2" style="color:Blue;">מתחלפת 1</font>
                </a>
              </td>
              <td>
              </td>
              <td>
                <a onclick="newwinGallery2WithDesc('news','image_url3','SImageText3');" href="javascript:void(0);">
                  <font size="2" style="color:Blue;">מתחלפת 2</font>
                </a>
              </td>
              <td>
              </td>
              <td>
                <a onclick="newwinGallery2WithDesc('news','image_url4','SImageText4');" href="javascript:void(0);">
                  <font size="2" style="color:Blue;">מתחלפת 3</font>
                </a>
              </td>
              <td>
              </td>
            </tr>
            <tr>
              <td>
                <input type="text" name="image_url" value="<%=sImage%>" style="width:118px;" tabindex="9">
              </td>
              <td>&nbsp;</td>
              <td>
                <input type="text" name="image_url2" value="<%=sImage2%>" style="width:118px;" tabindex="9">
              </td>
              <td>&nbsp;</td>
              <td>
                <input type="text" name="image_url3" value="<%=sImage3%>" style="width:118px;" tabindex="9">
              </td>
              <td>&nbsp;</td>
              <td>
                <input type="text" name="image_url4" value="<%=sImage4%>" style="width:118px;" tabindex="9">
              </td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td colspan="8" height="2"></td>
            </tr>
            <tr align="right">
<%if SImageText <> "" then
                                                 SImageText = replace(SImageText,"""","&quot;")
                                                 if SImageText="0" then
                                                    SImageText= ""
                                                 end if
                                                 if lcase(SImageText)="null" then
                                                    SImageText= ""
                                                 end if
                                              end if   
                                              
                                             if SImageText2 <> "" then
                                                 SImageText2 = replace(SImageText2,"""","&quot;")
                                             end if   
                                             
                                             if SImageText3 <> "" then
                                                 SImageText3 = replace(SImageText3,"""","&quot;")
                                             end if
                                             
                                             if SImageText4 <> "" then
                                                 SImageText4 = replace(SImageText4,"""","&quot;")
                                             end if 
                                             
                                             if SImageText2="null" then
                                                fixedSImageText2=""
                                             else
                                                fixedSImageText2=SImageText2
                                             end if
                                             
                                             if SImageText3="null" then
                                                fixedSImageText3=""
                                             else
                                                fixedSImageText3=SImageText3
                                             end if
                                             
                                             if SImageText4="null" then
                                                fixedSImageText4=""
                                             else
                                                fixedSImageText4=SImageText4
                                             end if
                                        %>
<td><textarea rows="2" cols="16" name="SImageText" tabindex="12" width="118px"
                  style="font-size:10px;">
<%=SImageText%>
</textarea></td>
              <td>&nbsp;</td>
              <td><textarea rows="2" cols="16" name="SImageText2" tabindex="12" width="118px"
                  style="font-size:10px;">
<%=fixedSImageText2%>
</textarea></td>
              <td>&nbsp;</td>
              <td><textarea rows="2" cols="16" name="SImageText3" tabindex="12" width="118px"
                  style="font-size:10px;">
<%=fixedSImageText3%>
</textarea></td>
              <td>&nbsp;</td>
              <td><textarea rows="2" cols="16" name="SImageText4" tabindex="12" width="118px"
                  style="font-size:10px;">
<%=fixedSImageText4%>
</textarea></td>
              <td>&nbsp;</td>
            </tr>


            <tr>
              <td>
                <a onclick="newwinGallery2WithDesc('news','image_url5','SImageText5');" href="javascript:void(0);">
                  <font size="2" style="color:Blue;">מתחלפת 4</font>
                </a>
              </td>
              <td>
              </td>
              <td>
                <a onclick="newwinGallery2WithDesc('news','image_url6','SImageText6');" href="javascript:void(0);">
                  <font size="2" style="color:Blue;">מתחלפת 5</font>
                </a>
              </td>
              <td>
              </td>
              <td></td>
              <td>
              </td>
              <td></td>
              <td>
              </td>
            </tr>
            <tr>
              <td>
                <input type="text" name="image_url5" value="<%=sImage5%>" style="width:118px;" tabindex="9">
              </td>
              <td>&nbsp;</td>
              <td>
                <input type="text" name="image_url6" value="<%=sImage6%>" style="width:118px;" tabindex="9">
              </td>
              <td>&nbsp;</td>
              <td></td>
              <td>&nbsp;</td>
              <td></td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td colspan="8" height="2"></td>
            </tr>
            <tr align="right">
<%                                             
                                             if SImageText5 <> "" then
                                                 SImageText5 = replace(SImageText5,"""","&quot;")
                                             end if   
                                             
                                             if SImageText6 <> "" then
                                                 SImageText6 = replace(SImageText6,"""","&quot;")
                                             end if
                                             
                                             if SImageText5="null" then
                                                fixedSImageText5=""
                                             else
                                                fixedSImageText5=SImageText5
                                             end if
                                             
                                             if SImageText6="null" then
                                                fixedSImageText6=""
                                             else
                                                fixedSImageText6=SImageText6
                                             end if
                                        %>
<td><textarea rows="2" cols="16" name="SImageText5" tabindex="12" width="118px"
                  style="font-size:10px;">
<%=fixedSImageText5%>
</textarea></td>
              <td>&nbsp;</td>
              <td><textarea rows="2" cols="16" name="SImageText6" tabindex="12" width="118px"
                  style="font-size:10px;">
<%=fixedSImageText6%>
</textarea></td>
              <td>&nbsp;</td>
              <td></td>
              <td>&nbsp;</td>
              <td></td>
              <td>&nbsp;</td>
            </tr>
          </table>
        </td>
      </tr>
      <tr align="right">
        <td dir="rtl" colspan="2" align="right">
          <input type="checkbox" name="Special" value="1" <%if Trim(Special)="1" Then response.write "checked"%> tabindex="10">
          <font size="2">מיוחד(תמונה בראש העמוד) </font>
          <!--input type="checkbox" name="RelatedSpecial" value="1" <%if Trim(RelatedSpecial)="1" Then response.write "checked"%> tabindex="10" ><font size="2">מיוחד(מסמכים 
		קשורים) </font-->
          <input type="checkbox" name="SpecialAffair" value="1" <%if Trim(SpecialAffair)="1" Then response.write "checked"%> tabindex="10">
          <font size="2">מיוחד(קישורי פרשות)&nbsp;&nbsp; </font>
          <input type="checkbox" name="ShowChat" value="1" <%if Trim(ShowChat)="1" Then response.write "checked"%> tabindex="11">
          <font size="2">כותרת מבזק צבע אדום</font>
        </td>
      </tr>
      <!--
                        <tr align="right">
                            <td valign="top" align="right" colspan="2">
                            <br>
                            <input type="button" name="InsImage" value='צרף תמונה' size="35" onclick='UploadeImage();'
                                tabindex="11"></td>
                        </tr>
                        -->
      <!--כאן היה הכיתוב-->
      <tr>
        <td height="20"></td>
      </tr>
      <tr align="right">
<%
	    if SVidioFile <> "" then
	     SVidioFile = replace(SVidioFile,"""","&quot;")
	    end if   
                            %>
<td colspan="2" dir="RTL">
          <table border="0" width="100%" id="table6" cellspacing="0" cellpadding="0">
            <tr>
              <td align="right" width="70">
                <a onclick="newwinGalleryVideo();" href="javascript:void(0);">
                  <font size="2" style="color:Blue;">סרטון</font>
                </a>
              </td>
              <td align="right" width="120">
                <font size="2">
                  <input type="text" name="SVidioFile" style="width: 100%;" value="<%=SVidioFile%>" tabindex="6">
              </td>

              <td align="right"><input type="checkbox" name="LargeVideo" value="1" <%if cstr(LargeVideo&"") = "1" then response.write "checked" %> />
                <font size="2">הצג עליון</font>
              </td>
              <td align="right" width="70" style="padding-right:5px;">
                <font size="2">קוד סרטון</font>
              </td>
              <td align="right"><input type="text" name="VideoYouTube" style="width:130px;" value="<%=replace(VideoYouTube & "", """", "&quot;")%>" />
                </font>
              </td>
            </tr>
            <tr>
              <td style="font-size:12px;">כיתוב סרטון</td>
              <td colspan="4"><input type="text" name="videoText" style="width: 390px;" value="<%=videoText%>"
                  tabindex="6"></td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td colspan="2" dir="rtl">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td align="right">
                <font size="2" face="Arial">הערות:</font>
              </td>
            </tr>
            <tr>
              <td align="right">
                <textarea name="adminComments" style="width: 100%;" dir="rtl" rows="2">
<%=adminComments%>
</textarea>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr align="right">
        <td colspan="2" valign="top">
          <font size="2">:היוצרה השדחה תא בותכל אנ</font>
          <font color='red'>*</font>
        </td>
      </tr>
      <tr align="right">
        <td colspan="2" valign="top">
          <div dir="rtl">
            <textarea rows="20" style="width: 100%;" name="fText" tabindex="13">
<%=sText%>
</textarea>
          </div>
        </td>
      </tr>
      <tr>
        <td align="right" colspan="2" dir="rtl" height="30">
          <a href="javascript:void(0);" onclick="javascript:openMainLinks();"><b>
              <font size="2" color="#000080">
                קישורים ידיעה ראשית</font>
            </b></a>
        </td>
      </tr>
      <tr align="right">
        <td colspan="2" height="20">
          <hr size="1" />
        </td>
      </tr>
      <tr align="right">
        <td colspan="2" valign="top" dir="rtl" align="right">
          <b>
            <font size="2" color="#000080">מראה מקום</font>
          </b>
        </td>
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
                                                    font-family: Arial">
                </font>
              </td>
              <td width="10%" align="right" dir="rtl">
                <font size="2">&nbsp;קישור:</font>
              </td>
              <td align="right">
                <font size="2">
                  <input type="text" name="NewreferenceText" size="30" value="" dir="rtl" style="font-size: 8pt;
                                                    font-family: Arial">
                </font>
              </td>
              <td width="10%" align="right" dir="rtl">
                <font size="2">&nbsp;תאור:</font>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr align="right">
        <td colspan="2" height="20">
          <hr size="1" />
        </td>
      </tr>
      <tr align="right">
        <td colspan="2" valign="top">
          <table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
            <tr>
              <td align="right" dir="rtl">
                <b>
                  <font size="2" color="#000080">מידע נוסף (סוג הליך, מראה מקום, הערות וכו')</font>
                </b>
              </td>
            </tr>
            <tr>
              <td align="right">
                <textarea rows="3" name="moreInfo" cols="60" dir="rtl">
<%=moreInfo%>
</textarea>
              </td>
            </tr>
            <tr align="right">
              <td height="20">
                <hr size="1" />
              </td>
            </tr>
            <tr>
              <td align="right" dir="rtl">
                <b>
                  <font size="2" color="#000080">אודות המחבר</font>
                </b>
              </td>
            </tr>
            <tr>
              <td align="right">
                <textarea rows="3" name="authorText" cols="60" dir="rtl">
<%=authorText%>
</textarea>
              </td>
            </tr>
            <tr align="right">
              <td height="20">
                <hr size="1" />
              </td>
            </tr>
            <input type="hidden" name="leads" value="">
          </table>
        </td>
      </tr>
      <input type="hidden" name="bottomImg" value="">
      <tr align="right">
        <td colspan="2" valign="top" dir="rtl" align="right">
          <b>
            <font size="2" color="#000080">החלפת מילה בקישור</font>
          </b>
        </td>
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
                                                    font-family: Arial">
                </font>
              </td>
              <td width="10%" align="right" dir="rtl">
                <font size="2">&nbsp;קישור:</font>
              </td>
              <td align="right">
                <font size="2">
                  <input type="text" name="TextReplace" size="20" value="" dir="rtl" style="font-size: 8pt;
                                                    font-family: Arial">
                </font>
              </td>
              <td width="10%" align="right" dir="rtl">
                <font size="2">&nbsp;מילה:</font>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10">
          <hr size="1">
        </td>
      </tr>
<%

    sSqlText= " SELECT * FROM T_NewsText " & _
              " WHERE docId="& request("DocId") & " AND " & _
              " subjectId=" & request("subjectId") & _
              " ORDER BY OrderByNum , id " 
	'Response.Write sSqlText
	set objDButils=server.CreateObject("yoav.com_logic")
	
	set rs=objDButils.GetFastRecoredset(ssqltext)
		 
		sPartnerShow="<tr><td align=right colspan=2 dir=rtl><font size=""-1"" color=""#0051BD""><u>להלן כותרות המידע שהוזנו עד כה:</font></u></td></tr>"
		
		countText=1
		
		while not rs.eof 
		
				sPartnerShow=sPartnerShow & "<tr>"
				sPartnerShow=sPartnerShow & "<td dir=rtl align=right colspan=2><font size=2> "& countText &"</font>  "
				sPartnerShow=sPartnerShow & "<font size='-1' color='#000000' dir='rtl'>" & rs("RightSign") & " " & rs("head")&"</font>  - <a href='javascript:AddNewsTextLinks("&iSubjectId&","&iDocId &","&rs("id")&")'><font size=2 face=arial>קישורים</a>&nbsp;&nbsp;<a href=""javascript:void(0);"" onclick=""javascript:window.open('ManageTextTables.asp?Locid=2&subjectId="&isubjectId&"&DocId="&iDocId&"&recordId="&rs("id")&"','updateTables','width=600,height=500,scrollbars=yes');""><font size=2 face=arial>טבלאות</a>&nbsp;&nbsp;<a href=""javascript:void(0);"" onclick=""javascript:window.open('AddNewsText.asp?subjectId="&isubjectId&"&DocId="&iDocId&"&recordId="&rs("id")&"&from=main','update','width=600,height=500,scrollbars=yes');""><font size=2 face=arial>עדכן</a>&nbsp;&nbsp;<a href=""javascript:void(0);"" onclick=""javascript:window.open('AddNewsText.asp?subjectId="&isubjectId&"&DocId="&iDocId&"&recordId="&rs("id")&"&DeletePartner=1&from=dmain','update','width=600,height=500,scrollbars=yes');""><font size=2 face=arial>מחק</a>&nbsp;&nbsp;</td></tr>"
			
			countText=countText+1
			rs.movenext
		wend
		
		response.write sPartnerShow 

	if countit < 100 then
                        %>
<tr align="right">
        <td colspan="2" valign="top" align="right" dir="rtl">
          <a href="javascript:setOrder2(<%=iSubjectId%>,<%=iDocId %>)" tabindex="5" style="font-size: 12px;">
            קבע סדר</a>&nbsp;&nbsp; <a href="javascript:AddNewsText(<%=iSubjectId%>,<%=iDocId %>)"
            style="font-size: 12px;">הוסף/עדכן פיסקה</a><br>
          <font size="2" color="red">ניתן להוסיף עד שלושים פיסקאות*</font>
      </tr>
      <!--Add articleText-->
<%
	end if
                        %>
<tr>
        <td colspan="2">
          <hr size="1">
        </td>
      </tr>
    </table>
    <input type="hidden" name="entityvalidate" value="">
    <table border="0" cellspacing="2" cellpadding="2" width="480">
<%


	a = errorList.Items
	    ' Response.Write "<center><table>"
	'Response.Write errorList.Count&"ppp"
	 For fcount = 0 To errorList.Count -1
	
      'Response.Write fcount
	 response.Write "<tr><td align='right' colspan=4><FONT COLOR='red' size='-1'><STRONG>"
      Select Case a(fcount)
     
	  Case "fname"
			'Response.Write "נא הכנס שם פרטי "
			Response.Write translatelogical2visual("נא הכנס שם פרטי ",20)& "&nbsp;<img src='..\img\bull.gif'>"
	  Case "lname"
			'Response.Write "נא הכנס שם משפחה"
			Response.Write translatelogical2visual("נא הכנס שם משפחה",20)& "&nbsp;<img src='..\img\bull.gif'>"
	  Case "Address"
			'Response.Write "נא הכנס כתובת"
			Response.Write translatelogical2visual("נא הכנס כתובת",20)& "&nbsp;<img src='..\img\bull.gif'>"
     Case "phone"
			'Response.Write "נא הכנס מספר טלפון"
			Response.Write translatelogical2visual("נא הכנס מספר טלפון",20)& "&nbsp;<img src='..\img\bull.gif'>"
	  Case "Email"
         ' response.Write "נא הכנס דואר אלקטרוני"
   			Response.Write translatelogical2visual("נא הכנס דואר אלקטרוני",30)& "&nbsp;<img src='..\img\bull.gif'>"
	 ' Case "username"
			'Response.Write "נא הכנס שם משתמש"
			'Response.Write translatelogical2visual("נא הכנס שם משתמש",20)& "<br>"
	 ' case "pswd1"
			''Response.Write "נא  ודא סיסמה"
			'Response.Write translatelogical2visual("נא  ודא סיסמה",20)& "<br>"
		'case "pswd2"
		''Response.Write "jjj"
		'	Response.Write translatelogical2visual(" נא הכנס סיסמה",20)& "<br>"
		'case "hebrewvalue"		
		'	Response.Write translatelogical2visual("נא הכנס ערך בעברית",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "englishvalue"
			Response.Write translatelogical2visual("נא הכנס תפקיד",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "area"
			Response.Write translatelogical2visual("נא בחר תחום",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "zipcode"
			Response.Write translatelogical2visual("נא הכנס מיקוד נכון",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "duty"
			Response.Write translatelogical2visual("נא הכנס תפקיד/עיסוק",20)& "&nbsp;<img src='..\img\bull.gif'>"
		case "classcification2"
			Response.Write translatelogical2visual("נא הכנס תואר/מקצוע",20)& "&nbsp;<img src='..\img\bull.gif'>"
      end select 
      response.Write "</STRONG></font></td></tr>"    
    next       


                        %>
<input type="hidden" name="englishvalue" value="<%=sEngVlaue%>" size="62" tabindex="12">
      <tr>
        <td align="right" colspan="5" dir="rtl" style="font-size:12px;">
          <b>כותבים נוספים:</b>
        </td>
      </tr>
<%if request("docid") <> "" and isnumeric(request("docid")) and request("subjectid") <> "" and isnumeric(request("subjectid")) then 
                                strSqlDoc="select * from T_DocsWriters where docid=" & request("docid") & " and subjectId= " & request("subjectid")
			                    set MoreWrites=objDButils.GetFastRecoredset(strSqlDoc)
			                    currentWriterNum = 2
			                    if not MoreWrites.EOF then
			                        do while not MoreWrites.EOF
			                            select case currentWriterNum
			                                case 2
			                                    autoUserInfo2 = cstr(MoreWrites("autoUserInfo"))
			                                    if trim(request("email2"))<>"" then
			                                        email2=replace(request("email2"),"""","&quot;")
			                                    else
			                                        email2= replace(MoreWrites("email") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("fname2"))<>"" then
			                                        fname2=replace(request("fname2"),"""","&quot;")
			                                    else
			                                        fname2= replace(MoreWrites("fname") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("lname2"))<>"" then
			                                        lname2=replace(request("lname2"),"""","&quot;")
			                                    else
			                                        lname2= replace(MoreWrites("lname") & "","""","&quot;")
			                                    end if

			                                case 3
			                                    autoUserInfo3 = cstr(MoreWrites("autoUserInfo"))
                                                'email3= replace(MoreWrites("email") & "","""","&quot;")
                                                'fname3= replace(MoreWrites("fname") & "","""","&quot;")
                                                'lname3= replace(MoreWrites("lname") & "","""","&quot;")
                                                
                                                if trim(request("email3"))<>"" then
			                                        email3=replace(request("email3"),"""","&quot;")
			                                    else
			                                        email3= replace(MoreWrites("email") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("fname3"))<>"" then
			                                        fname3=replace(request("fname3"),"""","&quot;")
			                                    else
			                                        fname3= replace(MoreWrites("fname") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("lname3"))<>"" then
			                                        lname3=replace(request("lname3"),"""","&quot;")
			                                    else
			                                        lname3= replace(MoreWrites("lname") & "","""","&quot;")
			                                    end if
                                                
			                                case 4
			                                    autoUserInfo4 = cstr(MoreWrites("autoUserInfo"))
                                                'email4= replace(MoreWrites("email") & "","""","&quot;")
                                                'fname4= replace(MoreWrites("fname") & "","""","&quot;")
                                                'lname4= replace(MoreWrites("lname") & "","""","&quot;")
                                                
                                                if trim(request("email4"))<>"" then
			                                        email4=replace(request("email4"),"""","&quot;")
			                                    else
			                                        email4= replace(MoreWrites("email") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("fname4"))<>"" then
			                                        fname4=replace(request("fname4"),"""","&quot;")
			                                    else
			                                        fname4= replace(MoreWrites("fname") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("lname4"))<>"" then
			                                        lname4=replace(request("lname4"),"""","&quot;")
			                                    else
			                                        lname4= replace(MoreWrites("lname") & "","""","&quot;")
			                                    end if
			                                    
			                                case 5
			                                    autoUserInfo5 = cstr(MoreWrites("autoUserInfo"))
                                                'email5= replace(MoreWrites("email") & "","""","&quot;")
                                                'fname5= replace(MoreWrites("fname") & "","""","&quot;")
                                                'lname5= replace(MoreWrites("lname") & "","""","&quot;")
                                                
                                                if trim(request("email5"))<>"" then
			                                        email5=replace(request("email5"),"""","&quot;")
			                                    else
			                                        email5= replace(MoreWrites("email") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("fname5"))<>"" then
			                                        fname5=replace(request("fname5"),"""","&quot;")
			                                    else
			                                        fname5= replace(MoreWrites("fname") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("lname5"))<>"" then
			                                        lname5=replace(request("lname5"),"""","&quot;")
			                                    else
			                                        lname5= replace(MoreWrites("lname") & "","""","&quot;")
			                                    end if
			                                case 6
			                                    autoUserInfo6 = cstr(MoreWrites("autoUserInfo"))
                                                'email6= replace(MoreWrites("email") & "","""","&quot;")
                                                'fname6= replace(MoreWrites("fname") & "","""","&quot;")
                                                'lname6= replace(MoreWrites("lname") & "","""","&quot;")
                                                
                                                if trim(request("email6"))<>"" then
			                                        email6=replace(request("email6"),"""","&quot;")
			                                    else
			                                        email6= replace(MoreWrites("email") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("fname6"))<>"" then
			                                        fname6=replace(request("fname6"),"""","&quot;")
			                                    else
			                                        fname6= replace(MoreWrites("fname") & "","""","&quot;")
			                                    end if
                                                
                                                if trim(request("lname6"))<>"" then
			                                        lname6=replace(request("lname6"),"""","&quot;")
			                                    else
			                                        lname6= replace(MoreWrites("lname") & "","""","&quot;")
			                                    end if
			                            end select
			                            currentWriterNum = currentWriterNum + 1
			                        MoreWrites.movenext
			                        loop
			                    else
			                         autoUserInfo2 = request("autoUserInfo2")
                                     email2= replace(request("email2"),"""","&quot;")
                                      fname2= replace(request("fname2"),"""","&quot;")
                                       lname2= replace(request("lname2"),"""","&quot;")
			                         autoUserInfo3 = request("autoUserInfo3")
                                     email3= replace(request("email3"),"""","&quot;")
                                      fname3= replace(request("fname3"),"""","&quot;")
                                       lname3= replace(request("lname3"),"""","&quot;")
			                         autoUserInfo4 = request("autoUserInfo4")
                                     email4= replace(request("email4"),"""","&quot;")
                                      fname4= replace(request("fname4"),"""","&quot;")
                                       lname4= replace(request("lname4"),"""","&quot;")
			                         autoUserInfo5 = request("autoUserInfo5")
                                     email5= replace(request("email5"),"""","&quot;")
                                      fname5= replace(request("fname5"),"""","&quot;")
                                       lname5= replace(request("lname5"),"""","&quot;")
			                         autoUserInfo6 = request("autoUserInfo6")
                                     email6= replace(request("email6"),"""","&quot;")
                                      fname6= replace(request("fname6"),"""","&quot;")
                                       lname6= replace(request("lname6"),"""","&quot;")
			                    end if
                        else 
                                autoUserInfo2 = request("autoUserInfo2")
                                 email2= replace(request("email2"),"""","&quot;")
                                  fname2= replace(request("fname2"),"""","&quot;")
                                   lname2= replace(request("lname2"),"""","&quot;")
			                     autoUserInfo3 = request("autoUserInfo3")
                                 email3= replace(request("email3"),"""","&quot;")
                                  fname3= replace(request("fname3"),"""","&quot;")
                                   lname3= replace(request("lname3"),"""","&quot;")
			                     autoUserInfo4 = request("autoUserInfo4")
                                 email4= replace(request("email4"),"""","&quot;")
                                  fname4= replace(request("fname4"),"""","&quot;")
                                   lname4= replace(request("lname4"),"""","&quot;")
			                     autoUserInfo5 = request("autoUserInfo5")
                                 email5= replace(request("email5"),"""","&quot;")
                                  fname5= replace(request("fname5"),"""","&quot;")
                                   lname5= replace(request("lname5"),"""","&quot;")
			                     autoUserInfo6 = request("autoUserInfo6")
                                 email6= replace(request("email6"),"""","&quot;")
                                  fname6= replace(request("fname6"),"""","&quot;")
                                   lname6= replace(request("lname6"),"""","&quot;")
                       end if %>
<tr align="right">
        <td colspan="5" align="right">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td align="right" valign="top" dir="rtl">
                <select name="autoUserInfo" tabindex="102">
                  <option value="1" style="background-color: Yellow;" <%if cstr(autoUserInfo&"")="1" then response.write "selected" %>>
                    אוטומטי</option>
                  <option value="0" <%if cstr(autoUserInfo&"")="0" then response.write "selected" %>>ידני</option>
                </select>
              </td>
              <td colspan="2"></td>
              <td width="2">
<% sFname= replace(sFname,"""","&quot;")%>
</td>
              <td dir="rtl">
                <nobr>
                  <font color="red">*</font>
                  <font size="2">פרטי:</font>
                  <input type="text" name="fname" value="<%=sFname%>" size="10" tabindex="101">
                </nobr>
              </td>
<% sLname= replace(sLname,"""","&quot;")%>
<td width="2">
              </td>
              <td dir="rtl" nowrap>
                <nobr>
                  <font color="red">*</font>
                  <font size="2">משפחה:</font>
                  <input type="text" name="lname" value="<%=sLname%>" size="10" tabindex="100"
                    onchange="javascript:if (document.news.autoUserInfo[1].selected == true) { showDetails(document.news.lname.value); }""></nobr>
                                        </td>
                                        <td dir=" rtl">
                  <a href="javascript:window.open('showShortNames.asp','ShortNames','width=750,height=450,scrollbars=yes,resizable=yes'); void(0);"
                    style="color: Blue; font-size: 9pt;">קיצורים</a>
              </td>
            </tr>
            <tr>
              <td align="right" valign="top" dir="rtl">
                <select name="autoUserInfo2" tabindex="106">
                  <option value="1" style="background-color: Yellow;" <%if autoUserInfo2="1" then response.write "selected" %>>
                    אוטומטי</option>
                  <option value="0" <%if autoUserInfo2="0" then response.write "selected" %>>
                    ידני</option>
                </select>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font size="2">דוא"ל:</font>
                  <input type="text" name="email2" value="<%=email2%>" size="15" tabindex="105">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font color="red">*</font>
                  <font size="2">פרטי:</font>
                  <input type="text" name="fname2" value="<%=fname2%>" size="10" tabindex="104">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl" nowrap>
                <nobr>
                  <font color="red">*</font>
                  <font size="2">משפחה:</font>
                  <input type="text" name="lname2" value="<%=lname2%>" size="10" tabindex="103">
                </nobr>
              </td>
              <td dir="rtl">
                <a href="javascript:window.open('showShortNames.asp?CurrentWriter=2','ShortNames','width=750,height=450,scrollbars=yes,resizable=yes'); void(0);"
                  style="color: Blue; font-size: 9pt;">קיצורים</a>
              </td>
            </tr>
            <tr>
              <td align="right" valign="top" dir="rtl">
                <select name="autoUserInfo3" tabindex="110">
                  <option value="1" style="background-color: Yellow;" <%if autoUserInfo3="1" then response.write "selected" %>>
                    אוטומטי</option>
                  <option value="0" <%if autoUserInfo3="0" then response.write "selected" %>>
                    ידני</option>
                </select>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font size="2">דוא"ל:</font>
                  <input type="text" name="email3" value="<%=email3%>" size="15" tabindex="109">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font color="red">*</font>
                  <font size="2">פרטי:</font>
                  <input type="text" name="fname3" value="<%=fname3%>" size="10" tabindex="108">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl" nowrap>
                <nobr>
                  <font color="red">*</font>
                  <font size="2">משפחה:</font>
                  <input type="text" name="lname3" value="<%=lname3%>" size="10" tabindex="107">
                </nobr>
              </td>
              <td dir="rtl">
                <a href="javascript:window.open('showShortNames.asp?CurrentWriter=3','ShortNames','width=750,height=450,scrollbars=yes,resizable=yes'); void(0);"
                  style="color: Blue; font-size: 9pt;">קיצורים</a>
              </td>
            </tr>
            <tr>
              <td align="right" valign="top" dir="rtl">
                <select name="autoUserInfo4" tabindex="114">
                  <option value="1" style="background-color: Yellow;" <%if autoUserInfo4="1" then response.write "selected" %>>
                    אוטומטי</option>
                  <option value="0" <%if autoUserInfo4="0" then response.write "selected" %>>
                    ידני</option>
                </select>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font size="2">דוא"ל:</font>
                  <input type="text" name="email4" value="<%=email4%>" size="15" tabindex="113">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font color="red">*</font>
                  <font size="2">פרטי:</font>
                  <input type="text" name="fname4" value="<%=fname4%>" size="10" tabindex="112">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl" nowrap>
                <nobr>
                  <font color="red">*</font>
                  <font size="2">משפחה:</font>
                  <input type="text" name="lname4" value="<%=lname4%>" size="10" tabindex="111">
                </nobr>
              </td>
              <td dir="rtl">
                <a href="javascript:window.open('showShortNames.asp?CurrentWriter=4','ShortNames','width=750,height=450,scrollbars=yes,resizable=yes'); void(0);"
                  style="color: Blue; font-size: 9pt;">קיצורים</a>
              </td>
            </tr>
            <tr style="display:none;">
              <td align="right" valign="top" dir="rtl">
                <select name="autoUserInfo5" tabindex="118">
                  <option value="1" style="background-color: Yellow;" <%if autoUserInfo5="1" then response.write "selected" %>>
                    אוטומטי</option>
                  <option value="0" <%if autoUserInfo5="0" then response.write "selected" %>>
                    ידני</option>
                </select>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font size="2">דוא"ל:</font>
                  <input type="text" name="email5" value="<%=email5%>" size="15" tabindex="117">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font color="red">*</font>
                  <font size="2">פרטי:</font>
                  <input type="text" name="fname5" value="<%=fname5%>" size="10" tabindex="116">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl" nowrap>
                <nobr>
                  <font color="red">*</font>
                  <font size="2">משפחה:</font>
                  <input type="text" name="lname5" value="<%=lname5%>" size="10" tabindex="115">
                </nobr>
              </td>
              <td dir="rtl">
                <a href="javascript:window.open('showShortNames.asp?CurrentWriter=5','ShortNames','width=750,height=450,scrollbars=yes,resizable=yes'); void(0);"
                  style="color: Blue; font-size: 9pt;">קיצורים</a>
              </td>
            </tr>
            <tr style="display:none;">
              <td align="right" valign="top" dir="rtl">
                <select name="autoUserInfo6" tabindex="122">
                  <option value="1" style="background-color: Yellow;" <%if autoUserInfo6="1" then response.write "selected" %>>
                    אוטומטי</option>
                  <option value="0" <%if autoUserInfo6="0" then response.write "selected" %>>
                    ידני</option>
                </select>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font size="2">דוא"ל:</font>
                  <input type="text" name="email6" value="<%=email6%>" size="15" tabindex="121">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl">
                <nobr>
                  <font color="red">*</font>
                  <font size="2">פרטי:</font>
                  <input type="text" name="fname6" value="<%=fname6%>" size="10" tabindex="120">
                </nobr>
              </td>
              <td width="2">
              </td>
              <td dir="rtl" nowrap>
                <nobr>
                  <font color="red">*</font>
                  <font size="2">משפחה:</font>
                  <input type="text" name="lname6" value="<%=lname6%>" size="10" tabindex="119">
                </nobr>
              </td>
              <td dir="rtl">
                <a href="javascript:window.open('showShortNames.asp?CurrentWriter=6','ShortNames','width=750,height=450,scrollbars=yes,resizable=yes'); void(0);"
                  style="color: Blue; font-size: 9pt;">קיצורים</a>
              </td>
            </tr>
          </table>
        </td>
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
          <font size="2">:זוחמ</font>
        </td>
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
          <font size="2">:הירוגטק</font>
        </td>
      </tr>
      <tr align="right">
        <td dir="LTR" valign="top">
          <font face='Arial (hebrew)'>
<%
			if cstr(request("area"))="" then
		    	showAreaid=split(iArea&"",",")
		  	else
		    	iArea=cstr(request("area")) 
				showAreaid=split(iArea,",")
		  	end if 
			call ShowAreaupdate(showAreaid,18)
  
                                    %>
</font>
        </td>
        <td valign="top">
          <font size="2">:םוחת</font>
          <font color='red'>*</font>
        </td>
        <td dir="LTR">
          <font face='Arial (hebrew)'>
            <%
	    '  if cstr(request("classcification2"))="" and isnull(iclasscification2)then
	         
	    '     call  ShowClassCombo("","classcification2","size=3 MULTIPLE tabindex=17")
	    '  else
	      
	    '   if cstr(request("classcification2"))="" then 
		'      showAreaid=split(iclasscification2,",")
		'	  call ShowClassUpdate(showAreaid,17)
		  
		'   else
		    
		'     iArea=cstr(request("classcification2")) 
		'	 showAreaid=split(iArea,",")
			
		'	 call ShowClassUpdate(showAreaid,17)
		  
		'   end if 
		' end if         
                                    %>
            <%
			if cstr(request("Affair_id"))="" then 
		    	showAffairID=split(Affair_id,",")
		  	else
		    	Affair_id=cstr(request("Affair_id")) 
				showAffairID=split(Affair_id,",")
		  	end if 
			call ShowAffairUpdate(showAffairID,18)


                                    %>
          </font>
        </td>
        <td align="right">
          <font size="2">:השרפ</font>
          <!--font size="2">:עוצקמ/ראות</font><font color='red'>*</font-->
        </td>
      </tr>
      <tr align="right" style="display: none;">
        <td dir="LTR">
          <input type="text" name="fax" value="<%=sFax%>" size="15" tabindex="20">
        </td>
        <td>
          <font size="2">:סקפ</font>
        </td>
        <td dir="LTR">
          <input type="text" name="phone" value="<%=sPhone%>" size="15" tabindex="19">
        </td>
        <td>
          <font size="2">:ןופלט</font>
        </td>
      </tr>
      <tr align="right" style="display: none;">
<% sSettling= replace(sSettling&"","""","&quot;")%>
<td dir="RTL">
          <input type="" name="settling" value="<%=sSettling%>" size="15" tabindex="24">
        </td>
        <td>
          <font size="2">:בושי</font>
        </td>
<% sAddress= replace(sAddress&"","""","&quot;")%>
<td dir="RTL">
          <input type="" name="address" value="<%=sAddress%>" size="15" tabindex="23">
        </td>
        <td>
          <font size="2">:תבותכ</font>
        </td>
      </tr>
      <tr align="right">
        <td dir="RTL">
          <input type="text" name="updatedate" value="<%=Date()%>" size="15" tabindex="26" readonly>
        </td>
        <td dir="rtl">
          <font size="2">תאריך עדכון:</font>
        </td>
        <td dir="ltr" valign="bottom" style="white-space:nowrap; font-family:Arial; font-size:12px;">
          <input type="" name="zipcode" value="<%=iZipcode%>" size="15" tabindex="25" style="display: none;"><input
            type="hidden" name="fDate" value="<%=sDate%>">
<%=day(YoavTime) & "/" & month(YoavTime) & "/" & year(YoavTime) & " " & formatdatetime(YoavTime, 4)   %>
</td>
        <td valign="top" dir="rtl">
          <font size="2">תאריך פרסום:</font>
          <!--
                                <font size="2">:דוקימ</font><font color='red'>*</font>-->
        </td>
      </tr>
      <tr align="right">
        <td valign="top" dir="rtl" style="font-size:12px;"><input type="checkbox" name="ShowFastNews" value="1"
            <%if cstr(ShowFastNews)="1" then %>checked<%end if %> />&nbsp;הפעל משיכת מבזקים:&nbsp;<input type="text"
            name="CountFastNews" value="<%=CountFastNews %>" style="width:50px; border:1px solid #cc0000;" /><br /><input
            type="text" name="StartFastNews" value="<%if StartFastNews=0 then
                                response.write "8"
                                else
                                response.write StartFastNews
                                end if %>" style="width:30px; border:1px solid #cc0000;" />
          שעות אחורה, <input type="text" name="EndFastNews" value="<%if EndFastNews=0 then
                                response.write "14"
                                else
                                response.write EndFastNews
                                end if %>"
            style="width:30px; border:1px solid #cc0000;" /> שעות קדימה
        </td>
        <td>
          <div dir="RTL">
          </div>
        </td>
        <td align="RIGHT">
          <input type="hidden" name="fHour" value="<%=sTime%>">
          <input type="button" value="העבר" name="updaterec1" onclick="javascript:ChangeHidden(3)" tabindex="60"><select
            name="permission" tabindex="27" dir="rtl">
            <option value="2" <%if request("status")="2" then response.write "selected"%>>מסמכים
              מאושרים</option>
            <option value="1" <%if request("status")="1" then response.write "selected"%>>לא מוצגים</option>
            <option value="6" <%if request("status")="6" then response.write "selected"%>>תיקיית מערכת</option>
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
        <td align="right">
          &nbsp;<font size="2">היקיתל הרבעה</span> </font>
        </td>
      </tr>
      <!--TR>
	  <TD COLSPAN = "4" align = "right">
	  
	  
	  
	  <font size="2">
      <INPUT type="checkbox"  name="AddToMailingList" <%if not MailListrs.eof and not trim(sEmail)="na" then Response.Write "checked"%> tabindex=30 value="ON">
	 רוויד תמישרל ףרצ	 
	 
	 
	 
	 </TD>
	</TR-->
      <tr>
        <td align="right" colspan="2" dir="rtl">
          <input type="checkbox" value="1" name="NoFindReplace" <%if cstr(NoFindReplace&"") = "1" then response.write "checked" %> />&nbsp;<font size="2">בטל החלפת/תיקון
            מילים</font>
        </td>
        <td align="right" dir="rtl" colspan="2">
          <input type="checkbox" value="1" name="ShowStarsOnLeft" <%if cstr(ShowStarsOnLeft&"") = "1" then response.write "checked" %> />&nbsp;<font size="2">הצג מי ומי
            בצד שמאל</font>
        </td>

      </tr>
      <tr>
        <td align="right" style="direction:rtl; white-space:nowrap;">
          <font size="2">אשכול מסמכים אוטומטית:</font>&nbsp;<input type="text" name="AutoAffairDocs"
            value="<%=AutoAffairDocs %>" style="width:50px; border:1px solid #cc0000;" />
        </td>
        <td align="right" colspan="2" style="direction:rtl;">
          <font size="2">מסמך בתשלום:</font>&nbsp;<select name="OnlyMembers">
            <option value="0" <%if cstr(OnlyMembers & "") = "0" then Response.Write "selected"%>>פתוח</option>
            <option value="1" <%if cstr(OnlyMembers & "") = "1" then Response.Write "selected"%>>מסמך בתשלום</option>
            <option value="2" <%if cstr(OnlyMembers & "") = "2" then Response.Write "selected"%>>ללא הגבלה</option>
          </select>
        </td>

      </tr>
<% if cint(request("status")) = 3 then %>
<tr align="center">
        <td colspan="4" dir="rtl">
          <br>
          <br>
<%if cstr(request("status"))="5" then %>
<input type="button" value="מחיקה" name="notgood" onclick="javascript:ChangeHiddenDel(5)" tabindex="31"
            style="color:red;" disabled="disabled">
<%else %>
<input type="button" value="מחיקה" name="notgood" onclick="javascript:ChangeHiddenDel(5)" tabindex="31">
<%end if %>
</td>
      </tr>
<% else %>
<tr align="center">
        <td colspan="4" dir="rtl">
          <br>
          <br>
          <input type="button" value="עדכון" name="updaterec" onclick="javascript:ChangeHidden(3)" tabindex="31"
            style="<%if cstr(request("status"))="3" then%>color:red;<%else%>color:black;<%end if%>">
          <%
                                    if Session("UserId")<>3 Then
                                    %>
          <%if cstr(request("status"))="2" then %>
          <input type="button" value="מאושרים" name="allright" onclick="javascript:ChangeHidden(1)" tabindex="32"
            style="color:red;" disabled="disabled">
<%else%>
<input type="button" value="מאושרים" name="allright" onclick="javascript:ChangeHidden(1)" tabindex="32">
          <%end if%>
          <%
	                                end if
                                    %>
          <%if cstr(request("status"))="1" then %>
          <input type="button" value="לא מוצגים" name="updaterec" onclick="javascript:ChangeHidden(4)" tabindex="33"
            style="color:red;" disabled="disabled">
<%else%>
<input type="button" value="לא מוצגים" name="updaterec" onclick="javascript:ChangeHidden(4)" tabindex="33">
          <%end if%>


          <%if cstr(request("status"))="6" then %>
          <input type="button" value="תקיית מערכת" name="updaterec2" onclick="ChangeHidden('6')" tabindex="34"
            style="color:red;" disabled="disabled">
<%else%>
<input type="button" value="תקיית מערכת" name="updaterec2" onclick="ChangeHidden('6')" tabindex="34">
          <%end if %>

          <%
		                            if Session("PrivateWaitingListPage")<>"" and IsNumeric(Session("PrivateWaitingListPage")) Then
		                                if cstr(Session("PrivateWaitingListPage"))=cstr(request("status")) then
                                    %>
          <input type="button" value="תקייה אישית" name="updaterec1" onclick="ChangeHidden('<%=Session("PrivateWaitingListPage")%>')"
            tabindex="35" style="color:red;" disabled="disabled">
<%
                                        else
                                        %>
<input type="button" value="תקייה אישית" name="updaterec1" onclick="ChangeHidden('<%=Session("PrivateWaitingListPage")%>')"
            tabindex="35">
          <%
                                        end if
		                            end if
                                    %>

          &nbsp;&nbsp;&nbsp;&nbsp;
          <%if cstr(request("status"))="5" then %>
          <input type="button" value="מחיקה" name="notgood" onclick="javascript:ChangeHiddenDel(5)" tabindex="36"
            style="color:red;" disabled="disabled">
<%else %>
<input type="button" value="מחיקה" name="notgood" onclick="javascript:ChangeHiddenDel(5)" tabindex="36">
<%end if %>
</td>
      </tr>
<%end if%>
</table>
    </FORM>
    <br>
    <div align="center">
      <a href="main.asp">ישארה טירפתל הרזח</a>
    </div>
    </td>
    <!-------------------- END MAIN TABLES ------------------------------------------->
    <!-------------------------------------- RIGHT TABLES ------------------------------------------->
    <td width="118" background="img/right_bg1.gif" bgcolor="#F5F5F5" valign="top" align="center"
      style="border-left: 1px solid #C0C0C0;">
      <table cellspacing="0" cellpadding="0" border="0">
        <tr>
          <!------------- News1 search ---------------->
          <td align="center" valign="top">
          </td>
          <!------------- end News1 search ---------------->
        </tr>
        <tr>
          <!------------- links table ------------------->
          <td align="right" bgcolor="#F0F0F0">
          </td>
          <!-------------------------------------- END RIGHT TABLES ------------------------------------------->
        </tr>
      </table>
  </div>
</body>
</html>
<%
function check(errorList)

'** validation of the new user form **
	'subjectid=con_news
	userId= session("UserId")
	status=cint(request("status")) 'when the document first goes in the db the status is 0
	permission=cstr(request("permission")) 'permission to update the doc comes from the managers only
	priority =0 'only managers give priority for homepage show
	parentDocID=request("parentDocID")
'====================check for news template=====goes into T_news====
	
	bottomImg=request("bottomImg")
	'TopTitle=Replace(request("TopTitle"),"'","''")
	TopTitle=request("TopTitle")

    TopTitleColor=request("TopTitleColor") & ""
                TopTitleFontSize=request("TopTitleFontSize") & ""
                Head1Color=request("Head1Color") & ""
                Head1FontSize=request("Head1FontSize") & ""
	
	'** head 1**
	'head1=replace(replace(request("fHead1"),"'",""),"-","&macaf")
	'head1=replace(request("fHead1"),"'","''")
	head1=request("fHead1")
	'Response.Write "pppppp"
	if head1="" then
		sErrorStr1="יש להכניס כותרת ראשית "
	else
		sErrorStr1=""
	end if
	'** head2 **
	'head2=replace(replace(request("fHead2"),"'",""),"-","&macaf")
		'head2=replace(request("fHead2"),"'","''")
		head2=request("fHead2")
	'** writername**
	writername=replace(request("fWriter"),"'","''")
	'** image**
	'upload
	SVidioFile=replace(request("SVidioFile"),"'","''")
	if SVidioFile="" then
	SVidioFile=""
	end if
	
	VideoYouTube=replace(request("VideoYouTube") & "", "'", "''")
	
	'** image**
	'upload
	SImageText=replace(request("SImageText"),"'","''")
	if SImageText="" then
	SImageText=""
	end if
	
	SImageText2=replace(request("SImageText2"),"'","''")
	if SImageText2="" then
	    SImageText2=""
	end if
	
	SImageText3=replace(request("SImageText3"),"'","''")
	if SImageText3="" then
	SImageText3=""
	end if
	
	SImageText4=replace(request("SImageText4"),"'","''")
	if SImageText4="" then
	SImageText4=""
	end if
	
	SImageText5=replace(request("SImageText5"),"'","''")
	if SImageText5="" then
	SImageText5=""
	end if
	
	SImageText6=replace(request("SImageText6"),"'","''")
	if SImageText6="" then
	SImageText6=""
	end if
	
	'** image**
	'upload
	image=replace(request("image_url"),"'","''")
	if image="" then
		image="null"
	end if
	
	image2=replace(request("image_url2"),"'","''")
	
	image3=replace(request("image_url3"),"'","''")
	
	image4=replace(request("image_url4"),"'","''")
	
	image5=replace(request("image_url5"),"'","''")
	
	image6=replace(request("image_url6"),"'","''")


    
    ShowFastNews=request("ShowFastNews") & ""
    CountFastNews=request("CountFastNews") & ""
    StartFastNews=request("StartFastNews") & ""
    EndFastNews=request("EndFastNews") & ""
    if ShowFastNews="1" then
        ShowFastNews=1
        if CountFastNews<>"" and isnumeric(CountFastNews) then
        else
            CountFastNews=50
        end if
        if StartFastNews<>"" and isnumeric(StartFastNews) then
        else
            StartFastNews=2
        end if
        if EndFastNews<>"" and isnumeric(EndFastNews) then
        else
            EndFastNews=14
        end if
    else
        ShowFastNews=0
        CountFastNews=50
        StartFastNews=2
        EndFastNews=14
    end if
	
	
	
	'** userdate**
	userdate=replace(request("fDate"),"'","''")
	if userdate="" then
		userdate=date()
	end if
	'========================================================================
	fdate=request("fDate")
	    if isdate(request("fdate")) then
		CreateDate = day(request("fdate")) & "/" & month(request("fDate")) & "/" & mid(year(request("fDate")),3,2)
		else
		CreateDate=request("fDate")
		end if
	if fdate="" then fDate=date() 'day(date) & "/" & month(date) & "/" & year(date)
	'** userhour**
	userhour=replace(request("fHour"),"'","''")
	'========================================================================
	if userhour="" then
		userhour=time()
	end if
	'** text**
	text=request("ftext")
	'text=replace(request("fText"),"'","''")
	if text="" then
		sErrorStr2="יש להכניס טקסט"
	else
		sErrorStr2=""
	end if
	'** linkname**
	linkname="aaaa"'replace(request("fLinkName"),"'","''")
	'** linkurl **
	linkurl="aaaa"'replace(request("fLinkUrl"),"'","''")
	'***permission**
	'permission=request("permission")

'============= check for entity===== goes into T_entity===============
	'into db :id,subject_id1
	'** hebrew dscr **
	'	hebrewvalue=replace(request("hebrewvalue"),"'","''")
		hebrewvalue=head1
		'if hebrewvalue="" then
		'Response.Write hebrewvalue
		'errorList.Add "6","hebrewvalue"
		'end if
	'**english dscr**

		englishvalue=replace(request("englishvalue"),"'","''")
		if englishvalue="" then
		   englishvalue= "-"
		'errorList.Add "7","englishvalue"
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
	duty =cint(2) 'request("duty")
	duty2=request("duty2")
	'if duty2="" then
	'    errorList.add "2","duty"
	'end if 
	'** firm name**
		firm_name=replace(request("firm_name"),"'","")
		JobTitle = replace(request("JobTitle") & "", "'", "''")
	'** year of birth**
		birthyear="0" 'request("birthyear")
		
	'** decease_date**
		decease_date="null" 'replace(request("decease"),"'","")

	'** birthday **		
		birthday="null" 'request("day")&"/"& request("month")&"/"& request("year")
		
	'** sex**
		sex="0" 'request("sex")
		
	'**nation1**
		nation1="0" 'request("nation1")
		
	'**nation2**
		nation2="0" 'request("nation2")
		
	'** Address **
	address=replace(request("address"),"'","''")
	'if Address="" then
		'errorList.Add "3","Address"
	'end if
	'**settling**
		settling=replace(request("settling"),"'","''")
	'**zipcode**
		zipcode=request("zipcode")
		if not IsNumeric(zipcode) then
			errorList.Add "10","zipcode"
		end if
	'** Telephone **
	phone=replace(request("phone"),"'","''")
	if phone="" then
		'errorList.Add "4","phone"
	end if
	'**fax**
	fax=replace(request("fax"),"'","''")
	
	'** E-mail **
	email=replace(request("email"),"'","''")
		'if email = "" or not email="na" then
		'	if not instr(1,email,"@")<> 0 Then
		'	errorList.Add "5","Email"	
		'	end if
		'end if
	'**site_url**
	   site_url=trim(request("site_url"))
	   if site_url="http://" or site_url="" then
		  site_url="0"
	   end if  
		  'request("site_url")
	'entry date comes automatic

	'**area**
		area=request("area")
		if area="" then
			'errorList.Add "8","area"
			area="כללי"
		end if	
		
		Affair_id=Request("Affair_id")
		Category_id=Request("Category_id")
		Province_id=Request("Province_id")
		moreinfo=request("moreinfo")
		referenceTEXT=""
		referenceURL=""
		referenceTEXT2=""
		referenceURL2=""
		referenceTEXT3=""
		referenceURL3=""
		videoText=request("videoText")
		
		LargeVideo = request("LargeVideo")
	if LargeVideo="" or not isnumeric(LargeVideo) then
	    LargeVideo=0
	end if
		
		
		if request("NoFindReplace") = "1" then
		    NoFindReplace = 1
		else
		    NoFindReplace = 0
		 end if  
		
		authorText=request("authorText")
		leads=request("leads")

	'===========================================================	
		
	'** classcification**
		classcification=cint(3)'request("classcification")
	   ' classcification2=request("classcification2")
	    'FLG
	  ' if classcification2="" then
	'	errorList.Add "11","classcification2"
	 '  end if
	    
		'	classcification="null"
		'end if
	'**deal type**
		dealtype="0"'request("dael_type")
		
	'**person_total_salary**
		person_total_salary="0"  'request("person_salary")
		
	'**person_total_date**	
		person_salary_date="null " 'request("salary_date")
		
	'**firm_total_balance**
		firm_total_balance="0" 'request("firm_balance")
		
	'**firm_balance_date**
		firm_balance_date="null"  'request("firm_balance")
		
	'**firm_total_income**
		firm_total_income="0" 'request("firm_income")
		
	'**firm_income_date**
		firm_income_date="null " 'request("firm_income")
		
	'**firm_managers_total**
		firm_managers_total="0"'replace(request("five_managers"),"'","")
		
	'update_date,doc_id
	'**eduction**
		'education=request("education")
		'if isnull(request("education")) then
		education="0"
		'end if
		placeToBeBorn="0"
		if isdate(request("updatedate")) then 
		strDate=month(request("updatedate")) & "/" & day(request("updatedate")) & "/" & year(request("updatedate"))
		strNow=strDate & " " & time()
		'strdate=date()
		else
		strDate=month(date) & "/" & day(date) & "/" & year(date)
		strNow=strDate & " " & time()
		end if
		'strNow=
	'=======end validation============	
	If errorList.Count >0 or sErrorStr<>"" Then
	        check = False
   		Exit Function
	else

		    
	Call UpdateHtmlStatus(request("docid"),request("subjectid"),0)
	
	sql=  " SELECT     docID, subjectID "
	sql= sql & " FROM   T_Linked_Documents "
	sql= sql & " WHERE  (linkedDocID = " & request("docid") & ") AND (linkedsubjectID = "&request("subjectid")&") "
	
	set objDbUtils=createObject("yoav.DbUtils")
	set linkedRS=objDbUtils.GetFastRecoredset(sql)
	do until linkedRS.EOF
		Call UpdateHtmlStatus(linkedRS("docid"),linkedRS("subjectid"),0)
		linkedRS.MoveNext
	Loop
	set objDbUtils=nothing
	
	UnLockDoc request("docid"),request("subjectid")
	Call WriteHistoryLog(request("docid"),request("subjectid"))
	
        if cstr(NoFindReplace&"") = "1" then
       
            head1=Replace(head1&"","'","''")
		 head2 =Replace(head2&"","'","''")
		 text =Replace(text&"","'","''")
		 moreInfo=Replace(moreInfo&"","'","''")
		authorText=Replace(authorText&"","'","''")
		 leads=replace(leads&""&"","'","''")
		 hebrewvalue=Replace(hebrewvalue&"","'","''")
		 TopTitle=Replace(TopTitle&"","'","''")  
       
       else
       
        head1=Replace(findReplace(head1),"'","''")
		 head2 =Replace(findReplace(head2),"'","''")
		 text =Replace(findReplace(text),"'","''")
		 moreInfo=Replace(findReplace(moreInfo),"'","''")
		authorText=Replace(findReplace(authorText),"'","''")
		 leads=replace(findReplace(leads)&"","'","''")
		 hebrewvalue=Replace(findReplace(hebrewvalue),"'","''")
		 TopTitle=Replace(findReplace(TopTitle),"'","''") 
       
       end if 

		 
		 
		 
		 referenceTEXT=Replace(referenceTEXT,"'","''")
		 referenceTEXT2=Replace(referenceTEXT2,"'","''")
		 referenceTEXT3=Replace(referenceTEXT3,"'","''")
		 videoText=Replace(videoText,"'","''")
		
		LargeVideo = request("LargeVideo")
	    if LargeVideo="" or not isnumeric(LargeVideo) then
	        LargeVideo=0
	    end if 
		 
		 
		 
		 
		 
		 
	
       	getPage = "https://www.news1.co.il/CloudFlareAPI/DeleteCache.aspx?url=https://www.news1.co.il/Archive/" & GenerateName(request("docid"),request("subjectid"),"")
        GetFastRemoteUrl(getPage) 
       	getPage = "https://www.news1.co.il/CloudFlareAPI/DeleteCache.aspx?url=https://m.news1.co.il/Archive/" & GenerateName(request("docid"),request("subjectid"),"")
		GetFastRemoteUrl(getPage) 
 





select case Request("RecAction")
	case "1"

	'אושר
		'		Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),2,request("parentDocId"),"T_news"
		
		status=2
	'	Com.UpdateT_newsPermission request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText
				Call UpdateT_newsPermission( request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText)
			'		 '===str for insert to entity table====
			StrEntity=request("subjectid") & spe & hebrewvalue & spe & englishvalue & spe & fname & spe & lname 
				StrEntity=StrEntity &  spe & duty & spe & firm_name & spe & birthyear  & spe & decease_date & spe & birthday & spe & sex & spe & nation1 & spe & nation2 
			StrEntity=StrEntity &  spe & Address & spe & settling & spe & zipcode & spe &  Phone & spe & fax & spe & Email '19
			StrEntity=StrEntity &  spe & site_url & spe & fdate & spe & area & spe & classcification & spe & dealtype '24
				StrEntity=StrEntity &  spe & person_total_salary & spe & person_salary_date & spe & firm_total_balance & spe & firm_balance_date '28
				StrEntity=StrEntity &  spe & firm_total_income & spe & firm_income_date & spe & firm_managers_total & spe & strNow & spe &  request("docid") & spe & education & spe & placeToBeBorn
				StrEntity=StrEntity &  spe & duty2 & spe & classcification2 & spe & replace(Affair_id&"","'","''")  & spe & Request("skill_id")  & spe & Request("skill_Parent_id")  & spe & Lang & spe &category_id & spe & province_id & spe & SearchWords
			StrEntity=StrEntity+   spe & moreInfo & spe & authorText
			
			'			Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),status,request("parentDocId"),"T_news"
			Call UpdateT_newsPermissionByDocAndSubject (request("docid"),request("subjectid"),status,request("parentDocId"),"T_news")
			
			set logoConn=createObject("adodb.connection")
			logoConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"


            sql = "update DocsTranslates set htmlStatus=0 where docid = " & request("docid") & " and subjectid = " & request("subjectid") 
			logoConn.execute(sql)
			
			sql = "delete from T_docsArea where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)
			sql = "delete from T_docsAffair where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)

            if request("subjectid")="35" then
                Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & ",35,153)"
			    logoConn.execute(Isql)
            end if
			
			
				if request("area") <> "" then
					areaId = getAreaIds(request("area"))
					areaIdArr = split(areaId, ",")
					for i = 0 to Ubound(areaIdArr)
						if areaIdArr(i) <> "" and isnumeric(areaIdArr(i)) then
                        if request("subjectid")="35" and cstr(areaIdArr(i) & "") = "153" then
                        else
							Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & "," & request("subjectid") & "," & areaIdArr(i) & ")"
							logoConn.execute(Isql)
                        end if
						end if
					next
				end if
				

                FirstAffair_id = 0

				if request("Affair_id") <> "" then
					AffairId = getAffairIds(request("Affair_id"))
					AffairIdArr = split(AffairId, ",")
					for i = 0 to Ubound(AffairIdArr)
						if AffairIdArr(i) <> "" and isnumeric(AffairIdArr(i)) then
                        if FirstAffair_id=0 then
                                FirstAffair_id = clng(AffairIdArr(i))
                            end if
							Isql = "insert into T_docsAffair (docid,subjectid,Affairid) values (" & request("docid") & "," & request("subjectid") & "," & AffairIdArr(i) & ")"
							logoConn.execute(Isql)
						end if
					next
				end if

                ShortNews = request("ShortNews")
		if ShortNews="" or not isnumeric(ShortNews) then
		    ShortNews = 0
		end if 

			 if cstr(ShortNews&"") = "1" and FirstAffair_id > 0 then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN T_news as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)

                sql = "update t_articles set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN t_articles as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)
                end if 

                if cstr(ShortNews&"") = "1" then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShortNews=1 and ShowInShortNews=0 and YoavTime>=dateadd(hh,-45,getdate())"
                logoConn.execute(sql)
                end if
			
			
			Special=Request("Special")
			RelatedSpecial=Request("RelatedSpecial")
			SpecialAffair=Request("SpecialAffair")
			adverStatus=request("adverStatus")
			adverStatus12Home=request("adverStatus12Home")
		
			if cstr(adverStatus&"") = "1" then
	            adverStatus = 1
            else
                adverStatus = 0
            end if   
         if cstr(adverStatus12Home&"") = "1" then
	            adverStatus12Home = "1"
            else
                adverStatus12Home = "0"
            end if   

    'new ddl
    invoiceStatus=request("invoiceStatus")
    if cstr(invoiceStatus&"") = "1" then
	    invoiceStatus = 1
    else if cstr(invoiceStatus&"") = "2" then
	    invoiceStatus = 2
     else if cstr(invoiceStatus&"") = "3" then
    invoiceStatus = 3
     else if cstr(invoiceStatus&"") = "4" then
    invoiceStatus = 4
     end if 
     end if 
     end if 
    end if 
			
			ShortNews = request("ShortNews")
			ShowInShortNews=0
			ShowInBursaNames=0
		    if ShortNews="" or not isnumeric(ShortNews) then
		        ShortNews = 0
		    else
		        if cstr(ShortNews&"") = "2" then
		             ShowInShortNews=1
		             ShortNews=0
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "1" then
		             ShowInShortNews=0
		             ShortNews=1
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "3" then
		             ShowInShortNews=0
		             ShortNews=2
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "4" then
		             ShowInShortNews=0
		             ShortNews=0
                     ShowInBursaNames=1
		        else
		             ShowInBursaNames=0
		             ShowInShortNews=0
		             ShortNews=0
		        end if  
		    end if
		      
			
			
			logoSql="update t_news set logo='"&session("imgLogo")&"',ShortNews=" & ShortNews & ",ShowInShortNews=" & ShowInShortNews & ",ShowInBursaNames=" & ShowInBursaNames & ", UpdatedByUserName='"&session("UpdatedByUserName")&"',TopTitle='"&TopTitle&"',bottomImg='"&bottomImg&"',Special='"&Special&"',RelatedSpecial='"&RelatedSpecial&"',SpecialAffair='"&SpecialAffair&"', "
			logoSql=logoSql &"linkReplace='',TextReplace='',image2='"&image2&"',image3='"&image3&"',image4='"&image4&"',videoText='"&videoText&"',LargeVideo="&LargeVideo&",VideoYouTube='" & VideoYouTube & "',ImageText2='" & SImageText2 & "',ImageText3='" & SImageText3 & "',ImageText4='" & SImageText4 & "', "  
			logoSql=logoSql &"referenceTEXT='" &referenceTEXT& "',referenceURL='"&referenceURL&"',referenceTEXT2='" &referenceTEXT2& "',referenceURL2='"&referenceURL2&"',referenceTEXT3='" &referenceTEXT3& "',referenceURL3='"&referenceURL3&"', leads='"&leads&"',adverStatus="&adverStatus&", invoiceStatus="& invoiceStatus 
			logoSql=logoSql & ",NoFindReplace=" & NoFindReplace & ",TopTitleColor='" & TopTitleColor & "',TopTitleFontSize='" & TopTitleFontSize & "',Head1Color='" & Head1Color & "',Head1FontSize='" & Head1FontSize & "'"
			logoSql=logoSql &",image5='"&image5&"',ImageText5='"&SImageText5&"',image6='"&image6&"',ImageText6='"&SImageText6&"',ShowFastNews=" & ShowFastNews & ",CountFastNews=" & CountFastNews & ",StartFastNews=" & StartFastNews & ",EndFastNews=" & EndFastNews & ",YoavTime=getdate() where docid="& request("docid")& " and subjectID="&request("subjectID")
			
			logoConn.execute logoSql
			
			logoSql="update t_entity set adminComments='" & replace(request("adminComments"), "'", "''") & "'"
			if request("ShowChat") = "1" then
			    logoSql = logoSql & ",ShowChat=1"
			else
			    logoSql = logoSql & ",ShowChat=0"
			end if
			if request("ShowStarsOnLeft") = "1" then
			    logoSql = logoSql & ",ShowStarsOnLeft=1"
			else
			    logoSql = logoSql & ",ShowStarsOnLeft=0"
			end if
		    logoSql = logoSql & ",OnlyMembers=" & request("OnlyMembers")
            if request("AutoAffairDocs") <> "" and isnumeric(request("AutoAffairDocs")) then
                logoSql = logoSql & ",AutoAffairDocs=" & request("AutoAffairDocs")
            else
                logoSql = logoSql & ",AutoAffairDocs=0"
            end if
			logoSql = logoSql & " where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
			logoConn.execute logoSql
			
			'logoConn.close
			'set logoConn=nothing
			
			
			
			
			
			'com.UpdateEntity(StrEntity)
			Call UpdateEntity(StrEntity)
			
		'set aConn=createObject("adodb.connection")
		'aConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
		
	if request("status")="2" or Request("RecAction") ="1" then
		
		txtToTags =  replace(TopTitle & "","''", "'") & chr(10) & replace(head1 & "","''", "'") & chr(10) & replace(head2 & "","''", "'") & chr(10) & replace(text & "","''", "'")
	   sqlP = "select head,head2,mtext from T_newsText where docid=" & request("docid") & " and subjectid=" & request("subjectid")
	   set rsPP = logoConn.execute(sqlP)
	   do while not rsPP.EOF
	    txtToTags = txtToTags & chr(10) & rsPP("head") & chr(10) & rsPP("head2") & chr(10) & rsPP("mtext")
	    rsPP.movenext
	    loop
	    set rsPP = nothing
	    
	
		 UpdateArticleTagsGlobal request("docid"),request("subjectid"),txtToTags
	end if
		
		if request("autoUserInfo") = "1" then
		    sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(fname & "", "'", "''") & "' and Link_LastName='" & replace(lname & "", "'", "''") & "'"
		    set rsA = logoConn.Execute(sql)
		    if not rsA.EOF then
		        email = rsA("email") 
		        firmName = rsA("firmName") 
		        website = rsA("website") 
		        street = rsA("street") 
		        city = rsA("city") 
		        if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
		            zipcode = 0
		        else
		            zipcode = rsA("zip")
		        end if  
		        phone = rsA("phone") 
		        fax = rsA("fax") 
		        JobTitle = rsA("duty")
		         
		        set rsU = server.CreateObject("ADODB.recordset")
		        sql = "select * from T_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		        rsU.open sql, logoConn, 1, 3
		        rsU("autoUserInfo") = 1
		        rsU("email") = email
		        rsU("firm_name") = firmName
		        rsU("JobTitle") = JobTitle
		        if fname="globes" then
                else
		        rsU("site_url") = website
                end if
		        rsU("address") = street
		        rsU("settling") = city
		        rsU("zipcode") = zipcode
		        rsU("phone") = phone
		        rsU("fax") = fax
		        rsU.update
		        set rsU = nothing  
		    end if
		    set rsA = nothing 
		else
		    sql = "update T_entity set autoUserInfo=0,JobTitle='" & JobTitle & "' where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		    logoConn.Execute(sql) 
		end if  
		
		sql = "delete T_DocsWriters where subjectid=" & request("subjectid") & " and docid=" & request("docid")
		logoConn.execute(sql)
		
		
		if request("fname2") <> "" and request("lname2") <> "" then
		    if request("autoUserInfo2") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname2"), "'", "''") & "' and Link_LastName='" & replace(request("lname2"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = request("email2")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname3") <> "" and request("lname3") <> "" then
		    if request("autoUserInfo3") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname3"), "'", "''") & "' and Link_LastName='" & replace(request("lname3"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = request("email3")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname4") <> "" and request("lname4") <> "" then
		    if request("autoUserInfo4") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname4"), "'", "''") & "' and Link_LastName='" & replace(request("lname4"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = request("email4")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname5") <> "" and request("lname5") <> "" then
		    if request("autoUserInfo5") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname5"), "'", "''") & "' and Link_LastName='" & replace(request("lname5"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = request("email5")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    
	    if request("fname6") <> "" and request("lname6") <> "" then
		    if request("autoUserInfo6") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname6"), "'", "''") & "' and Link_LastName='" & replace(request("lname6"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = request("email6")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if


        
            set rsUE = server.createobject("ADODB.Recordset")
            sql = "select * from t_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
            rsUE.open sql, logoConn, 1, 3

            if request("EventBoard") = "1" then
                rsUE("EventBoard") = 1
                 EventBoardStartDate = request("EventStartDate") & " " & request("EventStartH") & ":" & request("EventStartM")
                 EventBoardEndDate = request("EventEndDate") & " " & request("EventEndH") & ":" & request("EventEndM")

                 if isdate(EventBoardStartDate) then
                 rsUE("EventBoardStartDate") = cdate(EventBoardStartDate)
                 else
                 rsUE("EventBoardStartDate") = null
                 end if
                 if isdate(EventBoardEndDate) then
                    rsUE("EventBoardEndDate") = cdate(EventBoardEndDate)
                 else
                    rsUE("EventBoardEndDate") = null
                 end if
            else
                rsUE("EventBoard") = 0
                rsUE("EventBoardStartDate") = null
                rsUE("EventBoardEndDate") = null
            end if

            rsUE.update
            set rsUE = nothing
		
        logoConn.close
		set logoConn = nothing
			
			
			'update Affairs if selected
			'if trim(request("Affair_id")&"") <> "" and len(request("Affair_id")&"") > 1 then
			
			'	set objDButils = server.createobject("yoav.dbutils")
				
			'	arrAffair = split(request("Affair_id")&"",",")
			'	for i = 0 to Ubound(arrAffair)
			'		if trim(arrAffair(i)&"") <> "" then
						
			'			set rs=objDButils.GetFastRecoredset("SELECT id,dscr FROM T_area WHERE status=2 and dscr = '"&trim(arrAffair(i)&"")&"'")
			'			if not rs.EOF then
			'				call CreateOneHTMLPage("https://www.news1.co.il/createHtml/HtmlCreateAreaByID.aspx?areaid=" & rs("id"))
			'			end if
			'			set rs = nothing
			'		end if
			'	next
			'	set objDButils = nothing
				
			'end if
			
		
		
		call UpdateT_newsPermissionByDocAndSubject(request("docid"),request("subjectid"),2,request("parentDocId"),"T_news")
		com.sendEmailForConform email
    SendArticleByWriterNewsletter request("docid"),request("subjectid"),fname,lname
		Response.Redirect "currentpages.asp"
	case "2"  
		' לא אושר
		'			Com.DeleteT_newsPermissionByDocAndSubject request("docid"),request("subjectid")
		call DeleteT_newsPermissionByDocAndSubject(request("docid"),request("subjectid"))
		Response.Redirect  "main.asp"
	case "3" 
    
            set logoConn_1=createObject("adodb.connection")
			logoConn_1.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
			if cstr(adverStatus12Home&"") = "1" then
	            adverStatus12Home = "1"
            else
	            adverStatus12Home = "0"
            end if 
			logoSql="update t_news set adverStatus12Home='1' where docid="& request("docid")& " and subjectID="&request("subjectID")
			
			logoConn_1.execute logoSql
			
	        logoConn_1.close
    
    
    
    
    
       
			'עדכון
			'Response.Write "3"
			'Response.Write area
			If status=0 Then
				status=1
			Else
				if permission="1" then
					status=1
				Else
					if Cstr(status)="5" Then
						status=5
					else
						Status=Cint(permission)
					end if
				end if
			End If   
			'	Com.UpdateT_newsPermission request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText
				Call UpdateT_newsPermission( request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText)
			'		 '===str for insert to entity table====
			StrEntity=request("subjectid") & spe & hebrewvalue & spe & englishvalue & spe & fname & spe & lname 
				StrEntity=StrEntity &  spe & duty & spe & firm_name & spe & birthyear  & spe & decease_date & spe & birthday & spe & sex & spe & nation1 & spe & nation2 
			StrEntity=StrEntity &  spe & Address & spe & settling & spe & zipcode & spe &  Phone & spe & fax & spe & Email '19
			StrEntity=StrEntity &  spe & site_url & spe & fdate & spe & area & spe & classcification & spe & dealtype '24
				StrEntity=StrEntity &  spe & person_total_salary & spe & person_salary_date & spe & firm_total_balance & spe & firm_balance_date '28
				StrEntity=StrEntity &  spe & firm_total_income & spe & firm_income_date & spe & firm_managers_total & spe & strNow & spe &  request("docid") & spe & education & spe & placeToBeBorn
				StrEntity=StrEntity &  spe & duty2 & spe & classcification2 & spe & replace(Affair_id&"","'","''")  & spe & Request("skill_id")  & spe & Request("skill_Parent_id")  & spe & Lang & spe &category_id & spe & province_id & spe & SearchWords
			StrEntity=StrEntity+   spe & moreInfo & spe & authorText
			
			'			Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),status,request("parentDocId"),"T_news"
			Call UpdateT_newsPermissionByDocAndSubject (request("docid"),request("subjectid"),status,request("parentDocId"),"T_news")
			
			set logoConn=createObject("adodb.connection")
			logoConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
			
			sql = "update DocsTranslates set htmlStatus=0 where docid = " & request("docid") & " and subjectid = " & request("subjectid") 
			logoConn.execute(sql)


			sql = "delete from T_docsArea where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)
			sql = "delete from T_docsAffair where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)

            if request("subjectid")="35" then
                Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & ",35,153)"
			    logoConn.execute(Isql)
            end if
			
            FirstAffair_id = 0

		
				if request("area") <> "" then
					areaId = getAreaIds(request("area"))
					areaIdArr = split(areaId, ",")
					for i = 0 to Ubound(areaIdArr)
						if areaIdArr(i) <> "" and isnumeric(areaIdArr(i)) then
                        if request("subjectid")="35" and cstr(areaIdArr(i) & "") = "153" then
                        else
							Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & "," & request("subjectid") & "," & areaIdArr(i) & ")"
							logoConn.execute(Isql)
                        end if
						end if
					next
				end if
				
				if request("Affair_id") <> "" then
					AffairId = getAffairIds(request("Affair_id"))
					AffairIdArr = split(AffairId, ",")
					for i = 0 to Ubound(AffairIdArr)
						if AffairIdArr(i) <> "" and isnumeric(AffairIdArr(i)) then
                        if FirstAffair_id=0 then
                                FirstAffair_id = clng(AffairIdArr(i))
                            end if
							Isql = "insert into T_docsAffair (docid,subjectid,Affairid) values (" & request("docid") & "," & request("subjectid") & "," & AffairIdArr(i) & ")"
							logoConn.execute(Isql)
						end if
					next
				end if

			
                ShortNews = request("ShortNews")
		if ShortNews="" or not isnumeric(ShortNews) then
		    ShortNews = 0
		end if 

			 if cstr(ShortNews&"") = "1" and FirstAffair_id > 0 then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN T_news as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)

                sql = "update t_articles set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN t_articles as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)
                end if 

                if cstr(ShortNews&"") = "1" then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShortNews=1 and ShowInShortNews=0 and YoavTime>=dateadd(hh,-45,getdate())"
                logoConn.execute(sql)
                end if
			
			
			Special=Request("Special")
			RelatedSpecial=Request("RelatedSpecial")
			SpecialAffair=Request("SpecialAffair")
			adverStatus=request("adverStatus")
			
		
			if cstr(adverStatus&"") = "1" then
	            adverStatus = 1
            else
                adverStatus = 0
            end if
    
            if cstr(adverStatus12Home&"") = "1" then
	            adverStatus12Home = "1"
            else
	            adverStatus12Home = "0"
            end if  
    
     	'new ddl
    invoiceStatus=request("invoiceStatus")
    	if cstr(invoiceStatus&"") = "1" then
	    invoiceStatus = 1
    else if cstr(invoiceStatus&"") = "2" then
	    invoiceStatus = 2
     else if cstr(invoiceStatus&"") = "3" then
    invoiceStatus = 3
     else if cstr(invoiceStatus&"") = "4" then
    invoiceStatus = 4
     end if 
     end if 
     end if 
    end if   
			
			ShortNews = request("ShortNews")
			ShowInShortNews=0
			ShowInBursaNames=0
		    if ShortNews="" or not isnumeric(ShortNews) then
		        ShortNews = 0
		    else
		        if cstr(ShortNews&"") = "2" then
		             ShowInShortNews=1
		             ShortNews=0
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "1" then
		             ShowInShortNews=0
		             ShortNews=1
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "3" then
		             ShowInShortNews=0
		             ShortNews=2
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "4" then
		             ShowInShortNews=0
		             ShortNews=0
                     ShowInBursaNames=1
		        else
		             ShowInBursaNames=0
		             ShowInShortNews=0
		             ShortNews=0
		        end if
		    end if
		      
			
			
			logoSql="update t_news set logo='"&session("imgLogo")&"',ShortNews=" & ShortNews & ",ShowInShortNews=" & ShowInShortNews & ",ShowInBursaNames=" & ShowInBursaNames & ", UpdatedByUserName='"&session("UpdatedByUserName")&"',TopTitle='"&TopTitle&"',bottomImg='"&bottomImg&"',Special='"&Special&"',RelatedSpecial='"&RelatedSpecial&"',SpecialAffair='"&SpecialAffair&"', "
			logoSql=logoSql &"linkReplace='',TextReplace='',image2='"&image2&"',image3='"&image3&"',image4='"&image4&"',videoText='"&videoText&"',LargeVideo="&LargeVideo&",VideoYouTube='" & VideoYouTube & "',ImageText2='" & SImageText2 & "',ImageText3='" & SImageText3 & "',ImageText4='" & SImageText4 & "', "  
			logoSql=logoSql &"referenceTEXT='" &referenceTEXT& "',referenceURL='"&referenceURL&"',referenceTEXT2='" &referenceTEXT2& "',referenceURL2='"&referenceURL2&"',referenceTEXT3='" &referenceTEXT3& "',referenceURL3='"&referenceURL3&"', leads='"&leads&"',adverStatus="&adverStatus&",invoiceStatus="& invoiceStatus 
			logoSql=logoSql & ",NoFindReplace=" & NoFindReplace & ",TopTitleColor='" & TopTitleColor & "',TopTitleFontSize='" & TopTitleFontSize & "',Head1Color='" & Head1Color & "',Head1FontSize='" & Head1FontSize & "'"
			logoSql=logoSql &",image5='"&image5&"',ImageText5='"&SImageText5&"',image6='"&image6&"',ImageText6='"&SImageText6&"',ShowFastNews=" & ShowFastNews & ",CountFastNews=" & CountFastNews & ",StartFastNews=" & StartFastNews & ",EndFastNews=" & EndFastNews & " where docid="& request("docid")& " and subjectID="&request("subjectID")
			
			logoConn.execute logoSql
			
			logoSql="update t_entity set adminComments='" & replace(request("adminComments"), "'", "''") & "'"
			if request("ShowChat") = "1" then
			    logoSql = logoSql & ",ShowChat=1"
			else
			    logoSql = logoSql & ",ShowChat=0"
			end if
			if request("ShowStarsOnLeft") = "1" then
			    logoSql = logoSql & ",ShowStarsOnLeft=1"
			else
			    logoSql = logoSql & ",ShowStarsOnLeft=0"
			end if
		    logoSql = logoSql & ",OnlyMembers=" & request("OnlyMembers")
            if request("AutoAffairDocs") <> "" and isnumeric(request("AutoAffairDocs")) then
                logoSql = logoSql & ",AutoAffairDocs=" & request("AutoAffairDocs")
            else
                logoSql = logoSql & ",AutoAffairDocs=0"
            end if
			logoSql = logoSql & " where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
			logoConn.execute logoSql
			
			'logoConn.close
			'set logoConn=nothing
			
			
			
			
			
			'com.UpdateEntity(StrEntity)
			Call UpdateEntity(StrEntity)
			
		'set aConn=createObject("adodb.connection")
		'aConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
		if request("status")="2" then
		
		txtToTags =  replace(TopTitle & "","''", "'") & chr(10) & replace(head1 & "","''", "'") & chr(10) & replace(head2 & "","''", "'") & chr(10) & replace(text & "","''", "'")
	   sqlP = "select head,head2,mtext from T_newsText where docid=" & request("docid") & " and subjectid=" & request("subjectid")
	   set rsPP = logoConn.execute(sqlP)
	   do while not rsPP.EOF
	    txtToTags = txtToTags & chr(10) & rsPP("head") & chr(10) & rsPP("head2") & chr(10) & rsPP("mtext")
	    rsPP.movenext
	    loop
	    set rsPP = nothing
		
       
		 UpdateArticleTagsGlobal request("docid"),request("subjectid"),txtToTags
	end if
	
		
		if request("autoUserInfo") = "1" then
		    sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(fname & "", "'", "''") & "' and Link_LastName='" & replace(lname & "", "'", "''") & "'"
		    set rsA = logoConn.Execute(sql)
		    if not rsA.EOF then
		        email = rsA("email") 
		        firmName = rsA("firmName") 
		        website = rsA("website") 
		        street = rsA("street") 
		        city = rsA("city") 
		        if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
		            zipcode = 0
		        else
		            zipcode = rsA("zip")
		        end if  
		        phone = rsA("phone") 
		        fax = rsA("fax") 
		        JobTitle = rsA("duty")
		         
		        set rsU = server.CreateObject("ADODB.recordset")
		        sql = "select * from T_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		        rsU.open sql, logoConn, 1, 3
		        rsU("autoUserInfo") = 1
		        rsU("email") = email
		        rsU("firm_name") = firmName
		        rsU("JobTitle") = JobTitle
		        if fname="globes" then
                else
		        rsU("site_url") = website
                end if
		        rsU("address") = street
		        rsU("settling") = city
		        rsU("zipcode") = zipcode
		        rsU("phone") = phone
		        rsU("fax") = fax
		        rsU.update
		        set rsU = nothing  
		    end if
		    set rsA = nothing 
		else
		    sql = "update T_entity set autoUserInfo=0,JobTitle='" & JobTitle & "' where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		    logoConn.Execute(sql) 
		end if  
		
		sql = "delete T_DocsWriters where subjectid=" & request("subjectid") & " and docid=" & request("docid")
		logoConn.execute(sql)
		
		
		if request("fname2") <> "" and request("lname2") <> "" then
		    if request("autoUserInfo2") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname2"), "'", "''") & "' and Link_LastName='" & replace(request("lname2"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = request("email2")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname3") <> "" and request("lname3") <> "" then
		    if request("autoUserInfo3") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname3"), "'", "''") & "' and Link_LastName='" & replace(request("lname3"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = request("email3")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname4") <> "" and request("lname4") <> "" then
		    if request("autoUserInfo4") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname4"), "'", "''") & "' and Link_LastName='" & replace(request("lname4"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = request("email4")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname5") <> "" and request("lname5") <> "" then
		    if request("autoUserInfo5") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname5"), "'", "''") & "' and Link_LastName='" & replace(request("lname5"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = request("email5")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    
	    if request("fname6") <> "" and request("lname6") <> "" then
		    if request("autoUserInfo6") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname6"), "'", "''") & "' and Link_LastName='" & replace(request("lname6"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = request("email6")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if

         set rsUE = server.createobject("ADODB.Recordset")
            sql = "select EventBoard,EventBoardStartDate,EventBoardEndDate from t_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
            rsUE.open sql, logoConn, 1, 3

            if request("EventBoard") = "1" then
                rsUE("EventBoard") = 1
                 EventBoardStartDate = request("EventStartDate") & " " & request("EventStartH") & ":" & request("EventStartM")
                 EventBoardEndDate = request("EventEndDate") & " " & request("EventEndH") & ":" & request("EventEndM")

                 if isdate(EventBoardStartDate) then
                 rsUE("EventBoardStartDate") = cdate(EventBoardStartDate)
                 else
                 rsUE("EventBoardStartDate") = null
                 end if
                 if isdate(EventBoardEndDate) then
                    rsUE("EventBoardEndDate") = cdate(EventBoardEndDate)
                 else
                    rsUE("EventBoardEndDate") = null
                 end if
            else
                rsUE("EventBoard") = 0
                rsUE("EventBoardStartDate") = null
                rsUE("EventBoardEndDate") = null
            end if

            rsUE.update
            set rsUE = nothing
		
		logoConn.close
			set logoConn=nothing
			
			
			'update Affairs if selected
			'if trim(request("Affair_id")&"") <> "" and len(request("Affair_id")&"") > 1 then
			
			'	set objDButils = server.createobject("yoav.dbutils")
				
			'	arrAffair = split(request("Affair_id")&"",",")
			'	for i = 0 to Ubound(arrAffair)
			'		if trim(arrAffair(i)&"") <> "" then
						
			'			set rs=objDButils.GetFastRecoredset("SELECT id,dscr FROM T_area WHERE status=2 and dscr = '"&trim(arrAffair(i)&"")&"'")
			'			if not rs.EOF then
			'				call CreateOneHTMLPage("https://www.news1.co.il/createHtml/HtmlCreateAreaByID.aspx?areaid=" & rs("id"))
			'			end if
			'			set rs = nothing
			'		end if
			'	next
			'	set objDButils = nothing
				
			'end if
			set logoConn_1=createObject("adodb.connection")
			logoConn_1.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
    			

            if Request.form("adverStatus12Home") <> "" then
                logoSql="update t_news set adverStatus12Home='1' where docid="& request("docid")& " and subjectID="&request("subjectID")
            else
                logoSql="update t_news set adverStatus12Home='0' where docid="& request("docid")& " and subjectID="&request("subjectID")
            end if
    			
			
			logoConn_1.execute logoSql
			
	        logoConn_1.close
		
			
			if request("fromOpenWin") <> "YES" then
				docOpener=cstr(request("docOpener"))
				'if request("status")=0 then
				'	Response.Redirect "newdetails.asp"
				'end if
				
				if request("status")=0 then
					Response.Redirect "newdetails.asp"
				end if
				if request("status")=1 then
			'		Response.Redirect "waitinglistpage.asp?orderby=ArticlesAndNews"
	Response.Redirect "main.asp"
				end if
				if request("status")=2 then
					Response.Redirect "Main.asp" '"currentpages.asp"
				elseif request("status")=3 then
					Response.Redirect "Main.asp" '"currentpages.asp"
				elseif request("status")=5 then
				    Response.Redirect "Main.asp" '"CurrentPages.asp?status=5"
				elseif request("status")=6 then
				    Response.Redirect "Main.asp" '"CurrentPages.asp?status=6"
				else
					Response.Redirect "Main.asp" '"PrivateWaitingListPage.asp"
				end if
			end if
	Case "4"
	'מאושר לא להצגה
	'Response.Write "4"
	'		    Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),1,request("parentDocId"),"T_news"
	
	
	status=1
	'	Com.UpdateT_newsPermission request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText
				Call UpdateT_newsPermission( request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText)
			'		 '===str for insert to entity table====
			StrEntity=request("subjectid") & spe & hebrewvalue & spe & englishvalue & spe & fname & spe & lname 
				StrEntity=StrEntity &  spe & duty & spe & firm_name & spe & birthyear  & spe & decease_date & spe & birthday & spe & sex & spe & nation1 & spe & nation2 
			StrEntity=StrEntity &  spe & Address & spe & settling & spe & zipcode & spe &  Phone & spe & fax & spe & Email '19
			StrEntity=StrEntity &  spe & site_url & spe & fdate & spe & area & spe & classcification & spe & dealtype '24
				StrEntity=StrEntity &  spe & person_total_salary & spe & person_salary_date & spe & firm_total_balance & spe & firm_balance_date '28
				StrEntity=StrEntity &  spe & firm_total_income & spe & firm_income_date & spe & firm_managers_total & spe & strNow & spe &  request("docid") & spe & education & spe & placeToBeBorn
				StrEntity=StrEntity &  spe & duty2 & spe & classcification2 & spe & replace(Affair_id&"","'","''")  & spe & Request("skill_id")  & spe & Request("skill_Parent_id")  & spe & Lang & spe &category_id & spe & province_id & spe & SearchWords
			StrEntity=StrEntity+   spe & moreInfo & spe & authorText
			
			'			Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),status,request("parentDocId"),"T_news"
			Call UpdateT_newsPermissionByDocAndSubject (request("docid"),request("subjectid"),status,request("parentDocId"),"T_news")
			
			set logoConn=createObject("adodb.connection")
			logoConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
			
			sql = "update DocsTranslates set htmlStatus=0 where docid = " & request("docid") & " and subjectid = " & request("subjectid") 
			logoConn.execute(sql)


			sql = "delete from T_docsArea where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)
			sql = "delete from T_docsAffair where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)

            if request("subjectid")="35" then
                Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & ",35,153)"
			    logoConn.execute(Isql)
            end if
			
			FirstAffair_id = 0
				if request("area") <> "" then
					areaId = getAreaIds(request("area"))
					areaIdArr = split(areaId, ",")
					for i = 0 to Ubound(areaIdArr)
						if areaIdArr(i) <> "" and isnumeric(areaIdArr(i)) then
                         if request("subjectid")="35" and cstr(areaIdArr(i) & "") = "153" then
                        else
							Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & "," & request("subjectid") & "," & areaIdArr(i) & ")"
							logoConn.execute(sql)
						end if
                    end if
					next
				end if
				
				if request("Affair_id") <> "" then
					AffairId = getAffairIds(request("Affair_id"))
					AffairIdArr = split(AffairId, ",")
					for i = 0 to Ubound(AffairIdArr)
						if AffairIdArr(i) <> "" and isnumeric(AffairIdArr(i)) then
                        if FirstAffair_id=0 then
                                FirstAffair_id = clng(AffairIdArr(i))
                            end if
							Isql = "insert into T_docsAffair (docid,subjectid,Affairid) values (" & request("docid") & "," & request("subjectid") & "," & AffairIdArr(i) & ")"
							logoConn.execute(Isql)
						end if
					next
				end if

			
			
                ShortNews = request("ShortNews")
		if ShortNews="" or not isnumeric(ShortNews) then
		    ShortNews = 0
		end if 

			 if cstr(ShortNews&"") = "1" and FirstAffair_id > 0 then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN T_news as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)

                sql = "update t_articles set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN t_articles as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)
                end if 

                if cstr(ShortNews&"") = "1" then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShortNews=1 and ShowInShortNews=0 and YoavTime>=dateadd(hh,-45,getdate())"
                logoConn.execute(sql)
                end if
			
			Special=Request("Special")
			RelatedSpecial=Request("RelatedSpecial")
			SpecialAffair=Request("SpecialAffair")
			adverStatus=request("adverStatus")
			
		
			if cstr(adverStatus&"") = "1" then
	            adverStatus = 1
            else
                adverStatus = 0
            end if
            if cstr(adverStatus12Home&"") = "1" then
	            adverStatus12Home = "1"
            else
	            adverStatus12Home = "0"
            end if   
       'new ddl
    invoiceStatus=request("invoiceStatus") 
    

    	if cstr(invoiceStatus&"") = "1" then
	    invoiceStatus = 1
    else if cstr(invoiceStatus&"") = "2" then
	    invoiceStatus = 2
     else if cstr(invoiceStatus&"") = "3" then
    invoiceStatus = 3
     else if cstr(invoiceStatus&"") = "4" then
    invoiceStatus = 4
     end if 
     end if 
     end if 
    end if 
			
			ShortNews = request("ShortNews")
			ShowInShortNews=0
			ShowInBursaNames=0
		    if ShortNews="" or not isnumeric(ShortNews) then
		        ShortNews = 0
		    else
		        if cstr(ShortNews&"") = "2" then
		             ShowInShortNews=1
		             ShortNews=0
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "1" then
		             ShowInShortNews=0
		             ShortNews=1
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "3" then
		             ShowInShortNews=0
		             ShortNews=2
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "4" then
		             ShowInShortNews=0
		             ShortNews=0
                     ShowInBursaNames=1
		        else
		             ShowInBursaNames=0
		             ShowInShortNews=0
		             ShortNews=0
		        end if
		    end if
		      
			
			
			logoSql="update t_news set logo='"&session("imgLogo")&"',ShortNews=" & ShortNews & ",ShowInShortNews=" & ShowInShortNews & ",ShowInBursaNames=" & ShowInBursaNames & ", UpdatedByUserName='"&session("UpdatedByUserName")&"',TopTitle='"&TopTitle&"',bottomImg='"&bottomImg&"',Special='"&Special&"',RelatedSpecial='"&RelatedSpecial&"',SpecialAffair='"&SpecialAffair&"', "
			logoSql=logoSql &"linkReplace='',TextReplace='',image2='"&image2&"',image3='"&image3&"',image4='"&image4&"',videoText='"&videoText&"',LargeVideo="&LargeVideo&",VideoYouTube='" & VideoYouTube & "',ImageText2='" & SImageText2 & "',ImageText3='" & SImageText3 & "',ImageText4='" & SImageText4 & "', "  
			logoSql=logoSql &"referenceTEXT='" &referenceTEXT& "',referenceURL='"&referenceURL&"',referenceTEXT2='" &referenceTEXT2& "',referenceURL2='"&referenceURL2&"',referenceTEXT3='" &referenceTEXT3& "',referenceURL3='"&referenceURL3&"', leads='"&leads&"',adverStatus="&adverStatus&",invoiceStatus="& invoiceStatus 
			logoSql=logoSql & ",NoFindReplace=" & NoFindReplace & ",TopTitleColor='" & TopTitleColor & "',TopTitleFontSize='" & TopTitleFontSize & "',Head1Color='" & Head1Color & "',Head1FontSize='" & Head1FontSize & "'"
			logoSql=logoSql &",image5='"&image5&"',ImageText5='"&SImageText5&"',image6='"&image6&"',ImageText6='"&SImageText6&"',ShowFastNews=" & ShowFastNews & ",CountFastNews=" & CountFastNews & ",StartFastNews=" & StartFastNews & ",EndFastNews=" & EndFastNews & " where docid="& request("docid")& " and subjectID="&request("subjectID")
			
			logoConn.execute logoSql
			
			logoSql="update t_entity set adminComments='" & replace(request("adminComments"), "'", "''") & "'"
			if request("ShowChat") = "1" then
			    logoSql = logoSql & ",ShowChat=1"
			else
			    logoSql = logoSql & ",ShowChat=0"
			end if
			if request("ShowStarsOnLeft") = "1" then
			    logoSql = logoSql & ",ShowStarsOnLeft=1"
			else
			    logoSql = logoSql & ",ShowStarsOnLeft=0"
			end if
		    logoSql = logoSql & ",OnlyMembers=" & request("OnlyMembers")
            if request("AutoAffairDocs") <> "" and isnumeric(request("AutoAffairDocs")) then
                logoSql = logoSql & ",AutoAffairDocs=" & request("AutoAffairDocs")
            else
                logoSql = logoSql & ",AutoAffairDocs=0"
            end if
			logoSql = logoSql & " where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
			logoConn.execute logoSql
			
			'logoConn.close
			'set logoConn=nothing
			
			
			
			
			
			'com.UpdateEntity(StrEntity)
			Call UpdateEntity(StrEntity)
			
		'set aConn=createObject("adodb.connection")
		'aConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
		
	if request("status")="2" then	
		txtToTags =  replace(TopTitle & "","''", "'") & chr(10) & replace(head1 & "","''", "'") & chr(10) & replace(head2 & "","''", "'") & chr(10) & replace(text & "","''", "'")
	   sqlP = "select head,head2,mtext from T_newsText where docid=" & request("docid") & " and subjectid=" & request("subjectid")
	   set rsPP = logoConn.execute(sqlP)
	   do while not rsPP.EOF
	    txtToTags = txtToTags & chr(10) & rsPP("head") & chr(10) & rsPP("head2") & chr(10) & rsPP("mtext")
	    rsPP.movenext
	    loop
	    set rsPP = nothing
	
	       
		 UpdateArticleTagsGlobal request("docid"),request("subjectid"),txtToTags
	end if	
		
		if request("autoUserInfo") = "1" then
		    sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(fname & "", "'", "''") & "' and Link_LastName='" & replace(lname & "", "'", "''") & "'"
		    set rsA = logoConn.Execute(sql)
		    if not rsA.EOF then
		        email = rsA("email") 
		        firmName = rsA("firmName") 
		        website = rsA("website") 
		        street = rsA("street") 
		        city = rsA("city") 
		        if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
		            zipcode = 0
		        else
		            zipcode = rsA("zip")
		        end if  
		        phone = rsA("phone") 
		        fax = rsA("fax") 
		        JobTitle = rsA("duty")
		         
		        set rsU = server.CreateObject("ADODB.recordset")
		        sql = "select * from T_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		        rsU.open sql, logoConn, 1, 3
		        rsU("autoUserInfo") = 1
		        rsU("email") = email
		        rsU("firm_name") = firmName
		        rsU("JobTitle") = JobTitle
		        if fname="globes" then
                else
		        rsU("site_url") = website
                end if
		        rsU("address") = street
		        rsU("settling") = city
		        rsU("zipcode") = zipcode
		        rsU("phone") = phone
		        rsU("fax") = fax
		        rsU.update
		        set rsU = nothing  
		    end if
		    set rsA = nothing 
		else
		    sql = "update T_entity set autoUserInfo=0,JobTitle='" & JobTitle & "' where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		    logoConn.Execute(sql) 
		end if  
		
		sql = "delete T_DocsWriters where subjectid=" & request("subjectid") & " and docid=" & request("docid")
		logoConn.execute(sql)
		
		
		if request("fname2") <> "" and request("lname2") <> "" then
		    if request("autoUserInfo2") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname2"), "'", "''") & "' and Link_LastName='" & replace(request("lname2"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = request("email2")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname3") <> "" and request("lname3") <> "" then
		    if request("autoUserInfo3") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname3"), "'", "''") & "' and Link_LastName='" & replace(request("lname3"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = request("email3")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname4") <> "" and request("lname4") <> "" then
		    if request("autoUserInfo4") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname4"), "'", "''") & "' and Link_LastName='" & replace(request("lname4"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = request("email4")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname5") <> "" and request("lname5") <> "" then
		    if request("autoUserInfo5") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname5"), "'", "''") & "' and Link_LastName='" & replace(request("lname5"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = request("email5")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    
	    if request("fname6") <> "" and request("lname6") <> "" then
		    if request("autoUserInfo6") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname6"), "'", "''") & "' and Link_LastName='" & replace(request("lname6"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = request("email6")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if

         set rsUE = server.createobject("ADODB.Recordset")
            sql = "select * from t_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
            rsUE.open sql, logoConn, 1, 3

            if request("EventBoard") = "1" then
                rsUE("EventBoard") = 1
                 EventBoardStartDate = request("EventStartDate") & " " & request("EventStartH") & ":" & request("EventStartM")
                 EventBoardEndDate = request("EventEndDate") & " " & request("EventEndH") & ":" & request("EventEndM")

                 if isdate(EventBoardStartDate) then
                 rsUE("EventBoardStartDate") = cdate(EventBoardStartDate)
                 else
                 rsUE("EventBoardStartDate") = null
                 end if
                 if isdate(EventBoardEndDate) then
                    rsUE("EventBoardEndDate") = cdate(EventBoardEndDate)
                 else
                    rsUE("EventBoardEndDate") = null
                 end if
            else
                rsUE("EventBoard") = 0
                rsUE("EventBoardStartDate") = null
                rsUE("EventBoardEndDate") = null
            end if

            rsUE.update
            set rsUE = nothing
		
		logoConn.close
			set logoConn=nothing
			
			
			'update Affairs if selected
			'if trim(request("Affair_id")&"") <> "" and len(request("Affair_id")&"") > 1 then
			
			'	set objDButils = server.createobject("yoav.dbutils")
				
			'	arrAffair = split(request("Affair_id")&"",",")
			'	for i = 0 to Ubound(arrAffair)
			'		if trim(arrAffair(i)&"") <> "" then
						
			'			set rs=objDButils.GetFastRecoredset("SELECT id,dscr FROM T_area WHERE status=2 and dscr = '"&trim(arrAffair(i)&"")&"'")
			'			if not rs.EOF then
			'				call CreateOneHTMLPage("https://www.news1.co.il/createHtml/HtmlCreateAreaByID.aspx?areaid=" & rs("id"))
			'			end if
			'			set rs = nothing
			'		end if
			'	next
			'	set objDButils = nothing
				
			'end if
			
		
	
	Call UpdateT_newsPermissionByDocAndSubject(request("docid"),request("subjectid"),1,request("parentDocId"),"T_news")
	'Response.Redirect "WaitingListPage.asp"
	Response.Redirect "main.asp"
	Case "5"
    FirstAffair_id=0
    if request("Affair_id") <> "" then
					AffairId = getAffairIds(request("Affair_id"))
					AffairIdArr = split(AffairId, ",")
					for i = 0 to Ubound(AffairIdArr)
						if AffairIdArr(i) <> "" and isnumeric(AffairIdArr(i)) then
                        if FirstAffair_id=0 then
                                FirstAffair_id = clng(AffairIdArr(i))
                                exit for
                            end if
						end if
					next
				end if

			
			
                ShortNews = request("ShortNews")
		if ShortNews="" or not isnumeric(ShortNews) then
		    ShortNews = 0
		end if 

        set logoConn=createObject("adodb.connection")
			logoConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"


			 if cstr(ShortNews&"") = "1" and FirstAffair_id > 0 then
            sql = "update DocsTranslates set htmlStatus=0 where docid = " & request("docid") & " and subjectid = " & request("subjectid") 
			logoConn.execute(sql)

                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN T_news as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-30,getdate()))"

                logoConn.execute(sql)

                sql = "update t_articles set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN t_articles as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-30,getdate()))"

                logoConn.execute(sql)
                end if 

                if cstr(ShortNews&"") = "1" then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShortNews=1 and ShowInShortNews=0 and YoavTime>=dateadd(hh,-45,getdate())"
                logoConn.execute(sql)
                end if

		Call UpdateT_newsPermissionByDocAndSubject(request("docid"),request("subjectid"),5,0,"T_news")
	    'Response.Redirect "currentpages.asp"
	    Response.Redirect "CurrentPages.asp?status=5"
	'Case "6"
		'Call UpdateT_newsPermissionByDocAndSubject(request("docid"),request("subjectid"),6,0,"T_news")
	    'Response.Redirect "CurrentPages.asp?status=6"
	Case else
        'if Session("PrivateWaitingListPage")<>"" and IsNumeric(Session("PrivateWaitingListPage")) Then
        '    if cstr(Session("PrivateWaitingListPage"))=cstr(Request("RecAction")) then
        '        Call UpdateT_newsPermissionByDocAndSubject(request("docid"),request("subjectid"),Session("PrivateWaitingListPage"),0,"T_news")
        '        Response.Redirect "PrivateWaitingListPage.asp"
        '    end if
        'end if
        
        if cstr(Request("RecAction"))="6" or cstr(Session("PrivateWaitingListPage"))=cstr(Request("RecAction")) then
            
            status=0
            If cstr(Request("RecAction"))="6" Then
				status=6
			Else
			    if IsNumeric(Session("PrivateWaitingListPage")) and Session("PrivateWaitingListPage")<>"" then
			        if cstr(Session("PrivateWaitingListPage"))=cstr(Request("RecAction")) then
				        status=cint(Request("RecAction"))       
				    end if
				end if         
			End If   
			
			if status>0 then
			
			'	Com.UpdateT_newsPermission request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText
				Call UpdateT_newsPermission( request("docid"),request("subjectid"),Head1,Head2,Text,LinkName,Linkurl,image,writername,SVidioFile,SImageText)
			'		 '===str for insert to entity table====
			StrEntity=request("subjectid") & spe & hebrewvalue & spe & englishvalue & spe & fname & spe & lname 
				StrEntity=StrEntity &  spe & duty & spe & firm_name & spe & birthyear  & spe & decease_date & spe & birthday & spe & sex & spe & nation1 & spe & nation2 
			StrEntity=StrEntity &  spe & Address & spe & settling & spe & zipcode & spe &  Phone & spe & fax & spe & Email '19
			StrEntity=StrEntity &  spe & site_url & spe & fdate & spe & area & spe & classcification & spe & dealtype '24
				StrEntity=StrEntity &  spe & person_total_salary & spe & person_salary_date & spe & firm_total_balance & spe & firm_balance_date '28
				StrEntity=StrEntity &  spe & firm_total_income & spe & firm_income_date & spe & firm_managers_total & spe & strNow & spe &  request("docid") & spe & education & spe & placeToBeBorn
				StrEntity=StrEntity &  spe & duty2 & spe & classcification2 & spe & replace(Affair_id&"","'","''")  & spe & Request("skill_id")  & spe & Request("skill_Parent_id")  & spe & Lang & spe &category_id & spe & province_id & spe & SearchWords
			StrEntity=StrEntity+   spe & moreInfo & spe & authorText
			
			'			Com.UpdateT_newsPermissionByDocAndSubject request("docid"),request("subjectid"),status,request("parentDocId"),"T_news"
			Call UpdateT_newsPermissionByDocAndSubject (request("docid"),request("subjectid"),status,request("parentDocId"),"T_news")
			
			set logoConn=createObject("adodb.connection")
			logoConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
			
			sql = "update DocsTranslates set htmlStatus=0 where docid = " & request("docid") & " and subjectid = " & request("subjectid") 
			logoConn.execute(sql)

			sql = "delete from T_docsArea where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)
			sql = "delete from T_docsAffair where docid = " & request("docid") & " and subjectid = " & request("subjectid")
			logoConn.execute(sql)

               if request("subjectid")="35" then
                Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & ",35,153)"
			    logoConn.execute(Isql)
            end if
			
			FirstAffair_id = 0 


				if request("area") <> "" then
					areaId = getAreaIds(request("area"))
					areaIdArr = split(areaId, ",")
					for i = 0 to Ubound(areaIdArr)
						if areaIdArr(i) <> "" and isnumeric(areaIdArr(i)) then
                         if request("subjectid")="35" and cstr(areaIdArr(i) & "") = "153" then
                        else
							Isql = "insert into T_docsArea (docid,subjectid,areaid) values (" & request("docid") & "," & request("subjectid") & "," & areaIdArr(i) & ")"
							logoConn.execute(Isql)
                        end if
						end if
					next
				end if
				
				if request("Affair_id") <> "" then
					AffairId = getAffairIds(request("Affair_id"))
					AffairIdArr = split(AffairId, ",")
					for i = 0 to Ubound(AffairIdArr)
						if AffairIdArr(i) <> "" and isnumeric(AffairIdArr(i)) then
                        if FirstAffair_id=0 then
                                FirstAffair_id = clng(AffairIdArr(i))
                            end if
							Isql = "insert into T_docsAffair (docid,subjectid,Affairid) values (" & request("docid") & "," & request("subjectid") & "," & AffairIdArr(i) & ")"
							logoConn.execute(Isql)
						end if
					next
				end if

			
			
                ShortNews = request("ShortNews")
		if ShortNews="" or not isnumeric(ShortNews) then
		    ShortNews = 0
		end if 

			 if cstr(ShortNews&"") = "1" and FirstAffair_id > 0 then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN T_news as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)

                sql = "update t_articles set nfchtmlstatus=0,mobileHtmlStatus=0 where ShowFastNews=1 and CountFastNews>0 and docid in (select n.docid from T_docsAffair AS a INNER JOIN t_articles as n on a.docid=n.docId and a.subjectid=n.subjectId where Affairid=" & FirstAffair_id & " and ShowFastNews=1 and CountFastNews>0 and YoavTime>=dateadd(hh,-14,getdate()))"

                logoConn.execute(sql)
                end if 

                if cstr(ShortNews&"") = "1" then
                sql = "update t_news set nfchtmlstatus=0,mobileHtmlStatus=0 where ShortNews=1 and ShowInShortNews=0 and YoavTime>=dateadd(hh,-45,getdate())"
                logoConn.execute(sql)
                end if
			
			Special=Request("Special")
			RelatedSpecial=Request("RelatedSpecial")
			SpecialAffair=Request("SpecialAffair")
			adverStatus=request("adverStatus")
			
		
			if cstr(adverStatus&"") = "1" then
	            adverStatus = 1
            else
                adverStatus = 0
            end if 
    
            if cstr(adverStatus12Home&"") = "1" then
	            adverStatus12Home = "1"
            else
	            adverStatus12Home = "0"
            end if  

     'new ddl
        invoiceStatus=request("invoiceStatus") 
    	if cstr(invoiceStatus&"") = "1" then
	    invoiceStatus = 1
    else if cstr(invoiceStatus&"") = "2" then
	    invoiceStatus = 2
     else if cstr(invoiceStatus&"") = "3" then
    invoiceStatus = 3
     else if cstr(invoiceStatus&"") = "4" then
    invoiceStatus = 4
     end if 
     end if 
     end if 
    end if  
			
			ShortNews = request("ShortNews")
			ShowInShortNews=0
			ShowInBursaNames=0
		    if ShortNews="" or not isnumeric(ShortNews) then
		        ShortNews = 0
		    else
		        if cstr(ShortNews&"") = "2" then
		             ShowInShortNews=1
		             ShortNews=0
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "1" then
		             ShowInShortNews=0
		             ShortNews=1
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "3" then
		             ShowInShortNews=0
		             ShortNews=2
                     ShowInBursaNames=0
		        elseif cstr(ShortNews&"") = "4" then
		             ShowInShortNews=0
		             ShortNews=0
                     ShowInBursaNames=1
		        else
		             ShowInBursaNames=0
		             ShowInShortNews=0
		             ShortNews=0
		        end if
		    end if
		      
			
			
			logoSql="update t_news set logo='"&session("imgLogo")&"',ShortNews=" & ShortNews & ",ShowInShortNews=" & ShowInShortNews & ",ShowInBursaNames=" & ShowInBursaNames & ", UpdatedByUserName='"&session("UpdatedByUserName")&"',TopTitle='"&TopTitle&"',bottomImg='"&bottomImg&"',Special='"&Special&"',RelatedSpecial='"&RelatedSpecial&"',SpecialAffair='"&SpecialAffair&"', "
			logoSql=logoSql &"linkReplace='',TextReplace='',image2='"&image2&"',image3='"&image3&"',image4='"&image4&"',videoText='"&videoText&"',LargeVideo="&LargeVideo&",VideoYouTube='" & VideoYouTube & "',ImageText2='" & SImageText2 & "',ImageText3='" & SImageText3 & "',ImageText4='" & SImageText4 & "', "  
			logoSql=logoSql &"referenceTEXT='" &referenceTEXT& "',referenceURL='"&referenceURL&"',referenceTEXT2='" &referenceTEXT2& "',referenceURL2='"&referenceURL2&"',referenceTEXT3='" &referenceTEXT3& "',referenceURL3='"&referenceURL3&"', leads='"&leads&"',adverStatus="&adverStatus&",adverStatus12Home='"&adverStatus12Home&"',invoiceStatus="& invoiceStatus
			logoSql=logoSql & ",NoFindReplace=" & NoFindReplace & ",TopTitleColor='" & TopTitleColor & "',TopTitleFontSize='" & TopTitleFontSize & "',Head1Color='" & Head1Color & "',Head1FontSize='" & Head1FontSize & "'"
			logoSql=logoSql &",image5='"&image5&"',ImageText5='"&SImageText5&"',image6='"&image6&"',ImageText6='"&SImageText6&"',ShowFastNews=" & ShowFastNews & ",CountFastNews=" & CountFastNews & ",StartFastNews=" & StartFastNews & ",EndFastNews=" & EndFastNews & " where docid="& request("docid")& " and subjectID="&request("subjectID")
			
			logoConn.execute logoSql
			
			logoSql="update t_entity set adminComments='" & replace(request("adminComments"), "'", "''") & "'"
			if request("ShowChat") = "1" then
			    logoSql = logoSql & ",ShowChat=1"
			else
			    logoSql = logoSql & ",ShowChat=0"
			end if
			if request("ShowStarsOnLeft") = "1" then
			    logoSql = logoSql & ",ShowStarsOnLeft=1"
			else
			    logoSql = logoSql & ",ShowStarsOnLeft=0"
			end if
		    logoSql = logoSql & ",OnlyMembers=" & request("OnlyMembers")
            if request("AutoAffairDocs") <> "" and isnumeric(request("AutoAffairDocs")) then
                logoSql = logoSql & ",AutoAffairDocs=" & request("AutoAffairDocs")
            else
                logoSql = logoSql & ",AutoAffairDocs=0"
            end if
			logoSql = logoSql & " where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
			logoConn.execute logoSql
			
			'logoConn.close
			'set logoConn=nothing
			
			
			
			
			
			'com.UpdateEntity(StrEntity)
			Call UpdateEntity(StrEntity)
			
		'set aConn=createObject("adodb.connection")
		'aConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
		
		if request("status")="2" then
		txtToTags =  replace(TopTitle & "","''", "'") & chr(10) & replace(head1 & "","''", "'") & chr(10) & replace(head2 & "","''", "'") & chr(10) & replace(text & "","''", "'")
	   sqlP = "select head,head2,mtext from T_newsText where docid=" & request("docid") & " and subjectid=" & request("subjectid")
	   set rsPP = logoConn.execute(sqlP)
	   do while not rsPP.EOF
	    txtToTags = txtToTags & chr(10) & rsPP("head") & chr(10) & rsPP("head2") & chr(10) & rsPP("mtext")
	    rsPP.movenext
	    loop
	    set rsPP = nothing
		
       
		 UpdateArticleTagsGlobal request("docid"),request("subjectid"),txtToTags
		end if

		
		if request("autoUserInfo") = "1" then
		    sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(fname & "", "'", "''") & "' and Link_LastName='" & replace(lname & "", "'", "''") & "'"
		    set rsA = logoConn.Execute(sql)
		    if not rsA.EOF then
		        email = rsA("email") 
		        firmName = rsA("firmName") 
		        website = rsA("website") 
		        street = rsA("street") 
		        city = rsA("city") 
		        if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
		            zipcode = 0
		        else
		            zipcode = rsA("zip")
		        end if  
		        phone = rsA("phone") 
		        fax = rsA("fax") 
		        JobTitle = rsA("duty")
		         
		        set rsU = server.CreateObject("ADODB.recordset")
		        sql = "select * from T_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		        rsU.open sql, logoConn, 1, 3
		        rsU("autoUserInfo") = 1
		        rsU("email") = email
		        rsU("firm_name") = firmName
		        rsU("JobTitle") = JobTitle
		        if fname="globes" then
                else
		        rsU("site_url") = website
                end if
		        rsU("address") = street
		        rsU("settling") = city
		        rsU("zipcode") = zipcode
		        rsU("phone") = phone
		        rsU("fax") = fax
		        rsU.update
		        set rsU = nothing  
		    end if
		    set rsA = nothing 
		else
		    sql = "update T_entity set autoUserInfo=0,JobTitle='" & JobTitle & "' where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
		    logoConn.Execute(sql) 
		end if  
		
		sql = "delete T_DocsWriters where subjectid=" & request("subjectid") & " and docid=" & request("docid")
		logoConn.execute(sql)
		
		
		if request("fname2") <> "" and request("lname2") <> "" then
		    if request("autoUserInfo2") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname2"), "'", "''") & "' and Link_LastName='" & replace(request("lname2"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname2")
	                rsU("lname") = request("lname2")
	                rsU("email") = request("email2")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname3") <> "" and request("lname3") <> "" then
		    if request("autoUserInfo3") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname3"), "'", "''") & "' and Link_LastName='" & replace(request("lname3"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname3")
	                rsU("lname") = request("lname3")
	                rsU("email") = request("email3")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname4") <> "" and request("lname4") <> "" then
		    if request("autoUserInfo4") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname4"), "'", "''") & "' and Link_LastName='" & replace(request("lname4"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname4")
	                rsU("lname") = request("lname4")
	                rsU("email") = request("email4")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    if request("fname5") <> "" and request("lname5") <> "" then
		    if request("autoUserInfo5") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname5"), "'", "''") & "' and Link_LastName='" & replace(request("lname5"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname5")
	                rsU("lname") = request("lname5")
	                rsU("email") = request("email5")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if
	    
	    
	    if request("fname6") <> "" and request("lname6") <> "" then
		    if request("autoUserInfo6") = "1" then
	            sql = "select email,firmName,website,street,city,zip,phone,fax,duty from T_PeopleInNews where Link_FirstName='" & replace(request("fname6"), "'", "''") & "' and Link_LastName='" & replace(request("lname6"), "'", "''") & "'"
	            set rsA = logoConn.Execute(sql)
	            if not rsA.EOF then
	                email = rsA("email") 
	                JobTitle = rsA("duty")
	                firmName = rsA("firmName")
	                website = rsA("website") 
	                street = rsA("street") 
	                city = rsA("city") 
	                if not isnumeric(rsA("zip")) or cstr(rsA("zip")&"") = "" then
	                    zipcode = 0
	                else
	                    zipcode = rsA("zip")
	                end if  
	                phone = rsA("phone") 
	                fax = rsA("fax") 
    		        
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 1
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = email
	                rsU("firm_name") = firmName
	                rsU("site_url") = website
	                rsU("address") = street
	                rsU("settling") = city
	                rsU("zipcode") = zipcode
	                rsU("phone") = phone
	                rsU("fax") = fax
	                rsU("JobTitle") = JobTitle
	                rsU.update
	                set rsU = nothing 
	            end if
	            set rsA = nothing 
		    else
		            set rsU = server.CreateObject("ADODB.recordset")
	                sql = "select * from T_DocsWriters where Aid=-1"
	                rsU.open sql, logoConn, 1, 3
	                rsU.AddNew
	                rsU("docid") = request("docid")
	                rsU("subjectid") = request("subjectid")
	                rsU("autoUserInfo") = 0
	                rsU("fname") = request("fname6")
	                rsU("lname") = request("lname6")
	                rsU("email") = request("email6")
	                rsU.update
	                set rsU = nothing 
		    end if
	    end if

         set rsUE = server.createobject("ADODB.Recordset")
            sql = "select * from t_entity where subjectid=" & request("subjectid") & " and doc_id=" & request("docid")
            rsUE.open sql, logoConn, 1, 3

            if request("EventBoard") = "1" then
                rsUE("EventBoard") = 1
                 EventBoardStartDate = request("EventStartDate") & " " & request("EventStartH") & ":" & request("EventStartM")
                 EventBoardEndDate = request("EventEndDate") & " " & request("EventEndH") & ":" & request("EventEndM")

                 if isdate(EventBoardStartDate) then
                 rsUE("EventBoardStartDate") = cdate(EventBoardStartDate)
                 else
                 rsUE("EventBoardStartDate") = null
                 end if
                 if isdate(EventBoardEndDate) then
                    rsUE("EventBoardEndDate") = cdate(EventBoardEndDate)
                 else
                    rsUE("EventBoardEndDate") = null
                 end if
            else
                rsUE("EventBoard") = 0
                rsUE("EventBoardStartDate") = null
                rsUE("EventBoardEndDate") = null
            end if

            rsUE.update
            set rsUE = nothing
		
		logoConn.close
			set logoConn=nothing
			
			
			'update Affairs if selected
			'if trim(request("Affair_id")&"") <> "" and len(request("Affair_id")&"") > 1 then
			
			'	set objDButils = server.createobject("yoav.dbutils")
				
			'	arrAffair = split(request("Affair_id")&"",",")
			'	for i = 0 to Ubound(arrAffair)
			'		if trim(arrAffair(i)&"") <> "" then
			'			
			'			set rs=objDButils.GetFastRecoredset("SELECT id,dscr FROM T_area WHERE status=2 and dscr = '"&trim(arrAffair(i)&"")&"'")
			'			if not rs.EOF then
			'				call CreateOneHTMLPage("https://www.news1.co.il/createHtml/HtmlCreateAreaByID.aspx?areaid=" & rs("id"))
			'			end if
			'			set rs = nothing
			'		end if
			'	next
			'	set objDButils = nothing
				
			'end if
			
		
			
			if request("fromOpenWin") <> "YES" then
				docOpener=cstr(request("docOpener"))
				'if request("status")=0 then
				'	Response.Redirect "newdetails.asp"
				'end if
				
				if request("status")=0 then
					Response.Redirect "newdetails.asp"
				end if
				if request("status")=1 then
				'	Response.Redirect "waitinglistpage.asp?orderby=ArticlesAndNews"
					Response.Redirect "main.asp"
				end if
				if request("status")=2 then
					Response.Redirect "Main.asp" '"currentpages.asp"
				elseif request("status")=3 then
					Response.Redirect "Main.asp" '"currentpages.asp"
				elseif request("status")=5 then
				    Response.Redirect "Main.asp" '"CurrentPages.asp?status=5"
				elseif request("status")=6 then
				    Response.Redirect "Main.asp" '"CurrentPages.asp?status=6"
				else
					Response.Redirect "Main.asp" '"PrivateWaitingListPage.asp"
				end if
			end if
        
        end if
	end if
end select  
	



		   if request("AddToMailingList") = "on"  then
		   set objDButils=server.CreateObject("yoav.Com_Logic")
		   strSqlDoc="select * from T_MailingList where doc_Id=" & request("docid") & " and subjectId= " & request("subjectid")
		   set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)  
		   
		   if MailListrs.eof then
		        Com.InsertMailingList request("subjectid"),request("docid"),Email
		   end if
		  else
		       set objDButils=server.CreateObject("yoav.Com_Logic")
		       strSqlDoc="delete T_MailingList where doc_Id=" & request("docid") & " and subjectId= " & request("subjectid")
		       set MailListrs=objDButils.GetFastRecoredset(strSqlDoc)  
		   
		
		 end if


		End If
    
end function

function updateTxtLinksReplaceTable(txt, links, docid, subjectid)

	set tConn=createObject("adodb.connection")
	tConn.open "dsn=yoav;uid=sa;pwd=yoavrdbms;"
	txtArr = split(txt, ",")
	linkArr = split(links, ",")
    if cstr(adverStatus12Home&"") = "1" then
	adverStatus12Home = "1"
    else
	    adverStatus12Home = "0"
    end if
	
	select case cstr(subjectid)
		Case "1","20","7","12","13","40","41","42","44"
			Sql="update T_news set linkReplace='',TextReplace='' where docid="& docid & " and subjectID="& subjectid
		Case else
			Sql="update T_articles set linkReplace='',TextReplace='' where docid="& docid & " and subjectID="& subjectid
		end select
			
	tConn.execute Sql
	
	
	for iTxt = 0 to Ubound(txtArr)
		if trim(txtArr(iTxt)&"") <> "" and trim(txtArr(iTxt)&"") <> "," then
			sql = "insert into T_txtLinksReplace (docid,subjectid,txt,link) values (" & docid & "," & subjectid & ",'" & trim(txtArr(iTxt)&"") & "','" & trim(linkArr(iTxt)&"") & "')"
			tConn.execute Sql
		end if
	next
	
	tConn.close
	set tConn=nothing

end function


%>
<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
Const MaxPerPage=20
Dim strFileName
Dim totalPut,TotalPages
Dim rs, sql
Dim CoverID,UserSearch,Keyword,strField
Dim Action,FoundErr,ErrMsg
Dim tmpDays
keyword=Trim(Request("keyword"))
If keyword<>"" Then 
	keyword=oblog.filt_badstr(keyword)
End If
strField=Trim(Request("Field"))
UserSearch=Trim(Request("UserSearch"))
Action=Trim(Request("Action"))
CoverID=Trim(Request("CoverID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")

If UserSearch="" Then
	UserSearch=0
Else
	UserSearch=Clng(UserSearch)
End If
strFileName="Administrator_CoverProducts.asp?UserSearch=" & UserSearch
If strField<>"" Then
	strFileName=strFileName&"&Field="&strField
End If
If keyword<>"" Then
	strFileName=strFileName&"&keyword="&keyword
End If
If Request("page")<>"" Then
    currentPage=Cint(Request("page"))
Else
	currentPage=1
End If

%>
<html>
<head>
<title>ע������������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function unselectall()
{
    If(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    If (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
</SCRIPT>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>�� �� �� �� �� ��</strong></td>
  </tr>
  <form name="form1" action="Administrator_CoverProducts.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>���ٲ��������</strong></td>
      <td width="687" height="30">
	   <select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">���ע���500�����</option>
    
		  <option value="3">VIP���</option>

        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="Administrator_CoverProducts.asp">�������������ҳ</a>&nbsp;|&nbsp;<a href="?Action=Add">��������</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="Administrator_CoverProducts.asp">
  <tr class="tdbg">
    <td width="120"><strong>����߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
		  <!--<option value="UserName" Selected>�����</option>-->
		  <option value="CoverName" Selected>�������</option>
		  <option value="CoverID" >�����ID</option>
		  <option value="address" >����ַ</option>
		  <option value="beizhu" >����ע</option>
		  <option value="xuwei_beizhu" >����ϸ</option>

      </select>
	  
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10"> 
	  ��Ϊ�գ����ѯ�������</td>
  </tr>
</form>
</table>
<%
If Action="Add" Then
	Call AddCover()
ElseIf Action="SaveAdd" Then
	Call SaveAdd()
ElseIf Action="Modify" Then
	Call Modify()
ElseIf Action="SaveModify" Then
	Call SaveModify()
ElseIf Action="Del" Then
	Call DelUser()
ElseIf Action="Lock" Then
	Call LockUser()
ElseIf Action="UnLock" Then
	Call UnLockUser()
ElseIf Action="Move" Then
	Call MoveUser()
ElseIf Action="Update" Then
	Call UpdateUser()
ElseIf Action="DoUpdate" Then
	Call DoUpdate()
ElseIf Action="DoUpdatelog" Then
	Call DoUpdatelog()
ElseIf Action="gouser1" Then
	Call gouser1()
ElseIf Action="gouser2" Then
	Call gouser2()
ElseIf Action="ShowInfo" Then
	Call ShowInfo()


ElseIf Action="Discontinued" Then
	Call Discontinued()
ElseIf Action="UnDiscontinued" Then
	Call UnDiscontinued()
	
	
Else
	Call main()
End If

If FoundErr=True Then
	Call WriteErrMsg()
End If


Sub main()
	Dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Administrator_CoverProducts.asp'>ע������������</a>&nbsp;&gt;&gt;&nbsp;"
	Select Case UserSearch
		Case 0
			sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 Order By CoverID Desc"
			strGuide=strGuide & "���ע���500�����"
		Case 1
			sql="Select Top 100 * From Product_CoverProducts Where Discontinued=0 Order By log_count Desc"
			strGuide=strGuide & "������������ǰ100�����"
		Case 2
			sql="Select Top 100 * From  Product_CoverProducts Where Discontinued=0 Order By log_count"
			strGuide=strGuide & "�����������ٵ�100�����"
		Case 3
			sql="Select * From  Product_CoverProducts Where Discontinued=1 Order By CoverID Desc"
			strGuide=strGuide & "ǰ500�����������"
		Case 4
			sql="Select * From  Product_CoverProducts Where Discontinued=0 And user_level=9 Order By CoverID Desc"
			strGuide=strGuide & "ǰ̨����Ա"
		Case 5
			sql="Select * From  Product_CoverProducts Where Discontinued=0 And user_isbest=1 Order By CoverID Desc"
			strGuide=strGuide & "�Ƽ����"
		Case 6
			sql="Select * From Product_CoverProducts Where Discontinued=0 And User_Level=6 Order By CoverID  Desc"
			strGuide=strGuide & "�ȴ�������֤֤�����"
		Case 7
			sql="Select * From Product_CoverProducts Where Discontinued=0 And  LockUser =1 Order By CoverID  Desc"
			strGuide=strGuide & "���б���ס�����"
		Case 10
			If Keyword="" Then
				sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 Order By CoverID Desc"
				strGuide=strGuide & "�������"
			Else
				Select Case strField
				Case "CoverID"

					If IsNumeric(Keyword)=false Then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>���ID������������</li>"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And CoverID =" & Clng(Keyword)
						strGuide=strGuide & "���ID����<font color=red> " & Clng(Keyword) & " </font>�����"
					End If
				Case "CoverName"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And CoverName like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "������к��С� <font color=red>" & Keyword & "</font> �������"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And CoverName= '" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "��������ڡ� <font color=red>" & Keyword & "</font> �������"
					End If
					
				Case "address"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And address like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "ͨѶ��ַ�к��С� <font color=red>" & Keyword & "</font> �������"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And address='" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "ͨѶ��ַ���ڡ� <font color=red>" & Keyword & "</font> �������"
					End If
				Case "beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And beizhu like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "��ע�к��С� <font color=red>" & Keyword & "</font> �������"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And beizhu='" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "��ע���ڡ� <font color=red>" & Keyword & "</font> �������"
					End If
				Case "xuwei_beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And xuwei_beizhu like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "��ϸ�к��С� <font color=red>" & Keyword & "</font> �������"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And xuwei_beizhu='" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "��ϸ���ڡ� <font color=red>" & Keyword & "</font> �������"
					End If
				Case "logcount"
					sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And log_count < " & Clng(Keyword) & " Order By CoverID  Desc"
					strGuide=strGuide & "������С�ڡ� <font color=red>" & Keyword & "</font> �������"
				Case "logintimes"
					sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And logintimes < " & Clng(Keyword) & " Order By CoverID  Desc"
					strGuide=strGuide & "��¼����С�ڡ� <font color=red>" & Keyword & "</font> �������"
				Case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And datediff(d,lastlogintime,getdate())>"&Clng(Keyword)&" Order By CoverID  Desc"
				   Else 
				   	sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And datediff('d',lastlogintime,now())>"&Clng(Keyword)&" Order By CoverID  Desc"
				   End If 
					strGuide=strGuide & "�� <font color=red>" & Keyword & "</font> ������δ��¼�����"
				End Select
			End If
		Case Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	End Select
	strGuide=strGuide & "</td><td align='right'>"
	If FoundErr=True Then Exit Sub
	If Not IsObject(conn) Then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
'	Response.Write sql
'	Response.End 
	rs.Open sql,Conn,1,1
  	If rs.Eof And rs.Bof Then
		strGuide=strGuide & "���ҵ� <font color=red>0</font> �����</td></tr></table>"
		Response.Write strGuide
	Else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> �����</td></tr></table>"
		Response.Write strGuide
		If currentpage<1 Then
       		currentpage=1
    	End If
    	If (currentpage-1)*MaxPerPage>totalput Then
	   		If (totalPut mod MaxPerPage)=0 Then
	     		currentpage= totalPut \ MaxPerPage
		  	Else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		End If

    	End If
	    If currentPage=1 Then
        	showContent
        	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"�����")
   	 	Else
   	     	If (currentPage-1)*MaxPerPage<totalPut Then
         	   	rs.Move  (currentPage-1)*MaxPerPage
         		Dim Bookmark
           		Bookmark=rs.Bookmark
            	showContent
            	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"�����")
        	Else
	        	currentPage=1
           		showContent
           		Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"�����")
	    	End If
		End If
	End If
	rs.Close
	Set rs=Nothing
End Sub


Sub showContent()
   	Dim i
    i=0
%>
<table width='98%' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="Administrator_CoverProducts.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="39" align="center"><font color="#FFFFFF">ѡ��</font></td>
            <td width="34" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="181" height="22" align="center"><font color="#FFFFFF">�������</font></td>
            <td width="225" height="22" align="center"><font color="#FFFFFF">�������</font></td>
            <td width="130" align="center"><font color="#FFFFFF">�����Ĳ�ƷIDs (��ѡ)</font></td>
			<td width="125" height="22" align="center"><font color="#FFFFFF">���������� (��ѡ)</font></td>
            
			<td width="85" align="center"><font color="#cccccc">����ѡ����</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF">�����Żݼ۸�</font></td>
            <td width="120" height="22" align="center"><font color="#FFFFFF">����</font></td>
          </tr>
          <%do while Not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td  align="center"><input name='CoverID' type='checkbox' onClick="unselectall()" id="CoverID" value='<%=Cstr(rs("CoverID"))%>' /></td>
            <td  align="center"><%=rs("CoverID")%></td>
            <td  align="center"><%
			Response.Write "<a href='?action=Modify&CoverID="&rs("CoverID")&"'" 			
			Response.Write """  title='[���ͼƬ��ַ:"&rs("CoverPhoto")&"]'>" & rs("CoverName") & "</a>"
			%> </td>
            <td align="center" title="<%=rs("CoverDesc")%>"> <%=Left(rs("CoverDesc") ,20)%> ...</td>
            
			<td align="center"> <%=rs("Iclude_ProductIDs")%></td>
			
			<td align="center"> <%

		Response.Write rs("ClassIDs")& "&nbsp;"
	

	%> </td>
            
			<td align="center"> <%=rs("MaxNumber")%> ����</td>
			
            <td align="center"><%=FormatCurrency(rs("CoverPrice"))%></td>
            <td align="center">
			

			
			<%
			
			
			If rs("Discontinued")=1 Then
				Response.Write "<a href='Administrator_CoverProducts.asp?Action=Discontinued&CoverID=" & rs("CoverID") & "' title=''>"
				Response.Write("<font color=red title='������Ѿ�ֹͣ�������������뿪ʼ������'>��ֹͣ</font>")
			ElseIf rs("Discontinued")=0 Then
				Response.Write "<a href='Administrator_CoverProducts.asp?Action=UnDiscontinued&CoverID=" & rs("CoverID") & "' title=''>"
				Response.Write("<font color=green title='��������������У���������ֹͣ�������'>������</font>")
			End If
			Response.Write "</font></a>&nbsp;"
			
		Response.Write "<a href='Administrator_CoverProducts.asp?Action=Modify&CoverID=" & rs("CoverID") & "'>�޸�</a>&nbsp;"
		
		'Response.Write "<a href='Administrator_CoverProducts.asp?Action=gouser2&username=" & rs("username") & "' target='blank'>����̨</a>&nbsp;"

        Response.Write "<a href='Administrator_CoverProducts.asp?Action=Del&CoverID=" & rs("CoverID") & "' onClick='return confirm(""ȷ��Ҫɾ���������"");'>ɾ��</a>&nbsp;"
		
		%> </td>
          </tr>
          <%
	i=i+1
	If i>=MaxPerPage Then Exit do
	rs.movenext
loop
%>
        </table>  
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />
              ѡ�б�ҳ��ʾ���������</td>
            <td> <strong>������</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=True" />
              ɾ��&nbsp;&nbsp;&nbsp;&nbsp; 
              <!--<input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">�ƶ���
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">�������</option>
                <option value="7">ע����Ŀ����</option>
                <option value="8">VIP��Ŀ����</option>
                <option value="9">ǰ̨������</option>
              </select>-->
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" ִ �� " /> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
End Sub




Sub AddCover()

%>
<FORM name="Form1" Action="Administrator_CoverProducts.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">ע���� �������</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">�������</TD>
      <TD width="60%"><input name="CoverName" type="text" value="" size=20 maxlength=50 /></TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>���������</td>
      <td><textarea name="CoverDesc" cols="38" rows="8"></textarea></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����Ĳ�Ʒ�У�(��ѡ)</td>
      <td><!--<Select name="Iclude_ProductIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0,0,"Product_ClassChild")	���Ͷ�ѡselect�ؼ�.%>
			
		  </select>-->
	  		<input name="Iclude_ProductIDs" type="text" value="" size=6 maxlength=30 />
	  </td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���������ͣ�(��ѡ)</td>
      <td><!--<Select name="ClassIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0,0,"Product_Class")	���Ͷ�ѡselect�ؼ�.%>
			
		  </select>-->
			<input name="ClassIDs" type="text" value="" size=6 maxlength=30 />  
	  </td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���������ѡ��Ʒ������</td>
      <td><input name="MaxNumber" type="text" value="" size=3 maxlength=10 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����Żݼ۸�</td>
      <td><input name="CoverPrice" type="text" value="" size=10 maxlength=20 style="text-align:center;" />ԪRMB</td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>�����չʾͼƬ��</TD>
      <TD><input name="CoverPhoto" type="text" value="" size=50 maxlength=500 /> 
	  
	  </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">�����e����״̬��</TD>
      <TD width="60%"><INPUT name="Discontinued" type=radio value="0" checked="checked" />������&nbsp;&nbsp;&nbsp;<INPUT name="Discontinued" type=radio value="1" />ֹͣ����
      </TD>
    </TR>
   
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name=Submit   type=submit id="Submit" value="���������"></TD>
    </TR>
  </TABLE>
</form>
<%

End Sub




Sub Modify()
	Dim CoverID
	Dim rsCover,sqlCover
	CoverID=Trim(Request("CoverID"))
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		Exit Sub
	Else
		CoverID=Clng(CoverID)
	End If
	Set rsCover=Server.CreateObject("Adodb.RecordSet")
	sqlCover="Select * From Product_CoverProducts Where CoverID=" & CoverID
	If Not IsObject(conn) Then link_database
	rsCover.Open sqlCover,Conn,1,3
	If rsCover.Bof And rsCover.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���������</li>"
		rsCover.Close
		Set rsCover=Nothing
		Exit Sub
	End If
%>
<FORM name="Form1" Action="Administrator_CoverProducts.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">ע���� �������</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">�������</TD>
      <TD width="60%"><input name="CoverName" type="text" value="<%=rsCover("CoverName")%>" size=20 maxlength=50 /></TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>���������</td>
      <td><textarea name="CoverDesc" cols="38" rows="8"><%=rsCover("CoverDesc")%></textarea></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����Ĳ�Ʒ�У�(��ѡ)</td>
      <td><!--<Select name="Iclude_ProductIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0,0,"Product_ClassChild")		���Ϳؼ�%>
			
		  </select>-->
		  
		  <input name="Iclude_ProductIDs" type="text" value="<%=rsCover("Iclude_ProductIDs")%>" size=6 maxlength=30 />
	  </td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���������ͣ�(��ѡ)</td>
      <td><!--<Select name="ClassIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0, rsCover("ClassIDs"), "Product_Class")		���Ϳؼ�%>
			
		  </select>-->
	  		<input name="ClassIDs" type="text" value="<%=rsCover("ClassIDs")%>" size=6 maxlength=30 />
	  </td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���������ѡ��Ʒ������</td>
      <td><input name="MaxNumber" type="text" value="<%=rsCover("MaxNumber")%>" size=3 maxlength=10 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����Żݼ۸�</td>
      <td><input name="CoverPrice" type="text" value="<%=FormatCurrency(rsCover("CoverPrice"))%>" size=10 maxlength=20 style="text-align:center;" />ԪRMB</td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>�����չʾͼƬ��</TD>
      <TD><input name="CoverPhoto" type="text" value="<%=rsCover("CoverPhoto")%>" size=50 maxlength=500 /> 
	  
	  </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">�����e����״̬��</TD>
      <TD width="60%"><INPUT name="Discontinued" type=radio value="0" <% If rsCover("Discontinued")=0 Then Response.Write "checked" %> />������&nbsp;&nbsp;&nbsp;<INPUT name="Discontinued" type=radio value="1" <% If rsCover("Discontinued")=1 Then Response.Write "checked" %> />ֹͣ����
      </TD>
    </TR>
   
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="��������޸Ľ��"> <input name="CoverID" type="hidden" id="CoverID" value="<%=rsCover("CoverID")%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rsCover.Close
	Set rsCover=Nothing
End Sub



Sub gouser1()
If is_ot_user=1 Then
	Response.Write("�����ⲿ���ݿ����������ʹ�ô˹���"):Response.End()
End If
%>
<FORM name="Form1" action="Administrator_CoverProducts.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">��¼����Ŀ�����̨</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          ������������Ա��¼����Ŀ�Ĺ��������й���<br>
          ������Ŀ���������ϰ�ʱ���ɽ���������̨��<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">������Ŀ���ƣ�</td>
    <td height="25"><input name="username" type="text" id="username" value="" size="30" maxlength="50"></td>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value=" �ύ "></td>
  </tr>
</table>
</form>
  <%
End Sub
%>
</body>
</html>
<%



Sub SaveAdd()
	Dim rsCover,sqlCover
	Dim CoverName,CoverDesc,Iclude_ProductIDs,ClassIDs,MaxNumber,CoverPrice,CoverPhoto,Discontinued
	
	CoverName			=Trim(Request("CoverName"))
	CoverDesc			=Trim(Request("CoverDesc"))
	Iclude_ProductIDs	=Trim(Request("Iclude_ProductIDs"))
	ClassIDs			=Trim(Request("ClassIDs"))
	MaxNumber			=Trim(Request("MaxNumber"))
	CoverPrice			=Trim(Request("CoverPrice"))
	CoverPhoto			=Trim(Request("CoverPhoto"))
	Discontinued		=Trim(Request("Discontinued"))
	

	If CoverName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>������Ʋ���Ϊ��</li>"
	Else
		CoverName = Trim(CoverName)
	End If

	If CoverDesc="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>�����������Ϊ��</li>"
	Else
		CoverDesc = Trim(CoverDesc)
	End If
	
	If Not Instr(Iclude_ProductIDs,",")>0 Then
		Errmsg=Errmsg & "<br><li>��ѡ��2������2�����ϵĲ�Ʒ����Ϊ�����</li>"
		founderr=True
	Else
		Iclude_ProductIDs = Trim(Iclude_ProductIDs)
	End If

	If Not Len(ClassIDs)>0 Then
		Errmsg=Errmsg & "<br><li>��ѡ������1������������࣡</li>"
		founderr=True
	Else
		ClassIDs = Trim(ClassIDs)
	End If
	
	If MaxNumber="" Or Not isNumeric(MaxNumber) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��� ��ѡ�������Ʒ���� ����Ϊ�գ�����ҪΪ���֣�</li>"
	Else
		'UnitPrice = FormatCurrency(UnitPrice)
		MaxNumber = Cint(MaxNumber)
	End If
	
	If CoverPrice="" Or Not isNumeric(CoverPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��� �����Żݼ۸� ����Ϊ�գ�����ҪΪ�������֣�</li>"
	Else
		CoverPrice = FormatCurrency(CoverPrice)
	End If
	
	
	If CoverPhoto="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>���ͼƬ�ϴ�����Ϊ��</li>"
	Else
		CoverPhoto = Trim(CoverPhoto)
	End If
	
	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��ѡ�� �Ƿ��Ѿ������������ ����״̬��</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	

	If founderr=True Then
		Set rs=Nothing
		Exit Sub
	End If
	
	
	Set rsCover=Server.CreateObject("Adodb.RecordSet")
	sqlCover="Select * From Product_CoverProducts"
	If Not IsObject(conn) Then link_database
	rsCover.Open sqlCover,Conn,1,3

	rsCover.Addnew
	
		If Not CoverName="" Then  			rsCover("CoverName")			=CoverName
		If Not CoverDesc="" Then  			rsCover("CoverDesc")			=CoverDesc
		If Not Iclude_ProductIDs="" Then  	rsCover("Iclude_ProductIDs")	=Iclude_ProductIDs
		If Not ClassIDs="" Then  			rsCover("ClassIDs")				=ClassIDs
		If Not MaxNumber="" Then  			rsCover("MaxNumber")			=MaxNumber
		If Not CoverPrice="" Then  			rsCover("CoverPrice")			=CoverPrice
		If Not CoverPhoto="" Then  			rsCover("CoverPhoto")			=CoverPhoto
		If Not Discontinued="" Then  		rsCover("Discontinued")			=Discontinued

		
	rsCover.Update
	
	rsCover.Close
	Set rsCover=Nothing
	
	oblog.showok "ע��������ɹ�!",""
End Sub



Sub gouser2()
	If is_ot_user=1 Then
		Response.Write("�����ⲿ���ݿ����������ʹ�ô˹���"):Response.End()
	End If
	Dim rs,username
	username=oblog.filt_badstr(Trim(Request("username")))
	If username="" Then Response.Write("���������Ϊ��"):Response.End()
	If is_ot_user=0 Then
		Set rs=oblog.Execute("Select username,password From Product_CoverProducts Where username='"&username&"'")
	Else
		Set rs=ot_conn.Execute("Select "&ot_username&","&ot_password&" From "&ot_usertable&" Where "&ot_username&"='"&username&"'")
	End If
	If Not rs.Eof Then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		Set rs=Nothing
		Response.Redirect("../HomestayBackctrl-indexHome.asp")
	Else
		Set rs=Nothing
		Response.Write("�޴����"):Response.End()
	End If
	
End Sub


Sub SaveModify()
	Dim rsCover,sqlCover
	
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		Exit Sub
	Else
		CoverID=Clng(CoverID)
	End If
	
	Dim CoverName,CoverDesc,Iclude_ProductIDs,ClassIDs,MaxNumber,CoverPrice,CoverPhoto,Discontinued
	
	CoverName			=Trim(Request("CoverName"))
	CoverDesc			=Trim(Request("CoverDesc"))
	Iclude_ProductIDs	=Trim(Request("Iclude_ProductIDs"))
	ClassIDs			=Trim(Request("ClassIDs"))
	MaxNumber			=Trim(Request("MaxNumber"))
	CoverPrice			=Trim(Request("CoverPrice"))
	CoverPhoto			=Trim(Request("CoverPhoto"))
	Discontinued		=Trim(Request("Discontinued"))
	

	If CoverName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>������Ʋ���Ϊ��</li>"
	Else
		CoverName = Trim(CoverName)
	End If

	If CoverDesc="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>�����������Ϊ��</li>"
	Else
		CoverDesc = Trim(CoverDesc)
	End If
	
	If Not Instr(Iclude_ProductIDs,",")>0 Then
		Errmsg=Errmsg & "<br><li>��ѡ��2������2�����ϵĲ�Ʒ����Ϊ�����</li>"
		founderr=True
	Else
		Iclude_ProductIDs = Trim(Iclude_ProductIDs)
	End If

	If Not Len(ClassIDs)>0 Then
		Errmsg=Errmsg & "<br><li>��ѡ������1������������࣡</li>"
		founderr=True
	Else
		ClassIDs = Trim(ClassIDs)
	End If
	
	If MaxNumber="" Or Not isNumeric(MaxNumber) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��� ��ѡ�������Ʒ���� ����Ϊ�գ�����ҪΪ���֣�</li>"
	Else
		'UnitPrice = FormatCurrency(UnitPrice)
		MaxNumber = Cint(MaxNumber)
	End If
	
	If CoverPrice="" Or Not isNumeric(CoverPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��� �����Żݼ۸� ����Ϊ�գ�����ҪΪ�������֣�</li>"
	Else
		CoverPrice = FormatCurrency(CoverPrice)
	End If
	
	
	If CoverPhoto="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>���ͼƬ�ϴ�����Ϊ��</li>"
	Else
		CoverPhoto = Trim(CoverPhoto)
	End If
	
	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��ѡ�� �Ƿ��Ѿ������������ ����״̬��</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	

	If founderr=True Then
		Set rsuser=Nothing
		Exit Sub
	End If
	
	Set rsCover=Server.CreateObject("Adodb.RecordSet")
	sqlCover="Select * From Product_CoverProducts Where CoverID=" & CoverID
	If Not IsObject(conn) Then link_database
	rsCover.Open sqlCover,Conn,1,3
	If rsCover.Bof And rsCover.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���������</li>"
		rsCover.Close
		Set rsCover=Nothing
		Exit Sub
	End If
	
	
	If Not CoverName="" Then  			rsCover("CoverName")			=CoverName
	If Not CoverDesc="" Then  			rsCover("CoverDesc")			=CoverDesc
	If Not Iclude_ProductIDs="" Then  	rsCover("Iclude_ProductIDs")	=Iclude_ProductIDs
	If Not ClassIDs="" Then  			rsCover("ClassIDs")				=ClassIDs
	If Not MaxNumber="" Then  			rsCover("MaxNumber")			=MaxNumber
	If Not CoverPrice="" Then  			rsCover("CoverPrice")			=CoverPrice
	If Not CoverPhoto="" Then  			rsCover("CoverPhoto")			=CoverPhoto
	If Not Discontinued="" Then  		rsCover("Discontinued")			=Discontinued
		
	
	rsCover.Update
	rsCover.Close
	Set rsCover=Nothing
	oblog.showok "�޸ĳɹ�!",""
End Sub


Sub DelUser()
	Dim rs,i
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ�������</li>"
		Exit Sub
	End If
	If Instr(CoverID,",")>0 Then
		CoverID=Split(CoverID,",")
		for i=0 to Ubound(CoverID)
			deloneuser(CoverID(i))
		next
	Else
		deloneuser(CoverID)
	End If
	Response.Redirect "Administrator_CoverProducts.asp"
End Sub

Sub deloneuser(CoverID)
	CoverID=Clng(CoverID)
'	Dim rs,fso,f,uname,udir
'	Set rs=oblog.Execute("Select CoverID From Product_CoverProducts Where CoverID="&CoverID)
'	If Not rs.Eof Then
'		udir=rs(0)
'		uname=rs(1)
'		Set fso=Server.CreateObject("scripting.filesystemobject")
'		If fso.FolderExists(Server.MapPath(blogdir & udir&"/"&rs("user_folder"))) Then 
'			Set f = fso.GetFolder(Server.MapPath(blogdir & udir&"/"&rs("user_folder")))
'			f.Delete True
'		End If
'		Set f=Nothing
'		Set fso=Nothing
'		oblog.Execute("Delete From oblog_log Where CoverID="&CoverID)
'		oblog.Execute("Delete From oblog_comment Where CoverID="&CoverID)
'		oblog.Execute("Delete From oblog_message Where CoverID="&CoverID)
'		oblog.Execute("Delete From oblog_subject Where CoverID="&CoverID)
'		oblog.Execute("Delete From Product_CoverProducts Where CoverID=" & CoverID)
'		oblog.Execute("Delete From oblog_upfile Where CoverID=" & CoverID)
		oblog.Execute("Delete From Product_CoverProducts Where CoverID=" & CoverID)
'		oblog.Execute("Update Product_CoverProducts Set Discontinued=1 Where CoverID=" & CoverID)
'	End If
'	Set rs=Nothing
End Sub

Sub LockUser()
	Dim rs,udir
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������</li>"
		Exit Sub
	End If
	CoverID=Clng(CoverID)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select lockuser,user_dir From Product_CoverProducts Where CoverID="&CoverID,conn,1,3
	If Not rs.Eof Then
		Dim strtmp,fso
		rs(0)=1
		udir=rs(1)
		rs.Update
		strtmp="<script language=javascript>"&vbcrlf
		strtmp=strtmp&" window.location.replace('"&blogdir&"err.asp?message=������Ѿ�������������ϵ����Ա!');"&vbcrlf
		strtmp=strtmp&"</script>"&vbcrlf
		If udir<>"" Then
		Set fso=Server.CreateObject("scripting.filesystemobject")
			udir=Server.MapPath(udir)
			If fso.FolderExists(udir&"/"&CoverID&"/inc") Then
				Call oblog.BuildFile(udir&"/"&CoverID&"/inc/chkblogpassword.htm",strtmp)								
			End If
			Set fso=Nothing
		End If
	End If
	rs.Close
	Set rs=Nothing
	oblog.showok "��������ɹ�",""
End Sub

Sub UnLockUser()
	Dim rs,udir
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������</li>"
		Exit Sub
	End If
	CoverID=Clng(CoverID)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select lockuser,user_dir From Product_CoverProducts Where CoverID="&CoverID,conn,1,3
	If Not rs.Eof Then
		Dim strtmp,fso
		rs(0)=0
		udir=rs(1)
		rs.Update
		strtmp=" "
		If udir<>"" Then
		Set fso=Server.CreateObject("scripting.filesystemobject")
			udir=Server.MapPath(udir)
			If fso.FolderExists(udir&"/"&CoverID&"/inc") Then
				Call oblog.BuildFile(udir&"/"&CoverID&"/inc/chkblogpassword.htm",strtmp)				
			End If			
			Set fso=Nothing
		End If
	End If
	rs.Close
	Set rs=Nothing
	oblog.showok "��������ɹ�",""
End Sub



Sub Discontinued()
	CoverID=Clng(CoverID)
	oblog.Execute("Update Product_CoverProducts Set Discontinued=1 Where CoverID="& CoverID)
	Response.Redirect "Administrator_CoverProducts.asp"
End Sub


Sub UnDiscontinued()
	CoverID=Clng(CoverID)
	oblog.Execute("Update Product_CoverProducts Set Discontinued=0 Where CoverID="& CoverID)
	Response.Redirect "Administrator_CoverProducts.asp"
End Sub



'ʹ�����ӣ�Call Admin_ShowClass_Option_classid(0,ParentID,"Product_Class")		'ParentID=classid.
Sub Admin_ShowClass_Option_classid(ShowType,CurrentID,Table_Class)
	If ShowType=0 Then
	    Response.Write "<option value='0'"
		If CurrentID=0 Then Response.Write " selected"
		Response.Write ">�ޣ���Ϊһ����Ŀ��</option>"
	End If
	Dim rsClass,sqlClass,strTemp,tmpDepth,i
	Dim arrShowLine(20)
	For i=0 to Ubound(arrShowLine)
		arrShowLine(i)=False
	Next
	sqlClass="Select * From "& Table_Class &" Where idType=" & t &" Order By RootID,OrderID"
						't="0"   tName="��Ʒ     /���
	Set rsClass=conn.Execute(sqlClass)
	If rsClass.Bof And rsClass.Eof Then 
		Response.Write "<option value=''>���������Ŀ</option>"
	Else
		Do While Not rsClass.Eof
			tmpDepth=rsClass("Depth")
			If rsClass("NextID")>0 Then
				arrShowLine(tmpDepth)=True
			Else
				arrShowLine(tmpDepth)=False
			End If
				strTemp="<option value='" & rsClass("classid") & "'"		'WL
			If CurrentID>0 And rsClass("classid")=CurrentID Then		'WL
				 strTemp=strTemp & " selected"
			End If
			strTemp=strTemp & ">"
			
			If tmpDepth>0 Then
				For i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					If i=tmpDepth Then
						If rsClass("NextID")>0 Then
							strTemp=strTemp & "��&nbsp;"
						Else
							strTemp=strTemp & "��&nbsp;"
						End If
					Else
						If arrShowLine(i)=True Then
							strTemp=strTemp & "��"
						Else
							strTemp=strTemp & "&nbsp;"
						End If
					End If
				Next
			End If
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			Response.Write strTemp
			rsClass.MoveNext
		Loop
	End If
	rsClass.Close
	Set rsClass=Nothing
End Sub


Sub WriteErrMsg()
	Dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	Response.Write strErr
End Sub





%>

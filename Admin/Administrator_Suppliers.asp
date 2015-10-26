<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
Const MaxPerPage=20
Dim strFileName
Dim totalPut,TotalPages
Dim rs, sql
Dim SupplierID,UserSearch,Keyword,strField
Dim Action,FoundErr,ErrMsg
Dim tmpDays
keyword=Trim(Request("keyword"))
If keyword<>"" Then 
	keyword=oblog.filt_badstr(keyword)
End If
strField=Trim(Request("Field"))
UserSearch=Trim(Request("UserSearch"))
Action=Trim(Request("Action"))
SupplierID=Trim(Request("SupplierID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")

If UserSearch="" Then
	UserSearch=0
Else
	UserSearch=Clng(UserSearch)
End If
strFileName="Administrator_Suppliers.asp?UserSearch=" & UserSearch
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
<title>��Ӧ�̹���</title>
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
    <td height="22" colspan=2 align=center><strong>��Ӧ�� ����</strong></td>
  </tr>
  <form name="form1" action="Administrator_Suppliers.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>���ٲ��ҹ�Ӧ�̣�</strong></td>
      <td width="687" height="30">
	   <select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">���ע���500����Ӧ��</option>
    
		  <option value="3">��ѯ�ѷ����Ĺ�Ӧ��</option>

        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="Administrator_Suppliers.asp">��Ӧ�̹�����ҳ</a>&nbsp;|&nbsp;<a href="?Action=Add">����¹�Ӧ��</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="Administrator_Suppliers.asp">
  <tr class="tdbg">
    <td width="120"><strong>��Ӧ�̸߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
		  <!--<option value="UserName" Selected>��Ӧ����</option>-->
		  <option value="SupplierName" Selected>����Ӧ����</option>
		  <option value="SupplierID" >����Ӧ��ID</option>
		  <option value="address" >����ַ</option>
		  <option value="beizhu" >����ע</option>
		  <option value="xuwei_beizhu" >����ϸ</option>

      </select>
	  
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10"> 
	  ��Ϊ�գ����ѯ���й�Ӧ��</td>
  </tr>
</form>
</table>
<%
If Action="Add" Then
	Call AddSupplier()
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
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Administrator_Suppliers.asp'>ע�ṩӦ�̹���</a>&nbsp;&gt;&gt;&nbsp;"
	Select Case UserSearch
		Case 0
			sql="Select Top 500 * From Product_Suppliers Where Discontinued=0 Order By SupplierID Desc"
			strGuide=strGuide & "���ע���500����Ӧ��"
		Case 1
			sql="Select Top 100 * From Product_Suppliers Where Discontinued=0 Order By log_count Desc"
			strGuide=strGuide & "������������ǰ100����Ӧ��"
		Case 2
			sql="Select Top 500 * From  Product_Suppliers Where Discontinued=1 Order By SupplierID Desc"
			strGuide=strGuide & "�ѷ����Ĺ�Ӧ��"
		Case 3
			sql="Select Top 500 * From  Product_Suppliers Where Discontinued=1 Order By SupplierID Desc"
			strGuide=strGuide & "�ѷ����Ĺ�Ӧ��"
		Case 4
			sql="Select * From  Product_Suppliers Where Discontinued=0 And user_level=9 Order By SupplierID Desc"
			strGuide=strGuide & "ǰ̨����Ա"
		Case 5
			sql="Select * From  Product_Suppliers Where Discontinued=0 And user_isbest=1 Order By SupplierID Desc"
			strGuide=strGuide & "�Ƽ���Ӧ��"
		Case 6
			sql="Select * From Product_Suppliers Where Discontinued=0 And User_Level=6 Order By SupplierID  Desc"
			strGuide=strGuide & "�ȴ�������֤֤�Ĺ�Ӧ��"
		Case 7
			sql="Select * From Product_Suppliers Where Discontinued=0 And  LockUser =1 Order By SupplierID  Desc"
			strGuide=strGuide & "���б���ס�Ĺ�Ӧ��"
		Case 10
			If Keyword="" Then
				sql="Select Top 500 * From Product_Suppliers Where Discontinued=0 Order By SupplierID Desc"
				strGuide=strGuide & "���й�Ӧ��"
			Else
				Select Case strField
				Case "SupplierID"

					If IsNumeric(Keyword)=false Then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>��Ӧ��ID������������</li>"
					Else
						sql="Select * From Product_Suppliers Where Discontinued=0 And SupplierID =" & Clng(Keyword)
						strGuide=strGuide & "��Ӧ��ID����<font color=red> " & Clng(Keyword) & " </font>�Ĺ�Ӧ��"
					End If
				Case "SupplierName"
					If is_sqldata=1 Then
						sql="Select * From Product_Suppliers Where Discontinued=0 And SupplierName like '%" & Keyword & "%' Order By SupplierID  Desc"
						strGuide=strGuide & "��Ӧ�����к��С� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					Else
						sql="Select * From Product_Suppliers Where Discontinued=0 And SupplierName= '" & Keyword & "' Order By SupplierID  Desc"
						strGuide=strGuide & "��Ӧ�������ڡ� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					End If
					
				Case "address"
					If is_sqldata=1 Then
						sql="Select * From Product_Suppliers Where Discontinued=0 And address like '%" & Keyword & "%' Order By SupplierID  Desc"
						strGuide=strGuide & "ͨѶ��ַ�к��С� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					Else
						sql="Select * From Product_Suppliers Where Discontinued=0 And address='" & Keyword & "' Order By SupplierID  Desc"
						strGuide=strGuide & "ͨѶ��ַ���ڡ� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					End If
				Case "beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product_Suppliers Where Discontinued=0 And beizhu like '%" & Keyword & "%' Order By SupplierID  Desc"
						strGuide=strGuide & "��ע�к��С� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					Else
						sql="Select * From Product_Suppliers Where Discontinued=0 And beizhu='" & Keyword & "' Order By SupplierID  Desc"
						strGuide=strGuide & "��ע���ڡ� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					End If
				Case "xuwei_beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product_Suppliers Where Discontinued=0 And xuwei_beizhu like '%" & Keyword & "%' Order By SupplierID  Desc"
						strGuide=strGuide & "��ϸ�к��С� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					Else
						sql="Select * From Product_Suppliers Where Discontinued=0 And xuwei_beizhu='" & Keyword & "' Order By SupplierID  Desc"
						strGuide=strGuide & "��ϸ���ڡ� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
					End If
				Case "logcount"
					sql="Select Top 500 * From Product_Suppliers Where Discontinued=0 And log_count < " & Clng(Keyword) & " Order By SupplierID  Desc"
					strGuide=strGuide & "������С�ڡ� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
				Case "logintimes"
					sql="Select Top 500 * From Product_Suppliers Where Discontinued=0 And logintimes < " & Clng(Keyword) & " Order By SupplierID  Desc"
					strGuide=strGuide & "��¼����С�ڡ� <font color=red>" & Keyword & "</font> ���Ĺ�Ӧ��"
				Case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="Select Top 500 * From Product_Suppliers Where Discontinued=0 And datediff(d,lastlogintime,getdate())>"&Clng(Keyword)&" Order By SupplierID  Desc"
				   Else 
				   	sql="Select Top 500 * From Product_Suppliers Where Discontinued=0 And datediff('d',lastlogintime,now())>"&Clng(Keyword)&" Order By SupplierID  Desc"
				   End If 
					strGuide=strGuide & "�� <font color=red>" & Keyword & "</font> ������δ��¼�Ĺ�Ӧ��"
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
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ����Ӧ��</td></tr></table>"
		Response.Write strGuide
	Else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ����Ӧ��</td></tr></table>"
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
        	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"����Ӧ��")
   	 	Else
   	     	If (currentPage-1)*MaxPerPage<totalPut Then
         	   	rs.Move  (currentPage-1)*MaxPerPage
         		Dim Bookmark
           		Bookmark=rs.Bookmark
            	showContent
            	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"����Ӧ��")
        	Else
	        	currentPage=1
           		showContent
           		Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"����Ӧ��")
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
  <form name="myform" method="Post" action="Administrator_Suppliers.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ѡ��</font></td>
            <td width="29" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="190" height="22" align="center"><font color="#FFFFFF"> ��Ӧ��</font></td>
			<td width="160" height="22" align="center"><font color="#FFFFFF">��������</font></td>
            <td width="95" height="22" align="center"><font color="#FFFFFF">����������</font></td>
            <td width="95" align="center"><font color="#FFFFFF">������ְ��</font></td>
			<td width="110" height="22" align="center"><font color="#FFFFFF">ͨѶ��ַ</font></td>
            
			<td width="145" align="center"><font color="#cccccc">��ϵ�绰</font></td>
            <td width="100" height="22" align="center"><font color="#FFFFFF">��ҳ��ַ</font></td>
            <td width="138" height="22" align="center"><font color="#FFFFFF">����</font></td>
          </tr>
          <%do while Not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td align="center"><input name='SupplierID' type='checkbox' onClick="unselectall()" id="SupplierID" value='<%=Cstr(rs("SupplierID"))%>' /></td>
            <td align="center"><%=rs("SupplierID")%></td>
            <td align="center"><%
			Response.Write "<a href='?action=Modify&SupplierID="&rs("SupplierID")&"'" 			
			Response.Write """  title='[��Ӧ�����ڣ�"& rs("Country") &"|"& rs("City") &"][����:"& rs("Fax") &"]'>" & rs("SupplierName") & "</a>"
			%> </td>
            <td align="center"> <%=rs("Region")%> </td>
			
			<td align="center"> <%=rs("ContactName")%> </td>
            
			<td align="center"> <%=rs("ContactTitle")%></td>
			
			<td align="center"> <%=rs("Address")%> </td>
            
			<td align="center"> <%=rs("Phone")%></td>
			
            <td align="center"><a href="<%=rs("HomePage")%>" target="_blank">HomePage</a></td>
            
			<td align="center">
			<%
			
			
			If rs("Discontinued")=1 Then
				Response.Write "<a href='Administrator_Suppliers.asp?Action=UnDiscontinued&SupplierID=" & rs("SupplierID") & "' title=''>"
				Response.Write("<font color=red title='�ҹ�˾�Ѿ��ʹ˹�Ӧ��ֹͣ�˽��׹�ϵ�������������ô˹�Ӧ�̣�' onClick='return confirm(""ȷ��Ҫ���ô˹�Ӧ����"");'>ֹͣ����</font>")
			ElseIf rs("Discontinued")=0 Then
				Response.Write "<a href='Administrator_Suppliers.asp?Action=Discontinued&SupplierID=" & rs("SupplierID") & "' title=''>"
				Response.Write("<font color=green title='�ҹ�˾�Ѿ��˹�Ӧ�����ڽ�����������֮�У���������ֹͣ����Ľ��׹�ϵ��' onClick='return confirm(""ȷ��Ҫֹͣ�ʹ˹�Ӧ�̵Ľ��׹�ϵ��"");'>��ϵ����</font>")
			End If
			Response.Write "</font></a>&nbsp;"
			
		Response.Write "<a href='Administrator_Suppliers.asp?Action=Modify&SupplierID=" & rs("SupplierID") & "'>����</a>&nbsp;"
		
		'Response.Write "<a href='Administrator_Suppliers.asp?Action=gouser2&username=" & rs("username") & "' target='blank'>����̨</a>&nbsp;"

        Response.Write "<a href='Administrator_Suppliers.asp?Action=Del&SupplierID=" & rs("SupplierID") & "' onClick='return confirm(""ȷ��Ҫɾ���˹�Ӧ����"");'>ɾ��</a>&nbsp;"
		
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
              ѡ�б�ҳ��ʾ�����й�Ӧ��</td>
            <td> <strong>������</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=True" />
              ɾ��&nbsp;&nbsp;&nbsp;&nbsp; 
              <!--<input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">�ƶ���
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">��Ӧ��</option>
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




Sub AddSupplier()

%>
<FORM name="Form1" Action="Administrator_Suppliers.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">ע���� ��Ӧ��</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��Ӧ������<br />(��Ӧ��e��˾���������)</TD>
      <TD width="60%"><input name="SupplierName" type="text" value="" size=40 maxlength=40 /></TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%">��������</TD>
      <TD width="60%">
	  <input name="Region" type="text" value="" size=30 maxlength=30 />
	  		<!--<select name="Region">
	  			<option value="0">��ѡ����������</option>
		  	</select>-->
	  </TD>
    </TR>
	
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>���������֣�</td>
      <td><input name="ContactName" type="text" value="" size=30 maxlength=30 /></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>������ְ��</td>
      <td><input name="ContactTitle" type="text" value="" size=30 maxlength=30 /></td>
    </tr>
	
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ҳ��У�</td>
      <td><select name="Country">
	  		<option value="0">��ѡ�����</option>
		  </select>
		  
		  &nbsp;&nbsp;&nbsp;
		  
		  <select name="City">
	  		<option value="0">��ѡ�����</option>
		  </select>	  </td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>ͨѶ��ַ��</td>
      <td><input name="Address" type="text" value="" size=60 maxlength=60 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�������룺</td>
      <td><input name="PostalCode" type="text" value="" size=10 maxlength=10 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��ϵ�绰��</td>
      <td><input name="Phone" type="text" value="" size=24 maxlength=24 /></td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>������룺</TD>
      <TD><input name="Fax" type="text" value="" size=24 maxlength=24 />	  </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD>��ҳ��ַ��</TD>
      <TD><input name="HomePage" type="text" value="" size=50 maxlength=255 />	  </TD>
    </TR>
	
	
    <TR class="tdbg" > 
      <TD width="40%">��Ӧ��e״̬��</TD>
      <TD width="60%"><INPUT name="Discontinued" type=radio value="0" checked="checked" />��ϵ����&nbsp;&nbsp;&nbsp;<INPUT name="Discontinued" type=radio value="1" />ֹͣ����      </TD>
    </TR>
   
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name=Submit   type=submit id="Submit" value="�����¹�Ӧ��"></TD>
    </TR>
  </TABLE>
</form>
<%

End Sub




Sub Modify()
	Dim SupplierID
	Dim rsSupplier,sqlSupplier
	SupplierID=Trim(Request("SupplierID"))
	If SupplierID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		Exit Sub
	Else
		SupplierID=Clng(SupplierID)
	End If
	Set rsSupplier=Server.CreateObject("Adodb.RecordSet")
	sqlSupplier="Select * From Product_Suppliers Where SupplierID=" & SupplierID
	If Not IsObject(conn) Then link_database
	rsSupplier.Open sqlSupplier,Conn,1,3
	If rsSupplier.Bof And rsSupplier.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���Ĺ�Ӧ�̣�</li>"
		rsSupplier.Close
		Set rsSupplier=Nothing
		Exit Sub
	End If
%>
<FORM name="Form1" Action="Administrator_Suppliers.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">���� ��Ӧ��e����</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��Ӧ������<br />(��Ӧ��e��˾���������)</TD>
      <TD width="60%"><input name="SupplierName" type="text" value="<%=rsSupplier("SupplierName")%>" size=40 maxlength=40 /></TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%">��������</TD>
      <TD width="60%">
	  		<input name="Region" type="text" value="<%=rsSupplier("Region")%>" size=30 maxlength=30 />
	  </TD>
    </TR>
	
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>���������֣�</td>
      <td><input name="ContactName" type="text" value="<%=rsSupplier("ContactName")%>" size=30 maxlength=30 /></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>������ְ��</td>
      <td><input name="ContactTitle" type="text" value="<%=rsSupplier("ContactTitle")%>" size=30 maxlength=30 /></td>
    </tr>
	
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ҳ��У�</td>
      <td><select name="Country">
	  		<option value="0">��ѡ�����</option>
			<%'Call Admin_ShowClass_Option_classid(0, rsCover("ClassIDs"), "Product_Class")		���Ϳؼ�%>
		  </select>
		  
		  &nbsp;&nbsp;&nbsp;
		  
		  <select name="City">
	  		<option value="0">��ѡ�����</option>
			<%'Call Admin_ShowClass_Option_classid(0, rsCover("ClassIDs"), "Product_Class")		���Ϳؼ�%>
		  </select>	  </td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>ͨѶ��ַ��</td>
      <td><input name="Address" type="text" value="<%=rsSupplier("Address")%>" size=60 maxlength=60 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�������룺</td>
      <td><input name="PostalCode" type="text" value="<%=rsSupplier("PostalCode")%>" size=10 maxlength=10 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��ϵ�绰��</td>
      <td><input name="Phone" type="text" value="<%=rsSupplier("Phone")%>" size=24 maxlength=24 /></td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>������룺</TD>
      <TD><input name="Fax" type="text" value="<%=rsSupplier("Fax")%>" size=24 maxlength=24 />	  </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD>��ҳ��ַ��</TD>
      <TD><input name="HomePage" type="text" value="<%=rsSupplier("HomePage")%>" size=50 maxlength=50 />	  </TD>
    </TR>
	
	
    <TR class="tdbg" > 
      <TD width="40%">��Ӧ��e״̬��</TD>
      <TD width="60%"><INPUT name="Discontinued" type=radio value="0" <% If rsSupplier("Discontinued")=0 Then Response.Write "checked" %> />��ϵ����
	  &nbsp;&nbsp;&nbsp;
	  <INPUT name="Discontinued" type=radio value="1" <% If rsSupplier("Discontinued")=1 Then Response.Write "checked" %> />ֹͣ����      </TD>
    </TR>
   
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center">
	  <input name="Action" type="hidden" id="Action" value="SaveModify"> 
	  <input name=Submit   type=submit id="Submit" value="���湩Ӧ�̸��½��"> 
	  <input name="SupplierID" type="hidden" id="SupplierID" value="<%=rsSupplier("SupplierID")%>">
	  </TD>
    </TR>
  </TABLE>
</form>
<%
	rsSupplier.Close
	Set rsSupplier=Nothing
End Sub



Sub gouser1()
If is_ot_user=1 Then
	Response.Write("�����ⲿ���ݿ⹩Ӧ�̣�����ʹ�ô˹���"):Response.End()
End If
%>
<FORM name="Form1" action="Administrator_Suppliers.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">��¼����Ŀ�����̨</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          ������������Ա��¼����Ŀ�Ĺ��������й���<br>
          ������Ŀ���������ϰ�ʱ���ɽ���ù�Ӧ�̺�̨��<br>
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
	Dim rsSupplier,sqlSupplier
	Dim SupplierName,ContactName,ContactTitle,Address,Country,City,Region,PostalCode,Phone,Fax,HomePage,Discontinued
	
	SupplierName	=Trim(Request("SupplierName"))
	ContactName		=Trim(Request("ContactName"))
	ContactTitle	=Trim(Request("ContactTitle"))
	Address			=Trim(Request("Address"))
	Country			=Trim(Request("Country"))
	City			=Trim(Request("City"))
	Region			=Trim(Request("Region"))
	PostalCode		=Trim(Request("PostalCode"))
	Phone			=Trim(Request("Phone"))
	Fax				=Trim(Request("Fax"))
	HomePage		=Trim(Request("HomePage"))
	Discontinued	=Trim(Request("Discontinued"))
	

	If SupplierName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ӧ����(��˾�������)����Ϊ��</li>"
	Else
		SupplierName = Trim(SupplierName)
	End If

	If ContactName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ӧ�����������ֲ���Ϊ��</li>"
	Else
		ContactName = Trim(ContactName)
	End If
	
'	If ContactTitle="" Then
'		founderr=True
'		Errmsg=Errmsg & "<br><li>��Ӧ��������ְ�Ʋ���Ϊ��</li>"
'	Else
'		ContactTitle = Trim(ContactTitle)
'	End If
	
	If Phone="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ӧ�̵���ϵ�绰����Ϊ��</li>"
	Else
		Phone = Trim(Phone)
	End If

	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��ѡ�� ��Ӧ��e״̬��</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	

	If founderr=True Then
		Set rs=Nothing
		Exit Sub
	End If
	
	
	Set rsSupplier=Server.CreateObject("Adodb.RecordSet")
	sqlSupplier="Select * From Product_Suppliers"
	If Not IsObject(conn) Then link_database
	rsSupplier.Open sqlSupplier,Conn,1,3

	rsSupplier.Addnew
	
		If Not SupplierName="" Then  	rsSupplier("SupplierName")	=SupplierName
		If Not ContactName="" Then  	rsSupplier("ContactName")	=ContactName
		If Not ContactTitle="" Then  	rsSupplier("ContactTitle")	=ContactTitle
		If Not Address="" Then  		rsSupplier("Address")		=Address
		If Not Country="" Then  		rsSupplier("Country")		=Country
		If Not City="" Then  			rsSupplier("City")			=City
		If Not Region="" Then  			rsSupplier("Region")		=Region
		If Not PostalCode="" Then  		rsSupplier("PostalCode")	=PostalCode
		If Not Phone="" Then  			rsSupplier("Phone")			=Phone
		If Not Fax="" Then  			rsSupplier("Fax")			=Fax
		If Not HomePage="" Then  		rsSupplier("HomePage")		=HomePage
		If Not Discontinued="" Then  	rsSupplier("Discontinued")	=Discontinued

		
	rsSupplier.Update
	
	rsSupplier.Close
	Set rsSupplier=Nothing
	
	oblog.showok "ע���¹�Ӧ�̳ɹ�!",""
End Sub



Sub gouser2()
	If is_ot_user=1 Then
		Response.Write("�����ⲿ���ݿ⹩Ӧ�̣�����ʹ�ô˹���"):Response.End()
	End If
	Dim rs,username
	username=oblog.filt_badstr(Trim(Request("username")))
	If username="" Then Response.Write("��Ӧ��������Ϊ��"):Response.End()
	If is_ot_user=0 Then
		Set rs=oblog.Execute("Select username,password From Product_Suppliers Where username='"&username&"'")
	Else
		Set rs=ot_conn.Execute("Select "&ot_username&","&ot_password&" From "&ot_usertable&" Where "&ot_username&"='"&username&"'")
	End If
	If Not rs.Eof Then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		Set rs=Nothing
		Response.Redirect("../HomestayBackctrl-indexHome.asp")
	Else
		Set rs=Nothing
		Response.Write("�޴˹�Ӧ��"):Response.End()
	End If
	
End Sub


Sub SaveModify()
	Dim rsSupplier,sqlSupplier
	
	If SupplierID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		Exit Sub
	Else
		SupplierID=Clng(SupplierID)
	End If
	
	Dim SupplierName,ContactName,ContactTitle,Address,Country,City,Region,PostalCode,Phone,Fax,HomePage,Discontinued
	
	SupplierName	=Trim(Request("SupplierName"))
	ContactName		=Trim(Request("ContactName"))
	ContactTitle	=Trim(Request("ContactTitle"))
	Address			=Trim(Request("Address"))
	Country			=Trim(Request("Country"))
	City			=Trim(Request("City"))
	Region			=Trim(Request("Region"))
	PostalCode		=Trim(Request("PostalCode"))
	Phone			=Trim(Request("Phone"))
	Fax				=Trim(Request("Fax"))
	HomePage		=Trim(Request("HomePage"))
	Discontinued	=Trim(Request("Discontinued"))
	

	If SupplierName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ӧ����(��˾�������)����Ϊ��</li>"
	Else
		SupplierName = Trim(SupplierName)
	End If

	If ContactName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ӧ�����������ֲ���Ϊ��</li>"
	Else
		ContactName = Trim(ContactName)
	End If
	
'	If ContactTitle="" Then
'		founderr=True
'		Errmsg=Errmsg & "<br><li>��Ӧ��������ְ�Ʋ���Ϊ��</li>"
'	Else
'		ContactTitle = Trim(ContactTitle)
'	End If
	
	If Phone="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ӧ�̵���ϵ�绰����Ϊ��</li>"
	Else
		Phone = Trim(Phone)
	End If

	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��ѡ�� ��Ӧ��e״̬��</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	

	If founderr=True Then
		Set rsuser=Nothing
		Exit Sub
	End If
	
	Set rsSupplier=Server.CreateObject("Adodb.RecordSet")
	sqlSupplier="Select * From Product_Suppliers Where SupplierID=" & SupplierID
	If Not IsObject(conn) Then link_database
	rsSupplier.Open sqlSupplier,Conn,1,3
	If rsSupplier.Bof And rsSupplier.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���Ĺ�Ӧ�̣�</li>"
		rsSupplier.Close
		Set rsSupplier=Nothing
		Exit Sub
	End If
	
	
	If Not SupplierName="" Then  	rsSupplier("SupplierName")	=SupplierName
	If Not ContactName="" Then  	rsSupplier("ContactName")	=ContactName
	If Not ContactTitle="" Then  	rsSupplier("ContactTitle")	=ContactTitle
	If Not Address="" Then  		rsSupplier("Address")		=Address
	If Not Country="" Then  		rsSupplier("Country")		=Country
	If Not City="" Then  			rsSupplier("City")			=City
	If Not Region="" Then  			rsSupplier("Region")		=Region
	If Not PostalCode="" Then  		rsSupplier("PostalCode")	=PostalCode
	If Not Phone="" Then  			rsSupplier("Phone")			=Phone
	If Not Fax="" Then  			rsSupplier("Fax")			=Fax
	If Not HomePage="" Then  		rsSupplier("HomePage")		=HomePage
	If Not Discontinued="" Then  	rsSupplier("Discontinued")	=Discontinued
		
	
	rsSupplier.Update
	rsSupplier.Close
	Set rsSupplier=Nothing
	oblog.showok "���¹�Ӧ�����ϳɹ�!",""
End Sub


Sub DelUser()
	Dim rs,i
	If SupplierID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ���Ĺ�Ӧ��</li>"
		Exit Sub
	End If
	If Instr(SupplierID,",")>0 Then
		SupplierID=Split(SupplierID,",")
		for i=0 to Ubound(SupplierID)
			deloneuser(SupplierID(i))
		next
	Else
		deloneuser(SupplierID)
	End If
	Response.Redirect "Administrator_Suppliers.asp"
End Sub

Sub deloneuser(SupplierID)
	SupplierID=Clng(SupplierID)
'	Dim rs,fso,f,uname,udir
'	Set rs=oblog.Execute("Select SupplierID,SupplierName,ContactName From Product_Suppliers Where SupplierID="&SupplierID)
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
'		oblog.Execute("Delete From oblog_log Where SupplierID="&SupplierID)
'		oblog.Execute("Delete From oblog_comment Where SupplierID="&SupplierID)
'		oblog.Execute("Delete From oblog_message Where SupplierID="&SupplierID)
'		oblog.Execute("Delete From oblog_subject Where SupplierID="&SupplierID)
'		oblog.Execute("Delete From Product_Suppliers Where SupplierID=" & SupplierID)
'		oblog.Execute("Delete From oblog_upfile Where SupplierID=" & SupplierID)
		oblog.Execute("Delete From Product_Suppliers Where SupplierID=" & SupplierID)
		'oblog.Execute("Update Product_Suppliers Set Discontinued=1 Where SupplierID=" & SupplierID)
'	End If
'	Set rs=Nothing
End Sub

Sub LockUser()
	Dim rs,udir
	If SupplierID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�����Ĺ�Ӧ��</li>"
		Exit Sub
	End If
	SupplierID=Clng(SupplierID)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select lockuser,user_dir From Product_Suppliers Where SupplierID="&SupplierID,conn,1,3
	If Not rs.Eof Then
		Dim strtmp,fso
		rs(0)=1
		udir=rs(1)
		rs.Update
		strtmp="<script language=javascript>"&vbcrlf
		strtmp=strtmp&" window.location.replace('"&blogdir&"err.asp?message=�˹�Ӧ���Ѿ�������������ϵ����Ա!');"&vbcrlf
		strtmp=strtmp&"</script>"&vbcrlf
		If udir<>"" Then
		Set fso=Server.CreateObject("scripting.filesystemobject")
			udir=Server.MapPath(udir)
			If fso.FolderExists(udir&"/"&SupplierID&"/inc") Then
				Call oblog.BuildFile(udir&"/"&SupplierID&"/inc/chkblogpassword.htm",strtmp)								
			End If
			Set fso=Nothing
		End If
	End If
	rs.Close
	Set rs=Nothing
	oblog.showok "������Ӧ�̳ɹ�",""
End Sub

Sub UnLockUser()
	Dim rs,udir
	If SupplierID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�����Ĺ�Ӧ��</li>"
		Exit Sub
	End If
	SupplierID=Clng(SupplierID)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select lockuser,user_dir From Product_Suppliers Where SupplierID="&SupplierID,conn,1,3
	If Not rs.Eof Then
		Dim strtmp,fso
		rs(0)=0
		udir=rs(1)
		rs.Update
		strtmp=" "
		If udir<>"" Then
		Set fso=Server.CreateObject("scripting.filesystemobject")
			udir=Server.MapPath(udir)
			If fso.FolderExists(udir&"/"&SupplierID&"/inc") Then
				Call oblog.BuildFile(udir&"/"&SupplierID&"/inc/chkblogpassword.htm",strtmp)				
			End If			
			Set fso=Nothing
		End If
	End If
	rs.Close
	Set rs=Nothing
	oblog.showok "������Ӧ�̳ɹ�",""
End Sub



Sub Discontinued()
	SupplierID=Clng(SupplierID)
	oblog.Execute("Update Product_Suppliers Set Discontinued=1 Where SupplierID="& SupplierID)
	Response.Redirect "Administrator_Suppliers.asp"
End Sub


Sub UnDiscontinued()
	SupplierID=Clng(SupplierID)
	oblog.Execute("Update Product_Suppliers Set Discontinued=0 Where SupplierID="& SupplierID)
	Response.Redirect "Administrator_Suppliers.asp"
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

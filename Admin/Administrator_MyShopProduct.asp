<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
const MaxPerPage=20
Dim strFileName
Dim totalPut,TotalPages
Dim rs, sql
Dim ProductID,UserSearch,Keyword,strField
Dim Action,FoundErr,ErrMsg
Dim tmpDays
keyword=Trim(Request("keyword"))
If keyword<>"" Then 
	keyword=oblog.filt_badstr(keyword)
End If
strField=Trim(Request("Field"))
UserSearch=Trim(Request("UserSearch"))
Action=Trim(Request("Action"))
ProductID=Trim(Request("ProductID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")

If UserSearch="" Then
	UserSearch=0
Else
	UserSearch=Clng(UserSearch)
End If
strFileName="Administrator_MyShopProduct.asp?UserSearch=" & UserSearch
If strField<>"" Then
	strFileName=strFileName&"&Field="&strField
End If
If keyword<>"" Then
	strFileName=strFileName&"&keyword="&keyword
End If
If Request("page")<>"" Then
    currentPage=cint(Request("page"))
Else
	currentPage=1
End If

%>
<html>
<head>
<title>ע���Ʒ����</title>
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
  For (var i=0;i<form.elements.length;i++)
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
    <td height="22" colspan=2 align=center><strong>ע �� �� Ʒ �� ��</strong></td>
  </tr>
  <form name="form1" Action="Administrator_MyShopProduct.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>���ٲ��Ҳ�Ʒ��</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
        <option value=>��ѡ���ѯ����</option>
        <option value="0">���ע���500����Ʒ</option>
        <option value="1">����������Ĳ�Ʒ</option>
        <option value="3">�Ѿ������Ĳ�Ʒ</option>
      </select>        &nbsp;&nbsp;&nbsp;&nbsp;<a href="Administrator_MyShopProduct.asp">ע���Ʒ������ҳ</a>&nbsp;|&nbsp;<a href="Administrator_MyShopProduct.asp?Action=Add">����²�Ʒ</a></td>
    </tr>
  </form>
  <form name="form2" method="post" Action="Administrator_MyShopProduct.asp">
  <tr class="tdbg">
    <td width="120"><strong>��Ʒ�߼���ѯ��</strong></td>
    <td >
      <Select name="Field" id="Field">
		  
		  <option value="ProductName" selected>����Ʒ����</option>
		  <option value="ProductID" >����ƷID</option>
		  <option value="address" >����Ʒ�����</option>
		  <option value="beizhu" >����ƷС����</option>
		  <option value="xuwei_beizhu" >����Ӧ��ID</option>
		  <option value="lastlogintime" >����Ʒ����</option>
	
		  <option value="logintimes" >����Ʒ����</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10"> 
	  ��Ϊ�գ����ѯ���в�Ʒ</td>
  </tr>
</form>
</table>
<%
If Action="Add" Then
	Call AddProduct()
ElseIf Action="SaveAdd" Then
	Call SaveAdd()
ElseIf Action="Modify" Then
	Call Modify()
ElseIf Action="SaveModify" Then
	Call SaveModify()
ElseIf Action="Del" Then
	Call DelProduct()
ElseIf Action="Lock" Then
	Call Discontinued()
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
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Administrator_MyShopProduct.asp'>ע��ע���Ʒ����</a>&nbsp;&gt;&gt;&nbsp;"
	Select Case UserSearch
		Case 0
			sql="Select Top 500 * From Product Where Discontinued=0 Order By ProductID Desc"
			strGuide=strGuide & "���ע���500����Ʒ"
		Case 1
			sql="Select Top 100 * From Product Where Discontinued=0 Order By ProductID Desc"
			strGuide=strGuide & "����������Ĳ�Ʒ"
		Case 2
			sql="Select Top 100 * From  Product Where Discontinued=1 Order By ProductID Desc"
			strGuide=strGuide & "�Ѿ������Ĳ�Ʒ"
		Case 3
			sql="Select * From  Product Where Discontinued=1 Order By ProductID Desc"
			strGuide=strGuide & "�Ѿ������Ĳ�Ʒ"
		Case 4
			sql="Select * From  Product Where Discontinued=0 And user_level=9 Order By ProductID Desc"
			strGuide=strGuide & "�Ѿ������Ĳ�Ʒ"
		Case 5
			sql="Select * From  Product Where Discontinued=0 And user_isbest=1 Order By ProductID Desc"
			strGuide=strGuide & "�Ѿ������Ĳ�Ʒ"
		Case 6
			sql="Select * From Product Where Discontinued=0 And User_Level=6 Order By ProductID  Desc"
			strGuide=strGuide & "�Ѿ������Ĳ�Ʒ"
		Case 7
			sql="Select * From Product Where Discontinued=0 And  Discontinued =1 Order By ProductID  Desc"
			strGuide=strGuide & "�Ѿ������Ĳ�Ʒ"
		Case 10
			If Keyword="" Then
				sql="Select Top 500 * From Product Where Discontinued=0 Order By ProductID Desc"
				strGuide=strGuide & "���в�Ʒ"
			Else
				Select Case strField
				Case "ProductID"
'				If Instr(1,Keyword,"MH00")>0 Then Keyword=replace(Keyword,"MH00","")
'				If Instr(1,Keyword,"mh00")>0 Then Keyword=replace(Keyword,"mh00","")
				'''Response.Write Keyword
					If IsNumeric(Keyword)=False Then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>��ƷID������������</li>"
					Else
						sql="Select * From Product Where Discontinued=0 And ProductID=" & Clng(Keyword)
						strGuide=strGuide & "��ƷID����<font color=red>" & Clng(Keyword) & " </font>�Ĳ�Ʒ"
					End If
				Case "ProductName"
					If is_sqldata=1 Then
						sql="Select * From Product Where Discontinued=0 And ProductName Like '%" & Keyword & "%' Order By ProductID Desc"
						strGuide=strGuide & "��Ʒ�����к��С� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					Else
						sql="Select * From Product Where Discontinued=0 And ProductName='" & Keyword & "' Order By ProductID Desc"
						strGuide=strGuide & "��Ʒ���Ƶ��ڡ� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					End If
					
				Case "address"
					If is_sqldata=1 Then
						sql="Select * From Product Where Discontinued=0 And address Like '%" & Keyword & "%' Order By ProductID  Desc"
						strGuide=strGuide & "��Ʒ������к��С� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					Else
						sql="Select * From Product Where Discontinued=0 And address='" & Keyword & "' Order By ProductID  Desc"
						strGuide=strGuide & "��Ʒ�������ڡ� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					End If
				Case "beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product Where Discontinued=0 And beizhu Like '%" & Keyword & "%' Order By ProductID  Desc"
						strGuide=strGuide & "��ƷС�����к��С� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					Else
						sql="Select * From Product Where Discontinued=0 And beizhu='" & Keyword & "' Order By ProductID  Desc"
						strGuide=strGuide & "��ƷС������ڡ� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					End If
				Case "xuwei_beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product Where Discontinued=0 And xuwei_beizhu Like '%" & Keyword & "%' Order By ProductID  Desc"
						strGuide=strGuide & "����Ӧ��ID�к��С� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					Else
						sql="Select * From Product Where Discontinued=0 And xuwei_beizhu='" & Keyword & "' Order By ProductID  Desc"
						strGuide=strGuide & "����Ӧ��ID���ڡ� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
					End If
				Case "logcount"
					sql="Select Top 500 * From Product Where Discontinued=0 And log_count < " & Clng(Keyword) & " Order By ProductID  Desc"
					strGuide=strGuide & "���� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
				Case "logintimes"
					sql="Select Top 500 * From Product Where Discontinued=0 And logintimes < " & Clng(Keyword) & " Order By ProductID  Desc"
					strGuide=strGuide & "���� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
				Case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="Select Top 500 * From Product Where Discontinued=0 And datediff(d,lastlogintime,getdate())>"&Clng(Keyword)&" Order By ProductID  Desc"
				   Else 
				   	sql="Select Top 500 * From Product Where Discontinued=0 And datediff('d',lastlogintime,now())>"&Clng(Keyword)&" Order By ProductID  Desc"
				   End If 
					strGuide=strGuide & "���۵��ڣ��� <font color=red>" & Keyword & "</font> ���Ĳ�Ʒ"
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
'	Response.Write sql&"--"
'	Response.End 
	rs.Open sql,Conn,1,1
  	If rs.Eof And rs.Bof Then
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ����Ʒ</td></tr></table>"
		Response.Write strGuide
	Else
    	totalPut=rs.RecordCount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ����Ʒ</td></tr></table>"
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
        	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"����Ʒ")
   	 	Else
   	     	If (currentPage-1)*MaxPerPage<totalPut Then
         	   	rs.Move  (currentPage-1)*MaxPerPage
         		Dim Bookmark
           		Bookmark=rs.Bookmark
            	showContent
            	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"����Ʒ")
        	Else
	        	currentPage=1
           		showContent
           		Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"����Ʒ")
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
  <form name="myform" method="Post" Action="Administrator_MyShopProduct.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ѡ��</font></td>
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="268" height="22" align="center"><font color="#FFFFFF">��Ʒ����</font></td>
            <td width="109" height="22" align="center"><font color="#FFFFFF">������</font></td>
            <td width="140" align="center"><font color="#FFFFFF">С����</font></td>
			<td width="210" height="22" align="center"><font color="#FFFFFF">��Ӧ��ID</font></td>
            
			<td width="68" align="center"><font color="#cccccc">�ּ�</font></td>
            <td width="84" height="22" align="center"><font color="#FFFFFF">�г���</font></td>
            <td width="217" height="22" align="center"><font color="#FFFFFF"> 
              ����</font></td>
          </tr>
          <%Do While Not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><input name='ProductID' type='checkbox' onClick="unselectall()" id="ProductID" value='<%=Cstr(rs("ProductID"))%>'></td>
            <td width="30" align="center"><%=rs("ProductID")%></td>
            <td width="268" align="center"><%
			Response.Write "<a href='?Action=Modify&ProductID="&rs("ProductID")&"'" 			
			Response.Write """  title='[ע������:"&rs("AddDate")&"]'>" & rs("ProductName") & "</a>"
			%> </td>
            <td align="center"><%=rs("ClassID")%> </td>
            
			<td align="center"><%=rs("ClassChildID")%></td>
			
			<td align="center"><%=rs("SupplierID")%> </td>
            
			<td align="center"><%=FormatCurrency(rs("UnitPrice"))%></td>
			
            <td width="84" align="center"><%=FormatCurrency(rs("UnitPrice_Market"))%></td>
            <td width="217" align="center">
			

			
			<%
			
		Response.Write "<a href='Administrator_MyShopProduct.asp?Action=Modify&ProductID=" & rs("ProductID") & "'>�޸�</a>&nbsp;"


		If rs("Discontinued")=0 Then
			Response.Write "<a href='Administrator_MyShopProduct.asp?Action=Discontinued&ProductID="& rs("ProductID") &"'><font color=green>����������</font></a>&nbsp;"
		ElseIf rs("Discontinued")=1 Then
            Response.Write "<a href='Administrator_MyShopProduct.asp?Action=UnDiscontinued&ProductID="& rs("ProductID") &"'><font color=red><font color=red>�ѷ���</font></font></a>&nbsp;"
		End If
		
		
        Response.Write "<a href='Administrator_MyShopProduct.asp?Action=Del&ProductID=" & rs("ProductID") & "' onClick='return confirm(""ȷ��Ҫɾ���˲�Ʒ��"");'>ɾ��</a>&nbsp;"

		%></td>
          </tr>
          <%
	i=i+1
	If i>=MaxPerPage Then Exit Do
	rs.MoveNext
Loop
%>
        </table>  
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ��ʾ�����в�Ʒ</td>
            <td> <strong>������</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=True">
              ɾ��&nbsp;&nbsp;&nbsp;&nbsp; 
              <!--<input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=False">�ƶ���
              <Select name="User_Level" id="User_Level" disabled>
                <option value="6">ע���Ʒ</option>
                <option value="7">ע����Ŀ����</option>
                <option value="8">VIP��Ŀ����</option>
                <option value="9">ǰ̨������</option>
              </select>-->
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" ִ �� "> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
End Sub



Sub AddProduct()

%>
<FORM name="Form1" Action="Administrator_MyShopProduct.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">ע���²�Ʒ��Ϣ</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��Ʒ����</TD>
      <TD width="60%"><input name="ProductName" type="text" value="" size=20 maxlength=20 /></TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>���������ͣ�</td>
      <td><Select name="ClassID" >
	  		<%Call Admin_ShowClass_Option_classid(0,0,"Product_Class")%>

		  </select></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>����С���ͣ�</td>
      <td><Select name="ClassChildID" >
	  		<%Call Admin_ShowClass_Option_classid(0,0,"Product_ClassChild")%>
		  </select></td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ӧ��ID��</td>
      <td><Select name="SupplierID" >
	  		<%Call Admin_Show_Option_keyid(0,0,"Product_Suppliers","SupplierName","SupplierID")%>

		  </select></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����ּۣ�</td>
      <td><input name="UnitPrice" type="text" value="" size=6 maxlength=20 style="text-align:center;" />ԪRMB</td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����г��ۣ�</td>
      <td><input name="UnitPrice_Market" type="text" value="" size=6 maxlength=20 style="text-align:center;" />ԪRMB</td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>����ÿ��λ��Ʒ����������</TD>
      <TD><input name="QuantityPerUnit" type="text" value="1" size=2 maxlength=10 /> 
	  
	  ��ѡ��λ��
	  <Select name="QuantityPerUnit_Unit" >
	  	<option value="��">��</option>
		<option value="g">g</option>
		<option value="kg">kg</option>
		<option value="ml">ml</option>
		<option value="L">L</option>
	  </select>
	  </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��ǰ�������</TD>
      <TD width="60%"><INPUT name="UnitsInStock" type=text value="" size=10 maxlength=20 />(+) 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��ǰ��������</TD>
      <TD width="60%"><input name="UnitsOnOrder" type="text" value="0" size=10 maxlength=20 />(-) </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">ȡ��������������Ʒ������</TD>
      <TD width="60%"><input name="UnitsOnOrder_del" type="text" value="0" size=10 maxlength=20 />(+)</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�Ƿ��Ѿ������˲�Ʒ��<br />
	  (<font color="red">�����ϻ���</font>)
	  </TD>
      <TD width="60%"><font color="red"><input type="radio" name="Discontinued" value=1  />
        ��</font> &nbsp;&nbsp; <input type="radio" name="Discontinued" value=0 checked="checked" />
        ��
      </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��ɫ���ƣ�</TD>
      <TD width="60%"><input name="ColorName" type="text" value="" size=10 maxlength=20 /></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�ϴ���ɫͼƬ��</TD>
      <TD width="60%"><input name="ColorPhoto" type="text" value="***.jpg" size=30 maxlength=500 /></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">���ID���ߴ磩��</TD>
      <TD width="60%"><Select name="StandardID">
	  					  <%Call Admin_ShowClass_Option_classid(0,0,"Product_Standard")%>

					  </select></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ļ��������<br />
	  �磺ֱ��_�ۺ�ϸ����
	  </td>
      <td><input name="StandardDesc" type="text" value="" size=10 maxlength=20 /></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ʒ��������</td>
      <td>
	  		<textarea name="Description" cols="38" rows="8"></textarea>
	  </td>
    </tr>
    <TR class="tdbg" > 
      <TD>�ϴ���Ʒ��ʵ��ͼƬ��</TD>
      <TD><input name="Photo" type="text" value="***.jpg" size=30 maxlength=500 /></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">����ʱ�䣺</TD>
      <TD width="60%"><INPUT name="ComeInToTheMarketDate" value="<%=Now()%>" size=7 maxLength=20 />
        ע����ϵ������</TD>
    </TR>
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name=Submit   type=submit id="Submit" value="�����������"></TD>
    </TR>
  </TABLE>
</form>
<%

End Sub



Sub Modify()
	Dim ProductID
	Dim rsProduct,sqlProduct
	ProductID=Trim(Request("ProductID"))
	If ProductID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		Exit Sub
	Else
		ProductID=Clng(ProductID)
	End If
	Set rsProduct=Server.CreateObject("Adodb.RecordSet")
	sqlProduct="Select * From Product Where ProductID=" & ProductID
	If Not IsObject(conn) Then link_database
	rsProduct.Open sqlProduct,Conn,1,3
	If rsProduct.Bof And rsProduct.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���Ĳ�Ʒ��</li>"
		rsProduct.Close
		Set rsProduct=Nothing
		Exit Sub
	End If
%>
<FORM name="Form1" Action="Administrator_MyShopProduct.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">�޸Ĳ�Ʒ��Ϣ</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��Ʒ����</TD>
      <TD width="60%"><input name="ProductName" type="text" value="<%=rsProduct("ProductName")%>" size=20 maxlength=20 /></TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>���������ͣ�</td>
      <td><Select name="ClassID" >
	  		<%Call Admin_ShowClass_Option_classid(0,rsProduct("ClassID"),"Product_Class")%>

		  </select></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>����С���ͣ�</td>
      <td><Select name="ClassChildID" >
	  		<%Call Admin_ShowClass_Option_classid(0,rsProduct("ClassChildID"),"Product_ClassChild")%>
		  </select></td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ӧ��ID��</td>
      <td><Select name="SupplierID" >
	  		<%Call Admin_Show_Option_keyid(0,rsProduct("SupplierID"),"Product_Suppliers","SupplierName","SupplierID")%>

		  </select></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����ּۣ�</td>
      <td><input name="UnitPrice" type="text" value="<%=FormatCurrency(rsProduct("UnitPrice"))%>" size=6 maxlength=20 style="text-align:center;" />ԪRMB</td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�����г��ۣ�</td>
      <td><input name="UnitPrice_Market" type="text" value="<%=FormatCurrency(rsProduct("UnitPrice_Market"))%>" size=6 maxlength=20 style="text-align:center;" />ԪRMB</td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>����ÿ��λ��Ʒ����������</TD>
      <TD><input name="QuantityPerUnit" type="text" value="<%=rsProduct("QuantityPerUnit")%>" size=2 maxlength=10 /> 
	  
	  ��ѡ��λ��
	  <Select name="QuantityPerUnit_Unit" >
	  	<option value="��" <% If rsProduct("QuantityPerUnit_Unit")="��" Then Response.Write "selected=""selected""" %>>��</option>
		<option value="˫" <% If rsProduct("QuantityPerUnit_Unit")="˫" Then Response.Write "selected=""selected""" %>>˫</option>
		<option value="g" <% If rsProduct("QuantityPerUnit_Unit")="g" Then Response.Write "selected=""selected""" %>>g</option>
		<option value="kg" <% If rsProduct("QuantityPerUnit_Unit")="kg" Then Response.Write "selected=""selected""" %>>kg</option>
		<option value="ml" <% If rsProduct("QuantityPerUnit_Unit")="ml" Then Response.Write "selected=""selected""" %>>ml</option>
		<option value="L" <% If rsProduct("QuantityPerUnit_Unit")="L" Then Response.Write "selected=""selected""" %>>L</option>
	  </select>
	  </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��ǰ�������</TD>
      <TD width="60%"><INPUT name="UnitsInStock" type=text value="<%=rsProduct("UnitsInStock")%>" size=10 maxlength=20 />(+) 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��ǰ��������</TD>
      <TD width="60%"><input name="UnitsOnOrder" type="text" value="<%=rsProduct("UnitsOnOrder")%>" size=10 maxlength=20 />(-) </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">ȡ��������������Ʒ������</TD>
      <TD width="60%"><input name="UnitsOnOrder_del" type="text" value="<%=rsProduct("UnitsOnOrder_del")%>" size=10 maxlength=20 />(+)</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�Ƿ��Ѿ������˲�Ʒ��<br />
	  (<font color="red">�����ϻ���</font>)
	  </TD>
      <TD width="60%"><font color="red"><input type="radio" name="Discontinued" value=1 <% If rsProduct("Discontinued")=1 Then Response.Write "checked=""checked""" %> />
        ��</font> &nbsp;&nbsp; <input type="radio" name="Discontinued" value=0 <% If rsProduct("Discontinued")=0 Then Response.Write "checked=""checked""" %> />
        ��
      </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">��ɫ���ƣ�</TD>
      <TD width="60%"><input name="ColorName" type="text" value="<%=rsProduct("ColorName")%>" size=10 maxlength=20 /></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�ϴ���ɫͼƬ��</TD>
      <TD width="60%"><input name="ColorPhoto" type="text" value="<%=rsProduct("ColorPhoto")%>" size=30 maxlength=50 /></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">���ID���ߴ磩��</TD>
      <TD width="60%"><Select name="StandardID">
						  <%Call Admin_ShowClass_Option_classid(0,rsProduct("StandardID"),"Product_Standard")%>
					  </select></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ļ��������<br />
	  �磺ֱ��_�ۺ�ϸ����
	  </td>
      <td><input name="StandardDesc" type="text" value="<%=rsProduct("StandardDesc")%>" size=10 maxlength=20 /></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ʒ��������</td>
      <td>
	  		<textarea name="Description" cols="38" rows="8"><%=rsProduct("Description")%></textarea>
	  </td>
    </tr>
    <TR class="tdbg" > 
      <TD>�ϴ���Ʒ��ʵ��ͼƬ��</TD>
      <TD><input name="Photo" type="text" value="<%=rsProduct("Photo")%>" size=30 maxlength=50 /></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">����ʱ�䣺</TD>
      <TD width="60%"><INPUT name="ComeInToTheMarketDate" value="<%=rsProduct("ComeInToTheMarketDate")%>" size=7 maxLength=20 />
        ע����ϵ������</TD>
    </TR>
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center">
	   <input name="Action" type="hidden" id="Action" value="SaveModify" />
	   <input name="ProductID" type="hidden" id="ProductID" value="<%=rsProduct("ProductID")%>" /> 
	   
	   <input name=Submit   type=submit id="Submit" value="�����޸Ľ��" /></TD>
    </TR>
  </TABLE>
</form>
<%
	rsProduct.Close
	Set rsProduct=Nothing
End Sub


Sub UpdateUser()
	%>
	<FORM name="Form1" Action="Administrator_MyShopProduct.asp?Action=DoUpdate" method="post">
	  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
		<tr align="center" class="title"> 
		  <td height="22" colspan="2"><strong><font color="#FFFFFF">���²�Ʒ��̬ҳ��</font></strong></td>
	  </tr>
	  <tr class="tdbg"> 
		  <td colspan="2"><p>˵����<br>
			  1�����������������ɲ�Ʒ��̬ҳ�档<br>
			  2�����������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�<br>
		  	  3 �����������ݲ�Ʒ�����¡� </p>
		  </td>
	  </tr>
	  <tr class="tdbg"> 
		<td height="25">��ʼ��ƷID��</td>
		<td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
		  ��ƷID��������д�������һ��ID�ſ�ʼ���и���</td>
	  </tr>
	  <tr class="tdbg"> 
		<td height="25">������ƷID��</td>
		<td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
		  �����¿�ʼ������ID֮��Ĳ�Ʒ���ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
	  </tr>
	  <tr class="tdbg"> 
		<td height="25">&nbsp;</td>
		<td height="25"><input name="Submit" type="submit" id="Submit" value="���ɾ�̬ҳ��"></td>
	  </tr>
	</table>
	</form>
	
	<FORM name="Form1" Action="Administrator_MyShopProduct.asp?Action=DoUpdatelog" method="post">
	  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
		<tr align="center" class="title"> 
		  <td height="22" colspan="2"><strong><font color="#FFFFFF">�������¾�̬ҳ��</font></strong></td>
	  </tr>
	  <tr class="tdbg"> 
		  <td colspan="2"><p>˵����<br>
			  1�����������������ɲ�Ʒ��̬ҳ�档<br>
			  2�����������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�<br>
		  3���������������£����¡�</p>
		  </td>
	  </tr>
	  <tr class="tdbg"> 
		<td height="25">��ʼ����ID��</td>
		<td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
		  ��ƷID��������д�������һ��ID�ſ�ʼ���и���</td>
	  </tr>
	  <tr class="tdbg"> 
		<td height="25">��������ID��</td>
		<td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
		  �����¿�ʼ������ID֮�������ҳ�棬֮�����ֵ��ò�Ҫѡ�����</td>
	  </tr>
	  <tr class="tdbg"> 
		<td height="25">&nbsp;</td>
		<td height="25"><input name="Submit" type="submit" id="Submit" value="�������¾�̬ҳ��"></td>
	  </tr>
	</table>
	</form>
  <%
End Sub

Sub gouser1()
	If is_ot_user=1 Then
		Response.Write("�����ⲿ���ݿ��Ʒ������ʹ�ô˹���"):Response.End()
	End If
	%>
	<FORM name="Form1" Action="Administrator_MyShopProduct.asp?Action=gouser2" method="post" target="_blank">
	  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
		<tr align="center" class="title"> 
		  <td height="22" colspan="2"><strong><font color="#FFFFFF">��¼����Ŀ�����̨</font></strong></td>
	  </tr>
	  <tr class="tdbg"> 
		  <td colspan="2"><p>˵����<br>
			  ������������Ա��¼����Ŀ�Ĺ��������й���<br>
			  ������Ŀ���������ϰ�ʱ���ɽ���ò�Ʒ��̨��<br>
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
	Dim rsProduct,sqlProduct
	Dim ProductName,ClassID,ClassChildID,SupplierID,UnitPrice,UnitPrice_Market,QuantityPerUnit,QuantityPerUnit_Unit,UnitsInStock,UnitsOnOrder,UnitsOnOrder_del,Discontinued,ColorName,ColorPhoto,StandardID,StandardDesc,Description,Photo,ComeInToTheMarketDate
	
	ProductName		=Trim(Request("ProductName"))
	ClassID			=Trim(Request("ClassID"))
	ClassChildID	=Trim(Request("ClassChildID"))
	SupplierID		=Trim(Request("SupplierID"))
	UnitPrice		=Trim(Request("UnitPrice"))
	UnitPrice_Market=Trim(Request("UnitPrice_Market"))
	QuantityPerUnit	=Trim(Request("QuantityPerUnit"))
	QuantityPerUnit_Unit	=Trim(Request("QuantityPerUnit_Unit"))
	UnitsInStock	=Trim(Request("UnitsInStock"))
	UnitsOnOrder	=Trim(Request("UnitsOnOrder"))
	UnitsOnOrder_del=Trim(Request("UnitsOnOrder_del"))
	Discontinued	=Trim(Request("Discontinued"))
	ColorName		=Trim(Request("ColorName"))
	ColorPhoto		=Trim(Request("ColorPhoto"))
	StandardID		=Trim(Request("StandardID"))
	StandardDesc	=Trim(Request("StandardDesc"))
	Description		=Trim(Request("Description"))
	Photo			=Trim(Request("Photo"))
	ComeInToTheMarketDate	=Trim(Request("ComeInToTheMarketDate"))
	

	If ProductName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ���Ʋ���Ϊ��</li>"
	Else
		ProductName = ProductName
	End If

	If ClassID<>"" Then
		If Not isNumeric(ClassID) Or Cstr(ClassID)="0" Then	' Or Len(Cstr(OICQ))>10
			Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĵ���࣡����������û����Ҫ�����ȥ��������������</li>"
			founderr=True
		Else
			ClassID = Cint(ClassID)
		End If
	Else
		Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĵ���࣡</li>"
		founderr=True
	End If
	
	If ClassChildID<>"" Then
		If Not isNumeric(ClassChildID) Or Cstr(ClassChildID)="0" Then	' Or Len(Cstr(OICQ))>10
			Errmsg=Errmsg & "<br><li>��ѡ���Ʒ��С���࣡���С������û����Ҫ�����ȥС�������������</li>"
			founderr=True
		Else
			ClassChildID = Cint(ClassChildID)
		End If
	Else
		Errmsg=Errmsg & "<br><li>��ѡ���Ʒ��С���࣡</li>"
		founderr=True
	End If

	If SupplierID<>"" Then
		If Not isNumeric(SupplierID) Or Cstr(SupplierID)="0" Then	' Or Len(Cstr(OICQ))>10
			Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĺ�Ӧ�̣������Ӧ����û����Ҫ�����ȥ��Ӧ�̹���������</li>"
			founderr=True
		Else
			ClassID = Cint(ClassID)
		End If
	Else
		Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĺ�Ӧ�̣�</li>"
		founderr=True
	End If
	
	If UnitPrice="" Or Not isNumeric(UnitPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�ּ۲���Ϊ�գ�����ҪΪ���֣�</li>"
	Else
		UnitPrice = FormatCurrency(UnitPrice)
	End If
	
	If UnitPrice_Market="" Or Not isNumeric(UnitPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�г��۲���Ϊ�գ�����ҪΪ���֣�</li>"
	Else
		UnitPrice_Market = FormatCurrency(UnitPrice_Market)
	End If
	
	If QuantityPerUnit<>"" Then
		If Not isNumeric(QuantityPerUnit) Or Cstr(QuantityPerUnit)="0" Then
			Errmsg=Errmsg & "<br><li>����д ÿ��λ��Ʒ����������</li>"
			founderr=True
		Else
			QuantityPerUnit = Cint(QuantityPerUnit)
		End If
	Else
		Errmsg=Errmsg & "<br><li>����д ÿ��λ��Ʒ����������</li>"
		founderr=True
	End If
	
	If QuantityPerUnit_Unit="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>ÿ��λ��Ʒ��������e��λ ����Ϊ��</li>"
	Else
		QuantityPerUnit_Unit = Trim(QuantityPerUnit_Unit)
	End If
	
	If UnitsInStock="" Or isNumeric(UnitsInStock) Then
		UnitsInStock = UnitsInStock
	Else
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�Ŀ����ҪΪ���֣�</li>"
	End If
	
	If UnitsOnOrder="" Or isNumeric(UnitsOnOrder) Then
		UnitsOnOrder = UnitsOnOrder
	Else
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�ĵ�ǰ������ҪΪ���֣�</li>"
	End If
	
	If UnitsOnOrder_del="" Or isNumeric(UnitsOnOrder_del) Then
		UnitsOnOrder_del = UnitsOnOrder_del
	Else
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ��ȡ��������������Ʒ�� ������ҪΪ���֣�</li>"
	End If
	
	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��ѡ�� �Ƿ��Ѿ������˲�Ʒ��ѡ��״̬��</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	
	'ColorName
	'ColorPhoto
	'StandardID
	'StandardDesc
	'Description
	'Photo
	'ComeInToTheMarketDate
	
	

	If founderr=True Then
		Set rs=Nothing
		Exit Sub
	End If
	
	Set rsProduct=Server.CreateObject("Adodb.RecordSet")
	sqlProduct="Select * From Product"
	If Not IsObject(conn) Then link_database
	rsProduct.Open sqlProduct,Conn,1,3

	rsProduct.Addnew
	
		If Not ProductName="" Then  			rsProduct("ProductName")		=ProductName
		If Not ClassID="" Then  				rsProduct("ClassID")			=ClassID
		If Not ClassChildID="" Then  			rsProduct("ClassChildID")		=ClassChildID
		If Not SupplierID="" Then  				rsProduct("SupplierID")			=SupplierID
		If Not UnitPrice="" Then  				rsProduct("UnitPrice")			=UnitPrice
		If Not UnitPrice_Market="" Then  		rsProduct("UnitPrice_Market")	=UnitPrice_Market
		If Not QuantityPerUnit="" Then  		rsProduct("QuantityPerUnit")	=QuantityPerUnit
		If Not QuantityPerUnit_Unit="" Then  	rsProduct("QuantityPerUnit_Unit")	=QuantityPerUnit_Unit
		If Not UnitsInStock="" Then  			rsProduct("UnitsInStock")		=UnitsInStock
		If Not UnitsOnOrder="" Then  			rsProduct("UnitsOnOrder")		=UnitsOnOrder
		If Not UnitsOnOrder_del="" Then  		rsProduct("UnitsOnOrder_del")	=UnitsOnOrder_del
		If Not Discontinued="" Then  			rsProduct("Discontinued")		=Discontinued
		If Not ColorName="" Then 			    rsProduct("ColorName")			=ColorName
		If Not ColorPhoto="" Then  				rsProduct("ColorPhoto")			=ColorPhoto
		If Not StandardID="" Then  				rsProduct("StandardID")			=StandardID
		If Not StandardDesc="" Then  			rsProduct("StandardDesc")		=StandardDesc
		If Not Description="" Then  			rsProduct("Description")		=Description
		If Not Photo="" Then  					rsProduct("Photo")				=Photo
		If Not ComeInToTheMarketDate="" Then  	rsProduct("ComeInToTheMarketDate")	=ComeInToTheMarketDate
		
	rsProduct.Update
	
	rsProduct.Close
	Set rsProduct=Nothing
	
	oblog.showok "ע���²�Ʒ�ɹ�!",""
End Sub



Sub gouser2()
	If is_ot_user=1 Then
		Response.Write("�����ⲿ���ݿ��Ʒ������ʹ�ô˹���"):Response.End()
	End If
	Dim rs,username
	username=oblog.filt_badstr(Trim(Request("username")))
	If username="" Then Response.Write("��Ʒ������Ϊ��"):Response.End()
	If is_ot_user=0 Then
		Set rs=oblog.Execute("Select username,password From Product Where username='"&username&"'")
	Else
		Set rs=ot_conn.Execute("Select "&ot_username&","&ot_password&" From "&ot_usertable&" Where "&ot_username&"='"&username&"'")
	End If
	If Not rs.Eof Then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		Set rs=Nothing
		Response.Redirect("../HomestayBackctrl-indexHome.asp")
	Else
		Set rs=Nothing
		Response.Write("�޴˲�Ʒ"):Response.End()
	End If
	
End Sub


Sub SaveModify()
	Dim rsProduct,sqlProduct
	Dim ProductName,ClassID,ClassChildID,SupplierID,UnitPrice,UnitPrice_Market,QuantityPerUnit,QuantityPerUnit_Unit,UnitsInStock,UnitsOnOrder,UnitsOnOrder_del,Discontinued,ColorName,ColorPhoto,StandardID,StandardDesc,Description,Photo,ComeInToTheMarketDate
	
	ProductID=Trim(Request("ProductID"))
	If ProductID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		Exit Sub
	Else
		ProductID = Clng(ProductID)
	End If
	
	ProductName		=Trim(Request("ProductName"))
	ClassID			=Trim(Request("ClassID"))
	ClassChildID	=Trim(Request("ClassChildID"))
	SupplierID		=Trim(Request("SupplierID"))
	UnitPrice		=Trim(Request("UnitPrice"))
	UnitPrice_Market=Trim(Request("UnitPrice_Market"))
	QuantityPerUnit	=Trim(Request("QuantityPerUnit"))
	QuantityPerUnit_Unit	=Trim(Request("QuantityPerUnit_Unit"))
	UnitsInStock	=Trim(Request("UnitsInStock"))
	UnitsOnOrder	=Trim(Request("UnitsOnOrder"))
	UnitsOnOrder_del=Trim(Request("UnitsOnOrder_del"))
	Discontinued	=Trim(Request("Discontinued"))
	ColorName		=Trim(Request("ColorName"))
	ColorPhoto		=Trim(Request("ColorPhoto"))
	StandardID		=Trim(Request("StandardID"))
	StandardDesc	=Trim(Request("StandardDesc"))
	Description		=Trim(Request("Description"))
	Photo			=Trim(Request("Photo"))
	ComeInToTheMarketDate	=Trim(Request("ComeInToTheMarketDate"))
	

	If ProductName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ���Ʋ���Ϊ��</li>"
	Else
		ProductName = ProductName
	End If

	If ClassID<>"" Then
		If Not isNumeric(ClassID) Or Cstr(ClassID)="0" Then	' Or Len(Cstr(OICQ))>10
			Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĵ���࣡����������û����Ҫ�����ȥ��������������</li>"
			founderr=True
		Else
			ClassID = Cint(ClassID)
		End If
	Else
		Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĵ���࣡</li>"
		founderr=True
	End If
	
	If ClassChildID<>"" Then
		If Not isNumeric(ClassChildID) Or Cstr(ClassChildID)="0" Then	' Or Len(Cstr(OICQ))>10
			Errmsg=Errmsg & "<br><li>��ѡ���Ʒ��С���࣡���С������û����Ҫ�����ȥС�������������</li>"
			founderr=True
		Else
			ClassChildID = Cint(ClassChildID)
		End If
	Else
		Errmsg=Errmsg & "<br><li>��ѡ���Ʒ��С���࣡</li>"
		founderr=True
	End If

	If SupplierID<>"" Then
		If Not isNumeric(SupplierID) Or Cstr(SupplierID)="0" Then	' Or Len(Cstr(OICQ))>10
			Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĺ�Ӧ�̣������Ӧ����û����Ҫ�����ȥ��Ӧ�̹���������</li>"
			founderr=True
		Else
			ClassID = Cint(ClassID)
		End If
	Else
		Errmsg=Errmsg & "<br><li>��ѡ���Ʒ�Ĺ�Ӧ�̣�</li>"
		founderr=True
	End If
	
	If UnitPrice="" Or Not isNumeric(UnitPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�ּ۲���Ϊ�գ�����ҪΪ���֣�</li>"
	Else
		UnitPrice = FormatCurrency(UnitPrice)
	End If
	
	If UnitPrice_Market="" Or Not isNumeric(UnitPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�г��۲���Ϊ�գ�����ҪΪ���֣�</li>"
	Else
		UnitPrice_Market = FormatCurrency(UnitPrice_Market)
	End If
	
	If QuantityPerUnit<>"" Then
		If Not isNumeric(QuantityPerUnit) Or Cstr(QuantityPerUnit)="0" Then
			Errmsg=Errmsg & "<br><li>����д ÿ��λ��Ʒ����������</li>"
			founderr=True
		Else
			QuantityPerUnit = Cint(QuantityPerUnit)
		End If
	Else
		Errmsg=Errmsg & "<br><li>����д ÿ��λ��Ʒ����������</li>"
		founderr=True
	End If
	
	If QuantityPerUnit_Unit="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>ÿ��λ��Ʒ��������e��λ ����Ϊ��</li>"
	Else
		QuantityPerUnit_Unit = Trim(QuantityPerUnit_Unit)
	End If
	
	If UnitsInStock="" Or isNumeric(UnitsInStock) Then
		UnitsInStock = UnitsInStock
	Else
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�Ŀ����ҪΪ���֣�</li>"
	End If
	
	If UnitsOnOrder="" Or isNumeric(UnitsOnOrder) Then
		UnitsOnOrder = UnitsOnOrder
	Else
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ�ĵ�ǰ������ҪΪ���֣�</li>"
	End If
	
	If UnitsOnOrder_del="" Or isNumeric(UnitsOnOrder_del) Then
		UnitsOnOrder_del = UnitsOnOrder_del
	Else
		founderr=True
		Errmsg=Errmsg & "<br><li>��Ʒ��ȡ��������������Ʒ�� ������ҪΪ���֣�</li>"
	End If
	
	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>��ѡ�� �Ƿ��Ѿ������˲�Ʒ��ѡ��״̬��</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	
	'ColorName
	'ColorPhoto
	'StandardID
	'StandardDesc
	'Description
	'Photo
	'ComeInToTheMarketDate
	

	If founderr=True Then
		'Set rs=Nothing
		Exit Sub
	End If
	
	Set rsProduct=Server.CreateObject("Adodb.RecordSet")
	sqlProduct="Select * From Product Where ProductID=" & ProductID
	If Not IsObject(conn) Then link_database
	rsProduct.Open sqlProduct,Conn,1,3
	If rsProduct.Bof And rsProduct.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���Ĳ�Ʒ��</li>"
		rsProduct.Close
		Set rsProduct=Nothing
		Exit Sub
	End If
	
		If Not ProductName="" Then  			rsProduct("ProductName")		=ProductName
		If Not ClassID="" Then  				rsProduct("ClassID")			=ClassID
		If Not ClassChildID="" Then  			rsProduct("ClassChildID")		=ClassChildID
		If Not SupplierID="" Then  				rsProduct("SupplierID")			=SupplierID
		If Not UnitPrice="" Then  				rsProduct("UnitPrice")			=UnitPrice
		If Not UnitPrice_Market="" Then  		rsProduct("UnitPrice_Market")	=UnitPrice_Market
		If Not QuantityPerUnit="" Then  		rsProduct("QuantityPerUnit")	=QuantityPerUnit
		If Not QuantityPerUnit_Unit="" Then  	rsProduct("QuantityPerUnit_Unit")	=QuantityPerUnit_Unit
		If Not UnitsInStock="" Then  			rsProduct("UnitsInStock")		=UnitsInStock
		If Not UnitsOnOrder="" Then  			rsProduct("UnitsOnOrder")		=UnitsOnOrder
		If Not UnitsOnOrder_del="" Then  		rsProduct("UnitsOnOrder_del")	=UnitsOnOrder_del
		If Not Discontinued="" Then  			rsProduct("Discontinued")		=Discontinued
		If Not ColorName="" Then 			    rsProduct("ColorName")			=ColorName
		If Not ColorPhoto="" Then  				rsProduct("ColorPhoto")			=ColorPhoto
		If Not StandardID="" Then  				rsProduct("StandardID")			=StandardID
		If Not StandardDesc="" Then  			rsProduct("StandardDesc")		=StandardDesc
		If Not Description="" Then  			rsProduct("Description")		=Description
		If Not Photo="" Then  					rsProduct("Photo")				=Photo
		If Not ComeInToTheMarketDate="" Then  	rsProduct("ComeInToTheMarketDate")	=ComeInToTheMarketDate
		
	rsProduct.Update
	rsProduct.Close
	Set rsProduct=Nothing
	
	oblog.showok "��Ʒ�޸ĳɹ�!",""
End Sub


Sub DelProduct()
	Dim rs,i
	If ProductID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ���Ĳ�Ʒ</li>"
		Exit Sub
	End If
	If Instr(ProductID,",")>0 Then
		ProductID=Split(ProductID,",")
		For i=0 to Ubound(ProductID)
			DelOneProduct(ProductID(i))
		Next
	Else
		DelOneProduct(ProductID)
	End If
	Response.Redirect "Administrator_MyShopProduct.asp"
End Sub

Sub DelOneProduct(ProductID)
	ProductID=Clng(ProductID)
'	Dim rs,fso,f,uname,udir
'	Set rs=oblog.Execute("Select * From Product Where ProductID="&ProductID)
	If Not rs.Eof Then

		oblog.Execute("Delete From Product Where ProductID=" & ProductID)
'		oblog.Execute("Update Product Set Discontinued=1 Where ProductID="& ProductID)
	End If
'	Set rs=Nothing
End Sub





Sub MoveUser()
	Dim msg
	If ProductID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�ƶ��Ĳ�Ʒ</li>"
		Exit Sub
	End If
	Dim User_Level
	User_Level=Trim(Request("User_Level"))
	If User_Level="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ���Ʒ��</li>"
		Exit Sub
	Else
		User_Level=Clng(User_Level)
	End If
	If Instr(ProductID,",")>0 Then
		ProductID=replace(ProductID," ","")
		If User_Level=6 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>ע���Ʒ</font>����"
			sql="Update Product Set User_Level=6 Where ProductID in (" & ProductID & ")"
		ElseIf User_Level=7 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update Product Set User_Level=7 Where ProductID in (" & ProductID & ")"
		ElseIf User_Level=8 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update Product Set User_Level=8 Where ProductID in (" & ProductID & ")"
		ElseIf User_Level=9 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>ǰ̨������</font>����"
			sql="Update Product Set User_Level=9 Where ProductID in (" & ProductID & ")"
		End If
	Else
		If User_Level=6 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>ע���Ʒ</font>����"
			sql="Update Product Set User_Level=6 Where ProductID=" & ProductID 
		ElseIf User_Level=7 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update Product Set User_Level=7 Where ProductID="& ProductID
		ElseIf User_Level=8 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update Product Set User_Level=8 Where ProductID=" & ProductID 
		ElseIf User_Level=9 Then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ʒ��Ϊ��<font color=blue>ǰ̨������</font>����"
			sql="Update Product Set User_Level=9 Where ProductID=" & ProductID
		End If
	End If
	Conn.Execute sql
	Response.Redirect "Administrator_MyShopProduct.asp"
	'Call WriteSuccessMsg(msg)
End Sub

Sub DoUpdate()
	Server.ScriptTimeOut=999999999
	Dim BeginID,EndID,p1,rsProduct,blog,i
	BeginID=Trim(Request("BeginID"))
	EndID=Trim(Request("EndID"))
	If BeginID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ����ʼID</li>"
	Else
		BeginID=Clng(BeginID)
	End If
	If EndID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ������ID</li>"
	Else
		EndID=Clng(EndID)
	End If
	If FoundErr=True Then Exit Sub
	Set rsuser=oblog.Execute("Select count(ProductID) From Product Where ProductID>=" & Clng(BeginID) & " And ProductID<=" & Clng(EndID))
	p1=rsuser(0)
	Set rsuser=oblog.Execute("Select ProductID From Product Where ProductID>=" & Clng(BeginID) & " And ProductID<=" & Clng(EndID)&" Order By ProductID")
	Set blog=new class_blog
	Response.Write("<div style=""text-align: center;"">")
	Response.Write("<div class=""progress1""><div class=""progress2"" id=""progress1""></div></div><span id=""pstr1""></span><br><br>")
	i=1
	blog.progress_init
	Do While Not rsProduct.Eof
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""ȫ�����ȣ���ǰ��ƷID:"&rsuser(0)&""";</script>" & VbCrLf
		Response.Flush
		'blog.ProductID=rsProduct(0)
		blog.update_alllog_admin rsProduct(0)
		If P_BLOG_UPDATEPAUSE>0 Then Call Pause()
		rsProduct.MoveNext
		i=i+1
	Loop
	Response.Write("</div>")
	Set rsProduct=Nothing
	Set blog=Nothing
End Sub

Sub DoUpdatelog()
	Server.ScriptTimeOut=999999999
	Dim BeginID,EndID,p1,rs,blog,i
	BeginID=Trim(Request("BeginID"))
	EndID=Trim(Request("EndID"))
	If BeginID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ����ʼID</li>"
	Else
		BeginID=Clng(BeginID)
	End If
	If EndID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ������ID</li>"
	Else
		EndID=Clng(EndID)
	End If
	If FoundErr=True Then Exit Sub
	Set rs=oblog.Execute("Select count(logid) From oblog_log Where logid>=" & Clng(BeginID) & " And logid<=" & Clng(EndID))
	p1=rs(0)
	Set rs=oblog.Execute("Select logid,ProductID From oblog_log Where logid>=" & Clng(BeginID) & " And logid<=" & Clng(EndID)&" Order By logid")
	Set blog=new class_blog
	Response.Write("<div style=""text-align: center;"">")
	Response.Write("<div class=""progress1""><div class=""progress2"" id=""progress1""></div></div><span id=""pstr1""></span><br><br>")
	i=1
	'blog.progress_init
	Do While Not rs.Eof
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""���ȣ���ǰ����ID:"&rs(0)&""";</script>" & VbCrLf
		Response.Flush
		blog.ProductID=rs(1)
		blog.update_log rs(0),0
		If P_BLOG_UPDATEPAUSE>0 Then Call Pause()
		rs.MoveNext
		i=i+1
	Loop
	Response.Write("</div>")
	Set rs=Nothing
	Set blog=Nothing
End Sub




Sub ShowInfo()
	Dim ProductID
	Dim rsProduct,sqlProduct
	ProductID=Trim(Request("ProductID"))
	If ProductID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		Exit Sub
	Else
		ProductID=Clng(ProductID)
	End If
	Set rsProduct=Server.CreateObject("Adodb.RecordSet")
	sqlProduct="Select * From Product Where ProductID=" & ProductID
	If Not IsObject(conn) Then link_database
	rsProduct.Open sqlProduct,Conn,1,3
	If rsProduct.Bof And rsProduct.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���Ĳ�Ʒ��</li>"
		rsProduct.Close
		Set rsProduct=Nothing
		Exit Sub
	End If
	Dim strGuide
	strGuide=strGuide&"<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Administrator_MyShopProduct.asp'>ע��ע���Ʒ����</a>&nbsp;&gt;&gt;&nbsp;" & "�й���ͥ<font color=red>"& rsuser("blogname") &"</font>����ϸ��Ϣ</td></tr></table>"
	
	Response.Write(strGuide)
%>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">(<%=rsuser("ProductID")%>)ע���Ʒ��ϸ��Ϣ</font></b>&nbsp;&nbsp;</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">��ͥ��¼�ʺţ�</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsProduct("ProductName")%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">Email��ַ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsProduct("userEmail")%> </font>&nbsp;&nbsp;
        <a href="mailto:<%=rsProduct("userEmail")%>">���˲�Ʒ��һ������ʼ�</a> 
      </TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">��ϵ��ʽ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsProduct("tel")%> &nbsp;&nbsp;&nbsp;<%=rsProduct("mobile")%></font>
      </TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">ͨѶ��ַ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsProduct("address")%> &nbsp;&nbsp;&nbsp;�ʱࣺ<%=rsProduct("postnum")%></font>
      </TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">���֪�����ǣ�</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsProduct("howknowsite")%></font> 
      </TD>
    </TR>
	
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>��Ʒ������</td>
      <td><input name="user_domain" type="text" value="<%=oblog.filt_html(rsuser("user_domain"))%>" size=10 maxlength=20 /> <Select name="user_domainroot" ><option value="123">666</option></select></td>
    </tr>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>ѡ��Ӵ���ͥ���ࣺ</td>
      <td><%
			Dim FamilyTypeRequest,FamilyType
				FamilyType = rsuser("FamilyType")
				
				Select Case FamilyType
					Case "1"
					FamilyTypeRequest= "��ѽӴ���ͥ"
					Case "2"
					FamilyTypeRequest= "�շѽӴ���ͥ"
					Case "3"
					FamilyTypeRequest= "2008���˽Ӵ���ͥ"
					Case "1, 2"
					FamilyTypeRequest= "���/�շѽӴ���ͥ"
					Case "1, 3"
					FamilyTypeRequest= "��ѽӴ���ͥ/2008���˽Ӵ���ͥ"
					Case "1, 2, 3"
					FamilyTypeRequest= "���/�շ�/2008���˽Ӵ���ͥ"
					Case "2, 3"
					FamilyTypeRequest= "�շѽӴ���ͥ/2008���˽Ӵ���ͥ"
				End Select%>
	  <%=FamilyTypeRequest%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td><strong>��ʵ������</strong></td>
      <td><strong><%=rsuser("blogname")%></strong></td>
    </tr>
	<TR class="tdbg" > 
      <TD width="40%">��ע��</TD>
      <TD width="60%"><%=rsProduct("beizhu")%>
        </TD>
    </TR>
	
	
	<%If true_domain=1 Then%>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ʒ�󶨵Ķ���������</td>
      <td><input name=custom_domain   type=text id="custom_domain" value="<%=rsuser("custom_domain")%>" size=30 maxlength=20></td>
    </tr>
	<%End If%>
	
	<TR class="tdbg" > 
      <TD width="40%">���ڵ���(ʡ/��)��</TD>
      <TD width="60%"><%=rsProduct("province")%>&nbsp;<%=rsProduct("city")%> 
      </TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>ѡ��ѧϰ��Χ��</td>
      <td><Select name="usertype" id="usertype">
          <%If rsProduct("user_classid")<>"" Then
	  Response.Write(oblog.show_class("user",rsProduct("user_classid"),0))
	  Else
	  Response.Write(oblog.show_class("user",0,0))
	  End If
	  %>
        </select></td>
    </tr>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">����(����6λ)��<BR>
        ���������룬���ִ�Сд�� �벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ� </TD>
      <TD width="60%"> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">��������޸ģ�������(���ϲ�Ʒ�뵽��̳�޸�)</font> </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD>ȷ������(����6λ)��<br>
        ������һ��ȷ��</TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=12> <font color="#FF0000">��������޸ģ�������(���ϲ�Ʒ�뵽��̳�޸�)</font> </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">�������⣺<br>
        �����������ʾ����</TD>
      <TD width="60%"> <INPUT name="Question"   type=text value="<%=rsProduct("Question")%>" size=30>(���ϲ�Ʒ�뵽��̳�޸�) 
      </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">����𰸣�<BR>
        �����������ʾ����𰸣�����ȡ������</TD>
      <TD width="60%"> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">��������޸ģ�������(���ϲ�Ʒ�뵽��̳�޸�)</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�Ա�</TD>
      <TD width="60%"><%If rsProduct("Sex")=1 Then Response.Write "��"%>
        <%If rsProduct("Sex")=0 Then Response.Write "Ů"%>
        </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�������ڣ�</TD>
      <TD width="60%"><%=rsProduct("birthday")%>
        </TD>
    </TR>
	
    
	
	
	<TR class="tdbg" > 
      <TD width="40%">����ְҵ��</TD>
      <TD width="60%"><%=rsProduct("job")%> 
      </TD>
    </TR>
	
	
	
	<TR class="tdbg" > 
      <TD width="40%">��ͥ������</TD>
      <TD width="60%"><%=rsProduct("familynumber")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">��ͥ�ṹ��</TD>
      <TD width="60%"><%=rsProduct("familymember")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">���ݽ��ܣ�</TD>
      <TD width="60%">�ṹ��<%=rsProduct("houseframe")%> &nbsp;&nbsp;&nbsp;�����<%=rsProduct("housespace")%>
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�Ƿ��г��</TD>
      <TD width="60%"><%=rsProduct("pet")%>&nbsp;&nbsp;&nbsp;<%=rsProduct("PetType")%>
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">����Ա�</TD>
      <TD width="60%"><%=rsProduct("asksex")%>
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�Ƿ��˵���ģ�</TD>
      <TD width="60%"><%=rsProduct("issaychinese")%> 
      </TD>
    </TR>
	
	
	
	<TR class="tdbg" > 
      <TD width="40%">Ӣ��ˮƽ��</TD>
      <TD width="60%"><%=rsProduct("englishlevel")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�Ƿ���˽�ҳ���</TD>
      <TD width="60%"><%=rsProduct("car")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" style="display:none;" > 
      <TD width="40%">װ�ޱ�׼��</TD>
      <TD width="60%"><%=rsProduct("fitment")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">����������ͣ�</TD>
      <TD width="60%"><%=rsProduct("toilet")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�з���ԣ�</TD>
      <TD width="60%"><%=rsProduct("computer")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�з���������</TD>
      <TD width="60%"><%=rsProduct("internet")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�ͷ�����(��ע)��</TD>
      <TD width="60%"><%=rsProduct("remarkinfo")%> 
      </TD>
    </TR>
	
	
	
	
    
	
	
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">��ϸ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsProduct("xuwei_beizhu")%></font>
        </TD>
    </TR>
	
	
	
	
	<TR style="display:none;" class="tdbg" > 
      <TD width="40%">OICQ���룺</TD>
      <TD width="60%"> <INPUT name=OICQ value="<%=rsProduct("qq")%>" size=30 maxLength=20></TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">MSN��</TD>
      <TD width="60%"> <INPUT name=msn value="<%=rsProduct("Msn")%>" size=30 maxLength=50></TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">��Ʒ����</TD>
      <TD width="60%"><Select name="User_Level" id="User_Level">
          <option value="6" <%If Clng(rsProduct("user_level"))=6 Then Response.Write("selected")%>>ע���Ʒ</option>
          <option value="7" <%If Clng(rsProduct("user_level"))=7 Then Response.Write("selected")%>>ע����Ŀ����</option>
          <option value="8" <%If Clng(rsProduct("user_level"))=8 Then Response.Write("selected")%>>vip��Ŀ����</option>
          <option value="9" <%If Clng(rsProduct("user_level"))=9 Then Response.Write("selected")%>>ǰ̨������</option>
        </select></TD>
    </TR>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ռ�(kb)��</td>
      <td><input name=user_upfiles_max   type=text value="<%=rsuser("user_upfiles_max")%>" size=20 maxlength=20>
        Ϊ��ʱΪϵͳĬ�����ã����赥�����ã�������kbֵ</td>
    </tr>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ֽ�(�ֽ�)��</td>
      <td><input name=upfiles_size   type=text id="upfiles_size" value="<%=rsuser("user_upfiles_size")%>" size=20 maxlength=20></td>
    </tr>
    <TR class="tdbg" > 
      <TD>�Ƿ�Ϊ�Ƽ���Ʒ��</TD>
      <TD><%If rsProduct("user_isbest")=1 Then Response.Write "��"%>
        <%If rsProduct("user_isbest")<>1 Then Response.Write "��"%>
        </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">��ƷĿ¼��</TD>
      <TD width="60%"> <INPUT name=user_dir value="<%=rsProduct("user_dir")%>" size=30 maxLength=50>
        ���ޱ�Ҫ�벻Ҫ�޸ģ�������ɲ�ƷĿ¼����</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��Ʒ״̬��</TD>
      <TD width="60%"><%If rsProduct("Discontinued")=0 Then Response.Write "����"%>
        <%If rsProduct("Discontinued")=1 Then Response.Write "����"%>
        </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input type="button" onClick="history.go(-1);" value=" ������һҳ " style="cursor:hand;"></TD>
    </TR>
  </TABLE>

<%
	rsProduct.Close
	Set rsProduct=Nothing
End Sub


Sub Discontinued()
	ProductID=Clng(ProductID)
	oblog.Execute("Update Product Set Discontinued=1 Where ProductID="& ProductID)
	Response.Redirect "Administrator_MyShopProduct.asp"
End Sub


Sub UnDiscontinued()
	ProductID=Clng(ProductID)
	oblog.Execute("Update Product Set Discontinued=0 Where ProductID="& ProductID)
	Response.Redirect "Administrator_MyShopProduct.asp"
End Sub





'ʹ�����ӣ�Call Admin_ShowClass_Option_classid(0,ParentID,"Product_Class")		'ParentID=classid.
Sub Admin_ShowClass_Option_classid(ShowType,CurrentID,Table_Class)
Dim t
t = 0
	
	If ShowType=0 Then
	    Response.Write "<option value='0'"
		If CurrentID=0 Then Response.Write " selected"
		Response.Write ">�ޣ���ѡ����ࣩ</option>"
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
		Response.Write "<option value=''>���������̨���ѡ��</option>"
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



'ʹ�����ӣ�Call Admin_Show_Option_keyid(0,ParentID,"Product_Suppliers","SupplierName","SupplierID")		'ParentID=classid.
Sub Admin_Show_Option_keyid(ShowType,CurrentID,Table_Name,Field_Name,ID_Name)
	If ShowType=0 Then
	    Response.Write "<option value='0'"
		If CurrentID=0 Then Response.Write " selected"
		Response.Write ">�ޣ���ѡ����ࣩ</option>"
	End If
	Dim rsClass,sqlClass,strTemp

	sqlClass="Select * From "& Table_Name &" "
						't="0"   tName="��Ʒ     /���
	Set rsClass=conn.Execute(sqlClass)
	If rsClass.Bof And rsClass.Eof Then 
		Response.Write "<option value=''>���������̨���ѡ��</option>"
	Else
		Do While Not rsClass.Eof
			
			strTemp="<option value='" & rsClass( ID_Name ) & "'"		'WL
			If CurrentID>0 And rsClass( ID_Name )=CurrentID Then		'WL
				 strTemp=strTemp & " selected"
			End If
			strTemp=strTemp & ">"
			

			strTemp=strTemp & rsClass( Field_Name )
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

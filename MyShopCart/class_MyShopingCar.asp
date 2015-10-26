<%
Class class_ShopingCart
'Include
'-----------------------------------
	Public Sub CreateCart()
	  If IsObject(Session("ShopingCart")) = False Then
		 Set Session("ShopingCart") = CreateObject("Scripting.Dictionary")
	  End If
	End Sub
	
	Public Function CheckCart()
	  If IsObject(Session("ShopingCart")) Then
		 CheckCart = True
	  Else
		 CheckCart = False
	  End If
	End Function
	
	Public Function CheckItem(ID)
	  Dim aryCart,findFlag,i
	  If CheckCart = True Then
		  '//aryCart = Session("ShopingCart")
		  findFlag = False
		  If Session("ShopingCart").Exists(ID) = True Then
			 findFlag = True
		  End If
		  CheckItem = findFlag
	  End If
	End Function
	
	Public Sub RemoveItem(ID)
	  ID = Clng(ID)
	  If Not CheckItem(ID) Then
		 Exit Sub
	  End If
	  '//aryCart = Session("ShopingCart")
	  If Session("ShopingCart").Exists(ID) = True Then
		 Session("ShopingCart").Remove(ID)
	  End If
	End Sub
	
	Public Sub UpdateItem(ID,Num)
	  ID	= Clng(ID)
	  Num	= Clng(Num)
	  '//Dim aryUpdateCart
	  '//aryUpdateCart = Session("ShopingCart")
	  If Session("ShopingCart").Exists(ID) = True Then
		 Session("ShopingCart")(ID) = Num
	  End If
	End Sub
	
	Public Sub AddItem(ID,Num)
	  ID	= Clng(ID)
	  Num	= Clng(Num)  
	  Dim btnCartStatus,aryAddCart,newSize,i
	  btnCartStatus = CheckCart
	  If btnCartStatus = False Then
		 CreateCart
		 Session("ShopingCart").Add  ID,Num
		 Exit Sub
	  ElseIf btnCartStatus = True Then
	  
		 If CheckItem(ID) = True Then
			'//aryAddCart = Session("ShopingCart")
			If Session("ShopingCart").Exists(ID) = True Then
			   Dim NumTmp
			   NumTmp = Session("ShopingCart").Item(ID)		'ԭ������
			   Session("ShopingCart")(ID) = NumTmp + Num
			   Exit Sub
			End If
		 ElseIf CheckItem(ID) = False Then
			'//aryAddCart = Session("ShopingCart")
			Session("ShopingCart").Add  ID,Num
			Exit Sub
		 End If
		 
	  End If
	End Sub
	
	
	Public Sub showCart()
		If CheckCart = False Then
			Response.Write "���Ĺ��ﳵ��û����Ʒ"
			Exit Sub
		End if
		'//aryCart=Session("ShopingCart")
		'����˶�����Ϣ��
		Dim strBillNo
		strBillNo = CFTGetServerDate_BillNo	'WL�Լ���˾�����̻������ź���;
		Response.Write "<span style=""font-size:12px;"">����ǰ������ˮ�ţ�"& strBillNo &"</span><br />"
		
		Dim productCount,moneyCount
			
			'Then CheckCart=True
			'==========WL==========		Begin
			Dim IDs,Nums
			IDs = Session("ShopingCart").keys
			Nums= Session("ShopingCart").items
			%>
			
			<!--#include file="connShop_Product.asp"-->
			
			<%
			'Set rs_AllProduct = conn.Execute("Select * From Products Where Discontinued=0 Order By id Desc")
			'ProductID   	��ƷID
			'ProductName   	��Ʒ����
			'SupplierID   	��Ӧ��ID
			'CategoryID   	����ID
			'QuantityPerUnit   ÿ��λ�ĺ���
			'UnitPrice   	����
			'UnitsInStock   ��ǰ�����
			'UnitsOnOrder   ��ǰ������
			'ReorderLevel   �ٶ��� ˮƽ
			'Discontinued   �Ƿ��Ѿ������˲�Ʒ
			
			product_desc = ""
			total_money  = 0
		
		Response.Write "<style type=""text/css"">"&_
						"<!--"&_
						
						"#Table_Cart {"&_
							"font-size: 13px;"&_
							"color: #000000;"&_
							"/*font-weight: bold;*/"&_
							"font-family: Arial, Helvetica, sans-serif;"&_
							"width: 600px;"&_
							"margin-left: 3px;"&_
							"padding-left: 3px;"&_
							"background-color: #FFFFFF;"&_
						"}"&_
						
						"#Table_Cart caption {"&_
							"font-weight: normal;"&_
						"}"&_
						
						
						"#Table_Cart thead {"&_
							"background-color: #7F7573; /*background-image:url(117720930.jpg)*/"&_
							"margin-left: 18px;"&_
							"color: #ffffff;"&_
						"}"&_
						"#Table_Cart thead tr th {"&_
							"font-weight: bold;"&_
							"text-align: left;"&_
							
							"padding-top:   10px;"&_
							"padding-bottom: 8px;"&_
							
							"padding-left: 8px;"&_
							"padding-right: 8px;"&_
						"}"&_
						
						
						"#Table_Cart tfoot {"&_
							"background-color: #ECE6E6; /*background-image:url(117720930.jpg)*/"&_
							"margin-left: 18px;"&_
							"color: #ffffff;"&_
						"}"&_
						"#Table_Cart tfoot tr th {"&_
							"font-size: 11px;"&_
							"font-weight: bold;"&_
							"text-align: center;"&_
							
							"padding-top:   10px;"&_
							"padding-bottom: 8px;"&_
							
							"padding-left: 8px;"&_
							"padding-right: 8px;"&_
						"}"&_
						
						
						
						
						"#Table_Cart tbody tr td {"&_
							"border-bottom: 1px solid #ffffff;"&_
						"}"&_
						"#Table_Cart tbody tr td {"&_
							"font-size: 12px;"&_
							"font-weight: normal;"&_
							"text-align: left;"&_
							
							"padding-left: 8px;"&_
							"padding-right: 8px;"&_
							
							"line-height: 28px;"&_
							"background-color: #F7F4F4;"&_
						"}"&_
						
						"#Table_Cart tbody tr td a {"&_
							"border: 1px solid #ECE6E6;"&_
							"background-color: #ffffff;"&_
							"text-decoration: none;"&_
							"padding: 3px;"&_
							"color: #7F7573;"&_
						"}"&_
						
						
						"button {"&_
							"background-image:url(../image_Shop/ShopingCartdsdfsaf.gif);"&_
							"background-color: #7F7573;"&_
							"padding-top: 3px;"&_
							"color: #ffffff;"&_
							"font-weight: bold;"&_
							
							"border: 2px solid #ECE6E6;"&_
						"}"&_

						
						
						"-->"&_
						"</style>"
'						a:link { text-decoration: none;color: blue}
'���� a:active { text-decoration:blink}
'���� a:hover { text-decoration:underline;color: red} 
'���� a:visited { text-decoration: none;color: green}

		
		Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""��ժҪ"" id=""Table_Cart"">"&_
			  
			  "<caption>"&_
				"<img src=""../image_Shop/ShopingCart.gif"">���ĳ������У�&nbsp;&nbsp;&nbsp;������200Ԫ���˷� <img src=""../image_Shop/gwlc_shop__r3_c6.gif""> "&_
			  "&nbsp;&nbsp;&nbsp;<img src=""../image_Shop/www.gif""></caption>"&_
				
			    "<thead>"&_
					"<tr>"&_
						"<th>���</th>"&_
						"<th>��Ʒ����</th>"&_
						"<th>�ߴ�</th>"&_
						"<th>��������</th>"&_
						"<th>����</th>"&_
						"<th>С��</th>"&_
						"<th>��Ū�ҵĹ��ﳵ</th>"&_
					"</tr>"&_
				"</thead>"&_
				
				"<tfoot>"&_
					"<tr>"&_
						"<th colspan=""7"">"&_
						"Copyright 2007 - 2008 myhomestay.com.cn All Rights Reserved    ��ICP��07019323��"&_
						"</th>"&_
					"</tr>"&_
				"</tfoot>"&_
			  
			    "<tbody>"
				
			Dim strProducts
			strProducts = ""			
			For i=0 To Session("ShopingCart").Count-1
				
				'Begin ͳ��
				Set rs_Product = conn.Execute("Select * From Products Where Discontinued=0 And ProductID=" & IDs(i) )
				If Not rs_Product.Eof Then
		strNoHTML_ProductName = inHTML( noHTML( Trim(rs_Product("ProductName"))) )
		strProducts = "<tr>"&_
					
					"<td valign=""top"">"&_
					""& rs_Product("ProductID") &""&_
					"</td>"&_
					
					"<td valign=""top"">"&_
					""& strNoHTML_ProductName &_
					"</td>"&_
					
					"<td valign=""top"">"&_
					"--"&_
					"</td>"&_
					
					"<td valign=""top"">"&_
					""& Nums(i) &"��" &_
					"</td>"&_
					
					"<td valign=""top"">"&_
					"��"& rs_Product("UnitPrice") &_
					"</td>"&_
					
					"<td valign=""top"">"&_
					"��"& rs_Product("UnitPrice") &_
					"</td>"&_
					
					"<td valign=""top"">"&_
					" <a href=""MyShopingCar.asp?myAction=addProduct&ID="& IDs(i) &""">+</a>&nbsp;<a href=""MyShopingCar.asp?myAction=modifyNumber&ID="& IDs(i) &"&quantity="& Nums(i)-1 &""">-</a>&nbsp;<a href=""#"" onclick=""alert('Alarm Alarm !');"">�޸�</a>&nbsp;<a href=""MyShopingCar.asp?myAction=deleteProduct&ID="& IDs(i) &""" onClick='return confirm(""ȷ���Ƴ���"& strNoHTML_ProductName &"��Ʒ��"");'>�Ƴ�</a>&nbsp;"&_
					"</td>"&_
					
				  "</tr>" & strProducts
		
					product_desc = product_desc  & Trim(rs_Product("ProductName")) & "|"	'WL ������hidden���䣻
					total_money = total_money + ( rs_Product("UnitPrice")*100 * Nums(i) )		'WL ������hidden���䣻
					
				Else
					Response.Write "<font color=red>��Ʒ"& IDs(i) &"�쳣��ʧ���������������ӳ������...</font><br />"
				End If
				rs_Product.Close
				'End   ͳ��
				
				'//Response.Write IDs(i) &":"& Nums(i) &"<br />"
			Next
			
		Response.Write strProducts
			
			
			If Right(product_desc,1)="|" Then product_desc= Left(product_desc, Len(product_desc)-1)	'WL ȥ�����ұߵ�һ�� "|" ��
			
			'==========WL==========		End
			  
		Response.Write "<tr>"&_
		
							"<td valign=""top"" align=""left"">"&_
							" "&_
							"</td>"&_
							
							"<td valign=""top"" align=""left"">"&_
							" "&_
							"</td>"&_
							
							"<td valign=""top"" align=""left"">"&_
							" "&_
							"</td>"&_
							
							"<td valign=""top"" align=""left"">"&_
							" "&_
							"</td>"&_
							
							"<td valign=""top"" align=""left"">"&_
							" "&_
							"</td>"&_
							
							"<!--<td valign=""top"" align=""left"">"&_
							" "&_
							"</td>-->"&_
							
							"<td valign=""top"" align=""left"" colspan=""2"">"&_
							"�ܽ��&nbsp;&nbsp;<font color=red size=4><u>"& FormatNumber(total_money*0.01 , 2) & "</u> </font>&nbsp;Ԫ<font color=#003370 style=""font-weight: bold;"">RMB</font>." &_
							"</td>"&_
						
					  "</tr>"&_
					  
				  "<tbody>"&_
			  
			  "</table>"
			  
			  
		
		Response.Write "<form action=../tenpay_myhomestay/md5_request.asp method=post>"
			Response.Write ""&_
							"<input type=""hidden"" name=""str_BillNo"" value="""& strBillNo &""" />"&_
							"<input type=""hidden"" name=""product_desc"" value="""& product_desc &""" />"&_
							"<input type=""hidden"" name=""total_money"" value="""& total_money &""" />"

							
			Response.Write ""&_
							
							"&nbsp;&nbsp;&nbsp;"&_
							"<BUTTON type=""button"" name=""close"" onclick=""window.close();"">��������(Close)</BUTTON>"&_
							"&nbsp;&nbsp;&nbsp;"&_
							"<BUTTON type=""submit"" >����ȷ�ϣ�������һ�����֧����</BUTTON>"
'							"<input type=""submit"" name=""submit1"" value=""����ȷ�ϣ�������һ�����֧����"" /> "&_
'							"&nbsp;&nbsp;&nbsp;"&_
'							"<input type=""button"" name=""close"" value=""�ر�"" onclick=""window.close();"" />"	
		
		Response.Write "</form>"
		
		
		conn.Close
		
	End Sub
	
	
	
	
	
	' WL�Լ���˾�����̻������ź��� �� ԭ����ȡ������ ʱ�䣬��ʽMMDDHHMMSS ��10λ���ȣ�
	Public Function CFTGetServerDate_BillNo 
		Dim strTmp, iYear,iMonth,iDate 
		'//iYear = Year(Date) 
		'//iMonth = Month(Date) 
		'//iDate = Day(Date) 
		
		iHour = Hour(Now)
		iMinute = Minute(Now)
		iSecond = Second(Now)
		
		'//strTmp = CStr(iYear)
		strTmp = Right(SESSION.SessionID, 4)		'WL
		
	'//	If iMonth < 10 Then 
	'//		strTmp = strTmp & "0" & Cstr(iMonth)
	'//	Else 
	'//		strTmp = strTmp & Cstr(iMonth)
	'//	End If
	'//	
	'//	If iDate < 10 Then 
	'//		strTmp = strTmp & "0" & Cstr(iDate) 
	'//	Else 
	'//		strTmp = strTmp & Cstr(iDate) 
	'//	End If
	'//	
		
		
		If iHour < 10 Then 
			strTmp = strTmp & "0" & Cstr(iHour) 
		Else 
			strTmp = strTmp & Cstr(iHour) 
		End If
		
		If iMinute < 10 Then 
			strTmp = strTmp & "0" & Cstr(iMinute) 
		Else 
			strTmp = strTmp & Cstr(iMinute) 
		End If
		
		If iSecond < 10 Then 
			strTmp = strTmp & "0" & Cstr(iSecond) 
		Else 
			strTmp = strTmp & Cstr(iSecond) 
		End If
		
		CFTGetServerDate_BillNo = strTmp 
	End Function
	
	
	Function noHTML(str) 
		Dim re 
		Set re = new RegExp 
		re.IgnoreCase = True 
		re.Global = True 
		're.Pattern = "(\<.[^\<]*\>)" 
		'str = re.replace(str," ") 
		re.Pattern = "(\<img[^\<]*\>)" 
		str = re.Replace(str," ") 
		nohtml = str 
		Set re = Nothing 
	End Function
	
	Function inHTML(str)
		Dim sTemp
		sTemp = str
		inHTML = ""
		If IsNull(sTemp) = True Then
			Exit Function
		End If
		sTemp = Replace(sTemp, "&", "&amp;")
		sTemp = Replace(sTemp, "<", "&lt;")
		sTemp = Replace(sTemp, ">", "&gt;")
		sTemp = Replace(sTemp, Chr(34), "&quot;")
		sTemp = Replace(sTemp, "'", "&#39;")
		
		inHTML = sTemp
	End Function


	
End Class

%>

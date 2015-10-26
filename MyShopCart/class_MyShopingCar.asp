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
			   NumTmp = Session("ShopingCart").Item(ID)		'原来数量
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
			Response.Write "您的购物车里没有商品"
			Exit Sub
		End if
		'//aryCart=Session("ShopingCart")
		'输出此订单信息：
		Dim strBillNo
		strBillNo = CFTGetServerDate_BillNo	'WL自己公司生成商户订单号函数;
		Response.Write "<span style=""font-size:12px;"">您当前订单流水号："& strBillNo &"</span><br />"
		
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
			'ProductID   	产品ID
			'ProductName   	产品名字
			'SupplierID   	供应商ID
			'CategoryID   	种类ID
			'QuantityPerUnit   每单位的含量
			'UnitPrice   	单价
			'UnitsInStock   当前库存量
			'UnitsOnOrder   当前订购量
			'ReorderLevel   再订购 水平
			'Discontinued   是否已经放弃此产品
			
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
'　　 a:active { text-decoration:blink}
'　　 a:hover { text-decoration:underline;color: red} 
'　　 a:visited { text-decoration: none;color: green}

		
		Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""无摘要"" id=""Table_Cart"">"&_
			  
			  "<caption>"&_
				"<img src=""../image_Shop/ShopingCart.gif"">您的车筐里有：&nbsp;&nbsp;&nbsp;购物满200元免运费 <img src=""../image_Shop/gwlc_shop__r3_c6.gif""> "&_
			  "&nbsp;&nbsp;&nbsp;<img src=""../image_Shop/www.gif""></caption>"&_
				
			    "<thead>"&_
					"<tr>"&_
						"<th>序号</th>"&_
						"<th>产品名称</th>"&_
						"<th>尺寸</th>"&_
						"<th>购买数量</th>"&_
						"<th>单价</th>"&_
						"<th>小计</th>"&_
						"<th>摆弄我的购物车</th>"&_
					"</tr>"&_
				"</thead>"&_
				
				"<tfoot>"&_
					"<tr>"&_
						"<th colspan=""7"">"&_
						"Copyright 2007 - 2008 myhomestay.com.cn All Rights Reserved    京ICP备07019323号"&_
						"</th>"&_
					"</tr>"&_
				"</tfoot>"&_
			  
			    "<tbody>"
				
			Dim strProducts
			strProducts = ""			
			For i=0 To Session("ShopingCart").Count-1
				
				'Begin 统计
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
					""& Nums(i) &"个" &_
					"</td>"&_
					
					"<td valign=""top"">"&_
					"￥"& rs_Product("UnitPrice") &_
					"</td>"&_
					
					"<td valign=""top"">"&_
					"￥"& rs_Product("UnitPrice") &_
					"</td>"&_
					
					"<td valign=""top"">"&_
					" <a href=""MyShopingCar.asp?myAction=addProduct&ID="& IDs(i) &""">+</a>&nbsp;<a href=""MyShopingCar.asp?myAction=modifyNumber&ID="& IDs(i) &"&quantity="& Nums(i)-1 &""">-</a>&nbsp;<a href=""#"" onclick=""alert('Alarm Alarm !');"">修改</a>&nbsp;<a href=""MyShopingCar.asp?myAction=deleteProduct&ID="& IDs(i) &""" onClick='return confirm(""确定移除此"& strNoHTML_ProductName &"商品吗？"");'>移除</a>&nbsp;"&_
					"</td>"&_
					
				  "</tr>" & strProducts
		
					product_desc = product_desc  & Trim(rs_Product("ProductName")) & "|"	'WL 仅用于hidden传输；
					total_money = total_money + ( rs_Product("UnitPrice")*100 * Nums(i) )		'WL 仅用于hidden传输；
					
				Else
					Response.Write "<font color=red>产品"& IDs(i) &"异常丢失，请您将此情况反映给厂商...</font><br />"
				End If
				rs_Product.Close
				'End   统计
				
				'//Response.Write IDs(i) &":"& Nums(i) &"<br />"
			Next
			
		Response.Write strProducts
			
			
			If Right(product_desc,1)="|" Then product_desc= Left(product_desc, Len(product_desc)-1)	'WL 去掉最右边的一个 "|" ；
			
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
							"总金额&nbsp;&nbsp;<font color=red size=4><u>"& FormatNumber(total_money*0.01 , 2) & "</u> </font>&nbsp;元<font color=#003370 style=""font-weight: bold;"">RMB</font>." &_
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
							"<BUTTON type=""button"" name=""close"" onclick=""window.close();"">继续购物(Close)</BUTTON>"&_
							"&nbsp;&nbsp;&nbsp;"&_
							"<BUTTON type=""submit"" >我已确认，进入下一步完成支付！</BUTTON>"
'							"<input type=""submit"" name=""submit1"" value=""我已确认，进入下一步完成支付！"" /> "&_
'							"&nbsp;&nbsp;&nbsp;"&_
'							"<input type=""button"" name=""close"" value=""关闭"" onclick=""window.close();"" />"	
		
		Response.Write "</form>"
		
		
		conn.Close
		
	End Sub
	
	
	
	
	
	' WL自己公司生成商户订单号函数 （ 原理：获取服务器 时间，格式MMDDHHMMSS ，10位长度）
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

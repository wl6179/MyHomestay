<!--#include file="class_MyShopingCar.asp"-->

<%     
'可以把下面的代码放在一个asp文件里，include一下。
'通过 Request("myAction") 来判定要进行的处理！
'使用方法范例：
'===============================
Set ShopingCart = new class_ShopingCart
'ShopingCart.start

myAction = Trim(Request("myAction"))

Select Case myAction
Case "addProduct"
       ID 		= Trim(Request("ID"))
       quantity = Trim(Request("quantity"))
       If quantity = "" Then
          quantity = 1
       End If
       ShopingCart.AddItem  ID,quantity
	   
	   Response.Redirect "MyShopingCar.asp"
	   
Case "addManyProduct"
	   myNumber = Split(Request("quantity"), ",")
	   myID     = Split(Request("ID"), ",")
       For i=0 To Ubound(myID)
'//Response.Write "=="&myNumber(i)
		  If isNumeric(myNumber(i)) Then
		  	 If myNumber(i)>0 Then
		  	 	ShopingCart.AddItem  CLng(myID(i)),CLng(myNumber(i))
			 End If
		  End IF
       Next
	   
	   Response.Redirect "MyShopingCar.asp"
	   
Case "modifyNumber"
	   myNumber = Split(Request("quantity"), ",")
	   myID     = Split(Request("ID"), ",")
       For i=0 To Ubound(myID)
		  If isNumeric(myNumber(i)) Then
			  If myNumber(i)<=0 Then
				 myNumber(i) =1
			  Else
			  	 '没有特殊情况，保持数量数字；
			  End If
		  Else	'如果数量非法，则统一变成数量1
		  	  myNumber(i) = 1
		  End IF
		  ShopingCart.UpdateItem  CLng(myID(i)),CLng(myNumber(i))
       Next
	   
	   Response.Redirect "MyShopingCar.asp"
	   
Case "deleteProduct"
       ID = Trim(Request("ID"))
       ShopingCart.RemoveItem(Clng(ID))
	   
	   Response.Redirect "MyShopingCar.asp"
	   
Case "clearCart"
       Session("ShopingCart") = NULL
	   
	   Response.Redirect "MyShopingCar.asp"
	   
End Select
'session.Abandon


ShopingCart.showCart()
'===================================
%> 
 

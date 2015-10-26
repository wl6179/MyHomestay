
<%

Response.Write getrandStr()

'返回一个时间和随机数连接后的字符串,用于构建文件名
Function getrandStr()
   Dim RanNum
   Randomize
   RanNum = Int(90000*rnd)+10000
   getrandStr = Year(now)&Month(now)&Day(now) &"-"& Hour(now)&Minute(now)&Second(now) &"-"& RanNum
End Function
	
%>
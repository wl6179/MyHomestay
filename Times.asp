
<%

Response.Write getrandStr()

'����һ��ʱ�����������Ӻ���ַ���,���ڹ����ļ���
Function getrandStr()
   Dim RanNum
   Randomize
   RanNum = Int(90000*rnd)+10000
   getrandStr = Year(now)&Month(now)&Day(now) &"-"& Hour(now)&Minute(now)&Second(now) &"-"& RanNum
End Function
	
%>
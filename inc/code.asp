<%
'------------------
'ͼƬ��֤��
'netdust
'http://www.oyj.cn
'05-8-29 ����
'��ǰ��ͼƬ��֤������ױ��ƽ⣬������������Ͻ��иĽ�����ʹ֪��Դ����
'Ҳ�����ƽ⡣
'���Ŵ���ÿ��һ  ->  ��ٷ��Ĵ������Լ100*100��
'�ʵ��趨�ӵ㼸�ʿ�����Ч�������ʶ����ȷ��
'05-12-20 ������һ������
'------------------

Option Explicit

Call Com_CreatValidCode("GetCode")

Rem ������֤��ͼƬ
Sub Com_CreatValidCode(pSN)
	Const codeLen = 4						'��֤��λ��
	Const cOdds = 3 						'�ӵ���ֵĻ���
	Const dbtTimes = 1					'���Ŵ�������ȫ���ǣ���ò�ҪС��3��

	Const cAmount = 10 					'�ֿ�����
	Const cCode = "0123456789"	'�ֿ��Ӧ���ַ�
	Const UnitWidth = 16				'�ֿ�(ҪΪ4�ı���)
	Const UnitHeight = 13				'�ָ�
	Const DotsLimit = 5				'ÿ��ɾ����Ч�������(�����޷���Ϊʶ��)
	Const tryCount = 5					'����ɾ����Ч�㳬�����޵ĳ��Դ�������

	'-----------

	Randomize
	Dim i, ii, iii

	' ��ֹ����
	Response.Expires = -9999
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "cache-ctrol","no-cache"
	Response.ContentType = "Image/BMP"

	' ��ɫ������(�ַ�������)
	Dim vColorData(1)
	vColorData(0) = ChrB(0) & ChrB(0) & ChrB(0)  ' ��0����0����0����ɫ��
	vColorData(1) = ChrB(255) & ChrB(255) & ChrB(255) ' ��250����236����211��ǳ��ɫ��

	' �ַ�������(�����Լ��޸ģ�����޸��˳ߴ磬�ǵð�ǰ����趨Ҳ����)
	Dim vNumberData(9)
	vNumberData(0) = "1111000000001111111000000000011111100111111001111110011111100111111001111110011111100111111001111110011111100111111001111110011111100111111001111110011111100111111001111110011111100000000001111111000000001111"
	vNumberData(1) = "1111110001111111111100000111111111100000011111111100110001111111111111000111111111111100011111111111110001111111111111000111111111111100011111111111110001111111111111000111111111100000000011111110000000001111"
	vNumberData(2) = "1111110000011111111110000000111111110001110011111110001111001111111111111001111111111111001111111111111001111111111111001111111111111001111111111111001111001111111001111100111111100000000011111110000000001111"
	vNumberData(3) = "1111100000011111111100000000111111100111111001111110011111001111111111111001111111111110001111111111111000111111111111111001111111111111110011111110011111100111111001111110011111110000000011111111100000011111"
	vNumberData(4) = "1111111100111111111110110011111111110011001111111111001100111111111001110011111111001111001111111000000000000011100000000000001111111111001111111111111100111111111111110011111111111111001111111111111100111111"
	vNumberData(5) = "1110000000000111110011111111111111001111111111111100111111111111110011111111111111001100000011111100000111100111111111111110011111111111111001111111111111100111110011111110011111001111111001111110000000001111"
	vNumberData(6) = "1111110000011111111110000000111111110011111001111110011111111111111001111111111111100100000111111110000000001111111000111110011111100111111001111110011111100111111001111110011111110000000011111111100000011111"
	vNumberData(7) = "1110000000000111111000000000011111100111111001111110011111100111111111111100111111111111110011111111111110011111111111110011111111111111001111111111111100111111111111110011111111111111001111111111111100111111"
	vNumberData(8) = "1111100000011111111100000000111111100111111001111110011111100111111001111110011111110000000011111111000000001111111100111100111111100111111001111110011111100111111001111110011111110000000011111111100000011111"
	vNumberData(9) = "1111100000011111111100000000111111100111111001111110011111100111111001111110011111110000000001111111000000100111111111111110011111111111111001111111111111100111111001111100111111110000000011111111100000011111"

	' ��������ַ�
	Dim vCodes
	ReDim vCode(codeLen-1)
	For i = 0 To codeLen-1
	  vCode(i) = Int(Rnd * cAmount)
	  vCodes = vCodes & Mid(cCode, vCode(i) + 1, 1)
	  vCode(i) = pcd_doubter(vNumberData(vCode(i)),UnitWidth,UnitHeight,DotsLimit,tryCount,dbtTimes)
	Next

	Session(pSN) = vCodes  '��¼��Session

	' ���ͼ���ļ�ͷ
	Response.BinaryWrite ChrB(66) & ChrB(77) & Num2ChrB(54+UnitWidth*UnitHeight*CodeLen*3,4) & ChrB(0) & ChrB(0) &_
	  ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) &_
	  ChrB(0) & ChrB(0) & Num2ChrB(UnitWidth*CodeLen,4) & Num2ChrB(UnitHeight,4) &_
	  ChrB(1) & ChrB(0)

	' ���ͼ����Ϣͷ
	Response.BinaryWrite ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & Num2ChrB(UnitWidth*UnitHeight*CodeLen*3,4) &_
	  ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) &_
	  ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
	  ChrB(0) & ChrB(0)

	For i = UnitHeight-1 To 0 Step -1  ' ����������
	  For ii = 0 To codeLen-1  ' ����������
	   For iii = 1 To UnitWidth ' ������������
	    If Rnd * 99 + 1 >= cOdds Then	' ���С����֡������ص����ͼ������
	     Response.BinaryWrite vColorData(Mid(vCode(ii), i * UnitWidth + iii, 1))
	    Else ' ��������ӵ�
	     Response.BinaryWrite vColorData(1-CInt(Mid(vCode(ii), i * UnitWidth + iii, 1)))
	    End If
	   Next
	  Next
	Next
End Sub

Rem �Ե����ֵĵ�����и���
Rem ����˼�룺�ڵ���Χ���������2���˵㣬�������ߣ���λ�ƽϴ��һ�������ᣬ�Ƚ������ϵĵ�ɾ�����ٽ���ɾ��������᷽���Ϸ����·��ĵ㣨���ȷ��������ɾ���㣬�ƶ���Ŀհ��ñ���ɫ����
Function pcd_doubter(ByVal str,UnitWidth,UnitHeight,DotsLimit,tryCount,dbtTimes)
	Randomize
	Dim x1,x2,y1,y2,xOffSet,yOffSet,direction,flag,rows,step,yu,yuStr,i,ii,iii,f1,f2
	For f1=1 To dbtTimes	'���Ŵ���
		For f2=1 To tryCount	'����ɾ����Ч�㳬�����޵ĳ��Դ�������
			'���ȷ��2���˵�
			x1 = int(Rnd*UnitWidth)
			x2 = int(Rnd*UnitWidth)
			y1 = int(Rnd*UnitHeight)
			y2 = int(Rnd*UnitHeight)
			'x,yλ����
			xOffSet = Abs(x2-x1)
			yOffSet = Abs(y2-y1)
			If xOffSet >= yOffSet Then	'��λ�����ϴ�������
				direction = "x"
				ReDim ary(xOffSet)	'������¼���߸���yֵ
				'x2,y2�洢xֵ�ϴ�ĵ�
				If x2 < x1 Then
					i = x1
					x1 = x2
					x2 = i
					i = y1
					y1 = y2
					y2 = i
				End If
				'�жϴ�x1->x2�����᷽���������Ǽ�
				If y2 >= y1 Then
					flag = 1
				Else
					flag = -1
				End If
				'������������ϵ�ķֲ�������ƽ��������еĵ㣬Ȼ���������ʣ��ĵ㵽���У�
				rows = yOffSet + 1	'��ռ����
				step = (xOffSet+1) \ rows	'����ƽ������ĵ�
				yu = (xOffSet+1) Mod rows	'ʣ��ĵ���
				ReDim ary2(rows-1)	'������¼ʣ�����������
				While yu > 0
					i = int(Rnd*rows)
					ary2(i) = ary2(i)&"."		'�����䵽�������һ���ַ�"."
					yu = yu - 1
				WEnd
				iii = 0
				'�����ߵĵ���Ϣ��¼������
				For i=0 To rows-1
					For ii=1 To step+Len(ary2(i))
						ary(iii) = y1 + i*flag
						iii = iii + 1
					Next
				Next
				ii = 0
				'ͳ����������Ч�������
				For i=0 To xOffSet
					If pcd_getDot(x1+i,ary(i),str,UnitWidth) = "0" Then ii = ii + 1
				Next
			Else
				'��������yΪ���ᣬԭ����xʱ��ͬ
				direction = "y"
				ReDim ary(yOffSet)
				If y2 < y1 Then
					i = x1
					x1 = x2
					x2 = i
					i = y1
					y1 = y2
					y2 = i
				End If
				If x2 >= x1 Then
					flag = 1
				Else
					flag = -1
				End If
				rows = xOffSet + 1
				step = (yOffSet+1) \ rows
				yu = (yOffSet+1) Mod rows
				ReDim ary2(rows-1)
				While yu > 0
					i = int(Rnd*10)
					If i < rows Then
						ary2(i) = ary2(i)&"."
						yu = yu - 1
					End If
				WEnd
				iii = 0
				For i=0 To rows-1
					For ii=1 To step+Len(ary2(i))
						ary(iii) = x1 + i*flag
						iii = iii + 1
					Next
				Next
				ii = 0
				For i=0 To yOffSet
					If pcd_getDot(ary(i),y1+i,str,UnitWidth) = "0" Then ii = ii + 1
				Next
			End If
			'��δ������Ч������������ѭ����ִ�и���
			If ii <= DotsLimit Then Exit For
		Next

		If direction = "x" Then
			'���ȷ�������᷽���ϻ��½����ƶ�
			If int(Rnd*10) > 4 Then
				'���������ϵĵ�
				For i=0 To xOffSet
					'�����ƶ�
					For ii=ary(i) To 1 Step -1
						Call pcd_setDot(x1+i,ii,str,pcd_getDot(x1+i,ii-1,str,UnitWidth),UnitWidth)
					Next
					'���հ�
					Call pcd_setDot(x1+i,0,str,"1",UnitWidth)
				Next
			Else
				For i=0 To xOffSet
					For ii=ary(i) To UnitHeight-2
						Call pcd_setDot(x1+i,ii,str,pcd_getDot(x1+i,ii+1,str,UnitWidth),UnitWidth)
					Next
					Call pcd_setDot(x1+i,UnitHeight-1,str,"1",UnitWidth)
				Next
			End If
		Else
			If int(Rnd*10) > 4 Then
				For i=0 To yOffSet
					For ii=ary(i) To 1 Step -1
						Call pcd_setDot(ii,y1+i,str,pcd_getDot(ii-1,y1+i,str,UnitWidth),UnitWidth)
					Next
					Call pcd_setDot(0,y1+i,str,"1",UnitWidth)
				Next
			Else
				For i=0 To yOffSet
					For ii=ary(i) To UnitWidth-2
						Call pcd_setDot(ii,y1+i,str,pcd_getDot(ii+1,y1+i,str,UnitWidth),UnitWidth)
					Next
					Call pcd_setDot(UnitWidth-1,y1+i,str,"1",UnitWidth)
				Next
			End If
		End If
	Next
	pcd_doubter = str
End Function

Rem �õ�ĳ����ַ�
Function pcd_getDot(x,y,str,UnitWidth)
	pcd_getDot = Mid(str,x+1+y*UnitWidth,1)
End Function

Rem ����ĳ����ַ�
Sub pcd_setDot(x,y,ByRef str,newDot,UnitWidth)
	str = Left(str,x+y*UnitWidth) & newDot & Right(str,Len(str)-x-y*UnitWidth-1)
End Sub

Rem ������תΪbmp��Ҫ�ĸ�ʽ lens��Ŀ���ֽڳ���
Function Num2ChrB(num,lens)
	Dim ret,i
	ret = ""
	While (num>0)
		ret = ret & ChrB(num mod 256)
		num = num \ 256
	WEnd
	For i=Lenb(ret) To lens-1
		ret = ret & chrB(0)
	Next
	Num2ChrB = ret
End Function
%>
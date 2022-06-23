' 照抄 https://github.com/Bidai/ChineseNumber

Option Explicit
	Dim NumberChars		'中文数字
	NumberChars = Array(_
		Array("0", "０", "零", "〇"),_
		Array("1", "１", "一", "壹", "幺"),_
		Array("2", "２", "二", "贰", "两", "俩", "貮", "兩", "倆"),_
		Array("3", "３", "三", "叁"),_
		Array("4", "４", "四", "肆"),_
		Array("5", "５", "五", "伍"),_
		Array("6", "６", "六", "陆", "陸"),_
		Array("7", "７", "七", "柒"),_
		Array("8", "８", "八", "捌"),_
		Array("9", "９", "九", "玖")_
	)
	Dim UnitChars		'中文单位
	UnitChars  = Array(_
					Array("十", "拾"), _
					Array("廿", "念"), _
					Array("百", "佰"), _
					Array("千", "仟"), _
					Array("万", "萬"), _
					Array("亿", "億"), _
					Array("-"), _
					Array("-"), _
					Array("-"), _
					array("-")  _
	)
	Dim UnitValues : UnitValues = Array(10, 20, 100, 1000, 10000, 100000000)

'False, True  : -1
'False, False : 0
'True, True   : 0
'True, False  : 1
Function CompareValue(L, R)
	CompareValue = Abs(L and not R) - Abs(R and not L)
End Function

Function StrToChars(str)
	Dim Chars(), i 
	ReDim Chars(Len(str))
	For i = 1 To Len(str)
		Chars(i -1) = Mid(str, i, 1)
	Next
	StrToChars = Chars
End Function

Function GetChsValue(types, values, idx, [Len])
	
	Dim Result, LastPT
	Dim i, j, k
	LastPT = [Len]
	For i = [len] - 1 To 0 Step -1
		If types(idx + i) Then
			If (LastPT < [Len]) And (values(idx + i) <= values(idx + LastPT)) Then
				j = i - 1
				On Error Resume Next	'报错继续，为了忽略下标越界
				Do While (j >= 0) And (values(idx + j) <= values(idx + lastpt))		'会造成下标越界
					j = j - 1
					If (j <= 0) Then Exit Do
				Loop
				If j < 0 Then j = j + 1
				If j < 0 Or (types(idx + j) And values(idx + j) > values(idx + LastPT)) Then j = j + 1
				
				
				values(idx + j) = GetChsValue(types, values, idx + j, LastPT - j)
				types(idx + j) = False
				k = 1
				Do While lastpt + k <= len
					types(idx + j + k) = types(idx + lastpt + k - 1)
					values(idx + j + k) = values(idx + lastpt + k - 1)
					k = k + 1
				Loop
				[Len] = [Len] - (LastPT - j - 1)
				LastPT = j + 1
				i = j
			Else
				LastPT = i
			End If
		End If
	Next 'i
	
	'扫描完毕，开始计算
	For i = 0 To [Len] - 1
		If types(idx + i) Then '遇到单位
			If i <> 0 Then
				If types(idx + i - 1) Or (0 = values(idx + i - 1)) Then '非常规处理：单位紧挨如"千百”，或数值为0如"万零百"
					Result = Result + values(idx + i)
				Else
					Result = Result + values(idx + i - 1) * values(idx + i) '正常赋值
				End If
			Else '单位打头
				Result = values(idx + i)
			End If
'		ElseIf (i <> 0) And (Not types(idx + i - 1)) Then '遇到数字叠加, 会下标越界
		ElseIf i <> 0 Then
			If Not types(idx + i - 1) Then values(idx + i) = values(idx + i) + values(idx + i - 1) * 10
		End If

	Next 'i
	If Not types(idx + len - 1) Then Result = Result + values(idx + len - 1)

	GetChsValue = result
End Function

Function ChineseToULong(chs)
	Dim StrChars, CharsLength, Types(), Values(), StrLen, Result
	Dim ci, ni, NValue, UValue 'char index, number index
	Result = 0
	StrLen = Len(chs)
	
	If StrLen = 0 Then '空字符串
		ChineseToULong = 0
		Exit Function
	End If
	
	StrChars    = StrToChars(chs)	'string 转 char 数组
	CharsLength = UBound(StrChars)
	ReDim Types(CharsLength)	'Types 元素类型,跟 values 对应。false 表示数值，true 表示单位
	ReDim Values(CharsLength)

	'数字字符串信息记录
	For ci = 0 To CharsLength
		For ni = 0 To UBound(NumberChars)
			NValue = InStr(1, Join(NumberChars(ni), ""), StrChars(ci))	'查数字
			UValue = InStr(1, Join(UnitChars  (ni), ""), StrChars(ci))	'查单位
			Select Case CompareValue(NValue > 0, UValue > 0)
				Case -1		'单位
					Values(ci) = UnitValues(ni)
					Types(ci)  = True
					Exit For 
				Case  1		'数字
					Values(ci) = ni
					Types(ci)  = False
					Exit For
				Case Else	'啥玩意？
				
			End Select
		Next	'ni
	Next	'ci
	
	'二百五等情况修正
	If (CharsLength > 1) And (Not Types(CharsLength - 1)) And (Types(CharsLength - 2)) And (Values(CharsLength - 2) >= 100) Then
		Values(CharsLength - 1) = Values(CharsLength - 1) * (Values(CharsLength - 2) / 10)
	End If

	Result = GetChsValue(Types, Values, 0, CharsLength)
	ChineseToULong = Result
End Function

Sub test
  Dim test_array, i
  test_array = Array(_
    Array("十五万零二百五", 150250, 0), _
    Array("两万四", 24000, 0), _
    Array("贰仟贰佰零伍万", 22050000, 0), _
    Array("十一", 11, 0), _
    Array("三百零五万零三十", 3050030, 0), _
    Array("贰拾万", 200000, 0), _
    Array("贰仟贰佰零伍万", 22050000, 0), _
    Array("贰佰陆拾叁", 263, 0), _
    Array("贰万捌仟贰佰伍拾陆", 28256, 0), _
    Array("贰拾叁", 23, 0), _
    Array("壹仟柒佰零陆", 1706, 0), _
    Array("叁拾万零伍拾", 300050, 0), _
    Array("贰拾捌万贰仟玖佰叁拾壹", 282931, 0), _
    Array("廿万", 200000, 0), _
    Array("一八七九五四九一九七一", 18795491971, 0), _
    Array("壹亿贰仟叁佰肆拾伍万陆仟柒佰捌拾玖", 123456789, 0) _
  )

  For Each i In test_array
	  i(2) = ChineseToULong(i(0))
	  WScript.Echo Right(String(20, "　") & i(0), 20), Right(String(12, " ") & i(1), 12), "=", Left(i(2) & String(12, " "), 12), CStr(i(1) = i(2))
  Next
End Sub

test



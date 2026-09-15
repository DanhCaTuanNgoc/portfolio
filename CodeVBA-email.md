Private Sub Worksheet_Change(ByVal Target As Range)
	Dim rngProcess As Range
	Dim cell As Range
    
	'========================================================
	' 1. CHI KICH HOAT KHI NHAP EMAIL VAO COT C
	'========================================================
	Set rngProcess = Intersect(Target, Me.Columns("C"))
	If rngProcess Is Nothing Then Exit Sub
    
	Application.EnableEvents = False
	Application.ScreenUpdating = False
	On Error GoTo ErrorHandler
    
	'========================================================
	' 2. KET NOI TOI CAC SHEET DU LIEU
	'========================================================
	Dim wsUser As Worksheet, wsKho As Worksheet, wsRole As Worksheet
	On Error Resume Next
	Set wsUser = ThisWorkbook.Worksheets("User List")
	Set wsKho = ThisWorkbook.Worksheets("Danh sách kho")
	Set wsRole = ThisWorkbook.Worksheets("Role Mapping")
	On Error GoTo ErrorHandler
    
	If wsUser Is Nothing Or wsKho Is Nothing Or wsRole Is Nothing Then
		MsgBox "Khong tim thay mot trong cac sheet du lieu (User List, Danh sách kho, Role Mapping). Vui long kiem tra lai ten sheet!", vbCritical, "Loi ket noi Sheet"
		GoTo CleanUp
	End If
    
	Dim lastRowKho As Long, lastRowRole As Long
	lastRowKho = wsKho.Cells(wsKho.Rows.Count, "C").End(xlUp).Row
	If lastRowKho < 2 Then lastRowKho = wsKho.Cells(wsKho.Rows.Count, "B").End(xlUp).Row
	lastRowRole = wsRole.Cells(wsRole.Rows.Count, "A").End(xlUp).Row
    
	' Load du lieu Danh sach kho vao mang de toi uu toc do va chuan hoa
	Dim arrKho() As Variant, arrKhoNorm() As String
	If lastRowKho >= 2 Then
		arrKho = wsKho.Range("A2:G" & lastRowKho).Value
		ReDim arrKhoNorm(1 To UBound(arrKho, 1))
		Dim k As Long
		For k = 1 To UBound(arrKho, 1)
			arrKhoNorm(k) = ChuanHoaTenDoi(arrKho(k, 3))
		Next k
	End If
    
	' Load du lieu Role Mapping
	Dim arrRole() As Variant
	If lastRowRole >= 2 Then
		arrRole = wsRole.Range("A2:B" & lastRowRole).Value
	End If
    
	'========================================================
	' 3. XU LY TUNG EMAIL DUOC NHAP/DAN
	'========================================================
	For Each cell In rngProcess
		If cell.Row < 2 Then GoTo NextCell
        
		Dim emailInput As String
		emailInput = Trim(CStr(cell.Value))
        
		' Neu xoa email thi xoa trang thong tin tuong ung o cot A va D:J
		If emailInput = "" Then
			cell.Offset(0, -2).ClearContents
			cell.Offset(0, -1).ClearContents
			cell.Offset(0, 1).Resize(1, 7).ClearContents
			GoTo NextCell
		End If
        
		' TIM EMAIL TRONG COT D CUA USER LIST
		Dim matchRes As Variant
		matchRes = Application.Match(emailInput, wsUser.Columns("D"), 0)
        
		If IsError(matchRes) Then
			cell.Offset(0, -2).ClearContents
			cell.Offset(0, -1).ClearContents
			cell.Offset(0, 1).Value = "Not Found"
			cell.Offset(0, 2).Value = "Not Found"
			cell.Offset(0, 3).Resize(1, 5).ClearContents
			GoTo NextCell
		End If
        
		Dim rUser As Long
		rUser = CLng(matchRes)
        
		' --- LAY THONG TIN TU USER LIST ---
		Dim employeeCode As String, emailStr As String, nameStr As String
		Dim salesTeam As String, salesPos As String
        
		employeeCode = CStr(wsUser.Cells(rUser, "B").Value)
		emailStr = CStr(wsUser.Cells(rUser, "D").Value)
		nameStr = CStr(wsUser.Cells(rUser, "C").Value)
		salesTeam = CStr(wsUser.Cells(rUser, "L").Value)
		salesPos = CStr(wsUser.Cells(rUser, "M").Value)
        
		' --- BUOC A: DIEN THONG TIN CO BAN (COT D -> G) ---
		cell.Offset(0, -1).NumberFormat = "@"
		cell.Offset(0, -1).Value = employeeCode
		cell.Value = emailStr
		cell.Offset(0, 1).Value = nameStr
        
		cell.Offset(0, 2).NumberFormat = "yyyy-mm-dd"
		cell.Offset(0, 2).Value = Date
        
		cell.Offset(0, 3).Value = "AVNp@ssw0rd"
		cell.Offset(0, 4).Value = "Ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng"
        
		' --- BUOC B: TINH MAI VAI TRO (COT H) ---
		Dim roleVal As String
		Dim rRole As Long
		Dim foundRole As Boolean
		foundRole = False
        
		If salesPos <> "" And (Not Not arrRole) Then
			For rRole = 1 To UBound(arrRole, 1)
				If UCase(Trim(CStr(arrRole(rRole, 1)))) = UCase(Trim(salesPos)) Then
					roleVal = CStr(arrRole(rRole, 2))
					foundRole = True
					Exit For
				End If
			Next rRole
		End If
        
		If foundRole Then
			cell.Offset(0, 5).Value = roleVal
		Else
				cell.Offset(0, 5).Value = "Null"
			roleVal = ""
		End If
        
		' Dien Cot A: Loai tai khoan
		If HasRoleS(roleVal) Then
			cell.Offset(0, -2).Value = "eSalesSFA"
		Else
			cell.Offset(0, -2).Value = "eSalesBackOffice"
		End If
        
		' --- BUOC C: TINH MA KHO (COT I) VA MA DOI (COT J) ---
		Dim resMaDoi As String, resMaKho As String
		Dim arrInput() As String
		Dim i As Long, r As Long
		Dim searchStr As String, dbStr As String
		Dim foundKho As Boolean
        
		resMaDoi = ""
		resMaKho = ""
        
		If salesTeam <> "" And (Not Not arrKhoNorm) Then
			Dim cleanTeamInput As String
			cleanTeamInput = Replace(salesTeam, vbCrLf, ",")
			cleanTeamInput = Replace(cleanTeamInput, vbCr, ",")
			cleanTeamInput = Replace(cleanTeamInput, vbLf, ",")
			cleanTeamInput = Replace(cleanTeamInput, ";", ",")
            
			arrInput = Split(cleanTeamInput, ",")
            
			For i = LBound(arrInput) To UBound(arrInput)
				Dim itemTeam As String
				itemTeam = Trim(arrInput(i))
                
				If itemTeam <> "" Then
					searchStr = ChuanHoaTenDoi(itemTeam)
					foundKho = False
                    
					For r = 1 To UBound(arrKhoNorm)
						dbStr = arrKhoNorm(r)
						If dbStr <> "" And dbStr = searchStr Then
							Dim strMaDoi As String, strMaKho As String
							strMaDoi = Trim(CStr(arrKho(r, 7)))
							strMaKho = Trim(CStr(arrKho(r, 5)))
                            
							If strMaDoi = "" Then strMaDoi = "Null"
							If strMaKho = "" Then strMaKho = "Null"
                            
							resMaDoi = resMaDoi & strMaDoi & ", "
							resMaKho = resMaKho & strMaKho & ", "
							foundKho = True
							Exit For
						End If
					Next r
                    
					If Not foundKho Then
						resMaDoi = resMaDoi & "Null, "
						resMaKho = resMaKho & "Null, "
					End If
				End If
			Next i
		Else
			resMaDoi = "Null, "
			resMaKho = "Null, "
		End If
        
		If Len(resMaDoi) > 0 Then resMaDoi = Left(resMaDoi, Len(resMaDoi) - 2)
		If Len(resMaKho) > 0 Then
			resMaKho = Left(resMaKho, Len(resMaKho) - 2)
			resMaKho = LocTrungMaKho(resMaKho)
		End If
        
		If resMaKho = "" Then resMaKho = "Null"
		If resMaDoi = "" Then resMaDoi = "Null"
        
		cell.Offset(0, 6).NumberFormat = "@"
		cell.Offset(0, 7).NumberFormat = "@"
		cell.Offset(0, 6).Value = resMaKho
		cell.Offset(0, 7).Value = resMaDoi

NextCell:
	Next cell

CleanUp:
	Application.ScreenUpdating = True
	Application.EnableEvents = True
	Exit Sub

ErrorHandler:
	Application.ScreenUpdating = True
	Application.EnableEvents = True
	MsgBox "Co loi xay ra trong qua trinh xu ly: " & Err.Description, vbExclamation, "Thong bao loi"
End Sub

'============================================================
' CAC HAM PHU TRO DUNG DE LOC TRUNG VA CHUAN HOA
'============================================================
Function HasRoleS(ByVal roleStr As String) As Boolean
	If Trim(roleStr) = "" Or Trim(roleStr) = "Null" Then
		HasRoleS = False
		Exit Function
	End If
    
	Dim arr() As String
	Dim item As Variant
	arr = Split(roleStr, ",")
	For Each item In arr
		If UCase(Trim(CStr(item))) = "S" Then
			HasRoleS = True
			Exit Function
		End If
	Next item
	HasRoleS = False
End Function

Function LocTrungMaKho(ByVal txt As String) As String
	Dim arr() As String
	Dim col As Collection
	Dim i As Long
	Dim result As String
    
	If Trim(txt) = "" Or txt = "Null" Then
		LocTrungMaKho = txt
		Exit Function
	End If
    
	Set col = New Collection
	arr = Split(txt, ", ")
    
	On Error Resume Next
	For i = LBound(arr) To UBound(arr)
		Dim itemK As String
		itemK = Trim(arr(i))
		If itemK <> "" Then col.Add itemK, UCase(itemK)
	Next i
	On Error GoTo 0
    
	For i = 1 To col.Count
		result = result & col(i) & ", "
	Next i
    
	If Len(result) > 0 Then result = Left(result, Len(result) - 2)
	LocTrungMaKho = result
End Function

Function ChuanHoaTenDoi(ByVal text As Variant) As String
	Dim result As String
	If IsError(text) Or IsEmpty(text) Or IsNull(text) Then
		ChuanHoaTenDoi = ""
		Exit Function
	End If
	result = CStr(text)
    
	result = Replace(result, Chr(160), " ")
	result = Replace(result, Chr(10), "")
	result = Replace(result, Chr(13), "")
	result = Replace(result, Chr(9), " ")
	result = Replace(result, ChrW(8203), "")
	result = Replace(result, ChrW(8204), "")
	result = Replace(result, ChrW(8205), "")
	result = Replace(result, ChrW(65279), "")
    
	result = Replace(result, ChrW(8211), "-")
	result = Replace(result, ChrW(8212), "-")
	result = Replace(result, ChrW(8722), "-")
	result = Replace(result, ChrW(45), "-")
    
	Dim combMarks As Variant
	combMarks = Array(768, 769, 770, 771, 772, 774, 775, 776, 777, 778, 779, 780, 781, 782, 783, 784, 785, 786, 787, 788, 789, 790, 791, 792, 793, 794, 795, 803, 804, 805, 821)
	Dim idx As Long
	For idx = LBound(combMarks) To UBound(combMarks)
		result = Replace(result, ChrW(combMarks(idx)), "")
	Next idx
    
	result = BoDau(result)
	result = UCase(Trim(result))
    
	Do While InStr(result, "  ") > 0
		result = Replace(result, "  ", " ")
	Loop
    
	result = Replace(result, " - ", "-")
	result = Replace(result, " -", "-")
	result = Replace(result, "- ", "-")
	ChuanHoaTenDoi = Trim(result)
End Function

Function BoDau(ByVal text As String) As String
	text = ReplaceChars(text, Array(224, 225, 7843, 227, 7841, 259, 7857, 7855, 7859, 7861, 7863, 226, 7847, 7845, 7849, 7851, 7853, 192, 193, 7842, 195, 7840, 258, 7856, 7854, 7858, 7860, 7862, 194, 7846, 7844, 7848, 7850, 7852), "A")
	text = ReplaceChars(text, Array(232, 233, 7867, 7869, 7865, 234, 7873, 7871, 7875, 7877, 7879, 200, 201, 7866, 7868, 7864, 202, 7872, 7870, 7874, 7876, 7878), "E")
	text = ReplaceChars(text, Array(236, 237, 7881, 297, 7883, 204, 205, 7880, 296, 7882), "I")
	text = ReplaceChars(text, Array(242, 243, 7887, 245, 7885, 244, 7891, 7889, 7893, 7895, 7897, 417, 7899, 7901, 7903, 7905, 7907, 210, 211, 7886, 213, 7884, 212, 7890, 7888, 7892, 7894, 7896, 416, 7898, 7900, 7902, 7904, 7906), "O")
	text = ReplaceChars(text, Array(249, 250, 7911, 361, 7913, 432, 7915, 7917, 7919, 7921, 7923, 217, 218, 7910, 360, 7912, 431, 7914, 7916, 7918, 7920, 7922), "U")
	text = ReplaceChars(text, Array(7923, 253, 7927, 7929, 7925, 7922, 221, 7926, 7928, 7924), "Y")
	text = ReplaceChars(text, Array(273, 272), "D")
	BoDau = text
End Function

Function ReplaceChars(ByVal txt As String, ByVal codeArr As Variant, ByVal replaceWith As String) As String
	Dim i As Long
	For i = LBound(codeArr) To UBound(codeArr)
		txt = Replace(txt, ChrW(codeArr(i)), replaceWith)
	Next i
	ReplaceChars = txt
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngProcess As Range
    Dim cell As Range
    
    '========================================================
    ' 1. CHỈ KÍCH HOẠT KHI NHẬP VÀO CỘT B (MÃ NHÂN VIÊN)
    '========================================================
    Set rngProcess = Intersect(Target, Me.Columns("B"))
    If rngProcess Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '========================================================
    ' 2. KẾT NỐI TỚI CÁC SHEET DỮ LIỆU
    '========================================================
    Dim wsUser As Worksheet, wsDoi As Worksheet, wsRole As Worksheet
    On Error Resume Next
    Set wsUser = ThisWorkbook.Worksheets("User List")
    Set wsDoi = ThisWorkbook.Worksheets("Sale team list")
    Set wsRole = ThisWorkbook.Worksheets("Role Mapping")
    On Error GoTo ErrorHandler
    
    If wsUser Is Nothing Or wsDoi Is Nothing Or wsRole Is Nothing Then
        MsgBox "Khong tim thay mot trong cac sheet du lieu (User List, Sale team list, Role Mapping). Kiem tra dung ten sheet!", vbCritical, "Loi ket noi Sheet"
        GoTo CleanUp
    End If
    
    Dim lastRowDoi As Long, lastRowRole As Long
    lastRowDoi = wsDoi.Cells(wsDoi.Rows.Count, "C").End(xlUp).Row
    If lastRowDoi < 2 Then lastRowDoi = wsDoi.Cells(wsDoi.Rows.Count, "B").End(xlUp).Row
    lastRowRole = wsRole.Cells(wsRole.Rows.Count, "A").End(xlUp).Row
    
    ' Load dữ liệu Danh sách đội: dò tên đội ở C, lấy mã đội ở B và mã kho ở D
    Dim arrDoi() As Variant, arrDoiNorm() As String
    If lastRowDoi >= 2 Then
        arrDoi = wsDoi.Range("A2:D" & lastRowDoi).Value
        ReDim arrDoiNorm(1 To UBound(arrDoi, 1))
        Dim k As Long
        For k = 1 To UBound(arrDoi, 1)
            arrDoiNorm(k) = ChuanHoaTenDoi(arrDoi(k, 3)) ' Cột C: Tên đội
        Next k
    End If
    
    ' Load dữ liệu Role Mapping
    Dim arrRole() As Variant
    If lastRowRole >= 2 Then
        arrRole = wsRole.Range("A2:B" & lastRowRole).Value
    End If
    
    '========================================================
    ' 3. XỬ LÝ TỪNG NHÂN VIÊN ĐƯỢC NHẬP/DÁN
    '========================================================
    For Each cell In rngProcess
        If cell.Row < 2 Then GoTo NextCell
        
        Dim empCode As String
        empCode = Trim(CStr(cell.Value))
        
        ' Nếu xóa mã nhân viên thì xóa trắng toàn bộ thông tin tương ứng ở các cột A và C -> J
        If empCode = "" Then
            cell.Offset(0, -1).ClearContents
            cell.Offset(0, 1).Resize(1, 8).ClearContents
            GoTo NextCell
        End If
        
        ' --- TÌM MÃ NHÂN VIÊN (XỬ LÝ ĐỒNG THỜI TEXT VÀ SỐ) ---
        Dim matchRes As Variant
        matchRes = Application.Match(empCode, wsUser.Columns("B"), 0)
        
        If IsError(matchRes) And IsNumeric(empCode) Then
            matchRes = Application.Match(CDbl(empCode), wsUser.Columns("B"), 0)
        End If
        
        If IsError(matchRes) Then
            cell.Offset(0, -1).ClearContents
            cell.Offset(0, 1).Value = "Not Found"
            cell.Offset(0, 2).Value = "Not Found"
            cell.Offset(0, 3).Resize(1, 6).ClearContents
            GoTo NextCell
        End If
        
        Dim rUser As Long
        rUser = CLng(matchRes)
        
        ' --- LẤY THÔNG TIN TỪ USER LIST ---
        Dim emailStr As String, nameStr As String
        Dim salesTeam As String, salesPos As String
        
        emailStr = CStr(wsUser.Cells(rUser, "D").Value)  ' Cột D: HRMS account
        nameStr = CStr(wsUser.Cells(rUser, "C").Value)   ' Cột C: Full name
        salesTeam = CStr(wsUser.Cells(rUser, "L").Value) ' Cột L: Sales team name
        salesPos = CStr(wsUser.Cells(rUser, "M").Value)  ' Cột M: Sales position
        
        ' --- BƯỚC A: ĐIỀN THÔNG TIN CƠ BẢN (CỘT C -> G) ---
        cell.Offset(0, 1).Value = emailStr
        cell.Offset(0, 2).Value = nameStr
        
        cell.Offset(0, 3).NumberFormat = "yyyy-mm-dd"
        cell.Offset(0, 3).Value = Date
        
        cell.Offset(0, 4).Value = "AVNp@ssw0rd"
        
        ' Xuất "Hoạt động" chuẩn Unicode
        cell.Offset(0, 5).Value = "Ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng"
        
        ' --- BƯỚC B: TÍNH MÃ VAI TRÒ (CỘT H) TỪ SALES POSITION & LOẠI TÀI KHOẢN (CỘT A) ---
        Dim roleVal As String
        Dim rRole As Long
        Dim foundRole As Boolean: foundRole = False
        
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
            cell.Offset(0, 6).Value = roleVal
        Else
            cell.Offset(0, 6).Value = "Null"
            roleVal = ""
        End If
        
        ' Điền Cột A: Loại tài khoản (Nếu vai trò chứa role S thì eSalesSFA, ngược lại eSalesBackOffice)
        If HasRoleS(roleVal) Then
            cell.Offset(0, -1).Value = "eSalesSFA"
        Else
            cell.Offset(0, -1).Value = "eSalesBackOffice"
        End If
        
        ' --- BƯỚC C: TÍNH MÃ KHO (CỘT I) & MÃ ĐỘI (CỘT J) TỪ SALES TEAM NAME ---
        Dim resMaDoi As String, resMaKho As String
        Dim arrInput() As String
        Dim i As Long, r As Long
        Dim searchStr As String, dbStr As String
        Dim foundKho As Boolean
        
        resMaDoi = ""
        resMaKho = ""
        
        If salesTeam <> "" And (Not Not arrDoiNorm) Then
            ' Tách nhiều đội nếu có (hỗ trợ dấu phẩy, dấu chấm phẩy, xuống dòng)
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
                    
                    For r = 1 To UBound(arrDoiNorm)
                        dbStr = arrDoiNorm(r)
                        If dbStr <> "" And dbStr = searchStr Then
                            Dim strMaDoi As String, strMaKho As String
                            strMaDoi = Trim(CStr(arrDoi(r, 2))) ' Cột B: Mã đội
                            strMaKho = Trim(CStr(arrDoi(r, 4))) ' Cột D: Mã kho
                            
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
        
        ' Cắt dấu phẩy thừa ở cuối chuỗi
        If Len(resMaDoi) > 0 Then resMaDoi = Left(resMaDoi, Len(resMaDoi) - 2)
        If Len(resMaKho) > 0 Then
            resMaKho = Left(resMaKho, Len(resMaKho) - 2)
            resMaKho = LocTrungMaKho(resMaKho)
        End If
        
        If resMaKho = "" Then resMaKho = "Null"
        If resMaDoi = "" Then resMaDoi = "Null"
        
        ' Đặt format Text để giữ nguyên số 0 ở đầu các mã kho/đội
        cell.Offset(0, 7).NumberFormat = "@"
        cell.Offset(0, 8).NumberFormat = "@"
        
        cell.Offset(0, 7).Value = resMaKho
        cell.Offset(0, 8).Value = resMaDoi

NextCell:
    Next cell

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Có lỗi xảy ra trong quá trình xử lý: " & Err.Description, vbExclamation, "Thông báo lỗi"
End Sub

'============================================================
' CÁC HÀM PHỤ TRỢ DÙNG ĐỂ LỌC TRÙNG VÀ CHUẨN HÓA
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
        If itemK <> "" Then
            col.Add itemK, UCase(itemK)
        End If
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
    
    ' 1. Loại bỏ các ký tự điều khiển, khoảng trắng đặc biệt và ký tự ẩn
    result = Replace(result, Chr(160), " ")   ' Non-breaking space
    result = Replace(result, Chr(10), "")     ' Line Feed
    result = Replace(result, Chr(13), "")     ' Carriage Return
    result = Replace(result, Chr(9), " ")     ' Tab
    result = Replace(result, ChrW(8203), "")  ' Zero-width space
    result = Replace(result, ChrW(8204), "")  ' Zero-width non-joiner
    result = Replace(result, ChrW(8205), "")  ' Zero-width joiner
    result = Replace(result, ChrW(65279), "") ' Zero-width byte order mark (BOM)
    
    ' 2. Đồng nhất các loại dấu gạch ngang (En-dash, Em-dash, Hyphen...) về dấu gạch ngang chuẩn (-)
    result = Replace(result, ChrW(8211), "-") ' En Dash (–)
    result = Replace(result, ChrW(8212), "-") ' Em Dash (—)
    result = Replace(result, ChrW(8722), "-") ' Minus Sign (−)
    result = Replace(result, ChrW(45), "-")   ' Standard Hyphen (-)
    
    ' 3. Loại bỏ dấu tổ hợp (Decomposed Unicode NFD Combining Diacritical Marks)
    Dim combMarks As Variant
    combMarks = Array(768, 769, 770, 771, 772, 774, 775, 776, 777, 778, 779, 780, 781, 782, 783, 784, 785, 786, 787, 788, 789, 790, 791, 792, 793, 794, 795, 803, 804, 805, 821)
    Dim idx As Long
    For idx = LBound(combMarks) To UBound(combMarks)
        result = Replace(result, ChrW(combMarks(idx)), "")
    Next idx
    
    ' 4. Bỏ dấu tiếng Việt dạng Unicode chuẩn (NFC)
    result = BoDau(result)
    
    ' 5. Đồng nhất viết hoa và dọn dẹp khoảng trắng
    result = UCase(Trim(result))
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ' Chuẩn hóa khoảng cách quanh dấu gạch ngang (TT - BA RIA -> TT-BA RIA)
    result = Replace(result, " - ", "-")
    result = Replace(result, " -", "-")
    result = Replace(result, "- ", "-")
    
    ChuanHoaTenDoi = Trim(result)
End Function

Function BoDau(ByVal text As String) As String
    Dim uniChars As String, unsignChars As String, i As Long
    
    ' Sử dụng ChrW để thiết lập danh sách ký tự Tiếng Việt độc lập với bảng mã code page hệ thống
    ' a/A
    text = ReplaceChars(text, Array(224, 225, 7843, 227, 7841, 259, 7857, 7855, 7859, 7861, 7863, 226, 7847, 7845, 7849, 7851, 7853, 192, 193, 7842, 195, 7840, 258, 7856, 7854, 7858, 7860, 7862, 194, 7846, 7844, 7848, 7850, 7852), "A")
    ' e/E
    text = ReplaceChars(text, Array(232, 233, 7867, 7869, 7865, 234, 7873, 7871, 7875, 7877, 7879, 200, 201, 7866, 7868, 7864, 202, 7872, 7870, 7874, 7876, 7878), "E")
    ' i/I
    text = ReplaceChars(text, Array(236, 237, 7881, 297, 7883, 204, 205, 7880, 296, 7882), "I")
    ' o/O
    text = ReplaceChars(text, Array(242, 243, 7887, 245, 7885, 244, 7891, 7889, 7893, 7895, 7897, 417, 7899, 7901, 7903, 7905, 7907, 210, 211, 7886, 213, 7884, 212, 7890, 7888, 7892, 7894, 7896, 416, 7898, 7900, 7902, 7904, 7906), "O")
    ' u/U
    text = ReplaceChars(text, Array(249, 250, 7911, 361, 7913, 432, 7915, 7917, 7919, 7921, 7923, 217, 218, 7910, 360, 7912, 431, 7914, 7916, 7918, 7920, 7922), "U")
    ' y/Y
    text = ReplaceChars(text, Array(7923, 253, 7927, 7929, 7925, 7922, 221, 7926, 7928, 7924), "Y")
    ' d/D (đ, Đ)
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

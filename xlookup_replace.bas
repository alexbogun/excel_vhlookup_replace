' function to replace all VLOOKUPS / HLOOKUPS in current selection
Sub Replace_lookups_in_selection()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Call Replace_lookups_in_range(Selection)
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub



' function to replace all VLOOKUPS / HLOOKUPS in the whole workbook
Sub Replace_all_lookups()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim Sh As Worksheet
    Dim i As Integer, n As Integer
    
    n = ThisWorkbook.Worksheets.Count
    i = 1
    For Each Sh In ThisWorkbook.Worksheets
        Call Replace_lookups_in_range(Sh.UsedRange)
        Application.StatusBar = "Working on sheet " & i & " out of & n"
    Next Sh
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub



' function to replace all VLOOKUPS / HLOOKUPS in a given range
Sub Replace_lookups_in_range(R As Range)
    Dim s1 As String, s2 As String, FirstFind As String
    Dim Loc As Range

    With R
        Set Loc = .Cells.Find(What:="LOOKUP")
        If Not Loc Is Nothing Then
            FirstFind = Loc.Address
            Do
                s1 = Loc.Formula2R1C1
                s2 = String_replace_lookup(s1)
                If s2 <> s1 Then
                    Loc.FormulaR1C1 = s2
                End If
                Set Loc = .FindNext(Loc)
            Loop While Not Loc Is Nothing And Loc.Address <> FirstFind
        End If
    End With
End Sub




' help function to convert index numbers to r1c1 address
Function r1c1(ByVal r1, ByVal r2, ByVal c1, ByVal c2, r1_r, r2_r, c1_r, c2_r) 'first 4 args - position, last 4 relative/absolute address
    Dim sep As String
    sep = ":"
    'handling of relative address
    If r1_r Then r1 = "[" & r1 & "]"
    If r2_r Then r2 = "[" & r2 & "]"
    If c1_r Then c1 = "[" & c1 & "]"
    If c2_r Then c2 = "[" & c2 & "]"
    
    'handling of whole row/columns selection
    If r1 = 0 Then
        r1 = ""
    Else
        r1 = "R" & r1
    End If
    If r2 = 0 Then
        r2 = ""
    Else
        r2 = "R" & r2
    End If
    If c1 = 0 Then
        c1 = ""
    Else
        c1 = "C" & c1
    End If
    If c2 = 0 Then
        c2 = ""
    Else
        c2 = "C" & c2
    End If
    
    'handling of whole row/columns being in format "R1" "R[1]" instead of "R1R1" or "R[1]R[1]"
    If r1 = "" And c1 = c2 Then
        c2 = ""
        sep = ""
    End If
    If c1 = "" And r1 = r2 Then
        r2 = ""
        sep = ""
    End If
    
    'combining result
    r1c1 = r1 & c1 & sep & r2 & c2
End Function




' function that does actual parsing of the formula string and replacement
Function String_replace_lookup(ByVal s As String) As String
    Dim arg(4) As String
    Dim i As Integer, j As Integer, n_par As Integer
    Dim arg_n As Integer, found As Integer, first As Integer, last As Integer, from As Integer, m1 As Integer, m2 As Integer
    Dim look As String, s_vl As String, s_er As String, s_hl As String, iferr As String, mtch As String, ch As String
    Dim part0 As String, part1 As String, part2 As String, before As String, after As String
    Dim masked As Boolean, r1_r As Boolean, r2_r As Boolean, c1_r As Boolean, c2_r As Boolean
    Dim splt As Variant, splt1 As Variant, splt2 As Variant, r1 As Variant, r2 As Variant, c1 As Variant, c2 As Variant, col As Variant
    Dim rName As Name
    ' when I write vlookup it also applies to hlookup
    
    arg(4) = "1"    ' default value of 4th argument for vlookup
    arg_n = 1
    s_vl = "VLOOKUP("
    s_hl = "HLOOKUP("
    s_er = "IFERROR("
    part0 = "" 'other sheet reference
    iferr = "" 'iferror argument
    mtch = ""  'match type argument
    r1_r = False    'relative address?
    r2_r = False
    c1_r = False
    c2_r = False
    
    n_par = 1       ' number of parenthesis encountered
    found = InStr(s, s_vl) ' found vlookup
    If found Then   'vlookup
        look = "v"
    Else            'hlookup
        found = InStr(s, s_hl)
        If found Then
            s_vl = s_hl
            look = "h"
        Else
            String_replace_lookup = s
            Exit Function
        End If
    End If
        
    first = found   'found vlookup in position
    last = found    'end of the body of vlookup (will be populated later)
    from = found + Len(s_vl) 'start of current vlookup argument
    masked = False 'variable that handles ignoring quoted substrings e.g. vlookup(")",A:B,2) will correctly skip ")"
    
    ' extract arguments from vlookup
    If found > 0 Then
        Do ' so that we can exit do
            For i = (found + Len(s_vl)) To Len(s)   'Loop through chars of the string starting with body of vlookup
                ch = Mid(s, i, 1) ' current char
                If ch = """" Then ' if found quote start masking
                    masked = Not (masked)
                End If

                If Not masked Then ' ignore anything inside quotes
                
                    If ch = "(" Then 'found nested function, increase number of parenthesis by 1
                        n_par = n_par + 1
                    End If
                    
                    If ch = ")" Then 'closure of (nested) function, decrease number of parenthesis by 1
                        n_par = n_par - 1
                    End If
                    
                    If (n_par = 1 And ch = ",") Then 'if we are inside current vlookup function and found coma -> next argument
                        arg(arg_n) = Mid(s, from, i - from) 'extract previous argument
                        from = i + 1
                        arg_n = arg_n + 1
                    End If
                    
                    If (n_par = 0) Then 'if closed last parenthesis -> closure of vlookup
                        arg(arg_n) = Mid(s, from, i - from)
                        last = i
                        Exit For
                    End If
                End If
            Next i
            
            If InStr(arg(2), ":") = 0 Then 'possibly named range?
                For Each rName In Names
                    If InStr(arg(2), rName.Name) Then 'look in named ranges and replace
                        arg(2) = Replace(arg(2), rName.Name, Names(rName.Name).RefersToRange.Address(ReferenceStyle:=xlR1C1))
                    End If
                Next rName
            End If
            
            ' cleanup RC expressions
            arg(2) = Replace(arg(2), "RC", "R[0]C") 'if no number after R -> relative reference current row
            arg(2) = Replace(arg(2), "C:", "C[0]:") 'same for cols first part
            If Right(arg(2), 2) = "]C" Then 'second part
                arg(2) = arg(2) & "[0]"
            End If
            
            'split reference in part before and after semicolon
            splt = Split(arg(2), ":")
            If UBound(splt) <> 1 Then 'if not exactly 2 parts abort
                Exit Do
            End If
            
            part1 = splt(0)
            part2 = splt(1)
            
            'handle references to whole rows / columns
            If InStr(part1, "R") = 0 Then
                part1 = Replace(part1, "C", "R0C")
            End If
            If InStr(part2, "R") = 0 Then
                part2 = Replace(part2, "C", "R0C")
            End If
            If InStr(part1, "C") = 0 Then
                part1 = part1 & "C0"
            End If
            If InStr(part2, "C") = 0 Then
                part2 = part2 & "C0"
            End If
            
            'handle references to other sheets
            splt = Split(part1, "!")
            If UBound(splt) = 1 Then
                part0 = splt(0) & "!"
                part1 = splt(1)
            End If
            
            'extract numbers of rows and columns
            part1 = Replace(part1, "R", "")
            part2 = Replace(part2, "R", "")
            splt1 = Split(part1, "C")
            splt2 = Split(part2, "C")
            
            If (UBound(splt1) <> 1) Or (UBound(splt2) <> 1) Then
                Exit Do
            End If
                        
            before = Left(s, first - 1) 'part before vlookup
            after = Mid(s, last + 1, Len(s)) 'part after vlookup
            r1 = Trim(splt1(0)) 'extract numbers from r1c1:r2c2
            c1 = Trim(splt1(1))
            r2 = Trim(splt2(0))
            c2 = Trim(splt2(1))
            col = Trim(arg(3)) 'column index of (3rd arg of vlookup)
            
            'handling match type argument
            If ((arg(4) = "") Or (arg(4) = "TRUE") Or (arg(4) = "1")) Then
                mtch = ",-1"
            End If
            
            'handling relative addresses
            If InStr(r1, "[") Then
                r1 = Replace(Replace(r1, "[", ""), "]", "")
                r1_r = True
            End If
            If InStr(r2, "[") Then
                r2 = Replace(Replace(r2, "[", ""), "]", "")
                r2_r = True
            End If
            If InStr(c1, "[") Then
                c1 = Replace(Replace(c1, "[", ""), "]", "")
                c1_r = True
            End If
            If InStr(c2, "[") Then
                c2 = Replace(Replace(c2, "[", ""), "]", "")
                c2_r = True
            End If
            
            'check if all valid numbers
            If Not (IsNumeric(r1) And IsNumeric(r2) And IsNumeric(c1) And IsNumeric(c2) And IsNumeric(col)) Then
                Exit Do
            End If

            'look for "iferror" before vlookup
            If InStr(Right(before, 12), s_er) > 0 Then
                m1 = InStr(after, ",") + 1
                m2 = InStr(after, ")")
                iferr = Mid(after, m1, m2 - m1) 'iferror argument
                before = Left(before, Len(before) - Len(Right(before, 12))) & Left(Right(before, 12), InStr(Right(before, 12), s_er) - 1)
                after = Mid(after, m2 + 1, Len(after))
            End If
            
            If (Len(iferr) + Len(mtch) > 0) Then
                iferr = "," + iferr
            End If
            
            'actual replacement of the function
            If look = "v" Then
                col = CStr(CInt(c1) + CInt(col) - 1)
                s = before & "XLOOKUP(" & arg(1) & "," & part0 & r1c1(r1, r2, c1, c1, r1_r, r2_r, c1_r, c1_r) & "," & part0 & r1c1(r1, r2, col, col, r1_r, r2_r, c1_r, c1_r) & iferr & mtch & ")" & after
            Else
                col = CStr(CInt(r1) + CInt(col) - 1)
                s = before & "XLOOKUP(" & arg(1) & "," & part0 & r1c1(r1, r1, c1, c2, r1_r, r1_r, c1_r, c2_r) & "," & part0 & r1c1(col, col, c1, c2, r1_r, r1_r, c1_r, c2_r) & iferr & mtch & ")" & after
            End If
            
            s = String_replace_lookup(s) 'recursive call to handle multiple v/hlookups in the same formula
            
        Exit Do
        Loop
    End If
        
    String_replace_lookup = s
End Function

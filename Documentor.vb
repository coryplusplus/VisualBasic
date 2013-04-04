'Author:Cory Kelly
'Code used to generate documentation given a specific xml file
'Listed here as a reference and not intended for use.


Dim stand_type_map(34, 2)
Dim members() As String
Dim member_type As String
Dim Size As Integer
Dim message_name As String
Dim directory_name1 As String
Dim directory_name2 As String
Dim header_file As String
Dim class_name As String

Dim row_index As Integer

Dim message_row As Integer
Dim message_column As Integer

Dim common_struct_row As Integer
Dim common_struct_column As Integer

Dim enums_row As Integer
Dim enums_column As Integer

Dim class_directory_range As Range
Dim class_type_range As Range
Dim class_name_range As Range
Dim member_name_range As Range
Dim member_type_range As Range
Dim member_upper_range As Range
Dim member_lower_range As Range
Dim standard_type_id_range As Range
Dim messages As Worksheet
Dim commonstructs As Worksheet
Dim enums As Worksheet


Sub init()
Set class_directory_range = Range("W1:W14471")
Set class_type_range = Range("AD1:AD14471")
Set class_name_range = Range("AE1:AE14471")
Set member_name_range = Range("AR1:AR14471")
Set member_type_range = Range("AX1:AX14471")
Set member_upper_range = Range("AZ1:AZ14471")
Set member_lower_range = Range("AY1:AY14471")
Set standard_type_id_range = Range("HR14471:HR14507")
Set messages = Worksheets("Messages")
Set commonstructs = Worksheets("Common_Structures")
Set enums = Worksheets("Enums")

row_index = 2
message_row = 1
message_column = 1

common_struct_row = 1
common_struct_column = 1

enums_row = 1
enums_column = 1

i = 1
Do
    stand_type_map(i, 1) = standard_type_id_range.Cells(i)
    i = i + 1
    Loop While i < 35

i = 1
Do
    stand_type_map(i, 2) = standard_type_id_range.Offset(0, 1).Cells(i)
    i = i + 1
    Loop While i < 35


End Sub

Sub init2()
'Will have to change after set header
Set class_directory_range = Range("W1:W14471")
'Will need to check if this is empty, so it can stay the same
Set class_type_range = Range("AD1:AD14471")
'Set class_name_range = Range("CI1:CI14471")
Set member_name_range = Range("CV1:CV14471")
Set member_type_range = Range("DB1:DB14471")
Set member_upper_range = Range("DD1:DD14471")
Set member_lower_range = Range("DC1:DC14471")

row_index = 2



End Sub

Sub set_class_name()

class_name = "core::" & directory_name1 & "::" & message_name


End Sub

Sub set_header_file()

header_file = "core/" & directory_name2 & "/" & message_name & ".h"

End Sub

Sub get_members_info()
num_members = 0
members_index = row_index
Do
    If member_name_range.Cells(row_index) <> "" Then
        num_members = num_members + 1
    End If
    row_index = row_index + 1
    Loop While class_name_range.Cells(row_index) = message_name
    
If num_members > 0 Then
ReDim members(1 To num_members, 3)
    i = 1
    Do
        members(i, 1) = member_name_range.Cells(members_index)
        set_member_container members_index, i
        members(i, 3) = member_type_range.Cells(members_index)
                 
        i = i + 1
        members_index = members_index + 1
        Loop While i < num_members + 1
    
    set_member_type
Else
    ReDim members(1, 3)
    members(1, 1) = ""
    members(1, 2) = ""
    members(1, 3) = ""
    
    
End If

Size = 0
Size = UBound(members, 1) - LBound(members, 1) + 1

End Sub
Sub set_member_container(x, y)
    
    
    If member_upper_range.Cells(x) = 1 And member_lower_range.Cells(x) = 1 Then
        members(y, 2) = "none"
    End If

        
    If member_upper_range.Cells(x) = 1 And member_lower_range.Cells(x) = 0 Then
        members(y, 2) = "vector"
    End If

    
    If member_upper_range.Cells(x) <> 1 And member_upper_range.Cells(x) <> -1 Then
        members(y, 2) = "vector"
    End If
        
    If member_upper_range.Cells(x) = -1 Then
        members(y, 2) = "list"
    End If




End Sub

Sub set_member_type()

Size = 0
Size = UBound(members, 1) - LBound(members, 1) + 1

i = 1
Do
    mem_type = members(i, 3)
    If Left(mem_type, 1) = "G" Then
        get_standard_type (i)
    Else
        get_nonstandard_type (i)
    End If
    i = i + 1
    Loop While i < Size + 1




End Sub

Sub get_standard_type(x)

member_type = members(x, 3)
i = 1
b = 0
Do
    If stand_type_map(i, 1) = member_type Then
        members(x, 3) = stand_type_map(i, 2)
        b = 1
    End If
    i = i + 1
    Loop While i < 35 And b = 0
    
End Sub

Sub get_nonstandard_type(x)

member_type = members(x, 3)
i = 1
b = 0
Do
    If class_type_range.Cells(i) = member_type Then
        members(x, 3) = class_type_range.Offset(0, 1).Cells(i)
        b = 1
    End If
    i = i + 1
    Loop While i < 14471 And b = 0
    
i = 1
Set class_type_range = Range("CH1:CH14471")
Do
    If class_type_range.Cells(i) = member_type Then
        members(x, 3) = class_type_range.Offset(0, 1).Cells(i)
        b = 1
    End If
    i = i + 1
    Loop While i < 14471 And b = 0
  
Set class_type_range = Range("AD1:AD14471")
End Sub
Sub PopulateCommonStructure()

start_row = common_struct_row
Range(commonstructs.Cells(common_struct_row, common_struct_column), commonstructs.Cells(common_struct_row, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
Range(commonstructs.Cells(common_struct_row, common_struct_column), commonstructs.Cells(common_struct_row, 7)).Borders(xlEdgeTop).Weight = xlThick


commonstructs.Cells(common_struct_row, common_struct_column).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column) = "Structure Name:"
commonstructs.Cells(common_struct_row, common_struct_column + 1) = message_name
common_struct_row = common_struct_row + 1
commonstructs.Cells(common_struct_row, common_struct_column).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column) = "Purpose: "
common_struct_row = common_struct_row + 1
commonstructs.Cells(common_struct_row, common_struct_column).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column) = "Comments: "
common_struct_row = common_struct_row + 1
commonstructs.Cells(common_struct_row, common_struct_column).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column) = "Fully Qualified Class Name: "
commonstructs.Cells(common_struct_row, common_struct_column + 1) = class_name
common_struct_row = common_struct_row + 1
commonstructs.Cells(common_struct_row, common_struct_column).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column) = "Header File: "
commonstructs.Cells(common_struct_row, common_struct_column + 1) = header_file
common_struct_row = common_struct_row + 1
commonstructs.Cells(common_struct_row, common_struct_column + 1).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column + 1) = "Name"
commonstructs.Cells(common_struct_row, common_struct_column + 2).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column + 2) = "Type"
commonstructs.Cells(common_struct_row, common_struct_column + 3).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column + 3) = "Method"
commonstructs.Cells(common_struct_row, common_struct_column + 4).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column + 4) = "Units"
commonstructs.Cells(common_struct_row, common_struct_column + 5).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column + 5) = "Min/Max Range"
commonstructs.Cells(common_struct_row, common_struct_column + 6).Font.Bold = True
commonstructs.Cells(common_struct_row, common_struct_column + 6) = "Elaboration"
common_struct_row = common_struct_row + 1

i = 1
If members(i, 1) <> "" Then
Do
    temp_string = ""
    commonstructs.Cells(common_struct_row, common_struct_column + 1) = "m_" & members(i, 1)
    
    If members(i, 2) = "none" Then
        commonstructs.Cells(common_struct_row, common_struct_column + 2) = members(i, 3)
        temp_string = commonstructs.Cells(common_struct_row, common_struct_column + 2)
    ElseIf members(i, 2) = "vector" Then
        commonstructs.Cells(common_struct_row, common_struct_column + 2) = "std::vector<" & members(i, 3) & ">"
        temp_string = commonstructs.Cells(common_struct_row, common_struct_column + 2)
    ElseIf members(i, 2) = "list" Then
        commonstructs.Cells(common_struct_row, common_struct_column + 2) = "std::list<" & members(i, 3) & ">"
        temp_string = commonstructs.Cells(common_struct_row, common_struct_column + 2)
    Else
        commonstructs.Cells(common_struct_row, common_struct_column + 2) = members(i, 2) & members(i, 3)
    End If
    
    commonstructs.Cells(common_struct_row, common_struct_column + 3) = "'+ " & temp_string & " " & members(i, 1) & " ( ) const" & vbCrLf & "+ void " & members(i, 1) & "(" & temp_string & ")"

    common_struct_row = common_struct_row + 1
    i = i + 1
    Loop While i < Size + 1
Else
    commonstructs.Cells(common_struct_row, common_struct_column + 1) = "None"
    commonstructs.Cells(common_struct_row, common_struct_column + 2) = "None"
    commonstructs.Cells(common_struct_row, common_struct_column + 3) = "None"
    common_struct_row = common_struct_row + 1
    
End If
    
    
Range(commonstructs.Cells(start_row, 1), commonstructs.Cells(common_struct_row - 1, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(commonstructs.Cells(start_row, 1), commonstructs.Cells(common_struct_row - 1, 1)).Borders(xlEdgeLeft).Weight = xlThick

Range(commonstructs.Cells(start_row, 7), commonstructs.Cells(common_struct_row - 1, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
Range(commonstructs.Cells(start_row, 7), commonstructs.Cells(common_struct_row - 1, 7)).Borders(xlEdgeRight).Weight = xlThick

Range(commonstructs.Cells(common_struct_row - 1, 1), commonstructs.Cells(common_struct_row - 1, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(commonstructs.Cells(common_struct_row - 1, 1), commonstructs.Cells(common_struct_row - 1, 7)).Borders(xlEdgeBottom).Weight = xlThick
    
     
common_struct_row = common_struct_row + 1

End Sub

Sub PopulateMessages()

start_row = message_row
Range(messages.Cells(message_row, message_column), messages.Cells(message_row, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
Range(messages.Cells(message_row, message_column), messages.Cells(message_row, 7)).Borders(xlEdgeTop).Weight = xlThick

messages.Cells(message_row, message_column).Font.Bold = True
messages.Cells(message_row, message_column) = "Message:"
messages.Cells(message_row, message_column + 1) = message_name
message_row = message_row + 1
messages.Cells(message_row, message_column).Font.Bold = True
messages.Cells(message_row, message_column) = "Purpose: "
message_row = message_row + 1
messages.Cells(message_row, message_column).Font.Bold = True
messages.Cells(message_row, message_column) = "Comments: "
message_row = message_row + 1
messages.Cells(message_row, message_column).Font.Bold = True
messages.Cells(message_row, message_column) = "Fully Qualified Class Name: "
messages.Cells(message_row, message_column + 1) = class_name
message_row = message_row + 1
messages.Cells(message_row, message_column).Font.Bold = True
messages.Cells(message_row, message_column) = "Header File: "
messages.Cells(message_row, message_column + 1) = header_file
message_row = message_row + 1
messages.Cells(message_row, message_column + 1).Font.Bold = True
messages.Cells(message_row, message_column + 1) = "Name"
messages.Cells(message_row, message_column + 2).Font.Bold = True
messages.Cells(message_row, message_column + 2) = "Type"
messages.Cells(message_row, message_column + 3).Font.Bold = True
messages.Cells(message_row, message_column + 3) = "Method"
messages.Cells(message_row, message_column + 4).Font.Bold = True
messages.Cells(message_row, message_column + 4) = "Units"
messages.Cells(message_row, message_column + 5).Font.Bold = True
messages.Cells(message_row, message_column + 5) = "Min/Max Range"
messages.Cells(message_row, message_column + 6).Font.Bold = True
messages.Cells(message_row, message_column + 6) = "Elaboration"
message_row = message_row + 1

i = 1
    If members(i, 1) <> "" Then
        Do
            temp_string = ""
            messages.Cells(message_row, message_column + 1) = "m_" & members(i, 1)
            
            If members(i, 2) = "none" Then
                messages.Cells(message_row, message_column + 2) = members(i, 3)
                temp_string = messages.Cells(message_row, message_column + 2)
            ElseIf members(i, 2) = "vector" Then
                messages.Cells(message_row, message_column + 2) = "std::vector<" & members(i, 3) & ">"
                temp_string = messages.Cells(message_row, message_column + 2)
            ElseIf members(i, 2) = "list" Then
                messages.Cells(message_row, message_column + 2) = "std::list<" & members(i, 3) & ">"
                temp_string = messages.Cells(message_row, message_column + 2)
            Else
                messages.Cells(message_row, message_column + 2) = members(i, 2) & members(i, 3)
            End If
            
            messages.Cells(message_row, message_column + 3) = "'+ " & temp_string & " " & members(i, 1) & " ( ) const" & vbCrLf & "+ void " & members(i, 1) & "(" & temp_string & ")"
        
            message_row = message_row + 1
            i = i + 1
        Loop While i < Size + 1
    Else
        messages.Cells(message_row, message_column + 1) = "None"
        messages.Cells(message_row, message_column + 2) = "None"
        messages.Cells(message_row, message_column + 3) = "None"
        message_row = message_row + 1
    
    End If
    
    
    
Range(messages.Cells(start_row, 1), messages.Cells(message_row - 1, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(messages.Cells(start_row, 1), messages.Cells(message_row - 1, 1)).Borders(xlEdgeLeft).Weight = xlThick

Range(messages.Cells(start_row, 7), messages.Cells(message_row - 1, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
Range(messages.Cells(start_row, 7), messages.Cells(message_row - 1, 7)).Borders(xlEdgeRight).Weight = xlThick

Range(messages.Cells(message_row - 1, 1), messages.Cells(message_row - 1, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(messages.Cells(message_row - 1, 1), messages.Cells(message_row - 1, 7)).Borders(xlEdgeBottom).Weight = xlThick

message_row = message_row + 1

End Sub
Sub PopulateEnumerations()


start_row = enums_row

Range(enums.Cells(enums_row, enums_column), enums.Cells(enums_row, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
Range(enums.Cells(enums_row, enums_column), enums.Cells(enums_row, 2)).Borders(xlEdgeTop).Weight = xlThick


enums.Cells(enums_row, enums_column).Font.Bold = True
enums.Cells(enums_row, enums_column) = "Enumeration:"
enums.Cells(enums_row, enums_column + 1) = message_name
enums_row = enums_row + 1
enums.Cells(enums_row, enums_column).Font.Bold = True
enums.Cells(enums_row, enums_column) = "Purpose: "
enums_row = enums_row + 1
enums.Cells(enums_row, enums_column).Font.Bold = True
enums.Cells(enums_row, enums_column) = "Comments: "
enums_row = enums_row + 1
enums.Cells(enums_row, enums_column).Font.Bold = True
enums.Cells(enums_row, enums_column) = "Fully Qualified Class Name: "
enums.Cells(enums_row, enums_column + 1) = class_name
enums_row = enums_row + 1
enums.Cells(enums_row, enums_column).Font.Bold = True
enums.Cells(enums_row, enums_column) = "Header File: "
enums.Cells(enums_row, enums_column + 1) = header_file
enums_row = enums_row + 1
enums.Cells(enums_row, enums_column + 1).Font.Bold = True
enums.Cells(enums_row, enums_column + 1) = "Name"

enums_row = enums_row + 1

i = 1
Do
    temp_string = ""
    enums.Cells(enums_row, enums_column + 1) = members(i, 1)
    enums_row = enums_row + 1
    i = i + 1
    Loop While i < Size + 1
    
Range(enums.Cells(start_row, 1), enums.Cells(enums_row - 1, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(enums.Cells(start_row, 1), enums.Cells(enums_row - 1, 1)).Borders(xlEdgeLeft).Weight = xlThick

Range(enums.Cells(start_row, 2), enums.Cells(enums_row - 1, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Range(enums.Cells(start_row, 2), enums.Cells(enums_row - 1, 2)).Borders(xlEdgeRight).Weight = xlThick

Range(enums.Cells(enums_row - 1, 1), enums.Cells(enums_row - 1, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(enums.Cells(enums_row - 1, 1), enums.Cells(enums_row - 1, 2)).Borders(xlEdgeBottom).Weight = xlThick
enums_row = enums_row + 1

End Sub

Sub LinkMessages()

Dim mtype_range As Range
Dim ctype_range As Range
Dim etype_range As Range
Dim temp_string As String

Set mtype_range = messages.Range("C1:C15000")
Set ctype_range = commonstructs.Range("B1:B15000")
Set etype_range = enums.Range("B1:B15000")

c_count = 1
e_count = 1


For Each x In mtype_range
    If x.Value <> "" And x.Value <> "Type" Then
    
    temp_string = x.Value
        'Remove std::vector< and std::list<
        If Left(x.Value, 6) = "std::v" Then
            temp_string = Mid(temp_string, 13, Len(temp_string) - 1)
            temp_string = Left(temp_string, Len(temp_string) - 1)
        ElseIf Left(x.Value, 6) = "std::l" Then
            temp_string = Mid(temp_string, 11, Len(temp_string) - 1)
            temp_string = Left(temp_string, Len(temp_string) - 1)
        End If
        
        
        ' Only link if the type is a common_structure or enumeration
    
        ' If common structure
        If Left(temp_string, 2) = "c_" Then
            For Each commonstructX In ctype_range
                If temp_string = commonstructX.Value Then
                    messages.Hyperlinks.Add Anchor:=x, Address:="", SubAddress:="Common_Structures!A" & c_count, TextToDisplay:=x.Value
                End If
            c_count = c_count + 1
            Next commonstructX
            c_count = 1
        
        ' If enumeration
        ElseIf Left(temp_string, 1) = "e" Then
            For Each enumX In etype_range
                If temp_string = enumX.Value Then
                    messages.Hyperlinks.Add Anchor:=x, Address:="", SubAddress:="Enums!A" & e_count, TextToDisplay:=x.Value
                End If
            e_count = e_count + 1
            Next enumX
            e_count = 1
        End If
        
        
    
    
    
    
    
    End If
    
    Next x
    




End Sub

Sub LinkCommonStructures()

Dim ctype_range As Range
Dim etype_range As Range
Dim temp_string As String

Set cctype_range = commonstructs.Range("C1:C15000")
Set ctype_range = commonstructs.Range("B1:B15000")
Set etype_range = enums.Range("B1:B15000")

c_count = 1
e_count = 1


For Each x In cctype_range
    If x.Value <> "" And x.Value <> "Type" Then
    
    temp_string = x.Value
        'Remove std::vector< and std::list<
        If Left(x.Value, 6) = "std::v" Then
            temp_string = Mid(temp_string, 13, Len(temp_string) - 1)
            temp_string = Left(temp_string, Len(temp_string) - 1)
        ElseIf Left(x.Value, 6) = "std::l" Then
            temp_string = Mid(temp_string, 11, Len(temp_string) - 1)
            temp_string = Left(temp_string, Len(temp_string) - 1)
        End If
        
        
        ' Only link if the type is a common_structure or enumeration
    
        ' If common structure
        If Left(temp_string, 2) = "c_" Then
            For Each commonstructX In ctype_range
                If temp_string = commonstructX.Value Then
                    commonstructs.Hyperlinks.Add Anchor:=x, Address:="", SubAddress:="Common_Structures!A" & c_count, TextToDisplay:=x.Value
                End If
            c_count = c_count + 1
            Next commonstructX
            c_count = 1
        
        ' If enumeration
        ElseIf Left(temp_string, 1) = "e" Then
            For Each enumX In etype_range
                If temp_string = enumX.Value Then
                    commonstructs.Hyperlinks.Add Anchor:=x, Address:="", SubAddress:="Enums!A" & e_count, TextToDisplay:=x.Value
                End If
            e_count = e_count + 1
            Next enumX
            e_count = 1
        End If
        
        
    
    
    
    
    
    End If
    
    Next x



End Sub


Sub main()

init
message_name = class_name_range.Cells(row_index)
Do
    If class_name_range.Cells(row_index) <> "" Then
        message_name = class_name_range.Cells(row_index)
        directory_name1 = class_directory_range.Cells(row_index)
        directory_name2 = class_directory_range.Cells(row_index)
        set_class_name
        set_header_file
        get_members_info
    
    
       If Left(message_name, 2) = "c_" Then
       PopulateCommonStructure
       ElseIf Left(message_name, 1) = "e" Then
       PopulateEnumerations
       ElseIf Left(message_name, 2) = "cd" Then
       PopulateMessages
       ElseIf Left(message_name, 2) <> "c_" And Left(message_name, 1) <> "e" And Left(message_name, 2) <> "cd" Then
       PopulateMessages
       End If
    Else
        row_index = row_index + 1
    End If
Loop While row_index < 14471

' Do the same thing for sub directories. If program crashes, delete the rest.
init2
Do
    If class_name_range.Cells(row_index) = "" Then
        Set class_name_range = Range("CI1:CI14471")
        If class_name_range.Cells(row_index) <> "" Then
            message_name = class_name_range.Cells(row_index)
            directory_name1 = class_directory_range.Cells(row_index)
            temp = directory_name1
            Set class_directory_range = Range("CA1:CA14471")
            directory_name1 = directory_name1 & "::" & class_directory_range.Cells(row_index)
            directory_name2 = temp & "/" & class_directory_range.Cells(row_index)
            set_class_name
            set_header_file
            Set class_name_range = Range("CI1:CI14471")
            get_members_info
        
        
           If Left(message_name, 2) = "c_" Then
           PopulateCommonStructure
           ElseIf Left(message_name, 1) = "e" Then
           PopulateEnumerations
           ElseIf Left(message_name, 2) = "cd" Then
           PopulateMessages
           ElseIf Left(message_name, 2) <> "c_" And Left(message_name, 1) <> "e" And Left(message_name, 2) <> "cd" Then
           PopulateMessages
           End If
           
           Set class_directory_range = Range("W1:W14471")
        Else
            row_index = row_index + 1
       End If
    Set class_name_range = Range("AE1:AE14471")
    Else
        row_index = row_index + 1
    End If
Loop While row_index < 14471
    
    
    
    
    
    LinkMessages
    LinkCommonStructures
    
    
    
End Sub


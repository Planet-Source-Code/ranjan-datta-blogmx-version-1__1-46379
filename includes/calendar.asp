<%
Dim file_name
Dim redirect_file
file_name="archive.asp"
redirect_file="archive.asp"
Function GetDaysInMonth(iMonth, iYear)
Dim dTemp
dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
GetDaysInMonth = Day(dTemp)
End Function
Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
Dim dTemp
dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
GetWeekdayMonthStartsOn = WeekDay(dTemp)
End Function
Function SubtractOneMonth(dDate)
SubtractOneMonth = DateAdd("m", -1, dDate)
End Function
Function AddOneMonth(dDate)
AddOneMonth = DateAdd("m", 1, dDate)
End Function
Dim dDate     ' Date we're displaying calendar for
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table
If IsDate(Request.QueryString("date")) Then
dDate = CDate(Request.QueryString("date"))
Else
If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
Else
dDate = Date()
If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
Response.Write "The date you picked was not a valid date.  The calendar was set to today's date."
End If
End If
End If
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
%>
<%
Dim rsDlogDate
Dim rsDlogDate_numRows

Set rsDlogDate = Server.CreateObject("ADODB.Recordset")
rsDlogDate.ActiveConnection = MM_dreamConn_STRING
rsDlogDate.Source = "SELECT BlogDate FROM tblBlog"
rsDlogDate.CursorType = 0
rsDlogDate.CursorLocation = 2
rsDlogDate.LockType = 1
rsDlogDate.Open()

rsDlogDate_numRows = 0
%>

<%
strHTML = "<table cellspacing=" & Chr(34) & "0" & Chr(34) & " id=" & Chr(34) & "calendar"
strHTML = strHTML & Chr(34) & " summary=" & Chr(34) & "Event Calendar" & Chr(34) & ">" & Chr(10) 
strHTML = strHTML & "<tr id=" & Chr(34) & "title" & Chr(34) & ">" & Chr(10)
strHTML = strHTML & "<th colspan=" & Chr(34) & "7" & Chr(34) & ">" & Chr(10)
strHTML = strHTML & "<a accesskey=" & Chr(34) & "p" & Chr(34) & " title=" & Chr(34) 
strHTML = strHTML & "Previous Month" & Chr(34) & " href=" & Chr(34) & file_name & "?date=" & SubtractOneMonth(dDate) & Chr(34) & ">"
strHTML = strHTML & "&laquo;</a>&nbsp;" & MonthName(Month(dDate)) & "," & Year(dDate) & "&nbsp;<a accesskey=" & Chr(34) & "n" & Chr(34) 
strHTML = strHTML & " title=" & Chr(34) & "Next Month" & Chr(34) 
strHTML = strHTML & " href=" & Chr(34) & "./" & file_name & "?date=" & AddOneMonth(dDate) & Chr(34) & ">&raquo;</a>" & Chr(10) & "</th>"  
strHTML = strHTML & Chr(10) & "</tr>" & Chr(10) & "<tr id=" & Chr(34) & "days" & Chr(34) & ">" & Chr(10)
Response.Write(strHTML)
function DayName (iDay)
select case iDay
case 1
	DayName = "Sun"
case 2 
    DayName = "Mon"
case 3 
    DayName = "Tue"
case 4 
    DayName = "Wed"
case 5 
    DayName = "Thu"
case 6 
    DayName = "Fri"
case 7 
    DayName = "Sat"
end select
end function
For gDay = 1 To 7
Response.Write("<th>" & DayName(gDay)& "</th>" & Chr(10))
Next 
Response.Write("</tr>" & Chr(10))
If iDOW <> 1 Then
Response.Write("<tr>" & Chr(10))
iPosition = 1
Do While iPosition < iDOW
Response.Write("<td class=" & Chr(34) & "nodays" & Chr(34) & ">&nbsp;</td>" & Chr(10))
iPosition = iPosition + 1
Loop
End If
iCurrent = 1
iPosition = iDOW
Do While iCurrent <= iDIM
isEvent = FALSE
'Get the events for this date only
rsDlogDate.filter = 0
Dim iCheck
Dim chkStr
chkStr = (rsDlogDate.Fields.Item("BlogDate").Name)
iCheck = Month(dDate) & "/" & iCurrent & "/" & Year(dDate)
rsDlogDate.filter = chkStr & "='" & (iCheck) & "'"
'If there are events then set the event flag to true
if not(rsDlogDate.EOF) then isEvent = TRUE
If iPosition = 1 Then
Response.Write("<tr>"  & Chr(10))
End If
If isEvent = TRUE Then
Response.Write("<td class=" & Chr(34) & "thedays" & Chr(34) & ">" & Chr(10))
tmpHTML = "<a href=" & Chr(34) & redirect_file & "?month=" 
tmpHTML = tmpHTML & Month(dDate) & "&amp;day=" & iCurrent & "&amp;year=" & Year(dDate) & Chr(34) & "><strong>" 
tmpHTML = tmpHTML & iCurrent & "</strong></a>" & Chr(10) & "</td>" & Chr(10)
Response.Write(tmpHTML)
Else
Response.Write("<td>" & iCurrent & "</td>" & Chr(10))
End If
If iPosition = 7 Then
Response.Write("</tr>" & Chr(10))
iPosition = 0
End If
iCurrent = iCurrent + 1
iPosition = iPosition + 1
Loop
If iPosition <> 1 Then
Do While iPosition <= 7
Response.Write("<td class=" & Chr(34) & "nodays" & Chr(34) & ">&nbsp;</td>" & Chr(10))
iPosition = iPosition + 1
Loop
Response.Write("</tr>" & Chr(10))
End If
Response.Write("</table>")
%>
<%
rsDlogDate.Close()
Set rsDlogDate = Nothing
%>

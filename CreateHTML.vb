Option Compare Database

'======================================================================================================================================================================================================================
' Export table to HTML code
' Author : Savchenko Arthur
' Sub Tohtml(Strsql As String, Field1 As String, Field2 As String, Field3 As String, Field4 As String, Field5 As String, Field6 As String)
' Создание документа в НTML коде на основании таблица
'======================================================================================================================================================================================================================
Sub Tohtml()
Dim Ahtml As String   ' Аll Documents
Dim Hhtml As String   ' Head
Dim Bhtml As String   ' Body
Dim Mhtml As String   ' Table
Dim Fhtml As String   ' Footer
Dim Xhtml As String   ' XML File
Dim Shtml As String   ' Setting html
Dim Hfile As Long     ' Hfile
Dim Rs As Recordset   ' Recordset

Set Rs = CurrentDb().OpenRecordset("SELECT * FROM A_user ORDER BY FIO")

'Формирование заголовка документа и таблицы стилей
Shtml = "<HTML> " & vbCrLf _
      & "<HEAD> " & vbCrLf _
      & "<meta http-equiv='Content-Type' content='text/html; charset=windows-1251' />" & vbCrLf _
      & "<HEAD/> " & vbCrLf _
      & "<style> " & vbCrLf _
      & "html, table{font-family: calibri; font-size: 14px;}" & vbCrLf _
      & ".col1 {width: 20px; } " & vbCrLf _
      & ".coln {width: 60px; } " & vbCrLf _
      & "th    {background-color: #FBF7EE; } " & vbCrLf _
      & "</style> " & vbCrLf _
      & "<basefont face='calibri' color='#404040'  size='12px'   />" & vbCrLf

Hhtml = Shtml & "<h3>   Отчет о сотрудниках </h3> <br>" & vbCrLf _
              & "<span> Cотрудники нашеей организации</span> <br> " & vbCrLf _
              & "<table cellpadding='3' cellspacing='0' border='1px' border-collapse='collapse' >" _
              & "<col class='col1'> " _
              & "<col span='9 class='coln'> " & vbCrLf

' Оглавление таблицы
Mhtml = "<tr> <th> ID </th>   <th>Фамилия</th>  <th>Дивизион</th>     <th>Должность</th>   <th>Net Name</th>  <th>Tel</th> <th>E_mail</th>    <th>Mob Telefone </th></tr> " & vbCrLf

' Тело таблицы
Do While Not Rs.EOF()
   Mhtml = Mhtml & "<tr> <td> <a href='http://portal.winner.ua/Apps/Docflow/w_employees.aspx?id=" & Rs!ID & "' > " & Rs!ID & "</a> </td> <td>" & Rs!fio_full & "</td>  <td>" & Rs!Division_name & " </td> <td> " & Rs!Position_name & " </td> <td> " & Zamm(Rs!Serv) & " </td> <td> " & Zamm(Rs!Tel_vnutr) & " </td>  <td> " & Zamm(Rs!Emailwork) & " </td>  <td> " & Zamm(Rs!Tel_mob) & " </td> </tr> " & vbCrLf
   Rs.MoveNext
Loop


Mhtml = Mhtml & " </table>" & vbCrLf & "  <br> "
Fhtml = " </h4> Created File Savchenko Arthur </h4>" & vbCrLf _
      & " <!--[IF IE]> <span> IE <span> <![ENDIF]-->" & vbCrLf _
      & " </HTML>"

'Окончательное формирование документа
Ahtml = Hhtml & Mhtml & Fhtml

Hfile = 1
Open "C:\Employee.htm" For Output Access Write As Hfile
Print #Hfile, Ahtml
Close Hfile

End Sub

Function Zamm(Pole As Variant)

If IsNull(Pole) Or Pole = "" Or Len(Pole) = 0 Then
   Zamm = "&nbsp;"
Else
   Zamm = Pole
End If

End Function

Attribute VB_Name = "TestingCalOpts"
Function BuildHtmlBody()
    Dim OutMail As Outlook.MailItem

    Dim html, name, address, age, department


html = "   <!-- Microdata markup added by Google Structured Data Markup Helper. -->    " & vbCrLf
html = html & "   <html><head></head><body>   " & vbCrLf
html = html & "       <p> " & vbCrLf
html = html & "         Dear John, thanks for booking your Google I/O ticket with us. " & vbCrLf
html = html & "       </p>    " & vbCrLf
html = html & "           " & vbCrLf
html = html & "   <span itemscope itemtype=""http://schema.org/EventReservation""><p itemscope="""" itemtype=""http://schema.org/EventReservation"">  " & vbCrLf
html = html & "         BOOKING DETAILS<br/>  " & vbCrLf
html = html & "         Reservation number: <span itemprop=""reservationNumber"" itemprop=""reservationNumber"">IO12345</span><br/>   " & vbCrLf
html = html & "         Order for: <span itemprop=""underName"" itemscope itemtype=""http://schema.org/Person"" itemprop=""underName"" itemscope="""" itemtype=""http://schema.org/Person"">  " & vbCrLf
html = html & "           <span itemprop=""name"" itemprop=""name"">John Smith</span> " & vbCrLf
html = html & "         </span><br/>  " & vbCrLf
html = html & "         </p><div itemprop=""reservationFor"" itemscope itemtype=""http://schema.org/Event"" itemprop=""reservationFor"" itemscope="""" itemtype=""http://schema.org/Event"">  " & vbCrLf
html = html & "           Event: <span itemprop=""name"" itemprop=""name"">Google I/O 2013</span><br/>    " & vbCrLf
html = html & "           <time itemprop=""startDate"" datetime=""2013-05-15T08:30:00-08:00"">Start time:     " & vbCrLf
html = html & "   <span itemprop=""startDate"" content=""2013-05-15T08:00"">May 15th 2013 8:00am PST</span></time><br/>   " & vbCrLf
html = html & "           Venue:  " & vbCrLf
html = html & "   <span itemprop=""location"" itemscope itemtype=""http://schema.org/Place""> " & vbCrLf
html = html & "   <span itemprop=""address"" itemscope itemtype=""http://schema.org/PostalAddress""><span itemprop=""streetAddress"" itemprop=""location"" itemscope="""" itemtype=""http://schema.org/Place"">   " & vbCrLf
html = html & "             <span itemprop=""name"">Moscone Center</span> " & vbCrLf
html = html & "             <span itemprop=""address"" itemscope="""" itemtype=""http://schema.org/PostalAddress"">   " & vbCrLf
html = html & "               <span itemprop=""streetAddress"">800 Howard St.</span>, " & vbCrLf
html = html & "               <span itemprop=""addressLocality"">San Francisco</span>,    " & vbCrLf
html = html & "               <span itemprop=""addressRegion"">CA</span>, " & vbCrLf
html = html & "               <span itemprop=""postalCode"">94103</span>, " & vbCrLf
html = html & "               <span itemprop=""addressCountry"">US</span> " & vbCrLf
html = html & "             </span>   " & vbCrLf
html = html & "           </span></span></span>   " & vbCrLf
html = html & "         </div></span> " & vbCrLf
html = html & "       <p></p> " & vbCrLf
html = html & "       " & vbCrLf
html = html & "   </body></html>  " & vbCrLf
html = html & "       " & vbCrLf


    BuildHtmlBody = html

Set oOutlook = CreateObject("Outlook.Application")
Set OutMail = oOutlook.CreateItem(olMailItem)

With OutMail

     .BodyFormat = olFormatPlain
     .Body = html
'     .HTMLBody = html   ' this adds a bunch of html code from outlook
     .Display
End With
    
'    Debug.Print BuildHtmlBody
    
End Function

Sub MailURL()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    strbody = "<HTML><BODY>"
    strbody = strbody & "<A href=http://www.MrExcel.com>URL Text</A>"
    strbody = strbody & "</BODY></HTML>"
    On Error Resume Next
    With OutMail
        .To = "APerson@Somewhere.com"
        .Subject = "Testing URL"
        .HTMLBody = strbody
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub




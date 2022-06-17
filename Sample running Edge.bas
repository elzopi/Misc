'This is an example of how to use the classes
Sub runedge()

    'Start Browser
    Dim objBrowser As clsEdge
    Set objBrowser = New clsEdge
    Call objBrowser.start
    
    'Attach to any ("") or a specific page
    Call objBrowser.attach("")
    
    'navigate
    Call objBrowser.navigate("https://google.de")
    
    Call objBrowser.waitCompletion
    
    'evaluate javascript
    Call objBrowser.jsEval("alert(""hi"")")
    
    'fill search form (textbox is named q)
    Call objBrowser.jsEval("document.getElementsByName(""q"")[0].value=""automate edge vba""")
    
    'run search
    Call objBrowser.jsEval("document.getElementsByName(""q"")[0].form.submit()")
    
    'wait till search has finished
    Call objBrowser.waitCompletion
    

    'click on codeproject link
    Call objBrowser.jsEval("document.evaluate("".//h3[text()='Automate Chrome / Edge using VBA - CodeProject']"", document).iterateNext().click()")
    
    Call objBrowser.waitCompletion
    
    Dim strVotes As String
'if a javascript expression evaluates to a plain type it is passed back to VBA
    strVotes = objBrowser.jsEval("ctl00_RateArticle_VountCountHist.innerText")
    
    MsgBox ("finish! Vote count is " & strVotes)
    
    objBrowser.closeBrowser

    
End Sub
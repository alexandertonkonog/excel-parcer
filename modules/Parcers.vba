Function SendParser(url, site)
    Dim oHtml As New HTMLDocument, post As Object, price As String, name As String
    Dim xmlhttp As Object
    Dim result(3)
    
    Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
    
    xmlhttp.Open "GET", url, False
    xmlhttp.send
    
    If xmlhttp.Status <> 200 Then
        MsgBox ("Страница " & url & " не отвечает")
        Exit Function
    End If
    
    oHtml.body.innerHTML = xmlhttp.responseText
    
    If site = "White Goods" Then
        name = oHtml.getElementsByClassName("page-inner__title").Item(0).innerText
        price = oHtml.getElementsByClassName("product-page__price-new").Item(0).getElementsByTagName("span").Item(0).innerText
    End If
    
    If price = "" Or name = "" Then
        MsgBox ("Данные не обработаны, возможно верстка сайта изменилась")
        Exit Function
    End If
    
    result(0) = name
    result(1) = price
    result(2) = site
    result(3) = url
    
    SendParser = result()
End Function

Sub ShowForm()
    UF1.Show
End Sub
# VBA の HTMLDocument

- htmlDoc への html ソースコードの書き込み

  ```vb
  Dim iHtmlDoc As MSHTML.IHTMLDocument, htmlDoc As MSHTML.HTMLDocument
  Set iHtmlDoc = New MSHTML.HTMLDocument
  iHtmlDoc.write "HTMLソースコード"
  'iHtmlDoc.getElementsByTagName("html")(0).innerHTML = "ソースコード"
  Set htmlDoc = iHtmlDoc
  ```

- body への html ソースコードの書き込み

  ```vb
  Dim htmlDoc As MSHTML.HTMLDocument
  Set htmlDoc = New MSHTML.HTMLDocument
  htmlDoc.body.innerHTML = "HTMLソースコード"
  ```

- DocumentElement (the root element of the document)  
  例. htmlDoc.DocumentElement.outerHTML  
  https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms759095(v=vs.85)
- innerHTML, nodeValue, textContent, innerText  
  http://xahlee.info/js/js_textContent_innerHTML_innerText_nodeValue.html  
  https://www.w3schools.com/jsref/prop_node_textcontent.asp

Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
Set html = CreateObject("htmlfile")
Const local = "hu-hu"

url = "https://www.microsoft.com/" & local & "/software-download/windows8ISO"
With req
	.Open "GET", url, False
	.Send
	page = .ResponseText
End With

With html
	.write "<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">"
	.write page
	.close
	Do While .readyState = "loading"
		WScript.Sleep 100
	Loop
	sessId = .getElementById("session-id").value
	segment = .getElementById("SoftwareDownload_LanguageSelectionByProductEdition").getAttribute("data-host-segments")
	langPageId = .getElementById("SoftwareDownload_LanguageSelectionByProductEdition").getAttribute("data-defaultPageId")
	downPageId = .getElementById("SoftwareDownload_DownloadLinks").getAttribute("data-defaultPageId")
	prodId = .getElementById("product-edition").options.Item(1).value 'choose product
	For Each opt In .getElementById("product-edition").options
		If opt.value <> Empty Then
			'WScript.Echo opt.value & vbTab & opt.text
		End If
	Next
End With

base = "https://www.microsoft.com/" & local & "/api/controls/contentinclude/html?host=www.microsoft.com&segments=" & segment & "&sessionId=" & sessId & "&pageId="
url = base & langPageId & "&productEditionId=" & prodId
With req
	.Open "GET", url, False
	.Send
	page = .ResponseText
End With

With html
	.clear
	.write page
	.close
	Do While .readyState = "loading"
		WScript.Sleep 100
	Loop
	For Each opt In .getElementById("product-languages").options
		If opt.value <> Empty Then
			json = Split(opt.value, """")
			'WScript.Echo json(3) & vbTab & json(7) & vbTab & opt.text
		End If
	Next
	json = Split(.getElementById("product-languages").options.Item(22).value, """") 'choose language
End With

url = base & downPageId & "&skuId=" & json(3) & "&language=" & json(7)
With req
	.Open "GET", url, False
	.Send
	page = .ResponseText
End With

With html
	.clear
	.write page
	.close
	Do While .readyState = "loading"
		WScript.Sleep 100
	Loop
	If .getElementsByClassName("button").length > 0 Then
		For Each link In .getElementsByClassName("button")
			WScript.Echo link.getAttribute("href")
		Next
	Else .location = url
	End If
End With

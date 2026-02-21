Set ie = CreateObject("InternetExplorer.Application")

ie.Visible = True
ie.FullScreen = True

html = "<html><head>" & _
"<style>" & _
"body{background:black;color:white;font-family:Georgia;text-align:center;" & _
"font-size:32px;margin-top:10%;}" & _
"</style>" & _
"</head><body>" & _
"<div id='text'></div>" & _
"<audio autoplay loop>" & _
"<source src='music.mp3' type='audio/mpeg'>" & _
"</audio>" & _
"<script>" & _
"var poem = 'LOVE-LETTER-FOR-YOU\n\n" & _
"If I could fold my heart into an envelope,\n" & _
"I would seal it with the quiet hope\n" & _
"that when you open it,\n" & _
"you feel how softly you are held in my thoughts.\n\n" & _
"You are the place my heart returns to,\n" & _
"again and again.\n\n" & _
"Forever yours.';" & _
"var i=0;" & _
"function typeWriter(){" & _
"if(i<poem.length){" & _
"document.getElementById(""text"").innerHTML+=" & _
"(poem.charAt(i)=='\n'?'<br>':poem.charAt(i));" & _
"i++;setTimeout(typeWriter,40);" & _
"}else{setTimeout(function(){" & _
"document.body.innerHTML='<h1>I love you.</h1>';" & _
"},3000);" & _
"setTimeout(function(){window.close();},7000);" & _
"}};" & _
"typeWriter();" & _
"</script></body></html>"

ie.Navigate "about:blank"

Do While ie.Busy Or ie.ReadyState <> 4
    WScript.Sleep 100
Loop

ie.Document.Write html
ie.Document.Close

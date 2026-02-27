Option Explicit
Randomize

Function HtmlEncode(s)
  s = CStr(s)
  s = Replace(s, "&", "&amp;")
  s = Replace(s, "<", "&lt;")
  s = Replace(s, ">", "&gt;")
  s = Replace(s, """", "&quot;")
  HtmlEncode = s
End Function

Function RandomColor()
  Dim r, g, b
  r = Int(200 * Rnd) + 55  ' Lighter colors (55-255)
  g = Int(200 * Rnd) + 55
  b = Int(200 * Rnd) + 55
  RandomColor = "rgb(" & r & "," & g & "," & b & ")"
End Function

Dim ads, idx, ad, shell, fso, temp, htaPath, html, ts, bgColor
ads = Array( _
    "Upgrade now: Super Premium Vacuum 3000 — 90% off today!", _
    "You won a free cruise! Click OK to claim", _
    "Buy one pixel, get one free. Limited time offer!", _
    "Exclusive deal: Lifetime supply of coffee for early coders.", _
    "Introducing the new SmartHome Hub: Control your home from anywhere.", _
    "Get 50% off your first year of cloud backup. Protect your files today.", _
    "Try our advanced password manager—secure, simple, and free for 30 days.", _
    "Boost productivity with OfficeSuite Pro. Special pricing for new users.", _
    "Upgrade your internet speed with FiberMax. Check availability in your area.", _
    "Experience next-gen gaming with the UltraPlay Console. Pre-order now.", _
    "Save on energy bills with EcoLight LED bulbs. Order your starter pack.", _
    "Professional antivirus protection—get a free security assessment.", _
    "Stream unlimited movies and shows with StreamFlix Premium. Start your trial.", _
    "Find your perfect match! Join LoveConnect dating today.", _
    "Meet singles near you—start chatting now on HeartLink.", _
    "Discover true love with MatchMakers. Sign up for free!", _
    "恋人募集中！今すぐ登録して素敵な出会いを見つけよう。", _
    "新しい友達や恋人を探すなら、デートサイトにアクセス！", _
    "理想の相手と出会えるチャンス。無料登録はこちら。", _
    "寻找真爱？立即注册，开启浪漫之旅！", _
    "单身交友，轻松聊天，马上加入！", _
    "免费注册，发现附近的有缘人。", _
    "心动时刻，遇见理想伴侣。", _
    "加入我们，开启幸福人生！" _
)

idx = Int((UBound(ads) + 1) * Rnd)
ad = ads(idx)
bgColor = RandomColor()

Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

temp    = shell.ExpandEnvironmentStrings("%TEMP%")
htaPath = fso.BuildPath(temp, "fake_ad_" & Replace(CStr(Timer), ".", "") & ".hta")

html = ""
html = html & "<html>" & vbCrLf
html = html & "<head>" & vbCrLf
html = html & "<title>Advertisement</title>" & vbCrLf
html = html & "<hta:application id='app' caption='no' border='thin' sysmenu='no' singleinstance='yes' maximizebutton='no' minimizebutton='no' scroll='no'/>" & vbCrLf
html = html & "<style type='text/css'>" & vbCrLf
html = html & "  body{font-family:Segoe UI,Tahoma,Arial,sans-serif;font-size:12pt;margin:0;padding:0;background-color:" & bgColor & ";}" & vbCrLf
html = html & "  .container{padding:24px;}" & vbCrLf
html = html & "  #msg{margin:12px 0 20px 0;padding:10px;background-color:rgba(255,255,255,0.8);border-radius:5px;font-weight:bold;}" & vbCrLf
html = html & "  .btn{padding:8px 16px;border-radius:4px;cursor:pointer;font-weight:bold;background-color:white;border:2px solid #333;}" & vbCrLf
html = html & "  .btn:hover{background-color:#f0f0f0;}" & vbCrLf
html = html & "</style>" & vbCrLf
html = html & "<script type='text/javascript'>" & vbCrLf
html = html & "function pos(){" & vbCrLf
html = html & "  var w=450,h=180;" & vbCrLf
html = html & "  try{window.resizeTo(w,h);}catch(e){}" & vbCrLf
html = html & "  var sw=(screen.availWidth||screen.width||w);" & vbCrLf
html = html & "  var sh=(screen.availHeight||screen.height||h);" & vbCrLf
html = html & "  var x=Math.floor(Math.random()*Math.max(1,sw-w));" & vbCrLf
html = html & "  var y=Math.floor(Math.random()*Math.max(1,sh-h));" & vbCrLf
html = html & "  try{window.moveTo(x,y);}catch(e){}" & vbCrLf
html = html & "  document.getElementById('ok').focus();" & vbCrLf
html = html & "}" & vbCrLf
html = html & "</script>" & vbCrLf
html = html & "</head>" & vbCrLf
html = html & "<body onload='pos()'>" & vbCrLf
html = html & "  <div class='container'>" & vbCrLf
html = html & "    <div id='msg'>" & HtmlEncode(ad) & "</div>" & vbCrLf
html = html & "    <button id='ok' class='btn' onclick='window.close()'>OK</button>" & vbCrLf
html = html & "  </div>" & vbCrLf
html = html & "</body>" & vbCrLf
html = html & "</html>" & vbCrLf

Set ts = fso.CreateTextFile(htaPath, True)
ts.Write html
ts.Close

shell.Run "mshta.exe " & Chr(34) & htaPath & Chr(34), 1, True

On Error Resume Next
fso.DeleteFile htaPath, True
On Error GoTo 0

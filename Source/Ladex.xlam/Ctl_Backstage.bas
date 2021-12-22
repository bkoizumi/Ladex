Attribute VB_Name = "Ctl_Backstage"
Public Sub setLisence01(control As IRibbonControl, ByRef Screentip)
  Dim strBuf As String
  
  
  strBuf = strBuf & thisAppName & " Ver. " & thisAppVersion & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " Copyright (c) 2021 Bunpei.Koizumi" & vbCrLf
  strBuf = strBuf & " author:bunpei.koizumi@gmail.com" & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " The MIT License (MIT)" & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " Permission is hereby granted, free of charge, to any person obtaining a copy" & vbCrLf
  strBuf = strBuf & " of this software and associated documentation files (the ""Software""), to deal" & vbCrLf
  strBuf = strBuf & " in the Software without restriction, including without limitation the rights" & vbCrLf
  strBuf = strBuf & " to use, copy, modify, merge, publish, distribute, sublicense, and/or sell" & vbCrLf
  strBuf = strBuf & " copies of the Software, and to permit persons to whom the Software is" & vbCrLf
  strBuf = strBuf & " furnished to do so, subject to the following conditions:" & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " The above copyright notice and this permission notice shall be included in all" & vbCrLf
  strBuf = strBuf & " copies or substantial portions of the Software." & vbCrLf
  
  Screentip = strBuf

End Sub
Public Sub setLisence02(control As IRibbonControl, ByRef Screentip)
  Dim strBuf As String
  
  strBuf = strBuf & " THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR" & vbCrLf
  strBuf = strBuf & " IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY," & vbCrLf
  strBuf = strBuf & " FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE" & vbCrLf
  strBuf = strBuf & " AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER" & vbCrLf
  strBuf = strBuf & " LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING " & vbCrLf
  strBuf = strBuf & " FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER " & vbCrLf
  strBuf = strBuf & " DEALINGS IN THE SOFTWARE." & vbCrLf
  
  Screentip = strBuf

End Sub

Public Sub setCopyright01(control As IRibbonControl, ByRef Screentip)
  Dim strBuf As String
  
  strBuf = strBuf & "免責事項" & vbCrLf
  strBuf = strBuf & "・当ソフトウェアの利用に際し、いかなるトラブルが発生しても、作者は一切の責任を負いません。" & vbCrLf
  
  Screentip = strBuf
  
End Sub

<div align="center">

## Using The MCI API


</div>

### Description

This article will show you how to play almost any type of multimedia file using the Window API only. It will show you how to manipulate a variety of commands that in turn will allow you to create professional standard applications.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Barry French](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/barry-french.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/barry-french-using-the-mci-api__1-31694/archive/master.zip)





### Source Code

<p><font face="Courier New, Courier, mono" size="2" color="#000000"> Lets start
 out by the two most important API declares. These will allow you to manipulate
 any multimedia file and return error's directly from the API.</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 Ok, the only way to show you how to open multimedia files is to jump straight
 in. The examples are pretty self explanitory, so don't worry too much.</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 Public Declare Function mciSendString Lib &quot;winmm.dll&quot; Alias &quot;mciSendStringA&quot;
 (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength
 As Long, ByVal hwndCallback As Long) As Long</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 Public Declare Function mciGetErrorString Lib &quot;winmm.dll&quot; Alias &quot;mciGetErrorStringA&quot;
 (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long)
 As Long</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 'This is just an API call to get the short name of the path you specify. The
 MCI uses short path formats<br>
 Public Declare Function GetShortPathName Lib &quot;kernel32&quot; Alias &quot;GetShortPathNameA&quot;
 (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer
 As Long) As Long</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 'This constant is just a string that you pass to the MCI so it knows what file
 your want manipulated<br>
 Const Alias As String = &quot;Media&quot;</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 'This function will return the error of specified MCI error<br>
 Private Function GetMCIError(lError As Long) As String<br>
 Dim sBuffer As string 'We need this to store the returned error</font><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 sBuffer = String$(255, Chr(0)) 'This fills out buffer with null characters so
 the MCI has something to write the error on</font><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 mciGetErrorString lError, sReturn, Len(sReturn)<br>
 sBuffer = Replace$(sBuffer, Chr(0), &quot;&quot;)<br>
 End Function</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 Private Function OpenMP3(FileName As String) As String<br>
 Dim lResult As Long 'The return value of the MCI command<br>
 Dim sBuffer As String 'The Buffer used to get the short path, we use it in the
 same way as mciGetErrorString<br>
 sBuffer = String$(255, Chr(0))<br>
 GetShortPathName FileName, sBuffer, Len(sBuffer)<br>
 sBuffer = Replace$(sBuffer, Chr(0), &quot;&quot;)<br>
 lResult = mciSendString(&quot;OPEN &quot; &amp; FileName &amp; &quot; TYPE MPEGVideo
 ALIAS &quot; &amp; Alias, 0, 0, 0)<br>
 If lResult Then 'There was an error<br>
 'We make our function return the MCI error<br>
 OpenMP3 = GetMCIError(lResult)<br>
 Exit Function<br>
 Else 'There was no error<br>
 'Set the timeformat of the file to milliseconds so when we send a request to
 get the length of the file or the curent playing position it will return in
 something we can understand<br>
 mciSendString &quot;SET &quot; &amp; Alias &amp; &quot; TIME FORMAT TMSF&quot;,
 0, 0, 0<br>
 End Function</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 Private Sub CloseMP3()<br>
 'We dont need an error code for this becuase if it dosent close then there isnt
 much we can do about it<br>
 mciSendString &quot;CLOSE &quot; &amp; Alias, 0, 0, 0<br>
 End Sub</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 Private Sub PlayMP3(Optional lPosition As Long)<br>
 'We dont really need an error return code for this becuase if the file is playable
 the MCI would not have opened it in the first place<br>
 'The lPosition tells the MCI to play the MP3 from a certain position (in milliseconds)<br>
 mciSendString &quot;PLAY &quot; &amp; Alias &amp; &quot; FROM &quot; &amp; lPosition,
 0, 0, 0<br>
 End Sub</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 They are the basics of playing media files. I thought I'd show you an MP3 file
 becuase they are more fun. Now you have the basics you can incorporate it with
 the lst below. Below is a list of all the stuff you can do with the MCI.<br>
 All commands follow the same pattern e.g.</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 mciSendString &quot;You command string&quot; &amp; Your Alias &amp; &quot; Aditional
 Commands&quot;, 0, 0, 0<br>
 If you are requesting a return value remember you must use a buffer</font></p>
<p><font face="Courier New, Courier, mono" size="2" color="#000000"><br>
 <b>Command Strings</b><br>
 &quot;PAUSE&quot;<br>
 &quot;STOP&quot;<br>
 &quot;SEEK&quot; Same as PLAY ALIAS FROM<br>
 &quot;OPEN AS CDAUDIO&quot; opens it for a CD Audio<br>
 &quot;OPEN AS MPEGVideo&quot; Opens ANY MPEG File<br>
 &quot;SETAUDIO ALIAS LEFT VOLUME TO NUMBER&quot;<br>
 &quot;SETAUDIO ALIAS RIGHT VOLUME TO NUMBER&quot;<br>
 &quot;STATUS ALIAS LENGTH&quot;<br>
 &quot;STATUS ALIAS POSITION&quot; </font> </p>


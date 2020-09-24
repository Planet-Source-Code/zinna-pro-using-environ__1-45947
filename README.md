<div align="center">

## Using Environ


</div>

### Description

Many people ask how to get certain system information and environ is the solution.

Environ is a command that allows you to get system environmental information.

It can also get any of the Environment Variables from the [System Properties, Advanced, Environment Variables] settings in windows.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Zinna Pro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zinna-pro.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zinna-pro-using-environ__1-45947/archive/master.zip)





### Source Code

<p><b><font size="4">Using Environ</font></b></p>
<ul>
 <li>Environ is a command that allows you to get system environmental
 information.</li>
 <li>It can also get any of the Environment Variables from the [System
 Properties, Advanced, Environment Variables]&nbsp; settings in windows.</li>
</ul>
<p><font size="4"><b>Syntax: </b>Environ(expression)</font></p>
<blockquote>
 <p><i>Where expression is includes the environmental variable you wish to
 retrieve</i></p>
</blockquote>
<p><b><font size="4">Example:</font></b></p>
<blockquote>
 <p><font size="2">Call MsgBox(&quot;Your Temp Directory is: &quot; &amp; Environ(&quot;Temp&quot;),
 vbInformation, &quot;Temp Directory Finder&quot;)</font></p>
</blockquote>
<p><b><font size="4">List of Common Expressions and sample output:</font></b></p>
<table border="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
 <tr>
  <th width="25%" height="19"><b>Expression</b></th>
  <th width="75%" height="19"><b>What came from my computer</b></th>
 </tr>
 <tr>
  <td width="25%" height="19">ALLUSERSPROFILE</td>
  <td width="75%" height="19">H:\Documents and Settings\All Users</td>
 </tr>
 <tr>
  <td width="25%" height="19">APPDATA</td>
  <td width="75%" height="19">H:\Documents and Settings\ZinnaPro\Application
  Data</td>
 </tr>
 <tr>
  <td width="25%" height="19">CLIENTNAME</td>
  <td width="75%" height="19">Console</td>
 </tr>
 <tr>
  <td width="25%" height="19">CommonProgramFiles</td>
  <td width="75%" height="19">H:\Program Files\Common Files</td>
 </tr>
 <tr>
  <td width="25%" height="19">COMPUTERNAME</td>
  <td width="75%" height="19">LOCALHOST</td>
 </tr>
 <tr>
  <td width="25%" height="19">ComSpec</td>
  <td width="75%" height="19">H:\WINDOWS\system32\cmd.exe</td>
 </tr>
 <tr>
  <td width="25%" height="19">HOMEDRIVE</td>
  <td width="75%" height="19">H:</td>
 </tr>
 <tr>
  <td width="25%" height="19">HOMEPATH</td>
  <td width="75%" height="19">\Documents and Settings\ZinnaPro</td>
 </tr>
 <tr>
  <td width="25%" height="19">LOGONSERVER</td>
  <td width="75%" height="19">\\LOCALHOST</td>
 </tr>
 <tr>
  <td width="25%" height="19">NUMBER_OF_PROCESSORS</td>
  <td width="75%" height="19">1</td>
 </tr>
 <tr>
  <td width="25%" height="19">OS</td>
  <td width="75%" height="19">Windows_NT</td>
 </tr>
 <tr>
  <td width="25%" height="19">Path</td>
  <td width="75%" height="19">H:\WINDOWS\system32;H:\WINDOWS;H:\WINDOWS\System32\Wbem</td>
 </tr>
 <tr>
  <td width="25%" height="19">PATHEXT</td>
  <td width="75%" height="19">.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH</td>
 </tr>
 <tr>
  <td width="25%" height="19">PROCESSOR_ARCHITECTURE</td>
  <td width="75%" height="19">x86</td>
 </tr>
 <tr>
  <td width="25%" height="19">PROCESSOR_IDENTIFIER</td>
  <td width="75%" height="19">x86 Family 6 Model 6 Stepping 2, AuthenticAMD</td>
 </tr>
 <tr>
  <td width="25%" height="19">PROCESSOR_LEVEL</td>
  <td width="75%" height="19">6</td>
 </tr>
 <tr>
  <td width="25%" height="19">PROCESSOR_REVISION</td>
  <td width="75%" height="19">0602</td>
 </tr>
 <tr>
  <td width="25%" height="19">ProgramFiles</td>
  <td width="75%" height="19">H:\Program Files</td>
 </tr>
 <tr>
  <td width="25%" height="19">SystemDrive</td>
  <td width="75%" height="19">H:</td>
 </tr>
 <tr>
  <td width="25%" height="19">SystemRoot</td>
  <td width="75%" height="19">H:\WINDOWS</td>
 </tr>
 <tr>
  <td width="25%" height="19">TEMP</td>
  <td width="75%" height="19">H:\DOCUME~1\ZinnaPro\LOCALS~1\Temp</td>
 </tr>
 <tr>
  <td width="25%" height="19">TMP</td>
  <td width="75%" height="19">H:\DOCUME~1\ZinnaPro\LOCALS~1\Temp</td>
 </tr>
 <tr>
  <td width="25%" height="19">USERDOMAIN</td>
  <td width="75%" height="19">LOCALHOST</td>
 </tr>
 <tr>
  <td width="25%" height="19">USERNAME</td>
  <td width="75%" height="19">ZinnaPro</td>
 </tr>
 <tr>
  <td width="25%" height="19">USERPROFILE</td>
  <td width="75%" height="19">H:\Documents and Settings\ZinnaPro</td>
 </tr>
 <tr>
  <td width="25%" height="17">windir</td>
  <td width="75%" height="17">H:\WINDOWS</td>
 </tr>
</table>
<p><font size="4"><b>Full List of variables my computer can get:</b></font></p>
<blockquote>
 <ul>
  <li>ALLUSERSPROFILE</li>
  <li>APPDATA</li>
  <li>CLIENTNAME</li>
  <li>CommonProgramFiles</li>
  <li>COMPUTERNAME</li>
  <li>ComSpec</li>
  <li>DRIVERNETWORKS</li>
  <li>DRIVERWORKS</li>
  <li>HOMEDRIVE</li>
  <li>HOMEPATH</li>
  <li>include</li>
  <li>lib</li>
  <li>LOGONSERVER</li>
  <li>MSDevDir</li>
  <li>NUMBER_OF_PROCESSORS</li>
  <li>OS</li>
  <li>Path</li>
  <li>PATHEXT</li>
  <li>PROCESSOR_ARCHITECTURE</li>
  <li>PROCESSOR_IDENTIFIER</li>
  <li>PROCESSOR_LEVEL</li>
  <li>PROCESSOR_REVISION</li>
  <li>ProgramFiles</li>
  <li>SESSIONNAME</li>
  <li>SystemDrive</li>
  <li>SystemRoot</li>
  <li>TEMP</li>
  <li>TMP</li>
  <li>USERDOMAIN</li>
  <li>USERNAME</li>
  <li>USERPROFILE</li>
  <li>VTOOLSD</li>
  <li>windir</li>
 </ul>
</blockquote>
<p><b><font size="4">Code used to get all variables:</font></b></p>
<blockquote>
 <p>Dim environinfo as string <font color="#008080">'Declare environinfo
 variable as string<br>
 </font>On Error goto done <font color="#008080">'If an error occurs then goto
 end (error expected due to large for loop)</font><br>
 environinfo = &quot;&quot;<font color="#008080"> 'Set our variable to nothing</font><br>
 For i = 1 To 200<font color="#008080"> 'Do a large for loop and try 200
 environmental indexes augmenting i by 1</font><br>
&nbsp;&nbsp; environinfo = environinfo &amp; Environ(i) &amp; vbCrLf
 <font color="#008080">'Augment variable and add put on separate lines</font><br>
 Next i<font color="#008080"> 'Loop until i is 200</font><br>
 <br>
 done: <font color="#008080">'Where to go on error, means we are done in this
 case</font><br>
 clipboard.settext environinfo <font color="#008080">'Copy our variable
 information to clipboard</font><br>
 Call MsgBox(&quot;Your Environment Variables have been copied to clipboard&quot;,
 vbInformation, &quot;Environs Found&quot;) <font color="#008080">'alert the user that
 the variables are in clipboard</font><br>
&nbsp;</p>
</blockquote>


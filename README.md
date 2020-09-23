<div align="center">

## Setting and getting file attributes w/o affecting other attributes \- Updated


</div>

### Description

These two simple wrappers can be used for setting and retrieving individual or selected file attributes without affecting the other attributes of the file. For example, to set the Archive bit of a file you should not just set its attributes to vbArchive (32), as this will turn off any other attributes currently set. Normally you would need to get the file attributes, add the desired attribute to the current attributes, then set them again. These wrappers just hide the details of this process. Update thanks to redbird77.

----

With the GetAttrib wrapper you can easily test the current state of an attribute. You simply specify the attribute(s) and the function will return True if the specified attribute(s) is set to on. You can specify more than one attribute and True will be returned only if all specified attributes are turned on.

----

The SetAttrib wrapper simplifies the setting of an attribute to on or off, and will be set as desired irrespective of its current state. You can set more than one attibute at a time (eg. SetAttrib(sFile, vbReadOnly Or vbHidden) will set these attributes to on, no matter if they are on or off, without affecting other attributes that may be set.

----

My first version of the SetAttrib function used 'Xor' to turn attributes off, but thanks to redbird77, I have updated it to 'Not' which has eliminated a limitation of my version. The bottom line is that you can set a files attribute(s) to on or off without needing to know the current state of the specified attributes, while not affecting other attributes that may be set.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-setting-and-getting-file-attributes-w-o-affecting-other-attributes-updated__1-56686/archive/master.zip)





### Source Code

<code><font color="#006600">
<p>&#160;</p>
<p><nobr>' The return value is the sum of the attribute values</font><br />
<font color="#000099">Public Declare Function GetAttributes Lib "kernel32" _<br />
&#160; &#160; Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long</nobr></p>
<font color="#006600"><p><nobr>' Sets the Attributes argument whose sum specifies file attributes<br />
<font color=#186125>' An error occurs if you try to set the attributes of an open file</font><br />
</font><font color="#000099">Public Declare Function SetAttributes Lib "kernel32" _<br />
&#160; &#160; Alias "SetFileAttributesA" (ByVal lpSpec As String, _<br />
&#160; &#160; ByVal dwAttributes As Long) As Long</nobr></p>
<font color="#000099">
<p><nobr>Public Enum vbFileAttributes<br />
&#160; vbNormal = 0</font> &#160; &#160; &#160; &#160; <font color="#006600">' Normal</font><br /><font color="#000099">
&#160; vbReadOnly = 1</font> &#160; &#160; &#160; <font color="#006600">' Read-only</font><br /><font color="#000099">
&#160; vbHidden = 2</font> &#160; &#160; &#160; &#160; <font color="#006600">' Hidden</font><br /><font color="#000099">
&#160; vbSystem = 4</font> &#160; &#160; &#160; &#160; <font color="#006600">' System file</font><br /><font color="#000099">
&#160; vbVolume = 8</font> &#160; &#160; &#160; &#160; <font color="#006600">' Volume label</font><br /><font color="#000099">
&#160; vbDirectory = 16</font> &#160; &#160; <font color="#006600">' Directory or folder</font><br /><font color="#000099">
&#160; vbArchive = 32</font> &#160; &#160; &#160; <font color="#006600">' File has changed since last backup</font><br /><font color="#000099">
&#160; vbTemporary = &H100</font> &#160;<font color="#006600">' 256</font><br /><font color="#000099">
&#160; vbCompressed = &H800</font> <font color="#006600">' 2048</font><br /><font color="#000099">
End Enum</nobr></p>
<br><hr width="70%" size="1" align="left" />
<p><nobr>Public Function GetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes) As Boolean<br /></font>
&#160; <font color="#006600">' Returns True if the specified attribute(s) is currently set.</font><br /><font color="#000099">
&#160; If (LenB(sFileSpec) <> 0) Then<br />
&#160; &#160; GetAttrib = (GetAttributes(sFileSpec) And Attrib) = Attrib<br />
&#160; End If<br />
End Function</nobr></p>
<br><p>Public Sub SetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes, Optional fTurnOff As Boolean)<br /></font><nobr>
&#160; <font color="#006600">' Sets/clears the specified attribute(s) without affecting other attributes. You<br />
&#160; ' do not need to know the current state of an attribute to set it to on or off.</font><br /><font color="#000099">
&#160; If (LenB(sFileSpec) <> 0) Then<br />
&#160; &#160; If (Attrib = vbNormal) Then<br />
&#160; &#160; &#160; SetAttributes sFileSpec, vbNormal<br />
&#160; &#160; ElseIf fTurnOff Then<br />
&#160; &#160; &#160; SetAttributes sFileSpec, GetAttributes(sFileSpec) And (Not Attrib)<br />
&#160; &#160; Else<br />
&#160; &#160; &#160; SetAttributes sFileSpec, GetAttributes(sFileSpec) Or Attrib<br />
&#160; &#160; End If<br />
&#160; End If<br />
End Sub</nobr></p>
<br></font></code>


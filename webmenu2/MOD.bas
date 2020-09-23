Attribute VB_Name = "MOD"
'With some basic web scripting knowledge you can make some pretty sweet interface
'on a form using the microsoft webbrowser control. If you need any web scripting help check out this site: www.w3schools.com , it's a really nice site.

'menu variable
Public CurrentMenuTheme As String

Public Function Write2Web(WebControl As WebBrowser, What As String)
'this is a simple function to write to the webbrowser control
'allows us to set what webbrowser control & what to write
WebControl.Document.write What
End Function

Public Function RightClickScript(Link As String, WebControl As WebBrowser)
'web script for if user "right-clicks" ur web browser control (a.k.a ur menu)
Call Write2Web(WebControl, "<SCRIPT>" & vbCrLf & _
  "var navwin='" & Link & "';" & vbCrLf & _
  "function clickIE() {if (document.all) {window.navigate(navwin);return false;}}" & vbCrLf & _
  "function clickNS(e) {if" & vbCrLf & _
  "(document.layers||(document.getElementById&&!document.all)) {" & vbCrLf & _
  "if (e.which==2||e.which==3) {window.navigate(navwin);return false;}}}" & vbCrLf & _
  "if (document.layers)" & vbCrLf & _
  "{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;}" & vbCrLf & _
  "else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;}" & vbCrLf & _
  "</SCRIPT>")
End Function

'Public Function CssFormatting(WebControl As WebBrowser)
''web script for css style of ur menu
''I use basic HTML and Css script
'Call Write2Web(WebControl, "<STYLE type=text/css>" & vbCrLf & _
'  "TABLE.headertablemain {" & vbCrLf & _
'  "FONT-SIZE: 12px; BACKGROUND: #FFFFFF; COLOR: #ffffff; FONT-FAMILY: Tahoma; BORDER-COLLAPSE: collapse" & vbCrLf & _
'  "}" & vbCrLf & _
'  "TABLE.headertable {" & vbCrLf & _
'  "FONT-SIZE: 12px; BACKGROUND: #CCCCCC; COLOR: #ffffff; FONT-FAMILY: Tahoma; BORDER-COLLAPSE: collapse" & vbCrLf & _
'  "}" & vbCrLf & _
'  "<!--" & vbCrLf & _
'  "a:hover {color: #000000; text-decoration: none;}" & vbCrLf & _
' "a:link, a:active, a  {color: #0080c0; text-decoration: none;}" & vbCrLf & _
'  "-->" & vbCrLf & _
'  "</STYLE>")
'End Function

'Public Function MenuScript(WebControl As WebBrowser)
''web script for ur menu
''I use basic HTML and Css script
'Call Write2Web(WebControl, "<TABLE class=headertablemain width=90 border=1 bordercolor=#000000 cellpadding=0 cellspacing=0>" & vbCrLf & _
'  "<TD bgcolor=#000000><B><LI>Menu</LI></B></TD>" & vbCrLf & _
'  "<TR>" & vbCrLf & _
'  "<TD onMouseOver=style.backgroundColor='#999999' onMouseOut=style.backgroundColor='#FFFFFF'>&nbsp;&nbsp;<A href=http://127.0.0.1/WEBMENU>Web Menu</A></TD>" & vbCrLf & _
'  "<TR>" & vbCrLf & _
'  "<TD onMouseOver=style.backgroundColor='#999999' onMouseOut=style.backgroundColor='#FFFFFF'>&nbsp;&nbsp;<A href=http://127.0.0.1/WEBSCRIPT>Web Script</A></TD>" & vbCrLf & _
'  "<TR>" & vbCrLf & _
'  "<TD onMouseOver=style.backgroundColor='#999999' onMouseOut=style.backgroundColor='#FFFFFF'>&nbsp;&nbsp;<A href=http://127.0.0.1/CUSTOMIZE>Customize</A></TD>" & vbCrLf & _
'  "<TR>" & vbCrLf & _
'  "<TD onMouseOver=style.backgroundColor='#999999' onMouseOut=style.backgroundColor='#FFFFFF'>&nbsp;&nbsp;<A href=http://127.0.0.1/ABOUT>About</A></TD>" & vbCrLf & _
'  "<TR>" & vbCrLf & _
'  "<TD onMouseOver=style.backgroundColor='#999999' onMouseOut=style.backgroundColor='#FFFFFF'>&nbsp;&nbsp;<A href=http://www.planetsourcecode.com target=blank>Visit PSC</A></TD>" & vbCrLf & _
'  "<TR>" & vbCrLf & _
'  "<TD onMouseOver=style.backgroundColor='#999999' onMouseOut=style.backgroundColor='#FFFFFF'>&nbsp;&nbsp;<A href=mailto:v_b_h_e_l_p@yahoo.com?subject=Hello%20Sweet%20Project! target=blank>Email Me</A></TD>" & vbCrLf & _
'  "<TR>" & vbCrLf & _
'  "<TD bgcolor=#000000>&nbsp;</TD>" & vbCrLf & _
'  "</TR>" & vbCrLf & _
'  "</TABLE>")
'End Function

Public Function LoadMenuTheme(Location As String, CurrentTheme As String)
On Error Resume Next 'just incase an error occurs
Dim FileDat As String
If FileExists(Location) = False Then Exit Function 'make sure the Menu Theme is in this location to be able to load it
CurrentTheme = ""
  Open Location For Input As #1 'open the document (#1 is the var we set as the file)
    While Not EOF(1) 'Does loop untill it reaches the "End Of File" (EOF())
      Input #1, FileDat 'pulls the data
        DoEvents:
        CurrentTheme = CurrentTheme & vbCrLf & FileDat 'sets the data to our theme variable
    Wend
  Close #1 'closes the document once we are done reading from it
End Function

Public Function LoadThemes(Location As String, Combo As ComboBox)
On Error Resume Next
Dim Theme As String
If FileExists(Location & "\*.INI") = False Then 'check to make sure there are Menu Themes. If not end application
  Call MsgBox("No Menu Themes were found in Location: " & Location, vbOKOnly, "Critical!")
  End
  Exit Function
End If
inifile = Dir(Location & "\*.INI") 'the * is a "Wild-Card" variable which allows us to search all possible files in the specified location with the extention .INI
  While inifile <> ""
    Theme = Mid(inifile, 1, Len(inifile) - 4) 'using the mid() function lets pull out the files extention
    Call Combo.AddItem(Theme) 'add the theme to the combo box
    inifile = Dir()
  Wend
Combo.Text = Combo.List(0)
End Function

Public Function FileExists(Location As String) As Boolean
'checks to make sure a file exists
If Len(Dir(Location)) > 0 Then  'lets use the len() & dir() functions to check if if the file exists by checking if files length
  FileExists = True 'if exists
Else
  FileExists = False 'if not
End If
End Function

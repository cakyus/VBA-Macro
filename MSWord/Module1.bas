Option Explicit

' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License version 2 as
' published by the Free Software Foundation.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program. If not, see <http://www.gnu.org/licenses/>.

' ----------------------------------------------------------------------
' Name:
' Description:
' License: GPL-2.0-only
' Homepage: https://github.com/cakyus/VBA-Macro/
' References:
'  - Microsot Scripting Runtime
' ----------------------------------------------------------------------

Sub Command_PasteAsText()
  Dim sText
  sText = CreateObject("HTMLFile").parentWindow.clipboardData.getData("Text")
  ' remove line break
  sText = Replace(sText, Chr(10), " ")
  sText = Replace(sText, Chr(13), " ")
  ' remove tab
  sText = Replace(sText, Chr(9), " ")
  ' remove double spaces
  sText = Text_RemoveSpaces(sText)
  ' remove space at begining and the end
  sText = Trim(sText)
  Selection.Text = sText
End Sub

' Remove two or more spaces

Function Text_RemoveSpaces(sText)
  Dim sText1
  sText1 = Text_RemoveTwoSpace(sText)
  If Len(sText) = Len(sText1) Then
    Text_RemoveSpaces = sText
  Else
    Text_RemoveSpaces = Text_RemoveSpaces(sText1)
  End If
End Function

' Replace two spaces to a space

Function Text_RemoveTwoSpace(sText)
  Text_RemoveTwoSpace = Replace(sText, "  ", " ")
End Function

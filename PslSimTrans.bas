'' Changes the pseudo translation algorithm
'' Implements call back PSL_OnSimulateTranslation

'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND
'EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE

Option Explicit

' We want to check if Asian characters are processed correctly
Public Sub PSL_OnSimulateTranslation(SrcStr As PslSourceString, _
                      Text As String, X As Long, Y As Long, _
                      Cx As Long, Cy As Long, Handled As Boolean)

   Handled = True ' Yes we handle simulation

   If SrcStr.Type = "DialogFont" Then
       Text = "Arial Unicode MS" 'If font, choose a Unicode font
   Else
       Text = "[" & SrcStr.Text & ChrW(20442) & "]" ' Add chars
   End If
End Sub

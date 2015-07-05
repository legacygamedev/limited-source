Attribute VB_Name = "modText"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hdc As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hdc, GameFont)
    Call SetBkMode(hdc, vbTransparent)
    Call SetTextColor(hdc, RGB(0, 0, 0))
    Call TextOut(hdc, x + 2, y + 2, Text, Len(Text))
    Call TextOut(hdc, x + 1, y + 1, Text, Len(Text))
    Call SetTextColor(hdc, Color)
    Call TextOut(hdc, x, y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbNewLine & Msg
    frmCClient.txtChat.SelStart = Len(frmCClient.txtChat.Text)
    frmCClient.txtChat.SelColor = QBColor(Color)
    frmCClient.txtChat.SelText = s
    frmCClient.txtChat.SelStart = Len(frmCClient.txtChat.Text) - 1
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Function Parse(ByVal Num As Long, ByVal Data As String)
Dim i As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For i = 1 To Len(Data)
        If Mid(Data, i, 1) = SEP_CHAR Then
            If n = Num Then
                eChar = i
                Parse = Mid(Data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = i + 1
            n = n + 1
        End If
    Next i
End Function

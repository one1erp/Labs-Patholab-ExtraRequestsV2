VERSION 5.00
Begin VB.Form frmRemarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "הערות"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRemark 
      Alignment       =   1  'Right Justify
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private con As Connection



Public Sub Initialize(con_ As Connection, strRequestDataId As String)
10    On Error GoTo ERR_Initialize
          Dim rs As Recordset
          Dim sql As String
          
20        Set con = con_
          
30        sql = " select r.DESCRIPTION "
40        sql = sql & " from lims_sys.u_extra_request_data rd, "
50        sql = sql & "      lims_sys.u_extra_request_data_user rdu, "
60        sql = sql & "      lims_sys.u_extra_request r"
70        sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
80        sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
90        sql = sql & "  and   rd.U_EXTRA_REQUEST_DATA_ID=" & strRequestDataId

100       Set rs = con.Execute(sql)
          
110       If rs.EOF = True Then Exit Sub

120       txtRemark.Text = nte(rs("DESCRIPTION"))

130       Exit Sub
ERR_Initialize:
140   MsgBox "ERR_Initialize" & vbCrLf & Err.Description
End Sub



Private Function nte(e As Variant) As Variant
10        nte = IIf(IsNull(e), "", e)
End Function

Private Sub Form_Click()
10        Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
10        If KeyAscii = vbKeyEscape Then
20            Me.Hide
30        End If
End Sub

Private Sub txtRemark_Click()
10        Me.Hide
End Sub

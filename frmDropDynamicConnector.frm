VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDropDynamicConnector 
   Caption         =   "Drop Dynamic Connector"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4500
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDropDynamicConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    ' Initialize radio buttons and other controls
    OptionButton1.Caption = "Direct type"
    OptionButton2.Caption = "Reply type"
    OptionButton3.Caption = "Self activity type"
    TextBox1.Text = "12" ' Default font size
    TextBox2.Text = "0" ' Default line color (black)
    TextBox3.Text = "1.5" ' Default line width
End Sub

Private Sub CommandButton1_Click()
    ' Drop button logic
    Dim connectorType As String
    Dim arrowType As String
    Dim fontSize As String
    Dim lineColor As String
    Dim lineWidth As String
    
    If OptionButton1.Value Then
        connectorType = "0"
        arrowType = "1"
    ElseIf OptionButton2.Value Then
        connectorType = "4"
        arrowType = "2"
    ElseIf OptionButton3.Value Then
        connectorType = "0"
        arrowType = "0"
    End If
    
    fontSize = TextBox1.Text
    lineColor = TextBox2.Text
    lineWidth = TextBox3.Text
    
    Call DropDynamicConnector(connectorType, arrowType, fontSize, lineColor, lineWidth)
End Sub

Private Sub CommandButton2_Click()
    ' Close button logic
    Unload Me
End Sub

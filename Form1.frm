VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Address Validator"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Validate"
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   45
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1275
      TabIndex        =   0
      Text            =   "255.255.255.13"
      Top             =   45
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address :"
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   90
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Valid_IP(IP As String) As Boolean
    Dim i As Integer
    Dim dot_count As Integer
    Dim test_octet As String
    Dim byte_check
     IP = Trim$(IP)

     ' make sure the IP long enough before
     ' continuing
     If Len(IP) < 8 Then
        Valid_IP = False
        'Show Message
        MsgBox IP & " is Invalid", , "IP Validator"
        Exit Function
    End If

    i = 1
    dot_count = 0
    For i = 1 To Len(IP)
        If Mid$(IP, i, 1) = "." Then
            ' increment the dot count and
            ' clear the test octet variable
            dot_count = dot_count + 1
            test_octet = ""
            If i = Len(IP) Then
                ' we've ended with a dot
                ' this is not good
                Valid_IP = False
                'Show Message
                MsgBox IP & " is Invalid", , "IP Validator"
                Exit Function
            End If
        Else
            test_octet = test_octet & Mid$(IP, i, 1)
            On Error Resume Next
            byte_check = CByte(test_octet)
            If (Err) Then
                ' either the value is not numeric
                ' or exceeds the range of the byte
                ' data type.
                Valid_IP = False
                Exit Function
            End If
        End If
    Next i
     ' so far, so good
      ' did we get the correct number of dots?
    If dot_count <> 3 Then
        Valid_IP = False
        Exit Function
    End If
     ' we have a valid IP format!
    Valid_IP = True
        'Show Message
        MsgBox IP & " is Valid", , "IP Validator"
    
End Function

Private Sub Command1_Click()
    If Len(Text1) = 0 Then
        MsgBox "Please type an IP Address in the textbox.", , "IP Validator"
    Else
        'Call the Function
        Valid_IP Text1
    End If
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTinyComp 
   Caption         =   "Simple Lexical Scanner"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView LstView 
      Height          =   2385
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   4207
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   480
      Left            =   90
      TabIndex        =   1
      Top             =   4905
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmTinyComp.frx":0000
      Top             =   2580
      Width           =   5085
   End
End
Attribute VB_Name = "FrmTinyComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple Lexical Scanner
'By DreamVB

' Hi made this in the hope that it may be of some use to any one.
' wanted to know how to write a Basic Lexical for what ever purpose
' The meain Reason I made it was to build a small simple Compiler
' But you may use it for a scripting lanuage or anything you want/
' anyway Hope you like it

'Token Types
Enum Tok_Types
    NONE = 0
    LSTRING = 1
    DIGIT = 2
    QSTRING = 3
    LPARM = 4
    RPARM = 5
    LVARIABLE = 6
    COMMENT = 7
    KEYWORD = 8
    DEL = 9
    EOL = 10
    EOP = 11
End Enum

'Current processing char
Public CharPos As Long
'Source to scan
Dim Source As String

'Hold the returned Token
Dim Token As String
'Hold the Token Type
Dim TokType As Tok_Types
'Just hold some Keywords
Private Keywords(5) As String

Function GetStrToken(iTokT As Tok_Types) As String

    'All this does is return the string name of a Token Type ID
    
    Select Case iTokT
        Case NONE
            GetStrToken = "Nothing"
        Case LSTRING
            GetStrToken = "String"
        Case DIGIT
            GetStrToken = "Number"
        Case QSTRING
            GetStrToken = "Quote_String"
        Case LPARM
            GetStrToken = "LPARM"
        Case RPARM
            GetStrToken = "RPARM"
        Case LVARIABLE
            GetStrToken = "Variable"
        Case COMMENT
            GetStrToken = "Comment"
        Case KEYWORD
            GetStrToken = "IDENT"
        Case DEL
            GetStrToken = "Delimiter"
        Case EOL
            GetStrToken = "EOL"
        Case EOP
            GetStrToken = "EOP"
        Case Else
            GetStrToken = "UNKOWN"
    End Select
    
End Function

Function IsKeyword(Key As String) As Boolean
Dim x As Integer
    'Return true or false if we found a Keyword
    IsKeyword = False
    For x = 0 To 5
        If Keywords(x) = UCase(Key) Then IsKeyword = True: Exit For
    Next x
End Function

Public Function isAlpha(c As String) As Boolean
    'Return true if we only have letters a-z  A-Z
    isAlpha = UCase(c) >= "A" And UCase(c) <= "Z"
End Function

Public Function isWhite(c As String) As Boolean
    'Return true if we find a white space
    isWhite = (c = " ") Or (c = vbTab)
End Function

Public Function isDigit(c As String) As Boolean
    'Return true when we only have a digit
    isDigit = (c >= "0") And (c <= "9")
End Function

Function IsDelim(c As String) As Boolean
    'Return true if we have a Delimiter
    If InStr(" ,;<>+-/*%^=[]()&", c) Or c = vbCr Then IsDelim = True
End Function

Sub INC(Optional nMove As Integer = -1)
    If (nMove <> -1) Then
        CharPos = CharPos + nMove
    Else
        CharPos = CharPos + 1
    End If
End Sub

Sub GetToken()
    'This is the main part of the scanner
    ' scans the input source and builds the tokens, and assigns the types
    
    Token = ""
    TokType = NONE
    
    'If we are over the length of the souce we are at the end of the program
    If (CharPos > Len(Source)) Then
        'End of program
        TokType = EOP
        Token = Chr(11)
        Exit Sub
    End If
    
    'Skip over white-spaces
    Do While (CharPos <= Len(Source) And (isWhite(Mid(Source, CharPos, 1))))
       INC
        If CharPos > Len(Source) Then
            TokType = EOL
            Exit Sub
        End If
    Loop
    
    'Skip over Line Breaks
    If Mid(Source, CharPos, 1) = vbCr Then
        INC 2
        Token = vbCr
        TokType = EOL
        Exit Sub
    End If
    
    'Check for Delimiters  ,;<>+-/*%^=[]()&
    If IsDelim(Mid(Source, CharPos, 1)) Then
        Token = Token + Mid(Source, CharPos, 1)
        INC
        'Just something else I added to Ident opening and closeing bracets ()
        If (Token = "(") Then
            TokType = LPARM
        ElseIf (Token = ")") Then
            TokType = RPARM
        Else
            'Just return Delimiter type
            TokType = DEL
        End If
    'This part I added for my own purpose to ident a variable
    ElseIf Mid(Source, CharPos, 1) = "$" Then
        INC
        Token = "$"
        If isAlpha(Mid(Source, CharPos, 1)) Then
            While Not IsDelim(Mid(Source, CharPos, 1))
                Token = Token + Mid(Source, CharPos, 1)
                INC
            Wend
        End If
        TokType = LVARIABLE
        Exit Sub
    
    'Checks for only Alpha strings
    ElseIf isAlpha(Mid(Source, CharPos, 1)) Then
        While Not IsDelim(Mid(Source, CharPos, 1))
            Token = Token + Mid(Source, CharPos, 1)
            INC
            TokType = LSTRING
        Wend
        'Check if we have a keyword other wise it's a LSTRING
        If IsKeyword(Token) Then TokType = KEYWORD
        
        'Check for digits
    ElseIf isDigit(Mid(Source, CharPos, 1)) Then
        While Not IsDelim(Mid(Source, CharPos, 1))
            Token = Token + Mid(Source, CharPos, 1)
            If CharPos > Len(Source) Then Exit Sub
            INC
        Wend
        TokType = DIGIT
        Exit Sub
    
    'Check for quoted strings "hello world"
    ElseIf Mid(Source, CharPos, 1) = Chr(34) Then
        INC
        While Mid(Source, CharPos, 1) <> Chr(34) And Mid(Source, CharPos, 1) <> vbCr
            Token = Token + Mid(Source, CharPos, 1)
            INC
            If CharPos > Len(Source) Then
                Exit Sub
            End If
        Wend
        If Mid(Source, CharPos, 1) = vbCr Then
            Token = Chr(0)
            TokType = EOP
            Exit Sub
        End If
        INC
        TokType = QSTRING
        Exit Sub
    ElseIf (Mid(Source, CharPos, 1) = "'") Then
        While Mid(Source, CharPos, 1) <> vbCr
            Token = Token + Mid(Source, CharPos, 1)
            INC
            If (CharPos > Len(Source)) Then Exit Sub
        Wend
        TokType = COMMENT
    Else
        Beep
        Token = "PHASE_ERROR"
        TokType = EOP
    End If
    
End Sub

Private Sub Command1_Click()
Dim SubCnt As Integer
    On Error Resume Next
    
    'Code below Just added for a test
    LstView.ListItems.Clear
    LstView.ColumnHeaders.Clear
    
    LstView.ColumnHeaders.Add , , "Token", 2500
    LstView.ColumnHeaders.Add , , "Symbol"
    LstView.ColumnHeaders.Add , , "Symbol ID", 900

    CharPos = 1
    Source = Text1.Text
    
    Do
        'Loop until we hit the EOP token-type
        GetToken
        SubCnt = SubCnt + 1
        LstView.ListItems.Add , , Token
        LstView.ListItems(SubCnt).SubItems(1) = GetStrToken(TokType)
        LstView.ListItems(SubCnt).SubItems(2) = TokType
    Loop Until (TokType = EOP)
    
End Sub

Private Sub Form_Load()
    Keywords(0) = "PRINT": Keywords(1) = "END": Keywords(2) = "IF"
    Keywords(3) = "DO": Keywords(4) = "LOOP": Keywords(5) = "DIM"
End Sub

Attribute VB_Name = "Interpreter"
'������
'����ִ�г���
'���������ݽṹ��IsAlpha,IsDigit����

Option Explicit

'###������###
'����
Private Type Token
    Type As String
    Value As String
    Line As Long
    Column As Long
End Type

'����
Private Type Variable
    Name As String
    Type As String
    Value As Variant
End Type

'��������
Private Type ErrorData
    ErrorType As String
    Token As Token
    Log As String
End Type

'���ڵ�
Private Type Node
    Type As String
    Value As String
    Children() As Long
End Type

Private ErrorData As ErrorData '��������
Private IsDebugMode As Boolean '��־��debug�Ƿ���

'��������¼
Private Sub ClearError()
    ErrorData.ErrorType = ""
    ErrorData.Token = Token
    ErrorData.Log = ""
End Sub

'�������
Private Sub ClearVariable(ByVal Name As String, ByRef Variables() As Variable)
    Dim i As Long, j As Long
    For i = 1 To UBound(Variables)
        If Variables(i).Name = Name Then
            For j = i To UBound(Variables) - 1
                Variables(j) = Variables(j + 1)
            Next
            ReDim Preserve Variables(UBound(Variables) - 1)
            Exit Sub
        End If
    Next
End Sub

'��ʾTokenList
Private Sub DebugTokenList(TokenList() As Token)
    Dim i As Long
    Dim Text As String
    
    For i = 1 To UBound(TokenList)
        Text = Text & i & " " & TokenList(i).Type & " " & TokenList(i).Value & vbCrLf
    Next
    MsgBox Text
End Sub

'��ʾ����
Private Sub DebugVariables(Variables() As Variable)
    Dim i As Long
    Dim Text As String
    
    For i = 1 To UBound(Variables)
        Text = Text & Variables(i).Name & ":" & Variables(i).Value & vbCrLf
    Next
    If Text <> "" Then MsgBox Text
End Sub

'��ȡ����
Private Function GetVariable(ByVal Name As String, Variables() As Variable) As Variant
    Dim i As Long
    
    For i = 1 To UBound(Variables)
        If Variables(i).Name = Name Then
            GetVariable = Variables(i).Value
            Exit Function
        End If
    Next
    
    MsgBox "����ʱ����δ�ҵ�����""" & Name & """"
End Function

'��ȡ��������
Private Function GetVariableType(ByVal Name As String, ByRef Variables() As Variable) As String
    Dim i As Long
    
    For i = 1 To UBound(Variables)
        If Variables(i).Name = Name Then
            GetVariableType = Variables(i).Type
            Exit Function
        End If
    Next
    
    MsgBox "����ʱ����δ�ҵ�����""" & Variables(i).Name & """"
End Function

'����ִ��
Private Sub Interpret(AST() As Node)
    Dim Variables() As Variable
    Dim i As Long
    
    ReDim Variables(0)
    Visit AST, 1, Variables
    
    If ErrorData.ErrorType <> "" Then Exit Sub
    
    DebugVariables Variables
End Sub

'�ж��Ƿ�����ĸ
Public Function IsAlpha(Char As String) As Boolean
    If Char = "" Then
        IsAlpha = False
        Exit Function
    End If
    IsAlpha = Asc(Char) >= 65 And Asc(Char) <= 90 Or Asc(Char) >= 97 And Asc(Char) <= 122
End Function

'�ж��Ƿ�������
Public Function IsDigit(Char As String) As Boolean
    If Char = "" Then
        IsDigit = False
        Exit Function
    End If
    IsDigit = Asc(Char) >= 48 And Asc(Char) <= 57
End Function

'������
Private Function IsReserved(Text As String) As Boolean
    IsReserved = False
    
    If Text = "if" Then IsReserved = True 'if���
    If Text = "else" Then IsReserved = True 'else���
    If Text = "prog" Then IsReserved = True '�ӳ���
    If Text = "true" Then IsReserved = True '��
    If Text = "false" Then IsReserved = True '��
End Function

Private Sub Lexer(TokenList() As Token, ByVal Text As String) '�ʷ�������
    Dim Position As Long 'λ��
    Dim TokenLen As Long
    Dim Skip As Boolean
    
    Position = 1
    ReDim TokenList(0)
    
    Do Until Position > Len(Text)
        Skip = False
        
        If Not Skip And InStr(Mid(Text, Position), " ") = 1 Then
            Position = Position + 1
            Skip = True
        End If
        
        '����
        If Not Skip And IsDigit(Mid(Text, Position, 1)) Then
            TokenLen = 1
            Do While Position + TokenLen <= Len(Text)
                If IsDigit(Mid(Text, Position + TokenLen, 1)) Then
                    TokenLen = TokenLen + 1
                Else
                    Exit Do
                End If
            Loop
            
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Integer", Mid(Text, Position, TokenLen))
            Position = Position + TokenLen
            Skip = True
        End If
        
        '�ַ���
        If Not Skip And InStr(Mid(Text, Position), """") = 1 Then
            Position = Position + 1
            TokenLen = 0
            Do While Position + TokenLen <= Len(Text)
                If Not InStr(Mid(Text, Position + TokenLen), """") = 1 Then
                    TokenLen = TokenLen + 1
                Else
                    Exit Do
                End If
            Loop
            
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("String", Mid(Text, Position, TokenLen))
            Position = Position + TokenLen + 1
            Skip = True
        End If
        
        '���ƻ�����
        If Not Skip And IsAlpha(Mid(Text, Position, 1)) Then
            TokenLen = 1
            Do While Position + TokenLen <= Len(Text)
                If IsAlpha(Mid(Text, Position + TokenLen, 1)) Or IsDigit(Mid(Text, Position + TokenLen, 1)) Or Mid(Text, Position + TokenLen, 1) = "_" Then
                    TokenLen = TokenLen + 1
                Else
                    Exit Do
                End If
            Loop
            
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            If IsReserved(LCase(Mid(Text, Position, TokenLen))) Then
                TokenList(UBound(TokenList)) = Token("Reserved", LCase(Mid(Text, Position, TokenLen)))
            Else
                TokenList(UBound(TokenList)) = Token("ID", Mid(Text, Position, TokenLen))
            End If
            Position = Position + TokenLen
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "+") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Plus", "+")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "-") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Minus", "-")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "*") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Mul", "*")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "/") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Div", "/")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "(") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("LParen", "(")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), ")") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("RParen", ")")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "{") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("LBrace", "{")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "}") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("RBrace", "}")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), "=") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Assign", "=")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), ":") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Colon", ":")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), ";") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Semi", ";")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip And InStr(Mid(Text, Position), ",") = 1 Then
            ReDim Preserve TokenList(UBound(TokenList) + 1)
            TokenList(UBound(TokenList)) = Token("Comma", ",")
            Position = Position + 1
            Skip = True
        End If
        
        If Not Skip Then
            MsgBox "������󣺴ʷ�����ʧ��"
            Exit Sub
        End If
    Loop
    
    ReDim Preserve TokenList(UBound(TokenList) + 1)
    TokenList(UBound(TokenList)) = Token("Null", "")
End Sub

'�½��ڵ�
Private Function NodeCreate(Children() As Long, Optional ByVal NodeType As String = "Null", Optional ByVal Value As String) As Node
    With NodeCreate
        .Type = NodeType
        .Value = Value
        .Children = Children
    End With
End Function

'���ӽڵ�
Private Function NodeNoChild(Optional ByVal NodeType As String = "Null", Optional ByVal Value As String) As Node
    With NodeNoChild
        .Type = NodeType
        .Value = Value
        ReDim .Children(0)
    End With
End Function

'ɾ���ڵ�ָ��
Private Sub NodeDeleteChild(ByRef Node As Node, ID As Long)
    Dim i As Long
    
    For i = ID To UBound(Node.Children) - 1
        Node.Children(i) = Node.Children(i + 1)
    Next
    
    ReDim Preserve Node.Children(UBound(Node.Children) - 1)
End Sub

Private Sub Parser(TokenList() As Token, AST() As Node)   '�﷨������
    Dim Position As Long 'λ��
    
    Position = 1
    ParserStatementList TokenList, AST, Position
End Sub

'��ֵ���
Private Sub ParserAssign(TokenList() As Token, AST() As Node, ByRef Position As Long)
    Dim SubAST() As Node
    Dim Value As Variant
    
    '��ȡ��ֵ����ߵı���
    ParserVariable TokenList, SubAST, Position
    
    If TokenList(Position).Type = "Assign" Then
        Value = TokenList(Position).Value
        Position = Position + 1
        TreeCreateWithNode AST, NodeNoChild("Assign", Value)
        TreeMerge AST, SubAST, 1, 1
        '��ȡ��ֵ���ұߵĲ���
        ParserExpr TokenList, SubAST, Position
        TreeMerge AST, SubAST, 1, 2
    Else
        MsgBox "�������ȱ��="
        Exit Sub
    End If
End Sub

'�������
Private Sub ParserCompondStatement(TokenList() As Token, AST() As Node, ByRef Position As Long)
    If TokenList(Position).Type = "LBrace" Then
        Position = Position + 1
        ParserStatementList TokenList, AST, Position
        If TokenList(Position).Type = "RBrace" Then
            Position = Position + 1
        Else
            MsgBox "�������:ȱ��}"
        End If
        Exit Sub
    Else
    MsgBox "�������:ȱ��{"
    End If
End Sub

Private Sub ParserExpr(TokenList() As Token, AST() As Node, ByRef Position As Long) '���ʽ
    Dim SubAST() As Node
    Dim Value As Variant
    
    '��ȡ�Ӽ���ߵĲ���
    ParserTerm TokenList, AST, Position
    
    Do While Position <= UBound(TokenList)
        If TokenList(Position).Type = "Plus" Or TokenList(Position).Type = "Minus" Then
            Value = TokenList(Position).Value
            Position = Position + 1
            SubAST = AST
            TreeCreateWithNode AST, NodeNoChild("BinOp", Value)
            TreeMerge AST, SubAST, 1, UBound(AST(1).Children) + 1
            '��ȡ�Ӽ��ұߵĲ���
            ParserTerm TokenList, SubAST, Position
            TreeMerge AST, SubAST, 1, UBound(AST(1).Children) + 1
        Else
            Exit Sub
        End If
    Loop
End Sub

Private Sub ParserFactor(TokenList() As Token, AST() As Node, ByRef Position As Long) '���ֻ�����
    Dim SubAST() As Node
    Dim Value As Variant
    
    If Position > UBound(TokenList) Then
        MsgBox "�������:����Ľ�β"
        Exit Sub
    End If
    
    '����
    If TokenList(Position).Type = "Integer" Then
        Value = TokenList(Position).Value
        Position = Position + 1
        TreeCreateWithNode AST, NodeNoChild("Integer", Value)
        Exit Sub
    End If
    
    '����
    If TokenList(Position).Type = "Minus" Then
        Value = TokenList(Position).Value
        Position = Position + 1
        ParserFactor TokenList, SubAST, Position
        TreeCreateWithNode AST, NodeNoChild("UnaryOp", Value)
        TreeMerge AST, SubAST, 1, 1
        Exit Sub
    End If
    
    '����
    If TokenList(Position).Type = "ID" Then
        Value = TokenList(Position).Value
        Position = Position + 1
        TreeCreateWithNode AST, NodeNoChild("Variable", Value)
        Exit Sub
    End If
    
    '������
    If TokenList(Position).Type = "Reserved" Then
        If TokenList(Position).Value = "true" Or TokenList(Position).Value = "false" Then
            Value = TokenList(Position).Value
            Position = Position + 1
            TreeCreateWithNode AST, NodeNoChild("Boolean", Value)
            Exit Sub
        End If
    End If
    
    '�ַ���
    If TokenList(Position).Type = "String" Then
        Value = TokenList(Position).Value
        Position = Position + 1
        TreeCreateWithNode AST, NodeNoChild("String", Value)
        Exit Sub
    End If
    
    '����
    If TokenList(Position).Type = "LParen" Then
        Position = Position + 1
        ParserExpr TokenList, AST, Position
        If TokenList(Position).Type = "RParen" Then
            Position = Position + 1
        Else
            MsgBox "�������:ȱ��)"
        End If
        Exit Sub
    End If
    
    MsgBox "�������ȱ��ֵ"
End Sub

'if���
Private Sub ParserIf(TokenList() As Token, AST() As Node, ByRef Position As Long)
    Dim SubAST() As Node
    Dim Value As Variant
    
    '���if�ؼ���
    If TokenList(Position).Type = "Reserved" And TokenList(Position).Value = "if" Then
        Value = TokenList(Position).Value
        Position = Position + 1
        TreeCreateWithNode AST, NodeNoChild("If", Value)
        '��ȡ����
        ParserExpr TokenList, SubAST, Position
        TreeMerge AST, SubAST, 1, 1
        '���ð��
        If TokenList(Position).Type = "Colon" Then
            Position = Position + 1
            '��ȡ���
            ParserStatement TokenList, SubAST, Position
            TreeMerge AST, SubAST, 1, 2
            '���else�ؼ���
            If TokenList(Position).Type = "Reserved" And TokenList(Position).Value = "else" Then
                Position = Position + 1
                '��ȡ���
                ParserStatement TokenList, SubAST, Position
                TreeMerge AST, SubAST, 1, 3
            End If
        Else
            MsgBox "�������ȱ��:"
            Exit Sub
        End If
    Else
        MsgBox "�������ȱ��if"
        Exit Sub
    End If
End Sub

'�ӳ������
Private Sub ParserProgCall(TokenList() As Token, AST() As Node, ByRef Position As Long)
    Dim SubAST() As Node
    Dim Value As Variant
    
    '��ȡ�ӳ�����
    If TokenList(Position).Type = "ID" Then
        Value = TokenList(Position).Value
        TreeCreateWithNode AST, NodeNoChild("ProgCall", Value)
        Position = Position + 1
        
        'ʶ��������
        If TokenList(Position).Type = "LParen" Then
            Position = Position + 1
            '������������ž�ִ��
            If TokenList(Position).Type <> "RParen" Then
                'ʶ����ʽ
                ParserExpr TokenList, SubAST, Position
                If SubAST(1).Type <> "Null" Then
                    TreeMerge AST, SubAST, 1, UBound(AST(1).Children) + 1
                End If
                Do While Position <= UBound(TokenList)
                    '����ж��������ʶ����ȥ
                    If TokenList(Position).Type = "Comma" Then
                        Position = Position + 1
                        ParserExpr TokenList, SubAST, Position
                        If SubAST(1).Type <> "Null" Then
                            TreeMerge AST, SubAST, 1, UBound(AST(1).Children) + 1
                        End If
                    Else
                        '���������
                        If TokenList(Position).Type = "RParen" Then
                            Position = Position + 1
                            Exit Sub
                        Else
                            MsgBox "�������ȱ��)"
                            Exit Sub
                        End If
                    End If
                Loop
            End If
        End If
    Else
        MsgBox "�������ȱ�ٹ�����"
        Exit Sub
    End If
End Sub

'�ӳ�������
Private Sub ParserProgram(TokenList() As Token, AST() As Node, ByRef Position As Long)
    Dim SubAST() As Node
    Dim Value As Variant
    
    '���prog�ؼ���
    If TokenList(Position).Type = "Reserved" And TokenList(Position).Value = "prog" Then
        '��ȡ�ӳ�����
        Position = Position + 1
        If TokenList(Position).Type = "ID" Then
            Value = TokenList(Position).Value
            TreeCreateWithNode AST, NodeNoChild("Program", Value)
            Position = Position + 1
            '���ð��
            If TokenList(Position).Type = "Colon" Then
                Position = Position + 1
                '��ȡ���
                ParserStatement TokenList, SubAST, Position
                TreeMerge AST, SubAST, 1, 1
            Else
                MsgBox "�������ȱ��:"
                Exit Sub
            End If
        Else
            MsgBox "�������ȱ��:������"
            Exit Sub
        End If
    Else
        MsgBox "�������ȱ��if"
        Exit Sub
    End If
End Sub

'���
Private Sub ParserStatement(TokenList() As Token, AST() As Node, ByRef Position As Long)
    If TokenList(Position).Type = "ID" Then
        'ʶ��ϵͳ���
        If TokenList(Position).Value = "system" Then
            ParserProgCall TokenList, AST, Position
        Else
            'ʶ��ֵ���
            ParserAssign TokenList, AST, Position
        End If
        Exit Sub
    End If

    'ʶ�𸴺����
    If TokenList(Position).Type = "LBrace" Then
        ParserCompondStatement TokenList, AST, Position
        Exit Sub
    End If

    'ʶ��if���
    If TokenList(Position).Type = "Reserved" And TokenList(Position).Value = "if" Then
        ParserIf TokenList, AST, Position
        Exit Sub
    End If

    'ʶ���ӳ�������
    If TokenList(Position).Type = "Reserved" And TokenList(Position).Value = "prog" Then
        ParserProgram TokenList, AST, Position
        Exit Sub
    End If

    'ʶ������
    If TokenList(Position).Type = "Semi" Or TokenList(Position).Type = "RBrace" Or TokenList(Position).Type = "Null" Then
        TreeCreateWithNode AST, NodeNoChild("Null")
        Exit Sub
    End If

    MsgBox "�������:ȱ�����"
End Sub

'����б�
Private Sub ParserStatementList(TokenList() As Token, AST() As Node, ByRef Position As Long)
    Dim SubAST() As Node
    Dim Value As Variant
    
    'ʶ�����
    ParserStatement TokenList, SubAST, Position
    '���ʶ�𲻵�����򷵻ؿսڵ�
    If SubAST(1).Type <> "Null" Then
        TreeCreateWithNode AST, NodeNoChild("Compond", "")
        TreeMerge AST, SubAST, 1, 1
    Else
        TreeCreateWithNode AST, NodeNoChild("Null")
        Exit Sub
    End If
    
    Do While Position <= UBound(TokenList)
        '����зֺ������ʶ����ȥ
        If TokenList(Position).Type = "Semi" Then
            Position = Position + 1
            ParserStatement TokenList, SubAST, Position
            If SubAST(1).Type <> "Null" Then
                TreeMerge AST, SubAST, 1, UBound(AST(1).Children) + 1
            End If
        Else
            Exit Sub
        End If
    Loop
End Sub

Private Sub ParserTerm(TokenList() As Token, AST() As Node, ByRef Position As Long) '���г˳��ı��ʽ
    Dim SubAST() As Node
    Dim Value As Variant
    
    '��ȡ�˳���ߵĲ���
    ParserFactor TokenList, AST, Position
    
    Do While Position <= UBound(TokenList)
        If TokenList(Position).Type = "Mul" Or TokenList(Position).Type = "Div" Then
            Value = TokenList(Position).Value
            Position = Position + 1
            SubAST = AST
            TreeCreateWithNode AST, NodeNoChild("BinOp", Value)
            TreeMerge AST, SubAST, 1, UBound(AST(1).Children) + 1
            '��ȡ�˳��ұߵĲ���
            ParserFactor TokenList, SubAST, Position
            TreeMerge AST, SubAST, 1, UBound(AST(1).Children) + 1
        Else
            Exit Sub
        End If
    Loop
End Sub

Private Sub ParserVariable(TokenList() As Token, AST() As Node, ByRef Position As Long) '����
    Dim Value As Variant
    
    If TokenList(Position).Type = "ID" Then
        Value = TokenList(Position).Value
        Position = Position + 1
        TreeCreateWithNode AST, NodeNoChild("Variable", Value)
        Exit Sub
    End If
End Sub

'����ָ���ı���ִ��
Public Sub Run(ByVal Text As String, Optional ByVal DebugMode As Boolean = False)
    Dim TokenList() As Token
    Dim AST() As Node
    
    ClearError
    IsDebugMode = DebugMode
    
    Lexer TokenList, Text
    If ErrorData.ErrorType <> "" Then
'        MsgBox "�������"
        Exit Sub
    End If
    
    DebugTokenList TokenList
    Parser TokenList, AST
    If ErrorData.ErrorType <> "" Then
'        MsgBox "�������"
        Exit Sub
    End If
    
    Interpret AST
    If ErrorData.ErrorType <> "" Then
'        MsgBox "����ʱ����"
        Exit Sub
    End If
End Sub

'���ñ���
Private Sub SetVariable(Value As Variant, Variables() As Variable, ByVal Name As String, Optional ByVal VariableType = "Any")
    Dim i As Long
    
    For i = 1 To UBound(Variables)
        If Variables(i).Name = Name Then
            Variables(i).Value = Value
            Variables(i).Type = VariableType
            Exit Sub
        End If
    Next
    
    ReDim Preserve Variables(UBound(Variables) + 1)
    Variables(UBound(Variables)).Name = Name
    Variables(UBound(Variables)).Value = Value
    Variables(UBound(Variables)).Type = VariableType
End Sub

Private Function Token(Optional ByVal TokenType As String = "Null", Optional ByVal Value As String = "") As Token
    Token.Type = TokenType
    Token.Value = Value
End Function

'��ӽڵ�
Public Sub TreeAddChild(Tree() As Node, ByVal ID As Long, Node As Node, ByVal ChildId As Long)
    Dim NodeId As Long
    
    If TreeFindUnuse(Tree) <> -1 Then
        NodeId = TreeFindUnuse(Tree)
    Else
        ReDim Preserve Tree(UBound(Tree) + 1)
        NodeId = UBound(Tree)
    End If
       
    If UBound(Tree(ID).Children) < ChildId Then ReDim Preserve Tree(ID).Children(ChildId)
    Tree(ID).Children(ChildId) = NodeId
    
    Tree(NodeId) = Node
End Sub

'ɾ����ĳ�ڵ�Ϊ��������
Public Sub TreeDeleteNode(Tree() As Node, ByVal ID As Long)
    Dim i As Long, j As Long
    
    For i = 1 To UBound(Tree)
        For j = 1 To UBound(Tree(i).Children)
            If Tree(i).Children(j) = ID Then NodeDeleteChild Tree(i), j
        Next
    Next
End Sub

'�½�����
Public Sub TreeEmpty(Tree() As Node)
    ReDim Tree(1)
    Tree(1) = NodeNoChild
End Sub

'�½�ֻ��һ���ڵ����
Public Sub TreeCreateWithNode(Tree() As Node, Node As Node)
    ReDim Tree(1)
    Tree(1) = Node
End Sub

'���ر��ΪId�ڵ���ӽڵ�
Public Sub TreeFindSubTreeNodes(Tree() As Node, Nodes() As Boolean, ByVal ID As Long)
    Dim List() As Long
    Dim Current As Long
    Dim i As Long
    
    ReDim Nodes(UBound(Tree))
    ReDim List(1)
    List(1) = ID
    
    Do While UBound(List) >= 1
        '��ӽڵ㵽�ڵ��б�ɾ�������б��еĽڵ�
        Current = List(UBound(List))
        Nodes(Current) = True
        ReDim Preserve List(UBound(List) - 1)
        '����ӽڵ㵽�����б�
        For i = 1 To UBound(Tree(Current).Children)
            If Tree(Current).Children(i) > 0 Then
                ReDim Preserve List(UBound(List) + 1)
                List(UBound(List)) = Tree(Current).Children(i)
            End If
        Next
    Loop
End Sub

'����δռ�ýڵ����
Public Function TreeFindUnuse(Tree() As Node) As Long
    Dim i As Long
    Dim Used() As Boolean
    
    TreeFindSubTreeNodes Tree, Used, 1
    
    For i = 1 To UBound(Tree)
        If Not Used Then
            TreeFindUnuse = i
            Exit Function
        End If
    Next
    TreeFindUnuse = 0
End Function

'���������������������������ID�Žڵ�ĵ�ChildID��
Public Sub TreeMerge(Tree() As Node, SubTree() As Node, ByVal ID As Long, ByVal ChildId As Long)
    Dim LenTree As Long
    Dim i As Long, j As Long
    
    LenTree = UBound(Tree)
    ReDim Preserve Tree(LenTree + UBound(SubTree))
    
    'ָ���������ڵ�
    If UBound(Tree(ID).Children) < ChildId Then ReDim Preserve Tree(ID).Children(ChildId)
    Tree(ID).Children(ChildId) = LenTree + 1
    
    '��������
    For i = 1 To UBound(SubTree)
        Tree(LenTree + i).Type = SubTree(i).Type
        Tree(LenTree + i).Value = SubTree(i).Value
        
        ReDim Tree(LenTree + i).Children(UBound(SubTree(i).Children))
        For j = 1 To UBound(SubTree(i).Children)
            Tree(LenTree + i).Children(j) = SubTree(i).Children(j) + LenTree
        Next
    Next
End Sub

'����תд
Private Function VarTypeToString(Vartype As Long) As String
    Select Case Vartype
    Case vbInteger, vbLong, vbByte
        VarTypeToString = "Integer"
    Case vbSingle, vbDouble, vbCurrency, vbDecimal
        VarTypeToString = "Float"
    Case vbString
        VarTypeToString = "String"
    Case vbBoolean
        VarTypeToString = "Boolean"
    Case vbVariant
        VarTypeToString = "Any"
    Case Else
        VarTypeToString = "Null"
    End Select
End Function

Private Function Visit(AST() As Node, ByVal Position As Long, ByRef Variables() As Variable) As Variant '���ʽڵ�
    On Error GoTo ErrHandler
    Dim i As Long
    
    If ErrorData.ErrorType <> "" Then Exit Function
    
    '����
    If AST(Position).Type = "Integer" Then
        Visit = CLng(AST(Position).Value)
    End If
    
    '������
    If AST(Position).Type = "Float" Then
        Visit = CDbl(AST(Position).Value)
    End If
    
    '����
    If AST(Position).Type = "Boolean" Then
        Visit = CBool(AST(Position).Value)
    End If
    
    '�ַ���
    If AST(Position).Type = "String" Then
        Visit = CStr(AST(Position).Value)
    End If
    
    'һԪ�����
    If AST(Position).Type = "UnaryOp" Then
        Select Case AST(Position).Value
        Case "-"
            Visit = -Visit(AST, AST(Position).Children(1), Variables)
        End Select
    End If
    
    '��Ԫ�����
    If AST(Position).Type = "BinOp" Then
        Select Case AST(Position).Value
        Case "+"
            Visit = Visit(AST, AST(Position).Children(1), Variables) + Visit(AST, AST(Position).Children(2), Variables)
        Case "-"
            Visit = Visit(AST, AST(Position).Children(1), Variables) - Visit(AST, AST(Position).Children(2), Variables)
        Case "*"
            Visit = Visit(AST, AST(Position).Children(1), Variables) * Visit(AST, AST(Position).Children(2), Variables)
        Case "/"
            Visit = Visit(AST, AST(Position).Children(1), Variables) / Visit(AST, AST(Position).Children(2), Variables)
        Case "^"
            Visit = Visit(AST, AST(Position).Children(1), Variables) ^ Visit(AST, AST(Position).Children(2), Variables)
        End Select
    End If
    
    '����
    If AST(Position).Type = "Variable" Then
        Visit = GetVariable(AST(Position).Value, Variables)
    End If
    
    '��ֵ���
    If AST(Position).Type = "Assign" Then
        SetVariable Visit(AST, AST(Position).Children(2), Variables), Variables, CStr(AST(AST(Position).Children(1)).Value)
    End If
    
    '�������
    If AST(Position).Type = "Compond" Then
        For i = 1 To UBound(AST(Position).Children)
            Visit AST, AST(Position).Children(i), Variables
        Next
    End If
    
    'if���
    If AST(Position).Type = "If" Then
        If CBool(Visit(AST, AST(Position).Children(1), Variables)) Then
            Visit AST, AST(Position).Children(2), Variables
        Else
            If UBound(AST(Position).Children) >= 3 Then Visit AST, AST(Position).Children(3), Variables
        End If
    End If
    
    '�������
    If AST(Position).Type = "ProgCall" Then
        FormMain.DoActionSentence AST(AST(Position).Children(1)).Value
    End If
    
    Exit Function
ErrHandler:
    MsgBox "����ʱ����" & Err.Number & "��" & Err.Description
    ErrorData.ErrorType = "RunTimeError"
End Function

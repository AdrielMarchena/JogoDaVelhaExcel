VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmGame 
   Caption         =   "JOGO DA VELHA"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   OleObjectBlob   =   "FrmGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a As Integer
Public b As Integer
Public c As Integer
Public d As Integer
Public e As Integer
Public f As Integer
Public g As Integer
Public h As Integer
Public i As Integer
Public t As String
Public dev As Integer



Private Sub UserForm_Activate()
    dev = 1
    t = ""
End Sub
Public Function TERMINO()
        'para dizer o draw
        If ((Cmd1.Caption) <> t) And ((Cmd2.Caption) <> t) And ((Cmd3.Caption) <> t) And ((Cmd4.Caption) <> t) And ((Cmd5.Caption) <> t) And ((Cmd6.Caption) <> t) And ((Cmd7.Caption) <> t) And ((Cmd8.Caption) <> t) And ((Cmd9.Caption) <> t) Then
             LblWin.Caption = "DRAW"
             LblTime.Caption = ""
             BOTAO
        End If
            
            If dev = 1 Then
                LblTime.Caption = "VEZ DO X"
            Else
                LblTime.Caption = "VEZ DO O"
            End If
            'VERIFICA SE O x GANHOU
            'DIAGONAL
            If a = 1 And e = 1 And i = 1 Then
                 LblWin.Caption = "X WIN"
                 BOTAO
            End If
            If c = 1 And e = 1 And g = 1 Then
                LblWin.Caption = "X WIN"
                BOTAO
            End If
            'HORIZONTAL
            '1 -
            If a = 1 And b = 1 And c = 1 Then
                LblWin.Caption = "X WIN"
                BOTAO
            End If
            '2 -
            If d = 1 And e = 1 And f = 1 Then
                 LblWin.Caption = "X WIN"
                 BOTAO
            End If
            '3 -
            If g = 1 And h = 1 And i = 1 Then
                 LblWin.Caption = "X WIN"
                 BOTAO
            End If
            'VERTICAL
            '1 I
            If a = 1 And d = 1 And g = 1 Then
                 LblWin.Caption = "X WIN"
                 BOTAO
            End If
            ' 2 I
            If b = 1 And e = 1 And h = 1 Then
                 LblWin.Caption = "X WIN"
                 BOTAO
            End If
            ' 3 I
            If c = 1 And f = 1 And i = 1 Then
                 LblWin.Caption = "X WIN"
                 BOTAO
            End If
            
            'VERIFICA SE O o GANHOU
            'DIAGONAL
            If a = 11 And e = 11 And i = 11 Then
                 LblWin.Caption = "O WIN"
                 BOTAO
            End If
            If c = 11 And e = 11 And g = 11 Then
                LblWin.Caption = "O WIN"
                BOTAO
            End If
            'HORIZONTAL
            '1 -
            If a = 11 And b = 11 And c = 11 Then
                LblWin.Caption = "O WIN"
                BOTAO
            End If
            '2 -
            If d = 11 And e = 11 And f = 11 Then
                 LblWin.Caption = "O WIN"
                 BOTAO
            End If
            '3 -
            If g = 11 And h = 11 And i = 11 Then
                 LblWin.Caption = "O WIN"
                 BOTAO
            End If
            'VERTICAL
            '1 I
            If a = 11 And d = 11 And g = 11 Then
                 LblWin.Caption = "O WIN"
                 BOTAO
            End If
            ' 2 I
            If b = 11 And e = 11 And h = 11 Then
                 LblWin.Caption = "O WIN"
                 BOTAO
            End If
            ' 3 I
            If c = 11 And f = 11 And i = 11 Then
                 LblWin.Caption = "O WIN"
                 BOTAO
            End If
                      
fim:
End Function

Public Function BOTAO()
    LblTime.Caption = ""
    Cmd1.Enabled = False
    Cmd2.Enabled = False
    Cmd3.Enabled = False
    Cmd4.Enabled = False
    Cmd5.Enabled = False
    Cmd6.Enabled = False
    Cmd7.Enabled = False
    Cmd8.Enabled = False
    Cmd9.Enabled = False
    
End Function

Private Sub CmdClear_Click()
    'zerando variaveis
     a = 0
     b = 0
     c = 0
     d = 0
     e = 0
     f = 0
     g = 0
     h = 0
     i = 0
     'limpando label de vencedor
     LblWin.Caption = t
     'limpando label da vez da rodada
     LblTime.Caption = ""
     'limpando caption dos botoes
     Cmd1.Caption = t
     Cmd2.Caption = t
     Cmd3.Caption = t
     Cmd4.Caption = t
     Cmd5.Caption = t
     Cmd6.Caption = t
     Cmd7.Caption = t
     Cmd8.Caption = t
     Cmd9.Caption = t
     
    'Habilitando botoes
    Cmd1.Enabled = True
    Cmd2.Enabled = True
    Cmd3.Enabled = True
    Cmd4.Enabled = True
    Cmd5.Enabled = True
    Cmd6.Enabled = True
    Cmd7.Enabled = True
    Cmd8.Enabled = True
    Cmd9.Enabled = True
     
End Sub

Private Sub Cmd1_Click()
    
    If dev = 1 Then
        Cmd1.Caption = "x"
        dev = 2
        a = 1
        
        GoTo fim
    End If
    If dev = 2 Then
        Cmd1.Caption = "o"
        dev = 1
        a = 11
        GoTo fim
    End If
    
fim:
TERMINO
End Sub

Private Sub Cmd2_Click()
    

    If dev = 1 Then
        Cmd2.Caption = "x"
        dev = 2
        b = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd2.Caption = "o"
        dev = 1
        b = 11
        GoTo fim
    End If
    
fim:
    TERMINO
End Sub
Private Sub Cmd3_Click()
    
    If dev = 1 Then
        Cmd3.Caption = "x"
        dev = 2
        c = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd3.Caption = "o"
        dev = 1
        c = 11
        GoTo fim
    End If
fim:
    TERMINO
End Sub
Private Sub Cmd4_Click()
    
    If dev = 1 Then
        Cmd4.Caption = "x"
        dev = 2
        d = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd4.Caption = "o"
        dev = 1
        d = 11
        GoTo fim
    End If
fim:
    TERMINO
End Sub

Private Sub Cmd5_Click()
    
    If dev = 1 Then
        Cmd5.Caption = "x"
        dev = 2
        e = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd5.Caption = "o"
        dev = 1
        e = 11
        GoTo fim
    End If
fim:
    TERMINO
End Sub

Private Sub Cmd6_Click()
    
    If dev = 1 Then
        Cmd6.Caption = "x"
        dev = 2
        f = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd6.Caption = "o"
        dev = 1
        f = 11
        GoTo fim
    End If
fim:
    TERMINO
End Sub

Private Sub Cmd7_Click()
    
    If dev = 1 Then
        Cmd7.Caption = "x"
        dev = 2
        g = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd7.Caption = "o"
        dev = 1
        g = 11
        GoTo fim
    End If
fim:
    TERMINO

End Sub
Private Sub Cmd8_Click()
    
    TERMINO
    If dev = 1 Then
        Cmd8.Caption = "x"
        dev = 2
        h = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd8.Caption = "o"
        dev = 1
        h = 11
        GoTo fim
    End If
fim:
    TERMINO
End Sub

Private Sub Cmd9_Click()
    
    If dev = 1 Then
        Cmd9.Caption = "x"
        dev = 2
        i = 1
        GoTo fim
    End If
    If dev = 2 Then
        Cmd9.Caption = "o"
        dev = 1
        i = 11
        GoTo fim
    End If
fim:
    TERMINO
End Sub



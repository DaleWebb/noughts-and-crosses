Public Class Form1
    Dim GameNumber As Integer = 1
    Dim Player1Streak As Integer = 0
    Dim Player2Streak As Integer = 0
    Dim Player, Winner, P1P1, P1P2, P1P3, P2P1, P2P2, P2P3 As Integer
    Dim Symbol, Command As String
    Dim LastClicked As Label
    Dim AIChoice As Integer
    Dim rand As New Random()
    Dim AIGame As Boolean
    Dim Pos() As Integer = {0, 1, 2, 3, 4, 5, 6, 7, 8}


    'GAMES FUNCTIONS
    Public Sub SymbolSelect()
        Select Case Player
            Case 1
                Symbol = "O"
            Case 2
                Symbol = "X"
        End Select
    End Sub
    Public Sub ChangePlayer()
        If Player = 1 Then
            Player = 2
            PerkSeek()
        Else : Player = 1
            PerkSeek()
        End If
        Label11.Text = "Player " & Player & "'s Turn!"
    End Sub
    Public Sub Result()
        ChangePlayer()
        If Not (Winner = 0) Then
            ChangePlayer()
            Winner = Player
            Label11.Text = "End of Game"
            ListBox1.Items.Add("Game " & GameNumber & ": Player " & Winner & " Wins!")
            Label1.Enabled() = False
            Label2.Enabled() = False
            Label3.Enabled() = False
            Label4.Enabled() = False
            Label5.Enabled() = False
            Label6.Enabled() = False
            Label7.Enabled() = False
            Label8.Enabled() = False
            Label9.Enabled() = False
            Button1.Enabled() = True
            SymbolSelect()
            If Player = 1 Then
                Player1Streak = Player1Streak + 1
                Player2Streak = 0
            Else
                Player2Streak = Player2Streak + 1
                Player1Streak = 0
            End If

            Select Case Player1Streak
                Case 2
                    P1P1 = P1P1 + 1
                    PerkSeek()
                    MessageBox.Show("You have unlocked the 'Time Travel' Perk! The Time Travel Perk allows you to reset the round at any point in the game on your turn", "Perk Unlock")
                Case 3
                    P1P2 = P1P2 + 1
                    PerkSeek()
                    MessageBox.Show("You have unlocked the 'Symbol Sweep' Perk! The 'Symbol Sweep' Perk allows you to sweep the board with your symbol at any point in the game on your turn", "Perk Unlock")
                Case 4
                    P1P3 = P1P3 + 1
                    PerkSeek()
                    MessageBox.Show("You have unlocked the 'Reset Foe's Streak' Perk! The 'Reset Foe's Streak' Perk allows you to reset your oppenent's streak at any point in the game on your turn", "Perk Unlock")
            End Select

            Select Case Player2Streak
                Case 2
                    P2P1 = P2P1 + 1
                    PerkSeek()
                    MessageBox.Show("You have unlocked the 'Time Travel' Perk! The Time Travel Perk allows you to reset the round at any point in the game on your turn", "Perk Unlock")
                Case 3
                    P2P2 = P2P2 + 1
                    PerkSeek()
                    MessageBox.Show("You have unlocked the 'Symbol Sweep' Perk! The 'Symbol Sweep' Perk allows you to sweep the board with your symbol at any point in the game on your turn", "Perk Unlock")
                Case 4
                    P2P3 = P2P3 + 1
                    PerkSeek()
                    MessageBox.Show("You have unlocked the 'Reset Foe's Streak' Perk! The 'Reset Foe's Streak' Perk allows you to reset your oppenent's streak at any point in the game on your turn", "Perk Unlock")
            End Select
        End If
    End Sub
    Public Sub Play()
        If Label1.Text = "X" And Label2.Text = "X" And Label3.Text = "X" Then
            Winner = 2
        ElseIf Label4.Text = "X" And Label5.Text = "X" And Label6.Text = "X" Then
            Winner = 2
        ElseIf Label7.Text = "X" And Label8.Text = "X" And Label9.Text = "X" Then
            Winner = 2
        ElseIf Label1.Text = "X" And Label4.Text = "X" And Label7.Text = "X" Then
            Winner = 2
        ElseIf Label2.Text = "X" And Label5.Text = "X" And Label8.Text = "X" Then
            Winner = 2
        ElseIf Label3.Text = "X" And Label6.Text = "X" And Label9.Text = "X" Then
            Winner = 2
        ElseIf Label1.Text = "X" And Label5.Text = "X" And Label9.Text = "X" Then
            Winner = 2
        ElseIf Label3.Text = "X" And Label5.Text = "X" And Label7.Text = "X" Then
            Winner = 2

        ElseIf Label1.Text = "O" And Label2.Text = "O" And Label3.Text = "O" Then
            Winner = 1
        ElseIf Label4.Text = "O" And Label5.Text = "O" And Label6.Text = "O" Then
            Winner = 1
        ElseIf Label7.Text = "O" And Label8.Text = "O" And Label9.Text = "O" Then
            Winner = 1
        ElseIf Label1.Text = "O" And Label4.Text = "O" And Label7.Text = "O" Then
            Winner = 1
        ElseIf Label2.Text = "O" And Label5.Text = "O" And Label8.Text = "O" Then
            Winner = 1
        ElseIf Label3.Text = "O" And Label6.Text = "O" And Label9.Text = "O" Then
            Winner = 1
        ElseIf Label1.Text = "O" And Label5.Text = "O" And Label9.Text = "O" Then
            Winner = 1
        ElseIf Label3.Text = "O" And Label5.Text = "O" And Label7.Text = "O" Then
            Winner = 1
        ElseIf Not (Label1.Text = "   ") And Not (Label2.Text = "   ") And Not (Label3.Text = "   ") And Not (Label4.Text = "   ") And Not (Label5.Text = "   ") And Not (Label6.Text = "   ") And Not (Label7.Text = "   ") And Not (Label8.Text = "   ") And Not (Label9.Text = "   ") Then
            MessageBox.Show("Draw", "Draw")
            Label10.Text = "Draw!"
            ListBox1.Items.Add("Game " & GameNumber & ": Draw!")
            Button1.Enabled = True
            Player1Streak = 0
            Player2Streak = 0
        End If
    End Sub
    Public Sub Reset()
        Label1.Text = "   "
        Label2.Text = "   "
        Label3.Text = "   "
        Label4.Text = "   "
        Label5.Text = "   "
        Label6.Text = "   "
        Label7.Text = "   "
        Label8.Text = "   "
        Label9.Text = "   "
        Label1.Enabled() = True
        Label2.Enabled() = True
        Label3.Enabled() = True
        Label4.Enabled() = True
        Label5.Enabled() = True
        Label6.Enabled() = True
        Label7.Enabled() = True
        Label8.Enabled() = True
        Label9.Enabled() = True
        Label10.Text = ""
        ChangePlayer()
        Player = 1
        Label11.Text = "Player " + Convert.ToString(Player) + "'s Turn!"
        GameNumber = GameNumber + 1
        Button1.Enabled() = False
        Winner = 0
        PerkSeek()
    End Sub
    'ADDED GAME FUNCTION
    Public Sub PerkSeek()
        If Player = 1 Then
            Player2_Perk1.Enabled = False
            Player2_Perk2.Enabled = False
            Player2_Perk3.Enabled = False
            If P1P1 >= 1 Then
                Player1_Perk1.Enabled = True
            End If
            If P1P2 >= 1 Then
                Player1_Perk2.Enabled = True
            End If
            If P1P3 >= 1 Then
                Player1_Perk3.Enabled = True
            End If

        ElseIf Player = 2 Then
            Player1_Perk1.Enabled = False
            Player1_Perk2.Enabled = False
            Player1_Perk3.Enabled = False
            If P2P1 >= 1 Then
                Player2_Perk1.Enabled = True
            End If
            If P2P2 >= 1 Then
                Player2_Perk2.Enabled = True
            End If
            If P2P3 >= 1 Then
                Player2_Perk3.Enabled = True
            End If
        End If

    End Sub
    'ARTIFICAL INTELLIGENCE
    Public Sub AI()
        Randomize()
        AIChoice = Pos(rand.Next(0, 9))
        Select Case AIChoice
            Case 0
                If Label1.Enabled = True Then
                    SymbolSelect()
                    Label1.Text = Symbol
                    Label1.Enabled = False
                    Play()
                    LastClicked = Label1
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 1
                If Label2.Enabled = True Then
                    SymbolSelect()
                    Label2.Text = Symbol
                    Label2.Enabled = False
                    Play()
                    LastClicked = Label2
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 2
                If Label3.Enabled = True Then
                    SymbolSelect()
                    Label3.Text = Symbol
                    Label3.Enabled = False
                    Play()
                    LastClicked = Label3
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 3
                If Label4.Enabled = True Then
                    SymbolSelect()
                    Label4.Text = Symbol
                    Label4.Enabled = False
                    Play()
                    LastClicked = Label4
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 4
                If Label5.Enabled = True Then
                    SymbolSelect()
                    Label5.Text = Symbol
                    Label5.Enabled = False
                    Play()
                    LastClicked = Label5
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 5
                If Label6.Enabled = True Then
                    SymbolSelect()
                    Label6.Text = Symbol
                    Label6.Enabled = False
                    Play()
                    LastClicked = Label6
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 6
                If Label7.Enabled = True Then
                    SymbolSelect()
                    Label7.Text = Symbol
                    Label7.Enabled = False
                    Play()
                    LastClicked = Label7
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 7
                If Label8.Enabled = True Then
                    SymbolSelect()
                    Label8.Text = Symbol
                    Label8.Enabled = False
                    Play()
                    LastClicked = Label8
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 8
                If Label9.Enabled = True Then
                    SymbolSelect()
                    Label9.Text = Symbol
                    Label9.Enabled = False
                    Play()
                    LastClicked = Label9
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
        End Select
        Result()
    End Sub
    Public Sub AICase()
        Select AIChoice
            Case 0
                If Label1.Enabled = True Then
                    SymbolSelect()
                    Label1.Text = Symbol
                    Label1.Enabled = False
                    Play()
                    LastClicked = Label1
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 1
                If Label2.Enabled = True Then
                    SymbolSelect()
                    Label2.Text = Symbol
                    Label2.Enabled = False
                    Play()
                    LastClicked = Label2
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 2
                If Label3.Enabled = True Then
                    SymbolSelect()
                    Label3.Text = Symbol
                    Label3.Enabled = False
                    Play()
                    LastClicked = Label3
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 3
                If Label4.Enabled = True Then
                    SymbolSelect()
                    Label4.Text = Symbol
                    Label4.Enabled = False
                    Play()
                    LastClicked = Label4
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 4
                If Label5.Enabled = True Then
                    SymbolSelect()
                    Label5.Text = Symbol
                    Label5.Enabled = False
                    Play()
                    LastClicked = Label5
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 5
                If Label6.Enabled = True Then
                    SymbolSelect()
                    Label6.Text = Symbol
                    Label6.Enabled = False
                    Play()
                    LastClicked = Label6
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 6
                If Label7.Enabled = True Then
                    SymbolSelect()
                    Label7.Text = Symbol
                    Label7.Enabled = False
                    Play()
                    LastClicked = Label7
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 7
                If Label8.Enabled = True Then
                    SymbolSelect()
                    Label8.Text = Symbol
                    Label8.Enabled = False
                    Play()
                    LastClicked = Label8
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
            Case 8
                If Label9.Enabled = True Then
                    SymbolSelect()
                    Label9.Text = Symbol
                    Label9.Enabled = False
                    Play()
                    LastClicked = Label9
                Else
                    Randomize()
                    AIChoice = Pos(rand.Next(0, 9))
                    AICase()
                End If
        End Select
    End Sub
    'PERKS
    Public Sub Perk_TimeTravel()
        LastClicked.Enabled = True
        LastClicked.Text = "   "
        ListBox1.Items.Add("Player " & Player & " deleted the oppenents move!")

    End Sub
    Public Sub Perk_EnemyReset()
        ChangePlayer()
        If Player = 1 Then
            Player1Streak = 0
        Else
            Player2Streak = 0
        End If
        ListBox1.Items.Add("Player " & Player & "'s Streak has been ZEROED!")
        ChangePlayer()
    End Sub
    Public Sub Perk_SymbolSweep()
        If Player = 1 Then
            Symbol = "O"
            Label1.Text = Symbol
            Label2.Text = Symbol
            Label3.Text = Symbol
            Label4.Text = Symbol
            Label5.Text = Symbol
            Label6.Text = Symbol
            Label7.Text = Symbol
            Label8.Text = Symbol
            Label9.Text = Symbol
        ElseIf Player = 2 Then
            Symbol = "X"
            Label1.Text = Symbol
            Label2.Text = Symbol
            Label3.Text = Symbol
            Label4.Text = Symbol
            Label5.Text = Symbol
            Label6.Text = Symbol
            Label7.Text = Symbol
            Label8.Text = Symbol
            Label9.Text = Symbol
        End If
        ChangePlayer()
        Play()
    End Sub

    'GRID
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        If AIGame = True Then
            SymbolSelect()
            Label1.Text = Symbol
            Label1.Enabled = False
            Play()
            Result()
            LastClicked = Label1
            If Button1.Enabled = False Then
                AI()
            End If

        Else
            SymbolSelect()
            Label1.Text = Symbol
            Label1.Enabled = False
            Play()
            Result()
            LastClicked = Label1
        End If

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        If AIGame = True Then
            SymbolSelect()
            Label2.Text = Symbol
            Label2.Enabled = False
            Play()
            Result()
            LastClicked = Label2
            If Button1.Enabled = False Then
                AI()
            End If

        Else
            SymbolSelect()
            Label2.Text = Symbol
            Label2.Enabled = False
            Play()
            Result()
            LastClicked = Label2
        End If

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        If AIGame = True Then
            SymbolSelect()
            Label3.Text = Symbol
            Label3.Enabled = False
            Play()
            Result()
            LastClicked = Label3
            If Button1.Enabled = False Then
                AI()
            End If

        Else
            SymbolSelect()
            Label3.Text = Symbol
            Label3.Enabled = False
            Play()
            Result()
            LastClicked = Label3
        End If

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        If AIGame = True Then
            SymbolSelect()
            Label4.Text = Symbol
            Label4.Enabled = False
            Play()
            Result()
            LastClicked = Label4
            If Button1.Enabled = False Then
                AI()
            End If

        Else
            SymbolSelect()
            Label4.Text = Symbol
            Label4.Enabled = False
            Play()
            Result()
            LastClicked = Label4
        End If

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        If AIGame = True Then
            SymbolSelect()
            Label5.Text = Symbol
            Label5.Enabled = False
            Play()
            Result()
            LastClicked = Label5
            If Button1.Enabled = False Then
                AI()
            End If
        Else
            SymbolSelect()
            Label5.Text = Symbol
            Label5.Enabled = False
            Play()
            Result()
            LastClicked = Label5
        End If

    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        If AIGame = True Then
            SymbolSelect()
            Label6.Text = Symbol
            Label6.Enabled = False
            Play()
            Result()
            LastClicked = Label6
            If Button1.Enabled = False Then
                AI()
            End If

        Else
            SymbolSelect()
            Label6.Text = Symbol
            Label6.Enabled = False
            Play()
            Result()
            LastClicked = Label6
        End If

    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        If AIGame = True Then
            SymbolSelect()
            Label7.Text = Symbol
            Label7.Enabled = False
            Play()
            Result()
            LastClicked = Label7
            If Button1.Enabled = False Then
                AI()
            End If

        Else
            SymbolSelect()
            Label7.Text = Symbol
            Label7.Enabled = False
            Play()
            Result()
            LastClicked = Label7
        End If

    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        If AIGame = True Then
            SymbolSelect()
            Label8.Text = Symbol
            Label8.Enabled = False
            Play()
            Result()
            LastClicked = Label8
            If Button1.Enabled = False Then
                AI()
            End If

        Else
            SymbolSelect()
            Label8.Text = Symbol
            Label8.Enabled = False
            Play()
            Result()
            LastClicked = Label8
        End If

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        If AIGame = True Then
            SymbolSelect()
            Label9.Text = Symbol
            Label9.Enabled = False
            Play()
            Result()
            LastClicked = Label9
            If Button1.Enabled = False Then
                AI()
            End If
        Else
            SymbolSelect()
            Label9.Text = Symbol
            Label9.Enabled = False
            Play()
            Result()
            LastClicked = Label9
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Reset()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If MessageBox.Show("AI Game?", "Game Choice", MessageBoxButtons.YesNo) = 6 Then
            AIGame = True
            MessageBox.Show("AI Game", "AI Game")
        End If
        Player = 1
        SymbolSelect()
    End Sub

    'PERK BUTTONS

    Private Sub Player1_Perk1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Player1_Perk1.Click
        If (Winner = 0) And Not (Label1.Text = "   " And Label2.Text = "   " And Label3.Text = "   " _
                            And Label4.Text = "   " And Label5.Text = "   " And Label6.Text = "   " And _
                            Label1.Text = "   " And Label2.Text = "   " And Label3.Text = "   ") Then
            Perk_TimeTravel()
            P1P1 = P1P1 - 1
            PerkSeek()
            If P1P1 < 1 Then
                Player1_Perk1.Enabled = False
            End If
        Else : MessageBox.Show("You cannot use this Perk at this point in the game", "Perk Error")
        End If
    End Sub

    Private Sub Player2_Perk1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Player2_Perk1.Click
        If (Winner = 0) And Not (Label1.Text = "   " And Label2.Text = "   " And Label3.Text = "   " _
                            And Label4.Text = "   " And Label5.Text = "   " And Label6.Text = "   " And _
                            Label1.Text = "   " And Label2.Text = "   " And Label3.Text = "   ") Then
            Perk_TimeTravel()
            P2P1 = P2P1 - 1
            PerkSeek()
            If P2P1 < 1 Then
                Player2_Perk1.Enabled = False
            End If
        Else : MessageBox.Show("You cannot use this Perk at this point in the game", "Perk Error")
        End If
    End Sub

    Private Sub Player1_Perk2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Player1_Perk2.Click
        Perk_SymbolSweep()
        ChangePlayer()
        Result()
        P1P2 = P1P2 - 1
        PerkSeek()
        If P1P2 < 1 Then
            Player1_Perk2.Enabled = False
        End If
    End Sub

    Private Sub Player2_Perk2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Player2_Perk2.Click
        Perk_SymbolSweep()
        ChangePlayer()
        Result()
        P2P2 = P2P2 - 1
        PerkSeek()
        If P2P2 < 1 Then
            Player2_Perk2.Enabled = False
        End If
    End Sub

    Private Sub Player1_Perk3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Player1_Perk3.Click
        Perk_EnemyReset()
        P1P3 = P1P3 - 1
        PerkSeek()
        If P1P3 < 1 Then
            Player1_Perk3.Enabled = False
        End If
    End Sub

    Private Sub Player2_Perk3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Player2_Perk3.Click
        Perk_EnemyReset()
        P2P3 = P2P3 - 1
        PerkSeek()
        If P2P3 < 1 Then
            Player2_Perk3.Enabled = False
        End If
    End Sub
End Class

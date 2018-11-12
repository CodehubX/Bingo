Attribute VB_Name = "mdlBingo"
'Programmer: Alex Quach
'Date: June 7/2016
'Purpose: To create the game of bingo
Option Explicit
Const TEMPMAX = 15
Const LOW = 1
Global GameStatus As Boolean
Global Const HIGHSCOREMAX = 5
Global Const FNAME = "HIGHSCORES.txt"
Global HighestPlayer(1 To HIGHSCOREMAX) As String
Global HighestScore(1 To HIGHSCOREMAX) As Integer
Public Sub GetBingoNum(PNum() As Integer)
    
    Dim K As Integer, Counter As Integer, L As Integer
    Dim Num1 As Integer, Num2 As Integer, TempSwap As Integer, TempNums(1 To TEMPMAX) As Integer
    Dim TempHigh As Integer, TempLow As Integer
        
    Counter = 0
    TempHigh = 15
    TempLow = 1
    
    'Loop it by column (5 columns in total)
    For K = 1 To 5
    
        'Intializes the numbers to randomize
        For L = TempLow To TempHigh
        
            Counter = Counter + 1
            TempNums(Counter) = L
            
        Next L
        'Resets counter
        Counter = 0
        'Randomizes the 15 numbers generated
        For L = 1 To 1000
        
            Num1 = Int(Rnd() * (TEMPMAX - LOW + 1)) + LOW
            Num2 = Int(Rnd() * (TEMPMAX - LOW + 1)) + LOW
            TempSwap = TempNums(Num1)
            TempNums(Num1) = TempNums(Num2)
            TempNums(Num2) = TempSwap
            
        Next L
        'Inputs random numbers into array by column
        For L = 0 To 20 Step 5
        
            Counter = Counter + 1
            If K + L <> 13 Then
            
                PNum(K + L) = TempNums(Counter)
                
            End If
            
        Next L
        Counter = 0
        'Offsets the high and lows for the next number generation
        TempHigh = TempHigh + 15
        TempLow = TempLow + 15
        
    Next K
    
End Sub
Public Sub GetWinningNum(WNum() As Integer, ByVal WHigh As Integer)
    
    Dim K As Integer, L As Integer
    Dim Num1 As Integer, Num2 As Integer, TempSwap As Integer
    
    'Randomizes the array that was passed in
    For K = 1 To 1000
    
        Num1 = Int(Rnd() * (WHigh - LOW + 1)) + LOW
        Num2 = Int(Rnd() * (WHigh - LOW + 1)) + LOW
        TempSwap = WNum(Num1)
        WNum(Num1) = WNum(Num2)
        WNum(Num2) = TempSwap
        
    Next K
  
End Sub
Public Sub CoverBoard(ByVal X As Integer, ByVal L As Integer, Count() As Integer, Count2() As Integer)
    
    Dim K As Integer
    'Covers the two bingo boards
    For K = 1 To X
        frmMain.imgCover(K).Picture = frmMain.imgPHold1.Picture
        Count(K) = 0
        Count2(K) = 0
        frmMain.lblBingoNum(K).BackColor = &H8000000F
        frmMain.imgCover3(K).Picture = frmMain.imgPHold1.Picture
        frmMain.lblBingoNum2(K).BackColor = &H8000000F
    Next K
    'Covers the winning numbers board
    For K = 1 To L
    
        frmMain.imgCover2(K).Visible = True
        frmMain.imgCover2(K).Picture = frmMain.imgPHold2.Picture
        frmMain.imgCover2(K).Enabled = True
        frmMain.lblWinNum(K).BackColor = &H8000000F
        
    Next K
    'Disables the form so the user may not click the board before the
    'game starts
    frmMain.fraWinningNum.Enabled = False
    
End Sub
Public Sub MatchCheck(Count() As Integer, Count2() As Integer, ByVal W As Integer, BNum() As Integer, BNum2() As Integer, WNum() As Integer, ByVal I As Integer)
    
    Dim K As Integer
    
    'Sets the middle of the bingo boards to 1 as the middle
    'is a free match
    Count(13) = 1
    Count2(13) = 1
    
    'Checks to see if there is a match
    For K = 1 To W
    
        If BNum(K) = WNum(I) Then
        
            'Changes colour if there is a match
            frmMain.lblBingoNum(K).BackColor = &H80FF&
            frmMain.lblWinNum(I).BackColor = &H80FF&
            'Changes the win count value to 1
            '1 represents a match
            Count(K) = 1
            
        End If
        'Checks the second board
        If BNum2(K) = WNum(I) Then
        
            'Changes colour if there is a match
            frmMain.lblBingoNum2(K).BackColor = &H80FF&
            frmMain.lblWinNum(I).BackColor = &H80FF&
            'Changes the win count value to 1
            '1 represents a match
            Count2(K) = 1
            
        End If
        
    Next K
    
End Sub
Public Sub WinCheck(Count() As Integer, Total As Integer, Count2() As Integer)
    
    Dim K As Integer
    Dim C As Integer
    Dim TempTotal As Integer
    Dim TempTotal2 As Integer
    
    TempTotal = 0
    TempTotal2 = 0
    
    'Column checking
    For K = 1 To 5
        'Goes down the column on the first board to check if there is a win
        For C = 0 To 20 Step 5
            'Adds 1 to the win total
            If Count(K + C) = 1 Then
            
                TempTotal = TempTotal + 1
                
            End If
            
        Next C
        'Goes down the coloumn on the second board to check if there is a win
        For C = 0 To 20 Step 5
            'Adds 1 to the win total
            If Count2(K + C) = 1 Then
            
                TempTotal2 = TempTotal2 + 1
                
            End If
            
        Next C
        'Checks if there is a win on the first board
        If TempTotal = 5 Then
            'Changes colour of the winning column if there is a win on the
            'first board
            For C = 0 To 20 Step 5
            
                frmMain.lblBingoNum(K + C).BackColor = QBColor(9)
                
            Next C
            Total = TempTotal
            
        End If
        'Checks if there is a win on the second board
        If TempTotal2 = 5 Then
            'Changes colour of the winning column if there is a win on the
            'second board
            For C = 0 To 20 Step 5
            
                frmMain.lblBingoNum2(K + C).BackColor = QBColor(9)
                
            Next C
            Total = TempTotal2
            
        End If
        'Sets the win total for both boards to 0 for the next round of checking
        TempTotal = 0
        TempTotal2 = 0
        
    Next K
    
    'Row checking
    For C = 1 To 21 Step 5
        'Goes across the row on the first board
        For K = 0 To 4
            'Adds 1 to the win total
            If Count(K + C) = 1 Then
            
                TempTotal = TempTotal + 1
                
            End If
            
        Next K
        'Goes across the row on the second board
        For K = 0 To 4
            'Adds 1 to the win total
            If Count2(K + C) = 1 Then
            
                TempTotal2 = TempTotal2 + 1
                
            End If
            
        Next K
        'Checks if there is a win on the first board
        If TempTotal = 5 Then
            'Changes colour of the winning row
            For K = 0 To 4
            
                frmMain.lblBingoNum(K + C).BackColor = QBColor(9)
                
            Next K
            'Displays the appropriate messages if the
            'user has won
            Total = TempTotal
            
        End If
        'Checks if there is a win on the second board
        If TempTotal2 = 5 Then
            'Changes colour of the winning row
            For K = 0 To 4
            
                frmMain.lblBingoNum2(K + C).BackColor = QBColor(9)
                
            Next K
            'Displays the appropriate messages if the
            'user has won
            Total = TempTotal2
            
        End If
        'Sets the win total to 0 for the next round of checking
        TempTotal = 0
        TempTotal2 = 0
        
    Next C
    
    'Diagonal checking
    For C = 1 To 5 Step 4
        'Checks the first diagonal on the two boards
        If C = 1 Then
            'Goes down the diagonal on the first board
            'and adds 1 to the win total
            TempTotal = Count(C) + Count(C + 6) + Count(C + 12) + Count(C + 18) + Count(C + 24)
            'Checks if there is a win on the first board
            If TempTotal = 5 Then
                'Changes the colour of the winning diagonal row
                For K = 0 To 24 Step 6
                    
                    frmMain.lblBingoNum(K + C).BackColor = QBColor(9)
                    
                Next K
                Total = TempTotal
                
            End If
            'Goes down the diagonal on the second board
            'and adds 1 to the win total
            TempTotal2 = Count2(C) + Count2(C + 6) + Count2(C + 12) + Count2(C + 18) + Count2(C + 24)
            'Checks if there is a win on the second board
            If TempTotal2 = 5 Then
                'Changes the colour of the winning diagonal row
                For K = 0 To 24 Step 6
                    
                    frmMain.lblBingoNum2(K + C).BackColor = QBColor(9)
                    
                Next K
                Total = TempTotal2
                
            End If
        'Checks the second diagonal on the two boards
        ElseIf C = 5 Then
            'Goes down the diagonal on the first board
            'and adds 1 to the win total
            TempTotal = Count(C) + Count(C + 4) + Count(C + 8) + Count(C + 12) + Count(C + 16)
            'Checks if there is a win
            If TempTotal = 5 Then
                'Changes the colour of the winning diagonal row
                For K = 0 To 17 Step 4
                    
                    frmMain.lblBingoNum(K + C).BackColor = QBColor(9)
                    
                Next K
                Total = TempTotal
            End If
            'Goes down the diagonal on the second board
            'and adds 1 to the win total
            TempTotal2 = Count2(C) + Count2(C + 4) + Count2(C + 8) + Count2(C + 12) + Count2(C + 16)
            'Checks if there is a win on the second board
            If TempTotal2 = 5 Then
                'Changes the colour of the winning diagonal row
                For K = 0 To 17 Step 4
                    
                    frmMain.lblBingoNum2(K + C).BackColor = QBColor(9)
                    
                Next K
                Total = TempTotal2
            End If
            
        End If
        'Sets the win total to 0 for the next round of checking
        TempTotal = 0
        TempTotal2 = 0
        
    Next C
    
End Sub
Public Sub NewGame(BNum() As Integer, WNum() As Integer, WMax As Integer, BMax As Integer, WCount() As Integer, WCount2() As Integer, BNum2() As Integer)

    'Generates the bingo numbers, winning numbers, and covers the board
    GetBingoNum BNum()
    GetBingoNum BNum2()
    GetWinningNum WNum(), WMax
    CoverBoard BMax, WMax, WCount(), WCount2()
    DisplayNumbers BNum(), BNum2, WNum(), WMax

End Sub

Public Sub ReEnable(WinTot As Integer, C As Integer)
    
    'Re-enables menu items and intializes the win checking variables
    frmMain.mnuNewGame.Enabled = False
    frmMain.mnuReveal.Enabled = True
    WinTot = 0
    C = 0

End Sub

Public Sub DisplayNumbers(BNum() As Integer, BNum2() As Integer, WNum() As Integer, ByVal WHigh As Integer)
    
    Dim K As Integer
    Dim L As Integer
    Dim Counter As Integer
    
    Counter = 0
    
    'Displays numbers in the first board
    For K = 1 To 5
        'Displays by column
        For L = 0 To 20 Step 5
            'Makes sure nothing is displayed in the middle
            Counter = Counter + 1
            If K + L <> 13 Then
            
                frmMain.lblBingoNum(K + L).Caption = BNum(K + L)
                
            End If
            
        Next L
        Counter = 0
        
    Next K
    
    'Displays numbers in the second board
    For K = 1 To 5
        'Displays by column
        For L = 0 To 20 Step 5
            'Makes sure nothing is displayed in the middle
            Counter = Counter + 1
            If K + L <> 13 Then
            
                frmMain.lblBingoNum2(K + L).Caption = BNum2(K + L)
                
            End If
            
        Next L
        Counter = 0
        
    Next K
    
    'Displays numbers in the winning numbers board
    For L = 1 To WHigh
    
        frmMain.lblWinNum(L).Caption = WNum(L)
        
    Next L

End Sub

Public Sub Createfile()

    'Creates the file name if the file is not already created
    Dim K As Integer
    
    Open App.Path & "\" & FNAME For Output As #1
    
    'Writes default information to the file
    For K = 1 To HIGHSCOREMAX
    
        Write #1, "Anonymous", 999
    
    Next K

End Sub

Public Sub ReadFile()

    Dim X As Integer
    
    X = 0
    
    'Reads the current high score information in the file
    Open App.Path & "\" & FNAME For Input As #1
    Do While Not EOF(1)
        X = X + 1
        Input #1, HighestPlayer(X)
        Input #1, HighestScore(X)
        
    Loop
    Close #1

End Sub

Public Sub DisplayFile()

    Dim K As Integer

    'Displays the high score information that was read from the
    'high score file
    For K = 1 To HIGHSCOREMAX
    
        If Len(HighestPlayer(K)) > 16 Then
        
            HighestPlayer(K) = Left$(HighestPlayer(K), 16) & "..."
        
        End If
        frmHighScores.picData.Print HighestPlayer(K); Tab(25); Format$(HighestScore(K), "@@@@")
    
    Next K
    
End Sub

Public Sub WriteFile()

    Dim K As Integer
    
    'Writes the high scores to the file
    Open App.Path & "\" & FNAME For Output As #1
    For K = 1 To HIGHSCOREMAX
    
        Write #1, HighestPlayer(K)
        Write #1, HighestScore(K)
        
    Next K
    Close #1

End Sub

Public Sub ChangeHighScore(S As Integer, N As String)

    Dim K As Integer
    Dim TempName As String
    Dim TempNum As Integer
    
    For K = 1 To HIGHSCOREMAX
    
        If S < HighestScore(K) Then
            'Swaps the highest score
            TempNum = S
            S = HighestScore(K)
            HighestScore(K) = TempNum
            'Swaps the highest name
            TempName = N
            N = HighestPlayer(K)
            HighestPlayer(K) = TempName
            
        End If
    
    Next K
    
    'Writes the new highscores to the highscore file
    Open App.Path & "\" & FNAME For Output As #1
    For K = 1 To HIGHSCOREMAX
    
        Write #1, HighestPlayer(K)
        Write #1, HighestScore(K)
        
    Next K
    Close #1
    
End Sub


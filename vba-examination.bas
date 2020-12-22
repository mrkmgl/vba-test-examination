Attribute VB_Name = "Module2"
Option Explicit
Sub Derivatives_Exam()

Dim ansL1(1 To 11), ansL2(1 To 11), ansL3(1 To 7), qL1(1 To 11), qL2(1 To 11), qL3(1 To 7), ansUser As String
Dim score, wrong, i, j, exact, k, qtotal, q1total, q2total, q3total As Integer

Sheets("Sheet1").Range("B10", "C24").Clear 'to clear the cells where we will store the questions/answers

'Defining the Level 1 questions
qL1(1) = "What is the derivative of f = 2x^3+4x^2-5x+7"
qL1(2) = "What is the derivative of f = -7+5x-5/2x^2-3x^4"
qL1(3) = "What is the derivative of f = (-x+7)^4?"
qL1(4) = "What is the derivative of f = (x^2-3)^5"
qL1(5) = "What is the derivative of f = 1/(7x+2)^3?"
qL1(6) = "What is the derivative of f = 1/5x^(5/2)-1/3x^(3/2)?"
qL1(7) = "What is the derivative of f = (x^2-5)^(7/2)?"
qL1(8) = "What is the derivative of f = 3x^2+3x"
qL1(9) = "What is the derivative of f = x^2+1?"
qL1(10) = "What is the derivative of f = exp[-2x]?"
qL1(11) = "What is the derivative of f = x+2/x-2"

'Defining the Level 1 answers
ansL1(1) = "6x^2+8x-5"
ansL1(2) = "5-5x-12x^3"
ansL1(3) = "-4(7-x)^3"
ansL1(4) = "10x(x^2-3)^4"
ansL1(5) = "-21/(7x+2)^4"
ansL1(6) = "1/2(x^(3/2)-x^(1/2))"
ansL1(7) = "7x(x^2-5)^(5/2)"
ansL1(8) = "3(2x+1)"
ansL1(9) = "2x"
ansL1(10) = "-2exp[-2x]"
ansL1(11) = "-4/(x-2)^2"

'Defining the Level 2 questions
qL2(1) = "What is the derivative of f = cos(2x)Ð2sin(x)"
qL2(2) = "What is the derivative of f = tan(x)+(1/3)tan^3(x)"
qL2(3) = "What is the derivative of f = cos(x)-(1/3)cos^3(x)?"
qL2(4) = "What is the derivative of f = 1/cos^n(x)?"
qL2(5) = "What is the derivative of f = sin(x)+cos(x)?"
qL2(6) = "What is the derivative of f = cos(2)sin(x)?"
qL2(7) = "What is the derivative of f = cos(1/x)?"
qL2(8) = "What is the derivative of f = sin^3(x)+cos^3(x)?"
qL2(9) = "What is the derivative of f = tan(x/2)-cot(x/2)?"
qL2(10) = "What is the derivative of f = tan(x^2)-cot(x^2)?"
qL2(11) = "What is the derivative of f = tan(x/2)-cot(x/2)?"

'Defining the Level 2 answers
ansL2(1) = "-2cos(x)(sin(x)+1)"
ansL2(2) = "1+tan^2(x)/cos(2x)"
ansL2(3) = "-sin^3(x)"
ansL2(4) = "nsin(x)cos^(n+1)(x)"
ansL2(5) = "1/1+cos(x)"
ansL2(6) = "-sin((2)sin(x))cos(x)"
ansL2(7) = "1/x^2sin(1/x)"
ansL2(8) = "3/2sin(2x)(sin(x)-cos(x))"
ansL2(9) = "2/sin^2(x)"
ansL2(10) = "2tan^3(x)"
ansL2(11) = "2/sin^2(x)"

'Defining the Level 3 questions
qL3(1) = "What is the derivative of f = sinh(x)"
qL3(2) = "What is the derivative of f = cosh(x)"
qL3(3) = "What is the derivative of f = cosh(x)^(1/2)?"
qL3(4) = "What is the derivative of f = tanh(x)"
qL3(5) = "What is the derivative of f = coth(x)?"
qL3(6) = "What is the derivative of f = sech(x)?"
qL3(7) = "What is the derivative of f = csch(x)?"

'Defining the Level 3 answers
ansL3(1) = "cosh(x)"
ansL3(2) = "sinh(x)"
ansL3(3) = "(sinh(x)^(1/2))/(2(x)^(1/2))"
ansL3(4) = "1-tanh(x)^2"
ansL3(5) = "1-coth(x)^2"
ansL3(6) = "-sech(x)tanh(x)"
ansL3(7) = "-coth(x)csch(x)"

'We start the game by displaying the welcome message, instructions and formatting conditions.
MsgBox "Welcome to the Derivative Practice exam. We will test your preparation in derivatives. The exam is composed of 15 questions and 3 difficulty level. You must answer right 4 times in a row to pass to the next level. You cannot reply wrongly to more than 7 questions. Good luck!"
MsgBox "Before starting, the formatting is important! Please, check the rules written in this Sheet!"
wrong = 0

'Level 1 framework
'We create the first loop in the range of the maximum number of questions assuming the max wrong answers.
For i = 1 To 11
    
    'the user will be asked to reply to the questions
    ansUser = InputBox(qL1(i), "Level 1")
    
    'at each node we count the total number of questions, since the game stops at 15 questions
    'at each node we also count the total number of questions asked during the specific level (just for coding purpose*)
    qtotal = qtotal + 1
    q1total = q1total + 1
    
        'if the given answer is the same as the one in the array, the score updates, as well as the number of exact answers
        'the "exact" variable is used to store the exact answers until it arrives to 4 (4 in a row to pass to the next level). It becomes 0 if a wrong answer is registered
        If ansUser = ansL1(i) Then
            exact = exact + 1
            score = score + 1
        Else:
            'if wrong answer, we start again the counting of exact answer (4 in a row to reach the next level)
            'we update the number of wrong answer (max 7 wrong answers)
            'we also display a message to warn the user
            exact = 0
            wrong = wrong + 1
            MsgBox "Wrong answer!"
        End If
        
        If wrong = 7 Then
            'in case the wrong answers are 7, the game is terminated.
            MsgBox "You need to review your derivatives course, otherwise forget your internship at Goldman!"
            'the user will see displayed all the questions made with the right answers.
            'We use transpose because the array is a row vector, and we want to put it in column.
            '*here the q1total above indicated is used for diplaying the question/answer in the right cell number.
            Sheets("Sheet1").Range("B10", "B" & (9 + q1total)).Value = Application.Transpose(qL1)
            Sheets("Sheet1").Range("C10", "C" & (9 + q1total)).Value = Application.Transpose(ansL1)
            
            Exit Sub
        End If
        
        'Level 2 framework
        'If the user answers right 4 times in a row, it passes to the next level.
        'the "exact" variable is set to zero again, to restart the counting for the next level.
        'we also show a message to congratulate with the user.
        If exact = 4 Then
            MsgBox "You are doing really good! You passed to the next level!"
            exact = 0
            
            'We start the second loop going from 1 to 11 - wrong answers.
            'Even if not necessary, we point out (11-wrong) to make the code lighter.
            For j = 1 To (11 - wrong)
                
                'the user will be asked to reply to the questions
                ansUser = InputBox(qL2(j), "Level 2")
                
                'at each node we count the total number of questions, since the game stops at 15 questions
                'at each node we also count the total number of questions asked during the specific level (just for coding purpose*)
                qtotal = qtotal + 1
                q2total = q2total + 1
                
                    'if the given answer is the same as the one in the array, the score updates, as well as the number of exact answers
                    'the "exact" variable is used to store the exact answers until it arrives to 4 (4 in a row to pass to the next level). It becomes 0 if a wrong answer is registered
                    If ansUser = ansL2(j) Then
                        exact = exact + 1
                        score = score + 1
                    Else:
                        'if wrong answer, we start again the counting of exact answer (4 in a row to reach the next level)
                        'we update the number of wrong answer (max 7 wrong answers)
                        'we also display a message to warn the user
                        exact = 0
                        wrong = wrong + 1
                        MsgBox "Wrong answer!"
                    End If
                                
                        If wrong = 7 Then
                           'in case the wrong answers are 7, the game is terminated.
                           qtotal = 0 'just to avoid overlapping if qtotal = 15, see below's condition
                           'the user will see displayed all the questions made with the right answers.
                           'We use transpose because the array is a row vector, and we want to put it in column.
                           '*here the q2total above indicated is used for diplaying the question/answer in the right cell number.
                           MsgBox "You need to review your derivatives course, otherwise forget your internship at Goldman!"
                           Sheets("Sheet1").Range("B10", "B" & (9 + q1total)).Value = Application.Transpose(qL1)
                           Sheets("Sheet1").Range("C10", "C" & (9 + q1total)).Value = Application.Transpose(ansL1)
                           Sheets("Sheet1").Range("B" & 9 + q1total + 1, "B" & (9 + q1total + q2total)).Value = Application.Transpose(qL2)
                           Sheets("Sheet1").Range("C" & 9 + q1total + 1, "C" & (9 + q1total + q2total)).Value = Application.Transpose(ansL2)
                           Exit For
                        End If
                        
                        'if the total number of questions is 15, the game finishes.
                        'The user receives a message that the game is terminated + the overall score.
                        
                        If qtotal = 15 Then
                                                
                            'the user will see displayed all the questions made with the right answers.
                            'We use transpose because the array is a row vector, and we want to put it in column.
                            '*here the q2total above indicated is used for diplaying the question/answer in the right cell number.
                            MsgBox "The exam is terminated. Your score is " & score
                            Sheets("Sheet1").Range("B10", "B" & (9 + q1total)).Value = Application.Transpose(qL1)
                            Sheets("Sheet1").Range("C10", "C" & (9 + q1total)).Value = Application.Transpose(ansL1)
                            Sheets("Sheet1").Range("B" & 9 + q1total + 1, "B" & (9 + q1total + q2total)).Value = Application.Transpose(qL2)
                            Sheets("Sheet1").Range("C" & 9 + q1total + 1, "C" & (9 + q1total + q2total)).Value = Application.Transpose(ansL2)
                            Exit For
                            
                        End If
                        
                        'Level 3 framework
                        'If the user answers right 4 times in a row, it passes to the next level.
                        'the "exact" variable is not set to zero as before, because there is no next level.
                        'we also show a message to congratulate with the user.
                        If exact = 4 Then
                            MsgBox "You are doing really good! You passed to the last level!"
                            
                            'We start the second loop going from 1 to 7 - wrong answers.
                            'Even if not necessary, we point out (7-wrong) to make the code lighter.
                            For k = 1 To (7 - wrong)
                            
                            'the user will be asked to reply to the questions
                            ansUser = InputBox(qL3(k), "Level 3")
                            
                            'at each node we count the total number of questions, since the game stops at 15 questions
                            'at each node we also count the total number of questions asked during the specific level (just for coding purpose*)
                            qtotal = qtotal + 1
                            q3total = q3total + 1
                            
                                'if the given answer is the same as the one in the array, the score updates, as well as the number of exact answers
                                'the "exact" variable is not accounted anymore (there is no next level)
                                If ansUser = ansL3(k) Then
                                    score = score + 1
                                Else:
                                    'if the answer is wrong, update the number of wrong answer (max 7 wrong answers)
                                    'we also display a message to warn the user
                                    wrong = wrong + 1
                                    MsgBox "Wrong answer!"
                                End If
                                
                                    'in case the wrong answers are 7, the game is terminated.
                                    If wrong = 7 Then
                                    
                                        qtotal = 0 'just to avoid overlapping if qtotal = 15, see below's condition
                                        'the user will see displayed all the questions made with the right answers.
                                        'We use transpose because the array is a row vector, and we want to put it in column.
                                        '*here the q3total above indicated is used for diplaying the question/answer in the right cell number.
                                        MsgBox "You need to review your derivatives course, otherwise forget your internship at Goldman!"
                                        Sheets("Sheet1").Range("B10", "B" & (9 + q1total)).Value = Application.Transpose(qL1)
                                        Sheets("Sheet1").Range("C10", "C" & (9 + q1total)).Value = Application.Transpose(ansL1)
                                        Sheets("Sheet1").Range("B" & 9 + q1total + 1, "B" & (9 + q1total + q2total)).Value = Application.Transpose(qL2)
                                        Sheets("Sheet1").Range("C" & 9 + q1total + 1, "C" & (9 + q1total + q2total)).Value = Application.Transpose(ansL2)
                                        Sheets("Sheet1").Range("B" & 9 + q1total + q2total + 1, "B" & (9 + q1total + q2total + q3total)).Value = Application.Transpose(qL3)
                                        Sheets("Sheet1").Range("C" & 9 + q1total + q2total + 1, "C" & (9 + q1total + q2total + q3total)).Value = Application.Transpose(ansL3)
                                        Exit For
                                    
                                    End If
                                    
                                        'if the total number of questions is 15, the game finishes.
                                        'The user receives a message that the game is terminated + the overall score.
                                        If qtotal = 15 Then
                                        'the user will see displayed all the questions made with the right answers.
                                        'We use transpose because the array is a row vector, and we want to put it in column.
                                        '*here the q3total above indicated is used for diplaying the question/answer in the right cell number.
                                        Sheets("Sheet1").Range("B10", "B" & (9 + q1total)).Value = Application.Transpose(qL1)
                                        Sheets("Sheet1").Range("C10", "C" & (9 + q1total)).Value = Application.Transpose(ansL1)
                                        Sheets("Sheet1").Range("B" & 9 + q1total + 1, "B" & (9 + q1total + q2total)).Value = Application.Transpose(qL2)
                                        Sheets("Sheet1").Range("C" & 9 + q1total + 1, "C" & (9 + q1total + q2total)).Value = Application.Transpose(ansL2)
                                        Sheets("Sheet1").Range("B" & 9 + q1total + q2total + 1, "B" & (9 + q1total + q2total + q3total)).Value = Application.Transpose(qL3)
                                        Sheets("Sheet1").Range("C" & 9 + q1total + q2total + 1, "C" & (9 + q1total + q2total + q3total)).Value = Application.Transpose(ansL3)
                                        
                                        MsgBox "The exam is terminated. Your score is " & score
                                        Exit For
                            
                                        End If
                                        'if the condition is not met, we also show the current number of wrong answers.
                                        MsgBox "Your wrong answers are " & wrong
                              
                              Next k
                              Exit For
                              
                        End If
                        'if the condition is not met, we also show the current number of wrong answers.
                        MsgBox "Your wrong answers are " & wrong
                        
                Next j
                Exit For

        End If
        'if the condition is not met, we also show the current number of wrong answers.
        MsgBox "Your wrong answers are " & wrong
Next i

End Sub


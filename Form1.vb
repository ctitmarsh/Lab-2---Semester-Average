Option Strict On
Public Class frmSemesterGrades
    ' Net Develeopment - Lab 2
    ' By: Christian Titmarsh
    ' ID: 100527601
    ' This program is designed to take 6 numeric inputs within range for 6 courses and output 6 Letter grades, 
    ' semester average In percentage And Grade value.
#Region "Variable and constant Declaration"

    Const MINIMUM_GRADE As Double = 0.0
    Const MAXIMUM_GRADE As Double = 100.0
    Const ERROR_MESSAGE As String = "Invalid input, must be a number within 0 and 100, please try again"

    Dim grade As Double
    Dim letterGrade As String
    Dim grades(5) As Double
    Private gradeToLetter As Double

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
#End Region

#Region "Functions and subs"
    ''' <summary>
    '''     Resets form to initial state
    ''' </summary>
    Sub resetForm()
        txtGrade1.Text = ""
        txtGrade2.Text = ""
        txtGrade3.Text = ""
        txtGrade4.Text = ""
        txtGrade5.Text = ""
        txtGrade6.Text = ""
        lblCourse1Output.Text = ""
        lblCourse2Output.Text = ""
        lblCourse3Output.Text = ""
        lblCourse4Output.Text = ""
        lblCourse5Output.Text = ""
        lblCourse6Output.Text = ""
        lblSemesterAverage.Text = ""
        lblSemesterLetterGrade.Text = ""

        btnCalculate.Enabled = True

        txtGrade1.Focus()

    End Sub
    ''' <summary>
    '''     Validates the user input
    ''' </summary>
    ''' <param name="input">user input to be validated</param>
    ''' <returns>Whether the value is valid or invalid as boolean</returns>
    Private Function validateUserInput(ByVal input As String) As Boolean
        Dim inputValue As Double
        Dim isValidInput As Boolean = False

        Try
            inputValue = CDbl(input)
            If (inputValue >= MINIMUM_GRADE AndAlso inputValue <= MAXIMUM_GRADE) Then
                isValidInput = True
            End If
        Catch ex As Exception

        End Try

        Return isValidInput
    End Function
    ''' <summary>
    '''     Converts the grade achieved to a Letter value
    ''' </summary>
    ''' <param name="gradeToLetter"></param>
    ''' <returns></returns>
    Private Function gradeConversion(ByVal gradeToLetter As Double) As String
        Dim letterGrade As String = ""
        Select Case gradeToLetter
            Case 90 To 100
                letterGrade = "A+"
            Case 85 To 89
                letterGrade = "A"
            Case 80 To 84
                letterGrade = "A-"
            Case 77 To 79
                letterGrade = "B+"
            Case 73 To 76
                letterGrade = "B"
            Case 70 To 72
                letterGrade = "B-"
            Case 67 To 69
                letterGrade = "C+"
            Case 63 To 66
                letterGrade = "C"
            Case 60 To 62
                letterGrade = "C-"
            Case 57 To 59
                letterGrade = "D+"
            Case 53 To 56
                letterGrade = "D"
            Case 50 To 52
                letterGrade = "D-"
            Case 0 To 49
                letterGrade = "F"

        End Select

        Return letterGrade
    End Function
    ''' <summary>
    '''     Returns the average value of a provided array
    ''' </summary>
    ''' <param name="arrayToAverage"></param>
    ''' <returns></returns>
    Private Function averageArray(ByVal arrayToAverage() As Double) As Double
        Dim total As Double

        For Each studentGrade In arrayToAverage
            total += studentGrade
        Next

        averageArray = Math.Round(total / arrayToAverage.Length, 2)
    End Function

#End Region

#Region "Event Handlers"
    ''' <summary>
    '''     Handles the calculate button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        lblSemesterAverage.Text = averageArray(grades).ToString
        lblSemesterLetterGrade.Text = gradeConversion(grade)

    End Sub
    ''' <summary>
    '''     Handles the reset button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Call resetForm()
    End Sub
    ''' <summary>
    '''     Handles the exit button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub txtGrade1_Leave(sender As Object, e As EventArgs) Handles txtGrade1.Leave

        If (validateUserInput(txtGrade1.Text)) Then
            lblCourse1Output.Text = gradeConversion(gradeToLetter)
        Else
            lblErrorOutput.Text = ERROR_MESSAGE
        End If


    End Sub

    Private Sub txtGrade2_Leave(sender As Object, e As EventArgs) Handles txtGrade2.Leave
        If (validateUserInput(txtGrade2.Text)) Then
            lblCourse2Output.Text = gradeConversion(gradeToLetter)
        Else
            lblErrorOutput.Text = ERROR_MESSAGE
        End If


    End Sub

    Private Sub txtGrade3_Leave(sender As Object, e As EventArgs) Handles txtGrade3.Leave
        If (validateUserInput(txtGrade3.Text)) Then
            lblCourse3Output.Text = gradeConversion(gradeToLetter)
        Else
            lblErrorOutput.Text = ERROR_MESSAGE
        End If


    End Sub

    Private Sub txtGrade4_Leave(sender As Object, e As EventArgs) Handles txtGrade4.Leave
        If (validateUserInput(txtGrade4.Text)) Then
            lblCourse4Output.Text = gradeConversion(gradeToLetter)
        Else
            lblErrorOutput.Text = ERROR_MESSAGE
        End If


    End Sub

    Private Sub txtGrade5_Leave(sender As Object, e As EventArgs) Handles txtGrade5.Leave
        If (validateUserInput(txtGrade5.Text)) Then
            lblCourse5Output.Text = gradeConversion(gradeToLetter)
        Else
            lblErrorOutput.Text = ERROR_MESSAGE
        End If


    End Sub

    Private Sub txtGrade6_Leave(sender As Object, e As EventArgs) Handles txtGrade6.Leave
        If (validateUserInput(txtGrade6.Text)) Then
            lblCourse6Output.Text = gradeConversion(gradeToLetter)
        Else
            lblErrorOutput.Text = ERROR_MESSAGE
        End If


    End Sub
#End Region

End Class

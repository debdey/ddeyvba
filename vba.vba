Sub AddComment()
    Dim rngComent As Range, rngComent_EID As Range, rngComent_Name As Range, rngComent_CareerLevel As Range, rngComent_Project As Range, rngComent_Location As Range, rngComent_Supervisor As Range, rngComent_Tower As Range, rngComent_ContactNo As Range, rngComent_HESSID As Range
    Dim rng As Range
    Dim cm As comment, i As Integer
    Set rngComent = Sheets(1).Range("a2:a500")
    Set rngComent_EID = Sheets(1).Range("b2:b500")
    Set rngComent_Name = Sheets(1).Range("c2:c500")
    Set rngComent_CareerLevel = Sheets(1).Range("d2:d500")
    Set rngComent_Project = Sheets(1).Range("f2:f500")
    Set rngComent_Location = Sheets(1).Range("i2:i500")
    Set rngComent_Supervisor = Sheets(1).Range("k2:k500")
    Set rngComent_Tower = Sheets(1).Range("l2:l500")
    Set rngComent_ContactNo = Sheets(1).Range("p2:p500")
    Set rngComent_HESSID = Sheets(1).Range("v2:v500")
    For Each rng In rngComent
        i = i + 1
        If Not rng.comment Is Nothing Then
            rng.comment.Delete
        End If
        Set cm = rng.AddComment
        With cm
            .Visible = False
            .Text Text:="Enterprise ID :" & Chr(2) & rngComent_EID(i).Value & Chr(10) & "Name :" & Chr(2) & rngComent_Name(i).Value & Chr(10) & "Career Level :" & Chr(2) & rngComent_CareerLevel(i).Value & Chr(10) & "Project :" & Chr(2) & rngComent_Project(i).Value & Chr(10) & "Location :" & Chr(2) & rngComent_Location(i).Value & Chr(10) & "Supervisor :" & Chr(2) & rngComent_Supervisor(i).Value & Chr(10) & "Tower :" & Chr(2) & rngComent_Tower(i).Value & Chr(10) & "Contact number :" & Chr(2) & rngComent_ContactNo(i).Value & Chr(10) & "HESS ID :" & Chr(2) & rngComent_HESSID(i).Value & Chr(10)
        End With
    Next rng
    Call Sheet1.FitComments
End Sub

Sub FitComments()
Dim xComment As comment
For Each xComment In Application.ActiveSheet.Comments
    xComment.Shape.TextFrame.AutoSize = True
    xComment.Shape.AutoShapeType = msoShapeRoundedRectangle
    xComment.Shape.TextFrame.Characters.Font.Name =
    "Tahoma"
    xComment.Shape.TextFrame.Characters.Font.Size = 8
    xComment.Shape.TextFrame.Characters.Font.ColorIndex = 2
    xComment.Shape.Line.ForeColor.RGB = RGB(0, 0, 0)
    xComment.Shape.Line.BackColor.RGB = RGB(255, 255, 255)
    xComment.Shape.Fill.Visible = msoTrue
    xComment.Shape.Fill.ForeColor.RGB = RGB(58, 82, 184)
    xComment.Shape.Fill.OneColorGradient msoGradientDiagonalUp, 1, 0.23
Next
End Sub

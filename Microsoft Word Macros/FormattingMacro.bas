Attribute VB_Name = "Module1"
Public Sub replacebold()
'        .Replacement.Text = Replace("<b>^&</b>", Chr(10) & "</b>", "</b>" & Chr(10))

    With ActiveDocument.Content.Find
        .Font.bold = True
        .Text = "^13"
        .Replacement.Font.bold = False
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Font.bold = True
        .Text = "^11"
        .Replacement.Font.bold = False
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Font.bold = True
        .Replacement.Text = "<b>^&</b>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<b> "
        .Replacement.Text = " <b>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
        
    With ActiveDocument.Content.Find
        .Text = "</b><b>"
        .Replacement.Text = ""
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<b></b>"
        .Replacement.Text = ""
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "</b> <b>"
        .Replacement.Text = " "
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<B>"
        .Replacement.Text = "<b>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "</B>"
        .Replacement.Text = "</b>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
End Sub


Public Sub replaceitalic()
    With ActiveDocument.Content.Find
        .Font.italic = True
        .Text = "^13"
        .Replacement.Font.italic = False
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Font.italic = True
        .Text = "^11"
        .Replacement.Font.italic = False
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With

    With ActiveDocument.Content.Find
        .Font.italic = True
        .Replacement.Text = "<i>^&</i>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<i> "
        .Replacement.Text = " <i>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
        
    With ActiveDocument.Content.Find
        .Text = "<i></i>"
        .Replacement.Text = ""
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
        
    With ActiveDocument.Content.Find
        .Text = "</i><i>"
        .Replacement.Text = ""
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<I>"
        .Replacement.Text = "<i>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "</I>"
        .Replacement.Text = "</i>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
End Sub


Public Sub replaceunderline()
    With ActiveDocument.Content.Find
        .Font.underline = wdUnderlineSingle
        .Text = "^13"
        .Replacement.Font.underline = False
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Font.underline = wdUnderlineSingle
        .Text = "^11"
        .Replacement.Font.underline = False
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With

    With ActiveDocument.Content.Find
        .Font.underline = wdUnderlineSingle
        .Replacement.Text = "<u>^&</u>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
            With ActiveDocument.Content.Find
        .Text = "<u> "
        .Replacement.Text = " <u>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
        
    With ActiveDocument.Content.Find
        .Text = "</u><u>"
        .Replacement.Text = ""
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<u></u>"
        .Replacement.Text = ""
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<U>"
        .Replacement.Text = "<u>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "</U>"
        .Replacement.Text = "</u>"
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
End Sub


Sub replaceall()

Call replaceunderline
Call replacebold
Call replaceitalic

    With ActiveDocument.Content.Find
        .Text = " </b>"
        .Replacement.Text = "</b> "
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = " </i>"
        .Replacement.Text = "</i> "
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With

End Sub


Sub replacehyperlinks()
'
' hyper Macro
'
'



    Dim i As Long
    Dim strURL As String
    Dim strText As String
    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
        With ActiveDocument.Hyperlinks(i)
            strURL = .Address
            strText = .TextToDisplay
            .Range.Text = "<a href='" & strURL & "'" & " target='_blank'>" & strText & "</a>"
        End With
    Next i


End Sub

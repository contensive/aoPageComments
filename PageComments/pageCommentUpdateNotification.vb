Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace contensive.addon.pageComment
    '
    '
    '
    Public Class pageCommentUpdateNotificationClass
        Inherits AddonBaseClass
        '
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String
            Try
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim pageId As Integer = 0
                Dim pageName As String = String.Empty
                Dim memberName As String = String.Empty
                Dim memberEmail As String = String.Empty
                Dim commentDate As Date = Date.MinValue
                Dim comment As String = String.Empty
                '
                Dim emailHtmlBody As String = String.Empty
                Dim emailHtmlBodyHeader As String = String.Empty
                Dim emailHtmlBodyDetail As String = String.Empty
                '


                If cs.Open(cnPageComments, "commentStatus = 1 ", "pageID") Then
                    Do
                        '
                        pageId = cs.GetInteger("pageID")
                        pageName = cs.GetText("pageID")
                        memberName = cs.GetText("memberName")
                        memberEmail = cs.GetText("memberEmail")
                        commentDate = cs.GetDate("commentDate")
                        comment = cs.GetTextFile("comment")
                        '
                        emailHtmlBodyDetail &= " <tr>"
                        emailHtmlBodyDetail &= " <td> " & pageName & " </td> "
                        emailHtmlBodyDetail &= " <td> " & memberName & " </td> "
                        emailHtmlBodyDetail &= " <td> " & memberEmail & " </td> "
                        emailHtmlBodyDetail &= " <td> " & commentDate.ToString & " </td> "
                        emailHtmlBodyDetail &= " <td> " & comment & " </td> "
                        emailHtmlBodyDetail &= " </tr>" & vbCrLf
                        '
                        Call cs.GoNext()
                    Loop While cs.OK
                End If

                If Not String.IsNullOrEmpty(emailHtmlBodyDetail) Then
                    emailHtmlBodyHeader &= " <tr> "
                    emailHtmlBodyHeader &= " <th>Page</th> "
                    emailHtmlBodyHeader &= " <th>Member</th> "
                    emailHtmlBodyHeader &= " <th>Member Email</th> "
                    emailHtmlBodyHeader &= " <th>Comment Date</th> "
                    emailHtmlBodyHeader &= " <th>Comment</th> "
                    emailHtmlBodyHeader &= " </tr> " & vbCrLf
                    '
                    emailHtmlBody = emailHtmlBodyHeader & emailHtmlBodyDetail
                    '
                    ' Send the notofocation email
                    '
                    CP.Email.sendSystem("Page Comment Update Notification", emailHtmlBody)
                    '
                End If

                Call cs.Close()
                returnHtml = "OK"
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = "Error"
            End Try
            Return returnHtml
        End Function
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace

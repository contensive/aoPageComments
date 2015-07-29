
Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports contensive.BaseClasses

Namespace contensive.addon.pageComment
    Public Class pageCommentsForm
        Inherits AddonBaseClass
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Dim pageID As Integer = CP.Doc.PageId
            '
            Dim AllowCommentDisplay As Boolean = False
            'Dim CommentNeedApprove As Boolean = False
            'Dim CommentNotificationGroup As Integer = 0
            Dim AllowAnonymousComments As Boolean = False
            Dim forceLoginAnonymousComment As Boolean = False
            '
            Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
            Dim layout As CPBlockBaseClass = CP.BlockNew()
            Dim oneCommentLayout As CPBlockBaseClass = CP.BlockNew()
            Dim oneMessajeLayout As CPBlockBaseClass = CP.BlockNew()
            Dim authenticateFormLayout As CPBlockBaseClass = CP.BlockNew()
            Dim nonAuthenticateFormLayout As CPBlockBaseClass = CP.BlockNew()
            Dim commentList As String = ""
            Dim commentDate As DateTime
            Dim totalComments As Integer = 0
            Dim recaptchaHtml As String = String.Empty

            '
            Try
                '
                ' TCA Comments Layout
                '
                layout.OpenLayout("{A5B5B56C-FF44-45CB-8933-7FE6D9AD9D8A}")
                '
                oneCommentLayout.Load(layout.GetOuter("#js-onecomment"))

                ' check the comment settings for the actual page
                '

                Dim AnonymousCommentsOption As Integer = CP.Site.GetInteger("Page Comments - Option for Non-Authenticate Users", "0")
                '
                ' Default Values
                '
                AllowAnonymousComments = False
                forceLoginAnonymousComment = False
                '
                Select Case AnonymousCommentsOption
                    Case 0
                        ' leave dafault values
                    Case 1
                        AllowAnonymousComments = True
                    Case 2
                        AllowAnonymousComments = True
                        forceLoginAnonymousComment = True
                End Select



                If cs.Open(cnPageContent, " id = " & pageID.ToString) Then
                    '
                    AllowCommentDisplay = cs.GetBoolean("AllowCommentDisplay")
                    'CommentNeedApprove = cs.GetBoolean("CommentNeedApprove")
                    'CommentNotificationGroup = cs.GetInteger("CommentNotificationGroup")
                    '
                End If
                Call cs.Close()

                ' Check  if exist coments for actual page
                If cs.Open(cnPageComments, " pageID = " & pageID.ToString & " and commentStatus=2") Then
                    Do
                        '
                        oneCommentLayout.Load(layout.GetOuter("#js-onecomment"))
                        '
                        ' set member name
                        oneCommentLayout.SetInner(".fn", cs.GetText("memberName"))
                        '
                        ' set date
                        commentDate = cs.GetDate("commentDate")
                        oneCommentLayout.SetInner(".commentDate", MonthName(commentDate.Month) & " " & commentDate.Day & ", " & commentDate.Year & " at " & commentDate.ToString("hh:mm tt"))
                        ' 
                        ' Set comment
                        oneCommentLayout.SetInner(".comment-content", cs.GetText("comment"))

                        commentList &= oneCommentLayout.GetHtml
                        '
                        totalComments += 1
                        Call cs.GoNext()
                    Loop While cs.OK

                End If
                Call cs.Close()

                If totalComments <> 0 Then
                    If totalComments > 1 Then
                        layout.SetInner("#js-num-comments", totalComments.ToString & " comments")
                    Else
                        layout.SetInner("#js-num-comments", totalComments.ToString & " comment")
                    End If
                    '
                    layout.SetInner("#js-comment-title", CP.Doc.PageName)
                    '
                    layout.SetInner("#js-commentlist", commentList)
                    '
                Else
                    layout.SetOuter("#comments", "")
                End If

                ' show comment submit form
                If AllowCommentDisplay Then
                    ' show comment form
                    'layout.SetInner("#page", CP.Doc.PageId.ToString.Trim)

                    If CP.User.IsAuthenticated Then
                        layout.SetOuter("#js-messaje", "")
                        layout.SetOuter("#js-nonAuthenticate-reply", "")
                        layout.SetOuter("#js-nonAuthenticateLogin-reply", "")
                    Else
                        If AllowAnonymousComments Then
                            If forceLoginAnonymousComment Then
                                layout.SetOuter("#js-Authenticate-reply", "")
                                layout.SetOuter("#js-nonAuthenticate-reply", "")
                            Else
                                layout.SetOuter("#js-Authenticate-reply", "")
                                layout.SetOuter("#js-nonAuthenticateLogin-reply", "")
                                layout.SetOuter("#js-messaje", "")
                                ' recaptcha
                                recaptchaHtml = CP.Utils.ExecuteAddon("{500A1F57-86A2-4D47-B747-4EF4D30A83E2}")
                                layout.SetInner("#captchaArea", recaptchaHtml)
                            End If
                        Else
                            layout.SetOuter("#js-messaje", "")
                            layout.SetOuter("#js-Authenticate-reply", "")
                            layout.SetOuter("#js-nonAuthenticate-reply", "")
                            layout.SetOuter("#js-nonAuthenticateLogin-reply", "")
                        End If
                    End If


                    returnHtml = layout.GetHtml & CP.Html.Hidden("", CP.Doc.PageId.ToString, "", "page")
                End If

            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = "Contensive Addon - Error response"
            End Try
            Return returnHtml
        End Function

        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in contensive.addon.tca630Redesign.pageCommentsForm " & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub

    End Class
End Namespace

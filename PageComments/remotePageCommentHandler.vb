﻿
Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports contensive.BaseClasses
Imports System.Text.RegularExpressions

Namespace contensive.addon.pageComment
    Public Class remotePageCommentHandler
        Inherits AddonBaseClass
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Dim memberName As String = ""
            Dim memberEmail As String = ""
            Dim comment As String = CP.Doc.GetText("comment")
            Dim pageID As String = CP.Doc.GetText("page")
            Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
            Dim errorSTR As String = ""
            Dim errorStep As String = ""
            '
            Dim CommentNeedApprove As Boolean = True
            '

            Try
                errorStep = "11"
                If CP.User.IsAuthenticated Then
                    memberName = CP.User.Name
                    memberEmail = CP.User.Email
                Else
                    memberName = CP.Doc.GetText("author")
                    memberEmail = CP.Doc.GetText("email")
                    errorStep = "2"

                    If String.IsNullOrEmpty(memberName.Trim) Then
                        errorSTR &= "error-name,"
                    End If

                    errorStep = "3"
                    'Or Not IsValidEmailFormat(memberEmail.Trim)

                    If String.IsNullOrEmpty(memberEmail.Trim) Or Not IsValidEmailFormat(memberEmail.Trim) Then
                        errorSTR &= "error-mail,"
                    End If
                    errorStep = "5"

                    Dim mystrtmp As String = CP.Utils.ExecuteAddon("reCaptcha")

                    If Not CP.UserError.OK Then
                        errorSTR &= "error-captcha,"
                    End If
                End If


                errorStep = "44"
                If String.IsNullOrEmpty(comment.Trim) Then
                    errorSTR &= "error-comment,"
                End If


                errorStep = "66"
                If String.IsNullOrEmpty(errorSTR) Then

                    '
                    ' check page setting if comment need to be review
                    '

                    'Dim CommentNeedApprove As Boolean = True
                    'If cs.Open(cnPageContent, "id=" & pageID) Then
                    '    CommentNeedApprove = cs.GetBoolean("CommentNeedApprove")
                    'End If
                    'Call cs.Close()

                    CommentNeedApprove = CP.Site.GetBoolean("Page Comments - Comments Need to be Approve", "0")
                    '
                    If cs.Insert(cnPageComments) Then
                        cs.SetField("pageID", pageID.Trim)
                        cs.SetField("memberID", CP.User.Id.ToString)
                        cs.SetField("memberName", memberName)
                        cs.SetField("memberEmail", memberEmail)
                        cs.SetField("commentDate", Now.ToString)
                        cs.SetField("comment", comment)
                        If CommentNeedApprove Then
                            cs.SetField("commentStatus", "1")
                        Else
                            cs.SetField("commentStatus", "2")
                        End If

                    End If
                    Call cs.Close()
                    errorSTR = "OK"
                Else
                    errorSTR = errorSTR.Substring(0, errorSTR.Length - 1)
                End If
                returnHtml = errorSTR
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = "Contensive Addon - Error response" & errorStep
            End Try
            Return returnHtml
        End Function

        '
        '
        '
        Function IsValidEmailFormat(ByVal s As String) As Boolean
            Return Regex.IsMatch(s, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
        End Function
        '
        '
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
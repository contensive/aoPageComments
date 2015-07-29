Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace contensive.addon.pageComment
    '
    '
    '
    Public Class pageCommentLoginHandlerClass
        Inherits AddonBaseClass
        '
        '
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim cs As CPCSBaseClass = CP.CSNew
            Dim csAuth As CPCSBaseClass = CP.CSNew
            Dim returnHtml As String
            Dim username As String = ""
            Dim password As String = ""
            Dim msgLoginResult As Boolean = False
            Dim msgLoginError As String = ""

            Try
                username = CP.Doc.GetText("usn")
                password = CP.Doc.GetText("pwd")

                If cs.Open(cnPeople, "(username=" & CP.Db.EncodeSQLText(username) & ") and (password=" & CP.Db.EncodeSQLText(password) & ")", , False) Then
                    If cs.GetBoolean("active") Then
                        ' user is ok
                        CP.User.Login(username, password)
                        msgLoginError = ""
                        msgLoginResult = True
                    Else
                        msgLoginError = "User is not Active"
                    End If

                Else
                    msgLoginError = "Error in Login"
                End If
                Call cs.Close()

                returnHtml = "[{""result"":" & msgLoginResult.ToString.ToLower & ",""msgError"":""" & msgLoginError & """}]"
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = ""
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
                cp.Site.ErrorReport(ex, "Unexpected error in tcaLoginHandlerClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace

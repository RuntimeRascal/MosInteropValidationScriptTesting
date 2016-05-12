

Imports Microsoft.VisualBasic
Imports System
Imports System.Diagnostics
Imports System.Text
Imports System.IO
Imports System.Environment
Imports System.Xml
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports System.Collections.Generic
Imports Microsoft.Win32
Imports Microsoft.Office.Interop.Word

Namespace GMetrix.Dynamic.Validations

    Public Class DynamicQuestion

        Private _myOfficeApp As Microsoft.Office.Interop.Word.Application

        Public Sub New()
            MyBase.New
            _myOfficeApp = New Microsoft.Office.Interop.Word.Application
            MaximizeApp()
        End Sub

        Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.Word.Application
            Get
                Return Me._myOfficeApp
            End Get
            Set
                Me._myOfficeApp = Value
            End Set
        End Property

        Public Overridable Function ValidateQuestion(ByVal ParamArray Parameters As Object()) As [Object]
            'Juan 5/20/2013 
            Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
            Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
            Dim keyWords As List(Of String) = CType(Parameters(2), List(Of String))
            Dim blueWords As List(Of String) = CType(Parameters(3), List(Of String))
            Dim var2 As String = templates(0)

            var2 = Replace(var2, "Sample", keyWords(0))

            If Not isValidate Then ' Start Pre-Code
                Try
                    With Me.MyOfficeApp
                        .Visible = True
                    End With
                    If My.Computer.FileSystem.FileExists(var2) Then
                        My.Computer.FileSystem.DeleteFile(var2)
                    End If

                    Return True
                Catch ex As Exception
                    Return False
                End Try

            Else ' Start Validation-Code    

                Dim points(0)
                points(0) = False
                Dim var1 As String = blueWords(0)
                Try

                    'Crear un documento basado en la plantilla Adjacency Letter
                    If Me.MyOfficeApp.Documents.Count > 0 Then
                        If InStr(Me.MyOfficeApp.ActiveDocument.AttachedTemplate.Name, var1) Then
                            If My.Computer.FileSystem.FileExists(var2) Then
                                points(0) = True
                            End If
                        End If

                    End If
                Catch ex As Exception
                    Return points
                End Try
                Return points
            End If
        End Function

        Public Overridable Sub CloseDialogs()
            Dim procs() = Process.GetProcesses
            Dim pr As Process
            For Each pr In procs
                If (pr.ProcessName = "WINWORD") Then
                    Dim cdb As Integer
                    For cdb = 1 To 3
                        Try
                            AppActivate(pr.Id)
                            System.Windows.Forms.SendKeys.Send("{ESC}")
                            System.Threading.Thread.Sleep(50)
                            System.Windows.Forms.Application.DoEvents()
                        Catch ex As System.Exception
                            Exit For
                        End Try
                    Next
                End If
            Next

        End Sub

        Public Overridable Sub MaximizeApp()
            Try
                Me.MyOfficeApp.Visible = True
                Me.MyOfficeApp.WindowState = WdWindowState.wdWindowStateMaximize
            Catch Ex As System.Exception
            End Try

        End Sub

        Public Overridable Sub MinimizeApp()
            Try
                Me.MyOfficeApp.WindowState = WdWindowState.wdWindowStateMinimize
            Catch ex As System.Exception
            End Try

        End Sub

        Public Overridable Sub CloseApp()
            Try
                If MyOfficeApp IsNot Nothing Then
                    Dim procs() = Process.GetProcesses
                    Dim pr As Process
                    For Each pr In procs
                        If (pr.ProcessName = "WINWORD") Then
                            Dim a, b As Int16
                            b = MyOfficeApp.Documents.Count
                            For a = 1 To b
                                MyOfficeApp.ActiveDocument.Close(SaveChanges:=False)
                                b = MyOfficeApp.Documents.Count
                                If b < 1 Then
                                    Exit For
                                End If
                            Next
                            Exit For
                        End If
                    Next
                    MyOfficeApp.Quit()
                    If MyOfficeApp IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
                        MyOfficeApp = Nothing
                    End If
                End If
            Catch ex As System.Exception
            End Try

        End Sub

        Public Overridable Sub FinalCloseApp()
            Try
                Dim procs() = Process.GetProcesses
                Dim pr As Process
                For Each pr In procs
                    If (pr.ProcessName = "WINWORD") Then
                        If MyOfficeApp IsNot Nothing Then
                            Dim a, b As Int16
                            b = MyOfficeApp.Documents.Count
                            For a = 1 To b
                                MyOfficeApp.ActiveDocument.Close(SaveChanges:=False)
                                b = MyOfficeApp.Documents.Count
                                If b < 1 Then
                                    Exit For
                                End If
                            Next
                            Exit For
                        End If
                    End If
                Next
            Catch Ex As System.Exception
            End Try
            If MyOfficeApp IsNot Nothing Then
                MyOfficeApp.Quit()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
                MyOfficeApp = Nothing
            End If

        End Sub

        Private Sub GarbageCollector()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Sub
    End Class
End Namespace

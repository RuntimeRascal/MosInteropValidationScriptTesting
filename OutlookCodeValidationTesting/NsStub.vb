

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
Imports Microsoft.Office.Interop.Outlook

Namespace GMetrix.Dynamic.Validations

    Public Class DynamicQuestion

        Private _myOfficeApp As Microsoft.Office.Interop.Outlook.Application

        Public Sub New()
            MyBase.New
            _myOfficeApp = New Microsoft.Office.Interop.Outlook.Application
            Dim mapiNameSpace As Microsoft.Office.Interop.Outlook.NameSpace = Me.MyOfficeApp.GetNamespace("mapi")
            mapiNameSpace.Logon("GMetrix", "", True, True)
        End Sub

        Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.Outlook.Application
            Get
                Return Me._myOfficeApp
            End Get
            Set
                Me._myOfficeApp = Value
            End Set
        End Property

        Public Overridable Function ValidateQuestion(ByVal ParamArray Parameters As Object()) As [Object]
            ' Tommy 06/02/2014
            ' Question 3633


            Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
            Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
            Dim keyWords As List(Of String) = CType(Parameters(2), List(Of String))
            Dim blueWords As List(Of String) = CType(Parameters(3), List(Of String))


            If Not isValidate Then ' Start Pre-Code

                ' Variables
                Dim ofInbox As Microsoft.Office.Interop.Outlook.Folder = Nothing


                Try
                    Try
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Options\Spelling\", "Check", "0", RegistryValueKind.DWord)

                        System.Threading.Thread.Sleep(3000)
                        'Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\15.0\Outlook\Options\Spelling\", True).SetValue("Check", "0", RegistryValueKind.DWord)
                    Catch hkeyException As System.Exception
                    End Try

                    Try
                        ofInbox = Me.MyOfficeApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                        ofInbox.Display()
                        Me.MyOfficeApp.ActiveWindow().WindowState = OlWindowState.olMaximized
                        Me.MyOfficeApp.ActiveWindow().WindowState = OlWindowState.olNormalWindow
                        Me.MyOfficeApp.ActiveWindow().WindowState = OlWindowState.olMaximized
                    Catch
                    End Try

                    Return True
                Catch ex As System.Exception
                    Return True

                Finally
                    If ofInbox IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ofInbox)
                        ofInbox = Nothing
                    End If

                End Try


            Else ' Start Validation-Code
                Dim points(0)
                points(0) = False

                Try

                    Dim _spellingCheckSetting As String = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\15.0\Outlook\Options\Spelling\", True).GetValue("Check").ToString

                    If _spellingCheckSetting.ToString = "1" Then
                        points(0) = True
                    End If

                    Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\15.0\Outlook\Options\Spelling\", True).SetValue("Check", "0", RegistryValueKind.DWord)

                Catch ex As System.Exception
                    Return points

                Finally

                End Try
                Return points
            End If
        End Function

        Public Overridable Sub CloseDialogs()
            Dim procs() = Process.GetProcesses
            Dim pr As Process
            For Each pr In procs
                If (pr.ProcessName = "OUTLOOK") Then
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
                Me.MyOfficeApp.WindowState = OlWindowState.olMaximized
            Catch Ex As System.Exception
            End Try

        End Sub

        Public Overridable Sub MinimizeApp()
            Try
                Me.MyOfficeApp.WindowState = OlWindowState.olMinimized
            Catch ex As System.Exception
            End Try

        End Sub

        Public Overridable Sub CloseApp()
            Try
                If MyOfficeApp IsNot Nothing Then
                    Dim ofInbox As Microsoft.Office.Interop.Outlook.Folder = Me.MyOfficeApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                    If ofInbox IsNot Nothing Then
                        For Each item As Object In ofInbox.Items
                            item.Delete()
                        Next
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
                        MyOfficeApp = Nothing
                    End If
                    MyOfficeApp.Quit()
                End If
            Catch Ex As System.Exception
            End Try


        End Sub

        Public Overridable Sub FinalCloseApp()
            Try
                Dim ofInbox As Microsoft.Office.Interop.Outlook.Folder = Me.MyOfficeApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                If ofInbox IsNot Nothing Then
                    For Each item As Object In ofInbox.Items
                        item.Delete()
                    Next
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ofInbox)
                    ofInbox = Nothing
                End If
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

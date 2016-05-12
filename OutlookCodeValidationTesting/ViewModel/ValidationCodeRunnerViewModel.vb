Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports GalaSoft.MvvmLight
Imports GalaSoft.MvvmLight.CommandWpf
Imports Microsoft.Office.Interop.Outlook

Namespace ViewModel
    ''' <summary>
    '''     Class ValidationCodeRunnerViewModel.
    ''' </summary>
    ''' <seealso cref="GalaSoft.MvvmLight.ViewModelBase" />
    Public Class ValidationCodeRunnerViewModel
        Inherits ViewModelBase

#Region "Contructor"

        ''' <summary>
        '''     Initializes a new instance of the <see cref="ValidationCodeRunnerViewModel" /> class.
        ''' </summary>
        Public Sub New()
           'ValidateQuestionCommand = New RelayCommand(ValidateQuestion())

            PromptToFillArrayCommand = New RelayCommand(AddressOf PromptToFillArray)
            PrintParamsCommand = New RelayCommand(AddressOf PrintParams)

            SetupCommand = New RelayCommand(AddressOf Setup)
            RunValidationPreCodeCommand = New RelayCommand(AddressOf RunValidationPreCode)
            RunValidationPostCodeCommand = New RelayCommand(AddressOf RunValidationPostCode)
            EndCommand = New RelayCommand(AddressOf [End])
            CloseDialogsCommand = New RelayCommand(AddressOf CloseDialogs)
            MaximizeAppCommand = New RelayCommand(AddressOf MaximizeApp)
            MinimizeAppCommand = New RelayCommand(AddressOf MinimizeApp)
            CloseAppCommand = New RelayCommand(AddressOf CloseApp)
            FinalCloseAppCommand = New RelayCommand(AddressOf FinalCloseApp)
            GarbageCollectorCommand = New RelayCommand(AddressOf GarbageCollector)
        End Sub

#End Region

#Region "Fields"

        Dim _localParameters As Object() = New Object() {True, ' Compiler result (Pre-Code or Post-Code) True or False
                                                         New List(Of String) From {"c:\temp\document1.docx"},
                                                         New List(Of String), ' OrangeWords
                                                         New List(Of String), ' BlueWords
                                                         New List(Of String), ' RedWords
                                                         New List(Of String), ' GreenWords
                                                         New List(Of String), ' PurpleWords
                                                         New List(Of String) ' YellowWords
                                                        }

        ReadOnly _
            _parameterDictionary = New Dictionary(Of Integer, String) _
            From {{0, "(Pre-Code or Post-Code) True or False"},
            {1, "Sample files array"},
            {2, "OrangeWords"},
            {3, "BlueWords"},
            {4, "RedWords"},
            {5, "GreenWords"},
            {6, "PurpleWords"},
            {7, "YellowWords"}
            }

        Private _myOfficeApp As Microsoft.Office.Interop.Outlook.Application

#End Region

#Region "Properties"

        ''' <summary>
        '''     Gets or sets my office application.
        ''' </summary>
        ''' <value>My office application.</value>
        Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.Outlook.Application
            Get
                Return _myOfficeApp
            End Get
            Set
                [Set]( _myOfficeApp, Value,True,"MyOfficeApp" )
            End Set
        End Property

#End Region

#Region "Methods"
        Public Sub Setup()
            MyOfficeApp = New Microsoft.Office.Interop.Outlook.Application()

            If MyOfficeApp IsNot Nothing Then
                Console.WriteLine("Version: " + MyOfficeApp.Version)
                Console.WriteLine("Name: " + MyOfficeApp.Name)


                'Console.WriteLine("Caption: " + MyOfficeApp.Caption)

                '		If MyOfficeApp.ActiveWindow IsNot Nothing Then
                '			Console.WriteLine("HWND = " & MyOfficeApp.ActiveWindow.Hwnd.ToString())
                '		End If

                'Dim hwnd = MyOfficeApp.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, Nothing, MyOfficeApp, Nothing)
                'Console.WriteLine("HWND = " & hwnd.ToString())

               'Dim Caption = MyOfficeApp.GetType().InvokeMember("Caption", Reflection.BindingFlags.GetProperty, Nothing, MyOfficeApp, Nothing).ToString()
               
                'MyOfficeApp.Visible = True

'                Try
'                    If MyOfficeApp.ActiveWindow IsNot Nothing Then
'                        Console.WriteLine("HWND = " & MyOfficeApp.ActiveWindow().Hwnd.ToString())
'                        'MyOfficeApp.ActiveWindow
'                    End If
'                Catch Ex As COMException
'                End Try

            End If
            Console.WriteLine("")
        End Sub

        Public Sub RunValidationPreCode()
            Console.WriteLine(vbCrLf & "Executing the pre Code. . .")

            If MyOfficeApp IsNot Nothing Then

            End If
        End Sub

        Public Sub RunValidationPostCode()
            Console.WriteLine(vbCrLf & "Executing the post Code. . .")

            If MyOfficeApp IsNot Nothing Then

            End If
        End Sub

        Public Sub [End]()
            Console.WriteLine(vbCrLf & "Executing the termination Code. . .")

            If MyOfficeApp IsNot Nothing Then
                MyOfficeApp.Quit()
                Marshal.FinalReleaseComObject(MyOfficeApp)
                MyOfficeApp = Nothing
            End If
        End Sub

        Sub PromptToFillArray()
            _localParameters = New Object() {True, ' Compiler result (Pre-Code or Post-Code) True or False
                                             New List(Of String), ' Sample files array
                                             New List(Of String), ' OrangeWords
                                             New List(Of String), ' BlueWords
                                             New List(Of String), ' RedWords
                                             New List(Of String), ' GreenWords
                                             New List(Of String), ' PurpleWords
                                             New List(Of String) ' YellowWords
                                            }

            Dim Input = ""

            Console.WriteLine(vbCrLf & vbCrLf & "----------------- Populate Parameter Array --------------------")
            Console.WriteLine("Options:")
            For Each item In _parameterDictionary
                Console.WriteLine(vbTab & vbTab & vbTab & item.Key.ToString() & " => " & item.Value)
            Next
            Console.WriteLine(vbTab & vbTab & vbTab & "done => finsish populating array")

            While Input <> "done"
                Console.Write(vbCrLf & "Selection: ")
                Input = Console.ReadLine()

                If Input.ToLower().Contains("done") Then
                    Exit While
                End If

                Dim Key As Integer
                Integer.TryParse(Input, Key)

                If Not Integer.TryParse(Input, Key) Then
                    Console.WriteLine("Enter a number dummy")
                    Continue While
                ElseIf Not _parameterDictionary.ContainsKey(Key) Then
                    Console.WriteLine(Key + " is not valid")
                Else
                    Console.Write(vbCrLf & "Enter " + _parameterDictionary(Key) + " values: ")
                    Input = Console.ReadLine()
                    If String.IsNullOrEmpty(Input) Then
                        Continue While
                    End If

                    If Key = 0 Then
                        Dim Val = False
                        If Boolean.TryParse(Input, Val) Then
                            _localParameters(Key) = Val
                        Else
                            Console.WriteLine("Could'nt cast your value. Enter either 'True' or 'False'.")
                        End If
                    Else
                        Dim Tokens = Input.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                        Dim Collection = DirectCast(_localParameters(Key), List(Of String))

                        Collection.AddRange(Tokens)
                    End If
                End If
            End While

            PrintParams()
        End Sub

        Sub PrintParams()
            Console.WriteLine(vbCrLf & vbCrLf & "Parameter Array looks like this:")
            For index = 0 To _localParameters.Length - 1
                If index = 0 Then
                    Dim Val As Boolean = False
                    Boolean.TryParse(CStr(_localParameters(index)), Val)
                    Console.WriteLine(
                        vbTab & vbTab & vbTab & _parameterDictionary(index).ToString() & " = " & Val.ToString())
                Else
                    Dim Collection = DirectCast(_localParameters(index), List(Of String))
                    Dim Sb = New StringBuilder()
                    Dim First = True
                    For Each item In Collection
                        If First Then
                            Sb.Append(item)
                            First = False
                        Else
                            Sb.Append(" | " + item)
                        End If
                    Next
                    Console.WriteLine(vbTab & vbTab & vbTab & _parameterDictionary(index) & " = " & Sb.ToString())
                End If
            Next
            Console.WriteLine("")
        End Sub

        Public Overridable Sub CloseDialogs()
            Dim Procs() = Process.GetProcesses
            For Each Pr As Process In From Pr1 In Procs Where (Pr1.ProcessName = "OUTLOOK")
                For cdb = 1 To 3
                    Try
                        AppActivate(Pr.Id)
                        'TODO: SendKeys sends key strokes to the current active window.... Also, this is using Application of a windows forms app. We are in a Wpf. Find a more robust way
                        SendKeys.SendWait("{ESC}")
                        Thread.Sleep(50)
                        Forms.Application.DoEvents()
                    Catch Ex As System.Exception
                        Exit For
                    End Try
                Next
            Next
        End Sub

        Public Overridable Sub MaximizeApp()
            Try
                Dim MainExplorer = MyOfficeApp.ActiveWindow()
                If MainExplorer IsNot Nothing Then
                    MainExplorer.WindowState = OlWindowState.olMaximized
                End If
            Catch Ex As System.Exception
            End Try
        End Sub

        Public Overridable Sub MinimizeApp()
            Try
                Dim MainExplorer = MyOfficeApp.ActiveWindow()
                If MainExplorer IsNot Nothing Then
                    MainExplorer.WindowState = OlWindowState.olMinimized
                End If
            Catch Ex As System.Exception
            End Try
        End Sub

        Overridable Sub CloseApp()
            Try
                If MyOfficeApp IsNot Nothing Then
                    Dim OfInbox As Folder = MyOfficeApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                    If OfInbox IsNot Nothing Then
                        For Each Item As Object In OfInbox.Items
                            Item.Delete()
                        Next
                        Marshal.FinalReleaseComObject(MyOfficeApp)
                        MyOfficeApp = Nothing
                    End If

                    MyOfficeApp.Quit()
                End If
            Catch Ex As System.Exception
            End Try
        End Sub

        Overridable Sub FinalCloseApp()
            Try
                Dim OfInbox As Folder = MyOfficeApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                If OfInbox IsNot Nothing Then
                    For Each Item As Object In OfInbox.Items
                        Item.Delete()
                    Next
                    Marshal.FinalReleaseComObject(OfInbox)
                    OfInbox = Nothing
                End If
            Catch Ex As System.Exception
            End Try

            If MyOfficeApp IsNot Nothing Then
                MyOfficeApp.Quit()
                Marshal.FinalReleaseComObject(MyOfficeApp)
                MyOfficeApp = Nothing
            End If

        End Sub

        Private Shared Sub GarbageCollector()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Sub

#End Region

#Region "Functions"

        Public Overridable Function ValidateQuestion() As [Object]
            Return nothing
        End Function

        Public Overridable Function ValidateQuestion(ByVal ParamArray parameters As Object()) As [Object]
            If parameters Is Nothing Then
                If _localParameters Is Nothing Then
                    Console.WriteLine("The object array parameters is null. fill the parameters and try again.")
                Else
                    parameters = _localParameters
                End If
            End If

            ' -----------------------------------------------------------------------------------------------
            ' --------------------------  Place the validation code here  -----------------------------------
            ' -----------------------------------------------------------------------------------------------
            ' NOTE: Remove the Me in front of Me.MyOfficeApp

            Dim IsValidate = CType(parameters(0).ToString(), Boolean)

            Dim Templates = CType(parameters(1), List(Of String))
            Dim KeyWords = CType(parameters(2), List(Of String))
            Dim BlueWords = CType(parameters(3), List(Of String))

            If Not IsValidate Then ' Start Pre-Code
                Try
                    Return True
                Catch Ex As System.Exception
                    Return False
                End Try
            Else ' Start Validation-Code    

                Dim Points(0)
                Points(0) = False
                Try
                    Try
                    Catch Ex As System.Exception

                    End Try
                Catch Ex As System.Exception
                    Return Points
                End Try
                Return Points
            End If
        End Function

#End Region

#Region "Commands"
        Property SetupCommand() As ICommand
        Property RunValidationPreCodeCommand() As ICommand
        Property RunValidationPostCodeCommand() As ICommand
        Property EndCommand() As ICommand
        Property PromptToFillArrayCommand() As ICommand
        Property PrintParamsCommand() As ICommand
        Property CloseDialogsCommand() As ICommand
        Property MaximizeAppCommand() As ICommand
        Property MinimizeAppCommand() As ICommand
        Property CloseAppCommand() As ICommand
        Property FinalCloseAppCommand() As ICommand
        Property ValidateQuestionCommand() As ICommand
        Property GarbageCollectorCommand() As ICommand

#End Region
    End Class
End Namespace
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports GalaSoft.MvvmLight
Imports GalaSoft.MvvmLight.CommandWpf
Imports Microsoft.Office.Interop.Access
Imports Microsoft.Office.Interop.Access.Dao

Namespace ViewModel
    ''' <summary>
    '''     Class ValidationCodeRunnerViewModel.
    ''' </summary>
    ''' <seealso cref="GalaSoft.MvvmLight.ViewModelBase" />
    Public Class ValidationCodeRunnerViewModel
        Inherits ViewModelBase

        
#Region "Contructor"
        Public Sub New()
           ' ValidateQuestionCommand = New RelayCommand(ValidateQuestion())

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

        Private _myOfficeApp As Microsoft.Office.Interop.Access.Application
#End Region

#Region "Properties"
        Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.Access.Application
            Get
                Return _myOfficeApp
            End Get
            Set
                [Set]( _myOfficeApp, Value,True,"MyOfficeApp" )
               ' _myOfficeApp = Value
            End Set
        End Property
#End Region

#Region "Methods"
        Public Sub Setup()
            MyOfficeApp = New Microsoft.Office.Interop.Access.Application()

            If MyOfficeApp IsNot Nothing Then
                Console.WriteLine("Version: " + MyOfficeApp.Version)
                Console.WriteLine("Name: " + MyOfficeApp.Name)

                MyOfficeApp.Visible = True

                NativeMethods.SetWindowText(MyOfficeApp.hWndAccessApp(), "Did it work")

'                Try
'                    Dim HwndInt = MyOfficeApp.hWndAccessApp()
'                    Console.WriteLine("HWND = " & HwndInt)
'                Catch Ex As COMException
'                End Try

            End If
            Console.WriteLine("")
        End Sub

        Public Sub RunValidationPreCode()
            Console.WriteLine(vbCrLf & "Executing the pre Code. . .")

            If MyOfficeApp IsNot Nothing Then
                NativeMethods.SetWindowText(MyOfficeApp.hWndAccessApp(), "Did it work")

                MyOfficeApp.RefreshTitleBar()
            End If
        End Sub

        Function GetAccessApplicationHwnd() As IntPtr
            Return New IntPtr( MyOfficeApp.hWndAccessApp() )
        End Function

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
			Try
				Dim sessionId As Integer = Process.GetCurrentProcess.SessionId
				Dim processes = Process.GetProcesses().Where(Function(p) p.SessionId = sessionId AndAlso p.ProcessName.ToLower() = "MSACCESS".ToLower())
				For Each proc As Process In processes
					Try
						Microsoft.VisualBasic.AppActivate(proc.Id)
						System.Windows.Forms.SendKeys.Send("{ ESC}")
						System.Threading.Thread.Sleep(50)
						System.Windows.Forms.Application.DoEvents()
					Catch Ex As Exception
						'ErrorsCollection.Add("Exception occured in CloseDialogs.\n\tMessage: " + Ex.Message)
					End Try
				Next
			Catch Ex As Exception
				'ErrorsCollection.Add("Exception occured in CloseDialogs.\n\tMessage: " + Ex.Message)
			End Try
        End Sub

        Public Overridable Sub MaximizeApp()
            Try
                MyOfficeApp.Visible = True
                MyOfficeApp.RunCommand(AcCommand.acCmdAppMaximize)
            Catch Ex As System.Exception
                'ErrorsCollection.Add( "Exception in MaximizeApp.\n\tMessage: " + Ex.Message )
            End Try
        End Sub

        Public Overridable Sub MinimizeApp()
            Try
                MyOfficeApp.RunCommand(AcCommand.acCmdAppMinimize)
            Catch Ex As Exception
            End Try
        End Sub

        Overridable Sub CloseApp()
            Try
                If MyOfficeApp IsNot Nothing Then
                    If MyOfficeApp.CurrentObjectType <> AcObjectType.acDefault Then
                        MyOfficeApp.CloseCurrentDatabase()
                    End If
                    MyOfficeApp.Quit(Option:=Microsoft.Office.Interop.Access.AcQuitOption.acQuitSaveNone)
                End If
            Catch Ex As System.Exception
                'ErrorsCollection.Add("Exception occured in CloseApp while attempting to close the Office Application.\n\tMessage: " + Ex.Message)
            End Try
            Try
                If MyOfficeApp IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
                    MyOfficeApp = Nothing
                End If
            Catch Ex As System.Exception
                'ErrorsCollection.Add("Exception occured in CloseApp while attempting to realease the Office Application Com object Interface.\n\tMessage: " + Ex.Message)
            End Try
        End Sub

        Overridable Sub FinalCloseApp()
            			Try
				If MyOfficeApp IsNot Nothing Then
					If MyOfficeApp.CurrentObjectType <> AcObjectType.acDefault Then
					   MyOfficeApp.CloseCurrentDatabase()
					End If
				End If
			Catch Ex As System.Exception
				'ErrorsCollection.Add( "Exception in FinalCloseApp while attempting to close app.\n\tMessage: " + Ex.Message )
			End Try
			Try
				If MyOfficeApp IsNot Nothing Then
					MyOfficeApp.Quit()
					System.Runtime.InteropServices.Marshal.FinalReleaseComObject( MyOfficeApp )
					MyOfficeApp = Nothing
				End If
			Catch Ex As System.Exception
				'ErrorsCollection.Add( "Exception in FinalCloseApp while attemtping to realease Com object.\n\tMessage: " + Ex.Message )
			End Try


            If MyOfficeApp IsNot Nothing Then
                MyOfficeApp.Quit()
                Marshal.FinalReleaseComObject(MyOfficeApp)
                MyOfficeApp = Nothing
            End If
        End Sub

        Private Sub GarbageCollector()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Sub
#End Region

#Region "Functions"
        'Public Overridable Function ValidateQuestion(Optional ByVal ParamArray parameters As Object()) As [Object]
        Public Overridable Function ValidateQuestion() As [Object]
'            If parameters Is Nothing Then
'                If _localParameters Is Nothing Then
'                    Console.WriteLine("The object array parameters is null. fill the parameters and try again.")
'                Else
'                    parameters = _localParameters
'                End If
'            End If

            ' -----------------------------------------------------------------------------------------------
            ' --------------------------  Place the validation code here  -----------------------------------
            ' -----------------------------------------------------------------------------------------------
            ' NOTE: Remove the Me in front of Me.MyOfficeApp

            'Dim IsValidate = CType(parameters(0).ToString(), Boolean)

            'Dim Templates = CType(parameters(1), List(Of String))
            'Dim KeyWords = CType(parameters(2), List(Of String))
            'Dim BlueWords = CType(parameters(3), List(Of String))

            If Not true Then ' Start Pre-Code
                Try
                    Return True
                Catch Ex As Exception
                    Return False
                End Try
            Else ' Start Validation-Code    

                Dim Points(0)
                Points(0) = False
                Try
                Catch Ex As Exception
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




    Friend NotInheritable Class NativeMethods  
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
        Friend  Shared Function SetWindowText(ByVal hwnd As IntPtr, ByVal lpString As String) As Boolean
        End Function


    End Class
End Namespace
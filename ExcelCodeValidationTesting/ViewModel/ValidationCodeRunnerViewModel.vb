Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports GalaSoft.MvvmLight
Imports GalaSoft.MvvmLight.CommandWpf
Imports Microsoft.Office.Interop.Word

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
                    ValidateQuestionCommand = New RelayCommand(ValidateQuestion())

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

                Private _myOfficeApp As Microsoft.Office.Interop.Word.Application
        #End Region

        #Region "Properties"
                ''' <summary>
                '''     Gets or sets my office application.
                ''' </summary>
                ''' <value>My office application.</value>
                Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.Word.Application
                    Get
                        Return _myOfficeApp
                    End Get
                    Set
                        _myOfficeApp = Value
                    End Set
                End Property
        #End Region

        #Region "Methods"
                ''' <summary>
                '''     Setups this instance.
                ''' </summary>
                Public Sub Setup()
                    MyOfficeApp = New Microsoft.Office.Interop.Word.Application()

                    If MyOfficeApp IsNot Nothing Then
                        Console.WriteLine("Version: " + MyOfficeApp.Version)
                        Console.WriteLine("Name: " + MyOfficeApp.Name)
                        Console.WriteLine("Caption: " + MyOfficeApp.Caption)

                        '		if (System.Windows.Forms.MessageBox.Show("Open Inbox?", "Do Something?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        '		{
                        '			"Hit Done in MBox buddy!".Dump();		
                        '		}

                        '		If MyOfficeApp.ActiveWindow IsNot Nothing Then
                        '			Console.WriteLine("HWND = " & MyOfficeApp.ActiveWindow.Hwnd.ToString())
                        '		End If

                        'MyOfficeApp.GetType()
                        'Dim hwnd = MyOfficeApp.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, Nothing, MyOfficeApp, Nothing)
                        'Console.WriteLine("HWND = " & hwnd.ToString())
                        MyOfficeApp.Visible = True

                        Try
                            If MyOfficeApp.ActiveWindow IsNot Nothing Then
                                Console.WriteLine("HWND = " & MyOfficeApp.ActiveWindow.Hwnd.ToString())
                            End If
                        Catch Ex As COMException
                        End Try

                    End If
                    Console.WriteLine("")
                End Sub

                ''' <summary>
                '''     Runs the validation pre code.
                ''' </summary>
                Public Sub RunValidationPreCode()
                    Console.WriteLine(vbCrLf & "Executing the pre Code. . .")

                    If MyOfficeApp IsNot Nothing Then

                    End If
                End Sub

                ''' <summary>
                '''     Runs the validation post code.
                ''' </summary>
                Public Sub RunValidationPostCode()
                    Console.WriteLine(vbCrLf & "Executing the post Code. . .")

                    If MyOfficeApp IsNot Nothing Then

                    End If
                End Sub

                ''' <summary>
                '''     Ends this instance.
                ''' </summary>
                Public Sub [End]()
                    Console.WriteLine(vbCrLf & "Executing the termination Code. . .")

                    If MyOfficeApp IsNot Nothing Then
                        MyOfficeApp.Quit()
                        Marshal.FinalReleaseComObject(MyOfficeApp)
                        MyOfficeApp = Nothing
                    End If
                End Sub

                ''' <summary>
                '''     Prompts to fill array.
                ''' </summary>
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

                ''' <summary>
                '''     Prints the parameters.
                ''' </summary>
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

                ''' <summary>
                '''     Closes the dialogs.
                ''' </summary>
                Public Overridable Sub CloseDialogs()
                    Dim Procs() = Process.GetProcesses
                    Dim Pr As Process
                    For Each Pr In Procs
                        If (Pr.ProcessName = "WINWORD") Then
                            For cdb = 1 To 3
                                Try
                                    AppActivate(Pr.Id)
                                    'TODO: SendKeys sends key strokes to the current active window.... Also, this is using Application of a windows forms app. We are in a Wpf. Find a more robust way
                                    SendKeys.Send("{ESC}")
                                    Thread.Sleep(50)
                                    Forms.Application.DoEvents()
                                Catch Ex As Exception
                                    Exit For
                                End Try
                            Next
                        End If
                    Next
                End Sub

                ''' <summary>
                '''     Maximizes the application.
                ''' </summary>
                Public Overridable Sub MaximizeApp()
                    Try
                        MyOfficeApp.Visible = True
                        MyOfficeApp.WindowState = WdWindowState.wdWindowStateMaximize
                    Catch Ex As Exception
                    End Try
                End Sub

                ''' <summary>
                '''     Minimizes the application.
                ''' </summary>
                Public Overridable Sub MinimizeApp()
                    Try
                        MyOfficeApp.WindowState = WdWindowState.wdWindowStateMinimize
                    Catch Ex As Exception
                    End Try
                End Sub

                ''' <summary>
                '''     Closes the application.
                ''' </summary>
                Overridable Sub CloseApp()
                    Try
                        If MyOfficeApp IsNot Nothing Then
                            Dim Procs() = Process.GetProcesses
                            Dim Pr As Process
                            For Each Pr In Procs
                                If (Pr.ProcessName = "WINWORD") Then
                                    Dim B As Short
                                    B = MyOfficeApp.Documents.Count
                                    For a = 1 To B
                                        MyOfficeApp.ActiveDocument.Close(SaveChanges:=False)
                                        B = MyOfficeApp.Documents.Count
                                        If B < 1 Then
                                            Exit For
                                        End If
                                    Next
                                    Exit For
                                End If
                            Next
                            MyOfficeApp.Quit()
                            If MyOfficeApp IsNot Nothing Then
                                Marshal.FinalReleaseComObject(MyOfficeApp)
                                MyOfficeApp = Nothing
                            End If
                        End If
                    Catch Ex As Exception
                        ' Ignored
                    End Try
                End Sub

                ''' <summary>
                '''     Finals the close application.
                ''' </summary>
                Overridable Sub FinalCloseApp()
                    Try
                        Dim Procs() = Process.GetProcesses
                        Dim Pr As Process
                        For Each Pr In Procs
                            If (Pr.ProcessName = "WINWORD") Then
                                If MyOfficeApp IsNot Nothing Then
                                    Dim B As Int16
                                    B = MyOfficeApp.Documents.Count
                                    For a = 1 To B
                                        MyOfficeApp.ActiveDocument.Close(SaveChanges:=False)
                                        B = MyOfficeApp.Documents.Count
                                        If B < 1 Then
                                            Exit For
                                        End If
                                    Next
                                    Exit For
                                End If
                            End If
                        Next
                    Catch Ex As Exception
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
                ''' <summary>
                '''     Validates the question.
                ''' </summary>
                ''' <param name="parameters">The parameters.</param>
                ''' <returns>Object.</returns>
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
                            With MyOfficeApp
                                If Templates.Count >= 1 AndAlso Not String.IsNullOrEmpty(Templates(0)) Then
                                    .Documents.Open(Templates(0), [ReadOnly]:=False)
                                End If

                                .ActiveWindow.View.ReadingLayout = False
                                .Visible = True
                            End With

                            Try
                                _myOfficeApp.Options.PrintProperties = False
                                _myOfficeApp.Options.CheckSpellingAsYouType = True
                                _myOfficeApp.Options.AllowOpenInDraftView = False
                            Catch Ex As Exception
                                ' Ignored
                            End Try
                            Return True
                        Catch Ex As Exception
                            Return False
                        End Try
                    Else ' Start Validation-Code    

                        Dim Points(0)
                        Points(0) = False
                        Try
                            Try
                                If MyOfficeApp.Documents.Count > 0 Then
                                    If MyOfficeApp.ActiveDocument.Paragraphs.Count > 0 Then
                                        If MyOfficeApp.ActiveDocument.Paragraphs(2).Range.Font.Name = "Hey" Then
                                            Points(0) = True
                                        End If
                                    End If
                                End If
                            Catch Ex As Exception

                            End Try
                        Catch Ex As Exception
                            Return Points
                        End Try
                        Return Points
                    End If
                End Function
        #End Region

        #Region "Commands"
                ''' <summary>
                '''     Gets or sets the setup command.
                ''' </summary>
                ''' <value>The setup command.</value>
                Property SetupCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the run validation pre code command.
                ''' </summary>
                ''' <value>The run validation pre code command.</value>
                Property RunValidationPreCodeCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the run validation post code command.
                ''' </summary>
                ''' <value>The run validation post code command.</value>
                Property RunValidationPostCodeCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the end command.
                ''' </summary>
                ''' <value>The end command.</value>
                Property EndCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the prompt to fill array command.
                ''' </summary>
                ''' <value>The prompt to fill array command.</value>
                Property PromptToFillArrayCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the print parameters command.
                ''' </summary>
                ''' <value>The print parameters command.</value>
                Property PrintParamsCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the close dialogs command.
                ''' </summary>
                ''' <value>The close dialogs command.</value>
                Property CloseDialogsCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the maximize application command.
                ''' </summary>
                ''' <value>The maximize application command.</value>
                Property MaximizeAppCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the minimize application command.
                ''' </summary>
                ''' <value>The minimize application command.</value>
                Property MinimizeAppCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the close application command.
                ''' </summary>
                ''' <value>The close application command.</value>
                Property CloseAppCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the final close application command.
                ''' </summary>
                ''' <value>The final close application command.</value>
                Property FinalCloseAppCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the validate question command.
                ''' </summary>
                ''' <value>The validate question command.</value>
                Property ValidateQuestionCommand() As ICommand

                ''' <summary>
                '''     Gets or sets the garbage collector command.
                ''' </summary>
                ''' <value>The garbage collector command.</value>
                Property GarbageCollectorCommand() As ICommand
        #End Region
    End Class
End Namespace
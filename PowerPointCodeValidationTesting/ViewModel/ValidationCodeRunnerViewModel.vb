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
                   ValidateQuestionCommand = New RelayCommand(AddressOf ValidateQuestion)

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

                Private _myOfficeApp As Microsoft.Office.Interop.PowerPoint.Application
        #End Region

        #Region "Properties"
                ''' <summary>
                '''     Gets or sets my office application.
                ''' </summary>
                ''' <value>My office application.</value>
                Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.PowerPoint.Application
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
                    MyOfficeApp = New Microsoft.Office.Interop.PowerPoint.Application()

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
                        'MyOfficeApp.Visible = True

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
                End Sub

                ''' <summary>
                '''     Maximizes the application.
                ''' </summary>
                Public Overridable Sub MaximizeApp()
                End Sub

                ''' <summary>
                '''     Minimizes the application.
                ''' </summary>
                Public Overridable Sub MinimizeApp()
                End Sub

                ''' <summary>
                '''     Closes the application.
                ''' </summary>
                Overridable Sub CloseApp()
                End Sub

                ''' <summary>
                '''     Finals the close application.
                ''' </summary>
                Overridable Sub FinalCloseApp()
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
                Return new Object()
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

''' This is for a Project test
'Imports Microsoft.VisualBasic
'Imports System
'Imports System.Diagnostics
'Imports System.Text
'Imports System.IO
'Imports System.Environment
'Imports System.Xml
'Imports System.Windows.Forms
'Imports Microsoft.Office.Core
'Imports System.Collections.Generic
'Imports Microsoft.Win32
'Imports Microsoft.Office.Interop.PowerPoint

'Namespace GMetrix.Dynamic.Validations
   
'   Public Class DynamicQuestion
      
'      Private _myOfficeApp As Microsoft.Office.Interop.PowerPoint.Application
      
'      Public Sub New()
'         MyBase.New
'         Try
'Dim app As Microsoft.Office.Interop.PowerPoint.Application = Nothing
'app = DirectCast(System.Runtime.InteropServices.Marshal.GetActiveObject( "PowerPoint.Application" ), Microsoft.Office.Interop.PowerPoint.Application)
'If app IsNot Nothing Then
'    _myOfficeApp = app
'Else
'    _myOfficeApp = New Microsoft.Office.Interop.PowerPoint.Application()
'End If
'Catch ex As System.Runtime.InteropServices.COMException
''MessageBox.Show(ex.Message)
'_myOfficeApp = New Microsoft.Office.Interop.PowerPoint.Application()
'End Try

'         MaximizeApp()
'      End Sub
      
'      Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.PowerPoint.Application
'         Get
'            Return Me._myOfficeApp
'         End Get
'         Set
'            Me._myOfficeApp = value
'         End Set
'      End Property
      
'      Public Overridable Function Validate8706(ByVal ParamArray Parameters As Object()) As [Object]
'         '8706 Powerpoint Proj 1 V2
'Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
'Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
'Dim savedFile As String = String.Empty
'Dim DocsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal)

'Dim programingWords As List(Of String) = New List(Of String)

''Red Words
'Try
'    programingWords = CType(Parameters(5), List(Of String))
'	If programingWords.Count = 0 Then
'	  programingWords.Add("Eras Demi ITC") '0
'	  programingWords.Add("Century Gothic") '1
'	  programingWords.Add("WHY GO GREEN") '2
'	  programingWords.Add("Reduced Risk") '3
'	  programingWords.Add("HOW MUCH IT SAVES") '4
'	  programingWords.Add("Callout") ' 5
'	  programingWords.Add("Going Green Helps Everyone") '6
'	  programingWords.Add("Energy Star") '7
'	  programingWords.Add("When To Integrate") '8
'	  programingWords.Add("Eras Demi ITC") '9
'	  programingWords.Add("Save Electricity") '10
'	  programingWords.Add("Hybrid Car") '11
'	  programingWords.Add("Medium Style 2 - Accent 5") '12
'	  programingWords.Add("A Better World") '13
'	  programingWords.Add("Bird's Eye Scene") '14
'	  programingWords.Add("Colored Fill - Accent 4") '15
'	  programingWords.Add("Going Green Philosophy") '16
'	End If
'Catch Ex As Exception

'Try
    
'Catch
'End Try

'End Try

'            If Not isValidate Then ' Start Pre-Code
'                Try
'                    savedFile = Parameters(2).ToString()                 
'                    With Me.MyOfficeApp
'                        If Not String.IsNullOrEmpty(savedFile) Then
'								Me.MyOfficeApp.Presentations.Open(savedFile, ReadOnly:=False)
'                        Else

'                            Dim startingDoc As String
'                            If Not templates Is Nothing Then
'                                startingDoc = templates(0)
'                            Else
'                               startingDoc = DocsPath + "\GMetrixTemplates\Taking Your Company Green.pptx"
'                            End If
							
'                            .Presentations.Open(startingDoc)
'                            If programingWords(5) = "Llamada" Then
'                            	Dim fic As String = DocsPath + "\GMetrixTemplates\energy star.jpg"
'			                        If File.Exists(fic) Then
'			                            My.Computer.FileSystem.RenameFile(fic, "estrella de energía.jpg")
'			                        End If
'                            End If
'                        End If						
'                        .Visible = True
'                    End With

'						Me._myOfficeApp.ActivePresentation.PrintOptions.HighQuality = MsoTriState.msoFalse 
'						Me._myOfficeApp.DisplayGridLines = MsoTriState.msoFalse 
'						Me._myOfficeApp.ActivePresentation.PrintOptions.PrintFontsAsGraphics = MsoTriState.msoFalse
'						Me._myOfficeApp.ActivePresentation.SlideShowSettings.ShowType = PpSlideShowType.ppShowTypeWindow
						
'                    Return True
'                Catch ex As Exception
'                    Return False
'                End Try

'            Else ' Start Validation-Code    
'                Dim points(2)
'                Dim i As Integer = 0
'                For i = 0 To 2
'                    points(i) = False
'                Next i



'                'Design - Wisp Create a Presentation  519
'                Try
'                    If Me._myOfficeApp.ActivePresentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeLedgerPaper Then
'                        points(0) = True
'                    End If
'                Catch ex As Exception

'                End Try

'                ' Slide Master     Add Date to Footer 520
'                Try					
'                    If Me._myOfficeApp.ActivePresentation.SlideMaster.HeadersFooters.DisplayOnTitleSlide = MsoTriState.msoFalse Then
						
'                        If Me._myOfficeApp.ActivePresentation.SlideMaster.HeadersFooters.DateAndTime.Visible = MsoTriState.msoTrue Then
'                            points(1) = True
'                        End If
'                    End If
'                Catch ex As Exception
					
'                End Try
'                ' Modify Slide Master Using Slide Masters to apply default text styles
'                Try	
					
'                    If Me._myOfficeApp.ActivePresentation.SlideMaster.TextStyles(2).TextFrame.TextRange.Font.Name = programingWords(0) Then
						
'                        If Me._myOfficeApp.ActivePresentation.SlideMaster.TextStyles(2).TextFrame.TextRange.Font.Size = 36.0 Then
							
'                            If Me._myOfficeApp.ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange.Font.Name = programingWords(1) Then
								
'                                If Me._myOfficeApp.ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange.Font.Size = 18.0 Then
									
'                                    If Me._myOfficeApp.ActivePresentation.SlideMaster.CustomLayouts(1).Shapes(1).TextFrame.TextRange.Font.Size = 54.0 Then
'                                        points(2) = True
'                                    End If
'                                End If
'                            End If
'                        End If
'                    End If
'                Catch ex As Exception
					
'                End Try

'                Return points
'            End If
'      End Function
      
'      Public Overridable Function Validate8708(ByVal ParamArray Parameters As Object()) As [Object]
'         '8708
'Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
'Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
'Dim savedFile As String = String.Empty

''Red Words
'Dim programingWords As List(Of String) = CType(Parameters(5), List(Of String)) 
'If programingWords.Count = 0 Then
'  programingWords.Add("Eras Demi ITC") '0
'  programingWords.Add("Century Gothic") '1
'  programingWords.Add("WHY GO GREEN") '2
'  programingWords.Add("Reduced Risk") '3
'  programingWords.Add("HOW MUCH IT SAVES") '4
'  programingWords.Add("Callout") ' 5
'  programingWords.Add("Going Green Helps Everyone") '6
'  programingWords.Add("Energy Star") '7
'  programingWords.Add("When To Integrate") '8
'  programingWords.Add("Eras Demi ITC") '9
'  programingWords.Add("Save Electricity") '10
'  programingWords.Add("Hybrid Car") '11
'  programingWords.Add("Medium Style 2 - Accent 5") '12
'  programingWords.Add("A Better World") '13
'  programingWords.Add("Bird's Eye Scene") '14
'  programingWords.Add("Colored Fill - Accent 4") '15
'  programingWords.Add("Going Green Philosophy") '16
'End If

'            If Not isValidate Then ' Start Pre-Code
'                Try
'                    savedFile = Parameters(2).ToString()

'                    If Not String.IsNullOrEmpty(savedFile) Then
'                        Me.MyOfficeApp.Presentations.Open(savedFile, ReadOnly:=False)
'                    Else
'                        Me.MyOfficeApp.Presentations.Add()
'                    End If

'                    Return True
'                Catch ex As Exception
'                    Return False
'                End Try

'            Else ' Start Validation-Code    
'                Dim points(1)
'                Dim i As Integer = 0
'                For i = 0 To 1
'                    points(i) = False
'                Next i

'                ' Slide 2
'                Try
'                    'Title
					
'                    If InStr(Me._myOfficeApp.ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Text.ToUpper, programingWords(2).ToUpper) Then
'                        points(0) = True
'                    End If

'            'Bullet Types
					
'                    If InStr(Me._myOfficeApp.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange.Text.ToUpper, programingWords(3).ToUpper) Then
						
'                        If Me._myOfficeApp.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange.ParagraphFormat.Bullet.Character = 118 Then
'                            points(1) = True
'                        End If

'                    End If
'                Catch ex As Exception
					
'                End Try

'                Return points
'            End If
'      End Function
      
'      Public Overridable Function Validate8709(ByVal ParamArray Parameters As Object()) As [Object]
'         '8709
'Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
'Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
'Dim savedFile As String = String.Empty

''Red Words
'Dim programingWords As List(Of String) = CType(Parameters(5), List(Of String)) 
'If programingWords.Count = 0 Then
'  programingWords.Add("Eras Demi ITC") '0
'  programingWords.Add("Century Gothic") '1
'  programingWords.Add("WHY GO GREEN") '2
'  programingWords.Add("Reduced Risk") '3
'  programingWords.Add("HOW MUCH IT SAVES") '4
'  programingWords.Add("Callout") ' 5
'  programingWords.Add("Going Green Helps Everyone") '6
'  programingWords.Add("Energy Star") '7
'  programingWords.Add("When To Integrate") '8
'  programingWords.Add("Eras Demi ITC") '9
'  programingWords.Add("Save Electricity") '10
'  programingWords.Add("Hybrid Car") '11
'  programingWords.Add("Medium Style 2 - Accent 5") '12
'  programingWords.Add("A Better World") '13
'  programingWords.Add("Bird's Eye Scene") '14
'  programingWords.Add("Colored Fill - Accent 4") '15
'  programingWords.Add("Going Green Philosophy") '16
'End If

'            If Not isValidate Then ' Start Pre-Code
'                Try
'                    savedFile = Parameters(2).ToString()

'                    If Not String.IsNullOrEmpty(savedFile) Then
'                        Me.MyOfficeApp.Presentations.Open(savedFile, ReadOnly:=False)
'                    Else
'                        Me.MyOfficeApp.Presentations.Add()
'                    End If

'                    Return True
'                Catch ex As Exception
'                    Return False
'                End Try

'            Else ' Start Validation-Code
'                Dim points(3)
'                Dim i As Integer = 0
'                For i = 0 To 3
'                    points(i) = False
'                Next i



'                'Slide 3
'                Try
'                    'Title - Subheading 2 and Indent Level
'                    If InStr(Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(4).TextFrame.TextRange.Text.ToUpper, programingWords(4)) Then
'                        If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(3).TextFrame.TextRange.IndentLevel = 1 Then
'                            points(0) = True
'                        End If
'                    End If
'                Catch ex As Exception

'                End Try

'                Try
'                    'Chart Data And Style
'                    If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).HasChart = MsoTriState.msoTrue Then
'                        Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.ChartData.Activate()
'                        If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.ChartData.Workbook.Worksheets(1).Range("B2").Value2 = 6.0 Then
'                            If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.ChartData.Workbook.Worksheets(1).Range("C5").Value2 = 2.0 Then
'                                points(1) = True
'                                If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.ChartStyle = 209 Then
'                                    points(2) = True
'                                End If
'                            End If
'                        End If
'                    End If

'                Catch ex As Exception
'                End Try


'                'Slide 3
'                Try
'                    'Chart Layout
'                    If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.HasLegend = True Then
'                        If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.Legend.Position = Microsoft.Office.Interop.PowerPoint.XlLegendPosition.xlLegendPositionRight Then
'                            If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.HasTitle = True Then
'                                If Me._myOfficeApp.ActivePresentation.Slides(3).Shapes(5).Chart.HasTitle = True Then
'                                    points(3) = True
'                                End If
'                            End If
'                        End If
'                    End If

'                Catch ex As Exception

'                End Try

'                Return points
'            End If
'      End Function
      
'      Public Overridable Function Validate8710(ByVal ParamArray Parameters As Object()) As [Object]
'         '8710
'Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
'Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
'Dim savedFile As String = String.Empty

''Red Words
'Dim programingWords As List(Of String) = CType(Parameters(5), List(Of String)) 
'If programingWords.Count = 0 Then
'  programingWords.Add("Eras Demi ITC") '0
'  programingWords.Add("Century Gothic") '1
'  programingWords.Add("WHY GO GREEN") '2
'  programingWords.Add("Reduced Risk") '3
'  programingWords.Add("HOW MUCH IT SAVES") '4
'  programingWords.Add("Callout") ' 5
'  programingWords.Add("Going Green Helps Everyone") '6
'  programingWords.Add("Energy Star") '7
'  programingWords.Add("When To Integrate") '8
'  programingWords.Add("Eras Demi ITC") '9
'  programingWords.Add("Save Electricity") '10
'  programingWords.Add("Hybrid Car") '11
'  programingWords.Add("Medium Style 2 - Accent 5") '12
'  programingWords.Add("A Better World") '13
'  programingWords.Add("Bird's Eye Scene") '14
'  programingWords.Add("Colored Fill - Accent 4") '15
'  programingWords.Add("Going Green Philosophy") '16
'End If

'            If Not isValidate Then ' Start Pre-Code
'                Try
'                    savedFile = Parameters(2).ToString()

'                    If Not String.IsNullOrEmpty(savedFile) Then
'                        Me.MyOfficeApp.Presentations.Open(savedFile, ReadOnly:=False)
'                    Else
'                        Me.MyOfficeApp.Presentations.Add()
'                    End If

'                    Return True
'                Catch ex As Exception
'                    Return False
'                End Try

'            Else ' Start Validation-Code
'                Dim points(4)
'                Dim i As Integer = 0
'                For i = 0 To 4
'                    points(i) = False
'                Next i


'                'Slide 4

'                Try
'                    For Each shp As Microsoft.Office.Interop.PowerPoint.Shape In Me._myOfficeApp.ActivePresentation.Slides(4).Shapes
'                        'Insert Callout
'                        Try							
'                            If InStr(shp.Name, programingWords(5)) Or Instr(shp.Name, "Callout") Then 'Hardcoded "Callout" be...s.Add("Hybrid Car") '11
'  programingWords.Add("Medium Style 2 - Accent 5") '12
'  programingWords.Add("A Better World") '13
'  programingWords.Add("Bird's Eye Scene") '14
'  programingWords.Add("Colored Fill - Accent 4") '15
'  programingWords.Add("Going Green Philosophy") '16
'End If
'            If Not isValidate Then ' Start Pre-Code
'                Try
'                    savedFile = Parameters(2).ToString()

'                    If Not String.IsNullOrEmpty(savedFile) Then
'                        Me.MyOfficeApp.Presentations.Open(savedFile, ReadOnly:=False)
'                    Else
'                        Me.MyOfficeApp.Presentations.Add()
'                    End If

'                    Return True
'                Catch ex As Exception
'                    Return False
'                End Try

'            Else ' Start Validation-Code  
'                    Dim points(3)
'                    Dim i As Integer = 0
'                    For i = 0 To 3
'                        points(i) = False
'                    Next i



'                'Slide 7
'                    Try
'                    'Insert Text
'                    If Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).SmartArt.Nodes(3).TextFrame2.TextRange.Text.ToUpper = programingWords(13).ToUpper Then
'                        points(0) = True
'                    End If

'                    'SmartArt Scene					
'                    If Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).SmartArt.QuickStyle.Name = programingWords(14) Then
'                        points(1) = True
'                    End If

'                    ' SmartArt Style					
'                    If Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).SmartArt.Color.Description = programingWords(15) Or Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).SmartArt.Color.Description = "Colored Fill - Accent 4" Then
'                        points(2) = True
'                    End If
'                    'SmartArt Size
'                    If Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).Height > 425 And Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).Height < 427 Then
'                        If Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).Width > 639 And Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).Width < 641 Then
'                            points(3) = True
'                        End If
'                    End If

'                Catch ex As Exception
					
'                End Try
'                    Return points
'                End If
'      End Function
      
'      Public Overridable Function Validate8713(ByVal ParamArray Parameters As Object()) As [Object]
'         '8713 Has no Red Words
'Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
'Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
'Dim savedFile As String = String.Empty

''Red Words
'Dim programingWords As List(Of String) = CType(Parameters(5), List(Of String)) 
'If programingWords.Count = 0 Then
'  programingWords.Add("Eras Demi ITC") '0
'  programingWords.Add("Century Gothic") '1
'  programingWords.Add("WHY GO GREEN") '2
'  programingWords.Add("Reduced Risk") '3
'  programingWords.Add("HOW MUCH IT SAVES") '4
'  programingWords.Add("Callout") ' 5
'  programingWords.Add("Going Green Helps Everyone") '6
'  programingWords.Add("Energy Star") '7
'  programingWords.Add("When To Integrate") '8
'  programingWords.Add("Eras Demi ITC") '9
'  programingWords.Add("Save Electricity") '10
'  programingWords.Add("Hybrid Car") '11
'  programingWords.Add("Medium Style 2 - Accent 5") '12
'  programingWords.Add("A Better World") '13
'  programingWords.Add("Bird's Eye Scene") '14
'  programingWords.Add("Colored Fill - Accent 4") '15
'  programingWords.Add("Going Green Philosophy") '16
'End If

'            If Not isValidate Then ' Start Pre-Code
'                Try
'                    savedFile = Parameters(2).ToString()

'                    If Not String.IsNullOrEmpty(savedFile) Then
'                        Me.MyOfficeApp.Presentations.Open(savedFile, ReadOnly:=False)
'                    Else
'                        Me.MyOfficeApp.Presentations.Add()
'                    End If

'                    Return True
'                Catch ex As Exception
'                    Return False
'                End Try

'            Else 'Start Validation-Code    
'                Dim points(8)
'                Dim i As Integer = 0
'                For i = 0 To 8
'                    points(i) = False
'                Next i




'                'Wind Transition On All Slides
'                Try
'				'Added Or condition for Right to left languages
'					If Me._myOfficeApp.ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectWindRight Or Me._myOfficeApp.ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectWindLeft Then						
'                        If Me._myOfficeApp.ActivePresentation.Slides(5).SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectWindRight Or Me._myOfficeApp.ActivePresentation.Slides(5).SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectWindLeft Then
'                            If Me._myOfficeApp.ActivePresentation.Slides(1).SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectWindRight Or Me._myOfficeApp.ActivePresentation.Slides(1).SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectWindLeft Then
'                                points(0) = True
'                            End If
'                        End If
'                    End If

'                Catch ex As Exception
					
'                End Try


'                'Apply Animation - No Delay
'                Try
'                    If Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(1).AnimationSettings.EntryEffect = PpEntryEffect.ppEffectCut Then
'                        points(1) = True
'                    End If
'                Catch ex As Exception

'                End Try

'                'Apply Animation With Delay
'                Try
					
'                    If Me._myOfficeApp.ActivePresentation.Slides(7).Shapes(2).AnimationSettings.EntryEffect = PpEntryEffect.ppEffectCut Then
						
'                        If Me._myOfficeApp.ActivePresentation.Slides(7).TimeLine.MainSequence(2).Timing.TriggerDelayTime = 1.0 Then
							
'                            If Me._myOfficeApp.ActivePresentation.Slides(7).TimeLine.MainSequence(2).Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerAfterPrevious Then
'                                points(2) = True
'                            End If
'                        End If
'                    End If
'                Catch ex As Exception

'                End Try


'                Try
'                    'Print Truetype Fonts as graphics					
'                    If Me._myOfficeApp.ActivePresentation.PrintOptions.PrintFontsAsGraphics = MsoTriState.msoTrue Then
'                        points(3) = True
'                    End If

'                    'Advance Slideshow Manually
'                    If Me._myOfficeApp.ActivePresentation.SlideShowSettings.AdvanceMode = PpSlideShowAdvanceMode.ppSlideShowManualAdvance Then
'                        points(4) = True
'                    End If
'                    'Display Gridlines and Rulers
					
'                    If Me._myOfficeApp.DisplayGridLines = MsoTriState.msoTrue Then
'                        points(5) = True
'                    End If

'                    ' Set print greyscale
					
'                    If Me._myOfficeApp.ActivePresentation.PrintOptions.PrintColorType = PpPrintColorType.ppPrintBlackAndWhite Then
'                        points(6) = True
'                    End If

'                    'Loop continuously until 'Esc'
'                    If Me._myOfficeApp.ActivePresentation.SlideShowSettings.LoopUntilStopped Then
'                        points(7) = True
'                    End If

'                    'Slideshow Settings
'                    If Me._myOfficeApp.ActivePresentation.SlideShowSettings.ShowType = PpSlideShowType.ppShowTypeSpeaker Then
'                        points(8) = True
'                    End If

'                Catch ex As Exception

'                End Try


'                Return points
'            End If
'      End Function
      
'      Public Overridable Function Validate8736(ByVal ParamArray Parameters As Object()) As [Object]
'         '8736
'Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
'Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
'Dim savedFile As String = String.Empty

''Red Words
'Dim programingWords As List(Of String) = CType(Parameters(5), List(Of String)) 
'If programingWords.Count = 0 Then
'  programingWords.Add("Eras Demi ITC") '0
'  programingWords.Add("Century Gothic") '1
'  programingWords.Add("WHY GO GREEN") '2
'  programingWords.Add("Reduced Risk") '3
'  programingWords.Add("HOW MUCH IT SAVES") '4
'  programingWords.Add("Callout") ' 5
'  programingWords.Add("Going Green Helps Everyone") '6
'  programingWords.Add("Energy Star") '7
'  programingWords.Add("When To Integrate") '8
'  programingWords.Add("Eras Demi ITC") '9
'  programingWords.Add("Save Electricity") '10
'  programingWords.Add("Hybrid Car") '11
'  programingWords.Add("Medium Style 2 - Accent 5") '12
'  programingWords.Add("A Better World") '13
'  programingWords.Add("Bird's Eye Scene") '14
'  programingWords.Add("Colored Fill - Accent 4") '15
'  programingWords.Add("Going Green Philosophy") '16
'End If

'            If Not isValidate Then ' Start Pre-Code
'                Try
'                    savedFile = Parameters(2).ToString()

'                    If Not String.IsNullOrEmpty(savedFile) Then
'                        Me.MyOfficeApp.Presentations.Open(savedFile, ReadOnly:=False)
'                    Else
'                        Me.MyOfficeApp.Presentations.Add()
'                    End If

'                    Return True
'                Catch ex As Exception
'                    Return False
'                End Try

'            Else ' Start Validation-Code  
'                Dim points(2)
'                    Dim i As Integer = 0
'                For i = 0 To 2
'                    points(i) = False
'                Next i

'                'Slide 6
'                Try

'                    'Insert Slide from Outline And change layout					
'                    If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(1).TextFrame.TextRange.Text = programingWords(16) Then						
'                        If Me._myOfficeApp.ActivePresentation.Slides(6).Layout = PpSlideLayout.ppLayoutContentWithCaption Then
'                            points(0) = True
'                        End If

'                    End If
'                Catch ex As Exception
'                End Try
'                Try
'                    'Add a Video
'                    If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(2).Type = MsoShapeType.msoMedia Then
'                        If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(2).MediaType = PpMediaType.ppMediaTypeMovie Then
'                            points(1) = True
'                        End If
'                    End If
'                Catch ex As Exception

'                End Try

'                Try
'                    'Format a Video
'                    If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(2).Shadow.Blur = 15.0 Then
'                        If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(2).Shadow.Obscured = MsoTriState.msoFalse Then
'                            If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(2).Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow Then
'                                If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(2).AnimationSettings.PlaySettings.LoopUntilStopped = MsoTriState.msoTrue Then
'                                    If Me._myOfficeApp.ActivePresentation.Slides(6).Shapes(2).AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue Then
'                                        points(2) = True
'                                    End If
'                                End If
'                            End If
'                        End If
'                    End If
'                Catch ex As Exception

'                End Try
                
'                'Adding these lines of code to prevent save pop-ups
'                Try
'	                Me.MyOfficeApp.DisplayAlerts = False
'	    			Me.MyOfficeApp.ActivePresentation.Saved = True
'	    		Catch
'	    		End Try
	    		
'                Return points
'            End If
'      End Function
      
'      Public Overridable Sub CloseApp()
'         Try
'    If MyOfficeApp IsNot Nothing Then
'        MyOfficeApp.ActivePresentation.Close()
'        MyOfficeApp.WindowState = PpWindowState.ppWindowMinimized
'        MyOfficeApp.Visible = False
'        MyOfficeApp.Quit()
'        If MyOfficeApp IsNot Nothing Then
'            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
'            MyOfficeApp = Nothing
'        End If
'    End If
'Catch ex As System.Exception
'End Try

'      End Sub
      
'      Public Overridable Sub CloseDialogs()
'         Dim procs() = Process.GetProcesses
'Dim pr As Process
'For Each pr In procs
'    If (pr.ProcessName = "POWERPNT") Then
'        Dim cdb As Integer
'        For cdb = 1 To 3
'        Try
'            AppActivate(pr.Id)
'            System.Windows.Forms.SendKeys.Send("{ESC}")
'            System.Threading.Thread.Sleep(50)
'            System.Windows.Forms.Application.DoEvents()
'        Catch ex As System.Exception
'            Exit For
'        End Try
'        Next
'    End If
'Next

'      End Sub
      
'      Public Overridable Sub MaximizeApp()
'         Try
'    Me.MyOfficeApp.Visible = True
'    Me.MyOfficeApp.WindowState = PpWindowState.ppWindowMaximized
'Catch Ex As System.Exception
'End Try

'      End Sub
      
'      Private Sub GarbageCollector()
'         GC.Collect()
'GC.WaitForPendingFinalizers()
'GC.Collect()
'GC.WaitForPendingFinalizers()

'      End Sub
      
'      Private Function ExistsDirectory(ByVal path As [String]) As [Boolean]
'         Try
'If (Not Directory.Exists(path)) Then
'Directory.CreateDirectory(path)
'End If
'Return Directory.Exists(path)
'Catch ex As System.Exception
'Return False
'End Try

'      End Function
      
'      Private Function ExistsFile(ByVal file As [String]) As [Boolean]
'         Try
'    Return IO.File.Exists(file)
'Catch ex As System.Exception
'    Return False
'End Try

'      End Function
      
'      Public Overridable Sub SaveAs(ByVal ParamArray Parameters As Object())
'         Dim filePath As String = Parameters(0).ToString()
'Try
'    If MyOfficeApp Is Nothing Then
'        Return
'    End If
'    Me.MyOfficeApp.DisplayAlerts = False
'    With Me.MyOfficeApp.ActivePresentation
'        .SaveAs( FileName:=filePath, FileFormat:=PpSaveAsFileType.ppSaveAsOpenXMLPresentation )
'    End With
'    Me.MyOfficeApp.DisplayAlerts = True
'Catch ex As System.Exception
'    Me.MyOfficeApp.DisplayAlerts = True
'    MessageBox.Show(ex.Message)
'End Try

'      End Sub
      
'      Public Overridable Sub MinimizeApp()
'         Try
'   Try
'       MyOfficeApp.ActivePresentation.Close()
'   Catch ex As System.Exception
'   End Try
'   Try
'        Me.MyOfficeApp.WindowState = PpWindowState.ppWindowMinimized
'   Catch ex As System.Exception
'       msgbox(ex.Message)
'   End Try
'Catch ex As System.Exception
'End Try

'      End Sub
      
'      Public Overridable Sub FinalCloseApp()
'         Try
'    Dim procs() = Process.GetProcesses
'    Dim pr As Process
'    For Each pr In procs
'        If (pr.ProcessName = "POWERPNT") Then
'            If MyOfficeApp IsNot Nothing Then
'                Dim a, b As Int16
'                b = MyOfficeApp.Presentations.Count
'                For a = 1 To b
'                    MyOfficeApp.ActivePresentation.Final = False
'                    MyOfficeApp.ActivePresentation.Save()
'                    MyOfficeApp.ActivePresentation.Close()
'                    b = MyOfficeApp.Presentations.Count
'                    If b < 1 Then
'                        Exit For
'                    End If
'                Next
'                Exit For
'            End If
'        End If
'    Next
'Catch Ex As System.Exception
'End Try

'      End Sub
'   End Class
'End Namespace














''' This is for a Question Based Test 
'Imports Microsoft.VisualBasic
'Imports System
'Imports System.Diagnostics
'Imports System.Text
'Imports System.IO
'Imports System.Environment
'Imports System.Xml
'Imports System.Windows.Forms
'Imports Microsoft.Office.Core
'Imports System.Collections.Generic
'Imports Microsoft.Win32
'Imports Microsoft.Office.Interop.PowerPoint

'Namespace GMetrix.Dynamic.Validations
   
'   Public Class DynamicQuestion
      
'      Private _myOfficeApp As Microsoft.Office.Interop.PowerPoint.Application
      
'      Public Sub New()
'         MyBase.New
'         Try
'Dim app As Microsoft.Office.Interop.PowerPoint.Application = Nothing
'app = DirectCast(System.Runtime.InteropServices.Marshal.GetActiveObject( "PowerPoint.Application" ), Microsoft.Office.Interop.PowerPoint.Application)
'If app IsNot Nothing Then
'    _myOfficeApp = app
'Else
'    _myOfficeApp = New Microsoft.Office.Interop.PowerPoint.Application()
'End If
'Catch ex As System.Runtime.InteropServices.COMException
''MessageBox.Show(ex.Message)
'_myOfficeApp = New Microsoft.Office.Interop.PowerPoint.Application()
'End Try

'         MaximizeApp()
'      End Sub
      
'      Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.PowerPoint.Application
'         Get
'            Return Me._myOfficeApp
'         End Get
'         Set
'            Me._myOfficeApp = value
'         End Set
'      End Property
      
'      Public Overridable Function ValidateQuestion(ByVal ParamArray Parameters As Object()) As [Object]
'         'Juan 06/26/2013
'        Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
'        Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
'        Dim keyWords As List(Of String) = CType(Parameters(2), List(Of String))
'        Dim blueWords As List(Of String) = CType(Parameters(3), List(Of String))

'        If Not isValidate Then ' Start Pre-Code
'            Try
'                Me.MyOfficeApp.Presentations.Open(templates(0))
'                Me.MyOfficeApp.Visible = True

'                Return True
'            Catch ex As Exception
'                Return False
'            End Try
'        Else ' Start Validation-Code    

'            Dim points(0)
'            points(0) = False
'            Try
'                'Verificar el tamaño de la presentación
'                If InStr(Me.MyOfficeApp.ActivePresentation.PageSetup.SlideHeight, "720") Then
'                    If InStr(Me.MyOfficeApp.ActivePresentation.PageSetup.SlideWidth, "540") Then
'                        If Me.MyOfficeApp.ActivePresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationVertical Then
'                            points(0) = True
'                        End If
'                    End If
'                End If
               
'            Catch ex As Exception
'                Return points
'            End Try

'            Return points
'        End If
'      End Function
      
'      Public Overridable Sub CloseDialogs()
'         Dim procs() = Process.GetProcesses
'Dim pr As Process
'For Each pr In procs
'    If (pr.ProcessName = "POWERPNT") Then
'        Dim cdb As Integer
'        For cdb = 1 To 3
'        Try
'            AppActivate(pr.Id)
'            System.Windows.Forms.SendKeys.Send("{ESC}")
'            System.Threading.Thread.Sleep(50)
'            System.Windows.Forms.Application.DoEvents()
'        Catch ex As System.Exception
'            Exit For
'        End Try
'        Next
'    End If
'Next

'      End Sub
      
'      Public Overridable Sub MaximizeApp()
'         Try
'    Me.MyOfficeApp.Visible = True
'    Me.MyOfficeApp.WindowState = PpWindowState.ppWindowMaximized
'Catch Ex As System.Exception
'End Try

'      End Sub
      
'      Public Overridable Sub MinimizeApp()
'         Try
'   Try
'       MyOfficeApp.ActivePresentation.Close()
'   Catch ex As System.Exception
'   End Try
'   Try
'        Me.MyOfficeApp.WindowState = PpWindowState.ppWindowMinimized
'   Catch ex As System.Exception
'       msgbox(ex.Message)
'   End Try
'Catch ex As System.Exception
'End Try

'      End Sub
      
'      Public Overridable Sub CloseApp()
'         Try
'    If MyOfficeApp IsNot Nothing Then
'        MyOfficeApp.ActivePresentation.Close()
'        MyOfficeApp.WindowState = PpWindowState.ppWindowMinimized
'        MyOfficeApp.Visible = False
'        MyOfficeApp.Quit()
'        If MyOfficeApp IsNot Nothing Then
'            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
'            MyOfficeApp = Nothing
'        End If
'    End If
'Catch ex As System.Exception
'End Try

'      End Sub
      
'      Public Overridable Sub FinalCloseApp()
'         Try
'    Dim procs() = Process.GetProcesses
'    Dim pr As Process
'    For Each pr In procs
'        If (pr.ProcessName = "POWERPNT") Then
'            If MyOfficeApp IsNot Nothing Then
'                Dim a, b As Int16
'                b = MyOfficeApp.Presentations.Count
'                For a = 1 To b
'                    MyOfficeApp.ActivePresentation.Final = False
'                    MyOfficeApp.ActivePresentation.Save()
'                    MyOfficeApp.ActivePresentation.Close()
'                    b = MyOfficeApp.Presentations.Count
'                    If b < 1 Then
'                        Exit For
'                    End If
'                Next
'                Exit For
'            End If
'        End If
'    Next
'Catch Ex As System.Exception
'End Try

'      End Sub
      
'      Private Sub GarbageCollector()
'         GC.Collect()
'GC.WaitForPendingFinalizers()
'GC.Collect()
'GC.WaitForPendingFinalizers()

'      End Sub
'   End Class
'End Namespace

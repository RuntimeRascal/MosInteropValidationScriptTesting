Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports GalaSoft.MvvmLight
Imports GalaSoft.MvvmLight.CommandWpf
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Namespace ViewModel
    ''' <summary>
    '''     Class ValidationCodeRunnerViewModel.
    ''' </summary>
    ''' <seealso cref="GalaSoft.MvvmLight.ViewModelBase" />
    Public Class ValidationCodeRunnerViewModel
        Inherits ViewModelBase

#Region "Contructor"
        Public Sub New()
            ValidateQuestionCommand = New RelayCommand(AddressOf ValidateQuestion )

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

        Dim _localParameters As Object() = New Object() {False, ' Compiler result (Pre-Code = False or Post-Code) True or False
                                                         New List(Of String) From {"C:\Users\Tommy\Documents\GMetrixTemplates\VendasFT.xlsx"},
                                                         New List(Of String) From{ "C:\Users\Tommy\AppData\Roaming\GSavedProjects\1206997_19772545.xlsx"}, ' OrangeWords
                                                         New List(Of String), ' BlueWords
                                                         New List(Of String) , ' RedWords
                                                         New List(Of String)From { "Dados de Vendas",
                                                                                    "Totais do Trimestre",
                                                                                    "Registro do Funcionário",
                                                                                    "Adrian Parmalee",
                                                                                    "Mandrake Wilson",
                                                                                    "Víctor French",
                                                                                    "SOMA",
                                                                                    "TotaisTrimestre",
                                                                                    "='Dados de Vendas'!$B$16:$E$16",
                                                                                    "MAIOR VENDA",
                                                                                    "MAX",
                                                                                    "VENDAS POR VENDEDOR",
                                                                                    "Diminuir",
                                                                                    "Aumentar",  
                                                                                    "Registro do Funcionário",
                                                                                    "COPYRIGHT FUSION TOMO, TODOS OS DIREITOS RESERVADOS",
                                                                                    "CONTEÚDO CORRIGIDO ATÉ &amp; T",
                                                                                    "FUSION TOMO",
                                                                                    "FINALIZADO",
                                                                                    "13",
                                                                                    "SE",
                                                                                    "='Dados de Vendas'!B16",
                                                                                    "='Dados de Vendas'!C16",
                                                                                    "='Dados de Vendas'!D16",
                                                                                    "='Dados de Vendas'!E16",
                                                                                    "d-mmm-yy"    
                                                                                }, ' GreenWords
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

        Private _myOfficeApp As Microsoft.Office.Interop.Excel.Application

#End Region

#Region "Properties"

        Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.Excel.Application
            Get
                Return _myOfficeApp
            End Get
            Set
                _myOfficeApp = Value
            End Set
        End Property

#End Region

#Region "Methods"

        Public Sub Setup()
            MyOfficeApp = New Microsoft.Office.Interop.Excel.Application()

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

        Public Sub RunValidationPreCode()
            Console.WriteLine(vbCrLf & "Executing the pre Code. . .")

            If MyOfficeApp IsNot Nothing Then
                _localParameters(0) = False
                Question8700( _localParameters )
            End If
        End Sub

        Public Sub RunValidationPostCode()
            Console.WriteLine(vbCrLf & "Executing the post Code. . .")

            If MyOfficeApp IsNot Nothing Then
                _localParameters(0) = True
                Question8700( _localParameters )
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
            Dim Pr As Process
            For Each Pr In Procs
                If (Pr.ProcessName = "WINWORD") Then
                    For cdb = 1 To 3
                        Try
                            AppActivate(Pr.Id)
                            'TODO: SendKeys sends key strokes to the current active window.... Also, this is using Application of a windows forms app. We are in a Wpf. Find a more robust way
                            SendKeys.SendWait("{ESC}")
                            Thread.Sleep(50)
                            Forms.Application.DoEvents()
                        Catch Ex As Exception
                            Exit For
                        End Try
                    Next
                End If
            Next
        End Sub

        Public Overridable Sub MaximizeApp()
            Try
                MyOfficeApp.Visible = True
                MyOfficeApp.WindowState = XlWindowState.xlMaximized
            Catch Ex As Exception
            End Try
        End Sub

        Public Overridable Sub MinimizeApp()
            Try
                MyOfficeApp.WindowState = XlWindowState.xlMinimized
            Catch Ex As Exception
            End Try
        End Sub

        Overridable Sub CloseApp()
            Try
                If MyOfficeApp IsNot Nothing Then
                    Dim procs() = Process.GetProcesses
                    Dim pr As Process
                    For Each pr In procs
                        If (pr.ProcessName = "EXCEL") Then
                            Dim a, b As Int16
                            b = MyOfficeApp.Workbooks.Count
                            For a = 1 To b
                                MyOfficeApp.ActiveWorkbook.Close(SaveChanges:=False)
                                b = MyOfficeApp.Workbooks.Count
                                If b < 1 Then
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
            Catch ex As Exception
            End Try
        End Sub

        Overridable Sub FinalCloseApp()
            Try
                Dim procs() = Process.GetProcesses
                Dim pr As Process
                For Each pr In procs
                    If (pr.ProcessName = "EXCEL") Then
                        If MyOfficeApp IsNot Nothing Then
                            Dim a, b As Int16
                            b = MyOfficeApp.Workbooks.Count
                            For a = 1 To b
                                MyOfficeApp.ActiveWorkbook.Close(SaveChanges:=False)
                                b = MyOfficeApp.Workbooks.Count
                                If b < 1 Then
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


            If Not True Then ' Start Pre-Code
                Try
                    ''''''''''''''
                    Return True
                Catch Ex As Exception
                    Return False
                End Try
            Else ' Start Validation-Code    

                Dim Points(0)
                Points(0) = False
                Try
                    '''''
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



        Public Overridable Function Question8700(ByVal ParamArray Parameters As Object()) As [Object]
            '8700
            Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
            Dim pathTemplate As String = String.Empty
            Dim savedFile As String = CType(Parameters(2), List(Of String))(0)
            Dim programingWords As List(Of String) = New List(Of String)

            'Red Words
            Try
                programingWords = CType(Parameters(5), List(Of String))
                If programingWords.Count = 0 Then
                    programingWords.Add("Sales Data") '0
                    programingWords.Add("Quarter Totals") '1
                    programingWords.Add("Employee Record") '2
                    programingWords.Add("Adrian Parmalee") '3
                    programingWords.Add("Mandrake Wilson") '4
                    programingWords.Add("Víctor French") '5
                    programingWords.Add("SUM") '6
                    programingWords.Add("QuarterTotals") '7
                    programingWords.Add("='Sales Data'!$B$16:$E$16") '8
                    programingWords.Add("LARGEST SALE") '9
                    programingWords.Add("MAX") '10
                    programingWords.Add("SALES BY REP") '11
                    programingWords.Add("Decrease") '12
                    programingWords.Add("Increase") '13
                    programingWords.Add("Employee Record") '14
                    programingWords.Add("COPYRIGHT FUSION TOMO, ALL RIGHTS RESERVED") '15
                    programingWords.Add("CONTENT ACCURATE AS OF &T") '16
                    programingWords.Add("FUSION TOMO") '17
                    programingWords.Add("FINISHED") '18
                    programingWords.Add("13") '19 This is the font size when you change a style
                    programingWords.Add("IF") '20
                    programingWords.Add("='Sales Data'!B16") '21
                    programingWords.Add("='Sales Data'!C16") '22
                    programingWords.Add("='Sales Data'!D16") '23
                    programingWords.Add("='Sales Data'!E16") '24
                    programingWords.Add("d-mmm-yyyy") '25
                    programingWords.Add("Table1") '26
                End If
            Catch ex As Exception

            End Try

            If Not isValidate Then ' Start Pre-Code
                Try
                    'savedFile = Parameters(2).ToString()

                    With Me.MyOfficeApp
                        If String.IsNullOrEmpty(savedFile) Then
                            .Workbooks.Add()
                        Else
                            .Workbooks.Open(savedFile)
                        End If

                        .Visible = True
                    End With
                    Return True
                Catch ex As Exception
                    Return False
                End Try

            Else ' Start Validation-Code    

                Dim points(17)
                Dim i As Integer = 0

                For i = 0 To 17
                    points(i) = False
                Next i


                Try

                    With Me.MyOfficeApp


                        '3 In the Sales Data sheet, Merge and center cells A1 through F1
                        '3 Repeat the previous task for cells , A2 through F2, A3 through F3. and A4 through F4
                        Try
                            If .Worksheets(1).Range("A1").MergeCells = True Then
                                If .Worksheets(1).Range("A2").MergeCells = True Then
                                    If .Worksheets(1).Range("A3").MergeCells = True Then
                                        If .Worksheets(1).Range("A4").MergeCells = True Then
                                            points(0) = True

                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            points(0) = False
                        End Try

                        '4 Apply the Title style to cell A1	
                        '4 Apply the Explanatory... style to cells A2:A4	
                        '4 Apply the Heading 3 style to the header row, cells A5:F5	
                        '4 Apply the Total style to the total row, cells A16:F16
                        Try
                            If .Worksheets(1).Range("A1:F1").Font.Size = 18 Then
                                If .Worksheets(1).Range("A2:F4").Font.Italic = True Then
                                    If .Worksheets(1).Range("A5:F5").Font.Size = programingWords(19) And .Worksheets(1).Range("A5:F5").Font.Bold = True Then
                                        If .Worksheets(1).Range("A16:F16").Font.Bold = True And .Worksheets(1).Range("A16:F16").Font.Size = 11 Then
                                            points(1) = True
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As Exception

                            points(1) = False
                        End Try



                        '5 Change the column width of column A to 20
                        '5 Change the column widths of column B through F to 10
                        Try

                            If .Worksheets(1).Range("A1").ColumnWidth = 20 Then

                                If .Worksheets(1).Range("B1").ColumnWidth = 10 Then

                                    If .Worksheets(1).Range("F1").ColumnWidth = 10 Then
                                        points(2) = True
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            points(2) = False

                        End Try

                        '6 Sort the data by the Sales Rep column, from A to Z (data range A6:F15)
                        Try

                            If .Worksheets(1).Range("A6").Value = programingWords(3) Then

                                If .Worksheets(1).Range("A12").Value = programingWords(4) Then

                                    If .Worksheets(1).Range("A15").Value Like programingWords(5) Then
                                        points(3) = True
                                    End If
                                End If
                            End If
                        Catch ex As Exception

                            points(3) = False
                        End Try
                        '7 Use conditional formating to apply 3 Arrows (colored) to the Total column, F6:F15. Show green when the number value is greater than or equal to 30000. Show yellow when the number value is greater than or equal to 20000.
                        Try
                            If .Worksheets(1).range("F6:F15").FormatConditions.Count > 0 Then
                                If .Worksheets(1).range("F6:F15").FormatConditions(1).IconCriteria(2).Type = XlConditionValueTypes.xlConditionValueNumber Then
                                    If .Worksheets(1).range("F6:F15").FormatConditions(1).IconCriteria(2).Value = 20000 Then
                                        If .Worksheets(1).range("F6:F15").FormatConditions(1).IconCriteria(2).Operator = 7 Then
                                            If .Worksheets(1).range("F6:F15").FormatConditions(1).IconCriteria(3).Type = XlConditionValueTypes.xlConditionValueNumber Then
                                                If .Worksheets(1).range("F6:F15").FormatConditions(1).IconCriteria(3).Value = 30000 Then
                                                    If .Worksheets(1).range("F6:F15").FormatConditions(1).IconCriteria(3).Operator = 7 Then
                                                        points(4) = True

                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            points(4) = False
                        End Try
                        '8 In the total row, cell B16, insert a fomula that calculates the total for the Q1 column. (Use the SUM formula, cell range B6:B15)
                        '8 Repeat the previous task for cells C16 through F16.
                        Try
                            If InStr(.Worksheets(1).range("B16").FormulaLocal.ToUpper(), programingWords(6)) Then
                                If InStr(.Worksheets(1).range("B16").FormulaLocal, "B6:B15") Then
                                    If InStr(.Worksheets(1).range("C16").FormulaLocal.ToUpper(), programingWords(6)) Then
                                        If InStr(.Worksheets(1).range("C16").FormulaLocal, "C6:C15") Then
                                            If InStr(.Worksheets(1).range("F16").FormulaLocal.ToUpper(), programingWords(6)) Then
                                                If InStr(.Worksheets(1).range("F16").FormulaLocal, "F6:F15") Then
                                                    points(5) = True

                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            points(5) = False
                        End Try

                        '9 Create a named range for cell range B16:E16 named QuarterTotals
                        Try
                            Dim nms = .ActiveWorkbook.Names
                            Dim nm

                            If nms.Count > 1 Then

                                For Each nm In nms

                                    If nm.Name = programingWords(7) Then

                                        If InStr(nm.RefersTo, programingWords(8)) Then
                                            points(6) = True
                                            Exit For
                                        End If
                                    End If
                                Next nm
                            End If
                        Catch ex As Exception

                            points(6) = False
                        End Try

                        '10 In Cell A18, enter the text Largest Sale:
                        '10 In Cell B18, insert a formula that finds the Largest number in the Total column, not including the total row. (Use the MAX formula, cell range F6:F15)
                        Try
                            If InStr(.Worksheets(1).range("A18").Value.ToUpper(), programingWords(9)) Then
                                If InStr(.Worksheets(1).range("B18").Formula.ToUpper(), programingWords(10)) Then
                                    If .Worksheets(1).range("B18").value = 33888 Or .Worksheets(1).range("B18").value = 33800 Then
                                        points(7) = True
                                    End If
                                End If
                            End If
                        Catch
                            points(7) = False
                        End Try

                        '18 In the Sales Data sheet, insert a Clustered Column chart
                        '18 The chart data range should be the total column, F6:F15
                        '18 The Horizontal Axis Labels should be cell range A6:A15
                        '19 Change the Chart Title to read Sales by Rep	
                        '19 Add the title of Sales by Rep to the alt text of the chart
                        '20 Apply Chart Style 7 to the chart
                        '21 Show data labels on the outside end of the data.	
                        '21 Position the chart below the other data in the sheet. (Note: exact positioning is not required)	
                        Try
                            If .Worksheets(1).Shapes.Count > 0 Then
                                For i = 1 To .Worksheets(1).Shapes.Count ' loop so that the task validates no matter which order they inserted the chart/picture in.
                                    If .Worksheets(1).Shapes(i).Type = 3 Then
                                        .Sheets(1).Select()
                                        .Worksheets(1).Shapes(i).Select()
                                        .ActiveChart.ChartArea.Select()
                                        If .ActiveChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered Then
                                            points(8) = True
                                        End If

                                        If .ActiveChart.ChartTitle.Text.ToUpper() = programingWords(11) Then

                                            If .ActiveChart.ChartTitle.Caption.ToUpper() = programingWords(11) Then
                                                points(9) = True
                                            End If
                                        End If
                                        If .ActiveChart.ChartStyle = 207 Then
                                            points(10) = True

                                        End If
                                        If .ActiveChart.SeriesCollection(1).DataLabels.Count = 10 Then
                                            points(11) = True

                                        End If
                                    End If
                                Next i
                            End If
                        Catch ex As Exception

                        End Try



                        '22 Insert ftlogo.gif into the Sales Data sheet.	
                        '23 Crop the logo so that only the symbol at the left remains. the width should be 1.1	
                        '24 Position the image so that it is in the upper left corner. (Note: The position does not need to be exact. see the example doc for a suggestion)	
                        '25 Add a picture effect of Full Reflection, 4 pt offset.
                        Try
                            For i = 1 To .Worksheets(1).Shapes.Count ' loop so that the task validates no matter which order they inserted the chart/picture in.
                                If .Worksheets(1).Shapes(i).Type = 11 Or .Worksheets(1).Shapes(i).Type = 13 Then
                                    points(12) = True

                                    .Sheets(1).Select()
                                    If .Worksheets(1).Shapes(i).Width < 85 Then
                                        points(13) = True
                                        points(14) = True
                                    End If
                                    If .Worksheets(1).Shapes(i).Reflection.Type = MsoReflectionType.msoReflectionType6 Then
                                        points(15) = True
                                    End If
                                End If
                            Next i
                        Catch
                        End Try


                        '26 Change the Page Layout of the doument to Landscape, and the page size to Legal	
                        Try
                            If .Worksheets(1).PageSetup.Orientation = XlPageOrientation.xlLandscape Then
                                If .Worksheets(1).PageSetup.PaperSize = XlPaperSize.xlPaperLegal Then
                                    points(16) = True
                                End If
                            End If
                        Catch
                        End Try

                        '28 on the Sales Data Sheet, insert a page break on row 19, so that the data and the chart will print on seperate pages.	
                        Try

                            If .Worksheets(1).HPageBreaks.Count = 1 Then
                                points(17) = True
                            End If
                        Catch
                        End Try

                    End With
                Catch ex As Exception
                    Return points
                End Try
                Return points
            End If
            Return Nothing
        End Function






    End Class
End Namespace
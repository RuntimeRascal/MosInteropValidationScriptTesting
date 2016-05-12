

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
Imports Microsoft.Office.Interop.Access
Imports Microsoft.Office.Interop.Access.Dao

Namespace GMetrix.Dynamic.Validations

    Public Class DynamicQuestion

        Private _myOfficeApp As Microsoft.Office.Interop.Access.Application

        Private _userDbSavePath As [String]

        Public Sub New()
            MyBase.New
            _myOfficeApp = New Microsoft.Office.Interop.Access.Application
            MaximizeApp()
        End Sub

        Public Overridable Property MyOfficeApp() As Microsoft.Office.Interop.Access.Application
            Get
                Return Me._myOfficeApp
            End Get
            Set
                Me._myOfficeApp = Value
            End Set
        End Property

        Public Overridable Function Validate8817(ByVal ParamArray Parameters As Object()) As [Object]
            '8817

            Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)

            Dim templates As List(Of String) = CType(Parameters(1), List(Of String))

            Dim pathTemplate As String = String.Empty

            'Redwords for Access 1
            Dim programingWords As List(Of String) = New List(Of String)

            'Red Words
            Try
                programingWords = CType(Parameters(5), List(Of String))
                If programingWords.Count = 0 Then
                    programingWords.Add("Customers") '0
                    programingWords.Add("CustomerId") '1
                    programingWords.Add("PrimaryKey") '2
                    programingWords.Add("=Date()") '3
                    programingWords.Add("Orders") '4
                    programingWords.Add("OrderId") '5
                    programingWords.Add("ProductId") '6
                    programingWords.Add("US Customers") '7
                    programingWords.Add("*FROM CUSTOMERS*") '8
                    programingWords.Add("*WHERE*COUNTRY*=*USA*") '9
                    programingWords.Add("*ORDER BY*JOINDATE*DESC*") '10
                    programingWords.Add("Orders Processing") '11
                    programingWords.Add("*SELECT*CUSTOMERS*CUSTOMERID*ORDERS*ORDERSTATUS*FROM*CUSTOMERS*INNER JOIN*ORDERS*") '12
                    programingWords.Add("*WHERE*ORDERSTATUS*PROCESSING*") '13
                    programingWords.Add("*ORDER BY*ORDERDATE*") '14
                    programingWords.Add("*DATE*-*ORDERDATE*AS*DAYSSINCEORDER*") '15
                    programingWords.Add("New Customers Form") '16
                    programingWords.Add("Enter a New Customer") '17
                    programingWords.Add("Zip") '18
                    programingWords.Add("US Customers Report By State") '19
                    programingWords.Add("State") '20
                    programingWords.Add("CustomerName_Label") '21
                    programingWords.Add("Address") '22
                    programingWords.Add("ReportHeader") '23
                    programingWords.Add("PageHeaderSection") '24
                    programingWords.Add("Auto Compact") '25					
                End If
            Catch Ex As Exception
            End Try

            If Not isValidate Then ' Start Pre-Code
                Try
                    pathTemplate = Parameters(2).ToString()
                    With Me.MyOfficeApp
                        .Visible = True
                        If String.IsNullOrWhiteSpace(pathTemplate) Then
                            .OpenCurrentDatabase(templates(0))
                        Else
                            .OpenCurrentDatabase(pathTemplate)
                        End If
                    End With
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            Else ' Start Validation-Code
                Dim points(3)
                points(0) = False
                points(1) = False
                points(2) = False
                points(3) = False
                Dim vR As Integer = 0
                '4527#############################################################################################################
                '[177829]Create(Field)
                'Table:Customers()
                'Field(Name) : CustomerId()
                'Data Type: AutoNumber
                'Field Order: Top field in table
                Try
                    Dim aT As Integer = 0
                    For aT = 0 To Me.MyOfficeApp.CurrentData.AllTables.Count - 1
                        If Me.MyOfficeApp.CurrentData.AllTables(aT).Name = programingWords(0) Then
                            If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields.Count = 8 Then
                                If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(0).Name = programingWords(1) Then
                                    If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(0).Type = 4 Then
                                        points(0) = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try

                Try
                    Dim aT1 As Integer = 0
                    For aT1 = 0 To Me.MyOfficeApp.CurrentData.AllTables.Count - 1
                        If Me.MyOfficeApp.CurrentData.AllTables(aT1).Name = programingWords(0) Then
                            If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes.Count = 1 Then
                                If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes(0).Name = programingWords(2) Then
                                    If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes(0).Primary Then
                                        points(1) = True
                                        Exit For
                                    End If
                                End If
                            ElseIf Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes.Count = 2 Then
                                If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes(1).Name = programingWords(2) Then
                                    If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes(1).Primary Then
                                        points(1) = True
                                        Exit For
                                    End If
                                End If
                                'Below code is for Arabic validation it uses index 0 instead of 1
                                If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes(0).Name = programingWords(2) Then
                                    If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Indexes(0).Primary Then
                                        points(1) = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If

                    Next
                Catch ex As Exception
                End Try

                Dim rst, myarray
                Try
                    rst = Me.MyOfficeApp.CurrentDb.OpenRecordset(programingWords(0))
                    myarray = rst.GetRows(6)
                    rst.Close()
                    Dim cJD As Integer = 0
                    If InStr(myarray(7, 0), "T00:00:00") Then
                        cJD = cJD + 1
                    End If
                    If InStr(myarray(7, 2), "T00:00:00") Then
                        cJD = cJD + 1
                    End If
                    If InStr(myarray(7, 4), "T00:00:00") Then
                        cJD = cJD + 1
                    End If

                    If cJD = 0 Then
                        points(2) = True
                    End If

                Catch ex As Exception

                End Try

                Try
                    Dim aT2 As Integer = 0
                    For aT2 = 0 To Me.MyOfficeApp.CurrentData.AllTables.Count - 1
                        If Me.MyOfficeApp.CurrentData.AllTables(aT2).Name = programingWords(0) Then
                            'No borrar los Or que agregue, son los que valida en españon, Miguel
                            If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(7).Type = 8 Or Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(7).Type = 10 Then
                                If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(7).DefaultValue = programingWords(3) Or Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(6).DefaultValue = programingWords(3) Then
                                    If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(7).Required Or Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(0)).Fields(6).Required Then
                                        points(3) = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try
                Return points
            End If
        End Function

        Public Overridable Function Validate8818(ByVal ParamArray Parameters As Object()) As [Object]
            '8818			
            Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
            Dim pathTemplate As String = String.Empty

            'Redwords for Access 1
            Dim programingWords As List(Of String) = New List(Of String)

            'Red Words
            Try
                programingWords = CType(Parameters(5), List(Of String))
                If programingWords.Count = 0 Then
                    programingWords.Add("Customers") '0
                    programingWords.Add("CustomerId") '1
                    programingWords.Add("PrimaryKey") '2
                    programingWords.Add("=Date()") '3
                    programingWords.Add("Orders") '4
                    programingWords.Add("OrderId") '5
                    programingWords.Add("ProductId") '6
                    programingWords.Add("US Customers") '7
                    programingWords.Add("*FROM CUSTOMERS*") '8
                    programingWords.Add("*WHERE*COUNTRY*=*USA*") '9
                    programingWords.Add("*ORDER BY*JOINDATE*DESC*") '10
                    programingWords.Add("Orders Processing") '11
                    programingWords.Add("*SELECT*CUSTOMERS*CUSTOMERID*ORDERS*ORDERSTATUS*FROM*CUSTOMERS*INNER JOIN*ORDERS*") '12
                    programingWords.Add("*WHERE*ORDERSTATUS*PROCESSING*") '13
                    programingWords.Add("*ORDER BY*ORDERDATE*") '14
                    programingWords.Add("*DATE*-*ORDERDATE*AS*DAYSSINCEORDER*") '15
                    programingWords.Add("New Customers Form") '16
                    programingWords.Add("Enter a New Customer") '17
                    programingWords.Add("Zip") '18
                    programingWords.Add("US Customers Report By State") '19
                    programingWords.Add("State") '20
                    programingWords.Add("CustomerName_Label") '21
                    programingWords.Add("Address") '22
                    programingWords.Add("ReportHeader") '23
                    programingWords.Add("PageHeaderSection") '24
                    programingWords.Add("Auto Compact") '25
                End If
            Catch Ex As Exception
            End Try

            If Not isValidate Then ' Start Pre-Code
                Try
                    pathTemplate = Parameters(2).ToString()
                    With Me.MyOfficeApp
                        .Visible = True
                        If Not String.IsNullOrWhiteSpace(pathTemplate) Then
                            .OpenCurrentDatabase(pathTemplate)
                        End If
                    End With
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            Else ' Start Validation-Code
                Dim points(0)
                points(0) = False

                Try
                    Dim aT4 As Integer = 0
                    For aT4 = 0 To Me.MyOfficeApp.CurrentData.AllTables.Count - 1
                        If Me.MyOfficeApp.CurrentData.AllTables(aT4).Name = programingWords(4) Then
                            If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(4)).Fields(1).Type = 4 Then
                                If Me.MyOfficeApp.CurrentDb.TableDefs(programingWords(4)).Fields(2).Type = 8 Then
                                    points(0) = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try
                Return points
            End If
        End Function

        Public Overridable Function Validate8820(ByVal ParamArray Parameters As Object()) As [Object]
            '8820
			Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
            Dim pathTemplate As String = String.Empty
			
			'Redwords for Access 1
			Dim programingWords As List(Of String) = New List(Of String)

			'Red Words
			Try
				programingWords = CType(Parameters(5), List(Of String))
				If programingWords.Count = 0 Then                    
					programingWords.Add("Customers") '0
					programingWords.Add("CustomerId") '1
					programingWords.Add("PrimaryKey") '2
					programingWords.Add("=Date()") '3
					programingWords.Add("Orders") '4
					programingWords.Add("OrderId") '5
					programingWords.Add("ProductId") '6
					programingWords.Add("US Customers") '7
					programingWords.Add("*FROM CUSTOMERS*") '8
					programingWords.Add("*WHERE*COUNTRY*=*USA*") '9
					programingWords.Add("*ORDER BY*JOINDATE*DESC*") '10
					programingWords.Add("Orders Processing") '11
					programingWords.Add("*SELECT*CUSTOMERS*CUSTOMERID*ORDERS*ORDERSTATUS*FROM*CUSTOMERS*INNER JOIN*ORDERS*") '12
					programingWords.Add("*WHERE*ORDERSTATUS*PROCESSING*") '13
					programingWords.Add("*ORDER BY*ORDERDATE*") '14
					programingWords.Add("*DATE*-*ORDERDATE*AS*DAYSSINCEORDER*") '15
					programingWords.Add("New Customers Form") '16
					programingWords.Add("Enter a New Customer") '17
					programingWords.Add("Zip") '18
					programingWords.Add("US Customers Report By State") '19
					programingWords.Add("State") '20
					programingWords.Add("CustomerName_Label") '21
					programingWords.Add("Address") '22
					programingWords.Add("ReportHeader") '23
					programingWords.Add("PageHeaderSection") '24
					programingWords.Add("Auto Compact") '25
				End If
			Catch Ex As Exception
			End Try
			
            If Not isValidate Then ' Start Pre-Code
                Try
                    pathTemplate = Parameters(2).ToString()
                    With Me.MyOfficeApp
                        .Visible = True
                        If Not String.IsNullOrWhiteSpace(pathTemplate) Then
                            .OpenCurrentDatabase(pathTemplate)
                        End If
                    End With
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            Else ' Start Validation-Code
                Dim points(2)
                points(0) = False
                points(1) = False
				points(2) = False

				Try				
                    If CInt(Me.MyOfficeApp.CurrentDb.Relations.Count) > 1 Then
                        If Me.MyOfficeApp.CurrentDb.Relations(0).Attributes = 0 Then
                            If Me.MyOfficeApp.CurrentDb.Relations(0).Fields(0).ForeignName.ToString = programingWords(1) Then
                                If Me.MyOfficeApp.CurrentDb.Relations(0).Fields(0).Name.ToString = programingWords(1) Then
                                    points(0) = True
                                End If
                            End If
                        End If
                    End If
					'Below code is for arabic right to left
					If CInt(Me.MyOfficeApp.CurrentDb.Relations.Count) > 1 Then
                        If Me.MyOfficeApp.CurrentDb.Relations(3).Attributes = 0 Then
                            If Me.MyOfficeApp.CurrentDb.Relations(3).Fields(0).ForeignName.ToString = programingWords(1) Then
                                If Me.MyOfficeApp.CurrentDb.Relations(3).Fields(0).Name.ToString = programingWords(1) Then
                                    points(0) = True
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                End Try

				Try
                    If CInt(Me.MyOfficeApp.CurrentDb.Relations.Count) > 3 Then
                        If Me.MyOfficeApp.CurrentDb.Relations(3).Attributes = 2 Then
                            If Me.MyOfficeApp.CurrentDb.Relations(3).Fields(0).ForeignName.ToString = programingWords(5) Then
                                If Me.MyOfficeApp.CurrentDb.Relations(3).Fields(0).Name.ToString.ToUpper() = programingWords(5).ToUpper() Then
                                    points(1) = True
                                End If
                            End If
                        End If
                    End If
					'Below code is for arabic right to left adjustment
					If CInt(Me.MyOfficeApp.CurrentDb.Relations.Count) > 3 Then
                        If Me.MyOfficeApp.CurrentDb.Relations(2).Attributes = 2 Then
                            If Me.MyOfficeApp.CurrentDb.Relations(2).Fields(0).ForeignName.ToString = programingWords(5) Then
                                If Me.MyOfficeApp.CurrentDb.Relations(2).Fields(0).Name.ToString.ToUpper() = programingWords(5).ToUpper() Then
                                    points(1) = True
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                End Try

				Try
                    If CInt(Me.MyOfficeApp.CurrentDb.Relations.Count) > 4 Then
                        If Me.MyOfficeApp.CurrentDb.Relations(4).Attributes = 2 Then
                            If Me.MyOfficeApp.CurrentDb.Relations(4).Fields(0).ForeignName.ToString = programingWords(6) Then
                                If Me.MyOfficeApp.CurrentDb.Relations(4).Fields(0).Name.ToString = programingWords(6) Then
                                    points(2) = True
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception

                End Try
                     Return points
            End If
        End Function

        Public Overridable Function Validate8824(ByVal ParamArray Parameters As Object()) As [Object]
            '8824
            Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
            Dim pathTemplate As String = String.Empty

            'Redwords for Access 1
            Dim programingWords As List(Of String) = New List(Of String)

            'Red Words
            Try
                programingWords = CType(Parameters(5), List(Of String))
                If programingWords.Count = 0 Then
                    programingWords.Add("Customers") '0
                    programingWords.Add("CustomerId") '1
                    programingWords.Add("PrimaryKey") '2
                    programingWords.Add("=Date()") '3
                    programingWords.Add("Orders") '4
                    programingWords.Add("OrderId") '5
                    programingWords.Add("ProductId") '6
                    programingWords.Add("US Customers") '7
                    programingWords.Add("*FROM CUSTOMERS*") '8
                    programingWords.Add("*WHERE*COUNTRY*=*USA*") '9
                    programingWords.Add("*ORDER BY*JOINDATE*DESC*") '10
                    programingWords.Add("Orders Processing") '11
                    programingWords.Add("*SELECT*CUSTOMERS*CUSTOMERID*ORDERS*ORDERSTATUS*FROM*CUSTOMERS*INNER JOIN*ORDERS*") '12
                    programingWords.Add("*WHERE*ORDERSTATUS*PROCESSING*") '13
                    programingWords.Add("*ORDER BY*ORDERDATE*") '14
                    programingWords.Add("*DATE*-*ORDERDATE*AS*DAYSSINCEORDER*") '15
                    programingWords.Add("New Customers Form") '16
                    programingWords.Add("Enter a New Customer") '17
                    programingWords.Add("Zip") '18
                    programingWords.Add("US Customers Report By State") '19
                    programingWords.Add("State") '20
                    programingWords.Add("CustomerName_Label") '21
                    programingWords.Add("Address") '22
                    programingWords.Add("ReportHeader") '23
                    programingWords.Add("PageHeaderSection") '24
                    programingWords.Add("Auto Compact") '25
                End If
            Catch Ex As Exception
            End Try

            If Not isValidate Then ' Start Pre-Code
                Try
                    pathTemplate = Parameters(2).ToString()
                    With Me.MyOfficeApp
                        .Visible = True
                        If Not String.IsNullOrWhiteSpace(pathTemplate) Then
                            .OpenCurrentDatabase(pathTemplate)
                        End If
                    End With
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            Else ' Start Validation-Code
                Dim points(4)
                points(0) = False
                points(1) = False
                points(2) = False
                points(3) = False
                points(4) = False

                Dim aR As Int16 = 0

                'Adding a block of code that will close the report and then open the report again
                'It seems to be something that is needed in order to grade properly but we are not sure why
                Try
                    Me.MyOfficeApp.DoCmd.Close(3, programingWords(19)) ' 3 is a report object type https://msdn.microsoft.com/en-us/library/office/ff845495.aspx
                    Me.MyOfficeApp.DoCmd.OpenReport(programingWords(19), AcView.acViewDesign)
                Catch
                End Try

                For aR = 0 To Me.MyOfficeApp.Reports.Count - 1
                    Try
                        Try
                            If Me.MyOfficeApp.Reports(aR).Name.ToString = programingWords(19) Then
                                If InStr(Me.MyOfficeApp.Reports(aR).Controls(1).Caption.ToString, programingWords(20)) Then
                                    points(0) = True
                                End If
                            End If
                        Catch
                        End Try
                        '[177851]Edit Report Fields
                        'Remove Field: Page Header CustomerId, Detail: CustomerId
                        'Change Field Width: Page Header JoinDate, Detail JoinDate: 0.8"
                        'Change Field Text: Page Header, State/Province label, change to State
                        Try
                            If Me.MyOfficeApp.Reports(aR).Controls(3).Name.ToString = programingWords(21) Then
                                If Me.MyOfficeApp.Reports(aR).Controls(10).Name.ToString = programingWords(22) Then
                                    If Me.MyOfficeApp.Reports(aR).Controls(2).Width >= 1133 And Me.MyOfficeApp.Reports(aR).Controls(2).Width <= 1152 Then
                                        If Me.MyOfficeApp.Reports(aR).Controls(1).Caption.ToString = programingWords(20) Or Me.MyOfficeApp.Reports(aR).Controls(1).Name.ToString = programingWords(20) Then
                                            points(1) = True
                                        End If
                                    End If
                                End If
                            End If
                        Catch
                        End Try

                        '[177852]Format Report
                        'Report Header Height: 1"
                        'Report Header Back Color: Gold, Accent 4
                        Try
                            If Me.MyOfficeApp.Reports(aR).Section(1).Name = programingWords(23) Then
                                If Me.MyOfficeApp.Reports(aR).Section(1).Height >= 1418 And Me.MyOfficeApp.Reports(aR).Section(1).Height <= 1440 Then
                                    If Me.MyOfficeApp.Reports(aR).Section(1).BackColor = 49407 Then
                                        points(2) = True
                                    End If
                                End If
                            End If
                        Catch
                        End Try
                        '[177853]
                        'Page Header All Controls Font Size: 14
                        'Page Header All Controls Text Align: Center
                        Try
                            If Me.MyOfficeApp.Reports(aR).Section(3).Name = programingWords(24) Then
                                If Me.MyOfficeApp.Reports(aR).Section(3).Controls(0).FontSize = 14 Then
                                    If Me.MyOfficeApp.Reports(aR).Section(3).Controls(0).TextAlign = 2 Then
                                        points(3) = True
                                    End If
                                End If
                            End If
                        Catch
                        End Try
                        '[177854]Report Print Options
                        'Orientation: Landscape
                        'Paper Size: A3 
                        'Print Data Only
                        Try
                            If Me.MyOfficeApp.Reports(aR).Printer.Orientation = 2 Then
                                If Me.MyOfficeApp.Reports(aR).Printer.PaperSize = 5 Then
                                    points(4) = True
                                End If
                            End If
                        Catch
                        End Try

                    Catch
                    End Try
                Next
                Return points
            End If
        End Function

        Public Overridable Function Validate8825(ByVal ParamArray Parameters As Object()) As [Object]
            '8825
            Dim isValidate As Boolean = CType(Parameters(0).ToString(), Boolean)
            Dim templates As List(Of String) = CType(Parameters(1), List(Of String))
            Dim pathTemplate As String = String.Empty

            'Redwords for Access 1
            Dim programingWords As List(Of String) = New List(Of String)

            'Red Words
            Try
                programingWords = CType(Parameters(5), List(Of String))
                If programingWords.Count = 0 Then
                    programingWords.Add("Customers") '0
                    programingWords.Add("CustomerId") '1
                    programingWords.Add("PrimaryKey") '2
                    programingWords.Add("=Date()") '3
                    programingWords.Add("Orders") '4
                    programingWords.Add("OrderId") '5
                    programingWords.Add("ProductId") '6
                    programingWords.Add("US Customers") '7
                    programingWords.Add("*FROM CUSTOMERS*") '8
                    programingWords.Add("*WHERE*COUNTRY*=*USA*") '9
                    programingWords.Add("*ORDER BY*JOINDATE*DESC*") '10
                    programingWords.Add("Orders Processing") '11
                    programingWords.Add("*SELECT*CUSTOMERS*CUSTOMERID*ORDERS*ORDERSTATUS*FROM*CUSTOMERS*INNER JOIN*ORDERS*") '12
                    programingWords.Add("*WHERE*ORDERSTATUS*PROCESSING*") '13
                    programingWords.Add("*ORDER BY*ORDERDATE*") '14
                    programingWords.Add("*DATE*-*ORDERDATE*AS*DAYSSINCEORDER*") '15
                    programingWords.Add("New Customers Form") '16
                    programingWords.Add("Enter a New Customer") '17
                    programingWords.Add("Zip") '18
                    programingWords.Add("US Customers Report By State") '19
                    programingWords.Add("State") '20
                    programingWords.Add("CustomerName_Label") '21
                    programingWords.Add("Address") '22
                    programingWords.Add("ReportHeader") '23
                    programingWords.Add("PageHeaderSection") '24
                    programingWords.Add("Auto Compact") '25
                End If
            Catch Ex As Exception
            End Try

            If Not isValidate Then ' Start Pre-Code
                Try
                    pathTemplate = Parameters(2).ToString()
                    With Me.MyOfficeApp
                        .Visible = True
                        If Not String.IsNullOrWhiteSpace(pathTemplate) Then
                            .OpenCurrentDatabase(pathTemplate)
                        Else
                            .OpenCurrentDatabase(templates(0).ToString().Trim())
                        End If
                    End With
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            Else ' Start Validation-Code
                Dim points(0)
                points(0) = False

                '[177834]Change Database Options
                'Compact on Close
                Try
                    If Me.MyOfficeApp.Application.GetOption(programingWords(25)) = -1 Then
                        points(0) = True
                    End If
                Catch ex As Exception
                End Try
                Return points
            End If
        End Function

        Public Overridable Sub SetUserDbSavePath()
            Try
                If MyOfficeApp IsNot Nothing Then
                    _userDbSavePath = _myOfficeApp.CurrentDb().Name
                End If
            Catch ex As System.Exception
            End Try

        End Sub

        Public Overridable Function GetUserDbSavePath() As [String]
            Return _userDbSavePath
        End Function

        Public Overridable Sub CloseApp()
            Try
                If MyOfficeApp IsNot Nothing Then
                    If MyOfficeApp.CurrentObjectType <> AcObjectType.acDefault Then
                        MyOfficeApp.CloseCurrentDatabase()
                    End If
                    MyOfficeApp.Quit(Option:=Microsoft.Office.Interop.Access.AcQuitOption.acQuitSaveNone)
                    If MyOfficeApp IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
                        MyOfficeApp = Nothing
                    End If
                End If
            Catch ex As System.Exception
            End Try

        End Sub

        Public Overridable Sub CloseDialogs()
            Dim procs() = Process.GetProcesses
            Dim pr As Process
            For Each pr In procs
                If (pr.ProcessName = "MSACCESS") Then
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
                Me.MyOfficeApp.RunCommand(AcCommand.acCmdAppMaximize)
            Catch Ex As System.Exception
            End Try

        End Sub

        Private Sub GarbageCollector()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Sub

        Private Function ExistsDirectory(ByVal path As [String]) As [Boolean]
            Try
                If (Not Directory.Exists(path)) Then
                    Directory.CreateDirectory(path)
                End If
                Return Directory.Exists(path)
            Catch ex As System.Exception
                Return False
            End Try

        End Function

        Private Function ExistsFile(ByVal file As [String]) As [Boolean]
            Try
                Return IO.File.Exists(file)
            Catch ex As System.Exception
                Return False
            End Try

        End Function

        Public Overridable Sub SaveAs(ByVal ParamArray Parameters As Object())
            Dim filePath As String = Parameters(0).ToString()
            Dim sourcePath As String = Parameters(1).ToString()
            Try
                If MyOfficeApp Is Nothing Then
                    Return
                End If
                Me.MyOfficeApp.DoCmd.SetWarnings(False)
                Me.MyOfficeApp.DoCmd.Save()
                System.IO.File.Copy(sourcePath, filePath, True)
                MessageBox.Show(sourcePath)
                MessageBox.Show(filePath)
                Me.MyOfficeApp.DoCmd.SetWarnings(True)
            Catch ex As System.Exception
                Me.MyOfficeApp.DoCmd.SetWarnings(True)
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Overridable Sub MinimizeApp()
            Try
                Me.MyOfficeApp.RunCommand(AcCommand.acCmdAppMinimize)
            Catch ex As System.Exception
            End Try

        End Sub

        Public Overridable Sub FinalCloseApp()
            Try
                If MyOfficeApp IsNot Nothing Then
                    If MyOfficeApp.CurrentObjectType <> AcObjectType.acDefault Then
                        MyOfficeApp.CloseCurrentDatabase()
                    End If
                End If
            Catch Ex As System.Exception
            End Try
            If MyOfficeApp IsNot Nothing Then
                MyOfficeApp.Quit()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(MyOfficeApp)
                MyOfficeApp = Nothing
            End If

        End Sub
    End Class
End Namespace

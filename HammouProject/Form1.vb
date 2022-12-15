Imports DocumentProcessor90
Imports LFSO90Lib
Imports System.Data.SqlClient
Imports System.IO

Public Class Form1

    Dim tic As Boolean = False
    Public server As String
    Public rep As String
    Public user As String
    Public pass As String
    Public volum As String
    Public serverSqlName As String
    Public databaseName As String
    Public tablePatients As String
    Public tablevisit As String
    Public logFilePath As String
    Public fileNIPpath As String
    Public pathInLF As String

    Public templateName As String
    Public MedicalFileField As String
    Public FamilyField As String
    Public FirstField As String
    Public FatherField As String
    Public MotherField As String
    Public SexField As String
    Public BloodGroupField As String
    Public BirthDateField As String

    Public templateVisiName As String
    Public CaseNumberField As String
    Public AdmissionDateField As String
    Public DischargeDateField As String
    Public DepartmentField As String
    Public DoctorCodeField As String
    Public DoctorNameField As String
    Public ConsultantCodeField As String
    Public ConsultantNameField As String
    Public PatientTypeField As String
    Public ER_InpatientField As String
    Public CoverageField As String

    Public templatePlatTechnique As String
    Public Type1Field As String
    Public type2Field As String

    Public folNumberName As String = ""

    Public db As LFDatabase
    Public serv As LFServer
    Public conn As LFConnection

    Public MedicalFile As String
    Public MedicalFileNew As String
    Public Family As String
    Public First As String
    Public Father As String
    Public Mother As String
    Public Sex As String
    Public BloodGroup As String
    Public BirthDate As String
    Public flagBMB As String
    Public patientAutoID As String

    Public MedicalFileVisit As String
    Public CaseNumber As String
    Public CaseNumberNew As String
    Public AdmissionDate As String
    Public DischargeDate As String
    Public Department As String
    Public DoctorCode As String
    Public DoctorName As String
    Public ConsultantCode As String
    Public ConsultantName As String
    Public PatientType As String
    Public ER_Inpatient As String
    Public Coverage As String
    Public visitflagBMB As String
    Public visitAutoID As String

    Public doctorNameBeforeUpdate As String
    Public ConsulanttNameBeforeUpdate As String

    'Public arrDoctorsInMergeNIP() As String
    'Public arrConsultantInMergeNIP() As String

    Public pathPDF As String
    Public ProgressPDF As String

    Public Sub ParmFunction(ByVal path As String)
        Dim objReader As New System.IO.StreamReader(path)
        server = objReader.ReadLine.Split("=")(1)
        rep = objReader.ReadLine.Split("=")(1)
        user = objReader.ReadLine.Split("=")(1)
        pass = objReader.ReadLine.Split("=")(1)
        volum = objReader.ReadLine.Split("=")(1)
        serverSqlName = objReader.ReadLine.Split("=")(1)
        databaseName = objReader.ReadLine.Split("=")(1)
        tablePatients = objReader.ReadLine.Split("=")(1)
        tablevisit = objReader.ReadLine.Split("=")(1)
        logFilePath = objReader.ReadLine.Split("=")(1)

        objReader.Close()
    End Sub

    Public Sub ParmFunctionNip(ByVal path As String)
        Dim objReader As New System.IO.StreamReader(path)
        pathInLF = objReader.ReadLine.Split("=")(1)
        templateName = objReader.ReadLine.Split("=")(1)
        MedicalFileField = objReader.ReadLine.Split("=")(1)
        FamilyField = objReader.ReadLine.Split("=")(1)
        FirstField = objReader.ReadLine.Split("=")(1)
        FatherField = objReader.ReadLine.Split("=")(1)
        MotherField = objReader.ReadLine.Split("=")(1)
        SexField = objReader.ReadLine.Split("=")(1)
        BloodGroupField = objReader.ReadLine.Split("=")(1)
        BirthDateField = objReader.ReadLine.Split("=")(1)
        objReader.Close()
    End Sub

    Public Sub ParmFunctionVisits(ByVal path As String)
        Dim objReader As New System.IO.StreamReader(path)
        templateVisiName = objReader.ReadLine.Split("=")(1)
        CaseNumberField = objReader.ReadLine.Split("=")(1)
        AdmissionDateField = objReader.ReadLine.Split("=")(1)
        DischargeDateField = objReader.ReadLine.Split("=")(1)
        DepartmentField = objReader.ReadLine.Split("=")(1)
        DoctorCodeField = objReader.ReadLine.Split("=")(1)
        DoctorNameField = objReader.ReadLine.Split("=")(1)
        ConsultantCodeField = objReader.ReadLine.Split("=")(1)
        ConsultantNameField = objReader.ReadLine.Split("=")(1)
        PatientTypeField = objReader.ReadLine.Split("=")(1)
        ER_InpatientField = objReader.ReadLine.Split("=")(1)
        CoverageField = objReader.ReadLine.Split("=")(1)
        objReader.Close()
    End Sub

    Public Sub ParmFunctionPlateauTechnique(ByVal path As String)
        Dim objReader As New System.IO.StreamReader(path)
        templatePlatTechnique = objReader.ReadLine.Split("=")(1)
        Type1Field = objReader.ReadLine.Split("=")(1)
        type2Field = objReader.ReadLine.Split("=")(1)
        pathPDF = objReader.ReadLine.Split("=")(1)
        ProgressPDF = objReader.ReadLine.Split("=")(1)
        objReader.Close()
    End Sub

    ''' <summary>
    ''' Open connection with laserfiche
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub connDb()
        Try
            conn = New LFConnection
            ' Creates a new application object.
            Dim app As LFApplication = New LFApplication
            ' Finds the appropriate server.
            Dim serv As LFServer = app.GetServerByName(server)
            ' Gets the repository from the server.
            db = serv.GetDatabaseByName(rep)
            ' Creates a new LFConnection object.

            ' Sets the user name and password.
            conn.UserName = user
            conn.Password = pass
            ' Connects to repository.
            conn.Create(db)
        Catch ex As Exception
            Try
                If conn.IsTerminated = False Then
                    conn.Terminate()
                End If
            Catch ex1 As Exception

            End Try
        End Try
    End Sub

    Public Function returnFolderNumberName(ByVal numberAsString As String)

        folNumberName = ""
        Dim numberAsInteger = Convert.ToInt64(numberAsString)
        If numberAsInteger >= 0 And numberAsInteger <= 10000 Then
            folNumberName = "00000000-00010000"
            'MsgBox(numberAsInteger Mod 10000 & "-" & numberAsInteger / 10000)
        ElseIf numberAsInteger Mod 10000 <> 0 Then
            If numberAsInteger >= 10000 And numberAsInteger < 100000 Then
                If numberAsInteger > 90000 Then
                    folNumberName = "000" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-00" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000
                Else
                    folNumberName = "000" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-000" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000
                End If

            ElseIf numberAsInteger >= 100000 And numberAsInteger < 1000000 Then
                If numberAsInteger > 990000 Then
                    folNumberName = "00" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-0" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000
                Else
                    folNumberName = "00" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-00" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000
                End If

            ElseIf numberAsInteger >= 1000000 And numberAsInteger < 10000000 Then
                If numberAsInteger > 9990000 Then
                    folNumberName = "0" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000
                Else
                    folNumberName = "0" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-0" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000
                End If

            ElseIf numberAsInteger >= 10000000 Then
                folNumberName = (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000
            End If
            ' MsgBox((Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 1 & "-" & (Convert.ToInt32((numberAsInteger / 10000 - ((numberAsInteger Mod 10000) / 10000))) * 10000) + 10000)
        ElseIf numberAsInteger Mod 10000 = 0 And numberAsInteger > 10000 Then
            If numberAsInteger >= 10000 And numberAsInteger < 100000 Then
                folNumberName = "000" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 1 & "-000" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 10000
            ElseIf numberAsInteger >= 100000 And numberAsInteger < 1000000 Then
                folNumberName = "00" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 1 & "-00" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 10000
            ElseIf numberAsInteger >= 1000000 And numberAsInteger < 10000000 Then
                folNumberName = "0" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 1 & "-0" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 10000
            ElseIf numberAsInteger >= 10000000 Then
                folNumberName = (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 1 & "-" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 10000
            End If
            ' MsgBox((Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 1 & "-" & (Convert.ToInt32(((numberAsInteger - 1) / 10000 - (((numberAsInteger - 1) Mod 10000) / 10000))) * 10000) + 10000)
        End If

        Return folNumberName
    End Function

    ''' <summary>
    ''' Insert exceptions to the logfile
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="fileName"></param>
    ''' <remarks></remarks>
    Public Sub writeLogFile(ByVal text As String, ByVal fileName As String)

        If System.IO.File.Exists(fileName) = True Then
            Dim objWriter As New System.IO.StreamWriter(fileName, True)
            objWriter.WriteLine(text)
            objWriter.Close()
        End If
    End Sub

    Private Function getPatientViewRows() As DataTable
        Dim arrRows As ArrayList = New ArrayList
        Dim objConn As SqlConnection
        Dim connectionString As String
        'Dim StrConn As String = "Data Source=" & serverSqlName & ";Initial Catalog=" & databaseName & ";Integrated Security=True"
        Dim StrConn As String = "Data Source=" & serverSqlName & ";Initial Catalog=" & databaseName & ";Integrated Security=True"
        connectionString = StrConn
        objConn = New SqlConnection(connectionString)
        objConn.ConnectionString = connectionString
        Dim dt As New DataTable()
        Try
            objConn.Open()

            'Dim sql As String = "select BatchId ,InputDate,IAcc,IAccName ,IAmt ,IChqNo ,Cur from " & viewName & " where InputDate='" & myDate & "'"
            Dim sql As String = "SELECT [Medical File],[New Medical File],[Merged],[Family],[First],[Father],[Mother],[Sex],[Blood Group],[Birth Date],[FlagBMB],[AutoID] FROM " & tablePatients & " where FlagBMB=0 order by AutoID"

            Dim adp As New SqlDataAdapter(sql, objConn)

            adp.Fill(dt)
            adp.Dispose()
            objConn.Close()
            objConn.Dispose()

            'Dim dr As DataRow
            'For i As Integer = 0 To dt.Rows.Count - 1
            '    dr = dt.NewRow
            '    Dim aray(6) As String
            '    aray(0) = dt.Rows.Item(i)(0)
            '    aray(1) = dt.Rows.Item(i)(1)
            '    aray(2) = dt.Rows.Item(i)(2)
            '    aray(3) = dt.Rows.Item(i)(3)
            '    aray(4) = dt.Rows.Item(i)(4)
            '    aray(5) = dt.Rows.Item(i)(5)
            '    aray(6) = dt.Rows.Item(i)(6)
            '    arrRows.Add(aray)
            'Next

        Catch ex As Exception
            If objConn.State = ConnectionState.Open Then
                objConn.Close()
                objConn.Dispose()
            End If
        End Try

        Return dt
    End Function


    Private Sub updateRows(ByVal idp As Integer, ByVal type As String)
        Dim arrRows As ArrayList = New ArrayList
        Dim objConn As SqlConnection
        Dim connectionString As String
        'Dim StrConn As String = "Data Source=" & serverSqlName & ";Initial Catalog=" & databaseName & ";Integrated Security=True"
        Dim StrConn As String = "Data Source=" & serverSqlName & ";Initial Catalog=" & databaseName & ";Integrated Security=True"
        connectionString = StrConn
        objConn = New SqlConnection(connectionString)
        objConn.ConnectionString = connectionString

        Try
            objConn.Open()
            Dim myCommand As SqlCommand = objConn.CreateCommand
            Dim sql As String = ""
            If type = "P" Then
                sql = "update " & tablePatients & " set FlagBMB=1 where AutoID=" & idp
            Else
                sql = "update " & tablevisit & " set FlagBMB=1 where AutoID=" & idp
            End If

            myCommand.CommandText = sql
            myCommand.ExecuteNonQuery()
            objConn.Close()
            objConn.Dispose()


        Catch ex As Exception
            If objConn.State = ConnectionState.Open Then
                objConn.Close()
                objConn.Dispose()
            End If
        End Try


    End Sub

    Private Function getVisitViewRows() As DataTable
        Dim arrRows As ArrayList = New ArrayList
        Dim objConn As SqlConnection
        Dim connectionString As String
        'Dim StrConn As String = "Data Source=" & serverSqlName & ";Initial Catalog=" & databaseName & ";Integrated Security=True"
        Dim StrConn As String = "Data Source=" & serverSqlName & ";Initial Catalog=" & databaseName & ";Integrated Security=True"
        connectionString = StrConn
        objConn = New SqlConnection(connectionString)
        objConn.ConnectionString = connectionString
        Dim dt As New DataTable()
        Try
            objConn.Open()

            'Dim sql As String = "select BatchId ,InputDate,IAcc,IAccName ,IAmt ,IChqNo ,Cur from " & viewName & " where InputDate='" & myDate & "'"
            Dim sql As String = "SELECT [Medical File],[Case Number],[New Case Number],[Admission Date],[Discharge Date],[Department],[Doctor Code],[Doctor Name],[Consultant Code],[Consultant Name],[Patient Type],[ER_Inpatient],[Coverage],[FlagType],[FlagBMB],[AutoID] FROM " & tablevisit & " where FlagBMB=0 order by AutoID"

            Dim adp As New SqlDataAdapter(sql, objConn)

            adp.Fill(dt)
            adp.Dispose()
            objConn.Close()
            objConn.Dispose()

            'Dim dr As DataRow
            'For i As Integer = 0 To dt.Rows.Count - 1
            '    dr = dt.NewRow
            '    Dim aray(6) As String
            '    aray(0) = dt.Rows.Item(i)(0)
            '    aray(1) = dt.Rows.Item(i)(1)
            '    aray(2) = dt.Rows.Item(i)(2)
            '    aray(3) = dt.Rows.Item(i)(3)
            '    aray(4) = dt.Rows.Item(i)(4)
            '    aray(5) = dt.Rows.Item(i)(5)
            '    aray(6) = dt.Rows.Item(i)(6)
            '    arrRows.Add(aray)
            'Next

        Catch ex As Exception
            If objConn.State = ConnectionState.Open Then
                objConn.Close()
                objConn.Dispose()
            End If
        End Try

        Return dt
    End Function

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ParmFunction(".\param.ini")
        ParmFunctionNip(".\Niparam.ini")
        ParmFunctionVisits(".\Visitsparam.ini")
        ParmFunctionPlateauTechnique(".\TypesParam.ini")

    End Sub

    Private Sub btnGenerate_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerate.Click
        Timer1_Tick(sender, e)



    End Sub

    Public Sub CreateToLaserficheNIP(ByVal dt As DataTable)

        Dim textToWrite As String = ""

        Try
            If dt.Rows.Count <> 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim aray(8) As String
                    If dt.Rows.Item(i)(0) Is DBNull.Value = False Then
                        MedicalFile = dt.Rows.Item(i)(0)
                    Else
                        MedicalFile = ""
                    End If
                    If dt.Rows.Item(i)(1) Is DBNull.Value = False Then
                        MedicalFileNew = dt.Rows.Item(i)(1)
                    Else
                        MedicalFileNew = ""
                    End If

                    If dt.Rows.Item(i)(3) Is DBNull.Value = False Then
                        Family = dt.Rows.Item(i)(3)
                    Else
                        Family = ""
                    End If

                    If dt.Rows.Item(i)(4) Is DBNull.Value = False Then
                        First = dt.Rows.Item(i)(4)
                    Else
                        First = ""
                    End If

                    If dt.Rows.Item(i)(5) Is DBNull.Value = False Then
                        Father = dt.Rows.Item(i)(5)
                    Else
                        Father = ""
                    End If

                    If dt.Rows.Item(i)(6) Is DBNull.Value = False Then
                        Mother = dt.Rows.Item(i)(6)
                    Else
                        Mother = ""
                    End If

                    If dt.Rows.Item(i)(7) Is DBNull.Value = False Then
                        Sex = dt.Rows.Item(i)(7)
                    Else
                        Sex = ""
                    End If
                    If dt.Rows.Item(i)(8) Is DBNull.Value = False Then
                        BloodGroup = dt.Rows.Item(i)(8)
                    Else
                        BloodGroup = ""
                    End If
                    If dt.Rows.Item(i)(9) Is DBNull.Value = False Then
                        BirthDate = dt.Rows.Item(i)(9)
                    Else
                        BirthDate = ""
                    End If

                    If dt.Rows.Item(i)(10) Is DBNull.Value = False Then
                        flagBMB = dt.Rows.Item(i)(10)
                    Else
                        flagBMB = ""
                    End If
                    If dt.Rows.Item(i)(11) Is DBNull.Value = False Then
                        patientAutoID = dt.Rows.Item(i)(11)
                    Else
                        patientAutoID = ""
                    End If


                    Try
                        connDb()

                        Dim search As LFSearch = db.CreateSearch()
                        search.Command = "{[" & templateName & "]:[" & MedicalFileField & "]=""" & MedicalFileNew & """}"
                        search.BeginSearch(True)
                        Dim hits As ILFCollection = search.GetSearchHits()

                        Dim searchNouvNIP As LFSearch = db.CreateSearch()
                        searchNouvNIP.Command = "{[" & templateName & "]:[" & MedicalFileField & "]=""" & MedicalFile & """}"
                        searchNouvNIP.BeginSearch(True)
                        Dim hitsNouvNIP As ILFCollection = searchNouvNIP.GetSearchHits()

                        Dim search1 As LFSearch = db.CreateSearch()
                        search1.Command = "{[]:[" & MedicalFileField & "]=""" & MedicalFile & """}"
                        search1.BeginSearch(True)
                        Dim hits1 As ILFCollection = search1.GetSearchHits()

                        If hits.Count = 0 And hitsNouvNIP.Count = 0 Then
                            textToWrite = DateTime.Now & " : NIP Creation"
                            writeLogFile(textToWrite, fileNIPpath)
                            textToWrite = ""
                            returnFolderNumberName(MedicalFileNew)
                            createNumerFolders(folNumberName)

                            If createPatient() = True Then
                                updateRows(patientAutoID, "P")
                            End If
                        ElseIf hitsNouvNIP.Count = 0 And hits.Count <> 0 Then
                            'do nothing
                            'delete row
                            updateRows(patientAutoID, "P")
                        ElseIf ((hitsNouvNIP.Count <> 0 And hits.Count = 0) Or MedicalFile = MedicalFileNew) And MedicalFileNew <> "" Then
                            textToWrite = DateTime.Now & " : NIP Update"
                            writeLogFile(textToWrite, fileNIPpath)
                            textToWrite = ""
                            returnFolderNumberName(MedicalFileNew)
                            createNumerFolders(folNumberName)

                            If updatePatient(hits1) = True Then
                                Try
                                    'we make try catch because if the name exist we got error duplicat exist
                                    hitsNouvNIP.Item(1).Entry.Name = MedicalFileNew
                                Catch ex As Exception

                                End Try
                                Dim fTo As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName)
                                hitsNouvNIP.Item(1).Entry.Move(fTo, True)
                                hitsNouvNIP.Item(1).Entry.Dispose()

                                ' delete the row
                                updateRows(patientAutoID, "P")
                            End If
                        ElseIf hits.Count <> 0 And hitsNouvNIP.Count <> 0 And MedicalFile <> MedicalFileNew Then
                            textToWrite = DateTime.Now & " : NIP Merged"
                            writeLogFile(textToWrite, fileNIPpath)
                            textToWrite = ""
                            returnFolderNumberName(MedicalFile)
                            createNumerFolders(folNumberName)
                            Dim folFrom As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFile)

                            returnFolderNumberName(MedicalFileNew)
                            createNumerFolders(folNumberName)

                            Dim hitsFrom As ILFCollection = folFrom.GetChildren
                            Dim j As Integer = 0
                            For Each Hit In hitsFrom
                                Try
                                    Dim entry = db.GetEntryByID(Hit.id)
                                    Dim fields As LFFieldData = entry.FieldData
                                    Dim currentTemplate As LFTemplate = fields.Template
                                    ' Locks the object for writing.
                                    If currentTemplate.Name = templateVisiName Then
                                        If fields.Field(DoctorNameField) <> "" Then
                                            applySecurity("NIPMerged", fields.Field(DoctorNameField), "Docteur")
                                        End If
                                        If fields.Field(ConsultantNameField) <> "" Then
                                            applySecurity("NIPMerged", fields.Field(ConsultantNameField), "Docteur")
                                        End If

                                    End If
                                Catch ex As Exception

                                End Try
                            Next

                            Dim folTo As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileNew)

                            'search for all folders and document that have the NIP=nouveauNIP
                            Dim search1NouvNIP As LFSearch = db.CreateSearch()
                            search1NouvNIP.Command = "{[]:[" & MedicalFileField & "]=""" & MedicalFileNew & """}"
                            search1NouvNIP.BeginSearch(True)
                            Dim hits1NouNIP As ILFCollection = search1NouvNIP.GetSearchHits()

                            If updatePatient(hits1) = True Then

                                mergePatient(folFrom, folTo)
                                updatePatient(hits1NouNIP)
                                'delete row
                                updateRows(patientAutoID, "P")
                            End If
                        End If

                    Catch ex As Exception

                        textToWrite = DateTime.Now & " : " & "Error trying to connect due to : " & ex.Message
                        writeLogFile(textToWrite, fileNIPpath)
                        textToWrite = ""

                        Try
                            If conn.IsTerminated = False Then
                                conn.Terminate()
                            End If
                        Catch ex1 As Exception

                        End Try
                    End Try

                    Try
                        If conn.IsTerminated = False Then
                            conn.Terminate()
                        End If
                    Catch ex1 As Exception

                    End Try
                Next
            End If
        Catch ex As Exception
            textToWrite = DateTime.Now & " : " & "Error due to : " & ex.Message
            writeLogFile(textToWrite, fileNIPpath)
            textToWrite = ""

            Try
                If conn.IsTerminated = False Then
                    conn.Terminate()
                End If
            Catch ex1 As Exception

            End Try
        End Try

    End Sub

    Public Sub CreateToLaserficheVisite(ByVal dt As DataTable)

        Dim textToWrite As String = ""

        Try
            If dt.Rows.Count <> 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    Dim aray(12) As String
                    If dt.Rows.Item(i)(0) Is DBNull.Value = False Then
                        MedicalFileVisit = dt.Rows.Item(i)(0)
                    Else
                        MedicalFileVisit = ""
                    End If

                    If dt.Rows.Item(i)(1) Is DBNull.Value = False Then
                        CaseNumber = dt.Rows.Item(i)(1)
                    Else
                        CaseNumber = ""
                    End If
                    If dt.Rows.Item(i)(2) Is DBNull.Value = False Then
                        CaseNumberNew = dt.Rows.Item(i)(2)
                    Else
                        CaseNumberNew = ""
                    End If
                    If dt.Rows.Item(i)(3) Is DBNull.Value = False Then
                        AdmissionDate = dt.Rows.Item(i)(3)
                    Else
                        AdmissionDate = ""
                    End If

                    If dt.Rows.Item(i)(4) Is DBNull.Value = False Then
                        DischargeDate = dt.Rows.Item(i)(4)
                    Else
                        DischargeDate = ""
                    End If

                    If dt.Rows.Item(i)(5) Is DBNull.Value = False Then
                        Department = dt.Rows.Item(i)(5)
                    Else
                        Department = ""
                    End If
                    If dt.Rows.Item(i)(6) Is DBNull.Value = False Then
                        DoctorCode = dt.Rows.Item(i)(6)
                    Else
                        DoctorCode = ""
                    End If

                    If dt.Rows.Item(i)(7) Is DBNull.Value = False Then
                        DoctorName = dt.Rows.Item(i)(7)
                    Else
                        DoctorName = ""
                    End If

                    If dt.Rows.Item(i)(8) Is DBNull.Value = False Then
                        ConsultantCode = dt.Rows.Item(i)(8)
                    Else
                        ConsultantCode = ""
                    End If
                    If dt.Rows.Item(i)(9) Is DBNull.Value = False Then
                        ConsultantName = dt.Rows.Item(i)(9)
                    Else
                        ConsultantName = ""
                    End If
                    If dt.Rows.Item(i)(10) Is DBNull.Value = False Then
                        PatientType = dt.Rows.Item(i)(10)
                    Else
                        PatientType = ""
                    End If
                    If dt.Rows.Item(i)(11) Is DBNull.Value = False Then
                        ER_Inpatient = dt.Rows.Item(i)(11)
                    Else
                        ER_Inpatient = ""
                    End If
                    If dt.Rows.Item(i)(12) Is DBNull.Value = False Then
                        Coverage = dt.Rows.Item(i)(12)
                    Else
                        Coverage = ""
                    End If
                    If dt.Rows.Item(i)(14) Is DBNull.Value = False Then
                        visitflagBMB = dt.Rows.Item(i)(14)
                    Else
                        visitflagBMB = ""
                    End If

                    If dt.Rows.Item(i)(15) Is DBNull.Value = False Then
                        visitAutoID = dt.Rows.Item(i)(15)
                    Else
                        visitAutoID = ""
                    End If

                    Try
                        connDb()

                        Dim searchvisit As LFSearch = db.CreateSearch()
                        searchvisit.Command = "{[" & templateName & "]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """}"
                        searchvisit.BeginSearch(True)
                        Dim hitsVisit As ILFCollection = searchvisit.GetSearchHits()

                        Dim searchVisit1 As LFSearch = db.CreateSearch()
                        searchVisit1.Command = "{[" & templateVisiName & "]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """,[" & CaseNumberField & "]=""" & CaseNumber & """}"
                        searchVisit1.BeginSearch(True)
                        Dim hitVisit1 As ILFCollection = searchVisit1.GetSearchHits()

                        Dim searchVisit1Fake As LFSearch = db.CreateSearch()
                        searchVisit1Fake.Command = "{[" & templateVisiName & "]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """,[" & CaseNumberField & "]=""" & CaseNumberNew & """}"
                        searchVisit1Fake.BeginSearch(True)
                        Dim hitVisit1Fake As ILFCollection = searchVisit1Fake.GetSearchHits()

                        Dim searchVisitInheritance As LFSearch = db.CreateSearch()
                        searchVisitInheritance.Command = "{[]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """,[" & CaseNumberField & "]=""" & CaseNumber & """}"
                        searchVisitInheritance.BeginSearch(True)
                        Dim hitVisitInheritance As ILFCollection = searchVisitInheritance.GetSearchHits()

                        Dim searchVisitInheritanceFake As LFSearch = db.CreateSearch()
                        searchVisitInheritanceFake.Command = "{[]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """,[" & CaseNumberField & "]=""" & CaseNumberNew & """}"
                        searchVisitInheritanceFake.BeginSearch(True)
                        Dim hitVisitInheritanceFake As ILFCollection = searchVisitInheritanceFake.GetSearchHits()

                        returnFolderNumberName(MedicalFileVisit)
                        If hitsVisit.Count = 1 Then
                            If hitVisit1.Count = 0 And hitVisit1Fake.Count = 0 Then
                                textToWrite = DateTime.Now & " : visit Creation"
                                writeLogFile(textToWrite, fileNIPpath)
                                textToWrite = ""
                                If createVisits(hitsVisit.Item(1).Entry.Name) = True Then
                                    Dim searchAllVisitUnderNIP As LFSearch = db.CreateSearch()
                                    searchAllVisitUnderNIP.Command = "{[" & templateVisiName & "]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """}"
                                    searchAllVisitUnderNIP.BeginSearch(True)
                                    Dim hitAllvisitUnderNip As ILFCollection = searchAllVisitUnderNIP.GetSearchHits()

                                    If DoctorName <> "" Then
                                        createGroups(DoctorName)
                                        applySecurity("", DoctorName, "Docteur")
                                        For Each hit In hitAllvisitUnderNip
                                            Dim entry = db.GetEntryByID(hit.Entry.id)
                                            applySecurity(entry.Name, DoctorName, "Visite")
                                        Next
                                    End If
                                    If ConsultantName <> "" Then
                                        createGroups(ConsultantName)
                                        applySecurity("", DoctorName, "Docteur")
                                        For Each hit In hitAllvisitUnderNip
                                            Dim entry = db.GetEntryByID(hit.Entry.id)
                                            applySecurity(entry.Name, ConsultantName, "Visite")
                                        Next
                                    End If


                                    'delete row
                                    updateRows(visitAutoID, "V")
                                End If

                            ElseIf hitVisit1.Count = 0 And hitVisit1Fake.Count <> 0 Then
                                'do nothing
                                'delete row
                                updateRows(visitAutoID, "V")
                            ElseIf ((hitVisit1.Count <> 0 And hitVisit1Fake.Count = 0) Or CaseNumber = CaseNumberNew) And CaseNumberNew <> "" Then
                                textToWrite = DateTime.Now & " : update visit"
                                writeLogFile(textToWrite, fileNIPpath)
                                textToWrite = ""

                                If updatevisits(hitVisitInheritance, "") = True Then
                                    hitVisit1.Item(1).Entry.Name = CaseNumberNew

                                    If DoctorName <> "" Then
                                        If DoctorName <> doctorNameBeforeUpdate Then
                                            createGroups(DoctorName)
                                            Dim searchAllVisitUnderNIP As LFSearch = db.CreateSearch()
                                            searchAllVisitUnderNIP.Command = "{[" & templateVisiName & "]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """}"
                                            searchAllVisitUnderNIP.BeginSearch(True)
                                            Dim hitAllvisitUnderNip As ILFCollection = searchAllVisitUnderNIP.GetSearchHits()

                                            If doctorNameBeforeUpdate <> "" Then
                                                RemoveApplySecurity(CaseNumberNew, doctorNameBeforeUpdate, "Visite")
                                            End If
                                            applySecurity("", DoctorName, "Docteur")
                                            For Each hit In hitAllvisitUnderNip
                                                Dim entry = db.GetEntryByID(hit.Entry.id)
                                                applySecurity(entry.Name, DoctorName, "Visite")
                                            Next
                                        End If
                                    ElseIf DoctorName = "" Then
                                        If doctorNameBeforeUpdate <> "" Then
                                            RemoveApplySecurity(CaseNumberNew, doctorNameBeforeUpdate, "Visite")
                                        End If
                                    End If

                                    If ConsultantName <> "" Then
                                        If ConsultantName <> ConsulanttNameBeforeUpdate Then
                                            createGroups(ConsultantName)
                                            Dim searchAllVisitUnderNIP As LFSearch = db.CreateSearch()
                                            searchAllVisitUnderNIP.Command = "{[" & templateVisiName & "]:[" & MedicalFileField & "]=""" & MedicalFileVisit & """}"
                                            searchAllVisitUnderNIP.BeginSearch(True)
                                            Dim hitAllvisitUnderNip As ILFCollection = searchAllVisitUnderNIP.GetSearchHits()

                                            If ConsulanttNameBeforeUpdate <> "" Then
                                                RemoveApplySecurity(CaseNumberNew, ConsulanttNameBeforeUpdate, "Visite")
                                            End If
                                            applySecurity("", ConsultantName, "Docteur")
                                            For Each hit In hitAllvisitUnderNip
                                                Dim entry = db.GetEntryByID(hit.Entry.id)
                                                applySecurity(entry.Name, ConsultantName, "Visite")
                                            Next
                                        End If
                                    ElseIf ConsultantName = "" Then
                                        If ConsulanttNameBeforeUpdate <> "" Then
                                            RemoveApplySecurity(CaseNumberNew, ConsulanttNameBeforeUpdate, "Visite")
                                        End If
                                    End If

                                    'delete row
                                    updateRows(visitAutoID, "V")
                                End If
                            ElseIf hitVisit1.Count <> 0 And hitVisit1Fake.Count <> 0 And CaseNumber <> CaseNumberNew Then
                                textToWrite = DateTime.Now & " : merged visit"
                                writeLogFile(textToWrite, fileNIPpath)
                                textToWrite = ""
                                Dim folFrom As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumber)
                                Dim folTo As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumberNew)
                                If updatevisits(hitVisitInheritance, "") = True Then
                                    mergeVisits(folFrom, folTo)
                                    updatevisits(hitVisitInheritanceFake, "")
                                    updateRows(visitAutoID, "V")
                                End If
                            End If
                        ElseIf hitsVisit.Count > 1 Then
                            textToWrite = DateTime.Now & " : Duplication Data... The NIP " & MedicalFileVisit & " exist more than on time"
                            writeLogFile(textToWrite, fileNIPpath)
                            textToWrite = ""
                            updateRows(visitAutoID, "V")

                        ElseIf hitsVisit.Count = 0 Then
                            textToWrite = DateTime.Now & " : The NIP " & MedicalFileVisit & " is not created for this visit " & CaseNumber
                            writeLogFile(textToWrite, fileNIPpath)
                            textToWrite = ""
                            updateRows(visitAutoID, "V")
                            'ElseIf hitVisit1.Count <> 0 And hitVisit1Fake.Count <> 0 And CaseNumber <> CaseNumberNew Then
                            '    textToWrite = DateTime.Now & " : merged visit"
                            '    writeLogFile(textToWrite, fileNIPpath)
                            '    textToWrite = ""
                        End If

                    Catch ex As Exception

                        textToWrite = DateTime.Now & " : " & "Error trying to connect due to : " & ex.Message
                        writeLogFile(textToWrite, fileNIPpath)
                        textToWrite = ""

                        Try
                            If conn.IsTerminated = False Then
                                conn.Terminate()
                            End If
                        Catch ex1 As Exception

                        End Try
                    End Try

                    Try
                        If conn.IsTerminated = False Then
                            conn.Terminate()
                        End If
                    Catch ex1 As Exception

                    End Try
                Next
            End If

        Catch ex As Exception
            textToWrite = DateTime.Now & " : " & "Error due to : " & ex.Message
            writeLogFile(textToWrite, fileNIPpath)
            textToWrite = ""

            Try
                If conn.IsTerminated = False Then
                    conn.Terminate()
                End If
            Catch ex1 As Exception

            End Try
        End Try
    End Sub


    Public Function createPatient() As Boolean
        Dim textToWrite As String = ""
        Dim ok As Boolean = False

        Try

            Dim fol As New LFFolder
            Dim ParentFol As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName)
            fol.Create(MedicalFileNew, ParentFol, True)

            Dim pTempName As LFTemplate = db.GetTemplateByName(templateName)
            Dim fields As LFFieldData = fol.FieldData
            fields.LockObject(Lock_Type.LOCK_TYPE_WRITE)
            fields.Template = pTempName

            fields.Field(MedicalFileField) = MedicalFileNew
            fields.Field(FamilyField) = Family
            fields.Field(FirstField) = First
            fields.Field(FatherField) = Father
            fields.Field(MotherField) = Mother
            fields.Field(SexField) = Sex
            fields.Field(BloodGroupField) = BloodGroup
            fields.Field(BirthDateField) = BirthDate

            fields.Update()
            fol.Dispose()


            Dim foll As New LFFolder
            Dim ParentFoll As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileNew)
            foll.Create("Patient Personal ID", ParentFoll, True)
            foll.Dispose()

            Dim doc As New LFDocument
            Dim Vol As LFVolume = db.GetVolumeByName(volum)
            Dim ParentFolll As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileNew & "\Patient Personal ID")
            doc.Create(MedicalFileNew & "-Patient Personal ID", ParentFolll, Vol, True)


            'Dim fieldss As LFFieldData = doc.FieldData
            'fieldss.LockObject(Lock_Type.LOCK_TYPE_WRITE)
            'fieldss.Template = pTempName

            'fieldss.Field(MedicalFileField) = MedicalFileNew
            'fieldss.Field(FamilyField) = Family
            'fieldss.Field(FirstField) = First
            'fieldss.Field(FatherField) = Father
            'fieldss.Field(MotherField) = Mother
            'fieldss.Field(SexField) = Sex
            'fieldss.Field(BloodGroupField) = BloodGroup
            'fieldss.Field(BirthDateField) = BirthDate

            'fieldss.Update()
            doc.Dispose()

            ok = True
            textToWrite = DateTime.Now & " : " & "The medical file " & MedicalFileNew & " has been created"
            writeLogFile(textToWrite, fileNIPpath)
        Catch ex As Exception
            ok = False
            textToWrite = DateTime.Now & " : " & "The medical file " & MedicalFileNew & " could not created due to : " & ex.Message
            writeLogFile(textToWrite, fileNIPpath)
        End Try

        Return ok
    End Function

    Public Function updatePatient(ByVal hits As ILFCollection) As Boolean
        Dim ok As Boolean = False
        Dim textToWrite As String = ""

        For Each Hit As LFSearchHit In hits

            Dim entry = db.GetEntryByID(Hit.Entry.id)
            Try
                Dim fields As LFFieldData = entry.FieldData
                Dim currentTemplate As LFTemplate = fields.Template
                ' Locks the object for writing.

                fields.LockObject(Lock_Type.LOCK_TYPE_WRITE)
                fields.Field(MedicalFileField) = MedicalFileNew
                fields.Field(FamilyField) = Family
                fields.Field(FirstField) = First
                fields.Field(FatherField) = Father
                If currentTemplate.Name = templateName Then
                    fields.Field(MotherField) = Mother
                    fields.Field(SexField) = Sex
                    fields.Field(BloodGroupField) = BloodGroup
                    fields.Field(BirthDateField) = BirthDate
                End If
                fields.Update()
                entry.Dispose()
                ok = True
                textToWrite = DateTime.Now & " : " & "" & entry.Name & " has been updated from " & MedicalFile & " to " & MedicalFileNew
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            Catch ex As Exception
                ok = False
                textToWrite = DateTime.Now & " : " & "" & entry.Name & " could not be updated from " & MedicalFile & " to " & MedicalFileNew
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            End Try
        Next

        textToWrite = DateTime.Now & " : " & " all entries are updated from " & MedicalFile & " to " & MedicalFileNew
        writeLogFile(textToWrite, fileNIPpath)
        textToWrite = ""
        Return ok
    End Function

    Public Function mergePatient(ByVal folFrom As LFFolder, ByVal folTo As LFFolder) As Boolean
        Dim hitsFrom As ILFCollection = folFrom.GetChildren
        Dim ok As Boolean = False
        Dim textToWrite As String = ""

        For Each Hit In hitsFrom
            Dim entry = db.GetEntryByID(Hit.id)
            Try
                entry.Move(folTo, True)
                entry.Dispose()
                folTo.Dispose()
                ok = True
                textToWrite = DateTime.Now & " : " & "" & entry.Name & " has been merged from " & MedicalFile & " to " & MedicalFileNew
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            Catch ex As Exception
                ok = False
                textToWrite = DateTime.Now & " : " & "" & entry.Name & " could not be merged from " & MedicalFile & " to " & MedicalFileNew
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            End Try
        Next
        folFrom.Delete()
        textToWrite = DateTime.Now & " : All the entries are merged from " & MedicalFile & " to " & MedicalFileNew
        writeLogFile(textToWrite, fileNIPpath)
        textToWrite = ""
        Return ok
    End Function

    ''' <summary>
    ''' Create the visits folder under NIP into laserfiche and assign the visite template
    ''' </summary>
    ''' <param name="entryName"></param>
    ''' <remarks>Some fields are inherited from the parent folder(NIP)</remarks>
    Public Function createVisits(ByVal entryName As String) As Boolean

        Dim textToWrite As String = ""
        Dim ok As Boolean = False


        Try
            Dim folstrut As LFFolder = db.GetEntryByPath("\Folder structure")
            Dim folTypeTo As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & entryName & "\")
            Dim fol As New LFFolder
            Dim ParentFol As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & entryName)
            ' fol.Create(CaseNumberNew, ParentFol, True)
            fol.CreateCopyOf(folstrut, CaseNumberNew, folTypeTo, True)

            folTypeTo.Dispose()
            Dim visitTempName As LFTemplate = db.GetTemplateByName(templateVisiName)
            Dim fields As LFFieldData = fol.FieldData
            fields.LockObject(Lock_Type.LOCK_TYPE_WRITE)
            fields.Template = visitTempName

            Dim FD As LFFieldData = ParentFol.FieldData

            fields.Field(CaseNumberField) = CaseNumberNew
            fields.Field(AdmissionDateField) = AdmissionDate
            fields.Field(DischargeDateField) = DischargeDate
            fields.Field(DepartmentField) = Department
            fields.Field(DoctorCodeField) = DoctorCode
            fields.Field(DoctorNameField) = DoctorName
            fields.Field(ConsultantCodeField) = ConsultantCode
            fields.Field(ConsultantNameField) = ConsultantName
            fields.Field(PatientTypeField) = PatientType
            fields.Field(ER_InpatientField) = ER_Inpatient
            fields.Field(CoverageField) = Coverage

            fields.Field(MedicalFileField) = FD.Field(MedicalFileField)
            fields.Field(FirstField) = FD.Field(FirstField)
            fields.Field(FatherField) = FD.Field(FatherField)
            fields.Field(FamilyField) = FD.Field(FamilyField)

            fields.Update()
            fol.Dispose()



            'Visit Administration Scan Documents
            '' createNumerFoldersUnderVisits("Visit Administration Scan Documents", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            applySecurity("", "Admission", "Admission")
            ''createNumerFoldersUnderVisits("Patient Personal ID", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents")
            createDocumentUnderTypes("Patient Personal ID", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents\Patient Personal ID")

            '' createNumerFoldersUnderVisits("3rd Parties Approval", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents")
            createDocumentUnderTypes("3rd Parties Approval", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents\3rd Parties Approval")

            ''createNumerFoldersUnderVisits("Internal Approval Request", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents")
            createDocumentUnderTypes("Internal Approval Request", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents\Internal Approval Request")

            '' createNumerFoldersUnderVisits("Other Documents", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents")
            createDocumentUnderTypes("Other Documents", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Administration Scan Documents\Other Documents")

            'Visit Nursing Scan Records
            ''createNumerFoldersUnderVisits("Visit Nursing Scan Records", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''createNumerFoldersUnderVisits("Basic Nursing", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Nursing Scan Records")
            '''''  createDocumentUnderTypes("Basic Nursing", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Nursing Scan Records\Basic Nursing")

            '' createNumerFoldersUnderVisits("Nursing Observation", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Nursing Scan Records")
            ''''  createDocumentUnderTypes("Nursing Observation", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Nursing Scan Records\Nursing Observation")

            'Visit Physican Scan Records
            ''createNumerFoldersUnderVisits("Visit Physican Scan Records", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''''  createDocumentUnderTypes("Visit Physican Scan Records", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Physican Scan Records")

            'Visit  Surgical Scan  Records
            ''createNumerFoldersUnderVisits("Visit  Surgical Scan  Records", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''createNumerFoldersUnderVisits("General Surgery", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records")
            ''''  createDocumentUnderTypes("General Surgery", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records\General Surgery")

            ''createNumerFoldersUnderVisits("Cardio Vascular Surgery", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records")
            '''' createDocumentUnderTypes("Cardio Vascular Surgery", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records\Cardio Vascular Surgery")

            '' createNumerFoldersUnderVisits("Cardio Vascular Lab", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records")
            ''''  createDocumentUnderTypes("Cardio Vascular Lab", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records\Cardio Vascular Lab")

            ''createNumerFoldersUnderVisits("Labour & New Born", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records")
            ''''  createDocumentUnderTypes("Labour & New Born", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Surgical Scan  Records\Labour & New Born")

            'Visit  Laboratory Results
            ''createNumerFoldersUnderVisits("Visit  Laboratory Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''createNumerFoldersUnderVisits("Laboratory Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Laboratory Results")
            ''''  createDocumentUnderTypes("Laboratory Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Laboratory Results\Laboratory Results")

            ''createNumerFoldersUnderVisits("Pathology Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Laboratory Results")
            ''''  createDocumentUnderTypes("Pathology Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Laboratory Results\Pathology Results")

            'Visit  Radiology   Results
            '' createNumerFoldersUnderVisits("Visit  Radiology   Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''createNumerFoldersUnderVisits("General Radiology", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''   createDocumentUnderTypes("General Radiology", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\General Radiology")

            '' createNumerFoldersUnderVisits("ct-scan", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''  createDocumentUnderTypes("ct-scan", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\ct-scan")

            ''createNumerFoldersUnderVisits("mri", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''  createDocumentUnderTypes("mri", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\mri")

            ''createNumerFoldersUnderVisits("VCT", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''  createDocumentUnderTypes("VCT", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\VCT")

            ''createNumerFoldersUnderVisits("Echography", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''  createDocumentUnderTypes("Echography", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\Echography")

            '' createNumerFoldersUnderVisits("Nucluare Radiology", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''  createDocumentUnderTypes("Nucluare Radiology", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\Nucluare Radiology")

            ''createNumerFoldersUnderVisits("Osteo Radiology", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''   createDocumentUnderTypes("Osteo Radiology", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\Osteo Radiology")

            '' createNumerFoldersUnderVisits("IVF", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results")
            ''''   createDocumentUnderTypes("IVF", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Radiology   Results\IVF")

            'Visit  Other Exams Results
            ''createNumerFoldersUnderVisits("Visit  Other Exams Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''createNumerFoldersUnderVisits("EMG Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''   createDocumentUnderTypes("EMG Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\EMG Results")

            ''createNumerFoldersUnderVisits("EEG Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''   createDocumentUnderTypes("EEG Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\EEG Results")

            '' createNumerFoldersUnderVisits("EKG Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''   createDocumentUnderTypes("EKG Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\EKG Results")

            '' createNumerFoldersUnderVisits("Strees Test Result", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''   createDocumentUnderTypes("Strees Test Result", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\Strees Test Result")

            '' createNumerFoldersUnderVisits("Spiro Meter Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''   createDocumentUnderTypes("Spiro Meter Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\Spiro Meter Results")

            ''createNumerFoldersUnderVisits("Cardiography", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''   createDocumentUnderTypes("Cardiography", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\Cardiography")

            '' createNumerFoldersUnderVisits("Holter Monitor", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''    createDocumentUnderTypes("Holter Monitor", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\Holter Monitor")

            '' createNumerFoldersUnderVisits("Osteodensitometry Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results")
            ''''   createDocumentUnderTypes("Osteodensitometry Results", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit  Other Exams Results\Osteodensitometry Results")

            'Visit Discharge Documents
            '' createNumerFoldersUnderVisits("Visit Discharge Documents", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ' createDocumentUnderTypes("Visit Discharge Documents", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Discharge Documents")

            '' createNumerFoldersUnderVisits("Emergency", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Discharge Documents")
            ''''  createDocumentUnderTypes("Emergency", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Visit Discharge Documents\Emergency")

            'Other Medical  Documents
            ''createNumerFoldersUnderVisits("Other Medical  Documents", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''''  createDocumentUnderTypes("Other Medical  Documents", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Other Medical  Documents")

            'Dialysis
            ''createNumerFoldersUnderVisits("Dialysis", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew)
            ''''  createDocumentUnderTypes("Dialysis", pathInLF & "\" & folNumberName & "\" & entryName & "\" & CaseNumberNew & "\Dialysis")



            ok = True
            textToWrite = DateTime.Now & " : " & "The Visite " & CaseNumberNew & " has been created"
            writeLogFile(textToWrite, fileNIPpath)
            textToWrite = ""
        Catch ex As Exception
            ok = False
            textToWrite = DateTime.Now & " : " & "The Visite " & CaseNumber & " could not created due to : " & ex.Message
            writeLogFile(textToWrite, fileNIPpath)
        End Try

        Return ok

    End Function
    Public Function updatevisits(ByVal hits As ILFCollection, ByVal func As String) As Boolean
        Dim ok As Boolean = False
        Dim textToWrite As String = ""
        doctorNameBeforeUpdate = ""
        ConsulanttNameBeforeUpdate = ""
        For Each Hit As LFSearchHit In hits
            Dim entry = db.GetEntryByID(Hit.Entry.id)
            Try
                Dim fields As LFFieldData = entry.FieldData
                ' Locks the object for writing.
                Dim CurrentTemp As LFTemplate = fields.Template
                If CurrentTemp.Name = templateVisiName Then

                    doctorNameBeforeUpdate = fields.Field(DoctorNameField)
                    ConsulanttNameBeforeUpdate = fields.Field(ConsultantNameField)

                End If

                If CurrentTemp.Name <> templateName Then
                    fields.LockObject(Lock_Type.LOCK_TYPE_WRITE)
                    fields.Field(CaseNumberField) = CaseNumberNew
                    fields.Field(AdmissionDateField) = AdmissionDate
                    fields.Field(DischargeDateField) = DischargeDate
                    fields.Field(DepartmentField) = Department
                    fields.Field(DoctorCodeField) = DoctorCode
                    fields.Field(DoctorNameField) = DoctorName
                    fields.Field(ConsultantCodeField) = ConsultantCode
                    fields.Field(ConsultantNameField) = ConsultantName
                    fields.Field(PatientTypeField) = PatientType

                    If CurrentTemp.Name = templateVisiName Then
                        fields.Field(ER_InpatientField) = ER_Inpatient
                        fields.Field(CoverageField) = Coverage
                    End If
                End If
                fields.Update()
                entry.Dispose()

                ok = True
                textToWrite = DateTime.Now & " : " & "" & entry.Name & " has been updated "
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            Catch ex As Exception
                ok = False
                textToWrite = DateTime.Now & " : " & "" & entry.Name & " could not be updated  due to " & ex.Message
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            End Try
        Next

        textToWrite = DateTime.Now & " : All the entries are updated from " & CaseNumber & " to " & CaseNumberNew
        writeLogFile(textToWrite, fileNIPpath)
        textToWrite = ""

        Return ok
    End Function
    Public Sub mergeVisits(ByVal folFrom As LFFolder, ByVal folTo As LFFolder)
        Dim hitsFrom As ILFCollection = folFrom.GetChildren
        Dim hitsTo As ILFCollection = folTo.GetChildren

        'to find if the type exist to merge his content
        For i As Integer = 0 To hitsFrom.Count - 1
            Dim nameExist As Boolean = False
            Dim entryFrom As ILFEntry = hitsFrom.Item(i + 1)
            For j As Integer = 0 To hitsTo.Count - 1

                Dim entryto As ILFEntry = hitsTo.Item(j + 1)

                If entryFrom.Name <> entryto.Name Then
                    nameExist = False
                Else
                    nameExist = True
                    Exit For
                End If
            Next

            If nameExist = False Then
                Dim entry = db.GetEntryByID(entryFrom.ID)
                entry.Move(folTo, True)
                entry.Dispose()
                folTo.Dispose()
            Else
                Try
                    Dim folTypeFrom As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumber & "\" & entryFrom.Name)
                    Dim folTypeTo As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumberNew & "\" & entryFrom.Name)

                    'merger that dont have 2 subfolders confirm and not confirm

                    If entryFrom.Name = entryFrom.Name = "Visit Physican Scan Records" Or entryFrom.Name = "Other Medical  Documents" Or entryFrom.Name = "Dialysis" Then
                        For Each Hit In folTypeFrom.GetChildren
                            Dim entry = db.GetEntryByID(Hit.id)
                            entry.Move(folTypeTo, True)
                            entry.Dispose()
                            folTo.Dispose()
                        Next
                    Else
                        'merger that  have 2 subfolders confirm and not confirm
                        For Each Hit In folTypeFrom.GetChildren
                            Dim entry = db.GetEntryByID(Hit.id)

                            Dim folTypeFromConfirm As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumber & "\" & entryFrom.Name & "\" & entry.Name)
                            Dim folTypeToConfirm As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumberNew & "\" & entryFrom.Name & "\" & entry.Name)

                            For Each Hit1 In folTypeFromConfirm.GetChildren
                                Dim entryConf = db.GetEntryByID(Hit1.id)
                                entryConf.Move(folTypeToConfirm, True)
                                entryConf.Dispose()
                                folTypeToConfirm.Dispose()
                            Next
                        Next
                        folTypeFrom.Delete()

                    End If
                Catch ex As Exception
                    Dim entry = db.GetEntryByID(entryFrom.ID)
                    entry.Move(folTo, True)
                    entry.Dispose()
                    folTo.Dispose()
                End Try

            End If
        Next
        folFrom.Delete()
    End Sub
    Public Sub createNumerFolders(ByVal folderName As String)
        Try
            Dim fol As New LFFolder
            Dim ParentFol As LFFolder = db.GetEntryByPath(pathInLF)
            fol.Create(folderName, ParentFol, False)
            fol.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub createNumerFoldersUnderVisits(ByVal folderName As String, ByVal folderPath As String)
        Try
            Dim fol As New LFFolder
            Dim ParentFol As LFFolder = db.GetEntryByPath(folderPath)
            fol.Create(folderName, ParentFol, False)
            fol.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub createDocumentUnderTypes(ByVal docName As String, ByVal folderPath As String)
        Try
            'writeLogFile(DateTime.Now & " : ""Document preparation", fileNIPpath)
            Dim Doc As New LFDocument
            Dim Vol As LFVolume = db.GetVolumeByName(volum)
            Dim ParentFol As LFFolder = db.GetEntryByPath(folderPath)
            Doc.Create(docName, ParentFol, Vol, True)
            Doc.Dispose()

            Dim folPatient As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit)
            Dim folVisit As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumberNew)

            Dim visitTypeResultTempName As LFTemplate = db.GetTemplateByName(templatePlatTechnique)
            Dim fields As LFFieldData = Doc.FieldData
            fields.LockObject(Lock_Type.LOCK_TYPE_WRITE)
            fields.Template = visitTypeResultTempName
            'writeLogFile(DateTime.Now & " : Document has been created", fileNIPpath)
            Dim FD As LFFieldData = folPatient.FieldData
            Dim FDVisit As LFFieldData = folVisit.FieldData

            fields.Field(MedicalFileField) = FD.Field(MedicalFileField)
            fields.Field(FirstField) = FD.Field(FirstField)
            fields.Field(FamilyField) = FD.Field(FamilyField)
            fields.Field(FatherField) = FD.Field(FatherField)

            fields.Field(CaseNumberField) = FDVisit.Field(CaseNumberField)
            fields.Field(DoctorCodeField) = FDVisit.Field(DoctorCodeField)
            fields.Field(DoctorNameField) = FDVisit.Field(DoctorNameField)
            fields.Field(ConsultantCodeField) = FDVisit.Field(ConsultantCodeField)
            fields.Field(ConsultantNameField) = FDVisit.Field(ConsultantNameField)
            fields.Field(DepartmentField) = FDVisit.Field(DepartmentField)
            fields.Field(PatientTypeField) = FDVisit.Field(PatientTypeField)
            fields.Field(AdmissionDateField) = FDVisit.Field(AdmissionDateField)
            fields.Field(DischargeDateField) = FDVisit.Field(DischargeDateField)

            If folderPath.Split("\").Length > 6 Then
                fields.Field(Type1Field) = folderPath.Split("\")(folderPath.Split("\").Length - 2)
            Else
                fields.Field(Type1Field) = folderPath.Split("\")(folderPath.Split("\").Length - 1)
            End If

            fields.Field(type2Field) = docName
            fields.Update()
            Doc.Dispose()
            ' writeLogFile(DateTime.Now & " : ""fields has been updated", fileNIPpath)
        Catch ex As Exception

        End Try

    End Sub
    Public Sub createGroups(ByVal groupdeSec As String)
        Try
            If groupdeSec <> "" Then
                Dim newGroup As New LFGroup
                'connDb()
                newGroup.Create(db, groupdeSec)

                'Dim Group1 As LFGroup = db.GetTrusteeByName("Voir Tous les Patients")
                'Group1.AddTrustee(newGroup)
                'Group1.Update()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub applySecurity(ByVal entryName As String, ByVal groupdeSec As String, ByVal visiteOrTypeOrDoctor As String)
        Dim textToWrite As String = ""
        Try
            If visiteOrTypeOrDoctor = "Docteur" Then

                Dim secureGoupD As LFGroup = db.GetTrusteeByName(groupdeSec)
                secureGoupD.FeatureRight(Feature_Right.FEATURE_RIGHT_SEARCH) = True
                secureGoupD.Update()

                Try
                    Dim secureFolderD As LFFolder = db.GetEntryByPath(pathInLF)
                    Dim AccFolD As LFAccessRights = secureFolderD.Rights
                    Dim folDACE As LFAccessControlEntry
                    folDACE = AccFolD.AddAccessControlEntry(secureGoupD, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                    folDACE.Allow(Access_Right.ENTRY_READ) = True
                    folDACE.Allow(Access_Right.ENTRY_BROWSE) = True
                    AccFolD.Update()
                Catch ex As Exception

                End Try


                Try
                    Dim secureFolderPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName)
                    Dim secureGoupVPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                    Dim AccFolPP As LFAccessRights = secureFolderPP.Rights
                    Dim folACEPP As LFAccessControlEntry
                    folACEPP = AccFolPP.AddAccessControlEntry(secureGoupVPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                    folACEPP.Allow(Access_Right.ENTRY_READ) = True
                    folACEPP.Allow(Access_Right.ENTRY_BROWSE) = True
                    AccFolPP.Update()
                Catch ex As Exception

                End Try
                If entryName <> "NIPMerged" Then
                    Try
                        Dim secureFolderPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit)
                        Dim secureGoupVPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                        Dim AccFolPPP As LFAccessRights = secureFolderPPP.Rights
                        Dim folACEPPP As LFAccessControlEntry
                        folACEPPP = AccFolPPP.AddAccessControlEntry(secureGoupVPPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                        folACEPPP.Allow(Access_Right.ENTRY_READ) = True
                        folACEPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                        AccFolPPP.Update()
                    Catch ex As Exception

                    End Try
                    Try
                        Dim secureFolderPPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\Patient Personal ID")
                        Dim secureGoupVPPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                        Dim AccFolPPPP As LFAccessRights = secureFolderPPPP.Rights
                        Dim folACEPPPP As LFAccessControlEntry
                        folACEPPPP = AccFolPPPP.AddAccessControlEntry(secureGoupVPPPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                        folACEPPPP.Allow(Access_Right.ENTRY_READ) = True
                        folACEPPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                        AccFolPPPP.Update()
                    Catch ex As Exception

                    End Try

                Else
                    Try
                        Dim secureFolderPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileNew)
                        Dim secureGoupVPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                        Dim AccFolPPP As LFAccessRights = secureFolderPPP.Rights
                        Dim folACEPPP As LFAccessControlEntry
                        folACEPPP = AccFolPPP.AddAccessControlEntry(secureGoupVPPP, EntryAce_Scope.SCOPE_TOTAL)
                        folACEPPP.Allow(Access_Right.ENTRY_READ) = True
                        folACEPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                        AccFolPPP.Update()
                    Catch ex As Exception

                    End Try

                    Try
                        Dim secureFolderPPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileNew & "\Patient Personal ID")
                        Dim secureGoupVPPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                        Dim AccFolPPPP As LFAccessRights = secureFolderPPPP.Rights
                        Dim folACEPPPP As LFAccessControlEntry
                        folACEPPPP = AccFolPPPP.AddAccessControlEntry(secureGoupVPPPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                        folACEPPPP.Allow(Access_Right.ENTRY_READ) = True
                        folACEPPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                        AccFolPPPP.Update()
                    Catch ex As Exception

                    End Try
                End If


                'Try
                '    Dim secureFolderPPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & entryName)
                '    Dim secureGoupVPPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                '    Dim AccFolPPPP As LFAccessRights = secureFolderPPPP.Rights
                '    Dim folACEPPPP As LFAccessControlEntry
                '    folACEPPPP = AccFolPPPP.AddAccessControlEntry(secureGoupVPPPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                '    folACEPPPP.Allow(Access_Right.ENTRY_READ) = True
                '    folACEPPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                '    AccFolPPPP.Update()
                'Catch ex As Exception

                'End Try

                textToWrite = DateTime.Now & " : " & "" & "The docteur rights has been assigned to the patient : " & MedicalFileVisit
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            End If
            If visiteOrTypeOrDoctor = "Visite" Then

                Try
                    Dim secureFolderPPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & entryName)
                    Dim secureGoupVPPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                    Dim AccFolPPPP As LFAccessRights = secureFolderPPPP.Rights
                    Dim folACEPPPP As LFAccessControlEntry
                    folACEPPPP = AccFolPPPP.AddAccessControlEntry(secureGoupVPPPP, EntryAce_Scope.SCOPE_TOTAL)
                    folACEPPPP.Allow(Access_Right.ENTRY_READ) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                    AccFolPPPP.Update()
                Catch ex As Exception

                End Try

                textToWrite = DateTime.Now & " : " & "" & "The security " & groupdeSec & " has been assigned to the visite " & entryName
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            End If

            If visiteOrTypeOrDoctor = "Admission" Then
                Dim secureGoupD As LFGroup = db.GetTrusteeByName(groupdeSec)
                'secureGoupD.FeatureRight(Feature_Right.FEATURE_RIGHT_SEARCH) = True
                'secureGoupD.FeatureRight(Feature_Right.FEATURE_RIGHT_SCAN) = True
                'secureGoupD.FeatureRight(Feature_Right.FEATURE_RIGHT_PRINT) = True
                'secureGoupD.FeatureRight(Feature_Right.FEATURE_RIGHT_IMPORT) = True
                'secureGoupD.FeatureRight(Feature_Right.FEATURE_RIGHT_EXPORT) = True
                'secureGoupD.Update()

                Try
                    Dim secureFolderD As LFFolder = db.GetEntryByPath(pathInLF)
                    Dim AccFolD As LFAccessRights = secureFolderD.Rights
                    Dim folDACE As LFAccessControlEntry
                    folDACE = AccFolD.AddAccessControlEntry(secureGoupD, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                    folDACE.Allow(Access_Right.ENTRY_READ) = True
                    folDACE.Allow(Access_Right.ENTRY_BROWSE) = True
                    AccFolD.Update()
                Catch ex As Exception

                End Try


                Try
                    Dim secureFolderPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName)
                    Dim secureGoupVPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                    Dim AccFolPP As LFAccessRights = secureFolderPP.Rights
                    Dim folACEPP As LFAccessControlEntry
                    folACEPP = AccFolPP.AddAccessControlEntry(secureGoupVPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                    folACEPP.Allow(Access_Right.ENTRY_READ) = True
                    folACEPP.Allow(Access_Right.ENTRY_BROWSE) = True
                    AccFolPP.Update()
                Catch ex As Exception

                End Try

                Try
                    Dim secureFolderPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit)
                    Dim secureGoupVPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                    Dim AccFolPPP As LFAccessRights = secureFolderPPP.Rights
                    Dim folACEPPP As LFAccessControlEntry
                    folACEPPP = AccFolPPP.AddAccessControlEntry(secureGoupVPPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                    folACEPPP.Allow(Access_Right.ENTRY_READ) = True
                    folACEPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                    AccFolPPP.Update()
                Catch ex As Exception

                End Try

                Try
                    Dim secureFolderPPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumberNew)
                    Dim secureGoupVPPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                    Dim AccFolPPPP As LFAccessRights = secureFolderPPPP.Rights
                    Dim folACEPPPP As LFAccessControlEntry
                    folACEPPPP = AccFolPPPP.AddAccessControlEntry(secureGoupVPPPP, EntryAce_Scope.SCOPE_ENTRY_ONLY)
                    folACEPPPP.Allow(Access_Right.ENTRY_READ) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                    AccFolPPPP.Update()
                Catch ex As Exception

                End Try

                Try
                    Dim secureFolderPPPP As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & CaseNumberNew & "\Visit Administration Scan Documents")
                    Dim secureGoupVPPPP As LFGroup = db.GetTrusteeByName(groupdeSec)
                    Dim AccFolPPPP As LFAccessRights = secureFolderPPPP.Rights
                    Dim folACEPPPP As LFAccessControlEntry
                    folACEPPPP = AccFolPPPP.AddAccessControlEntry(secureGoupVPPPP, EntryAce_Scope.SCOPE_TOTAL)
                    folACEPPPP.Allow(Access_Right.ENTRY_READ) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_BROWSE) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_ANNOTATE) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_REMOVE_PAGE) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_SEE_ANNOTATIONS) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_WRITE_CONTENT) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_CREATE_DOC) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_ADD_PAGE) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_CREATE_FOLDER) = True
                    folACEPPPP.Allow(Access_Right.ENTRY_WRITE_PROP) = True
                    AccFolPPPP.Update()
                Catch ex As Exception

                End Try
                textToWrite = DateTime.Now & " : " & "" & "The security " & groupdeSec & " has been assigned to the admission folder "
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            End If

        Catch ex As Exception
            textToWrite = DateTime.Now & " : " & "" & "Error while applying the security " & groupdeSec & " Due To : " & ex.Message
            writeLogFile(textToWrite, fileNIPpath)
            textToWrite = ""
        End Try
    End Sub

    Public Sub RemoveApplySecurity(ByVal entryName As String, ByVal groupdeSec As String, ByVal visiteOrTypeOrDoctor As String)
        Dim textToWrite As String = ""
        Try
            If visiteOrTypeOrDoctor = "Visite" Then

                Try


                    Dim secureFolder As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & MedicalFileVisit & "\" & entryName)
                    Dim secureGoupV As LFGroup = db.GetTrusteeByName(groupdeSec)
                    Dim AccFol As LFAccessRights = secureFolder.Rights
                    Dim AnACE4 As LFAccessControlEntry = AccFol.RetrieveAccessControlEntry(secureGoupV, EntryAce_Scope.SCOPE_TOTAL, secureFolder)
                    AccFol.RemoveAccessControlEntry(AnACE4)
                    AccFol.Update()
                    secureFolder.Dispose()
                Catch ex As Exception

                End Try
                textToWrite = DateTime.Now & " : " & "" & "The security " & groupdeSec & " has been removed from the the visite " & CaseNumberNew
                writeLogFile(textToWrite, fileNIPpath)
                textToWrite = ""
            End If

        Catch ex As Exception
            textToWrite = DateTime.Now & " : " & "" & "Error while removing the security " & groupdeSec & " Due To : " & ex.Message
            writeLogFile(textToWrite, fileNIPpath)
            textToWrite = ""
        End Try
    End Sub

    Public Function pathReport(ByVal codeId As String) As String
        Dim pathRepport As String = ""

        Select Case codeId
            Case "0"
                pathRepport = "Visit  Laboratory Results\Laboratory Results"
            Case "19"
                pathRepport = "Visit  Laboratory Results\Laboratory Results"
            Case "1"
                pathRepport = "Visit  Laboratory Results\Pathology Results"
            Case "2"
                pathRepport = "Visit  Radiology   Results\General Radiology"
            Case "3"
                pathRepport = "Visit  Radiology   Results\mri"
            Case "4"
                pathRepport = "Visit  Radiology   Results\ct-scan"
            Case "5"
                pathRepport = "Visit  Other Exams Results\Cardiography"
            Case "6"
                pathRepport = "Visit  Radiology   Results\Echography"
            Case "7"
                pathRepport = "Visit  Radiology   Results\Echography"
            Case "17"
                pathRepport = "Visit  Radiology   Results\Echography"
            Case "21"
                pathRepport = "Visit  Radiology   Results\Echography"
            Case "8"
                pathRepport = "Visit  Radiology   Results\Nucluare Radiology"
            Case "9"
                pathRepport = "Visit  Other Exams Results\Strees Test Result"
            Case "10"
                pathRepport = "Visit  Other Exams Results\Holter Monitor"
            Case "11"
                pathRepport = "Visit  Other Exams Results\EEG Results"
            Case "12"
                pathRepport = "Visit  Other Exams Results\EMG Results"
            Case "13"
                pathRepport = "Visit  Other Exams Results\Osteodensitometry Results"
            Case "14"
                pathRepport = "Visit  Other Exams Results\Spiro Meter Results"
            Case "15"
                pathRepport = "Visit  Surgical Scan  Records\Cardio Vascular Lab"
            Case "16"
                pathRepport = "Visit  Radiology   Results\IVF"
            Case "18"
                pathRepport = "Visit Discharge Documents\Emergency"
            Case "20"
                pathRepport = "Visit  Surgical Scan  Records\Cardio Vascular Lab"
            Case "22"
                pathRepport = "Visit  Radiology   Results\VCT"
            Case "23"
                pathRepport = "Visit  Surgical Scan  Records\General Surgery"
            Case "24"
                pathRepport = "Visit  Surgical Scan  Records\General Surgery"
            Case "000"
                pathRepport = "Visit  Laboratory Results\Laboratory Results"
            Case Else
                pathRepport = ""
        End Select

        Return pathRepport
    End Function

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        'MsgBox("tick")
        Timer1.Enabled = False
        Me.btnGenerate.Enabled = False



        ' MsgBox("gg")

        Dim s As String = Date.Today
        fileNIPpath = ""
        fileNIPpath = logFilePath & s.Replace("/", "-") & ".txt"
        If Not System.IO.File.Exists(fileNIPpath) Then
            System.IO.File.Create(fileNIPpath).Dispose()
        End If
        Dim textToWrite As String = ""

        Try
            Dim dtPatient As New DataTable()
            dtPatient = getPatientViewRows()
            CreateToLaserficheNIP(dtPatient)

            Dim dtVisit As New DataTable()
            dtVisit = getVisitViewRows()
            CreateToLaserficheVisite(dtVisit)

            ImportPDFFiles()

            Try
                If conn.IsTerminated = False Then
                    conn.Terminate()
                End If
            Catch ex1 As Exception

            End Try

        Catch ex As Exception
            writeLogFile("Error in recursive function : " & ex.Message, fileNIPpath)
        End Try


        Timer1.Enabled = True
        Me.btnGenerate.Enabled = True

    End Sub

    Public Sub ImportPDFFiles()
        Dim mf As String
        Dim cn As String
        Dim cc As String
        Dim examNumber As String
        Dim extractDate As String
        Dim textToWrite As String
        Dim dirInfo As New IO.DirectoryInfo(pathPDF)
        For Each file As FileInfo In dirInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly)
            If FileIsLocked(pathPDF & file.Name) = False Then
                Try
                    mf = file.Name.Split("-")(0)
                    cn = file.Name.Split("-")(1)
                    cc = pathReport(file.Name.Split("-")(2))
                    examNumber = file.Name.Split("-")(3)

                    Dim dd As String = file.Name.Split("-")(4).Split(".")(0)
                    If dd.Length = 8 Then
                        extractDate = dd.Trim().Substring(0, 2) & "/" & dd.Trim().Substring(2, 2) & "/" & dd.Trim().Substring(4, 4)
                    Else
                        extractDate = ""
                    End If

                    returnFolderNumberName(mf)

                    Try
                        connDb()
                        ' textToWrite = DateTime.Now & " : " & "" & "testttttttt  " & pathInLF & "\" & folNumberName & "\" & mf & "\" & cn & "\" & cc
                        ' writeLogFile(textToWrite, fileNIPpath)
                        Try
                            Dim foll As New LFFolder
                            Dim ParentFol As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & mf & "\" & cn)
                            foll.Create(cc.Split("\")(0), ParentFol, False)
                            foll.Dispose()
                        Catch ex As Exception

                        End Try

                        Try
                            Dim foll As New LFFolder
                            Dim ParentFol As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & mf & "\" & cn & "\" & cc.Split("\")(0))
                            foll.Create(cc.Split("\")(1), ParentFol, False)
                            foll.Dispose()
                        Catch ex As Exception

                        End Try

                        Dim fol As LFFolder = db.GetEntryByPath(pathInLF & "\" & folNumberName & "\" & mf & "\" & cn & "\" & cc)
                        'textToWrite = DateTime.Now & " : " & "" & "testttttttt 222222222222222 " & pathInLF & "\" & folNumberName & "\" & mf & "\" & cn & "\" & cc
                        ' writeLogFile(textToWrite, fileNIPpath)
                        Dim doc As New LFDocument
                        Dim Vol As LFVolume = db.GetVolumeByName(volum)
                        doc.Create(mf & "-" & cn & "-" & examNumber, fol, Vol, True)
                        Dim DocImporter As New DocumentImporter
                        DocImporter.Document = doc
                        DocImporter.ImportElectronicFile(pathPDF & file.Name)

                        doc.Dispose()
                        '  textToWrite = DateTime.Now & " : " & "" & "testttttttt 33333333333 " & pathInLF & "\" & folNumberName & "\" & mf & "\" & cn & "\" & cc
                        ' writeLogFile(textToWrite, fileNIPpath)
                        Dim folPatient As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & mf)
                        Dim folVisit As LFFolder = db.GetEntryByPath("\" & pathInLF & "\" & folNumberName & "\" & mf & "\" & cn)

                        Dim visitTypeResultTempName As LFTemplate = db.GetTemplateByName(templatePlatTechnique)
                        Dim fields As LFFieldData = doc.FieldData
                        fields.LockObject(Lock_Type.LOCK_TYPE_WRITE)
                        fields.Template = visitTypeResultTempName

                        Dim FD As LFFieldData = folPatient.FieldData
                        Dim FDVisit As LFFieldData = folVisit.FieldData

                        fields.Field(MedicalFileField) = FD.Field(MedicalFileField)
                        fields.Field(FirstField) = FD.Field(FirstField)
                        fields.Field(FamilyField) = FD.Field(FamilyField)
                        fields.Field(FatherField) = FD.Field(FatherField)

                        fields.Field(CaseNumberField) = FDVisit.Field(CaseNumberField)
                        fields.Field(DoctorCodeField) = FDVisit.Field(DoctorCodeField)
                        fields.Field(DoctorNameField) = FDVisit.Field(DoctorNameField)
                        fields.Field(ConsultantCodeField) = FDVisit.Field(ConsultantCodeField)
                        fields.Field(ConsultantNameField) = FDVisit.Field(ConsultantNameField)
                        fields.Field(DepartmentField) = FDVisit.Field(DepartmentField)
                        fields.Field(PatientTypeField) = FDVisit.Field(PatientTypeField)
                        fields.Field(AdmissionDateField) = FDVisit.Field(AdmissionDateField)
                        fields.Field(DischargeDateField) = FDVisit.Field(DischargeDateField)

                        fields.Field("Exam Number") = examNumber
                        fields.Field("Extraction Date") = extractDate
                        fields.Field(Type1Field) = cc.Split("\")(0)
                        fields.Field(type2Field) = cc.Split("\")(1)

                        fields.Update()
                        doc.Dispose()

                        textToWrite = DateTime.Now & " : " & "" & "The report has been imported under medicale file  " & mf & " , case number " & cn & " and code center " & cc
                        writeLogFile(textToWrite, fileNIPpath)

                    Catch ex As Exception
                        textToWrite = DateTime.Now & " : " & "" & "The medicale file " & mf & " or case number " & cn & " or code center " & file.Name.Split("-")(2) & " doesn't exist"

                        'textToWrite = DateTime.Now & " : " & "" & "Error   : " & ex.Message
                        writeLogFile(textToWrite, fileNIPpath)
                    End Try


                Catch ex As Exception
                    textToWrite = DateTime.Now & " : Error in the import function for the file " & file.Name
                    writeLogFile(textToWrite, fileNIPpath)
                End Try

                Dim FileToMove As String
                Dim MoveLocation As String
                FileToMove = pathPDF & file.Name
                MoveLocation = ProgressPDF & file.Name

                If System.IO.File.Exists(FileToMove) = True Then
                    If System.IO.File.Exists(MoveLocation) = True Then
                        System.IO.File.Delete(MoveLocation)
                    End If
                    System.IO.File.Move(FileToMove, MoveLocation)
                End If

            End If
            Try
                If conn.IsTerminated = False Then
                    conn.Terminate()
                End If
            Catch ex1 As Exception

            End Try
        Next

        Try
            If conn.IsTerminated = False Then
                conn.Terminate()
            End If
        Catch ex1 As Exception

        End Try
    End Sub

    Public Function FileIsLocked(ByVal fileFullPathName As String) As Boolean
        Dim isLocked As Boolean = False
        Dim fileObj As System.IO.FileStream

        Try
            fileObj = New System.IO.FileStream( _
                                    fileFullPathName, _
                                    System.IO.FileMode.Open, _
                                    System.IO.FileAccess.ReadWrite, _
                                    System.IO.FileShare.None)
        Catch
            isLocked = True
        Finally
            If fileObj IsNot Nothing Then
                fileObj.Close()
            End If
        End Try

        Return isLocked
    End Function
End Class

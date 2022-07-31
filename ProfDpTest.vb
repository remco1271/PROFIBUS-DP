Option Strict On

Imports ProfDpClass.MProfDp                     ' MProfDp Class
Imports ProfDpClass.MProfDp.MDpResult           ' MProfDp Result codes
Imports ProfDpClass.MProfDp.MDpeRejectCodes     ' MProfDp reject codes
Imports ProfDpClass.MProfDp.MDpeFaultCodes      ' MProfDp fault codes

Public Class ProfDpTest
    Inherits System.Windows.Forms.Form

#Region " Vom Windows Form Designer generierter Code "

    Public Sub New()
        MyBase.New()

        ' Dieser Aufruf ist für den Windows Form-Designer erforderlich.
        InitializeComponent()

        ' Initialisierungen nach dem Aufruf InitializeComponent() hinzufügen

    End Sub

    ' Die Form überschreibt den Löschvorgang der Basisklasse, um Komponenten zu bereinigen.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Für Windows Form-Designer erforderlich
    Private components As System.ComponentModel.IContainer

    'HINWEIS: Die folgende Prozedur ist für den Windows Form-Designer erforderlich
    'Sie kann mit dem Windows Form-Designer modifiziert werden.
    'Verwenden Sie nicht den Code-Editor zur Bearbeitung.
    Friend WithEvents DllNameEdit As System.Windows.Forms.TextBox
    Friend WithEvents DllName As System.Windows.Forms.Label
    Friend WithEvents InitStringEdit As System.Windows.Forms.TextBox
    Friend WithEvents InitString As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CloseForm As System.Windows.Forms.Button
    Friend WithEvents DpCreate As System.Windows.Forms.Button
    Friend WithEvents DpDestroy As System.Windows.Forms.Button
    Friend WithEvents GroupProfDp As System.Windows.Forms.GroupBox
    Friend WithEvents Input1 As System.Windows.Forms.Label
    Friend WithEvents InputEdit1 As System.Windows.Forms.TextBox
    Friend WithEvents Input2 As System.Windows.Forms.Label
    Friend WithEvents InputEdit2 As System.Windows.Forms.TextBox
    Friend WithEvents Input3 As System.Windows.Forms.Label
    Friend WithEvents InputEdit3 As System.Windows.Forms.TextBox
    Friend WithEvents Input4 As System.Windows.Forms.Label
    Friend WithEvents InputEdit4 As System.Windows.Forms.TextBox
    Friend WithEvents InputEdit5 As System.Windows.Forms.TextBox
    Friend WithEvents Input5 As System.Windows.Forms.Label
    Friend WithEvents Input6 As System.Windows.Forms.Label
    Friend WithEvents InputEdit6 As System.Windows.Forms.TextBox
    Friend WithEvents LabelCommands As System.Windows.Forms.Label
    Friend WithEvents ComboCommands As System.Windows.Forms.ComboBox
    Friend WithEvents ExecFunc As System.Windows.Forms.Button
    Friend WithEvents Messages As System.Windows.Forms.ListView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ProfDpTest))
        Me.DpCreate = New System.Windows.Forms.Button
        Me.DllNameEdit = New System.Windows.Forms.TextBox
        Me.DllName = New System.Windows.Forms.Label
        Me.InitStringEdit = New System.Windows.Forms.TextBox
        Me.InitString = New System.Windows.Forms.Label
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.CloseForm = New System.Windows.Forms.Button
        Me.DpDestroy = New System.Windows.Forms.Button
        Me.GroupProfDp = New System.Windows.Forms.GroupBox
        Me.InputEdit1 = New System.Windows.Forms.TextBox
        Me.Input1 = New System.Windows.Forms.Label
        Me.InputEdit2 = New System.Windows.Forms.TextBox
        Me.Input2 = New System.Windows.Forms.Label
        Me.InputEdit3 = New System.Windows.Forms.TextBox
        Me.Input3 = New System.Windows.Forms.Label
        Me.InputEdit4 = New System.Windows.Forms.TextBox
        Me.Input4 = New System.Windows.Forms.Label
        Me.InputEdit5 = New System.Windows.Forms.TextBox
        Me.Input5 = New System.Windows.Forms.Label
        Me.Input6 = New System.Windows.Forms.Label
        Me.InputEdit6 = New System.Windows.Forms.TextBox
        Me.LabelCommands = New System.Windows.Forms.Label
        Me.ComboCommands = New System.Windows.Forms.ComboBox
        Me.ExecFunc = New System.Windows.Forms.Button
        Me.Messages = New System.Windows.Forms.ListView
        Me.GroupProfDp.SuspendLayout()
        Me.SuspendLayout()
        '
        'DpCreate
        '
        Me.DpCreate.Location = New System.Drawing.Point(24, 24)
        Me.DpCreate.Name = "DpCreate"
        Me.DpCreate.Size = New System.Drawing.Size(88, 23)
        Me.DpCreate.TabIndex = 0
        Me.DpCreate.Text = "&Create"
        '
        'DllNameEdit
        '
        Me.DllNameEdit.Location = New System.Drawing.Point(216, 24)
        Me.DllNameEdit.Name = "DllNameEdit"
        Me.DllNameEdit.Size = New System.Drawing.Size(128, 20)
        Me.DllNameEdit.TabIndex = 2
        Me.DllNameEdit.Text = ""
        '
        'DllName
        '
        Me.DllName.Location = New System.Drawing.Point(136, 24)
        Me.DllName.Name = "DllName"
        Me.DllName.Size = New System.Drawing.Size(72, 16)
        Me.DllName.TabIndex = 1
        Me.DllName.Text = "Dll Name"
        '
        'InitStringEdit
        '
        Me.InitStringEdit.Location = New System.Drawing.Point(448, 24)
        Me.InitStringEdit.Name = "InitStringEdit"
        Me.InitStringEdit.Size = New System.Drawing.Size(128, 20)
        Me.InitStringEdit.TabIndex = 4
        Me.InitStringEdit.Text = ""
        '
        'InitString
        '
        Me.InitString.Location = New System.Drawing.Point(368, 24)
        Me.InitString.Name = "InitString"
        Me.InitString.Size = New System.Drawing.Size(72, 16)
        Me.InitString.TabIndex = 3
        Me.InitString.Text = "Init String"
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Messages"
        Me.ColumnHeader1.Width = 844
        '
        'CloseForm
        '
        Me.CloseForm.Location = New System.Drawing.Point(16, 624)
        Me.CloseForm.Name = "CloseForm"
        Me.CloseForm.Size = New System.Drawing.Size(88, 23)
        Me.CloseForm.TabIndex = 40
        Me.CloseForm.Text = "&Close"
        '
        'DpDestroy
        '
        Me.DpDestroy.Location = New System.Drawing.Point(24, 56)
        Me.DpDestroy.Name = "DpDestroy"
        Me.DpDestroy.Size = New System.Drawing.Size(88, 23)
        Me.DpDestroy.TabIndex = 5
        Me.DpDestroy.Text = "&Destroy"
        '
        'GroupProfDp
        '
        Me.GroupProfDp.Controls.Add(Me.InputEdit6)
        Me.GroupProfDp.Controls.Add(Me.InputEdit1)
        Me.GroupProfDp.Controls.Add(Me.Input1)
        Me.GroupProfDp.Controls.Add(Me.InputEdit2)
        Me.GroupProfDp.Controls.Add(Me.Input2)
        Me.GroupProfDp.Controls.Add(Me.InputEdit3)
        Me.GroupProfDp.Controls.Add(Me.Input3)
        Me.GroupProfDp.Controls.Add(Me.InputEdit4)
        Me.GroupProfDp.Controls.Add(Me.Input4)
        Me.GroupProfDp.Controls.Add(Me.InputEdit5)
        Me.GroupProfDp.Controls.Add(Me.Input5)
        Me.GroupProfDp.Controls.Add(Me.LabelCommands)
        Me.GroupProfDp.Controls.Add(Me.ComboCommands)
        Me.GroupProfDp.Controls.Add(Me.ExecFunc)
        Me.GroupProfDp.Controls.Add(Me.Input6)
        Me.GroupProfDp.Location = New System.Drawing.Point(16, 96)
        Me.GroupProfDp.Name = "GroupProfDp"
        Me.GroupProfDp.Size = New System.Drawing.Size(848, 216)
        Me.GroupProfDp.TabIndex = 6
        Me.GroupProfDp.TabStop = False
        Me.GroupProfDp.Text = "ProfDp Functions"
        '
        'InputEdit1
        '
        Me.InputEdit1.Location = New System.Drawing.Point(192, 24)
        Me.InputEdit1.Name = "InputEdit1"
        Me.InputEdit1.Size = New System.Drawing.Size(640, 20)
        Me.InputEdit1.TabIndex = 1
        Me.InputEdit1.Text = ""
        '
        'Input1
        '
        Me.Input1.Location = New System.Drawing.Point(8, 24)
        Me.Input1.Name = "Input1"
        Me.Input1.Size = New System.Drawing.Size(176, 16)
        Me.Input1.TabIndex = 2
        Me.Input1.Text = "Input1"
        '
        'InputEdit2
        '
        Me.InputEdit2.Location = New System.Drawing.Point(192, 48)
        Me.InputEdit2.Name = "InputEdit2"
        Me.InputEdit2.Size = New System.Drawing.Size(640, 20)
        Me.InputEdit2.TabIndex = 3
        Me.InputEdit2.Text = ""
        '
        'Input2
        '
        Me.Input2.Location = New System.Drawing.Point(8, 48)
        Me.Input2.Name = "Input2"
        Me.Input2.Size = New System.Drawing.Size(176, 16)
        Me.Input2.TabIndex = 4
        Me.Input2.Text = "Input2"
        '
        'InputEdit3
        '
        Me.InputEdit3.Location = New System.Drawing.Point(192, 72)
        Me.InputEdit3.Name = "InputEdit3"
        Me.InputEdit3.Size = New System.Drawing.Size(640, 20)
        Me.InputEdit3.TabIndex = 5
        Me.InputEdit3.Text = ""
        '
        'Input3
        '
        Me.Input3.Location = New System.Drawing.Point(8, 72)
        Me.Input3.Name = "Input3"
        Me.Input3.Size = New System.Drawing.Size(176, 16)
        Me.Input3.TabIndex = 6
        Me.Input3.Text = "Input3"
        '
        'InputEdit4
        '
        Me.InputEdit4.Location = New System.Drawing.Point(192, 96)
        Me.InputEdit4.Name = "InputEdit4"
        Me.InputEdit4.Size = New System.Drawing.Size(640, 20)
        Me.InputEdit4.TabIndex = 7
        Me.InputEdit4.Text = ""
        '
        'Input4
        '
        Me.Input4.Location = New System.Drawing.Point(8, 96)
        Me.Input4.Name = "Input4"
        Me.Input4.Size = New System.Drawing.Size(176, 16)
        Me.Input4.TabIndex = 8
        Me.Input4.Text = "Input4"
        '
        'InputEdit5
        '
        Me.InputEdit5.Location = New System.Drawing.Point(192, 120)
        Me.InputEdit5.Name = "InputEdit5"
        Me.InputEdit5.Size = New System.Drawing.Size(640, 20)
        Me.InputEdit5.TabIndex = 9
        Me.InputEdit5.Text = ""
        '
        'Input5
        '
        Me.Input5.Location = New System.Drawing.Point(8, 120)
        Me.Input5.Name = "Input5"
        Me.Input5.Size = New System.Drawing.Size(176, 16)
        Me.Input5.TabIndex = 10
        Me.Input5.Text = "Input5"
        '
        'Input6
        '
        Me.Input6.Location = New System.Drawing.Point(8, 144)
        Me.Input6.Name = "Input6"
        Me.Input6.Size = New System.Drawing.Size(176, 16)
        Me.Input6.TabIndex = 11
        Me.Input6.Text = "Input6"
        '
        'InputEdit6
        '
        Me.InputEdit6.Location = New System.Drawing.Point(192, 144)
        Me.InputEdit6.Name = "InputEdit6"
        Me.InputEdit6.Size = New System.Drawing.Size(640, 20)
        Me.InputEdit6.TabIndex = 12
        Me.InputEdit6.Text = ""
        '
        'LabelCommands
        '
        Me.LabelCommands.Location = New System.Drawing.Point(8, 184)
        Me.LabelCommands.Name = "LabelCommands"
        Me.LabelCommands.Size = New System.Drawing.Size(176, 16)
        Me.LabelCommands.TabIndex = 20
        Me.LabelCommands.Text = "Command"
        '
        'ComboCommands
        '
        Me.ComboCommands.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboCommands.Location = New System.Drawing.Point(192, 176)
        Me.ComboCommands.Name = "ComboCommands"
        Me.ComboCommands.Size = New System.Drawing.Size(296, 21)
        Me.ComboCommands.TabIndex = 21
        '
        'ExecFunc
        '
        Me.ExecFunc.Location = New System.Drawing.Point(512, 176)
        Me.ExecFunc.Name = "ExecFunc"
        Me.ExecFunc.TabIndex = 22
        Me.ExecFunc.Text = "&Execute"
        '
        'Messages
        '
        Me.Messages.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.Messages.FullRowSelect = True
        Me.Messages.GridLines = True
        Me.Messages.Location = New System.Drawing.Point(16, 328)
        Me.Messages.Name = "Messages"
        Me.Messages.Size = New System.Drawing.Size(848, 280)
        Me.Messages.TabIndex = 30
        Me.Messages.View = System.Windows.Forms.View.Details
        '
        'ProfDpTest
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(880, 661)
        Me.Controls.Add(Me.DpCreate)
        Me.Controls.Add(Me.DllName)
        Me.Controls.Add(Me.DllNameEdit)
        Me.Controls.Add(Me.InitStringEdit)
        Me.Controls.Add(Me.InitString)
        Me.Controls.Add(Me.DpDestroy)
        Me.Controls.Add(Me.Messages)
        Me.Controls.Add(Me.CloseForm)
        Me.Controls.Add(Me.GroupProfDp)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ProfDpTest"
        Me.Text = "ProfDpTest"
        Me.GroupProfDp.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    ' Type for test function call
    Private Enum ProfDpCallType
        Call_Init       ' Init Screen
        Call_Execute    ' Execute Program
    End Enum

    Public WithEvents ProfDpDrv As ProfDpClass.MProfDp

    ' list of test function
    Private TestFuncList As New ArrayList

    ' delegate for calling EventDPCallback
    Delegate Sub EventDPDel()
    Private EventDPDelegate As New EventDPDel(AddressOf EventDPCallback)

    ' delegate for calling EventDPCycleCallback
    Delegate Sub EventCycleDPDel()
    Private EventDPCycleDelegate As New EventCycleDPDel(AddressOf EventDPCycleCallback)

    ' delegate for calling EventDPECallback
    Delegate Sub EventDPEDel()
    Private EventDPEDelegate As New EventDPEDel(AddressOf EventDPECallback)

    ' last used slave address
    Public LastSlaveAddress As Byte

    ' last used CRef
    Public LastCRef As Integer

    ' last used slot
    Public LastSlot As Byte

    ' last used index
    Public LastIndex As Byte

    ' last used length
    Public LastLength As Byte

    ' function is called if main window opens
    Private Sub ProfDpTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.InitStringEdit.Text = "COM2"
        Me.LastSlaveAddress = 3
        Me.LastCRef = 0
        Me.LastSlot = 0
        Me.LastIndex = 0
        Me.LastLength = 1

        ' disable clickable headers
        Me.Messages.HeaderStyle = ColumnHeaderStyle.Nonclickable

        ' Create list of test functions
        TestFuncList.Add(New TestFunc.TestFuncDPIdentifySlave(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPChangeAddress(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPGetCfg(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPSendPrm(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPOpenSlave(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPCloseSlave(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPReadSlaveData(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPWriteSlaveData(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPReadResult(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPSetFDLReceiveTimeout(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPCheckConvFeature(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEnableDPEvent(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEnableDPCycleEvent(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEnableDPEEvent(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPECreateConnection(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEDestroyConnection(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEReadResult(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEResetReq(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEInitReq(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEAbortReq(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEReadReq(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPEWriteReq(Me))
        TestFuncList.Add(New TestFunc.TestFuncDPETransReq(Me))

        ' fill the combo box with the test function names
        Dim myEnumerator As System.Collections.IEnumerator = _
           TestFuncList.GetEnumerator()
        While myEnumerator.MoveNext()
            Me.ComboCommands.Items.Add(CType(myEnumerator.Current, TestFunc.TestFuncBase).FuncName)
        End While

        ' select first entry
        Me.ComboCommands.SelectedIndex = 0

    End Sub

    Public Function ErrorCode2Text(ByVal ErrorCode As Integer) As String

        Select Case ErrorCode
            Case MDP_NO_ERROR           ' no error
                ErrorCode2Text = "No error"

            Case MDP_GENERIC_ERROR      ' generic error
                ErrorCode2Text = "Generic error"

            Case MDP_CALLED_TWICE       ' DP function have been called double (only with DPSetDispatchMessage())
                ErrorCode2Text = "Function called twice"

            Case MDP_OPEN_FAILED        ' device (FDL) opening failed
                ErrorCode2Text = "FDL opening failed"

            Case MDP_FDL_LOAD_FAILED    ' unable to load FDL library
                ErrorCode2Text = "Unable to call FDL library"

            Case MDP_NOT_OPEN           ' FDL not opened
                ErrorCode2Text = "FDL not opened"

            Case MDP_THREAD_ENABLED     ' error, background thread enabled: function disabled
                ErrorCode2Text = "Function disabled because thread is enabled"

            Case MDP_NO_THREAD_ENABLED  ' error, background thread is not enabled: function disabled
                ErrorCode2Text = "Function disabled because thread is not enabled"

            Case MDP_OUT_OF_MEMORY      ' out of memory
                ErrorCode2Text = "Out of memory"

            Case MDP_CONVERTER_FAILED   ' PROFIBUS converter not found, failed or bad type
                ErrorCode2Text = "PROFIBUS converter not found or bad type"

            Case MDP_ILL_ADDRESS        ' illegal DP slave address
                ErrorCode2Text = "Illegal DP slave address"

            Case MDP_SAP_OPEN_FAILED    ' PROFIBUS SAP opening failed
                ErrorCode2Text = "PROFIBUS SAP opening failed"

            Case MDP_SLAVE_OPEN         ' slave is already open
                ErrorCode2Text = "Slave is already open"

            Case MDP_SLAVE_NOT_OPEN     ' slave is not opened
                ErrorCode2Text = "Slave is not opened"

            Case MDP_SLAVE_NOT_FOUND    ' slave not found
                ErrorCode2Text = "Slave not found"

            Case MDP_ILLEGAL_CONFIG     ' illegal PROFIBUS configuration
                ErrorCode2Text = "Illegal PROFIBUS configuration"

            Case MDP_CFG_ERROR          ' configuration error on startup
                ErrorCode2Text = "Configuration error on startup"

            Case MDP_PRM_ERROR          ' parameter error on startup
                ErrorCode2Text = "Parameter error on startup"

            Case MDP_CFG_AND_PRM_ERROR  ' configuration and parameter error on startup
                ErrorCode2Text = "Configuration and parameter error on startup"

            Case MDP_BAD_LENGTH         ' data length incorrect
                ErrorCode2Text = "Data length incorrect"

            Case MDP_INCOMPLETE         ' background operation initiated
                ErrorCode2Text = "Background operation initiated"

            Case MDP_NORESULT           ' no result pending
                ErrorCode2Text = "No result pending"

            Case Else
                ErrorCode2Text = [String].Format("Unknown error code {0:X04}", CInt(ErrorCode))
        End Select

    End Function

    ' convert reject codes to text
    Public Function RejectCode2Text(ByVal RejectCode As MDpeRejectCodes) As String
        Select Case RejectCode
            Case MDPE_REJ_LE            ' max_data_len_unit overflow
                RejectCode2Text = "max_data_len_unit overflow"

            Case MDPE_REJ_PS            ' number of parallel service requests exceeded
                RejectCode2Text = "number of parallel service requests exceeded"

            Case MDPE_REJ_SE            ' MSAC1 service reject
                RejectCode2Text = "MSAC1 service reject"

            Case MDPE_REJ_ABORT         ' MSAC1 abort reject
                RejectCode2Text = "MSAC1 abort reject"

            Case MDPE_REJ_IV            ' MSAC1 invalid service reject
                RejectCode2Text = "MSAC1 invalid service reject"

            Case Else
                RejectCode2Text = [String].Format("Unknown reject code {0:X04}", CInt(RejectCode))

        End Select
    End Function

    ' convert fault codes to text
    Public Function FaultCode2Text(ByVal FaultCode As MDpeFaultCodes) As String
        Select Case FaultCode
            Case MDPE_DDLM_FDL_FAULT    ' FDL fault
                FaultCode2Text = "FDL fault"

            Case MDPE_MSAC1_FAULT       ' number of parallel service requests exceeded
                FaultCode2Text = "MSAC1 processing fault"

            Case Else
                FaultCode2Text = [String].Format("Unknown fault code {0:X04}", CInt(FaultCode))

        End Select
    End Function

    ' print Error code to the output window
    Public Sub PrintError(ByVal ErrorCode As Integer)

        If ErrorCode <> MDP_NO_ERROR Then
            ' error occured -> display it
            DisplayText("Error: " + ErrorCode2Text(ErrorCode))
        End If

    End Sub

    ' print DpData structure
    Public Sub PrintDpData(ByVal DpData As MDpData)
        Dim Output As String

        DisplayText([String].Format(" DPV1_Supported={0}", DpData.m_DPV1_Supported))
        DisplayText([String].Format(" Max_Channel_Data_Length={0}", DpData.m_Max_Channel_Data_Length))
        DisplayText([String].Format(" Extra_Alarm_SAP={0}", DpData.m_Extra_Alarm_SAP))
        DisplayText([String].Format(" C1_Response_Timeout={0}", DpData.m_C1_Response_Timeout))

        If Not (DpData.m_NormParameters Is Nothing) Then
            Output = " NormParameters="
            If DpData.m_NormParameters.Length > 0 Then
                For i As Integer = 0 To DpData.m_NormParameters.Length - 1
                    Output += DpData.m_NormParameters(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
        End If

        If Not (DpData.m_UserParameters Is Nothing) Then
            Output = " UserParameters="
            If DpData.m_UserParameters.Length > 0 Then
                For i As Integer = 0 To DpData.m_UserParameters.Length - 1
                    Output += DpData.m_UserParameters(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
        End If

        If Not (DpData.m_Configuration Is Nothing) Then
            Output = " Configuration="
            If DpData.m_Configuration.Length > 0 Then
                For i As Integer = 0 To DpData.m_Configuration.Length - 1
                    Output += DpData.m_Configuration(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
        End If

        If Not (DpData.m_OutputData Is Nothing) Then
            Output = " OutputData="
            If DpData.m_OutputData.Length > 0 Then
                For i As Integer = 0 To DpData.m_OutputData.Length - 1
                    Output += DpData.m_OutputData(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
        End If

        If Not (DpData.m_InputData Is Nothing) Then
            Output = " InputData="
            If DpData.m_InputData.Length > 0 Then
                For i As Integer = 0 To DpData.m_InputData.Length - 1
                    Output += DpData.m_InputData(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
        End If

        If Not (DpData.m_NormDiagnosis Is Nothing) Then
            Output = " NormDiagnosis="
            If DpData.m_NormDiagnosis.Length > 0 Then
                For i As Integer = 0 To DpData.m_NormDiagnosis.Length - 1
                    Output += DpData.m_NormDiagnosis(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
        End If

        If Not (DpData.m_UserDiagnosis Is Nothing) Then
            Output = " UserDiagnosis="
            If DpData.m_UserDiagnosis.Length > 0 Then
                For i As Integer = 0 To DpData.m_UserDiagnosis.Length - 1
                    Output += DpData.m_UserDiagnosis(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
        End If

        DisplayText([String].Format(" Communicate={0}", DpData.m_Communicate))
        DisplayText([String].Format(" Operate={0}", DpData.m_Operate))
        DisplayText([String].Format(" New Diag={0}", DpData.m_NewDiag))
    End Sub

    ' print PROFIBUS DP result information
    Public Sub PrintDPResult(ByVal DPResult As MDpResultBase)

        DisplayText([String].Format("DPResult ID={0:X04}, Error='{1}'", DPResult.m_ID, ErrorCode2Text(DPResult.m_ResultCode)))

        If (TypeOf DPResult Is MDpResultOpenSlave) Then
            ' result of MDPOpenSlave
            DisplayText(" DPOpenSlave")
            Exit Sub
        End If

        If (TypeOf DPResult Is MDpResultCloseSlave) Then
            ' result of MDPCloseSlave
            DisplayText(" DPCloseSlave")
            Exit Sub
        End If

        If (TypeOf DPResult Is MDpResultGetCfg) Then
            ' result of get cfg command
            Dim Output As String
            Dim Config() As Byte

            DisplayText(" DPGetCfg")

            If (DPResult.m_ResultCode <> MDP_NO_ERROR) Then
                ' error occured don't print more info
                Exit Sub
            End If

            Config = CType(DPResult, MDpResultGetCfg).m_Config
            Output = " Config="
            If Config.Length > 0 Then
                For i As Integer = 0 To Config.Length - 1
                    Output += Config(i).ToString("X02") + " "
                Next
            End If
            DisplayText(Output)
            Exit Sub
        End If

        If (TypeOf DPResult Is MDpResultSendPrm) Then
            ' result of MDPSendPrm
            DisplayText(" DPSendPrm")
            Exit Sub
        End If

        If (TypeOf DPResult Is MDpResultIdent) Then
            ' result of get ident command
            DisplayText(" DPIdentifySlave")

            If (DPResult.m_ResultCode <> MDP_NO_ERROR) Then
                ' error occured don't print more info
                Exit Sub
            End If

            DisplayText([String].Format(" IdentNumber (hex)={0:X04}", CType(DPResult, MDpResultIdent).m_IdentNumber))
            Exit Sub
        End If

        If (TypeOf DPResult Is MDpResultChangeAddress) Then
            ' result of MDPChangeAddress
            DisplayText(" DPChangeAddress")
            Exit Sub
        End If

        DisplayText(" Unknown Result Type")

    End Sub

    ' print negative confirmation
    Public Sub PrintConfirmationNeg(ByVal RespConfNegBase As MDpeRespConfNegBase)
        DisplayText([String].Format(" ErrorDecode (hex)={0:X02}", RespConfNegBase.m_ErrorDecode))
        DisplayText([String].Format(" ErrorCode1 (hex)={0:X02}", RespConfNegBase.m_ErrorCode1))
        DisplayText([String].Format(" ErrorCode2 (hex)={0:X02}", RespConfNegBase.m_ErrorCode2))
    End Sub

    ' print Error code to the output window
    Public Sub PrintDPEResult(ByVal DPEResult As MDpeRespBase)

        DisplayText([String].Format("DPEResult ID={0:X08}", DPEResult.m_CRef))
        If (TypeOf DPEResult Is MDpeRespMSAC2ConfInitPos) Then
            ' MSAC2 initiate confirmation positive
            Dim RespMSAC2ConfInitPos As MDpeRespMSAC2ConfInitPos
            Dim Output As String

            DisplayText(" MSAC2 Init Confirmation Positive")

            RespMSAC2ConfInitPos = CType(DPEResult, MDpeRespMSAC2ConfInitPos)

            DisplayText([String].Format(" FeaturesSupported 1,2 (hex)={0:X02},{1:X02}", _
                RespMSAC2ConfInitPos.m_FeaturesSupported1, RespMSAC2ConfInitPos.m_FeaturesSupported2))
            DisplayText([String].Format(" ProfileFeaturesSupported 1,2 (hex)={0:X02},{1:X02}", _
                RespMSAC2ConfInitPos.m_ProfileFeaturesSupported1, RespMSAC2ConfInitPos.m_ProfileFeaturesSupported2))
            DisplayText([String].Format(" ProfileIdentNumber (hex)={0:X04}", _
                RespMSAC2ConfInitPos.m_ProfileIdentNumber))
            DisplayText([String].Format(" MaxLenDataUnit={0}", RespMSAC2ConfInitPos.m_MaxLenDataUnit))
            If Not (RespMSAC2ConfInitPos.m_AddAddrParam Is Nothing) Then
                DisplayText(" AddAddrParam:")
                ' source address
                DisplayText([String].Format("  Source API (hex)={0:X02}", _
                    RespMSAC2ConfInitPos.m_AddAddrParam.m_SAPI))
                DisplayText([String].Format("  Source SCL (hex)={0:X02}", _
                    RespMSAC2ConfInitPos.m_AddAddrParam.m_SSCL))

                If Not (RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr Is Nothing) Then
                    Output = "  Source NetAddr="
                    If RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr.Length > 0 Then
                        For i As Integer = 0 To RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr.Length - 1
                            Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr(i).ToString("X02") + " "
                        Next
                    End If
                    DisplayText(Output)
                End If

                If Not (RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr Is Nothing) Then
                    Output = "  Source MACAddr="
                    If RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr.Length > 0 Then
                        For i As Integer = 0 To RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr.Length - 1
                            Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr(i).ToString("X02") + " "
                        Next
                    End If
                    DisplayText(Output)
                End If
                ' destination address
                DisplayText([String].Format("  Destination API (hex)={0:X02}", _
                    RespMSAC2ConfInitPos.m_AddAddrParam.m_DAPI))
                DisplayText([String].Format("  Destination SCL (hex)={0:X02}", _
                    RespMSAC2ConfInitPos.m_AddAddrParam.m_DSCL))

                If Not (RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr Is Nothing) Then
                    Output = "  Destination NetAddr="
                    If RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr.Length > 0 Then
                        For i As Integer = 0 To RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr.Length - 1
                            Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr(i).ToString("X02") + " "
                        Next
                    End If
                    DisplayText(Output)
                End If

                If Not (RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr Is Nothing) Then
                    Output = "  Destination MACAddr="
                    If RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr.Length > 0 Then
                        For i As Integer = 0 To RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr.Length - 1
                            Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr(i).ToString("X02") + " "
                        Next
                    End If
                    DisplayText(Output)
                End If
            End If

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2ConfInitNeg) Then
            ' MSAC2 initiate confirmation negative
            DisplayText(" MSAC2 Init Confirmation Negative")

            PrintConfirmationNeg(CType(DPEResult, MDpeRespConfNegBase))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSACXConfReadPos) Then
            ' MSACX read confirmation positive
            Dim RespMSAC2ConfReadPos As MDpeRespMSACXConfReadPos
            Dim Output As String

            DisplayText(" MSACX Read Confirmation Positive")

            RespMSAC2ConfReadPos = CType(DPEResult, MDpeRespMSACXConfReadPos)

            If Not (RespMSAC2ConfReadPos.m_Data Is Nothing) Then
                Output = " Data="
                If RespMSAC2ConfReadPos.m_Data.Length > 0 Then
                    For i As Integer = 0 To RespMSAC2ConfReadPos.m_Data.Length - 1
                        Output += RespMSAC2ConfReadPos.m_Data(i).ToString("X02") + " "
                    Next
                End If
                DisplayText(Output)
            End If

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSACXConfReadNeg) Then
            ' MSACX read confirmation negative
            DisplayText(" MSACX Read Confirmation Negative")

            PrintConfirmationNeg(CType(DPEResult, MDpeRespConfNegBase))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSACXConfWritePos) Then
            ' MSACX read confirmation positive
            Dim RespMSAC2ConfWritePos As MDpeRespMSACXConfWritePos

            DisplayText(" MSACX Write Confirmation Positive")

            RespMSAC2ConfWritePos = CType(DPEResult, MDpeRespMSACXConfWritePos)

            DisplayText([String].Format(" Length={0}", RespMSAC2ConfWritePos.m_Length))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSACXConfWriteNeg) Then
            ' MSACX write confirmation negative
            DisplayText(" MSACX Write Confirmation Negative")

            PrintConfirmationNeg(CType(DPEResult, MDpeRespConfNegBase))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2ConfTransPos) Then
            ' MSAC2 data transport confirmation positive
            Dim RespMSAC2ConfTransPos As MDpeRespMSAC2ConfTransPos
            Dim Output As String

            DisplayText(" MSAC2 Data Transport Confirmation Positive")

            RespMSAC2ConfTransPos = CType(DPEResult, MDpeRespMSAC2ConfTransPos)

            If Not (RespMSAC2ConfTransPos.m_Data Is Nothing) Then
                Output = " Data="
                If RespMSAC2ConfTransPos.m_Data.Length > 0 Then
                    For i As Integer = 0 To RespMSAC2ConfTransPos.m_Data.Length - 1
                        Output += RespMSAC2ConfTransPos.m_Data(i).ToString("X02") + " "
                    Next
                End If
                DisplayText(Output)
            End If

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2ConfTransNeg) Then
            ' MSAC2 data transport confirmation negative
            DisplayText(" MSAC2 Data Transport Confirmation Negative")

            PrintConfirmationNeg(CType(DPEResult, MDpeRespConfNegBase))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2ConfReset) Then
            ' MSAC2 reset confirmation
            DisplayText(" MSAC2 Reset Confirmation")

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2IndiAbort) Then
            ' MSAC2 abort indication
            Dim RespMSAC2IndiAbort As MDpeRespMSAC2IndiAbort

            DisplayText(" MSAC2 Abort Indication")

            RespMSAC2IndiAbort = CType(DPEResult, MDpeRespMSAC2IndiAbort)

            DisplayText([String].Format(" LocallyGenerated={0}", RespMSAC2IndiAbort.m_LocallyGenerated))
            DisplayText([String].Format(" Subnet (hex)={0:X02}", RespMSAC2IndiAbort.m_Subnet))
            DisplayText([String].Format(" ReasonCode (hex)={0:X02}", RespMSAC2IndiAbort.m_ReasonCode))
            DisplayText([String].Format(" AdditionalDetail (hex)={0:X04}", RespMSAC2IndiAbort.m_AdditionalDetail))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2ConfReset) Then
            ' MSAC2 reset confirmation
            DisplayText(" MSAC2 Reset Confirmation")

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2IndiClosed) Then
            ' MSAC2 closed indication
            DisplayText(" MSAC2 Closed Indication")

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2IndiReject) Then
            ' MSAC2 reject indication
            Dim RespMSAC2IndiReject As MDpeRespMSAC2IndiReject

            DisplayText(" MSAC2 Reject Indication")

            RespMSAC2IndiReject = CType(DPEResult, MDpeRespMSAC2IndiReject)

            DisplayText([String].Format(" ReasonCode={0}", RejectCode2Text(RespMSAC2IndiReject.m_ReasonCode)))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMSAC2IndiFault) Then
            ' MSAC2 fault indication
            DisplayText(" MSAC2 Fault Indication")

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMASC1IndiStarted) Then
            ' MSAC2 started indication
            DisplayText(" MSAC1 Started Indication")

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMASC1IndiStopped) Then
            ' MSAC2 stopped indication
            DisplayText(" MSAC1 Stopped Indication")

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMASC1IndiReject) Then
            ' MSAC1 reject indication
            Dim RespMASC1IndiReject As MDpeRespMASC1IndiReject

            DisplayText(" MSAC1 Reject Indication")

            RespMASC1IndiReject = CType(DPEResult, MDpeRespMASC1IndiReject)

            DisplayText([String].Format(" Status={0}", RejectCode2Text(RespMASC1IndiReject.m_Status)))

            Exit Sub
        End If

        If (TypeOf DPEResult Is MDpeRespMASC1IndiFault) Then
            ' MSAC1 fault indication
            Dim RespMASC1IndiFault As MDpeRespMASC1IndiFault

            DisplayText(" MSAC1 Fault Indication")

            RespMASC1IndiFault = CType(DPEResult, MDpeRespMASC1IndiFault)

            DisplayText([String].Format(" Status={0}", FaultCode2Text(RespMASC1IndiFault.m_Status)))

            Exit Sub
        End If

        DisplayText(" Unknown Result Type")
    End Sub

    ' print given text to the output window
    Public Sub DisplayText(ByVal Text As String)

        Dim ListViewItemNew As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(Text)

        ListViewItemNew.Tag = Text
        Me.Messages.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItemNew})
        Me.Messages.EnsureVisible(Me.Messages.Items.Count() - 1)

    End Sub

    ' call the specified test function with init or execute
    Private Sub CallSelectedFunction(ByVal Selection As String, ByVal CallType As ProfDpCallType)

        Dim TestFunc As TestFunc.TestFuncBase

        ' search the function name in the list of test functions
        Dim myEnumerator As System.Collections.IEnumerator = _
           TestFuncList.GetEnumerator()
        While myEnumerator.MoveNext()
            TestFunc = CType(myEnumerator.Current, TestFunc.TestFuncBase)
            If TestFunc.FuncName = Selection Then
                ' function found
                Select Case CallType
                    Case ProfDpCallType.Call_Init  ' init display
                        Dim FieldName As String = Nothing    ' text beside edit line
                        Dim FieldText As String = Nothing    ' text in edit field

                        ' init edit line 1
                        TestFunc.InitLine1(FieldName, FieldText)
                        If FieldName = Nothing Then
                            ' line not used -> disable it
                            Me.Input1.Text = ""
                            Me.InputEdit1.Text = ""
                            Me.InputEdit1.Enabled = False
                        Else
                            Me.Input1.Text = FieldName
                            Me.InputEdit1.Text = FieldText
                            Me.InputEdit1.Enabled = True
                        End If

                        ' init edit line 2
                        TestFunc.InitLine2(FieldName, FieldText)
                        If FieldName = Nothing Then
                            ' line not used -> disable it
                            Me.Input2.Text = ""
                            Me.InputEdit2.Text = ""
                            Me.InputEdit2.Enabled = False
                        Else
                            Me.Input2.Text = FieldName
                            Me.InputEdit2.Text = FieldText
                            Me.InputEdit2.Enabled = True
                        End If

                        ' init edit line 3
                        TestFunc.InitLine3(FieldName, FieldText)
                        If FieldName = Nothing Then
                            ' line not used -> disable it
                            Me.Input3.Text = ""
                            Me.InputEdit3.Text = ""
                            Me.InputEdit3.Enabled = False
                        Else
                            Me.Input3.Text = FieldName
                            Me.InputEdit3.Text = FieldText
                            Me.InputEdit3.Enabled = True
                        End If

                        ' init edit line 4
                        TestFunc.InitLine4(FieldName, FieldText)
                        If FieldName = Nothing Then
                            ' line not used -> disable it
                            Me.Input4.Text = ""
                            Me.InputEdit4.Text = ""
                            Me.InputEdit4.Enabled = False
                        Else
                            Me.Input4.Text = FieldName
                            Me.InputEdit4.Text = FieldText
                            Me.InputEdit4.Enabled = True
                        End If

                        ' init edit line 5
                        TestFunc.InitLine5(FieldName, FieldText)
                        If FieldName = Nothing Then
                            ' line not used -> disable it
                            Me.Input5.Text = ""
                            Me.InputEdit5.Text = ""
                            Me.InputEdit5.Enabled = False
                        Else
                            Me.Input5.Text = FieldName
                            Me.InputEdit5.Text = FieldText
                            Me.InputEdit5.Enabled = True
                        End If

                        ' init edit line 6
                        TestFunc.InitLine6(FieldName, FieldText)
                        If FieldName = Nothing Then
                            ' line not used -> disable it
                            Me.Input6.Text = ""
                            Me.InputEdit6.Text = ""
                            Me.InputEdit6.Enabled = False
                        Else
                            Me.Input6.Text = FieldName
                            Me.InputEdit6.Text = FieldText
                            Me.InputEdit6.Enabled = True
                        End If

                    Case ProfDpCallType.Call_Execute   ' execute function
                        ' check if ProfDpDrv is present
                        If ProfDpDrv Is Nothing Then
                            DisplayText("Driver not opened")
                            Exit Sub
                        End If
                        TestFunc.RunTest()          ' call test function

                    Case Else                       ' unknown call type
                        DisplayText("Unknown Call Type")

                End Select
                Exit Sub
            End If
        End While

        ' function not found
        DisplayText("Function not found")
    End Sub

    ' create profdp class
    Private Sub DpCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DpCreate.Click

        Try
            If Not (ProfDpDrv Is Nothing) Then
                DisplayText("Class is already created")
                Exit Sub
            End If
            ProfDpDrv = New ProfDpClass.MProfDp(Me.DllNameEdit.Text, Me.InitStringEdit.Text)
            ProfDpDrv.MDPEnableDPEvent(True)
            ProfDpDrv.MDPEnableDPEEvent(True)
            DisplayText("Create(): OK")
        Catch ex As CMProfDpException
            PrintError(ex.m_ErrorCode)
            Exit Sub
        End Try

    End Sub

    ' destroy profdp class
    Private Sub DpDestroy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DpDestroy.Click
        ' perform manual cleanup to reduce resouces hold in sub dll's
        If Not (ProfDpDrv Is Nothing) Then
            ProfDpDrv.Dispose()
            ProfDpDrv = Nothing
        End If
        DisplayText("Destroy(): OK")
    End Sub

    ' new test function selected
    Private Sub ComboCommands_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboCommands.SelectedIndexChanged
        ' update edit fields
        CallSelectedFunction(Me.ComboCommands.Text, ProfDpCallType.Call_Init)

    End Sub

    ' call of test function requested
    Private Sub ExecFunc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExecFunc.Click

        ' call test function
        CallSelectedFunction(Me.ComboCommands.Text, ProfDpCallType.Call_Execute)

    End Sub

    ' close button
    Private Sub CloseForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseForm.Click

        Me.Close()

    End Sub

    ' called when form should close
    Private Sub Form_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If Not (ProfDpDrv Is Nothing) Then
            ProfDpDrv.Dispose()
            ProfDpDrv = Nothing
        End If

    End Sub

    ' callback for ProfDpEventDP
    Private Sub EventDPCallback()

        Try
            Dim DPResult As MDpResultBase = Nothing

            If (ProfDpDrv Is Nothing) Then
                Exit Sub
            End If
            ' let's read the acync result
            ' normaly the enless loop is not required, because we get an event for
            ' every async result, except we have disabled events for a while ...
            While True
                ProfDpDrv.MDPReadResult(DPResult)
                PrintDPResult(DPResult)
            End While
        Catch ex As CMProfDpException
            Exit Sub
        End Try

    End Sub

    ' event: PROIFBUS DP operation finished
    ' warning: this event is generated by a different thread
    ' so we use a delegate to access the user interface
    ' Always use BeginInvoke here and not Invoke to prevent a deadlock!
    Private Sub ProfDpEventDP() Handles ProfDpDrv.MDPEventDP

        Messages.BeginInvoke(EventDPDelegate)

    End Sub

    ' callback for ProfDpEventDPCycle
    Private Sub EventDPCycleCallback()

        DisplayText("PROFIBUS DP Cycle completed!")

    End Sub

    ' event: PROFIBUS DP cycle completed
    ' warning: this event is generated by a different thread
    ' so we use a delegate to access the user interface
    ' Always use BeginInvoke here and not Invoke to prevent a deadlock!
    Private Sub ProfDpEventDPCycle() Handles ProfDpDrv.MDPEventDPCycle

        Messages.BeginInvoke(EventDPCycleDelegate)

    End Sub

    ' callback for ProfDpEventDPE
    Private Sub EventDPECallback()

        Try
            Dim DPEResult As MDpeRespBase = Nothing

            If (ProfDpDrv Is Nothing) Then
                Exit Sub
            End If

            ' let's read the DPE result
            ' normaly the enless loop is not required, because we get an event for
            ' every DPE event, except we have disabled events for a while ...
            While True
                ProfDpDrv.MDPEReadResult(DPEResult)
                PrintDPEResult(DPEResult)
            End While

        Catch ex As CMProfDpException
            Exit Sub
        End Try

    End Sub

    ' event: DPE event occured
    ' warning: this event is generated by a different thread
    ' so we use a delegate to access the user interface
    ' Always use BeginInvoke here and not Invoke to prevent a deadlock!
    Private Sub ProfDpEventDPE() Handles ProfDpDrv.MDPEventDPE

        Messages.BeginInvoke(EventDPEDelegate)

    End Sub

End Class

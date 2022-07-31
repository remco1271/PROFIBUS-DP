Option Strict On

Imports ProfDpClass.MProfDp             ' MProfDp Class
Imports ProfDpClass.MProfDp.MDpResult   ' MProfDp Result codes

Public Class TestFunc
    ' Base class for test functions
    ' this base class is never called directly
    Public Class TestFuncBase
        Public Sub New(ByRef TestClass As ProfDpTest)
            ' store pointer to main class
            DrvTestClass = TestClass
        End Sub

        ' Pointer to main class of test app
        Protected DrvTestClass As ProfDpTest
        ' Name of test function for ComboBox
        Public FuncName As String

        ' Init edit line 1
        Public Overridable Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            ' disable fields
            FieldName = Nothing
            FieldText = Nothing
        End Sub

        ' Init edit line 2
        Public Overridable Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            ' disable fields
            FieldName = Nothing
            FieldText = Nothing
        End Sub

        ' Init edit line 3
        Public Overridable Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            ' disable fields
            FieldName = Nothing
            FieldText = Nothing
        End Sub

        ' Init edit line 4
        Public Overridable Sub InitLine4(ByRef FieldName As String, ByRef FieldText As String)
            ' disable fields
            FieldName = Nothing
            FieldText = Nothing
        End Sub

        ' Init edit line 5
        Public Overridable Sub InitLine5(ByRef FieldName As String, ByRef FieldText As String)
            ' disable fields
            FieldName = Nothing
            FieldText = Nothing
        End Sub

        ' Init edit line 6
        Public Overridable Sub InitLine6(ByRef FieldName As String, ByRef FieldText As String)
            ' disable fields
            FieldName = Nothing
            FieldText = Nothing
        End Sub

        ' run the test
        Public Overridable Sub RunTest()

        End Sub
    End Class

    Public Class TestFuncDPIdentifySlave
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPIdentifySlave"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "No Retry"
            FieldText = "0"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim SlaveAddress As Byte
                Dim NoRetry As Boolean
                Dim ID As Short

                Try
                    SlaveAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                    NoRetry = CBool(Val(DrvTestClass.InputEdit2.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPIdentifySlave(SlaveAddress, NoRetry, ID)

                DrvTestClass.LastSlaveAddress = SlaveAddress

                DrvTestClass.DisplayText([String].Format("DPIdentifySlave({0}, {1}): ID={2:X04}", SlaveAddress, NoRetry, ID))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPChangeAddress
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPChangeAddress"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Old Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            Dim NewAddress As Integer

            FieldName = "New Address"

            NewAddress = DrvTestClass.LastSlaveAddress + 1
            If (NewAddress > 125) Then
                NewAddress = 0
            End If

            FieldText = [String].Format("{0}", NewAddress)
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim OldAddress As Byte
                Dim NewAddress As Byte
                Dim ID As Short

                Try
                    OldAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                    NewAddress = CByte(Val(DrvTestClass.InputEdit2.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPChangeAddress(OldAddress, NewAddress, ID)

                DrvTestClass.LastSlaveAddress = NewAddress

                DrvTestClass.DisplayText([String].Format("DPChangeAddress({0}, {1}): ID={2:X04}", OldAddress, NewAddress, ID))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPGetCfg
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPGetCfg"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim SlaveAddress As Byte
                Dim ID As Short

                Try
                    SlaveAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPGetCfg(SlaveAddress, ID)

                DrvTestClass.LastSlaveAddress = SlaveAddress

                DrvTestClass.DisplayText([String].Format("DPGetCfg({0}): ID={1:X04}", SlaveAddress, ID))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPSendPrm
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPSendPrm"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Norm Parameters"
            FieldText = "88 10 10 37 00 00 00"
        End Sub

        Public Overrides Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "User Parameters"
            FieldText = ""
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim SlaveAddress As Byte
                Dim NormParameters() As Byte = Nothing
                Dim UserParameters() As Byte = Nothing
                Dim ID As Short

                Try
                    SlaveAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                Try
                    GetByteArray(DrvTestClass.InputEdit2, 0, NormParameters)
                    GetByteArray(DrvTestClass.InputEdit3, 0, UserParameters)
                Catch ex As System.Exception
                    DrvTestClass.DisplayText(ex.Message)
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPSendPrm(SlaveAddress, NormParameters, UserParameters, ID)

                DrvTestClass.LastSlaveAddress = SlaveAddress

                DrvTestClass.DisplayText([String].Format("DPSendPrm({0}): ID={1:X04}", SlaveAddress, ID))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPOpenSlave
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPOpenSlave"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "C1 Timeout [10ms]"
            FieldText = "100"
        End Sub

        Public Overrides Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "User Parameters"
            FieldText = ""
        End Sub

        Public Overrides Sub InitLine4(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Configuration (empty=readback)"
            FieldText = ""
        End Sub

        Public Overrides Sub InitLine5(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Open Timeout [s] (0=don't wait)"
            FieldText = "0"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim SlaveAddress As Byte
                Dim OpenTimeout As Integer
                Dim DpData As MDpData
                Dim ID As Short

                DpData = New MDpData
                DpData.m_DPV1_Supported = True
                DpData.m_Max_Channel_Data_Length = 244
                DpData.m_Extra_Alarm_SAP = False

                DpData.m_NormParameters = New Byte(6) {}
                DpData.m_NormParameters(0) = &H88     ' watchdog
                DpData.m_NormParameters(1) = &H10
                DpData.m_NormParameters(2) = &H10
                DpData.m_NormParameters(3) = &H37
                DpData.m_NormParameters(4) = &H0
                DpData.m_NormParameters(5) = &H0      ' ident read is automatic
                DpData.m_NormParameters(6) = &H0

                Try
                    SlaveAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                    DpData.m_C1_Response_Timeout = CInt(Val(DrvTestClass.InputEdit2.Text()))
                    OpenTimeout = CInt(Val(DrvTestClass.InputEdit5.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                Try
                    GetByteArray(DrvTestClass.InputEdit3, 0, DpData.m_UserParameters)
                    GetByteArray(DrvTestClass.InputEdit4, 0, DpData.m_Configuration)
                Catch ex As System.Exception
                    DrvTestClass.DisplayText(ex.Message)
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPOpenSlave(SlaveAddress, DpData, OpenTimeout, ID)

                DrvTestClass.LastSlaveAddress = SlaveAddress
                ' CRef is identical with slave address for DPV1 connections
                DrvTestClass.LastCRef = SlaveAddress

                DrvTestClass.DisplayText([String].Format("DPOpenSlave({0}): ID={1:X04}", SlaveAddress, ID))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPCloseSlave
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPCloseSlave"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim SlaveAddress As Byte
                Dim ID As Short

                Try
                    SlaveAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPCloseSlave(SlaveAddress, ID)

                DrvTestClass.LastSlaveAddress = SlaveAddress

                DrvTestClass.DisplayText([String].Format("DPCloseSlave({0}): ID={1:X04}", SlaveAddress, ID))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPReadSlaveData
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPReadSlaveData"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Reset Diag"
            FieldText = "1"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim SlaveAddress As Byte
                Dim ResetDiag As Boolean
                Dim DpData As MDpData = Nothing

                Try
                    SlaveAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                    ResetDiag = CBool(Val(DrvTestClass.InputEdit2.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPReadSlaveData(SlaveAddress, ResetDiag, DpData)

                DrvTestClass.LastSlaveAddress = SlaveAddress

                DrvTestClass.DisplayText([String].Format("DPReadSlaveData({0}): OK", SlaveAddress))
                DrvTestClass.PrintDpData(DpData)

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPWriteSlaveData
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPWriteSlaveData"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            Try
                Dim DpData As MDpData = Nothing

                FieldName = "Output Data"
                FieldText = ""

                ' we try to preset with current output data
                If Not (DrvTestClass.ProfDpDrv Is Nothing) Then
                    DrvTestClass.ProfDpDrv.MDPReadSlaveData(DrvTestClass.LastSlaveAddress, False, DpData)

                    If DpData.m_OutputData.Length > 0 Then
                        For i As Integer = 0 To DpData.m_OutputData.Length - 1
                            FieldText += DpData.m_OutputData(i).ToString("X02") + " "
                        Next
                    End If
                End If

            Catch ex As CMProfDpException
                Exit Sub
            End Try
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim SlaveAddress As Byte
                Dim OutputData() As Byte = Nothing

                Try
                    SlaveAddress = CByte(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                Try
                    GetByteArray(DrvTestClass.InputEdit2, 0, OutputData)
                Catch ex As System.Exception
                    DrvTestClass.DisplayText(ex.Message)
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPWriteSlaveData(SlaveAddress, OutputData)

                DrvTestClass.LastSlaveAddress = SlaveAddress

                DrvTestClass.DisplayText([String].Format("DPWriteSlaveData({0}): OK", SlaveAddress))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPReadResult
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPReadResult"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPResult As MDpResultBase = Nothing

                DrvTestClass.ProfDpDrv.MDPReadResult(DPResult)

                DrvTestClass.DisplayText([String].Format("DPReadResult(): OK"))
                DrvTestClass.PrintDPResult(DPResult)

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPSetFDLReceiveTimeout
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPSetFDLReceiveTimeout"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Timeout [ms]"
            FieldText = "1000"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim Timeout As Integer

                Try
                    Timeout = CInt(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPSetFDLReceiveTimeout(Timeout)

                DrvTestClass.DisplayText([String].Format("DPSetFDLReceiveTimeout({0}): OK", Timeout))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPCheckConvFeature
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPCheckConvFeature"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Feature (hex)"
            FieldText = "02"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim Feature As Byte

                Try
                    Feature = CByte(Val("&H" + DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPCheckConvFeature(Feature)

                DrvTestClass.DisplayText([String].Format("DPCheckConvFeature({0:X02}): OK", Feature))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEnableDPEvent
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPEnableDPEvent"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Enable"
            FieldText = "1"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim Enable As Boolean

                Try
                    Enable = CBool(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPEnableDPEvent(Enable)

                DrvTestClass.DisplayText([String].Format("DPEnableAcycEvent({0}): OK", Enable))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEnableDPCycleEvent
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPEnableDPCycleEvent"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Enable"
            FieldText = "1"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim Enable As Boolean

                Try
                    Enable = CBool(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPEnableDPCycleEvent(Enable)

                DrvTestClass.DisplayText([String].Format("DPEnableDpCycleEvent({0}): OK", Enable))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEnableDPEEvent
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPEnableDPEEvent"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Enable"
            FieldText = "1"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim Enable As Boolean

                Try
                    Enable = CBool(Val(DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPEnableDPEEvent(Enable)

                DrvTestClass.DisplayText([String].Format("DPEnableDPEEvent({0}): OK", Enable))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPECreateConnection
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPECreateConnection"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim CRef As Integer

                DrvTestClass.ProfDpDrv.MDPECreateConnection(CRef)

                DrvTestClass.DisplayText([String].Format("DPECreateConnection(): CRef={0:X08}", CRef))
                DrvTestClass.LastCRef = CRef

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEDestroyConnection
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPEDestroyConnection"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "CRef (hex)"
            FieldText = [String].Format("{0:X08}", DrvTestClass.LastCRef)
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim CRef As Integer

                Try
                    CRef = CInt(Val("&H" + DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPEDestroyConnection(CRef)

                DrvTestClass.DisplayText([String].Format("DPEDestroyConnection({0:X08}): OK", CRef))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEReadResult
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPEReadResult"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPEResult As MDpeRespBase = Nothing

                DrvTestClass.ProfDpDrv.MDPEReadResult(DPEResult)

                DrvTestClass.DisplayText([String].Format("DPEReadResult(): OK"))
                DrvTestClass.PrintDPEResult(DPEResult)

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEResetReq
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPE Reset Request"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "CRef (hex)"
            FieldText = [String].Format("{0:X08}", DrvTestClass.LastCRef)
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPEReqReset As MDpeReqMASC2Reset

                DPEReqReset = New MDpeReqMASC2Reset
                Try
                    DPEReqReset.m_CRef = CInt(Val("&H" + DrvTestClass.InputEdit1.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqReset)

                DrvTestClass.DisplayText([String].Format("DPEResetRequest({0:X08}): OK", DPEReqReset.m_CRef))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEInitReq
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPE Init Request"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "CRef (hex)"
            FieldText = [String].Format("{0:X08}", DrvTestClass.LastCRef)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slave Address"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlaveAddress)
        End Sub

        Public Overrides Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Send Timeout [10ms]"
            FieldText = "100"
        End Sub

        Public Overrides Sub InitLine4(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Features Supported 1 2"
            FieldText = "01 00"
        End Sub

        Public Overrides Sub InitLine5(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Prof. Features Supported 1 2"
            FieldText = "00 00"
        End Sub

        Public Overrides Sub InitLine6(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Prof. Ident Number (hex)"
            FieldText = "0000"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPEReqInit As MDpeReqMASC2Init
                Dim FeatSupp As Byte() = Nothing
                Dim ProfFeatSupp As Byte() = Nothing

                DPEReqInit = New MDpeReqMASC2Init

                Try
                    DPEReqInit.m_CRef = CInt(Val("&H" + DrvTestClass.InputEdit1.Text()))
                    DPEReqInit.m_RemAdd = CByte(Val(DrvTestClass.InputEdit2.Text()))
                    DPEReqInit.m_SendTimeout = CInt(Val(DrvTestClass.InputEdit3.Text()))
                    DPEReqInit.m_ProfileIdentNumber = CInt(Val(DrvTestClass.InputEdit6.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                Try
                    GetByteArray(DrvTestClass.InputEdit4, 2, FeatSupp)
                    GetByteArray(DrvTestClass.InputEdit5, 2, ProfFeatSupp)
                Catch ex As System.Exception
                    DrvTestClass.DisplayText(ex.Message)
                    Exit Sub
                End Try

                DPEReqInit.m_FeaturesSupported1 = FeatSupp(0)
                DPEReqInit.m_FeaturesSupported2 = FeatSupp(1)
                DPEReqInit.m_ProfileFeaturesSupported1 = ProfFeatSupp(0)
                DPEReqInit.m_ProfileFeaturesSupported2 = ProfFeatSupp(1)

                DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqInit)

                DrvTestClass.DisplayText([String].Format("DPEInitReq({0:X08}): OK", DPEReqInit.m_CRef))

                DrvTestClass.LastSlaveAddress = DPEReqInit.m_RemAdd

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEAbortReq
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPE Abort Request"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "CRef (hex)"
            FieldText = [String].Format("{0:X08}", DrvTestClass.LastCRef)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Subnet (hex)"
            FieldText = "00"
        End Sub

        Public Overrides Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Reason (hex)"
            FieldText = "00"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPEReqAbort As MDpeReqMASC2Abort

                DPEReqAbort = New MDpeReqMASC2Abort
                Try
                    DPEReqAbort.m_CRef = CInt(Val("&H" + DrvTestClass.InputEdit1.Text()))
                    DPEReqAbort.m_Subnet = CByte(Val("&H" + DrvTestClass.InputEdit2.Text()))
                    DPEReqAbort.m_Reason = CByte(Val("&H" + DrvTestClass.InputEdit3.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqAbort)

                DrvTestClass.DisplayText([String].Format("DPEAbortRequest({0:X08}): OK", DPEReqAbort.m_CRef))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEReadReq
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPE Read Request"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "CRef (hex)"
            FieldText = [String].Format("{0:X08}", DrvTestClass.LastCRef)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slot"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlot)
        End Sub

        Public Overrides Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Index"
            FieldText = [String].Format("{0}", DrvTestClass.LastIndex)
        End Sub

        Public Overrides Sub InitLine4(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Length"
            FieldText = [String].Format("{0}", DrvTestClass.LastLength)
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPEReqRead As MDpeReqMASCXRead

                DPEReqRead = New MDpeReqMASCXRead
                Try
                    DPEReqRead.m_CRef = CInt(Val("&H" + DrvTestClass.InputEdit1.Text()))
                    DPEReqRead.m_SlotNum = CByte(Val(DrvTestClass.InputEdit2.Text()))
                    DPEReqRead.m_Index = CByte(Val(DrvTestClass.InputEdit3.Text()))
                    DPEReqRead.m_Length = CByte(Val(DrvTestClass.InputEdit4.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqRead)

                DrvTestClass.LastCRef = DPEReqRead.m_CRef
                DrvTestClass.LastSlot = DPEReqRead.m_SlotNum
                DrvTestClass.LastIndex = DPEReqRead.m_Index
                DrvTestClass.LastLength = DPEReqRead.m_Length

                DrvTestClass.DisplayText([String].Format("DPEReqRead({0:X08}): OK", DPEReqRead.m_CRef))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPEWriteReq
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPE Write Request"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "CRef (hex)"
            FieldText = [String].Format("{0:X08}", DrvTestClass.LastCRef)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slot"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlot)
        End Sub

        Public Overrides Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Index"
            FieldText = [String].Format("{0}", DrvTestClass.LastIndex)
        End Sub

        Public Overrides Sub InitLine4(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Data (hex)"
            FieldText = "00"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPEReqWrite As MDpeReqMASCXWrite

                DPEReqWrite = New MDpeReqMASCXWrite
                Try
                    DPEReqWrite.m_CRef = CInt(Val("&H" + DrvTestClass.InputEdit1.Text()))
                    DPEReqWrite.m_SlotNum = CByte(Val(DrvTestClass.InputEdit2.Text()))
                    DPEReqWrite.m_Index = CByte(Val(DrvTestClass.InputEdit3.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                Try
                    GetByteArray(DrvTestClass.InputEdit4, 0, DPEReqWrite.m_PduData)
                Catch ex As System.Exception
                    DrvTestClass.DisplayText(ex.Message)
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqWrite)

                DrvTestClass.LastCRef = DPEReqWrite.m_CRef
                DrvTestClass.LastSlot = DPEReqWrite.m_SlotNum
                DrvTestClass.LastIndex = DPEReqWrite.m_Index

                DrvTestClass.DisplayText([String].Format("DPEReqWrite({0:X08}): OK", DPEReqWrite.m_CRef))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    Public Class TestFuncDPETransReq
        Inherits TestFuncBase

        Public Sub New(ByRef TestClass As ProfDpTest)
            MyBase.New(TestClass)
            FuncName = "DPE Transport Request"
        End Sub

        Public Overrides Sub InitLine1(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "CRef (hex)"
            FieldText = [String].Format("{0:X08}", DrvTestClass.LastCRef)
        End Sub

        Public Overrides Sub InitLine2(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Slot"
            FieldText = [String].Format("{0}", DrvTestClass.LastSlot)
        End Sub

        Public Overrides Sub InitLine3(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Index"
            FieldText = [String].Format("{0}", DrvTestClass.LastIndex)
        End Sub

        Public Overrides Sub InitLine4(ByRef FieldName As String, ByRef FieldText As String)
            FieldName = "Data (hex)"
            FieldText = "00"
        End Sub

        Public Overrides Sub RunTest()
            Try
                Dim DPEReqTransport As MDpeReqMASC2Transport

                DPEReqTransport = New MDpeReqMASC2Transport
                Try
                    DPEReqTransport.m_CRef = CInt(Val("&H" + DrvTestClass.InputEdit1.Text()))
                    DPEReqTransport.m_SlotNum = CByte(Val(DrvTestClass.InputEdit2.Text()))
                    DPEReqTransport.m_Index = CByte(Val(DrvTestClass.InputEdit3.Text()))
                Catch
                    DrvTestClass.DisplayText("Illegal Value")
                    Exit Sub
                End Try

                Try
                    GetByteArray(DrvTestClass.InputEdit4, 0, DPEReqTransport.m_PduData)
                Catch ex As System.Exception
                    DrvTestClass.DisplayText(ex.Message)
                    Exit Sub
                End Try

                DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqTransport)

                DrvTestClass.LastCRef = DPEReqTransport.m_CRef
                DrvTestClass.LastSlot = DPEReqTransport.m_SlotNum
                DrvTestClass.LastIndex = DPEReqTransport.m_Index

                DrvTestClass.DisplayText([String].Format("DPEReqTrans({0:X08}): OK", DPEReqTransport.m_CRef))

            Catch ex As CMProfDpException
                DrvTestClass.PrintError(ex.m_ErrorCode)
                Exit Sub
            End Try

        End Sub
    End Class

    ' get array of bytes from the given edit form
    ' if length is 0 no length check is performed
    Private Shared Sub GetByteArray(ByVal TextBox As System.Windows.Forms.TextBox, ByRef Length As Integer, ByRef ByteArray As Byte())

        Dim i As Integer
        Dim delimStr As String = " ,"
        Dim delimiter As Char() = delimStr.ToCharArray()
        Dim SubStrings() As String
        Dim SubStringCount As Integer

        ' generate list of delimiters
        SubStrings = TextBox.Text.Split(delimiter)

        ' count nummber of sub strings
        SubStringCount = 0
        For Each s As String In SubStrings
            If s.Length > 0 Then
                ' non empty string
                SubStringCount += 1
            End If
        Next

        If ((Length > 0) And (SubStringCount <> Length)) Then
            ' length specified and incorrect
            Throw New System.Exception("Invalid Parameter Length")
            Exit Sub
        End If
        ' allocate memory of calculated size (also zero size)
        ByteArray = CType(Array.CreateInstance(GetType(Byte), SubStringCount), Byte())

        ' convert sub strings
        i = 0
        For Each s As String In SubStrings
            If s.Length > 0 Then
                ' non empty string
                Try
                    ByteArray(i) = CByte(Val("&H" + s))
                Catch
                    ' conversion not possible
                    Throw New System.Exception("Illegal Value")
                    Exit Sub
                End Try
                i += 1
            End If
        Next

    End Sub

End Class

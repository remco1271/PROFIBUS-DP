using System;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

using static ProfDpClass.MProfDp;             // MProfDp Class

namespace ProfDpTest
{

    public class TestFunc
    {
        // Base class for test functions
        // this base class is never called directly
        public class TestFuncBase
        {
            public TestFuncBase(ref ProfDpTest TestClass)
            {
                // store pointer to main class
                DrvTestClass = TestClass;
            }

            // Pointer to main class of test app
            protected ProfDpTest DrvTestClass;
            // Name of test function for ComboBox
            public string FuncName;

            // Init edit line 1
            public virtual void InitLine1(ref string FieldName, ref string FieldText)
            {
                // disable fields
                FieldName = null;
                FieldText = null;
            }

            // Init edit line 2
            public virtual void InitLine2(ref string FieldName, ref string FieldText)
            {
                // disable fields
                FieldName = null;
                FieldText = null;
            }

            // Init edit line 3
            public virtual void InitLine3(ref string FieldName, ref string FieldText)
            {
                // disable fields
                FieldName = null;
                FieldText = null;
            }

            // Init edit line 4
            public virtual void InitLine4(ref string FieldName, ref string FieldText)
            {
                // disable fields
                FieldName = null;
                FieldText = null;
            }

            // Init edit line 5
            public virtual void InitLine5(ref string FieldName, ref string FieldText)
            {
                // disable fields
                FieldName = null;
                FieldText = null;
            }

            // Init edit line 6
            public virtual void InitLine6(ref string FieldName, ref string FieldText)
            {
                // disable fields
                FieldName = null;
                FieldText = null;
            }

            // run the test
            public virtual void RunTest()
            {

            }
        }

        public class TestFuncDPIdentifySlave : TestFuncBase
        {

            public TestFuncDPIdentifySlave(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPIdentifySlave";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "No Retry";
                FieldText = "0";
            }

            public override void RunTest()
            {
                try
                {
                    byte SlaveAddress;
                    bool NoRetry;
                    short ID;

                    try
                    {
                        SlaveAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                        NoRetry = Conversions.ToBoolean(Conversion.Val(DrvTestClass.InputEdit2.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPIdentifySlave(SlaveAddress, NoRetry, out ID);

                    DrvTestClass.LastSlaveAddress = SlaveAddress;

                    DrvTestClass.DisplayText(string.Format("DPIdentifySlave({0}, {1}): ID={2:X04}", SlaveAddress, NoRetry, ID));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPChangeAddress : TestFuncBase
        {

            public TestFuncDPChangeAddress(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPChangeAddress";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Old Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                int NewAddress;

                FieldName = "New Address";

                NewAddress = DrvTestClass.LastSlaveAddress + 1;
                if (NewAddress > 125)
                {
                    NewAddress = 0;
                }

                FieldText = string.Format("{0}", NewAddress);
            }

            public override void RunTest()
            {
                try
                {
                    byte OldAddress;
                    byte NewAddress;
                    short ID;

                    try
                    {
                        OldAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                        NewAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit2.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPChangeAddress(OldAddress, NewAddress, out ID);

                    DrvTestClass.LastSlaveAddress = NewAddress;

                    DrvTestClass.DisplayText(string.Format("DPChangeAddress({0}, {1}): ID={2:X04}", OldAddress, NewAddress, ID));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPGetCfg : TestFuncBase
        {

            public TestFuncDPGetCfg(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPGetCfg";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void RunTest()
            {
                try
                {
                    byte SlaveAddress;
                    short ID;

                    try
                    {
                        SlaveAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPGetCfg(SlaveAddress, out ID);

                    DrvTestClass.LastSlaveAddress = SlaveAddress;

                    DrvTestClass.DisplayText(string.Format("DPGetCfg({0}): ID={1:X04}", SlaveAddress, ID));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPSendPrm : TestFuncBase
        {

            public TestFuncDPSendPrm(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPSendPrm";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "Norm Parameters";
                FieldText = "88 10 10 37 00 00 00";
            }

            public override void InitLine3(ref string FieldName, ref string FieldText)
            {
                FieldName = "User Parameters";
                FieldText = "";
            }

            public override void RunTest()
            {
                try
                {
                    byte SlaveAddress;
                    byte[] NormParameters = null;
                    byte[] UserParameters = null;
                    short ID;

                    try
                    {
                        SlaveAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    try
                    {
                        int argLength = 0;
                        GetByteArray(DrvTestClass.InputEdit2, ref argLength, ref NormParameters);
                        int argLength1 = 0;
                        GetByteArray(DrvTestClass.InputEdit3, ref argLength1, ref UserParameters);
                    }
                    catch (Exception ex)
                    {
                        DrvTestClass.DisplayText(ex.Message);
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPSendPrm(SlaveAddress, NormParameters, UserParameters, out ID);

                    DrvTestClass.LastSlaveAddress = SlaveAddress;

                    DrvTestClass.DisplayText(string.Format("DPSendPrm({0}): ID={1:X04}", SlaveAddress, ID));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPOpenSlave : TestFuncBase
        {

            public TestFuncDPOpenSlave(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPOpenSlave";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "C1 Timeout [10ms]";
                FieldText = "100";
            }

            public override void InitLine3(ref string FieldName, ref string FieldText)
            {
                FieldName = "User Parameters";
                FieldText = "";
            }

            public override void InitLine4(ref string FieldName, ref string FieldText)
            {
                FieldName = "Configuration (empty=readback)";
                FieldText = "";
            }

            public override void InitLine5(ref string FieldName, ref string FieldText)
            {
                FieldName = "Open Timeout [s] (0=don't wait)";
                FieldText = "0";
            }

            public override void RunTest()
            {
                try
                {
                    byte SlaveAddress;
                    int OpenTimeout;
                    MDpData DpData;
                    short ID;

                    DpData = new MDpData();
                    DpData.m_DPV1_Supported = true;
                    DpData.m_Max_Channel_Data_Length = 244;
                    DpData.m_Extra_Alarm_SAP = false;

                    DpData.m_NormParameters = new byte[7];
                    DpData.m_NormParameters[0] = 0x88;     // watchdog
                    DpData.m_NormParameters[1] = 0x10;
                    DpData.m_NormParameters[2] = 0x10;
                    DpData.m_NormParameters[3] = 0x37;
                    DpData.m_NormParameters[4] = 0x0;
                    DpData.m_NormParameters[5] = 0x0;      // ident read is automatic
                    DpData.m_NormParameters[6] = 0x0;

                    try
                    {
                        SlaveAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                        DpData.m_C1_Response_Timeout = (int)Math.Round(Conversion.Val(DrvTestClass.InputEdit2.Text));
                        OpenTimeout = (int)Math.Round(Conversion.Val(DrvTestClass.InputEdit5.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    try
                    {
                        int argLength = 0;
                        GetByteArray(DrvTestClass.InputEdit3, ref argLength, ref DpData.m_UserParameters);
                        int argLength1 = 0;
                        GetByteArray(DrvTestClass.InputEdit4, ref argLength1, ref DpData.m_Configuration);
                    }
                    catch (Exception ex)
                    {
                        DrvTestClass.DisplayText(ex.Message);
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPOpenSlave(SlaveAddress, DpData, OpenTimeout, out ID);

                    DrvTestClass.LastSlaveAddress = SlaveAddress;
                    // CRef is identical with slave address for DPV1 connections
                    DrvTestClass.LastCRef = SlaveAddress;

                    DrvTestClass.DisplayText(string.Format("DPOpenSlave({0}): ID={1:X04}", SlaveAddress, ID));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPCloseSlave : TestFuncBase
        {

            public TestFuncDPCloseSlave(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPCloseSlave";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void RunTest()
            {
                try
                {
                    byte SlaveAddress;
                    short ID;

                    try
                    {
                        SlaveAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPCloseSlave(SlaveAddress, out ID);

                    DrvTestClass.LastSlaveAddress = SlaveAddress;

                    DrvTestClass.DisplayText(string.Format("DPCloseSlave({0}): ID={1:X04}", SlaveAddress, ID));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPReadSlaveData : TestFuncBase
        {

            public TestFuncDPReadSlaveData(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPReadSlaveData";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "Reset Diag";
                FieldText = "1";
            }

            public override void RunTest()
            {
                try
                {
                    byte SlaveAddress;
                    bool ResetDiag;
                    MDpData DpData = null;

                    try
                    {
                        SlaveAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                        ResetDiag = Conversions.ToBoolean(Conversion.Val(DrvTestClass.InputEdit2.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPReadSlaveData(SlaveAddress, ResetDiag, out DpData);

                    DrvTestClass.LastSlaveAddress = SlaveAddress;

                    DrvTestClass.DisplayText(string.Format("DPReadSlaveData({0}): OK", SlaveAddress));
                    DrvTestClass.PrintDpData(DpData);
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPWriteSlaveData : TestFuncBase
        {

            public TestFuncDPWriteSlaveData(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPWriteSlaveData";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                try
                {
                    MDpData DpData = null;

                    FieldName = "Output Data";
                    FieldText = "";

                    // we try to preset with current output data
                    if (DrvTestClass.ProfDpDrv is not null)
                    {
                        DrvTestClass.ProfDpDrv.MDPReadSlaveData(DrvTestClass.LastSlaveAddress, false, out DpData);

                        if (DpData.m_OutputData.Length > 0)
                        {
                            for (int i = 0, loopTo = DpData.m_OutputData.Length - 1; i <= loopTo; i++)
                                FieldText += DpData.m_OutputData[i].ToString("X02") + " ";
                        }
                    }
                }

                catch (CMProfDpException ex)
                {
                    return;
                }
            }

            public override void RunTest()
            {
                try
                {
                    byte SlaveAddress;
                    byte[] OutputData = null;

                    try
                    {
                        SlaveAddress = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    try
                    {
                        int argLength = 0;
                        GetByteArray(DrvTestClass.InputEdit2, ref argLength, ref OutputData);
                    }
                    catch (Exception ex)
                    {
                        DrvTestClass.DisplayText(ex.Message);
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPWriteSlaveData(SlaveAddress, OutputData);

                    DrvTestClass.LastSlaveAddress = SlaveAddress;

                    DrvTestClass.DisplayText(string.Format("DPWriteSlaveData({0}): OK", SlaveAddress));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPReadResult : TestFuncBase
        {

            public TestFuncDPReadResult(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPReadResult";
            }

            public override void RunTest()
            {
                try
                {
                    MDpResultBase DPResult = null;

                    DrvTestClass.ProfDpDrv.MDPReadResult(out DPResult);

                    DrvTestClass.DisplayText(string.Format("DPReadResult(): OK"));
                    DrvTestClass.PrintDPResult(DPResult);
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPSetFDLReceiveTimeout : TestFuncBase
        {

            public TestFuncDPSetFDLReceiveTimeout(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPSetFDLReceiveTimeout";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Timeout [ms]";
                FieldText = "1000";
            }

            public override void RunTest()
            {
                try
                {
                    int Timeout;

                    try
                    {
                        Timeout = (int)Math.Round(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPSetFDLReceiveTimeout(Timeout);

                    DrvTestClass.DisplayText(string.Format("DPSetFDLReceiveTimeout({0}): OK", Timeout));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPCheckConvFeature : TestFuncBase
        {

            public TestFuncDPCheckConvFeature(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPCheckConvFeature";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Feature (hex)";
                FieldText = "02";
            }

            public override void RunTest()
            {
                try
                {
                    byte Feature;

                    try
                    {
                        Feature = (byte)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPCheckConvFeature(Feature);

                    DrvTestClass.DisplayText(string.Format("DPCheckConvFeature({0:X02}): OK", Feature));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEnableDPEvent : TestFuncBase
        {

            public TestFuncDPEnableDPEvent(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPEnableDPEvent";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Enable";
                FieldText = "1";
            }

            public override void RunTest()
            {
                try
                {
                    bool Enable;

                    try
                    {
                        Enable = Conversions.ToBoolean(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPEnableDPEvent(Enable);

                    DrvTestClass.DisplayText(string.Format("DPEnableAcycEvent({0}): OK", Enable));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEnableDPCycleEvent : TestFuncBase
        {

            public TestFuncDPEnableDPCycleEvent(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPEnableDPCycleEvent";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Enable";
                FieldText = "1";
            }

            public override void RunTest()
            {
                try
                {
                    bool Enable;

                    try
                    {
                        Enable = Conversions.ToBoolean(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPEnableDPCycleEvent(Enable);

                    DrvTestClass.DisplayText(string.Format("DPEnableDpCycleEvent({0}): OK", Enable));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEnableDPEEvent : TestFuncBase
        {

            public TestFuncDPEnableDPEEvent(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPEnableDPEEvent";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "Enable";
                FieldText = "1";
            }

            public override void RunTest()
            {
                try
                {
                    bool Enable;

                    try
                    {
                        Enable = Conversions.ToBoolean(Conversion.Val(DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPEnableDPEEvent(Enable);

                    DrvTestClass.DisplayText(string.Format("DPEnableDPEEvent({0}): OK", Enable));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPECreateConnection : TestFuncBase
        {

            public TestFuncDPECreateConnection(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPECreateConnection";
            }

            public override void RunTest()
            {
                try
                {
                    int CRef;

                    DrvTestClass.ProfDpDrv.MDPECreateConnection(out CRef);

                    DrvTestClass.DisplayText(string.Format("DPECreateConnection(): CRef={0:X08}", CRef));
                    DrvTestClass.LastCRef = CRef;
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEDestroyConnection : TestFuncBase
        {

            public TestFuncDPEDestroyConnection(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPEDestroyConnection";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "CRef (hex)";
                FieldText = string.Format("{0:X08}", DrvTestClass.LastCRef);
            }

            public override void RunTest()
            {
                try
                {
                    int CRef;

                    try
                    {
                        CRef = (int)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPEDestroyConnection(CRef);

                    DrvTestClass.DisplayText(string.Format("DPEDestroyConnection({0:X08}): OK", CRef));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEReadResult : TestFuncBase
        {

            public TestFuncDPEReadResult(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPEReadResult";
            }

            public override void RunTest()
            {
                try
                {
                    MDpeRespBase DPEResult = null;

                    DrvTestClass.ProfDpDrv.MDPEReadResult(ref DPEResult);

                    DrvTestClass.DisplayText(string.Format("DPEReadResult(): OK"));
                    DrvTestClass.PrintDPEResult(DPEResult);
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEResetReq : TestFuncBase
        {

            public TestFuncDPEResetReq(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPE Reset Request";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "CRef (hex)";
                FieldText = string.Format("{0:X08}", DrvTestClass.LastCRef);
            }

            public override void RunTest()
            {
                try
                {
                    MDpeReqMASC2Reset DPEReqReset;

                    DPEReqReset = new MDpeReqMASC2Reset();
                    try
                    {
                        DPEReqReset.m_CRef = (int)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqReset);

                    DrvTestClass.DisplayText(string.Format("DPEResetRequest({0:X08}): OK", DPEReqReset.m_CRef));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEInitReq : TestFuncBase
        {

            public TestFuncDPEInitReq(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPE Init Request";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "CRef (hex)";
                FieldText = string.Format("{0:X08}", DrvTestClass.LastCRef);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slave Address";
                FieldText = string.Format("{0}", DrvTestClass.LastSlaveAddress);
            }

            public override void InitLine3(ref string FieldName, ref string FieldText)
            {
                FieldName = "Send Timeout [10ms]";
                FieldText = "100";
            }

            public override void InitLine4(ref string FieldName, ref string FieldText)
            {
                FieldName = "Features Supported 1 2";
                FieldText = "01 00";
            }

            public override void InitLine5(ref string FieldName, ref string FieldText)
            {
                FieldName = "Prof. Features Supported 1 2";
                FieldText = "00 00";
            }

            public override void InitLine6(ref string FieldName, ref string FieldText)
            {
                FieldName = "Prof. Ident Number (hex)";
                FieldText = "0000";
            }

            public override void RunTest()
            {
                try
                {
                    MDpeReqMASC2Init DPEReqInit;
                    byte[] FeatSupp = null;
                    byte[] ProfFeatSupp = null;

                    DPEReqInit = new MDpeReqMASC2Init();

                    try
                    {
                        DPEReqInit.m_CRef = (int)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                        DPEReqInit.m_RemAdd = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit2.Text));
                        DPEReqInit.m_SendTimeout = (int)Math.Round(Conversion.Val(DrvTestClass.InputEdit3.Text));
                        DPEReqInit.m_ProfileIdentNumber = (int)Math.Round(Conversion.Val(DrvTestClass.InputEdit6.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    try
                    {
                        int argLength = 2;
                        GetByteArray(DrvTestClass.InputEdit4, ref argLength, ref FeatSupp);
                        int argLength1 = 2;
                        GetByteArray(DrvTestClass.InputEdit5, ref argLength1, ref ProfFeatSupp);
                    }
                    catch (Exception ex)
                    {
                        DrvTestClass.DisplayText(ex.Message);
                        return;
                    }

                    DPEReqInit.m_FeaturesSupported1 = FeatSupp[0];
                    DPEReqInit.m_FeaturesSupported2 = FeatSupp[1];
                    DPEReqInit.m_ProfileFeaturesSupported1 = ProfFeatSupp[0];
                    DPEReqInit.m_ProfileFeaturesSupported2 = ProfFeatSupp[1];

                    DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqInit);

                    DrvTestClass.DisplayText(string.Format("DPEInitReq({0:X08}): OK", DPEReqInit.m_CRef));

                    DrvTestClass.LastSlaveAddress = DPEReqInit.m_RemAdd;
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEAbortReq : TestFuncBase
        {

            public TestFuncDPEAbortReq(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPE Abort Request";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "CRef (hex)";
                FieldText = string.Format("{0:X08}", DrvTestClass.LastCRef);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "Subnet (hex)";
                FieldText = "00";
            }

            public override void InitLine3(ref string FieldName, ref string FieldText)
            {
                FieldName = "Reason (hex)";
                FieldText = "00";
            }

            public override void RunTest()
            {
                try
                {
                    MDpeReqMASC2Abort DPEReqAbort;

                    DPEReqAbort = new MDpeReqMASC2Abort();
                    try
                    {
                        DPEReqAbort.m_CRef = (int)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                        DPEReqAbort.m_Subnet = (byte)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit2.Text));
                        DPEReqAbort.m_Reason = (byte)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit3.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqAbort);

                    DrvTestClass.DisplayText(string.Format("DPEAbortRequest({0:X08}): OK", DPEReqAbort.m_CRef));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEReadReq : TestFuncBase
        {

            public TestFuncDPEReadReq(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPE Read Request";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "CRef (hex)";
                FieldText = string.Format("{0:X08}", DrvTestClass.LastCRef);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slot";
                FieldText = string.Format("{0}", DrvTestClass.LastSlot);
            }

            public override void InitLine3(ref string FieldName, ref string FieldText)
            {
                FieldName = "Index";
                FieldText = string.Format("{0}", DrvTestClass.LastIndex);
            }

            public override void InitLine4(ref string FieldName, ref string FieldText)
            {
                FieldName = "Length";
                FieldText = string.Format("{0}", DrvTestClass.LastLength);
            }

            public override void RunTest()
            {
                try
                {
                    MDpeReqMASCXRead DPEReqRead;

                    DPEReqRead = new MDpeReqMASCXRead();
                    try
                    {
                        DPEReqRead.m_CRef = (int)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                        DPEReqRead.m_SlotNum = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit2.Text));
                        DPEReqRead.m_Index = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit3.Text));
                        DPEReqRead.m_Length = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit4.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqRead);

                    DrvTestClass.LastCRef = DPEReqRead.m_CRef;
                    DrvTestClass.LastSlot = DPEReqRead.m_SlotNum;
                    DrvTestClass.LastIndex = DPEReqRead.m_Index;
                    DrvTestClass.LastLength = DPEReqRead.m_Length;

                    DrvTestClass.DisplayText(string.Format("DPEReqRead({0:X08}): OK", DPEReqRead.m_CRef));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPEWriteReq : TestFuncBase
        {

            public TestFuncDPEWriteReq(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPE Write Request";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "CRef (hex)";
                FieldText = string.Format("{0:X08}", DrvTestClass.LastCRef);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slot";
                FieldText = string.Format("{0}", DrvTestClass.LastSlot);
            }

            public override void InitLine3(ref string FieldName, ref string FieldText)
            {
                FieldName = "Index";
                FieldText = string.Format("{0}", DrvTestClass.LastIndex);
            }

            public override void InitLine4(ref string FieldName, ref string FieldText)
            {
                FieldName = "Data (hex)";
                FieldText = "00";
            }

            public override void RunTest()
            {
                try
                {
                    MDpeReqMASCXWrite DPEReqWrite;

                    DPEReqWrite = new MDpeReqMASCXWrite();
                    try
                    {
                        DPEReqWrite.m_CRef = (int)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                        DPEReqWrite.m_SlotNum = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit2.Text));
                        DPEReqWrite.m_Index = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit3.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    try
                    {
                        int argLength = 0;
                        GetByteArray(DrvTestClass.InputEdit4, ref argLength, ref DPEReqWrite.m_PduData);
                    }
                    catch (Exception ex)
                    {
                        DrvTestClass.DisplayText(ex.Message);
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqWrite);

                    DrvTestClass.LastCRef = DPEReqWrite.m_CRef;
                    DrvTestClass.LastSlot = DPEReqWrite.m_SlotNum;
                    DrvTestClass.LastIndex = DPEReqWrite.m_Index;

                    DrvTestClass.DisplayText(string.Format("DPEReqWrite({0:X08}): OK", DPEReqWrite.m_CRef));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        public class TestFuncDPETransReq : TestFuncBase
        {

            public TestFuncDPETransReq(ref ProfDpTest TestClass) : base(ref TestClass)
            {
                FuncName = "DPE Transport Request";
            }

            public override void InitLine1(ref string FieldName, ref string FieldText)
            {
                FieldName = "CRef (hex)";
                FieldText = string.Format("{0:X08}", DrvTestClass.LastCRef);
            }

            public override void InitLine2(ref string FieldName, ref string FieldText)
            {
                FieldName = "Slot";
                FieldText = string.Format("{0}", DrvTestClass.LastSlot);
            }

            public override void InitLine3(ref string FieldName, ref string FieldText)
            {
                FieldName = "Index";
                FieldText = string.Format("{0}", DrvTestClass.LastIndex);
            }

            public override void InitLine4(ref string FieldName, ref string FieldText)
            {
                FieldName = "Data (hex)";
                FieldText = "00";
            }

            public override void RunTest()
            {
                try
                {
                    MDpeReqMASC2Transport DPEReqTransport;

                    DPEReqTransport = new MDpeReqMASC2Transport();
                    try
                    {
                        DPEReqTransport.m_CRef = (int)Math.Round(Conversion.Val("&H" + DrvTestClass.InputEdit1.Text));
                        DPEReqTransport.m_SlotNum = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit2.Text));
                        DPEReqTransport.m_Index = (byte)Math.Round(Conversion.Val(DrvTestClass.InputEdit3.Text));
                    }
                    catch
                    {
                        DrvTestClass.DisplayText("Illegal Value");
                        return;
                    }

                    try
                    {
                        int argLength = 0;
                        GetByteArray(DrvTestClass.InputEdit4, ref argLength, ref DPEReqTransport.m_PduData);
                    }
                    catch (Exception ex)
                    {
                        DrvTestClass.DisplayText(ex.Message);
                        return;
                    }

                    DrvTestClass.ProfDpDrv.MDPESendRequest(DPEReqTransport);

                    DrvTestClass.LastCRef = DPEReqTransport.m_CRef;
                    DrvTestClass.LastSlot = DPEReqTransport.m_SlotNum;
                    DrvTestClass.LastIndex = DPEReqTransport.m_Index;

                    DrvTestClass.DisplayText(string.Format("DPEReqTrans({0:X08}): OK", DPEReqTransport.m_CRef));
                }

                catch (CMProfDpException ex)
                {
                    DrvTestClass.PrintError((int)ex.m_ErrorCode);
                    return;
                }

            }
        }

        // get array of bytes from the given edit form
        // if length is 0 no length check is performed
        private static void GetByteArray(TextBox TextBox, ref int Length, ref byte[] ByteArray)
        {

            int i;
            string delimStr = " ,";
            char[] delimiter = delimStr.ToCharArray();
            string[] SubStrings;
            int SubStringCount;

            // generate list of delimiters
            SubStrings = TextBox.Text.Split(delimiter);

            // count nummber of sub strings
            SubStringCount = 0;
            foreach (string s in SubStrings)
            {
                if (s.Length > 0)
                {
                    // non empty string
                    SubStringCount += 1;
                }
            }

            if (Length > 0 & SubStringCount != Length)
            {
                // length specified and incorrect
                throw new Exception("Invalid Parameter Length");
                return;
            }
            // allocate memory of calculated size (also zero size)
            ByteArray = (byte[])Array.CreateInstance(typeof(byte), SubStringCount);

            // convert sub strings
            i = 0;
            foreach (string s in SubStrings)
            {
                if (s.Length > 0)
                {
                    // non empty string
                    try
                    {
                        ByteArray[i] = (byte)Math.Round(Conversion.Val("&H" + s));
                    }
                    catch
                    {
                        // conversion not possible
                        throw new Exception("Illegal Value");
                        return;
                    }
                    i += 1;
                }
            }

        }

    }
}
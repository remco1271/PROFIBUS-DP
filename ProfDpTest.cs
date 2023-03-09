using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using static ProfDpClass.MProfDp;                     // MProfDp Class

namespace ProfDpTest
{

    public class ProfDpTest : Form
    {

        #region  Vom Windows Form Designer generierter Code 

        public ProfDpTest() : base()
        {
            EventDPDelegate = new EventDPDel(EventDPCallback);
            EventDPCycleDelegate = new EventCycleDPDel(EventDPCycleCallback);
            EventDPEDelegate = new EventDPEDel(EventDPECallback);
            Load += ProfDpTest_Load;
            Closing += Form_Closing;
            DpCreate.Click += DpCreate_Click;
            DpDestroy.Click += DpDestroy_Click;
            ComboCommands.SelectedIndexChanged += ComboCommands_SelectedIndexChanged;
            ExecFunc.Click += ExecFunc_Click;
            CloseForm.Click += CloseForm_Click;
            ProfDpDrv.MDPEventDP += ProfDpEventDP;
            ProfDpDrv.MDPEventDPCycle += ProfDpEventDPCycle;
            ProfDpDrv.MDPEventDPE += ProfDpEventDPE;

            // Dieser Aufruf ist für den Windows Form-Designer erforderlich.
            InitializeComponent();

            // Initialisierungen nach dem Aufruf InitializeComponent() hinzufügen

        }

        // Die Form überschreibt den Löschvorgang der Basisklasse, um Komponenten zu bereinigen.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components is not null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        // Für Windows Form-Designer erforderlich
        private System.ComponentModel.IContainer components;

        // HINWEIS: Die folgende Prozedur ist für den Windows Form-Designer erforderlich
        // Sie kann mit dem Windows Form-Designer modifiziert werden.
        // Verwenden Sie nicht den Code-Editor zur Bearbeitung.
        private TextBox _DllNameEdit;

        internal virtual TextBox DllNameEdit
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _DllNameEdit;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _DllNameEdit = value;
            }
        }
        private Label _DllName;

        internal virtual Label DllName
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _DllName;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _DllName = value;
            }
        }
        private TextBox _InitStringEdit;

        internal virtual TextBox InitStringEdit
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InitStringEdit;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InitStringEdit = value;
            }
        }
        private Label _InitString;

        internal virtual Label InitString
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InitString;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InitString = value;
            }
        }
        private ColumnHeader _ColumnHeader1;

        internal virtual ColumnHeader ColumnHeader1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _ColumnHeader1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _ColumnHeader1 = value;
            }
        }
        private Button _CloseForm;

        internal virtual Button CloseForm
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _CloseForm;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_CloseForm != null)
                {
                    _CloseForm.Click -= CloseForm_Click;
                }

                _CloseForm = value;
                if (_CloseForm != null)
                {
                    _CloseForm.Click += CloseForm_Click;
                }
            }
        }
        private Button _DpCreate;

        internal virtual Button DpCreate
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _DpCreate;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_DpCreate != null)
                {
                    _DpCreate.Click -= DpCreate_Click;
                }

                _DpCreate = value;
                if (_DpCreate != null)
                {
                    _DpCreate.Click += DpCreate_Click;
                }
            }
        }
        private Button _DpDestroy;

        internal virtual Button DpDestroy
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _DpDestroy;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_DpDestroy != null)
                {
                    _DpDestroy.Click -= DpDestroy_Click;
                }

                _DpDestroy = value;
                if (_DpDestroy != null)
                {
                    _DpDestroy.Click += DpDestroy_Click;
                }
            }
        }
        private GroupBox _GroupProfDp;

        internal virtual GroupBox GroupProfDp
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _GroupProfDp;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _GroupProfDp = value;
            }
        }
        private Label _Input1;

        internal virtual Label Input1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Input1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Input1 = value;
            }
        }
        private TextBox _InputEdit1;

        internal virtual TextBox InputEdit1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InputEdit1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InputEdit1 = value;
            }
        }
        private Label _Input2;

        internal virtual Label Input2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Input2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Input2 = value;
            }
        }
        private TextBox _InputEdit2;

        internal virtual TextBox InputEdit2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InputEdit2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InputEdit2 = value;
            }
        }
        private Label _Input3;

        internal virtual Label Input3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Input3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Input3 = value;
            }
        }
        private TextBox _InputEdit3;

        internal virtual TextBox InputEdit3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InputEdit3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InputEdit3 = value;
            }
        }
        private Label _Input4;

        internal virtual Label Input4
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Input4;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Input4 = value;
            }
        }
        private TextBox _InputEdit4;

        internal virtual TextBox InputEdit4
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InputEdit4;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InputEdit4 = value;
            }
        }
        private TextBox _InputEdit5;

        internal virtual TextBox InputEdit5
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InputEdit5;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InputEdit5 = value;
            }
        }
        private Label _Input5;

        internal virtual Label Input5
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Input5;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Input5 = value;
            }
        }
        private Label _Input6;

        internal virtual Label Input6
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Input6;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Input6 = value;
            }
        }
        private TextBox _InputEdit6;

        internal virtual TextBox InputEdit6
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _InputEdit6;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _InputEdit6 = value;
            }
        }
        private Label _LabelCommands;

        internal virtual Label LabelCommands
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _LabelCommands;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _LabelCommands = value;
            }
        }
        private ComboBox _ComboCommands;

        internal virtual ComboBox ComboCommands
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _ComboCommands;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_ComboCommands != null)
                {
                    _ComboCommands.SelectedIndexChanged -= ComboCommands_SelectedIndexChanged;
                }

                _ComboCommands = value;
                if (_ComboCommands != null)
                {
                    _ComboCommands.SelectedIndexChanged += ComboCommands_SelectedIndexChanged;
                }
            }
        }
        private Button _ExecFunc;

        internal virtual Button ExecFunc
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _ExecFunc;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_ExecFunc != null)
                {
                    _ExecFunc.Click -= ExecFunc_Click;
                }

                _ExecFunc = value;
                if (_ExecFunc != null)
                {
                    _ExecFunc.Click += ExecFunc_Click;
                }
            }
        }
        private ListView _Messages;

        internal virtual ListView Messages
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Messages;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Messages = value;
            }
        }
        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            var resources = new System.Resources.ResourceManager(typeof(ProfDpTest));
            _DpCreate = new Button();
            _DpCreate.Click += new EventHandler(DpCreate_Click);
            _DllNameEdit = new TextBox();
            _DllName = new Label();
            _InitStringEdit = new TextBox();
            _InitString = new Label();
            _ColumnHeader1 = new ColumnHeader();
            _CloseForm = new Button();
            _CloseForm.Click += new EventHandler(CloseForm_Click);
            _DpDestroy = new Button();
            _DpDestroy.Click += new EventHandler(DpDestroy_Click);
            _GroupProfDp = new GroupBox();
            _InputEdit1 = new TextBox();
            _Input1 = new Label();
            _InputEdit2 = new TextBox();
            _Input2 = new Label();
            _InputEdit3 = new TextBox();
            _Input3 = new Label();
            _InputEdit4 = new TextBox();
            _Input4 = new Label();
            _InputEdit5 = new TextBox();
            _Input5 = new Label();
            _Input6 = new Label();
            _InputEdit6 = new TextBox();
            _LabelCommands = new Label();
            _ComboCommands = new ComboBox();
            _ComboCommands.SelectedIndexChanged += new EventHandler(ComboCommands_SelectedIndexChanged);
            _ExecFunc = new Button();
            _ExecFunc.Click += new EventHandler(ExecFunc_Click);
            _Messages = new ListView();
            _GroupProfDp.SuspendLayout();
            SuspendLayout();
            // 
            // DpCreate
            // 
            _DpCreate.Location = new Point(24, 24);
            _DpCreate.Name = "_DpCreate";
            _DpCreate.Size = new Size(88, 23);
            _DpCreate.TabIndex = 0;
            _DpCreate.Text = "&Create";
            // 
            // DllNameEdit
            // 
            _DllNameEdit.Location = new Point(216, 24);
            _DllNameEdit.Name = "_DllNameEdit";
            _DllNameEdit.Size = new Size(128, 20);
            _DllNameEdit.TabIndex = 2;
            _DllNameEdit.Text = "";
            // 
            // DllName
            // 
            _DllName.Location = new Point(136, 24);
            _DllName.Name = "_DllName";
            _DllName.Size = new Size(72, 16);
            _DllName.TabIndex = 1;
            _DllName.Text = "Dll Name";
            // 
            // InitStringEdit
            // 
            _InitStringEdit.Location = new Point(448, 24);
            _InitStringEdit.Name = "_InitStringEdit";
            _InitStringEdit.Size = new Size(128, 20);
            _InitStringEdit.TabIndex = 4;
            _InitStringEdit.Text = "";
            // 
            // InitString
            // 
            _InitString.Location = new Point(368, 24);
            _InitString.Name = "_InitString";
            _InitString.Size = new Size(72, 16);
            _InitString.TabIndex = 3;
            _InitString.Text = "Init String";
            // 
            // ColumnHeader1
            // 
            _ColumnHeader1.Text = "Messages";
            _ColumnHeader1.Width = 844;
            // 
            // CloseForm
            // 
            _CloseForm.Location = new Point(16, 624);
            _CloseForm.Name = "_CloseForm";
            _CloseForm.Size = new Size(88, 23);
            _CloseForm.TabIndex = 40;
            _CloseForm.Text = "&Close";
            // 
            // DpDestroy
            // 
            _DpDestroy.Location = new Point(24, 56);
            _DpDestroy.Name = "_DpDestroy";
            _DpDestroy.Size = new Size(88, 23);
            _DpDestroy.TabIndex = 5;
            _DpDestroy.Text = "&Destroy";
            // 
            // GroupProfDp
            // 
            _GroupProfDp.Controls.Add(_InputEdit6);
            _GroupProfDp.Controls.Add(_InputEdit1);
            _GroupProfDp.Controls.Add(_Input1);
            _GroupProfDp.Controls.Add(_InputEdit2);
            _GroupProfDp.Controls.Add(_Input2);
            _GroupProfDp.Controls.Add(_InputEdit3);
            _GroupProfDp.Controls.Add(_Input3);
            _GroupProfDp.Controls.Add(_InputEdit4);
            _GroupProfDp.Controls.Add(_Input4);
            _GroupProfDp.Controls.Add(_InputEdit5);
            _GroupProfDp.Controls.Add(_Input5);
            _GroupProfDp.Controls.Add(_LabelCommands);
            _GroupProfDp.Controls.Add(_ComboCommands);
            _GroupProfDp.Controls.Add(_ExecFunc);
            _GroupProfDp.Controls.Add(_Input6);
            _GroupProfDp.Location = new Point(16, 96);
            _GroupProfDp.Name = "_GroupProfDp";
            _GroupProfDp.Size = new Size(848, 216);
            _GroupProfDp.TabIndex = 6;
            _GroupProfDp.TabStop = false;
            _GroupProfDp.Text = "ProfDp Functions";
            // 
            // InputEdit1
            // 
            _InputEdit1.Location = new Point(192, 24);
            _InputEdit1.Name = "_InputEdit1";
            _InputEdit1.Size = new Size(640, 20);
            _InputEdit1.TabIndex = 1;
            _InputEdit1.Text = "";
            // 
            // Input1
            // 
            _Input1.Location = new Point(8, 24);
            _Input1.Name = "_Input1";
            _Input1.Size = new Size(176, 16);
            _Input1.TabIndex = 2;
            _Input1.Text = "Input1";
            // 
            // InputEdit2
            // 
            _InputEdit2.Location = new Point(192, 48);
            _InputEdit2.Name = "_InputEdit2";
            _InputEdit2.Size = new Size(640, 20);
            _InputEdit2.TabIndex = 3;
            _InputEdit2.Text = "";
            // 
            // Input2
            // 
            _Input2.Location = new Point(8, 48);
            _Input2.Name = "_Input2";
            _Input2.Size = new Size(176, 16);
            _Input2.TabIndex = 4;
            _Input2.Text = "Input2";
            // 
            // InputEdit3
            // 
            _InputEdit3.Location = new Point(192, 72);
            _InputEdit3.Name = "_InputEdit3";
            _InputEdit3.Size = new Size(640, 20);
            _InputEdit3.TabIndex = 5;
            _InputEdit3.Text = "";
            // 
            // Input3
            // 
            _Input3.Location = new Point(8, 72);
            _Input3.Name = "_Input3";
            _Input3.Size = new Size(176, 16);
            _Input3.TabIndex = 6;
            _Input3.Text = "Input3";
            // 
            // InputEdit4
            // 
            _InputEdit4.Location = new Point(192, 96);
            _InputEdit4.Name = "_InputEdit4";
            _InputEdit4.Size = new Size(640, 20);
            _InputEdit4.TabIndex = 7;
            _InputEdit4.Text = "";
            // 
            // Input4
            // 
            _Input4.Location = new Point(8, 96);
            _Input4.Name = "_Input4";
            _Input4.Size = new Size(176, 16);
            _Input4.TabIndex = 8;
            _Input4.Text = "Input4";
            // 
            // InputEdit5
            // 
            _InputEdit5.Location = new Point(192, 120);
            _InputEdit5.Name = "_InputEdit5";
            _InputEdit5.Size = new Size(640, 20);
            _InputEdit5.TabIndex = 9;
            _InputEdit5.Text = "";
            // 
            // Input5
            // 
            _Input5.Location = new Point(8, 120);
            _Input5.Name = "_Input5";
            _Input5.Size = new Size(176, 16);
            _Input5.TabIndex = 10;
            _Input5.Text = "Input5";
            // 
            // Input6
            // 
            _Input6.Location = new Point(8, 144);
            _Input6.Name = "_Input6";
            _Input6.Size = new Size(176, 16);
            _Input6.TabIndex = 11;
            _Input6.Text = "Input6";
            // 
            // InputEdit6
            // 
            _InputEdit6.Location = new Point(192, 144);
            _InputEdit6.Name = "_InputEdit6";
            _InputEdit6.Size = new Size(640, 20);
            _InputEdit6.TabIndex = 12;
            _InputEdit6.Text = "";
            // 
            // LabelCommands
            // 
            _LabelCommands.Location = new Point(8, 184);
            _LabelCommands.Name = "_LabelCommands";
            _LabelCommands.Size = new Size(176, 16);
            _LabelCommands.TabIndex = 20;
            _LabelCommands.Text = "Command";
            // 
            // ComboCommands
            // 
            _ComboCommands.DropDownStyle = ComboBoxStyle.DropDownList;
            _ComboCommands.Location = new Point(192, 176);
            _ComboCommands.Name = "_ComboCommands";
            _ComboCommands.Size = new Size(296, 21);
            _ComboCommands.TabIndex = 21;
            // 
            // ExecFunc
            // 
            _ExecFunc.Location = new Point(512, 176);
            _ExecFunc.Name = "_ExecFunc";
            _ExecFunc.TabIndex = 22;
            _ExecFunc.Text = "&Execute";
            // 
            // Messages
            // 
            _Messages.Columns.AddRange(new ColumnHeader[] { _ColumnHeader1 });
            _Messages.FullRowSelect = true;
            _Messages.GridLines = true;
            _Messages.Location = new Point(16, 328);
            _Messages.Name = "_Messages";
            _Messages.Size = new Size(848, 280);
            _Messages.TabIndex = 30;
            _Messages.View = View.Details;
            // 
            // ProfDpTest
            // 
            AutoScaleBaseSize = new Size(5, 13);
            ClientSize = new Size(880, 661);
            Controls.Add(_DpCreate);
            Controls.Add(_DllName);
            Controls.Add(_DllNameEdit);
            Controls.Add(_InitStringEdit);
            Controls.Add(_InitString);
            Controls.Add(_DpDestroy);
            Controls.Add(_Messages);
            Controls.Add(_CloseForm);
            Controls.Add(_GroupProfDp);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "ProfDpTest";
            Text = "ProfDpTest";
            _GroupProfDp.ResumeLayout(false);
            ResumeLayout(false);

        }

        #endregion

        // Type for test function call
        private enum ProfDpCallType
        {
            Call_Init,       // Init Screen
            Call_Execute    // Execute Program
        }

        private ProfDpClass.MProfDp _ProfDpDrv;

        public virtual ProfDpClass.MProfDp ProfDpDrv
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _ProfDpDrv;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_ProfDpDrv != null)
                {
                    _ProfDpDrv.MDPEventDP -= ProfDpEventDP;
                    _ProfDpDrv.MDPEventDPCycle -= ProfDpEventDPCycle;
                    _ProfDpDrv.MDPEventDPE -= ProfDpEventDPE;
                }

                _ProfDpDrv = value;
                if (_ProfDpDrv != null)
                {
                    _ProfDpDrv.MDPEventDP += ProfDpEventDP;
                    _ProfDpDrv.MDPEventDPCycle += ProfDpEventDPCycle;
                    _ProfDpDrv.MDPEventDPE += ProfDpEventDPE;
                }
            }
        }

        // list of test function
        private ArrayList TestFuncList = new ArrayList();

        // delegate for calling EventDPCallback
        public delegate void EventDPDel();

        private EventDPDel EventDPDelegate;

        // delegate for calling EventDPCycleCallback
        public delegate void EventCycleDPDel();

        private EventCycleDPDel EventDPCycleDelegate;

        // delegate for calling EventDPECallback
        public delegate void EventDPEDel();

        private EventDPEDel EventDPEDelegate;

        // last used slave address
        public byte LastSlaveAddress;

        // last used CRef
        public int LastCRef;

        // last used slot
        public byte LastSlot;

        // last used index
        public byte LastIndex;

        // last used length
        public byte LastLength;

        // function is called if main window opens
        private void ProfDpTest_Load(object sender, EventArgs e)
        {

            InitStringEdit.Text = "COM2";
            LastSlaveAddress = 3;
            LastCRef = 0;
            LastSlot = 0;
            LastIndex = 0;
            LastLength = 1;

            // disable clickable headers
            Messages.HeaderStyle = ColumnHeaderStyle.Nonclickable;

            // Create list of test functions
            var argTestClass = this;
            TestFuncList.Add(new TestFunc.TestFuncDPIdentifySlave(ref argTestClass));
            var argTestClass1 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPChangeAddress(ref argTestClass1));
            var argTestClass2 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPGetCfg(ref argTestClass2));
            var argTestClass3 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPSendPrm(ref argTestClass3));
            var argTestClass4 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPOpenSlave(ref argTestClass4));
            var argTestClass5 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPCloseSlave(ref argTestClass5));
            var argTestClass6 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPReadSlaveData(ref argTestClass6));
            var argTestClass7 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPWriteSlaveData(ref argTestClass7));
            var argTestClass8 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPReadResult(ref argTestClass8));
            var argTestClass9 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPSetFDLReceiveTimeout(ref argTestClass9));
            var argTestClass10 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPCheckConvFeature(ref argTestClass10));
            var argTestClass11 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEnableDPEvent(ref argTestClass11));
            var argTestClass12 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEnableDPCycleEvent(ref argTestClass12));
            var argTestClass13 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEnableDPEEvent(ref argTestClass13));
            var argTestClass14 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPECreateConnection(ref argTestClass14));
            var argTestClass15 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEDestroyConnection(ref argTestClass15));
            var argTestClass16 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEReadResult(ref argTestClass16));
            var argTestClass17 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEResetReq(ref argTestClass17));
            var argTestClass18 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEInitReq(ref argTestClass18));
            var argTestClass19 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEAbortReq(ref argTestClass19));
            var argTestClass20 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEReadReq(ref argTestClass20));
            var argTestClass21 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPEWriteReq(ref argTestClass21));
            var argTestClass22 = this;
            TestFuncList.Add(new TestFunc.TestFuncDPETransReq(ref argTestClass22));

            // fill the combo box with the test function names
            var myEnumerator = TestFuncList.GetEnumerator();
            while (myEnumerator.MoveNext())
                ComboCommands.Items.Add(((TestFunc.TestFuncBase)myEnumerator.Current).FuncName);

            // select first entry
            ComboCommands.SelectedIndex = 0;

        }

        public string ErrorCode2Text(int ErrorCode)
        {
            string ErrorCode2TextRet = default;

            switch (ErrorCode)
            {
                case (int)MDpResult.MDP_NO_ERROR:           // no error
                    {
                        ErrorCode2TextRet = "No error";
                        break;
                    }

                case (int)MDpResult.MDP_GENERIC_ERROR:      // generic error
                    {
                        ErrorCode2TextRet = "Generic error";
                        break;
                    }

                case (int)MDpResult.MDP_CALLED_TWICE:       // DP function have been called double (only with DPSetDispatchMessage())
                    {
                        ErrorCode2TextRet = "Function called twice";
                        break;
                    }

                case (int)MDpResult.MDP_OPEN_FAILED:        // device (FDL) opening failed
                    {
                        ErrorCode2TextRet = "FDL opening failed";
                        break;
                    }

                case (int)MDpResult.MDP_FDL_LOAD_FAILED:    // unable to load FDL library
                    {
                        ErrorCode2TextRet = "Unable to call FDL library";
                        break;
                    }

                case (int)MDpResult.MDP_NOT_OPEN:           // FDL not opened
                    {
                        ErrorCode2TextRet = "FDL not opened";
                        break;
                    }

                case (int)MDpResult.MDP_THREAD_ENABLED:     // error, background thread enabled: function disabled
                    {
                        ErrorCode2TextRet = "Function disabled because thread is enabled";
                        break;
                    }

                case (int)MDpResult.MDP_NO_THREAD_ENABLED:  // error, background thread is not enabled: function disabled
                    {
                        ErrorCode2TextRet = "Function disabled because thread is not enabled";
                        break;
                    }

                case (int)MDpResult.MDP_OUT_OF_MEMORY:      // out of memory
                    {
                        ErrorCode2TextRet = "Out of memory";
                        break;
                    }

                case (int)MDpResult.MDP_CONVERTER_FAILED:   // PROFIBUS converter not found, failed or bad type
                    {
                        ErrorCode2TextRet = "PROFIBUS converter not found or bad type";
                        break;
                    }

                case (int)MDpResult.MDP_ILL_ADDRESS:        // illegal DP slave address
                    {
                        ErrorCode2TextRet = "Illegal DP slave address";
                        break;
                    }

                case (int)MDpResult.MDP_SAP_OPEN_FAILED:    // PROFIBUS SAP opening failed
                    {
                        ErrorCode2TextRet = "PROFIBUS SAP opening failed";
                        break;
                    }

                case (int)MDpResult.MDP_SLAVE_OPEN:         // slave is already open
                    {
                        ErrorCode2TextRet = "Slave is already open";
                        break;
                    }

                case (int)MDpResult.MDP_SLAVE_NOT_OPEN:     // slave is not opened
                    {
                        ErrorCode2TextRet = "Slave is not opened";
                        break;
                    }

                case (int)MDpResult.MDP_SLAVE_NOT_FOUND:    // slave not found
                    {
                        ErrorCode2TextRet = "Slave not found";
                        break;
                    }

                case (int)MDpResult.MDP_ILLEGAL_CONFIG:     // illegal PROFIBUS configuration
                    {
                        ErrorCode2TextRet = "Illegal PROFIBUS configuration";
                        break;
                    }

                case (int)MDpResult.MDP_CFG_ERROR:          // configuration error on startup
                    {
                        ErrorCode2TextRet = "Configuration error on startup";
                        break;
                    }

                case (int)MDpResult.MDP_PRM_ERROR:          // parameter error on startup
                    {
                        ErrorCode2TextRet = "Parameter error on startup";
                        break;
                    }

                case (int)MDpResult.MDP_CFG_AND_PRM_ERROR:  // configuration and parameter error on startup
                    {
                        ErrorCode2TextRet = "Configuration and parameter error on startup";
                        break;
                    }

                case (int)MDpResult.MDP_BAD_LENGTH:         // data length incorrect
                    {
                        ErrorCode2TextRet = "Data length incorrect";
                        break;
                    }

                case (int)MDpResult.MDP_INCOMPLETE:         // background operation initiated
                    {
                        ErrorCode2TextRet = "Background operation initiated";
                        break;
                    }

                case (int)MDpResult.MDP_NORESULT:           // no result pending
                    {
                        ErrorCode2TextRet = "No result pending";
                        break;
                    }

                default:
                    {
                        ErrorCode2TextRet = string.Format("Unknown error code {0:X04}", ErrorCode);
                        break;
                    }
            }

            return ErrorCode2TextRet;

        }

        // convert reject codes to text
        public string RejectCode2Text(MDpeRejectCodes RejectCode)
        {
            string RejectCode2TextRet = default;
            switch (RejectCode)
            {
                case MDpeRejectCodes.MDPE_REJ_LE:            // max_data_len_unit overflow
                    {
                        RejectCode2TextRet = "max_data_len_unit overflow";
                        break;
                    }

                case MDpeRejectCodes.MDPE_REJ_PS:            // number of parallel service requests exceeded
                    {
                        RejectCode2TextRet = "number of parallel service requests exceeded";
                        break;
                    }

                case MDpeRejectCodes.MDPE_REJ_SE:            // MSAC1 service reject
                    {
                        RejectCode2TextRet = "MSAC1 service reject";
                        break;
                    }

                case MDpeRejectCodes.MDPE_REJ_ABORT:         // MSAC1 abort reject
                    {
                        RejectCode2TextRet = "MSAC1 abort reject";
                        break;
                    }

                case MDpeRejectCodes.MDPE_REJ_IV:            // MSAC1 invalid service reject
                    {
                        RejectCode2TextRet = "MSAC1 invalid service reject";
                        break;
                    }

                default:
                    {
                        RejectCode2TextRet = string.Format("Unknown reject code {0:X04}", (int)RejectCode);
                        break;
                    }

            }

            return RejectCode2TextRet;
        }

        // convert fault codes to text
        public string FaultCode2Text(MDpeFaultCodes FaultCode)
        {
            string FaultCode2TextRet = default;
            switch (FaultCode)
            {
                case MDpeFaultCodes.MDPE_DDLM_FDL_FAULT:    // FDL fault
                    {
                        FaultCode2TextRet = "FDL fault";
                        break;
                    }

                case MDpeFaultCodes.MDPE_MSAC1_FAULT:       // number of parallel service requests exceeded
                    {
                        FaultCode2TextRet = "MSAC1 processing fault";
                        break;
                    }

                default:
                    {
                        FaultCode2TextRet = string.Format("Unknown fault code {0:X04}", (int)FaultCode);
                        break;
                    }

            }

            return FaultCode2TextRet;
        }

        // print Error code to the output window
        public void PrintError(int ErrorCode)
        {

            if (ErrorCode != (int)MDpResult.MDP_NO_ERROR)
            {
                // error occured -> display it
                DisplayText("Error: " + ErrorCode2Text(ErrorCode));
            }

        }

        // print DpData structure
        public void PrintDpData(MDpData DpData)
        {
            string Output;

            DisplayText(string.Format(" DPV1_Supported={0}", DpData.m_DPV1_Supported));
            DisplayText(string.Format(" Max_Channel_Data_Length={0}", DpData.m_Max_Channel_Data_Length));
            DisplayText(string.Format(" Extra_Alarm_SAP={0}", DpData.m_Extra_Alarm_SAP));
            DisplayText(string.Format(" C1_Response_Timeout={0}", DpData.m_C1_Response_Timeout));

            if (DpData.m_NormParameters is not null)
            {
                Output = " NormParameters=";
                if (DpData.m_NormParameters.Length > 0)
                {
                    for (int i = 0, loopTo = DpData.m_NormParameters.Length - 1; i <= loopTo; i++)
                        Output += DpData.m_NormParameters[i].ToString("X02") + " ";
                }
                DisplayText(Output);
            }

            if (DpData.m_UserParameters is not null)
            {
                Output = " UserParameters=";
                if (DpData.m_UserParameters.Length > 0)
                {
                    for (int i = 0, loopTo1 = DpData.m_UserParameters.Length - 1; i <= loopTo1; i++)
                        Output += DpData.m_UserParameters[i].ToString("X02") + " ";
                }
                DisplayText(Output);
            }

            if (DpData.m_Configuration is not null)
            {
                Output = " Configuration=";
                if (DpData.m_Configuration.Length > 0)
                {
                    for (int i = 0, loopTo2 = DpData.m_Configuration.Length - 1; i <= loopTo2; i++)
                        Output += DpData.m_Configuration[i].ToString("X02") + " ";
                }
                DisplayText(Output);
            }

            if (DpData.m_OutputData is not null)
            {
                Output = " OutputData=";
                if (DpData.m_OutputData.Length > 0)
                {
                    for (int i = 0, loopTo3 = DpData.m_OutputData.Length - 1; i <= loopTo3; i++)
                        Output += DpData.m_OutputData[i].ToString("X02") + " ";
                }
                DisplayText(Output);
            }

            if (DpData.m_InputData is not null)
            {
                Output = " InputData=";
                if (DpData.m_InputData.Length > 0)
                {
                    for (int i = 0, loopTo4 = DpData.m_InputData.Length - 1; i <= loopTo4; i++)
                        Output += DpData.m_InputData[i].ToString("X02") + " ";
                }
                DisplayText(Output);
            }

            if (DpData.m_NormDiagnosis is not null)
            {
                Output = " NormDiagnosis=";
                if (DpData.m_NormDiagnosis.Length > 0)
                {
                    for (int i = 0, loopTo5 = DpData.m_NormDiagnosis.Length - 1; i <= loopTo5; i++)
                        Output += DpData.m_NormDiagnosis[i].ToString("X02") + " ";
                }
                DisplayText(Output);
            }

            if (DpData.m_UserDiagnosis is not null)
            {
                Output = " UserDiagnosis=";
                if (DpData.m_UserDiagnosis.Length > 0)
                {
                    for (int i = 0, loopTo6 = DpData.m_UserDiagnosis.Length - 1; i <= loopTo6; i++)
                        Output += DpData.m_UserDiagnosis[i].ToString("X02") + " ";
                }
                DisplayText(Output);
            }

            DisplayText(string.Format(" Communicate={0}", DpData.m_Communicate));
            DisplayText(string.Format(" Operate={0}", DpData.m_Operate));
            DisplayText(string.Format(" New Diag={0}", DpData.m_NewDiag));
        }

        // print PROFIBUS DP result information
        public void PrintDPResult(MDpResultBase DPResult)
        {

            DisplayText(string.Format("DPResult ID={0:X04}, Error='{1}'", DPResult.m_ID, ErrorCode2Text((int)DPResult.m_ResultCode)));

            if (DPResult is MDpResultOpenSlave)
            {
                // result of MDPOpenSlave
                DisplayText(" DPOpenSlave");
                return;
            }

            if (DPResult is MDpResultCloseSlave)
            {
                // result of MDPCloseSlave
                DisplayText(" DPCloseSlave");
                return;
            }

            if (DPResult is MDpResultGetCfg)
            {
                // result of get cfg command
                string Output;
                byte[] Config;

                DisplayText(" DPGetCfg");

                if (DPResult.m_ResultCode != MDpResult.MDP_NO_ERROR)
                {
                    // error occured don't print more info
                    return;
                }

                Config = ((MDpResultGetCfg)DPResult).m_Config;
                Output = " Config=";
                if (Config.Length > 0)
                {
                    for (int i = 0, loopTo = Config.Length - 1; i <= loopTo; i++)
                        Output += Config[i].ToString("X02") + " ";
                }
                DisplayText(Output);
                return;
            }

            if (DPResult is MDpResultSendPrm)
            {
                // result of MDPSendPrm
                DisplayText(" DPSendPrm");
                return;
            }

            if (DPResult is MDpResultIdent)
            {
                // result of get ident command
                DisplayText(" DPIdentifySlave");

                if (DPResult.m_ResultCode != MDpResult.MDP_NO_ERROR)
                {
                    // error occured don't print more info
                    return;
                }

                DisplayText(string.Format(" IdentNumber (hex)={0:X04}", ((MDpResultIdent)DPResult).m_IdentNumber));
                return;
            }

            if (DPResult is MDpResultChangeAddress)
            {
                // result of MDPChangeAddress
                DisplayText(" DPChangeAddress");
                return;
            }

            DisplayText(" Unknown Result Type");

        }

        // print negative confirmation
        public void PrintConfirmationNeg(MDpeRespConfNegBase RespConfNegBase)
        {
            DisplayText(string.Format(" ErrorDecode (hex)={0:X02}", RespConfNegBase.m_ErrorDecode));
            DisplayText(string.Format(" ErrorCode1 (hex)={0:X02}", RespConfNegBase.m_ErrorCode1));
            DisplayText(string.Format(" ErrorCode2 (hex)={0:X02}", RespConfNegBase.m_ErrorCode2));
        }

        // print Error code to the output window
        public void PrintDPEResult(MDpeRespBase DPEResult)
        {

            DisplayText(string.Format("DPEResult ID={0:X08}", DPEResult.m_CRef));
            if (DPEResult is MDpeRespMSAC2ConfInitPos)
            {
                // MSAC2 initiate confirmation positive
                MDpeRespMSAC2ConfInitPos RespMSAC2ConfInitPos;
                string Output;

                DisplayText(" MSAC2 Init Confirmation Positive");

                RespMSAC2ConfInitPos = (MDpeRespMSAC2ConfInitPos)DPEResult;

                DisplayText(string.Format(" FeaturesSupported 1,2 (hex)={0:X02},{1:X02}", RespMSAC2ConfInitPos.m_FeaturesSupported1, RespMSAC2ConfInitPos.m_FeaturesSupported2));
                DisplayText(string.Format(" ProfileFeaturesSupported 1,2 (hex)={0:X02},{1:X02}", RespMSAC2ConfInitPos.m_ProfileFeaturesSupported1, RespMSAC2ConfInitPos.m_ProfileFeaturesSupported2));
                DisplayText(string.Format(" ProfileIdentNumber (hex)={0:X04}", RespMSAC2ConfInitPos.m_ProfileIdentNumber));
                DisplayText(string.Format(" MaxLenDataUnit={0}", RespMSAC2ConfInitPos.m_MaxLenDataUnit));
                if (RespMSAC2ConfInitPos.m_AddAddrParam is not null)
                {
                    DisplayText(" AddAddrParam:");
                    // source address
                    DisplayText(string.Format("  Source API (hex)={0:X02}", RespMSAC2ConfInitPos.m_AddAddrParam.m_SAPI));
                    DisplayText(string.Format("  Source SCL (hex)={0:X02}", RespMSAC2ConfInitPos.m_AddAddrParam.m_SSCL));

                    if (RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr is not null)
                    {
                        Output = "  Source NetAddr=";
                        if (RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr.Length > 0)
                        {
                            for (int i = 0, loopTo = RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr.Length - 1; i <= loopTo; i++)
                                Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_SNetAddr[i].ToString("X02") + " ";
                        }
                        DisplayText(Output);
                    }

                    if (RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr is not null)
                    {
                        Output = "  Source MACAddr=";
                        if (RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr.Length > 0)
                        {
                            for (int i = 0, loopTo1 = RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr.Length - 1; i <= loopTo1; i++)
                                Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_SMACAddr[i].ToString("X02") + " ";
                        }
                        DisplayText(Output);
                    }
                    // destination address
                    DisplayText(string.Format("  Destination API (hex)={0:X02}", RespMSAC2ConfInitPos.m_AddAddrParam.m_DAPI));
                    DisplayText(string.Format("  Destination SCL (hex)={0:X02}", RespMSAC2ConfInitPos.m_AddAddrParam.m_DSCL));

                    if (RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr is not null)
                    {
                        Output = "  Destination NetAddr=";
                        if (RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr.Length > 0)
                        {
                            for (int i = 0, loopTo2 = RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr.Length - 1; i <= loopTo2; i++)
                                Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_DNetAddr[i].ToString("X02") + " ";
                        }
                        DisplayText(Output);
                    }

                    if (RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr is not null)
                    {
                        Output = "  Destination MACAddr=";
                        if (RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr.Length > 0)
                        {
                            for (int i = 0, loopTo3 = RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr.Length - 1; i <= loopTo3; i++)
                                Output += RespMSAC2ConfInitPos.m_AddAddrParam.m_DMACAddr[i].ToString("X02") + " ";
                        }
                        DisplayText(Output);
                    }
                }

                return;
            }

            if (DPEResult is MDpeRespMSAC2ConfInitNeg)
            {
                // MSAC2 initiate confirmation negative
                DisplayText(" MSAC2 Init Confirmation Negative");

                PrintConfirmationNeg((MDpeRespConfNegBase)DPEResult);

                return;
            }

            if (DPEResult is MDpeRespMSACXConfReadPos)
            {
                // MSACX read confirmation positive
                MDpeRespMSACXConfReadPos RespMSAC2ConfReadPos;
                string Output;

                DisplayText(" MSACX Read Confirmation Positive");

                RespMSAC2ConfReadPos = (MDpeRespMSACXConfReadPos)DPEResult;

                if (RespMSAC2ConfReadPos.m_Data is not null)
                {
                    Output = " Data=";
                    if (RespMSAC2ConfReadPos.m_Data.Length > 0)
                    {
                        for (int i = 0, loopTo4 = RespMSAC2ConfReadPos.m_Data.Length - 1; i <= loopTo4; i++)
                            Output += RespMSAC2ConfReadPos.m_Data[i].ToString("X02") + " ";
                    }
                    DisplayText(Output);
                }

                return;
            }

            if (DPEResult is MDpeRespMSACXConfReadNeg)
            {
                // MSACX read confirmation negative
                DisplayText(" MSACX Read Confirmation Negative");

                PrintConfirmationNeg((MDpeRespConfNegBase)DPEResult);

                return;
            }

            if (DPEResult is MDpeRespMSACXConfWritePos)
            {
                // MSACX read confirmation positive
                MDpeRespMSACXConfWritePos RespMSAC2ConfWritePos;

                DisplayText(" MSACX Write Confirmation Positive");

                RespMSAC2ConfWritePos = (MDpeRespMSACXConfWritePos)DPEResult;

                DisplayText(string.Format(" Length={0}", RespMSAC2ConfWritePos.m_Length));

                return;
            }

            if (DPEResult is MDpeRespMSACXConfWriteNeg)
            {
                // MSACX write confirmation negative
                DisplayText(" MSACX Write Confirmation Negative");

                PrintConfirmationNeg((MDpeRespConfNegBase)DPEResult);

                return;
            }

            if (DPEResult is MDpeRespMSAC2ConfTransPos)
            {
                // MSAC2 data transport confirmation positive
                MDpeRespMSAC2ConfTransPos RespMSAC2ConfTransPos;
                string Output;

                DisplayText(" MSAC2 Data Transport Confirmation Positive");

                RespMSAC2ConfTransPos = (MDpeRespMSAC2ConfTransPos)DPEResult;

                if (RespMSAC2ConfTransPos.m_Data is not null)
                {
                    Output = " Data=";
                    if (RespMSAC2ConfTransPos.m_Data.Length > 0)
                    {
                        for (int i = 0, loopTo5 = RespMSAC2ConfTransPos.m_Data.Length - 1; i <= loopTo5; i++)
                            Output += RespMSAC2ConfTransPos.m_Data[i].ToString("X02") + " ";
                    }
                    DisplayText(Output);
                }

                return;
            }

            if (DPEResult is MDpeRespMSAC2ConfTransNeg)
            {
                // MSAC2 data transport confirmation negative
                DisplayText(" MSAC2 Data Transport Confirmation Negative");

                PrintConfirmationNeg((MDpeRespConfNegBase)DPEResult);

                return;
            }

            if (DPEResult is MDpeRespMSAC2ConfReset)
            {
                // MSAC2 reset confirmation
                DisplayText(" MSAC2 Reset Confirmation");

                return;
            }

            if (DPEResult is MDpeRespMSAC2IndiAbort)
            {
                // MSAC2 abort indication
                MDpeRespMSAC2IndiAbort RespMSAC2IndiAbort;

                DisplayText(" MSAC2 Abort Indication");

                RespMSAC2IndiAbort = (MDpeRespMSAC2IndiAbort)DPEResult;

                DisplayText(string.Format(" LocallyGenerated={0}", RespMSAC2IndiAbort.m_LocallyGenerated));
                DisplayText(string.Format(" Subnet (hex)={0:X02}", RespMSAC2IndiAbort.m_Subnet));
                DisplayText(string.Format(" ReasonCode (hex)={0:X02}", RespMSAC2IndiAbort.m_ReasonCode));
                DisplayText(string.Format(" AdditionalDetail (hex)={0:X04}", RespMSAC2IndiAbort.m_AdditionalDetail));

                return;
            }

            if (DPEResult is MDpeRespMSAC2ConfReset)
            {
                // MSAC2 reset confirmation
                DisplayText(" MSAC2 Reset Confirmation");

                return;
            }

            if (DPEResult is MDpeRespMSAC2IndiClosed)
            {
                // MSAC2 closed indication
                DisplayText(" MSAC2 Closed Indication");

                return;
            }

            if (DPEResult is MDpeRespMSAC2IndiReject)
            {
                // MSAC2 reject indication
                MDpeRespMSAC2IndiReject RespMSAC2IndiReject;

                DisplayText(" MSAC2 Reject Indication");

                RespMSAC2IndiReject = (MDpeRespMSAC2IndiReject)DPEResult;

                DisplayText(string.Format(" ReasonCode={0}", RejectCode2Text(RespMSAC2IndiReject.m_ReasonCode)));

                return;
            }

            if (DPEResult is MDpeRespMSAC2IndiFault)
            {
                // MSAC2 fault indication
                DisplayText(" MSAC2 Fault Indication");

                return;
            }

            if (DPEResult is MDpeRespMASC1IndiStarted)
            {
                // MSAC2 started indication
                DisplayText(" MSAC1 Started Indication");

                return;
            }

            if (DPEResult is MDpeRespMASC1IndiStopped)
            {
                // MSAC2 stopped indication
                DisplayText(" MSAC1 Stopped Indication");

                return;
            }

            if (DPEResult is MDpeRespMASC1IndiReject)
            {
                // MSAC1 reject indication
                MDpeRespMASC1IndiReject RespMASC1IndiReject;

                DisplayText(" MSAC1 Reject Indication");

                RespMASC1IndiReject = (MDpeRespMASC1IndiReject)DPEResult;

                DisplayText(string.Format(" Status={0}", RejectCode2Text(RespMASC1IndiReject.m_Status)));

                return;
            }

            if (DPEResult is MDpeRespMASC1IndiFault)
            {
                // MSAC1 fault indication
                MDpeRespMASC1IndiFault RespMASC1IndiFault;

                DisplayText(" MSAC1 Fault Indication");

                RespMASC1IndiFault = (MDpeRespMASC1IndiFault)DPEResult;

                DisplayText(string.Format(" Status={0}", FaultCode2Text(RespMASC1IndiFault.m_Status)));

                return;
            }

            DisplayText(" Unknown Result Type");
        }

        // print given text to the output window
        public void DisplayText(string Text)
        {

            var ListViewItemNew = new ListViewItem(Text);

            ListViewItemNew.Tag = Text;
            Messages.Items.AddRange(new ListViewItem[] { ListViewItemNew });
            Messages.EnsureVisible(Messages.Items.Count - 1);

        }

        // call the specified test function with init or execute
        private void CallSelectedFunction(string Selection, ProfDpCallType CallType)
        {

            TestFunc.TestFuncBase TestFunc;

            // search the function name in the list of test functions
            var myEnumerator = TestFuncList.GetEnumerator();
            while (myEnumerator.MoveNext())
            {
                TestFunc = (TestFunc.TestFuncBase)myEnumerator.Current;
                if ((TestFunc.FuncName ?? "") == (Selection ?? ""))
                {
                    // function found
                    switch (CallType)
                    {
                        case ProfDpCallType.Call_Init:  // init display
                            {
                                string FieldName = null;    // text beside edit line
                                string FieldText = null;    // text in edit field

                                // init edit line 1
                                TestFunc.InitLine1(ref FieldName, ref FieldText);
                                if (string.IsNullOrEmpty(FieldName))
                                {
                                    // line not used -> disable it
                                    Input1.Text = "";
                                    InputEdit1.Text = "";
                                    InputEdit1.Enabled = false;
                                }
                                else
                                {
                                    Input1.Text = FieldName;
                                    InputEdit1.Text = FieldText;
                                    InputEdit1.Enabled = true;
                                }

                                // init edit line 2
                                TestFunc.InitLine2(ref FieldName, ref FieldText);
                                if (string.IsNullOrEmpty(FieldName))
                                {
                                    // line not used -> disable it
                                    Input2.Text = "";
                                    InputEdit2.Text = "";
                                    InputEdit2.Enabled = false;
                                }
                                else
                                {
                                    Input2.Text = FieldName;
                                    InputEdit2.Text = FieldText;
                                    InputEdit2.Enabled = true;
                                }

                                // init edit line 3
                                TestFunc.InitLine3(ref FieldName, ref FieldText);
                                if (string.IsNullOrEmpty(FieldName))
                                {
                                    // line not used -> disable it
                                    Input3.Text = "";
                                    InputEdit3.Text = "";
                                    InputEdit3.Enabled = false;
                                }
                                else
                                {
                                    Input3.Text = FieldName;
                                    InputEdit3.Text = FieldText;
                                    InputEdit3.Enabled = true;
                                }

                                // init edit line 4
                                TestFunc.InitLine4(ref FieldName, ref FieldText);
                                if (string.IsNullOrEmpty(FieldName))
                                {
                                    // line not used -> disable it
                                    Input4.Text = "";
                                    InputEdit4.Text = "";
                                    InputEdit4.Enabled = false;
                                }
                                else
                                {
                                    Input4.Text = FieldName;
                                    InputEdit4.Text = FieldText;
                                    InputEdit4.Enabled = true;
                                }

                                // init edit line 5
                                TestFunc.InitLine5(ref FieldName, ref FieldText);
                                if (string.IsNullOrEmpty(FieldName))
                                {
                                    // line not used -> disable it
                                    Input5.Text = "";
                                    InputEdit5.Text = "";
                                    InputEdit5.Enabled = false;
                                }
                                else
                                {
                                    Input5.Text = FieldName;
                                    InputEdit5.Text = FieldText;
                                    InputEdit5.Enabled = true;
                                }

                                // init edit line 6
                                TestFunc.InitLine6(ref FieldName, ref FieldText);
                                if (string.IsNullOrEmpty(FieldName))
                                {
                                    // line not used -> disable it
                                    Input6.Text = "";
                                    InputEdit6.Text = "";
                                    InputEdit6.Enabled = false;
                                }
                                else
                                {
                                    Input6.Text = FieldName;
                                    InputEdit6.Text = FieldText;
                                    InputEdit6.Enabled = true;
                                }

                                break;
                            }

                        case ProfDpCallType.Call_Execute:   // execute function
                            {
                                // check if ProfDpDrv is present
                                if (ProfDpDrv is null)
                                {
                                    DisplayText("Driver not opened");
                                    return;
                                }
                                TestFunc.RunTest();          // call test function
                                                             // unknown call type
                                break;
                            }

                        default:
                            {
                                DisplayText("Unknown Call Type");
                                break;
                            }

                    }
                    return;
                }
            }

            // function not found
            DisplayText("Function not found");
        }

        // create profdp class
        private void DpCreate_Click(object sender, EventArgs e)
        {

            try
            {
                if (ProfDpDrv is not null)
                {
                    DisplayText("Class is already created");
                    return;
                }
                ProfDpDrv = new ProfDpClass.MProfDp(DllNameEdit.Text, InitStringEdit.Text);
                ProfDpDrv.MDPEnableDPEvent(true);
                ProfDpDrv.MDPEnableDPEEvent(true);
                DisplayText("Create(): OK");
            }
            catch (CMProfDpException ex)
            {
                PrintError((int)ex.m_ErrorCode);
                return;
            }

        }

        // destroy profdp class
        private void DpDestroy_Click(object sender, EventArgs e)
        {
            // perform manual cleanup to reduce resouces hold in sub dll's
            if (ProfDpDrv is not null)
            {
                ProfDpDrv.Dispose();
                ProfDpDrv = null;
            }
            DisplayText("Destroy(): OK");
        }

        // new test function selected
        private void ComboCommands_SelectedIndexChanged(object sender, EventArgs e)
        {
            // update edit fields
            CallSelectedFunction(ComboCommands.Text, ProfDpCallType.Call_Init);

        }

        // call of test function requested
        private void ExecFunc_Click(object sender, EventArgs e)
        {

            // call test function
            CallSelectedFunction(ComboCommands.Text, ProfDpCallType.Call_Execute);

        }

        // close button
        private void CloseForm_Click(object sender, EventArgs e)
        {

            Close();

        }

        // called when form should close
        private void Form_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

            if (ProfDpDrv is not null)
            {
                ProfDpDrv.Dispose();
                ProfDpDrv = null;
            }

        }

        // callback for ProfDpEventDP
        private void EventDPCallback()
        {

            try
            {
                MDpResultBase DPResult = null;

                if (ProfDpDrv is null)
                {
                    return;
                }
                // let's read the acync result
                // normaly the enless loop is not required, because we get an event for
                // every async result, except we have disabled events for a while ...
                while (true)
                {
                    ProfDpDrv.MDPReadResult(out DPResult);
                    PrintDPResult(DPResult);
                }
            }
            catch (CMProfDpException ex)
            {
                return;
            }

        }

        // event: PROIFBUS DP operation finished
        // warning: this event is generated by a different thread
        // so we use a delegate to access the user interface
        // Always use BeginInvoke here and not Invoke to prevent a deadlock!
        private void ProfDpEventDP()
        {

            Messages.BeginInvoke(EventDPDelegate);

        }

        // callback for ProfDpEventDPCycle
        private void EventDPCycleCallback()
        {

            DisplayText("PROFIBUS DP Cycle completed!");

        }

        // event: PROFIBUS DP cycle completed
        // warning: this event is generated by a different thread
        // so we use a delegate to access the user interface
        // Always use BeginInvoke here and not Invoke to prevent a deadlock!
        private void ProfDpEventDPCycle()
        {

            Messages.BeginInvoke(EventDPCycleDelegate);

        }

        // callback for ProfDpEventDPE
        private void EventDPECallback()
        {

            try
            {
                MDpeRespBase DPEResult = null;

                if (ProfDpDrv is null)
                {
                    return;
                }

                // let's read the DPE result
                // normaly the enless loop is not required, because we get an event for
                // every DPE event, except we have disabled events for a while ...
                while (true)
                {
                    ProfDpDrv.MDPEReadResult(ref DPEResult);
                    PrintDPEResult(DPEResult);
                }
            }

            catch (CMProfDpException ex)
            {
                return;
            }

        }

        // event: DPE event occured
        // warning: this event is generated by a different thread
        // so we use a delegate to access the user interface
        // Always use BeginInvoke here and not Invoke to prevent a deadlock!
        private void ProfDpEventDPE()
        {

            Messages.BeginInvoke(EventDPEDelegate);

        }

    }
}
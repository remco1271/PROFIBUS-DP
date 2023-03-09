using System;
using System.ComponentModel;
using System.Diagnostics;

namespace ProfDpTest.My
{
    internal static partial class MyProject
    {
        internal partial class MyForms
        {

            [EditorBrowsable(EditorBrowsableState.Never)]
            public ProfDpTest m_ProfDpTest;

            public ProfDpTest ProfDpTest
            {
                [DebuggerHidden]
                get
                {
                    m_ProfDpTest = Create__Instance__(m_ProfDpTest);
                    return m_ProfDpTest;
                }
                [DebuggerHidden]
                set
                {
                    if (ReferenceEquals(value, m_ProfDpTest))
                        return;
                    if (value is not null)
                        throw new ArgumentException("Property can only be set to Nothing");
                    Dispose__Instance__(ref m_ProfDpTest);
                }
            }

        }


    }
}
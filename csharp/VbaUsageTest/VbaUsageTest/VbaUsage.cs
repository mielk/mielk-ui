using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Runtime.InteropServices;

namespace VbaUsageTest
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    //The 3 GUIDs in this file need to be unique for your COM object.
    //Generate new ones using Tools->Create GUID in VS2010
    [Guid("FEF70ECB-AB37-4605-86E1-A1EC57D33DED")]
    public interface IVbaUsage
    {
        event EventHandler Test;  
        string format(string FormatString, [Optional]object arg0, [Optional]object arg1, [Optional]object arg2, [Optional]object arg3);
        void testMethodz(string param);
    }

    //TODO: Change me!
    [Guid("2753B87A-48C2-4B20-8EB8-C507DD67E9C4")] //Change this 
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    interface ICSharpComEvents
    {
        //Add event definitions here. Add [DispId(1..n)] attributes
        //before each event declaration.
    }



    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ICSharpComEvents))]
    //TODO: Change me!
    [Guid("E63352E0-5B3C-41BD-ABB8-87FFB2F486A5")]
    public sealed class VbaUsage : IVbaUsage
    {

        private Object obj;
        public event EventHandler Test;

        public string format(string FormatString, [Optional]object arg0, [Optional]object arg1, [Optional]object arg2, [Optional]object arg3)
        {
            return string.Format(FormatString, arg0, arg1, arg2, arg3);
        }

        public void testMethodz(string param)
        {
            EventHandler handler = Test;
            if (handler != null) handler(this, EventArgs.Empty);
        }

    }
}

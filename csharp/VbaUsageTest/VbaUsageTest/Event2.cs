using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace COMVisibleEvents
{
    [ComVisible(true)]
    [Guid("8403D952-E751-4DE1-BD91-F35DEE19206E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEvents2
    {
        [DispId(1)]
        void TestFinished(string a);
    }

    [ComVisible(true)]
    [Guid("2BF7DB6B-DDB3-42A5-BD65-92EE93ABB473")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDemoEvents2
    {
        [DispId(1)]
        void TestEvent(string a);
    }

    [ComVisible(true)]
    [Guid("56D41646-10CB-4188-979D-23F70E0FFDF5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEvents2))]
    [ProgId("COMVisibleEvents.DemoEvents")]
    public class DemoEvents2 : IDemoEvents2
    {
        public delegate void TestFinishedDelegate(string a);
        private event TestFinishedDelegate TestFinished;

        public void TestEvent(string a)
        {
            TestFinished(a + 'b');
        }

    }
}
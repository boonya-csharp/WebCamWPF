using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Drawing;
using System.Security;

namespace WebCamWPF
{
    public static class WindowStyles
    {
        public static int WS_CHILD = 0x40000000;
        public static int WS_CLIPCHILDREN = 0x02000000;
    }

    public static class Clsid
    {
        public static readonly Guid SystemDeviceEnum =
            new Guid(0x62BE5D10, 0x60EB, 0x11D0, 0xBD, 0x3B, 0x00, 0xA0, 0xC9, 0x11, 0xCE, 0x86);

        public static readonly Guid FilterGraph =
            new Guid(0xE436EBB3, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid SampleGrabber =
            new Guid(0xC1F400A0, 0x3F08, 0x11D3, 0x9F, 0x0B, 0x00, 0x60, 0x08, 0x03, 0x9E, 0x37);

        public static readonly Guid CaptureGraphBuilder2 =
            new Guid(0xBF87B6E1, 0x8C27, 0x11D0, 0xB3, 0xF0, 0x00, 0xAA, 0x00, 0x37, 0x61, 0xC5);

        public static readonly Guid AsyncReader =
            new Guid(0xE436EBB5, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid VideoInputDeviceCategory = new Guid("860BB310-5D01-11d0-BD3B-00A0C911CE86");
    }

    public static class MediaType
    {
        public static readonly Guid Video =
            new Guid(0x73646976, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);

        public static readonly Guid Interleaved =
            new Guid(0x73766169, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);

        public static readonly Guid Audio =
            new Guid(0x73647561, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);

        public static readonly Guid Text =
            new Guid(0x73747874, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);

        public static readonly Guid Stream =
            new Guid(0xE436EB83, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);
    }

    public static class MediaSubType
    {
        public static readonly Guid YUYV =
            new Guid(0x56595559, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);

        public static readonly Guid IYUV =
            new Guid(0x56555949, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);

        public static readonly Guid DVSD =
            new Guid(0x44535644, 0x0000, 0x0010, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71);

        public static readonly Guid RGB1 =
            new Guid(0xE436EB78, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid RGB4 =
            new Guid(0xE436EB79, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid RGB8 =
            new Guid(0xE436EB7A, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid RGB565 =
            new Guid(0xE436EB7B, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid RGB555 =
            new Guid(0xE436EB7C, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid RGB24 =
            new Guid(0xE436Eb7D, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid RGB32 =
            new Guid(0xE436EB7E, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid Avi =
            new Guid(0xE436EB88, 0x524F, 0x11CE, 0x9F, 0x53, 0x00, 0x20, 0xAF, 0x0B, 0xA7, 0x70);

        public static readonly Guid Asf =
            new Guid(0x3DB80F90, 0x9412, 0x11D1, 0xAD, 0xED, 0x00, 0x00, 0xF8, 0x75, 0x4B, 0x99);
    }

    internal class PinCategory
    {
        public static readonly Guid Capture =
            new Guid(0xFB6C4281, 0x0353, 0x11D1, 0x90, 0x5F, 0x00, 0x00, 0xC0, 0xCC, 0x16, 0xBA);
        public static readonly Guid Preview = new Guid("fb6c4282-0353-11d1-905f-0000c0cc16ba");

    }

    public enum DsEvCode
    {
        None,
        Complete = 0x01,      // EC_COMPLETE
        DeviceLost = 0x1F,    // EC_DEVICE_LOST
    }

    [ComImport]
    [Guid("56a868c0-0ad4-11ce-b03a-0020af0ba770")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMediaEventEx
    {
        [PreserveSig]
        int GetEventHandle(out IntPtr hEvent);

        [PreserveSig]
        int GetEvent([Out] [MarshalAs(UnmanagedType.I4)] out DsEvCode lEventCode, [Out] out IntPtr lParam1, [Out] out IntPtr lParam2, int msTimeout);

        [PreserveSig]
        int WaitForCompletion(int msTimeout, [Out] out int pEvCode);

        [PreserveSig]
        int CancelDefaultHandling(int lEvCode);

        [PreserveSig]
        int RestoreDefaultHandling(int lEvCode);

        [PreserveSig]
        int FreeEventParams([In] [MarshalAs(UnmanagedType.I4)] DsEvCode lEvCode, IntPtr lParam1, IntPtr lParam2);

        [PreserveSig]
        int SetNotifyWindow(IntPtr hwnd, int lMsg, IntPtr lInstanceData);

        [PreserveSig]
        int SetNotifyFlags(int lNoNotifyFlags);

        [PreserveSig]
        int GetNotifyFlags(out int lplNoNotifyFlags);
    }

    [ComImport]
    [Guid("56A868B1-0AD4-11CE-B03A-0020AF0BA770")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMediaControl
    {
        [PreserveSig]
        int Run();

        [PreserveSig]
        int Pause();

        [PreserveSig]
        int Stop();

        [PreserveSig]
        int GetState(int timeout, out int filterState);

        [PreserveSig]
        int RenderFile(string fileName);

        [PreserveSig]
        int AddSourceFilter([In] string fileName, [Out] [MarshalAs(UnmanagedType.IDispatch)] out object filterInfo);

        [PreserveSig]
        int get_FilterCollection([Out] [MarshalAs(UnmanagedType.IDispatch)] out object collection);

        [PreserveSig]
        int get_RegFilterCollection([Out] [MarshalAs(UnmanagedType.IDispatch)] out object collection);

        [PreserveSig]
        int StopWhenReady();
    }

    [ComImport]
    [Guid("56A868B4-0AD4-11CE-B03A-0020AF0BA770")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IVideoWindow
    {
        [PreserveSig]
        int put_Caption(string caption);

        [PreserveSig]
        int get_Caption([Out] out string caption);

        [PreserveSig]
        int put_WindowStyle(int windowStyle);

        [PreserveSig]
        int get_WindowStyle(out int windowStyle);

        [PreserveSig]
        int put_WindowStyleEx(int windowStyleEx);

        [PreserveSig]
        int get_WindowStyleEx(out int windowStyleEx);

        [PreserveSig]
        int put_AutoShow([In] [MarshalAs(UnmanagedType.Bool)] bool autoShow);

        [PreserveSig]
        int get_AutoShow([Out] [MarshalAs(UnmanagedType.Bool)] out bool autoShow);

        [PreserveSig]
        int put_WindowState(int windowState);

        [PreserveSig]
        int get_WindowState(out int windowState);

        [PreserveSig]
        int put_BackgroundPalette([In] [MarshalAs(UnmanagedType.Bool)] bool backgroundPalette);

        [PreserveSig]
        int get_BackgroundPalette([Out] [MarshalAs(UnmanagedType.Bool)] out bool backgroundPalette);

        [PreserveSig]
        int put_Visible([In] [MarshalAs(UnmanagedType.I4)] int visible);

        [PreserveSig]
        int get_Visible([Out] [MarshalAs(UnmanagedType.I4)] out int visible);

        [PreserveSig]
        int put_Left(int left);

        [PreserveSig]
        int get_Left(out int left);

        [PreserveSig]
        int put_Width(int width);

        [PreserveSig]
        int get_Width(out int width);

        [PreserveSig]
        int put_Top(int top);

        [PreserveSig]
        int get_Top(out int top);

        [PreserveSig]
        int put_Height(int height);

        [PreserveSig]
        int get_Height(out int height);

        [PreserveSig]
        int put_Owner(IntPtr owner);

        [PreserveSig]
        int get_Owner(out IntPtr owner);
        
        [PreserveSig]
        int put_MessageDrain(IntPtr drain);

        [PreserveSig]
        int get_MessageDrain(out IntPtr drain);

        [PreserveSig]
        int get_BorderColor(out int color);

        [PreserveSig]
        int put_BorderColor(int color);

        [PreserveSig]
        int get_FullScreenMode([Out] [MarshalAs(UnmanagedType.Bool)] out bool fullScreenMode);

        [PreserveSig]
        int put_FullScreenMode([In] [MarshalAs(UnmanagedType.Bool)] bool fullScreenMode);

        [PreserveSig]
        int SetWindowForeground(int focus);

        [PreserveSig]
        int NotifyOwnerMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        [PreserveSig]
        int SetWindowPosition(int left, int top, int width, int height);

        [PreserveSig]
        int GetWindowPosition(out int left, out int top, out int width, out int height);

        [PreserveSig]
        int GetMinIdealImageSize(out int width, out int height);

        [PreserveSig]
        int GetMaxIdealImageSize(out int width, out int height);

        [PreserveSig]
        int GetRestorePosition(out int left, out int top, out int width, out int height);

        [PreserveSig]
        int HideCursor([In] [MarshalAs(UnmanagedType.Bool)] bool hideCursor);

        [PreserveSig]
        int IsCursorHidden([Out] [MarshalAs(UnmanagedType.Bool)] out bool hideCursor);
    }

    public class FilterInfo : IComparable
    {
        [DllImport("ole32.dll")]
        public static extern
        int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        public static extern int MkParseDisplayName(IBindCtx pbc, string szUserName, ref int pchEaten, out IMoniker ppmk);

        public readonly string Name;
        public readonly string MonikerString;
        
        public FilterInfo(string monikerString)
        {
            MonikerString = monikerString;
            Name = GetName(monikerString);
        }

        internal FilterInfo(IMoniker moniker)
        {
            MonikerString = GetMonikerString(moniker);
            Name = GetName(moniker);
        }

        public int CompareTo(object value)
        {
            FilterInfo f = (FilterInfo)value;

            if (f == null)
                return 1;

            return (this.Name.CompareTo(f.Name));
        }

        public static object CreateFilter(string filterMoniker)
        {
            object filterObject = null;
            IBindCtx bindCtx = null;
            IMoniker moniker = null;

            int n = 0;
            if (CreateBindCtx(0, out bindCtx) == 0)
            {
                if (MkParseDisplayName(bindCtx, filterMoniker, ref n, out moniker) == 0)
                {
                    Guid filterId = typeof(IBaseFilter).GUID;
                    moniker.BindToObject(null, null, ref filterId, out filterObject);

                    Marshal.ReleaseComObject(moniker);
                }
                Marshal.ReleaseComObject(bindCtx);
            }
            return filterObject;
        }

        private string GetMonikerString(IMoniker moniker)
        {
            string str;
            moniker.GetDisplayName(null, null, out str);
            return str;
        }

        private string GetName(IMoniker moniker)
        {
            Object bagObj = null;
            IPropertyBag bag = null;

            try
            {
                Guid bagId = typeof(IPropertyBag).GUID;
                moniker.BindToStorage(null, null, ref bagId, out bagObj);
                bag = (IPropertyBag)bagObj;

                object val = "";
                int hr = bag.Read("FriendlyName", ref val, IntPtr.Zero);
                if (hr != 0)
                    Marshal.ThrowExceptionForHR(hr);

                string ret = (string)val;
                if ((ret == null) || (ret.Length < 1))
                    throw new ApplicationException();

                return ret;
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                bag = null;
                if (bagObj != null)
                {
                    Marshal.ReleaseComObject(bagObj);
                    bagObj = null;
                }
            }
        }

        private string GetName(string monikerString)
        {
            IBindCtx bindCtx = null;
            IMoniker moniker = null;
            String name = "";
            int n = 0;

            if (CreateBindCtx(0, out bindCtx) == 0)
            {
                if (MkParseDisplayName(bindCtx, monikerString, ref n, out moniker) == 0)
                {
                    name = GetName(moniker);

                    Marshal.ReleaseComObject(moniker);
                    moniker = null;
                }
                Marshal.ReleaseComObject(bindCtx);
                bindCtx = null;
            }
            return name;
        }
    }

    [ComVisible(false)]
    public enum PinDirection
    {
        Input,
        Output
    }

    [ComVisible(false)]
    [StructLayout(LayoutKind.Sequential)]
    public class AMMediaType : IDisposable
    {
        public Guid MajorType;
        public Guid SubType;
        [MarshalAs(UnmanagedType.Bool)]
        public bool FixedSizeSamples = true;
        [MarshalAs(UnmanagedType.Bool)]
        public bool TemporalCompression;
        public int SampleSize = 1;
        public Guid FormatType;
        public IntPtr unkPtr;
        public int FormatSize;
        public IntPtr FormatPtr;
        
        ~AMMediaType()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (FormatSize != 0)
                Marshal.FreeCoTaskMem(FormatPtr);
            if (unkPtr != IntPtr.Zero)
                Marshal.Release(unkPtr);
        }
    }

    [ComVisible(false)]
    [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]
    public class PinInfo
    {
        public IBaseFilter Filter;
        public PinDirection Direction;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string Name;
    }

    [ComVisible(false)]
    [StructLayout(LayoutKind.Sequential)]
    public struct VideoInfoHeader
    {
        public RECT SrcRect;
        public RECT TargetRect;
        public int BitRate;
        public int BitErrorRate;
        public long AverageTimePerFrame;
        public BitmapInfoHeader BmiHeader;
    }

    [ComVisible(false)]
    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    public struct BitmapInfoHeader
    {
        public int Size;
        public int Width;
        public int Height;
        public short Planes;
        public short BitCount;
        public int Compression;
        public int ImageSize;
        public int XPelsPerMeter;
        public int YPelsPerMeter;
        public int ColorsUsed;
        public int ColorsImportant;
    }

    [ComVisible(false)]
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [ComImport]
    [Guid("56A86891-0AD4-11CE-B03A-0020AF0BA770")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPin
    {
        [PreserveSig]
        int Connect([In] IPin receivePin, [In, MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);
        [PreserveSig]
        int ReceiveConnection([In] IPin receivePin, [In, MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);
        [PreserveSig]
        int Disconnect();
        [PreserveSig]
        int ConnectedTo([Out] out IPin pin);
        [PreserveSig]
        int ConnectionMediaType([Out, MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);
        [PreserveSig]
        int QueryPinInfo([Out, MarshalAs(UnmanagedType.LPStruct)] PinInfo pinInfo);
        [PreserveSig]
        int QueryDirection(out PinDirection pinDirection);
        [PreserveSig]
        int QueryId([Out, MarshalAs(UnmanagedType.LPWStr)] out string id);
        [PreserveSig]
        int QueryAccept([In, MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);
        [PreserveSig]
        int EnumMediaTypes(IntPtr enumerator);
        [PreserveSig]
        int QueryInternalConnections(IntPtr apPin, [In, Out] ref int nPin);
        [PreserveSig]
        int EndOfStream();
        [PreserveSig]
        int BeginFlush();
        [PreserveSig]
        int EndFlush();
        [PreserveSig]
        int NewSegment(long start, long stop, double rate);
    }

    [ComImport]
    [Guid("56A86892-0AD4-11CE-B03A-0020AF0BA770")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumPins
    {
        [PreserveSig]
        int Next([In] int cPins, [Out] [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IPin[] pins, [Out] out int pinsFetched);

        [PreserveSig]
        int Skip([In] int cPins);

        [PreserveSig]
        int Reset();

        [PreserveSig]
        int Clone([Out] out IEnumPins enumPins);
    }

    [ComImport]
    [Guid("56A868A6-0Ad4-11CE-B03A-0020AF0BA770")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileSourceFilter
    {
        [PreserveSig]
        int Load([In] [MarshalAs(UnmanagedType.LPWStr)] string fileName, [In] [MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);

        [PreserveSig]
        int GetCurFile([Out] [MarshalAs(UnmanagedType.LPWStr)] out string fileName, [Out] [MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);
    }

    [ComImport]
    [Guid("56A8689F-0AD4-11CE-B03A-0020AF0BA770")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFilterGraph
    {
        [PreserveSig]
        int AddFilter([In] IBaseFilter filter, [In] [MarshalAs(UnmanagedType.LPWStr)] string name);

        [PreserveSig]
        int RemoveFilter([In] IBaseFilter filter);

        [PreserveSig]
        int EnumFilters([Out] out IntPtr enumerator);

        [PreserveSig]
        int FindFilterByName([In, MarshalAs(UnmanagedType.LPWStr)] string name, [Out] out IBaseFilter filter);

        [PreserveSig]
        int ConnectDirect([In] IPin pinOut, [In] IPin pinIn, [In, MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);

        [PreserveSig]
        int Reconnect([In] IPin pin);

        [PreserveSig]
        int Disconnect([In] IPin pin);

        [PreserveSig]
        int SetDefaultSyncSource();
    }

    [ComImport]
    [Guid("56A86895-0AD4-11CE-B03A-0020AF0BA770")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IBaseFilter
    {
        [PreserveSig]
        int GetClassID([Out] out Guid ClassID);

        [PreserveSig]
        int Stop();

        [PreserveSig]
        int Pause();

        [PreserveSig]
        int Run(long start);

        [PreserveSig]
        int GetState(int milliSecsTimeout, [Out] out int filterState);

        [PreserveSig]
        int SetSyncSource([In] IntPtr clock);

        [PreserveSig]
        int GetSyncSource([Out] out IntPtr clock);

        [PreserveSig]
        int EnumPins([Out] out IEnumPins enumPins);

        [PreserveSig]
        int FindPin([In, MarshalAs(UnmanagedType.LPWStr)] string id, [Out] out IPin pin);

        [PreserveSig]
        int QueryFilterInfo([Out] FilterInfo filterInfo);

        [PreserveSig]
        int JoinFilterGraph([In] IFilterGraph graph, [In, MarshalAs(UnmanagedType.LPWStr)] string name);

        [PreserveSig]
        int QueryVendorInfo([Out, MarshalAs(UnmanagedType.LPWStr)] out string vendorInfo);
    }

    [ComImport]
    [Guid("55272A00-42CB-11CE-8135-00AA004BB851")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyBag
    {
        [PreserveSig]
        int Read([MarshalAs(UnmanagedType.LPWStr)] string name, ref object value, object error);
        [PreserveSig]
        int Write([MarshalAs(UnmanagedType.LPWStr)] string name, object value);
    }

    [ComImport]
    [Guid("29840822-5B84-11D0-BD3B-00A0C911CE86")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ICreateDevEnum
    {
        [PreserveSig]
        int CreateClassEnumerator( [In] ref Guid type, [Out] out IEnumMoniker enumMoniker, [In] int flags );
    }

    [ComImport]
    [Guid("62BE5D10-60EB-11d0-BD3B-00A0C911CE86")]
    public class CreateDevEnum
    {
    }

    [ComImport]
    [Guid("56a868a9-0ad4-11ce-b03a-0020af0ba770")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IGraphBuilder
    {
        [PreserveSig]
        int AddFilter([In] IBaseFilter filter, [In] [MarshalAs(UnmanagedType.LPWStr)] string name);

        [PreserveSig]
        int RemoveFilter([In] IBaseFilter filter);

        [PreserveSig]
        int EnumFilters([Out] out IntPtr enumerator);

        [PreserveSig]
        int FindFilterByName([In] [MarshalAs(UnmanagedType.LPWStr)] string name, [Out] out IBaseFilter filter);

        [PreserveSig]
        int ConnectDirect([In] IPin pinOut, [In] IPin pinIn, [In] [MarshalAs(UnmanagedType.LPStruct)] AMMediaType mediaType);

        [PreserveSig]
        int Reconnect([In] IPin pin);

        [PreserveSig]
        int Disconnect([In] IPin pin);

        [PreserveSig]
        int SetDefaultSyncSource();

        [PreserveSig]
        int Connect([In] IPin pinOut, [In] IPin pinIn);

        [PreserveSig]
        int Render([In] IPin pinOut);

        [PreserveSig]
        int RenderFile([In] [MarshalAs(UnmanagedType.LPWStr)] string file, [In] [MarshalAs(UnmanagedType.LPWStr)] string playList);

        [PreserveSig]
        int AddSourceFilter([In] [MarshalAs(UnmanagedType.LPWStr)] string fileName, [In] [MarshalAs(UnmanagedType.LPWStr)] string filterName, [Out] out IBaseFilter filter);

        [PreserveSig]
        int SetLogFile(IntPtr hFile);

        [PreserveSig]
        int Abort();

        [PreserveSig]
        int ShouldOperationContinue();
    }

    [ComImport]
    [Guid("93E5A4E0-2D50-11d2-ABFA-00A0C9C6E38D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ICaptureGraphBuilder2
    {
        [PreserveSig]
        int SetFiltergraph([In] IGraphBuilder graphBuilder);

        [PreserveSig]
        int GetFiltergraph([Out] out IGraphBuilder graphBuilder);

        [PreserveSig]
        int SetOutputFileName([In] [MarshalAs(UnmanagedType.LPStruct)] Guid type, [In] [MarshalAs(UnmanagedType.LPWStr)] string fileName, [Out] out IBaseFilter baseFilter, [Out] out IntPtr fileSinkFilter);

        [PreserveSig]
        int FindInterface([In] [MarshalAs(UnmanagedType.LPStruct)] Guid category, [In] [MarshalAs(UnmanagedType.LPStruct)] Guid type, [In] IBaseFilter baseFilter, [In] [MarshalAs(UnmanagedType.LPStruct)] Guid interfaceID, [Out] [MarshalAs(UnmanagedType.IUnknown)] out object retInterface);

        [PreserveSig]
        int RenderStream([In] [MarshalAs(UnmanagedType.LPStruct)] Guid category, [In] [MarshalAs(UnmanagedType.LPStruct)] Guid mediaType, [In] [MarshalAs(UnmanagedType.IUnknown)] object source, [In] IBaseFilter compressor, [In] IBaseFilter renderer);

        [PreserveSig]
        int ControlStream([In] [MarshalAs(UnmanagedType.LPStruct)] Guid category, [In] [MarshalAs(UnmanagedType.LPStruct)] Guid mediaType, [In] [MarshalAs(UnmanagedType.Interface)] IBaseFilter filter, [In] long start, [In] long stop, [In] short startCookie, [In] short stopCookie);

        [PreserveSig]
        int AllocCapFile([In] [MarshalAs(UnmanagedType.LPWStr)] string fileName, [In] long size);

        [PreserveSig]
        int CopyCaptureFile([In] [MarshalAs(UnmanagedType.LPWStr)] string oldFileName, [In] [MarshalAs(UnmanagedType.LPWStr)] string newFileName, [In] [MarshalAs(UnmanagedType.Bool)] bool allowEscAbort, [In] IntPtr callback);

        [PreserveSig]
        int FindPin([In] [MarshalAs(UnmanagedType.IUnknown)] object source, [In] PinDirection pinDirection, [In] [MarshalAs(UnmanagedType.LPStruct)] Guid category, [In] [MarshalAs(UnmanagedType.LPStruct)] Guid mediaType, [In] [MarshalAs(UnmanagedType.Bool)] bool unconnected, [In] int index, [Out] [MarshalAs(UnmanagedType.Interface)] out IPin pin);
    }

    [ComImport]
    [Guid("BF87B6E1-8C27-11d0-B3F0-00AA003761C5")]
    public class CaptureGraphBuilder
    {
    }

    [ComImport]
    [Guid("e436ebb3-524f-11ce-9f53-0020af0ba770")]
    public class FilterGraph
    {
    }

    [ComImport]
    [Guid("6B652FFF-11FE-4fce-92AD-0266B5D7C78F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISampleGrabber
    {
        [PreserveSig]
        int SetOneShot([In, MarshalAs(UnmanagedType.Bool)] bool OneShot);

        [PreserveSig]
        int SetMediaType([In, MarshalAs(UnmanagedType.LPStruct)] AMMediaType pmt);

        [PreserveSig]
        int GetConnectedMediaType([Out, MarshalAs(UnmanagedType.LPStruct)] AMMediaType pmt);

        [PreserveSig]
        int SetBufferSamples([In, MarshalAs(UnmanagedType.Bool)] bool BufferThem);

        [PreserveSig]
        int GetCurrentBuffer(ref int pBufferSize, IntPtr pBuffer);

        [PreserveSig]
        int GetCurrentSample(out IMediaSample ppSample);

        [PreserveSig]
        int SetCallback(ISampleGrabberCB pCallback, int WhichMethodToCallback);
    }

    [ComImport]
    [Guid("0579154A-2B53-4994-B0D0-E773148EFF85")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISampleGrabberCB
    {
        [PreserveSig]
        int SampleCB(double SampleTime, IMediaSample pSample);

        [PreserveSig]
        int BufferCB(double SampleTime, IntPtr pBuffer, int BufferLen);
    }

    [ComImport]
    [Guid("56a8689a-0ad4-11ce-b03a-0020af0ba770")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMediaSample
    {
        [PreserveSig]
        int GetPointer([Out] out IntPtr ppBuffer); // BYTE **

        [PreserveSig]
        int GetSize();

        [PreserveSig]
        int GetTime([Out] out long pTimeStart, [Out] out long pTimeEnd);

        [PreserveSig]
        int SetTime([In, MarshalAs(UnmanagedType.LPStruct)] long? pTimeStart, [In, MarshalAs(UnmanagedType.LPStruct)] long? pTimeEnd);

        [PreserveSig]
        int IsSyncPoint();

        [PreserveSig]
        int SetSyncPoint([In, MarshalAs(UnmanagedType.Bool)] bool bIsSyncPoint);

        [PreserveSig]
        int IsPreroll();

        [PreserveSig]
        int SetPreroll([In, MarshalAs(UnmanagedType.Bool)] bool bIsPreroll);

        [PreserveSig]
        int GetActualDataLength();

        [PreserveSig]
        int SetActualDataLength([In] int len);

        [PreserveSig]
        int GetMediaType([Out, MarshalAs(UnmanagedType.LPStruct)] out AMMediaType ppMediaType);

        [PreserveSig]
        int SetMediaType([In, MarshalAs(UnmanagedType.LPStruct)] AMMediaType pMediaType);

        [PreserveSig]
        int IsDiscontinuity();

        [PreserveSig]
        int SetDiscontinuity([In, MarshalAs(UnmanagedType.Bool)] bool bDiscontinuity);

        [PreserveSig]
        int GetMediaTime([Out] out long pTimeStart, [Out] out long pTimeEnd);

        [PreserveSig]
        int SetMediaTime([In, MarshalAs(UnmanagedType.LPStruct)] long? pTimeStart, [In, MarshalAs(UnmanagedType.LPStruct)] long? pTimeEnd);
    }

    public static class FormatType
    {
        public static readonly Guid Null = Guid.Empty;

        public static readonly Guid None = new Guid(0x0F6417D6, 0xc318, 0x11d0, 0xa4, 0x3f, 0x00, 0xa0, 0xc9, 0x22, 0x31, 0x96);
        public static readonly Guid VideoInfo = new Guid(0x05589f80, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static readonly Guid VideoInfo2 = new Guid(0xf72a76A0, 0xeb0a, 0x11d0, 0xac, 0xe4, 0x00, 0x00, 0xc0, 0xcc, 0x16, 0xba);
        public static readonly Guid WaveEx = new Guid(0x05589f81, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static readonly Guid MpegVideo = new Guid(0x05589f82, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static readonly Guid MpegStreams = new Guid(0x05589f83, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static readonly Guid DvInfo = new Guid(0x05589f84, 0xc356, 0x11ce, 0xbf, 0x01, 0x00, 0xaa, 0x00, 0x55, 0x59, 0x5a);
        public static readonly Guid AnalogVideo = new Guid(0x0482dde0, 0x7817, 0x11cf, 0x8a, 0x03, 0x00, 0xaa, 0x00, 0x6e, 0xcb, 0x65);
        public static readonly Guid Mpeg2Video = new Guid(0xe06d80e3, 0xdb46, 0x11cf, 0xb4, 0xd1, 0x00, 0x80, 0x5f, 0x6c, 0xbb, 0xea);
        public static readonly Guid DolbyAC3 = new Guid(0xe06d80e4, 0xdb46, 0x11cf, 0xb4, 0xd1, 0x00, 0x80, 0x5f, 0x6c, 0xbb, 0xea);
        public static readonly Guid Mpeg2Audio = new Guid(0xe06d80e5, 0xdb46, 0x11cf, 0xb4, 0xd1, 0x00, 0x80, 0x5f, 0x6c, 0xbb, 0xea);
        public static readonly Guid WSS525 = new Guid(0xc7ecf04d, 0x4582, 0x4869, 0x9a, 0xbb, 0xbf, 0xb5, 0x23, 0xb6, 0x2e, 0xdf);
        public static readonly Guid ETDTFilter_Tagged = new Guid(0xC4C4C4D1, 0x0049, 0x4E2B, 0x98, 0xFB, 0x95, 0x37, 0xF6, 0xCE, 0x51, 0x6D);
        public static readonly Guid CPFilters_Processed = new Guid(0x6739b36f, 0x1d5f, 0x4ac2, 0x81, 0x92, 0x28, 0xbb, 0xe, 0x73, 0xd1, 0x6a);
    }

    public class DsRect
    {
        public int left;

        public int top;

        public int right;

        public int bottom;

        /// <summary>
        /// Empty contructor. Initialize all fields to 0
        /// </summary>
        public DsRect()
        {
            this.left = 0;
            this.top = 0;
            this.right = 0;
            this.bottom = 0;
        }

        /// <summary>
        /// A parametred constructor. Initialize fields with given values.
        /// </summary>
        /// <param name="left">the left value</param>
        /// <param name="top">the top value</param>
        /// <param name="right">the right value</param>
        /// <param name="bottom">the bottom value</param>
        public DsRect(int left, int top, int right, int bottom)
        {
            this.left = left;
            this.top = top;
            this.right = right;
            this.bottom = bottom;
        }

        /// <summary>
        /// A parametred constructor. Initialize fields with a given <see cref="System.Drawing.Rectangle"/>.
        /// </summary>
        /// <param name="rectangle">A <see cref="System.Drawing.Rectangle"/></param>
        /// <remarks>
        /// Warning, DsRect define a rectangle by defining two of his corners and <see cref="System.Drawing.Rectangle"/> define a rectangle with his upper/left corner, his width and his height.
        /// </remarks>
        public DsRect(Rectangle rectangle)
        {
            this.left = rectangle.Left;
            this.top = rectangle.Top;
            this.right = rectangle.Right;
            this.bottom = rectangle.Bottom;
        }

        /// <summary>
        /// Provide de string representation of this DsRect instance
        /// </summary>
        /// <returns>A string formated like this : [left, top - right, bottom]</returns>
        public override string ToString()
        {
            return string.Format("[{0}, {1} - {2}, {3}]", this.left, this.top, this.right, this.bottom);
        }

        public override int GetHashCode()
        {
            return this.left.GetHashCode() | this.top.GetHashCode() | this.right.GetHashCode()
                   | this.bottom.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (obj is DsRect)
            {
                DsRect cmp = (DsRect)obj;

                return right == cmp.right && bottom == cmp.bottom && left == cmp.left && top == cmp.top;
            }

            if (obj is Rectangle)
            {
                Rectangle cmp = (Rectangle)obj;

                return right == cmp.Right && bottom == cmp.Bottom && left == cmp.Left && top == cmp.Top;
            }

            return false;
        }

        /// <summary>
        /// Define implicit cast between DirectShowLib.DsRect and System.Drawing.Rectangle for languages supporting this feature.
        /// VB.Net doesn't support implicit cast. <see cref="DirectShowLib.DsRect.ToRectangle"/> for similar functionality.
        /// <code>
        ///   // Define a new Rectangle instance
        ///   Rectangle r = new Rectangle(0, 0, 100, 100);
        ///   // Do implicit cast between Rectangle and DsRect
        ///   DsRect dsR = r;
        ///
        ///   Console.WriteLine(dsR.ToString());
        /// </code>
        /// </summary>
        /// <param name="r">a DsRect to be cast</param>
        /// <returns>A casted System.Drawing.Rectangle</returns>
        public static implicit operator Rectangle(DsRect r)
        {
            return r.ToRectangle();
        }

        public static implicit operator RECT(DsRect r)
        {
            return r.ToRECT();
        }

        /// <summary>
        /// Define implicit cast between System.Drawing.Rectangle and DirectShowLib.DsRect for languages supporting this feature.
        /// VB.Net doesn't support implicit cast. <see cref="DirectShowLib.DsRect.FromRectangle"/> for similar functionality.
        /// <code>
        ///   // Define a new DsRect instance
        ///   DsRect dsR = new DsRect(0, 0, 100, 100);
        ///   // Do implicit cast between DsRect and Rectangle
        ///   Rectangle r = dsR;
        ///
        ///   Console.WriteLine(r.ToString());
        /// </code>
        /// </summary>
        /// <param name="r">A System.Drawing.Rectangle to be cast</param>
        /// <returns>A casted DsRect</returns>
        public static implicit operator DsRect(Rectangle r)
        {
            return new DsRect(r);
        }

        /// <summary>
        /// Get the System.Drawing.Rectangle equivalent to this DirectShowLib.DsRect instance.
        /// </summary>
        /// <returns>A System.Drawing.Rectangle</returns>
        public Rectangle ToRectangle()
        {
            return new Rectangle(this.left, this.top, (this.right - this.left), (this.bottom - this.top));
        }

        public RECT ToRECT()
        {
            return new RECT()
            {
                Left = this.left,
                Top = this.top,
                Right = this.right,
                Bottom = this.bottom
            };
        }

        /// <summary>
        /// Get a new DirectShowLib.DsRect instance for a given <see cref="System.Drawing.Rectangle"/>
        /// </summary>
        /// <param name="r">The <see cref="System.Drawing.Rectangle"/> used to initialize this new DirectShowLib.DsGuid</param>
        /// <returns>A new instance of DirectShowLib.DsGuid</returns>
        public static DsRect FromRectangle(Rectangle r)
        {
            return new DsRect(r);
        }
    }

    [ComImport, SuppressUnmanagedCodeSecurity,
    Guid("C6E13340-30AC-11d0-A18C-00A0C9118956"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAMStreamConfig
    {
        [PreserveSig]
        int SetFormat([In, MarshalAs(UnmanagedType.LPStruct)] AMMediaType pmt);

        [PreserveSig]
        int GetFormat([Out] out AMMediaType pmt);

        [PreserveSig]
        int GetNumberOfCapabilities(out int piCount, out int piSize);

        [PreserveSig]
        int GetStreamCaps(
            [In] int iIndex,
            [Out] out AMMediaType ppmt,
            [In] IntPtr pSCC
            );
    }
}

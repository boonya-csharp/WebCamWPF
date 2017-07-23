using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices.ComTypes;
using System.ComponentModel;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using System.Threading;
using System.Windows.Forms.Integration;
using System.Net;
using Newtonsoft.Json;
using System.Dynamic;

namespace WebCamWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private IGraphBuilder graph;
        private ICaptureGraphBuilder2 capture;
        private SampleGrabberCallback callback;

        public MainWindow()
        {            
            InitializeComponent();           
        }

        private void CaptureVideo()
        {
            int retVal;

            graph = (IGraphBuilder)new FilterGraph();
            capture = (ICaptureGraphBuilder2)new CaptureGraphBuilder();
            
            IMediaControl control = (IMediaControl)graph;
            IMediaEventEx eventEx = (IMediaEventEx)graph;
            
            retVal = capture.SetFiltergraph(graph);

            Dictionary<string, IMoniker> devices = EnumDevices(Clsid.VideoInputDeviceCategory);
            IMoniker moniker = devices.First().Value;
            object obj = null;
            moniker.BindToObject(null, null, typeof(IBaseFilter).GUID, out obj);

            IBaseFilter baseFilter = (IBaseFilter)obj;
            retVal = graph.AddFilter(baseFilter, devices.First().Key);

            Guid CLSID_SampleGrabber = new Guid("{C1F400A0-3F08-11D3-9F0B-006008039E37}");
            IBaseFilter grabber = Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_SampleGrabber)) as IBaseFilter;
            
            var media = new AMMediaType();
            media.MajorType = MediaType.Video;
            media.SubType = MediaSubType.RGB24;
            media.FormatType = FormatType.VideoInfo;
            retVal = ((ISampleGrabber)grabber).SetMediaType(media);

            object configObj;
            retVal = capture.FindInterface(PinCategory.Capture, MediaType.Video, baseFilter, typeof(IAMStreamConfig).GUID, out configObj);
            IAMStreamConfig config = (IAMStreamConfig)configObj;
            
            AMMediaType pmt;
            retVal = config.GetFormat(out pmt);

            var header = (VideoInfoHeader)Marshal.PtrToStructure(pmt.FormatPtr, typeof(VideoInfoHeader));
            var width = header.BmiHeader.Width;
            var height = header.BmiHeader.Height;
            var stride = width * (header.BmiHeader.BitCount / 8);
            callback = new SampleGrabberCallback() { Width = width, Height = height, Stride = stride };
            retVal = ((ISampleGrabber)grabber).SetCallback(callback, 0);

            retVal = graph.AddFilter(grabber, "SampleGrabber");

            IPin output = GetPin(baseFilter, p => p.Name == "Capture");
            IPin input = GetPin(grabber, p => p.Name == "Input");
            IPin preview = GetPin(grabber, p => p.Name == "Output");            
            //retVal = graph.ConnectDirect(output, input, pmt);
            //retVal = graph.Connect(output, input);

            retVal = capture.RenderStream(PinCategory.Preview, MediaType.Video, baseFilter, grabber, null);


            //var wih = new WindowInteropHelper(this);
            var panel = FindName("PART_VideoPanel") as System.Windows.Forms.Panel;

            
            IVideoWindow window = (IVideoWindow)graph;
            retVal = window.put_Owner(panel.Handle);
            retVal = window.put_WindowStyle(WindowStyles.WS_CHILD | WindowStyles.WS_CLIPCHILDREN);
            retVal = window.SetWindowPosition(0, 0, (int)panel.ClientSize.Width, (int)panel.ClientSize.Height);
            retVal = window.put_MessageDrain(panel.Handle);
            retVal = window.put_Visible(-1); //OATRUE

            retVal = control.Run();
        }

        private IPin GetPin(IBaseFilter filter, Func<PinInfo, bool> pred)
        {
            IEnumPins enumPins;
            var retVal = filter.EnumPins(out enumPins);

            IPin[] pins = new IPin[1];
            int fetched;
            while ((retVal = enumPins.Next(1, pins, out fetched)) == 0)
            {
                if (fetched > 0)
                {
                    PinInfo info = new PinInfo();
                    retVal = pins[0].QueryPinInfo(info);
                    if (retVal == 0 && pred(info))
                        return pins[0];
                }
            }

            return null;
        }

        private Dictionary<string, IMoniker> EnumDevices(Guid type)
        {
            Dictionary<string, IMoniker> results = new Dictionary<string, IMoniker>();

            ICreateDevEnum enumDevices = (ICreateDevEnum)new CreateDevEnum();
            IEnumMoniker enumMoniker;
            enumDevices.CreateClassEnumerator(type, out enumMoniker, 0);

            IMoniker[] monikers = new IMoniker[1];
            IntPtr ptr = IntPtr.Zero;
            while (enumMoniker.Next(1, monikers, ptr) == 0)
            {
                object obj;
                monikers[0].BindToStorage(null, null, typeof(IPropertyBag).GUID, out obj);
                IPropertyBag props = (IPropertyBag)obj;

                object friendlyName = null;
                object description = null;
                props.Read("FriendlyName", ref friendlyName, null);
                props.Read("Description", ref description, null);

                results.Add((string)friendlyName, monikers[0]);
            }

            return results;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            CaptureVideo();
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (graph != null)
            {
                var panel = FindName("PART_VideoPanel") as System.Windows.Forms.Panel;
                IVideoWindow window = (IVideoWindow)graph;
                var retVal = window.SetWindowPosition(0, 0, (int)panel.ClientSize.Width, (int)panel.ClientSize.Height);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            callback?.Trigger?.Set();
        }
    }

    class SampleGrabberCallback : ISampleGrabberCB
    {
        public int Width { get; set; }
        public int Height { get; set; }
        public int Stride { get; set; }

        public readonly ManualResetEvent Trigger = new ManualResetEvent(false);

        private static readonly JsonSerializer serializer = JsonSerializer.Create();

        public int SampleCB(double SampleTime, IMediaSample pSample)
        {
            if (pSample == null)
                return -1;

            if (Trigger.WaitOne(0))
            {
                int len = pSample.GetActualDataLength();
                if (len > 0)
                {
                    IntPtr buf;
                    if (pSample.GetPointer(out buf) == 0)
                    {
                        byte[] buffer = new byte[len];
                        Marshal.Copy(buf, buffer, 0, len);

                        using (var bmp = new Bitmap(Width, Height, Stride, System.Drawing.Imaging.PixelFormat.Format24bppRgb, buf))
                        {
                            bmp.RotateFlip(RotateFlipType.Rotate180FlipX);
                            
                            using (var ms = new MemoryStream())
                            {
                                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

                                byte[] data = ms.ToArray();

                                var uri = new Uri($"");
                                var req = (HttpWebRequest)HttpWebRequest.Create(uri);
                                req.Method = "POST";
                                req.ContentType = "application/octet-stream";
                                req.Headers.Add("Ocp-Apim-Subscription-Key", "");
                                req.ContentLength = data.Length;
                                using (var stm = req.GetRequestStream())
                                    stm.Write(data, 0, data.Length);
                                using (var res = req.GetResponse())
                                using (var stm = res.GetResponseStream())
                                using (var sr = new StreamReader(stm))
                                using (var jr = new JsonTextReader(sr))
                                {
                                    var obj = serializer.Deserialize<ExpandoObject[]>(jr);

                                }
                            }
                        }
                    }
                }
                Trigger.Reset();
            }

            Marshal.ReleaseComObject(pSample);

            return 0;
        }

        public int BufferCB(double SampleTime, IntPtr pBuffer, int BufferLen)
        {
            return 0;
        }
    }
}

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

namespace WebCamWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private IGraphBuilder graph;
        private ICaptureGraphBuilder2 capture;
        
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


            retVal = capture.RenderStream(PinCategory.Preview, MediaType.Video, baseFilter, null, null);

            var wih = new WindowInteropHelper(this);

            IVideoWindow window = (IVideoWindow)graph;
            retVal = window.put_Owner(wih.Handle);
            retVal = window.put_WindowStyle(WindowStyles.WS_CHILD | WindowStyles.WS_CLIPCHILDREN);
            retVal = window.SetWindowPosition(0, 0, (int)Width, (int)Height);
            retVal = window.put_MessageDrain(wih.Handle);
            retVal = window.put_Visible(-1); //OATRUE

            retVal = control.Run();
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
                IVideoWindow window = (IVideoWindow)graph;
                var retVal = window.SetWindowPosition(0, 0, (int)Width, (int)Height);
            }
        }
    }


}

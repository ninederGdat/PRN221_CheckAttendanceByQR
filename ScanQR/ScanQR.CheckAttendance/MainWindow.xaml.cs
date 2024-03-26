using AForge.Video;
using AForge.Video.DirectShow;
using System.ComponentModel;
using System.Drawing;
using System.IO;
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
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
namespace ScanQR.CheckAttendance
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FilterInfoCollection videoDivces;
        private VideoCaptureDevice videoSource;
        private readonly BarcodeReaderGeneric barcodeReader;
        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;


        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            if (videoSource != null && videoSource.IsRunning)
            {
                videoSource.Stop();
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            videoDivces = new FilterInfoCollection(FilterCategory.VideoInputDevice);

            if (videoDivces.Count > 0)
            {
                videoSource = new VideoCaptureDevice(videoDivces[0].MonikerString);
                videoSource.NewFrame += VideoSource_NewFrame;
                videoSource.Start();
            }
        }

        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            fFrame.Dispatcher.Invoke(() =>
            {
                fFrame.Content = new System.Windows.Controls.Image() { Source = ToBitMapImage(eventArgs.Frame) };
            });
        }

        private System.Windows.Media.Imaging.BitmapImage ToBitMapImage(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;

                System.Windows.Media.Imaging.BitmapImage bitmapImage = new System.Windows.Media.Imaging.BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }
        private async void btnScan_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Capture current frame from the video source
                var bitmap = (fFrame.Content as System.Windows.Controls.Image)?.Source as BitmapImage;
                if (bitmap != null)
                {
                    var result = await Task.Run(() => DecodeQRCode(bitmap));
                    if (result != null)
                    {
                        txtStatus.Text = "QR Code found: " + result.Text;
                        // You can perform further actions with the QR code result here
                    }
                    else
                    {
                        txtStatus.Text = "No QR code found.";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private Result DecodeQRCode(BitmapImage bitmap)
        {
            try
            {
                //var writableBitmap = new WriteableBitmap(bitmap);
                //var luminanceSource = new ZXing.Presentation.WriteableBitmapLuminanceSource(writableBitmap);
                //var binaryBitmap = new ZXing.BinaryBitmap(new ZXing.Common.HybridBinarizer(luminanceSource));
                //return barcodeReader.Decode(binaryBitmap);
            }
            catch (Exception ex)
            {
                throw new Exception("Error decoding QR code: " + ex.Message);
            }
        }
    }
}
}
using AForge.Video;
using AForge.Video.DirectShow;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.ComponentModel;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Media.Imaging;
using ZXing;
using MessageBox = System.Windows.MessageBox;

namespace ScanQR.CheckAttendance
{
    public partial class MainWindow : Window
    {
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;
        private readonly BarcodeReaderGeneric barcodeReader;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;

            // Initialize the BarcodeReaderGeneric
            barcodeReader = new BarcodeReaderGeneric();
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (videoSource != null && videoSource.IsRunning)
            {
                videoSource.SignalToStop();

            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

            if (videoDevices.Count > 0)
            {
                videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
                videoSource.NewFrame += VideoSource_NewFrame;
                videoSource.Start();
            }
        }

        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap videoCapture = (Bitmap)eventArgs.Frame.Clone();
            Image1.Dispatcher.Invoke(() => DisplayImageInImageView(videoCapture, Image1));
        }

        private void DisplayImageInImageView(Bitmap frame, System.Windows.Controls.Image imageView)
        {
            Dispatcher.Invoke(() =>
            {
                BitmapImage bitmapImage = ToBitmapImage(frame);
                if (bitmapImage != null)
                {
                    imageView.Source = bitmapImage;
                }
            });
        }

        private BitmapImage ToBitmapImage(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;

                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }



        private ZXing.Result DecodeQRCode(BitmapImage bitmap)
        {
            try
            {
                ZXing.Windows.Compatibility.BarcodeReader barcodeReader = new ZXing.Windows.Compatibility.BarcodeReader();
                // Convert BitmapImage to Bitmap
                BitmapSource bitmapSource = bitmap;
                Bitmap bmp = BitmapFromSource(bitmapSource);

                // Decode QR code
                return barcodeReader.Decode(bmp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error decoding QR code: " + ex.Message);
                return null;
            }
        }

        // Helper method to convert BitmapSource to Bitmap
        private Bitmap BitmapFromSource(BitmapSource bitmapsource)
        {
            Bitmap bitmap;
            using (MemoryStream outStream = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapsource));
                enc.Save(outStream);
                bitmap = new Bitmap(outStream);
            }
            return bitmap;
        }


        private async void btnScan_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                // Capture current frame from the video source
                var bitmap = Image1.Source as BitmapImage;
                if (bitmap != null)
                {
                    var result = await Task.Run(() => DecodeQRCode(bitmap));
                    if (result != null)
                    {
                        txtStatus.Text = result.Text;
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
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }


        public static SheetsService API()
        {
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
                {
                    ClientId = "675809992056-pkh4tpldsqo8o6vth1ngcgft1g0ep7da.apps.googleusercontent.com",
                    ClientSecret = "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"
                }
            , Scopes, "user", CancellationToken.None, new FileDataStore("MyAppsToken")).Result,
                ApplicationName = "Google Sheets API .NET ",
            });
            return service;
        }



        private IList<IList<object>> GetValueAPI()
        {
            var service = API();
            var request = service.Spreadsheets.Values.Get("1tCq6_2pLGCAY433epZ5Iiznd6HczxrTrq123vth0uGU", "Sheet 1!A:G");
            var values = request.Execute().Values;
            return values;
        }

        private void btnCheck_Click(object sender, RoutedEventArgs e)
        {
            var sheetname = txtSheet.Text;
            var service = API();
            var request = service.Spreadsheets.Values.Get("1tCq6_2pLGCAY433epZ5Iiznd6HczxrTrq123vth0uGU", $"{sheetname}!A:G");
            var val = request.Execute().Values;
            int count = 0;
            string? timein;
            string? timeout;

            bool found = false;
            foreach (var Column in val)
            {
                count++;

                if (Column[1].ToString().Equals(txtStatus.Text))

                {
                    var date = DateTime.Now;
                    var attendanceTime = Column[5].ToString();
                    var endTime = Column[6].ToString();

                    if (date >= Convert.ToDateTime(attendanceTime) && date <= Convert.ToDateTime(endTime))
                    {
                        TimeSpan delay = DateTime.Now - Convert.ToDateTime(attendanceTime);
                        string status;
                        if (delay.TotalMinutes <= 15)
                            status = "Present"; // Đến đúng giờ
                        else
                            status = "Late"; // Đến trễ
                        var rowNumber = count;
                        var rangeVal = $"{sheetname}!D{rowNumber}:E{rowNumber}";
                        var valueRange = new ValueRange();
                        valueRange.Values = new List<IList<object>> { new List<object> { status, date.ToString("HH:mm:ss") } };
                        var updatePlank = service.Spreadsheets.Values.Update(valueRange, "1tCq6_2pLGCAY433epZ5Iiznd6HczxrTrq123vth0uGU", rangeVal);
                        updatePlank.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                        var UpdateResponse = updatePlank.Execute();
                        SpreadsheetsResource.ValuesResource.GetRequest getRequest = request;
                        System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender2, X509Certificate certificate, X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors) { return true; };
                        ValueRange getRespone = getRequest.Execute();
                        found = true;
                    }
                    else if (date > Convert.ToDateTime(endTime))
                    {
                        var status = "Absent";
                        var rowNumber = count;
                        var range = $"{sheetname}!D{rowNumber}:E{rowNumber}";
                        var ValueRange = new ValueRange();
                        ValueRange.Values = new List<IList<object>> { new List<object> { status, date.ToString("HH:mm:ss") } };
                        var updatePlank = service.Spreadsheets.Values.Update(ValueRange, "1tCq6_2pLGCAY433epZ5Iiznd6HczxrTrq123vth0uGU", range);
                        updatePlank.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                        var UpdateResponse = updatePlank.Execute();

                        SpreadsheetsResource.ValuesResource.GetRequest getRequest = request;
                        System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender2, X509Certificate certificate, X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors) { return true; };
                        ValueRange getRespone = getRequest.Execute();
                        MessageBox.Show($"Đã kết thúc lớp học, Hãy gửi email cho thầy  để được điểm danh ");

                        found = true;
                    }
                    else
                    {
                        MessageBox.Show("Lớp học chưa bắt đầu ");
                    }
                }

            }
            if (found == false)
            {
                MessageBox.Show("Bạn không có tên trong danh sách lớp này");
            }
            GetInfo();
        }


        private void GetInfo()
        {
            var val = GetValueAPI();
            foreach (var position in val)
            {
                string ID = position[0].ToString();
                string Code = position[1].ToString();
                string name = position[2].ToString();
                string attendance = position[3].ToString();
                string time = position[4].ToString();

                if (val != null && val.Count > 0)
                {
                    if (Code.Equals(txtStatus.Text))
                    {
                        txtID.Text = ID;
                        txtName.Text = name;
                        txtAttendance.Text = attendance;
                        txtTime.Text = time;
                        //return;
                    }

                }
            }
        }

        private void txtSheet_Loaded(object sender, RoutedEventArgs e)
        {
            int sheetCount = 1; // Biến đếm số thứ tự của sheet, bắt đầu từ 1
            DateTime timecheck = DateTime.Now;

            while (sheetCount < 4)
            {
                var sheetName = "Sheet " + sheetCount; // Tạo tên sheet từ số thứ tự
                var service = API();
                var request = service.Spreadsheets.Values.Get("1tCq6_2pLGCAY433epZ5Iiznd6HczxrTrq123vth0uGU", $"{sheetName}!F2:G2");
                var val = request.Execute().Values;



                foreach (var Column in val)
                {

                    string timein = Column[0].ToString();
                    string timeout = Column[1].ToString();

                    if (timein != "Time in" && timeout != "Time Out")
                    {
                        DateTime timeInDateTime, timeOutDateTime;
                        if (DateTime.TryParse(timein, out timeInDateTime) && DateTime.TryParse(timeout, out timeOutDateTime)) // Chuyển đổi thời gian sang DateTime
                        {
                            if (timecheck >= timeInDateTime && timecheck <= timeOutDateTime)
                            {
                                txtSheet.Text = "Sheet " + sheetCount;
                                return;
                            }
                        }
                    }

                }

                sheetCount++;
            }
        }

        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            txtAttendance.Text = null;
            txtName.Text = null;
            txtID.Text = null;
            txtStatus.Text = null;
            txtTime.Text = null;
        }
    }
}

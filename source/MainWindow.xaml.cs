using System;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;
using System.IO;
using System.Threading.Tasks;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace WpfApplication1 {
    public partial class MainWindow : Window {
        private const double Dpi = 300;
        private const double DisplayDpi = 96;
        private const double DpiK = Dpi / DisplayDpi;
        private static readonly PixelFormat Color = PixelFormats.Pbgra32;
        public MainWindow() {
            InitializeComponent();
        }

        private void BtnConvertClick(object sender, RoutedEventArgs e) {
            try {
                var destdir = TxtDestination.Text;
                var source = TxtSource.Text;
                if (!File.Exists(source)) {
                    MessageBox.Show("Source file does not exist.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                var input = new XpsDocument(source, FileAccess.Read).GetFixedDocumentSequence().DocumentPaginator;
                var cnt = 0;
                Parallel.ForEach(Enumerable.Range(0, input.PageCount)
                        .AsParallel()
                        .AsOrdered()
                        .Select(a => (FrameworkElement)input.GetPage(a).Visual)
                        .Select((page, index) => new {
                            FE = page,
                            Index = index,
                            RTB = new RenderTargetBitmap(C(page.Width), C(page.Height), Dpi, Dpi, Color)
                        }),
                    a => {
                        a.RTB.Render(a.FE);
                        using (var fs = File.Create(destdir + a.Index + ".png")) {
                            var png = new PngBitmapEncoder();
                            png.Frames.Add(BitmapFrame.Create(a.RTB));
                            png.Save(fs);
                            fs.Flush();
                        }
                        cnt++;
                    }
                );
                MessageBox.Show(String.Format("Rendering complete. Proceed {0} pages", cnt));
            }
            catch (Exception ex) {
                MessageBox.Show("Error occured: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private int C(double i) {
            var o = Convert.ToInt32(i * DpiK);
            return o;
        }

        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e) {
            var d = new OpenFileDialog();
            d.Filter = "XPS files (*.xps)|*.xps|All files (*.*)|*.*";
            if ( d.ShowDialog().Value )
                TxtSource.Text = d.FileName;
        }

        private void BtnBrowseDestination_Click(object sender, RoutedEventArgs e) {
            var d = new FolderBrowserDialog();
            if ( d.ShowDialog() == System.Windows.Forms.DialogResult.OK )
                TxtDestination.Text = d.SelectedPath;
        }
    }
}

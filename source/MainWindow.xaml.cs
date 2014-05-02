using System;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;
using System.IO;
using System.Threading.Tasks;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace XpsToPng {
    public partial class MainWindow : Window {
        private const double Dpi = 300;
        private const double DisplayDpi = 96;
        private const double DpiK = Dpi / DisplayDpi;
        private static readonly PixelFormat Color = PixelFormats.Pbgra32;
        public MainWindow() {
            InitializeComponent();
        }

        private void BtnConvertClick(object sender, RoutedEventArgs e) {
            new Thread( this.ConvertHandler().Wait ).Start();
        }

        private async Task ConvertHandler() {
            this.Dispatcher.Invoke(() => this.UpdateControlEnabled(false));
            try {
                string destdir = null;
                string source = null;
                this.Dispatcher.Invoke(
                    () => {
                        destdir = this.TxtDestination.Text;
                        source = this.TxtSource.Text;
                    } );
                if ( !File.Exists( source ) )
                    this.Dispatcher.Invoke( () => MessageBox.Show( "Source file does not exist.", "Error", MessageBoxButton.OK, MessageBoxImage.Error ) );
                else
                    this.ConvertCore( source, destdir );
            }
            catch ( Exception ex ) {
                this.Dispatcher.Invoke( () => MessageBox.Show( "Error occured: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error ) );
            }
            this.Dispatcher.Invoke( () => this.UpdateControlEnabled(true) );
        }

        private void UpdateControlEnabled( bool state ) {
            this.TxtDestination.IsEnabled = state;
            this.TxtSource.IsEnabled = state;
            this.BtnConvert.IsEnabled = state;
        }

        private void ConvertCore( string source, string destdir ) {
            var doc = new XpsDocument( source, FileAccess.Read );
            var input = doc.GetFixedDocumentSequence().DocumentPaginator;
            var cnt = 0;
            for (int a = 0; a < input.PageCount; a++) {
            //Parallel.For( 0, input.PageCount, a=>{
                var path = Path.Combine( destdir, a + ".png" );
                if ( File.Exists( path ) )
                    return;
                FrameworkElement fe = null;
                double h = 0D, w = 0D;
                fe = (FrameworkElement) input.GetPage( a ).Visual;
                h = fe.Height;
                w = fe.Width;
                var rtb = new RenderTargetBitmap( this.C( w ), this.C( h ), Dpi, Dpi, Color );
                rtb.Render( fe );
                using ( var fs = File.Create( path ) ) {
                    var png = new PngBitmapEncoder();
                    png.Frames.Add( BitmapFrame.Create( rtb ) );
                    png.Save( fs );
                    fs.Flush();
                }
                cnt++;
            }
            //);
            this.Dispatcher.Invoke( () => MessageBox.Show( String.Format( "Rendering complete. Proceed {0} pages", cnt ) ) );
        }

        private int C(double i) {
            var o = Convert.ToInt32(i * DpiK);
            return o;
        }

        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e) {
            var d = new OpenFileDialog { Filter = "XPS files (*.xps)|*.xps|All files (*.*)|*.*" };
            if (d.ShowDialog().Value)
                TxtSource.Text = d.FileName;
        }

        private void BtnBrowseDestination_Click(object sender, RoutedEventArgs e) {
            var d = new FolderBrowserDialog();
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                TxtDestination.Text = d.SelectedPath;
        }
    }
}

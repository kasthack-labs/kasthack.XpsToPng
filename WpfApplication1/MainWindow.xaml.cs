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
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.IO.Packaging;
using System.IO;
using System.Threading.Tasks;
namespace WpfApplication1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static double DPI = 300, DISPLAY_DPI = 96, DPI_K ;
        static PixelFormat color = PixelFormats.Pbgra32;//DEFAULT
        public MainWindow()
        {
            DPI_K = DPI / DISPLAY_DPI;
            InitializeComponent();
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                
                #region Check if file exists
                string outputmask = textBox2.Text;
                if (!File.Exists(textBox1.Text))
                {
                    MessageBox.Show("No file");
                    return;
                }
                #endregion
                #region Load
                var input = new XpsDocument(textBox1.Text, FileAccess.Read).GetFixedDocumentSequence().DocumentPaginator;
                #endregion
                #region GetPages
                int count = input.PageCount;
                var pages = Enumerable.Range(0, count).
#if !DEBUG
                     AsParallel().
                     AsOrdered().
#endif
Select(a => (FrameworkElement) input.GetPage(a).Visual).ToArray();
                GC.Collect();
                #endregion
       //         MessageBox.Show(_("Got {0} pages", pages.Length));
                #region GetRenderFrames
                var prerenders =
                    pages.
                    #if !DEBUG
                                        AsParallel().
                                        AsOrdered().
                    #endif
                    Select(b => new { FE = b, RTB = new RenderTargetBitmap(C(b.Width), C(b.Height), DPI, DPI, color) }).ToArray();
                pages = null;
                GC.Collect();
                #endregion
       //         MessageBox.Show(_("Got {0} bitmaps", prerenders.Length));
                #region Render
                int rndr = 0;
                Parallel.ForEach(prerenders, a =>
                    {
                        GC.Collect();
                        a.RTB.Render(a.FE);
                        rndr++;
                    });
                #endregion
       //         MessageBox.Show(_("Rendered {0}", rndr));
                #region Create frames
                var Frames = prerenders.
#if !DEBUG
                AsParallel().
                AsOrdered().
#endif
Select(a => BitmapFrame.Create(a.RTB)).ToArray();
                #endregion
       //         MessageBox.Show(_("Got {0} frames", Frames.Length));
                #region Save
                int i = 0;
                Parallel.ForEach(Frames, a =>
                    {
                        PngBitmapEncoder png = new PngBitmapEncoder();
                        png.Frames.Add(a);
                        var fs = File.Create(outputmask + (i++) + ".png");
                        png.Save(fs);
                        fs.Flush();
                        fs.Close();
                    });
                #endregion
                MessageBox.Show(_("Win: {0} pages", i));
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private int C(double i)
        {
            int o = Convert.ToInt32(i * DPI_K);
            return o;
        }
        private string _(string p, int count)
        {
            return String.Format(p, count);
        }
    }
}

using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace ScreenCapture
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // キャプチャ対象のデータグリッドにダミーデータを設定
            DataTable table = new DataTable();
            string[] columns = { "id", "name", "memo" };
            table.Columns.AddRange(columns.Select(n => new DataColumn(n)).ToArray());
            table.Rows.Add("id1", "name1", "memo1");
            table.Rows.Add("id2", "name2", "memo2");
            table.Rows.Add("id3", "name3", "memo3");
            this.dgTarget.ItemsSource = table.DefaultView;
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            // Windowが表示されるまで待機したいので、タイマーで1秒後に処理を実行
            var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            timer.Start();
            timer.Tick += (s, args) =>
            {
                // タイマーの停止
                timer.Stop();

                // キャプチャ開始
                this.CaptureStart();
            };
        }

        private void CaptureStart()
        {
            // キャプチャ対象のXY座標を取得
            System.Windows.Point targetPoint = this.dgTarget.PointToScreen(new System.Windows.Point(0.0d, 0.0d));

            // キャプチャ領域の生成
            Rect targetRect = new Rect(targetPoint.X, targetPoint.Y, this.dgTarget.ActualWidth, this.dgTarget.ActualHeight);

            // 画面のキャプチャ
            BitmapSource bitmap = this.CaptureScreen(targetRect);

            // キャプチャしたビットマップをImageコントロールに設定
            this.imgResult.Source = bitmap;

            // ビットマップをPNGで保存
            using (Stream stream = new FileStream("result.png", FileMode.Create))
            {
                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                encoder.Save(stream);
            }
        }

        private BitmapSource CaptureScreen(Rect rect)
        {
            // 引数rectの領域をキャプチャする
            using (var screenBmp = new Bitmap((int)rect.Width, (int)rect.Height,
                                              System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                using (var bmpGraphics = Graphics.FromImage(screenBmp))
                {
                    // キャプチャ結果をBitmapSourceで返す
                    bmpGraphics.CopyFromScreen((int)rect.X, (int)rect.Y, 0, 0, screenBmp.Size);
                    return Imaging.CreateBitmapSourceFromHBitmap(screenBmp.GetHbitmap(),
                                                                 IntPtr.Zero, Int32Rect.Empty,
                                                                 BitmapSizeOptions.FromEmptyOptions());
                }
            }
        }
    }
}
using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Tesseract;

namespace OCR
{
    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("User32", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("User32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr Parent, IntPtr Child, string lpszClass, string lpszWindows);

        public const int WM_MOUSEMOVE = 0x200;
        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_LBUTTONUP = 0x202;


        public Bitmap test = null;

        public Bitmap reservation = null;
        public Bitmap Confirm = null;
        public Bitmap Confirm2 = null;
        public Bitmap x = null;
        public Form1()
        {
            InitializeComponent();
            test = new Bitmap(@"test\test1.PNG");
            
            reservation = new Bitmap(@"test\reservation.PNG");
            Confirm = new Bitmap(@"test\Confirm.PNG");
            Confirm2 = new Bitmap(@"test\Confirm2.PNG");
            x = new Bitmap(@"test\x.PNG");
            
        }

        public static Bitmap resizeImage(Bitmap image)
        {
            if (image != null)
            {
                Bitmap croppedBitmap = new Bitmap(image);
                croppedBitmap = croppedBitmap.Clone(
                        //new Rectangle(0,(int)(image.Height*0.59), image.Width, image.Height- ((int)(image.Height * 0.59)+ (int)(image.Height * 0.315))),
                        new Rectangle((int)(image.Width*0.2), (int)(image.Height * 0.58), image.Width- (int)(image.Width * 0.4), image.Height - ((int)(image.Height * 0.58) + (int)(image.Height * 0.310))),
                        PixelFormat.Format32bppArgb);
               
                return croppedBitmap;
            }
            else
            {
                return image;
            }
        }

        private string OcrProcess(Bitmap oc)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", "kor", EngineMode.TesseractOnly))
                {
                    using(var page = engine.Process(oc))
                    {
                        return page.GetText();
                    }
                }
            }catch(Exception ex)
            {
                return ex.ToString();
            }
        }

        string AppPlayerName = "NoxPlayer2";
        
        //녹스플레이어 크기를 저정할 변수
        double full_width = 0;
        double full_height = 0;
        //녹스플레이어 지정한 크기
        double pix_width = 572;
        double pix_height = 1020;
        Bitmap bmp;
        double change_size;
        public void getBmp(Bitmap img,int num)
        {
            if (num == 5&& max != 0)
            {
                int k = n - 1;
                if (k < 0)
                    k = AppPlayerNames.Count - 1;
                AppPlayerName = AppPlayerNames[k];
            }

            IntPtr findwindow = FindWindow(null, AppPlayerName);
            if (findwindow != IntPtr.Zero)
            {
                //찾은 플레이어를 바탕으로 Graphics 정보를 가져옵니다.
                Graphics Graphicsdata = Graphics.FromHwnd(findwindow);

                //찾은 플레이어 창 크기 및 위치를 가져옵니다. 
                Rectangle rect = Rectangle.Round(Graphicsdata.VisibleClipBounds);

                full_width = rect.Width;
                full_height = rect.Height;

                if (num == 0)//화면을 내린다.
                {
                    NoxDrag(rect.Width / 2, rect.Height / 2, rect.Width / 2, 0);
                    Thread.Sleep(100);
                }

                //플레이어 창 크기 만큼의 비트맵을 선언해줍니다.
                bmp = new Bitmap(rect.Width, rect.Height);

                

                //비트맵을 바탕으로 그래픽스 함수로 선언해줍니다.
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    //찾은 플레이어의 크기만큼 화면을 캡쳐합니다.
                    IntPtr hdc = g.GetHdc();
                    PrintWindow(findwindow, hdc, 0x2);
                    g.ReleaseHdc(hdc);
                }

                System.Drawing.Size resize = new System.Drawing.Size((int)pix_width, (int)pix_height);
                bmp = new Bitmap(bmp, resize);
                change_size = full_width / pix_width;

                //pictureBox1.Image = bmp;

                if (num == 0)
                {
                    getBmp(reservation, 1);
                }
                else if (num == 1|| num == 3)
                {
                    if (searchIMG(bmp, img) >= 0.75)
                    {
                        //이미지 정중앙 클릭
                        InClick((int)((maxloc.X + FindMat.Width / 2) * change_size), (int)((maxloc.Y + FindMat.Height / 2) * change_size));
                    }
                    else
                    {
                        if(num == 1)
                            getBmp(null, 0);
                    }
                }else if (num == 2|| num == 5 )
                {
                    bmp = resizeImage(bmp);
                }
            }
        }

        public void InClick(int x, int y)
        {
            //클릭이벤트를 발생시킬 플레이어를 찾습니다.
            IntPtr findwindow = FindWindow(null, AppPlayerName);
            if (findwindow != IntPtr.Zero)
            {
                //플레이어를 찾았을 경우 클릭이벤트를 발생시킬 핸들을 가져옵니다.
                IntPtr lparam = new IntPtr(x | (y << 16));
                //플레이어 핸들에 클릭 이벤트를 전달합니다.
                SendMessage(findwindow, WM_LBUTTONDOWN, 1, lparam);
                SendMessage(findwindow, WM_LBUTTONUP, 0, lparam);
            }
        }

        static OpenCvSharp.Point minloc, maxloc;
        static Mat FindMat;

        public double searchIMG(Bitmap screen_img, Bitmap find_img)
        {
            //find_img 크기를 스크린 이미지에 맞게 조절
            Bitmap clone = find_img.Clone(new Rectangle(0, 0, find_img.Width, find_img.Height), PixelFormat.Format32bppArgb);

            Mat ScreenMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(screen_img);

            FindMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(clone);

            //스크린 이미지에서 FindMat 이미지를 찾아라
            using (Mat res = ScreenMat.MatchTemplate(FindMat, TemplateMatchModes.CCoeffNormed))
            {
                //찾은 이미지의 유사도를 담을 더블형 최대 최소 값을 선언합니다.
                double minval, maxval = 0;
                //찾은 이미지의 유사도 및 위치 값을 받습니다. 
                Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);

                //Debug.WriteLine("찾은 이미지의 유사도 : " + maxval);

                return maxval;
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, IntPtr lParam);

        public void NoxDrag(int X, int Y, int to_X, int to_Y)
        {
            IntPtr findwindow = FindWindow(null, AppPlayerName);
            Y -= 30;
            to_Y -= 30;
            PostMessage(findwindow, WM_LBUTTONDOWN, 1, new IntPtr(Y * 0x10000 + X));
            PostMessage(findwindow, WM_LBUTTONDOWN, 1, new IntPtr(to_Y * 0x10000 + to_X));
            PostMessage(findwindow, WM_LBUTTONUP, 0, new IntPtr(to_Y * 0x10000 + to_X));
        }

        static int max=0;
        static int n = 0;

        static List<string> AppPlayerNames;
        public void Process_All()
        {
            bunifuCustomLabel1.Text = "";
            AppPlayerNames = new List<string>();

            try
            {
                Process[] allProc = Process.GetProcesses();    //시스템의 모든 프로세스 정보 출력

                foreach (Process processInfo in allProc)
                {
                    if (processInfo.MainWindowTitle.Contains("Nox"))
                    {
                        AppPlayerNames.Add(processInfo.MainWindowTitle);
                    }
                }
            }
            catch (Exception ex)
            {
            }

            foreach (string s in AppPlayerNames)
            {
                bunifuCustomLabel1.Text += s + "\r\n";
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            
            Process_All();

            pictureBox1.Image= resizeImage(test);
            bmp = resizeImage(test);
            
            string ocr = "";
            if (InvokeRequired)
            {
                this.Invoke(new MethodInvoker(delegate ()
                {
                    double dZoomPercent = 200;// 200% 확대
                    Bitmap _bmpScaled = new Bitmap(bmp, (int)(bmp.Size.Width * dZoomPercent / 100), (int)(bmp.Size.Height * dZoomPercent / 100));

                    //Bitmap oc = (Bitmap)pictureBox1.Image;
                    Bitmap oc = _bmpScaled;
                    ocr = OcrProcess(oc);
                    ocr = Regex.Split(ocr, @"인원")[1].Substring(1).Trim();
                    //ocr = Regex.Replace(ocr, @"\D", "");
                    bunifuCustomLabel2.Text = ocr;

                    waiting = Enumerable.Repeat<int>(0, AppPlayerNames.Count).ToArray<int>();
                    waiting[n] = max;
                    foreach (int wait in waiting)
                        bunifuCustomLabel4.Text += wait.ToString() + "\r\n";
                }));
            }
            else
            {
                double dZoomPercent = 200;// 200% 확대
                Bitmap _bmpScaled = new Bitmap(bmp, (int)(bmp.Size.Width * dZoomPercent / 100), (int)(bmp.Size.Height * dZoomPercent / 100));

                //Bitmap oc = (Bitmap)pictureBox1.Image;
                Bitmap oc = _bmpScaled;
                ocr = OcrProcess(oc);
                //ocr = Regex.Split(Regex.Split(ocr, @"인원")[1], @"덤")[0].Trim();
                ocr = (Regex.Split(ocr, @"인원")[1]).Trim();
                ocr = ocr.Substring(0, ocr.Length - 1);
                //ocr = Regex.Replace(ocr, @"\D", "");
                bunifuCustomLabel2.Text = ocr;
                /*
                waiting = Enumerable.Repeat<int>(0, AppPlayerNames.Count).ToArray<int>();
                waiting[n] = int.Parse(ocr);
                bunifuCustomLabel4.Text = "";
                foreach (int wait in waiting)
                    bunifuCustomLabel4.Text += wait.ToString() + "\r\n";
                */
            }
        }

        static int[] waiting;

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            Process_All();
            AppPlayerName = AppPlayerNames[0];

            //waiting = new int[AppPlayerNames.Count];
            waiting = Enumerable.Repeat<int>(0, AppPlayerNames.Count).ToArray<int>();

            Thread.Sleep(1000);

            Thread acceptThread = new Thread(() => start());
            acceptThread.IsBackground = true;   // 부모 종료시 스레드 종료
            acceptThread.Start();
        }

        private void bunifuSlider1_ValueChanged(object sender, EventArgs e)
        {
            bunifuCustomLabel3.Text = bunifuSlider1.Value.ToString();
        }

        public void start()
        {
            while (true)
            {
                Thread.Sleep(5000);
                if (InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        getBmp(null, 4);
                        if (searchIMG(bmp, Confirm) >= 0.75)
                        {
                            //이미지 정중앙 클릭
                            InClick((int)((maxloc.X + FindMat.Width / 2) * change_size), (int)((maxloc.Y + FindMat.Height / 2) * change_size));
                        }else if(searchIMG(bmp, Confirm2) >= 0.75)
                        {
                            //이미지 정중앙 클릭
                            InClick((int)((maxloc.X + FindMat.Width / 2) * change_size), (int)((maxloc.Y + FindMat.Height / 2) * change_size));
                        }
                        else if (searchIMG(bmp, x) >= 0.75)
                        {
                            //이미지 정중앙 클릭
                            InClick((int)((maxloc.X + FindMat.Width / 2) * change_size), (int)((maxloc.Y + FindMat.Height / 2) * change_size));
                        }
                    }));
                }
                else
                {
                    getBmp(null, 4);
                    if (searchIMG(bmp, Confirm) >= 0.75)
                    {
                        //이미지 정중앙 클릭
                        InClick((int)((maxloc.X + FindMat.Width / 2) * change_size), (int)((maxloc.Y + FindMat.Height / 2) * change_size));
                    }
                    else if (searchIMG(bmp, x) >= 0.75)
                    {
                        //이미지 정중앙 클릭
                        InClick((int)((maxloc.X + FindMat.Width / 2) * change_size), (int)((maxloc.Y + FindMat.Height / 2) * change_size));
                    }
                    else if (searchIMG(bmp, Confirm2) >= 0.75)
                    {
                        //이미지 정중앙 클릭
                        InClick((int)((maxloc.X + FindMat.Width / 2) * change_size), (int)((maxloc.Y + FindMat.Height / 2) * change_size));
                    }
                }

                Thread.Sleep(1000);
                if (InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        getBmp(reservation,1);
                    }));
                }
                else
                {
                    getBmp(reservation, 1);
                }
                
                Thread.Sleep(1000);
                if (InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        getBmp(null,2);
                    }));
                }
                else
                {
                    getBmp(null, 2);
                }

                
                pictureBox1.BeginInvoke(new MethodInvoker(delegate
                {
                    pictureBox1.Image = bmp;
                }));

                label4.BeginInvoke(new MethodInvoker(delegate
                {
                    label4.Text = AppPlayerName + " 캡쳐 화면";
                }));
                
                string ocr="";

                if (InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        try
                        {
                            double dZoomPercent = 200;// 200% 확대
                            Bitmap _bmpScaled = new Bitmap(bmp, (int)(bmp.Size.Width * dZoomPercent / 100), (int)(bmp.Size.Height * dZoomPercent / 100));

                            //Bitmap oc = (Bitmap)pictureBox1.Image;
                            Bitmap oc = _bmpScaled;
                            ocr = OcrProcess(oc);
                            //ocr = Regex.Split(Regex.Split(ocr, @"인원")[1], @"덤")[0].Trim();
                            ocr = (Regex.Split(ocr, @"인원")[1]).Trim();
                            ocr = ocr.Substring(0, ocr.Length - 1);
                            bunifuCustomLabel2.Text = ocr;
                        }
                        catch (Exception ex)
                        {
                            bunifuCustomLabel2.Text = "";
                        }
                    }));
                }
                else
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        try
                        {
                            double dZoomPercent = 200;// 200% 확대
                            Bitmap _bmpScaled = new Bitmap(bmp, (int)(bmp.Size.Width * dZoomPercent / 100), (int)(bmp.Size.Height * dZoomPercent / 100));

                            //Bitmap oc = (Bitmap)pictureBox1.Image;
                            Bitmap oc = _bmpScaled;
                            ocr = OcrProcess(oc);
                            //ocr = Regex.Split(Regex.Split(ocr, @"인원")[1], @"덤")[0].Trim();
                            ocr = (Regex.Split(ocr, @"인원")[1]).Trim();
                            ocr = ocr.Substring(0, ocr.Length - 1);
                            bunifuCustomLabel2.Text = ocr;
                        }
                        catch (Exception ex)
                        {
                            bunifuCustomLabel2.Text = "";
                        }
                    }));
                }


                Thread.Sleep(1000);
                if (InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        getBmp(null, 5);
                    }));
                }
                else
                {
                    getBmp(null, 5);
                    
                }
                Thread.Sleep(1000);

                pictureBox3.BeginInvoke(new MethodInvoker(delegate
                {
                    pictureBox3.Image = bmp;
                }));

                label7.BeginInvoke(new MethodInvoker(delegate
                {
                    label7.Text = "이전 "+ AppPlayerName + " 캡쳐 화면";
                }));

                AppPlayerName = AppPlayerNames[n];

                string ocr2 = "";

                if (InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        try
                        {
                            double dZoomPercent = 200;// 200% 확대
                            Bitmap _bmpScaled = new Bitmap(bmp, (int)(bmp.Size.Width * dZoomPercent / 100), (int)(bmp.Size.Height * dZoomPercent / 100));

                            //Bitmap oc = (Bitmap)pictureBox1.Image;
                            Bitmap oc = _bmpScaled;
                            ocr2 = OcrProcess(oc);
                            //ocr2 = Regex.Split(Regex.Split(ocr2, @"인원")[1], @"덤")[0].Trim();
                            ocr2 = (Regex.Split(ocr2, @"인원")[1]).Trim();
                            ocr2 = ocr2.Substring(0, ocr2.Length - 1);
                            bunifuCustomLabel5.Text = ocr2;
                        }
                        catch (Exception ex)
                        {
                            bunifuCustomLabel5.Text = "";
                        }
                    }));
                }
                else
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        try
                        {
                            double dZoomPercent = 200;// 200% 확대
                            Bitmap _bmpScaled = new Bitmap(bmp, (int)(bmp.Size.Width * dZoomPercent / 100), (int)(bmp.Size.Height * dZoomPercent / 100));

                            //Bitmap oc = (Bitmap)pictureBox1.Image;
                            Bitmap oc = _bmpScaled;
                            ocr2 = OcrProcess(oc);
                            //ocr2 = Regex.Split(Regex.Split(ocr2, @"인원")[1], @"덤")[0].Trim();
                            ocr2 = (Regex.Split(ocr2, @"인원")[1]).Trim();
                            ocr2 = ocr2.Substring(0, ocr2.Length - 1);
                            bunifuCustomLabel5.Text = ocr2;
                        }
                        catch (Exception ex)
                        {
                            bunifuCustomLabel5.Text = "";
                        }
                    }));
                }


                if (InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        try
                        {
                            if (bunifuCustomLabel2.Text != "")
                            {
                                //if (max == 0 || int.Parse(ocr) > max + int.Parse(bunifuCustomLabel3.Text))
                                if (bunifuCustomLabel5.Text == "" || max == 0 || Math.Abs(int.Parse(bunifuCustomLabel2.Text) - int.Parse(bunifuCustomLabel5.Text)) >= int.Parse(bunifuCustomLabel3.Text))
                                {
                                    max = int.Parse(bunifuCustomLabel2.Text);
                                    waiting[n] = max;
                                    n++;
                                    if (n > AppPlayerNames.Count - 1)
                                        n = 0;
                                    
                                    AppPlayerName = AppPlayerNames[n];
                                    bunifuCustomLabel4.Text = "";
                                    foreach (int wait in waiting)
                                        bunifuCustomLabel4.Text += wait.ToString()+"\r\n";
                                    
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }));
                }
                else
                {
                    try
                    {
                        if (bunifuCustomLabel2.Text != "")
                        {
                            //if (max == 0 || int.Parse(ocr) > max + int.Parse(bunifuCustomLabel3.Text))
                            if (bunifuCustomLabel5.Text == "" || max == 0 || Math.Abs(int.Parse(bunifuCustomLabel2.Text) - int.Parse(bunifuCustomLabel5.Text)) >= int.Parse(bunifuCustomLabel3.Text))
                            {
                                n++;
                                if (n > AppPlayerNames.Count - 1)
                                    n = 0;
                                max = int.Parse(bunifuCustomLabel2.Text);
                                AppPlayerName = AppPlayerNames[n];
                                waiting[n] = max;
                                bunifuCustomLabel4.Text = "";
                                foreach (int wait in waiting)
                                    bunifuCustomLabel4.Text += wait.ToString() + "\r\n";

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }

                
            }
        }

    }
}

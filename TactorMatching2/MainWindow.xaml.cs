using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Syncfusion.XlsIO;

namespace TactorMatching2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Tactor Methods
        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll")]
        public static extern IntPtr GetVersionNumber();

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern int Discover(int type);

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern int Connect(string name, int type, IntPtr _callback);

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern int InitializeTI();

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern int Pulse(int deviceID, int tacNum, int msDuration, int delay);

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr GetDiscoveredDeviceName(int index);

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern int DiscoverLimited(int type, int amount);

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
           CallingConvention = CallingConvention.Cdecl)]
        public static extern int ChangeGain(int deviceID, int tacNum, int gainval, int delay);

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern int CloseAll();

        [DllImport(@"C:\Users\minisim\Desktop\Tactors\TDKAPI_1.0.6.0\libraries\Windows\TactorInterface.dll",
            CallingConvention = CallingConvention.Cdecl)]
        public static extern int UpdateTI();
        #endregion

        #region enums
        enum matchings
        {
            Pan_Back,
            Back_Pan,
            Pan_Belt,
            Belt_Pan,
            Pan_Wear,
            Wear_Pan,
            Belt_Wear,
            Wear_Belt,
            Belt_Back,
            Back_Belt,
            Wear_Back,
            Back_Wear
        }

        enum matchings2
        {
            Back_Pan,
            Pan_Back,
            Belt_Pan,
            Pan_Belt,
            Wear_Pan,
            Pan_Wear,
            Wear_Belt,
            Belt_Wear,
            Back_Belt,
            Belt_Back,
            Back_Wear,
            Wear_Back
        }

        enum intensity
        {
            low,
            mid,
            high
        }
        #endregion
        #region variables
        private const int low = 51;
        private const int mid = 128;
        private const int high = 204;
        private int lowIndex;
        private int midIndex;
        private int highIndex;
        private const int minGain = 0;
        private const int maxGain = 255;
        Random rand;
        private int[,] lowData;
        private int[,] midData;
        private int[,] highData;
        private int[] checkOff;
        private int currentGain;
        private int currentIntensity;
        private int buzzIntensity;
        private Button[] matchButtons;
        private bool startedTask;
        private matchings currentMatch;
        private int matchesCompleted = 0;

        private ExcelEngine excelEngine;
        private IApplication application;
        private IWorkbook workbook;
        private IWorksheet sheet;

        #endregion

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            CloseAll();
        }

        public void setupTactors()
        {
            if (InitializeTI() == 0)
            {
                System.Diagnostics.Debug.WriteLine("TDK Initialized");
            }

            System.Diagnostics.Debug.WriteLine("Tactors: " + Discover(1));
            string name = Marshal.PtrToStringAnsi((IntPtr)GetDiscoveredDeviceName(0));
            System.Diagnostics.Debug.WriteLine("Name: " + name );


            if (Connect(name, 1, IntPtr.Zero) >= 0)
            {
                System.Diagnostics.Debug.WriteLine("Connected");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Not connected");
            }
        }

        public void setUpVariables()
        {
            // 12 rows to store each type of matching
            // 4 columns to store low, mid, high, and avg
            lowData = new int[12,4];
            midData = new int[12, 4];
            highData = new int[12, 4];
            checkOff = new int[9];

            lowIndex = 0;
            midIndex = 0;
            highIndex = 0;

            rand = new Random();
            startedTask = false;

            matchButtons = new Button[]
            {
                panToBackBtn,
                backToPanBtn,
                panToBeltBtn,
                beltToPanBtn,
                panToWearableBtn,
                wearableToPanBtn,
                beltToWearableBtn,
                wearableToBeltBtn,
                BeltToBackBtn,
                backToBeltBtn,
                wearableToBackBtn,
                backToWearableBtn
            };

            submitBtn.Visibility = Visibility.Hidden;
            buzzBtn.Visibility = Visibility.Hidden;
            countLbl.Visibility = Visibility.Hidden;
            currentGain = 0;
        }

        public void setupSheet()
        {
            excelEngine = new ExcelEngine();
            application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;
            workbook = application.Workbooks.Open(@"C:\Gaojian\TactorMatching2\TactorMatching2_Template(DO NOT MOVE OR DELETE.xlsx");
            sheet = workbook.Worksheets[0];
        }

        public MainWindow()
        {
            InitializeComponent();
            setupTactors();
            setUpVariables();
            setupSheet();

            this.KeyDown += (sender, e) => grid_KeyDown(sender, e);
        }

        #region pulse methods
        public void pulsePan()
        {
            Pulse(0, 13, 500, 0);
            Pulse(0, 14, 500, 0);
            Pulse(0, 15, 500, 0);
            Pulse(0, 16, 500, 0);
        }

        public void pulseBack()
        {
            Pulse(0, 9, 500, 0);
            Pulse(0, 10, 500, 0);
            Pulse(0, 11, 500, 0);
            Pulse(0, 12, 500, 0);
        }

        public void pulseWearable()
        {
            Pulse(0, 5, 500, 0);
            Pulse(0, 6, 500, 0);
            Pulse(0, 7, 500, 0);
            Pulse(0, 8, 500, 0);
        }

        public void pulseBelt()
        {
            Pulse(0, 1, 500, 0);
            Pulse(0, 2, 500, 0);
            Pulse(0, 3, 500, 0);
            Pulse(0, 4, 500, 0);
        }
        #endregion

        #region change gain methods
        public void changePanGain(int gain)
        {
            ChangeGain(0, 13, gain, 0);
            ChangeGain(0, 14, gain, 0);
            ChangeGain(0, 15, gain, 0);
            ChangeGain(0, 16, gain, 0);
        }

        public void changeBackGain(int gain)
        {
            ChangeGain(0, 9, gain, 0);
            ChangeGain(0, 10, gain, 0);
            ChangeGain(0, 11, gain, 0);
            ChangeGain(0, 12, gain, 0);
        }
        
        public void changeWearableGain(int gain)
        {
            ChangeGain(0, 5, gain, 0);
            ChangeGain(0, 6, gain, 0);
            ChangeGain(0, 7, gain, 0);
            ChangeGain(0, 8, gain, 0);
        }

        public void changeBeltGain(int gain)
        {
            ChangeGain(0, 1, gain, 0);
            ChangeGain(0, 2, gain, 0);
            ChangeGain(0, 3, gain, 0);
            ChangeGain(0, 4, gain, 0);
        }
        #endregion

        public void hideButtons()
        {
            foreach(dynamic b in matchButtons)
            {
                if(b != null)
                {
                    b.Visibility = Visibility.Hidden;
                }

            }
        }

        public void showButtons()
        {
            foreach(dynamic b in matchButtons)
            {
                if(b != null)
                {
                    b.Visibility = Visibility.Visible;
                }
            }
        }

        public void setBuzzGain()
        {
            switch (currentIntensity)
            {
                case 0:
                    buzzIntensity = low;
                    break;
                case 1:
                    buzzIntensity = mid;
                    break;
                case 2:
                    buzzIntensity = high;
                    break;
            }

            switch (currentMatch)
            {
                case matchings.Pan_Back:
                    changeBackGain(buzzIntensity);
                    break;
                case matchings.Back_Pan:
                    changePanGain(buzzIntensity);
                    break;
                case matchings.Pan_Belt:
                    changeBeltGain(buzzIntensity);
                    break;
                case matchings.Belt_Pan:
                    changePanGain(buzzIntensity);
                    break;
                case matchings.Pan_Wear:
                    changeWearableGain(buzzIntensity);
                    break;
                case matchings.Wear_Pan:
                    changePanGain(buzzIntensity);
                    break;
                case matchings.Belt_Wear:
                    changeWearableGain(buzzIntensity);
                    break;
                case matchings.Wear_Belt:
                    changeBeltGain(buzzIntensity);
                    break;
                case matchings.Belt_Back:
                    changeBackGain(buzzIntensity);
                    break;
                case matchings.Back_Belt:
                    changeBeltGain(buzzIntensity);
                    break;
                case matchings.Wear_Back:
                    changeBackGain(buzzIntensity);
                    break;
                case matchings.Back_Wear:
                    changeWearableGain(buzzIntensity);
                    break;
            }
        }

        public void startMatch()
        {
            startedTask = true;
            submitBtn.Visibility = Visibility.Visible;
            buzzBtn.Visibility = Visibility.Visible;
            countLbl.Visibility = Visibility.Visible;
            countLbl.Content = "0/9";
            currentIntensity = rand.Next(3);
            intensityLbl.Content = "Matching Intensity: " + Enum.GetName(typeof(intensity), currentIntensity);
            currentGainLbl.Content = "Current: 0";
            setBuzzGain();
            hideButtons();
        }

        #region Button events
        private void PanToBackBtn_Click(object sender, RoutedEventArgs e)
        {
            //1
            currentMatch = matchings.Pan_Back;
            startMatch();
        }

        private void BackToPanBtn_Click(object sender, RoutedEventArgs e)
        {
            //2
            currentMatch = matchings.Back_Pan;
            startMatch();
        }

        private void PanToBeltBtn_Click(object sender, RoutedEventArgs e)
        {
            //3
            currentMatch = matchings.Pan_Belt;
            startMatch();
        }

        private void BeltToPanBtn_Click(object sender, RoutedEventArgs e)
        {
            //4
            currentMatch = matchings.Belt_Pan;
            startMatch();
        }

        private void PanToWearableBtn_Click(object sender, RoutedEventArgs e)
        {
            //5
            currentMatch = matchings.Pan_Wear;
            startMatch();
        }

        private void WearableToPanBtn_Click(object sender, RoutedEventArgs e)
        {
            //6
            currentMatch = matchings.Wear_Pan;
            startMatch();
        }

        private void BeltToWearableBtn_Click(object sender, RoutedEventArgs e)
        {
            //7
            currentMatch = matchings.Belt_Wear;
            startMatch();
        }

        private void WearableToBeltBtn_Click(object sender, RoutedEventArgs e)
        {
            //8
            currentMatch = matchings.Wear_Belt;
            startMatch();
        }

        private void BeltToBackBtn_Click(object sender, RoutedEventArgs e)
        {
            //9
            currentMatch = matchings.Belt_Back;
            startMatch();
        }

        private void BackToBeltBtn_Click(object sender, RoutedEventArgs e)
        {
            //10
            currentMatch = matchings.Back_Belt;
            startMatch();
        }

        private void WearableToBackBtn_Click(object sender, RoutedEventArgs e)
        {
            //11
            currentMatch = matchings.Wear_Back;
            startMatch();
        }

        private void BackToWearableBtn_Click(object sender, RoutedEventArgs e)
        {
            //12
            currentMatch = matchings.Back_Wear;
            startMatch();
        }
        #endregion

        private void changeGain()
        {
            System.Diagnostics.Debug.WriteLine("currentMatch: " + currentMatch);

            switch (currentMatch)
            {
                case matchings.Pan_Back:
                    changePanGain(currentGain);
                    pulsePan();
                    break;
                case matchings.Back_Pan:
                    changeBackGain(currentGain);
                    pulseBack();
                    break;
                case matchings.Pan_Belt:
                    changePanGain(currentGain);
                    pulsePan();
                    break;
                case matchings.Belt_Pan:
                    changeBeltGain(currentGain);
                    pulseBelt();
                    break;
                case matchings.Pan_Wear:
                    changePanGain(currentGain);
                    pulsePan();
                    break;
                case matchings.Wear_Pan:
                    changeWearableGain(currentGain);
                    pulseWearable();
                    break;
                case matchings.Belt_Wear:
                    changeBeltGain(currentGain);
                    pulseBelt();
                    break;
                case matchings.Wear_Belt:
                    changeWearableGain(currentGain);
                    pulseWearable();
                    break;
                case matchings.Belt_Back:
                    changeBeltGain(currentGain);
                    pulseBelt();
                    break;
                case matchings.Back_Belt:
                    changeBackGain(currentGain);
                    pulseBack();
                    break;
                case matchings.Wear_Back:
                    changeWearableGain(currentGain);
                    pulseWearable();
                    break;
                case matchings.Back_Wear:
                    changeBackGain(currentGain);
                    pulseBack();
                    break;
            }
        }

        private void grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (startedTask)
            {
                if (e.Key == Key.Left)
                {
                    currentGain -= 15;
                    if(currentGain <= minGain)
                    {
                        currentGain = minGain;
                    }
                    changeGain();
                }
                if (e.Key == Key.Right)
                {
                    currentGain += 15;
                    if(currentGain >= maxGain)
                    {
                        currentGain = maxGain;
                    }
                    changeGain();
                }
                if(e.Key == Key.B)
                {
                    switch (currentMatch)
                    {
                        case matchings.Pan_Back:
                            pulseBack();
                            break;
                        case matchings.Back_Pan:
                            pulsePan();
                            break;
                        case matchings.Pan_Belt:
                            pulseBelt();
                            break;
                        case matchings.Belt_Pan:
                            pulsePan();
                            break;
                        case matchings.Pan_Wear:
                            pulseWearable();
                            break;
                        case matchings.Wear_Pan:
                            pulsePan();
                            break;
                        case matchings.Belt_Wear:
                            pulseWearable();
                            break;
                        case matchings.Wear_Belt:
                            pulseBelt();
                            break;
                        case matchings.Belt_Back:
                            pulseBack();
                            break;
                        case matchings.Back_Belt:
                            pulseBelt();
                            break;
                        case matchings.Wear_Back:
                            pulseBack();
                            break;
                        case matchings.Back_Wear:
                            pulseWearable();
                            break;
                    }
                }

                currentGainLbl.Content = "Current: " + currentGain;
            }
        }

        public void disableButton(int m)
        {
            switch (m)
            {
                case (int)matchings.Pan_Back:
                    matchButtons[0] = null;
                    break;
                case (int)matchings.Back_Pan:
                    matchButtons[1] = null;
                    break;
                case (int)matchings.Pan_Belt:
                    matchButtons[2] = null;
                    break;
                case (int)matchings.Belt_Pan:
                    matchButtons[3] = null;
                    break;
                case (int)matchings.Pan_Wear:
                    matchButtons[4] = null;
                    break;
                case (int)matchings.Wear_Pan:
                    matchButtons[5] = null;
                    break;
                case (int)matchings.Belt_Wear:
                    matchButtons[6] = null;
                    break;
                case (int)matchings.Wear_Belt:
                    matchButtons[7] = null;
                    break;
                case (int)matchings.Belt_Back:
                    matchButtons[8] = null;
                    break;
                case (int)matchings.Back_Belt:
                    matchButtons[9] = null;
                    break;
                case (int)matchings.Wear_Back:
                    matchButtons[10] = null;
                    break;
                case (int)matchings.Back_Wear:
                    matchButtons[11] = null;
                    break;
            }
        }


        private void applyDataToSheet()
        {
            string time = DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss");
            int row = 3;

            sheet.Range["A2"].Text = time;

            for(int x = 0; x < 12; x++)
            {
                sheet.Range["C" + row].Number = lowData[x, 0];
                sheet.Range["D" + row].Number = lowData[x, 1];
                sheet.Range["E" + row].Number = lowData[x, 2];
                sheet.Range["F" + row].Number = (lowData[x, 0] + lowData[x, 1] + lowData[x, 2]) / 3;

                sheet.Range["H" + row].Number = midData[x, 0];
                sheet.Range["I" + row].Number = midData[x, 1];
                sheet.Range["J" + row].Number = midData[x, 2];
                sheet.Range["K" + row].Number = (midData[x, 0] + midData[x, 1] + midData[x, 2]) / 3;

                sheet.Range["M" + row].Number = highData[x, 0];
                sheet.Range["N" + row].Number = highData[x, 1];
                sheet.Range["O" + row].Number = highData[x, 2];
                sheet.Range["P" + row].Number = (highData[x, 0] + highData[x, 1] + highData[x, 2]) / 3;

                row += 4;
            }
        }

        private void SubmitBtn_Click(object sender, RoutedEventArgs e)
        {
            switch (currentIntensity)
            {
                case 0:
                    if (lowIndex < 3)
                    {
                        System.Diagnostics.Debug.WriteLine("saving gain: " + currentGain);
                        lowData[(int)currentMatch, lowIndex] = currentGain;
                        lowIndex++;
                        countLbl.Content = (lowIndex + midIndex + highIndex) + "/9";
                    }
                    break;
                case 1:
                    if (midIndex < 3)
                    {
                        System.Diagnostics.Debug.WriteLine("saving gain: " + currentGain);
                        midData[(int)currentMatch, midIndex] = currentGain;
                        midIndex++;
                        countLbl.Content = (lowIndex + midIndex + highIndex) + "/9";

                    }
                    break;
                case 2:
                    if (highIndex < 3)
                    {
                        System.Diagnostics.Debug.WriteLine("saving gain: " + currentGain);
                        highData[(int)currentMatch, highIndex] = currentGain;
                        highIndex++;
                        countLbl.Content = (lowIndex + midIndex + highIndex) + "/9";

                    }
                    break;
            }

            if (lowIndex == 3 && midIndex == 3 && highIndex == 3)
            {
                //finished this matching
                currentGain = 0;
                submitBtn.Visibility = Visibility.Hidden;
                buzzBtn.Visibility = Visibility.Hidden;
                countLbl.Visibility = Visibility.Hidden;
                currentGainLbl.Content = "";
                intensityLbl.Content = "";
                startedTask = false;

                disableButton((int)currentMatch);
                showButtons();

                for (int x = 0; x < 3; x++)
                {
                    System.Diagnostics.Debug.Write(lowData[(int)currentMatch, x] + ", ");
                }
                System.Diagnostics.Debug.WriteLine("");
                for (int x = 0; x < 3; x++)
                {
                    System.Diagnostics.Debug.Write(midData[(int)currentMatch, x] + ", ");
                }
                System.Diagnostics.Debug.WriteLine("");

                for (int x = 0; x < 3; x++)
                {
                    System.Diagnostics.Debug.Write(highData[(int)currentMatch, x] + ", ");
                }
                System.Diagnostics.Debug.WriteLine("");


                lowIndex = 0;
                midIndex = 0;
                highIndex = 0;
                matchesCompleted++;
                if(matchesCompleted == 12)
                {
                    CloseAll();
                    applyDataToSheet();
                    workbook.SaveAs(@"C:\Gaojian\TactorMatching2\Data\" + DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss") + ".xlsx");
                    this.Close();
                }
                return;
            }


            currentIntensity = rand.Next(3);
            switch (currentIntensity)
            {
                case 0:
                    if(lowIndex == 3)
                    {
                        while (currentIntensity == 0)
                        {
                            currentIntensity = rand.Next(3);
                            switch (currentIntensity)
                            {
                                case 1:
                                    if(midIndex == 3)
                                    {
                                        currentIntensity = 0;
                                    }
                                    break;
                                case 2:
                                    if(highIndex == 3)
                                    {
                                        currentIntensity = 0;
                                    }
                                    break;  
                            }
                        }
                    }
                    break;
                case 1:
                    if (midIndex == 3)
                    {
                        while (currentIntensity == 1)
                        {
                            currentIntensity = rand.Next(3);
                            switch (currentIntensity)
                            {
                                case 0:
                                    if (lowIndex == 3)
                                    {
                                        currentIntensity = 1;
                                    }
                                    break;
                                case 2:
                                    if (highIndex == 3)
                                    {
                                        currentIntensity = 1;
                                    }
                                    break;
                            }
                        }
                    }
                    break;
                case 2:
                    if (highIndex == 3)
                    {
                        while (currentIntensity == 2)
                        {
                            currentIntensity = rand.Next(3);
                            switch (currentIntensity)
                            {
                                case 0:
                                    if (lowIndex == 3)
                                    {
                                        currentIntensity = 2;
                                    }
                                    break;
                                case 1:
                                    if (midIndex == 3)
                                    {
                                        currentIntensity = 2;
                                    }
                                    break;
                            }
                        }
                    }
                    break;
            }

            intensityLbl.Content = "Intensity: " + Enum.GetName(typeof(intensity), currentIntensity);
            setBuzzGain();
            currentGain = 0;
            currentGainLbl.Content = "Current: 0";
        }

        private void BuzzBtn_Click(object sender, RoutedEventArgs e)
        {
            switch (currentMatch)
            {
                case matchings.Pan_Back:
                    pulseBack();
                    break;
                case matchings.Back_Pan:
                    pulsePan();
                    break;
                case matchings.Pan_Belt:
                    pulseBelt();
                    break;
                case matchings.Belt_Pan:
                    pulsePan();
                    break;
                case matchings.Pan_Wear:
                    pulseWearable();
                    break;
                case matchings.Wear_Pan:
                    pulsePan();
                    break;
                case matchings.Belt_Wear:
                    pulseWearable ();
                    break;
                case matchings.Wear_Belt:
                    pulseBelt();
                    break;
                case matchings.Belt_Back:
                    pulseBack();
                    break;
                case matchings.Back_Belt:
                    pulseBelt();
                    break;
                case matchings.Wear_Back:
                    pulseBack();
                    break;
                case matchings.Back_Wear:
                    pulseWearable();
                    break;
            }
        }
    }
}

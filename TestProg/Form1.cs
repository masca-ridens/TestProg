using Broadcom_lib;
using Ivi.Visa.Interop;
using lib_Win32DeviceManagement;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace TestProg
{
    public partial class Form1 : Form
    {
        static SerialPort usb;
        private readonly ResourceManager ioMgr = new ResourceManager();
        List<FormattedIO488> instruments = new List<FormattedIO488>();
        List<string> instrumentLog = new List<string>();
        static readonly string myDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        readonly string resultsFile = Path.Combine(myDesktop, "results.txt");
        bool automaticSwitching = false;

        readonly Dictionary<int, string> SpectrumAnalyserMode = new Dictionary<int, string>()
            {
                {1, "Spectrum Analyzer" },
                {2, "Real Time Spectrum Analyzer" },
                {8, @"I//Q Analyzer (Basic)" },
                {14, "Phase Noise" },
                {101, "89601 VSA" },
                { 219, "Noise Figure"},
                {234, "Analog Demod" }
            };

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string comPort = string.Empty;
            try
            {
                foreach (Win32DeviceMgmt.DeviceInfo di in Win32DeviceMgmt.GetAllCOMPorts())
                {
                    string bus_description = di.bus_description;
                    if (bus_description.Contains("Multigen"))
                    {
                        comPort = di.name;
                        rtb1.AppendText("\"" + bus_description + "\" found on port " + comPort + Environment.NewLine);
                    }
                }
            }
            catch (Exception ex)
            {
                comPort = "COM7";
            }
            finally
            {
                usb = new SerialPort
                {
                    PortName = comPort,
                    ReadTimeout = 250,
                    WriteTimeout = 250,
                    BaudRate = 115200,
                    Parity = Parity.None,
                    StopBits = StopBits.One,
                    DataBits = 8,
                    Handshake = Handshake.None,
                    NewLine = "\r"
                };
                usb.Open();
            }

            string[] resources = ioMgr.FindRsrc("?*");
            cbInstrumentSA1.Items.Clear();
            cbInstrumentSG1.Items.Clear();
            cbInstrumentSG2.Items.Clear();
            cbInstrumentPickering.Items.Clear();

            cbInstrumentSA1.Items.AddRange(resources);
            cbInstrumentSG1.Items.AddRange(resources);
            cbInstrumentSG2.Items.AddRange(resources);
            cbInstrumentPickering.Items.AddRange(resources);

            tabControl1.SelectedTab = tabPage2;

            checkBox1.Checked = true;
        }

        private void BDCOn_Click(object sender, EventArgs e)
        {
            try
            {
                usb.WriteLine("OE\r");
                rtb1.AppendText(usb.ReadLine() + Environment.NewLine);
            }
            catch { MessageBox.Show("Multigen timed out"); }
        }

        private void BUnlock_Click(object sender, EventArgs e)
        {
            bool success = FSK.Unlock(usb, 'F', 1);
            if (success)
                rtb1.AppendText("Unlocked" + Environment.NewLine);
            else rtb1.AppendText("Unlock failed" + Environment.NewLine);
        }

        private void BStatus_Click(object sender, EventArgs e)
        {
            string[] s = FSK.ReadStatus0x10(usb, 'F', 1);
            rtb1.AppendText(s[0] + Environment.NewLine);
            rtb1.AppendText(s[1] + Environment.NewLine);
        }

        private void BBS_Click(object sender, EventArgs e)
        {
            bool success = FSK.SetStackingMode(usb, 1, "Bandstack");
            if (success) rtb1.AppendText("Changed to B/S mode" + Environment.NewLine);
            else rtb1.AppendText("Failed to change to B/S mode" + Environment.NewLine);
        }

        private void BCS_Click(object sender, EventArgs e)
        {
            var relevantButton = gbRx.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            int rx = int.Parse(relevantButton.Text);

            bool success = FSK.SetStackingMode(usb, rx, "Channelstack");
            if (success) rtb1.AppendText("Changed port " + rx.ToString() + " to C/S mode" + Environment.NewLine);
            else rtb1.AppendText("Failed to change port " + rx.ToString() + " to C/S mode" + Environment.NewLine);
        }

        private void BPolling_Click(object sender, EventArgs e)
        {
            var relevantButton = gbRx.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            int rx = int.Parse(relevantButton.Text);

            bool success = FSK.DisableTdmaPolling(usb, 'F', rx, true);
            if (success)
                rtb1.AppendText("Polling disabled" + Environment.NewLine);
            else rtb1.AppendText("Polling disable failed" + Environment.NewLine);
        }

        private void BAllocate_Click(object sender, EventArgs e)
        {
            var relevantButton = gbRx.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            int rx = int.Parse(relevantButton.Text);

            int channel = (int)(numChannel.Value);

            bool success = FSK.AllocateChannel(usb, 'F', rx, channel);
            if (success)
                rtb1.AppendText("Channel allocated" + Environment.NewLine);
            else
                rtb1.AppendText("Channel allocation failed" + Environment.NewLine);
        }

        private void BDeallocate_Click(object sender, EventArgs e)
        {
            var relevantButton = gbRx.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            int rx = int.Parse(relevantButton.Text);

            int channel = (int)(numChannel.Value);

            bool success = FSK.DeallocateChannel(usb, 'F', rx, channel);
            if (success)
                rtb1.AppendText("Channel deallocated" + Environment.NewLine);
            else
                rtb1.AppendText("Channel deallocation failed" + Environment.NewLine);
        }

        private void BReset_Click(object sender, EventArgs e)
        {
            bool success = FSK.SoftReset(usb, 'F', 1);
            if (success)
                rtb1.AppendText("Reset succeeded" + Environment.NewLine);
            else
                rtb1.AppendText("Reset failed" + Environment.NewLine);
        }

        private void bSerialNo_Click(object sender, EventArgs e)
        {
            string[] serialNo = FSK.ReadSerialNumbers(usb, 'F', 1);
            rtb1.AppendText("Dish Serial # = " + serialNo[0]);
            rtb1.AppendText(Environment.NewLine);
            rtb1.AppendText("GI Serial # = " + serialNo[1] + Environment.NewLine);
        }

        private void FrequencySelect_Click(object sender, EventArgs e)
        {
            var relevantButton = gbRx.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            int rx = int.Parse(relevantButton.Text);

            relevantButton = gbLNB.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            string lnb = relevantButton.Text;

            int channel = (int)(numChannel.Value);

            decimal MHz = numFrequency.Value;
            int KHz = 1000 * (int)MHz;
            if (lnb == "C" && KHz >= 1650000)
                KHz += 28500;

            bool success = FSK.Command38(usb, 'F', rx, lnb, channel, KHz, null, null);
            if (success)
                rtb1.AppendText(MHz.ToString() + " MHz selected" + Environment.NewLine);
            else
                rtb1.AppendText("Frequency selection failed" + Environment.NewLine);
        }

        private void BBandSelect_Click(object sender, EventArgs e)
        {
            var relevantButton = gbLNB.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            string lnb = relevantButton.Text;

            relevantButton = gbBand.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
            string band = relevantButton.Text;

            bool isUpper = (rbU.Checked == true) ? true : false;

            List<int> rx = new List<int>();
            if (chRx1.Checked == true) rx.Add(1);
            if (chRx2.Checked == true) rx.Add(2);

            List<string> bands = new List<string>();
            if (chLow.Checked == true) bands.Add("Low");
            if (chMid.Checked == true) bands.Add("Mid");
            if (chHigh.Checked == true) bands.Add("High");

            foreach (int r in rx)
            {
                foreach (string b in bands)
                {
                    bool success = FSK.Command38(usb, 'D', r, lnb, null, null, isUpper, b);
                    if (success)
                        rtb1.AppendText("LNB " + lnb + " " + (isUpper ? "Upper" : "Lower") + " to Rx" + r.ToString() + " " + b + Environment.NewLine);
                    else
                        rtb1.AppendText("Frequencies assignment failed" + Environment.NewLine);
                }
            }
        }

        private void BDCOff_Click(object sender, EventArgs e)
        {
            try
            {
                usb.WriteLine("OD\r");
                rtb1.AppendText(usb.ReadLine() + Environment.NewLine);
            }
            catch { MessageBox.Show("Multigen timed out"); }
        }

        private void BRead_Click(object sender, EventArgs e)
        {
            try
            {
                string reply = usb.ReadLine();
                rtb1.AppendText(reply + Environment.NewLine);
            }
            catch (TimeoutException) { };
        }

        private void BUbGrid_Click(object sender, EventArgs e)
        {
            int[] grid = FSK.RequestUserBlockGrid(usb, 'F', 1);
            rtb1.AppendText("[" + grid[0].ToString() + " : " + grid[1].ToString() + "]" + Environment.NewLine);
        }

        private void Binstrument_Click(object sender, EventArgs e)
        {
            // Identify the controls we need to deal with...

            Button b = sender as Button;
            if (instrumentLog.Contains(b.Tag.ToString())) return;
            GroupBox gb = b.Parent as GroupBox;
            var label = gb.Controls.OfType<Label>().FirstOrDefault(r => r.Tag != null && r.Tag.ToString() == "idn?");
            var cbox = gb.Controls.OfType<ComboBox>().FirstOrDefault(r => r.Tag != null && r.Tag.ToString() == "Instrument");

            // Validate the instrument address...

            string address = cbox.SelectedItem.ToString();
            if (string.IsNullOrEmpty(address))
            {
                MessageBox.Show("Select an instrument");
                return;
            }

            instrumentLog.Add(b.Tag.ToString());

            FormattedIO488 instrument = new Ivi.Visa.Interop.FormattedIO488();

            try
            {
                instrument.IO = (IMessage)ioMgr.Open(address, AccessMode.NO_LOCK, 1000, "");
                instruments.Add(instrument);
                instruments[instruments.Count - 1].IO.SendEndEnabled = false;
                instruments[instruments.Count - 1].IO.TerminationCharacterEnabled = true;
                instruments[instruments.Count - 1].WriteString("*IDN?", true);
                label.Text = instruments[instruments.Count - 1].ReadString();
                b.Enabled = false;
            }
            catch (COMException ex)
            {
                switch (gb.Tag.ToString())
                {
                    case "SA1":
                        {
                            MessageBox.Show("Failed to connect to the Spectrum Analyser at address " + address);
                            rtb1.AppendText(ex.ToString() + Environment.NewLine);
                            break;
                        }
                    case "SG1":
                    case "SG2":
                        {
                            MessageBox.Show("Failed to connect to the Signal Generator at address " + address);
                            break;
                        }
                    case "Pickering":
                        {
                            MessageBox.Show("Failed to connect to the Pickering switch. Connect ports manually.");
                            automaticSwitching = false;
                            break;
                        }
                }
                return;
            }
        }

        private void BSpurTest_Click(object sender, EventArgs e)
        {
            char temperature = PreliminaryActions("TBS", out string test, out List<int> ports, out List<string> LNBs, out List<string> bands, out string serialNumber);

            File.AppendAllText(resultsFile,
                "LNB".PadRight(4) + "Pol.".PadRight(6) + "Rx".PadRight(4) + "Band".PadRight(7) +
                    "Input_MHz".PadRight(17) +
                    (test == "In-band" ? "Wanted_MHz" : "Unwanted").PadRight(17) +
                    (test == "In-band" ? "Worst_MHz".PadRight(17) : string.Empty) +
                    "Worst_dBm".PadRight(17) +
                    (test == "In-band" ? "Worst_dBc" : string.Empty) + Environment.NewLine);

            // Sort out the ProgressBar...
            progressBar1.Maximum = LNBs.Count * ports.Count * 2 * bands.Count;

            // Set each Rx to BANDSTACKED
            foreach (int r in ports)
            {
                bool success = FSK.SetStackingMode(usb, r, "BandStacked");
                if (!success) MessageBox.Show("Failed to set BS mode for Rx " + r.ToString());
            }

            // Set up the signal generator...

            Binstrument_Click(bSG1, new EventArgs());
            decimal power1 = radioButton1.Checked ? numSGPower1.Value : numSGPower2.Value;
            decimal fixtureLoss1 = numFixtureLoss1.Value;
            SetupSignalGenerator(bSG1, power1 + fixtureLoss1);
            int SigGen1 = instrumentLog.FindIndex(element => element == "SG1");

            // START THE MEASUREMENTS...

            foreach (int r in ports)
            {
                // Route the proper port to the spectrum analyser...

                if (automaticSwitching)
                {
                    int switchNo = IdentifySwitch(r);
                    for (int w = 1; w <= 3; w++)
                    {
                        if (w == r) RouteSignal(w, true);
                        else RouteSignal(w, false);
                    }

                }
                else MessageBox.Show("Connect Rx port " + r.ToString());

                foreach (string l in LNBs)
                {
                    if (chIntervene.Checked)
                        MessageBox.Show("Attach LNB " + l);
                    foreach (bool isUpper in new[] { false, true })
                    {
                        foreach (string b in bands)
                        {
                            // Set up the Destacker
                            bool success = FSK.Command38(usb, 'D', r, l, null, null, isUpper, b);

                            // Set up the spectrum analyser...
                            int[] specAnSpan = SetupSpectrumAnalyser(test, cbRbw.SelectedItem.ToString(), b);
                            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");

                            // Set the signal generator start/stop frequencies

                            decimal startF = isUpper ? 1650 : 950;
                            decimal stopF = isUpper ? 2150 : 1450;
                            decimal stepF = numStep.Value;

                            // Prepare a list of 'known-spurs'...
                            List<decimal> knownSpurs = new List<decimal>();

                            knownSpurs.Add(1650m);
                            knownSpurs.Add(specAnSpan[0] + 1);
                            knownSpurs.Add(specAnSpan[1] - 1);
                            List<decimal[]> results = MeasureSpurs(SigGen1, SpecAn1, test, ref startF, ref stopF, ref stepF, knownSpurs);

                            string[] data;
                            if (test != "In-band") data = new string[] { results[0][0].ToString(), results[1][0].ToString(), results[1][1].ToString("F1") };
                            else
                            {
                                string dBc_WorstSpur = (results[1][1] - results[2][1]).ToString("F1");
                                data = new string[] { results[0][0].ToString(), results[1][0].ToString(), results[2][0].ToString(), results[2][1].ToString("F1"), dBc_WorstSpur };
                            }

                            string dataRow = l.PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4) + b.PadRight(7);

                            foreach (string s in data)
                            {
                                dataRow += s.PadRight(17);
                            }

                            // Save results to a file...

                            File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                            progressBar1.Value += 1;
                        }
                    }
                }
            }
            File.AppendAllText(resultsFile, Environment.NewLine + "Finished at " + DateTime.Now);
            RenameResultsFile("TBS", ports, serialNumber, test, temperature);
            System.Media.SystemSounds.Beep.Play();
        }

        private List<decimal[]> MeasureSpurs(int SigGen1, int SpecAn1, string test, ref decimal startF, ref decimal stopF, ref decimal stepF, List<decimal> knownSpurs)
        {
            // Define some variables for the results...
            decimal[] markerPeak;
            decimal[] markerSpur = new decimal[2];
            decimal[] worstWanted = new decimal[2] { 0, -999 };
            decimal[] worstUnwanted = new decimal[2] { 0, -999 };
            decimal worstInput = 0;
            knownSpurs.Add(0); // place-holder for later

            for (decimal f = startF; f < stopF; f += stepF)
            {
                // Set the Signal Generator frequency...
                instruments[SigGen1].WriteString(":FREQ " + f + "E6;*OPC?");
                _ = instruments[SigGen1].ReadString();

                // Wait for the sig gen...
                Thread.Sleep(0);    // 'Yield to other threads'

                // Do a sweep and wait for it to finish...
                instruments[SpecAn1].WriteString("INIT:IMM;*OPC?");
                _ = instruments[SpecAn1].ReadString();

                // Find the biggest signal & add to the ignore list...

                instruments[SpecAn1].WriteString("CALC:MARK1:MAX");
                markerPeak = ReadMarker(instruments[SpecAn1]);   // MHz and dB
                knownSpurs.Insert(knownSpurs.Count - 1, markerPeak[0]);

                int margin = 1;

                if (test == "In-band")
                {
                    for (int i = 0; i < 100; i++)
                    {
                        instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                        instruments[SpecAn1].WriteString("SYST:ERR?");
                        string error = instruments[SpecAn1].ReadString();
                        if (error.Contains("No Peak Found"))
                        {
                            markerSpur[0] = 999;
                            markerSpur[1] = -99;
                        }
                        markerSpur = ReadMarker(instruments[SpecAn1]);  // MHz and dB

                        // Is the frequency in the 'known-spurs' list?

                        if (IsKnownSpur(markerSpur[0], knownSpurs, margin))
                        {
                            continue;
                        }
                        else break;
                    }
                    //if (markerSpur[1] > -25) MessageBox.Show("");
                    if (markerSpur[1] > worstUnwanted[1])
                    {
                        worstUnwanted = markerSpur;
                        worstWanted = markerPeak;
                        worstInput = f;
                    }
                }

                else
                {
                    while (true)
                    {
                        if (IsKnownSpur(markerPeak[0], knownSpurs, margin))
                        {
                            instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                            markerPeak = ReadMarker(instruments[SpecAn1]);
                            continue;
                        }
                        else break;
                    }

                    if (markerPeak[1] > worstWanted[1])
                    {
                        worstWanted = markerPeak;
                        worstInput = f;
                    }
                }
            }
            List<decimal[]> results = new List<decimal[]>();
            results.Add(new decimal[] { worstInput, 0 });
            results.Add(worstWanted);
            results.Add(worstUnwanted);
            return results;
        }
        private char PreliminaryActions(string type, out string test, out List<int> ports, out List<string> LNBs, out List<string> bands, out string serialNumber)
        {
            // Reset the progress bar to 0
            progressBar1.Value = 0;

            // Establish the temperature
            char temperature = '?';
            Form2 f = new Form2();
            var result = f.ShowDialog();
            if (result == DialogResult.OK)
            {
                temperature = f.t.ToUpper()[0];
            }

            // Sort out the results file...
            if (File.Exists(resultsFile))
            {
                DialogResult dialogResult = MessageBox.Show("Delete existing results file " + resultsFile + " ?", "?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes) File.Delete(resultsFile);
            }

            // Which tests are we performing?
            switch (tabControl1.SelectedTab.Name)
            {
                case "tpSpurs":
                    {
                        var relevantButton = gbTest.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked == true);
                        test = relevantButton.Text;
                        break;
                    }
                case "tpIM3":
                    {
                        test = "IM3";
                        break;
                    }
                case "tpMER":
                    {
                        test = "MER";
                        break;
                    }
                default:
                    {
                        test = "Unknown";
                        break;
                    }
            }

            // Which rx are we measuring?
            ports = new List<int>();
            if (chRx1.Checked == true) ports.Add(1);
            if (chRx2.Checked == true) ports.Add(2);
            if (ports.Count < 1) MessageBox.Show("No Rx ports selected");

            // Which LNBs are we testing?
            LNBs = new List<string>();
            var checkedBoxes = gbInclude.Controls.OfType<CheckBox>().Where(r => r.Checked);
            foreach (CheckBox ch in checkedBoxes)
            {
                LNBs.Add(ch.Tag.ToString());
            }
            LNBs.Sort();

            // Which bands are we measuring?
            bands = new List<string>();
            checkedBoxes = gbTo.Controls.OfType<CheckBox>().Where(r => r.Checked);
            foreach (CheckBox ch in checkedBoxes)
            {
                bands.Add(ch.Text);
            }

            // Unlock in order to access serial number later...
            bool success = FSK.Unlock(usb, 'F', ports[0]);
            if (!success) MessageBox.Show("Failed to unlock at port " + ports[0].ToString());

            // Get the serial number...
            serialNumber = FSK.ReadSerialNumbers(usb, 'F', ports[0])[1];
            if (string.IsNullOrEmpty(serialNumber)) MessageBox.Show("Failed to get serial number");

            decimal power = radioButton1.Checked ? numSGPower1.Value : numSGPower2.Value;

            File.AppendAllText(resultsFile, DateTime.Now.ToString() + Environment.NewLine);
            File.AppendAllText(resultsFile, serialNumber + Environment.NewLine);
            File.AppendAllText(resultsFile, "Test is: " + test + Environment.NewLine);
            if (!test.ToUpper().Contains("MER"))
            {
                File.AppendAllText(resultsFile, "Power into Stacker is " + power + Environment.NewLine);
                File.AppendAllText(resultsFile, "Frequency step is " + numStep.Value.ToString() + " MHz" + Environment.NewLine);

            }
            File.AppendAllText(resultsFile, (chIntervene.Checked ? "Intervention" : "No intervention") + Environment.NewLine);
            File.AppendAllText(resultsFile, type + " mode" + Environment.NewLine);
            File.AppendAllText(resultsFile, "Temperature is " + temperature + Environment.NewLine + Environment.NewLine);
            return temperature;
        }

        private bool IsKnownSpur(decimal f, List<decimal> knownSpurs, decimal margin)
        {
            foreach (decimal ks in knownSpurs)
            {
                if (Math.Abs(f - ks) < margin)
                {
                    return true;
                }
            }

            return false;
        }
        private void RenameResultsFile(string type, List<int> rxPorts, string serialNo, string test, char temperature)
        {
            string ports = string.Empty;
            foreach (int p in rxPorts)
            {
                ports += p.ToString() + " ";
            }
            ports.TrimEnd(' ');

            decimal power = radioButton1.Checked ? numSGPower1.Value : numSGPower2.Value;

            string newFileName = "#" + serialNo.Substring(serialNo.Length - 3, 3) + " " + type;
            newFileName += " [" + temperature + "] [" + ports + "] [" + power + "] [" + test + "].txt";
            newFileName = Path.Combine(myDesktop, newFileName);

            try
            {
                File.Move(resultsFile, newFileName);
            }

            catch
            {
                MessageBox.Show("File already exists. Please rename or move existing file.");
                File.Move(resultsFile, newFileName);
            }

            // Backup the file to a network drive...

            string backupDir = @"\\server-Ste-fs\shared-data\projects\SW48 DON GenII\Qualification\Removal of ADA amplifiers\Harmonic & non-harmonic spurious\WfH tests\";
            // Remove path from the file name.
            string fName = newFileName.Substring(myDesktop.Length + 1);

            try
            {
                // Will not overwrite if the destination file already exists.
                File.Copy(Path.Combine(myDesktop, fName), Path.Combine(backupDir, fName));
            }

            // Catch exception if the file was already copied.
            catch (IOException copyError)
            {
                Console.WriteLine(copyError.Message);
            }
        }
        private int[] SetupSpectrumAnalyser(string test, string resBw, string rxBand = null, params int[] centreSpan)
        {
            int[] startStop = new int[2];

            switch (test)
            {
                case "In-band":
                    {
                        if (rxBand == "Low")
                        {
                            startStop = new int[] { 949, 1451 };
                        }

                        else if (rxBand == "Mid")
                        {
                            startStop = new int[] { 1649, 2151 };
                        }

                        else if (rxBand == "High")
                        {
                            startStop = new int[] { 2499, 3001 };
                        }

                        else
                        {
                            startStop = new int[] { 949, 2151 };
                        }
                        break;
                    }
                case "50 to 875 MHz":
                    {
                        startStop = new int[] { 50, 875 };
                        break;
                    }
                case "3 to 7.5 GHz":
                    {
                        startStop = new int[] { 3000, 7500 };
                        break;
                    }
                case "3 to 3.05 GHz":
                    {
                        startStop = new int[] { 3000, 3050 };
                        break;
                    }
                case "3.05 to 7.5 GHz":
                    {
                        startStop = new int[] { 3050, 7500 };
                        break;
                    }
                case "11.7 to 12.7 GHz":
                    {
                        startStop = new int[] { 11700, 12700 };
                        break;
                    }
                case "IM3-TBS":
                    {
                        startStop[0] = centreSpan[0] - centreSpan[1] / 2;
                        startStop[1] = centreSpan[0] + centreSpan[1] / 2;
                        break;
                    }
                case "IM3-DCS":
                    {
                        startStop[0] = centreSpan[0] - centreSpan[1] / 2;
                        startStop[1] = centreSpan[0] + centreSpan[1] / 2;
                        break;
                    }
            }

            Binstrument_Click(bSA, new EventArgs());
            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");
            if (SpecAn1 == -1)
            {
                MessageBox.Show("Did not find a connected Spectrum Analyser");
            }

            if (chSaDisplay.Checked == true)
                instruments[SpecAn1].WriteString("DISP:ENAB 0");
            else instruments[SpecAn1].WriteString("DISP:ENAB 1");

            instruments[SpecAn1].WriteString("POW:ATT 0");

            string command = "BAND " + resBw + " KHz";
            instruments[SpecAn1].WriteString(command);
            command = "FREQ:STAR " + startStop[0].ToString() + " MHz";
            instruments[SpecAn1].WriteString(command);
            command = "FREQ:STOP " + startStop[1].ToString() + " MHz";
            instruments[SpecAn1].WriteString(command);
            instruments[SpecAn1].WriteString("SWE:POIN 8192");
            instruments[SpecAn1].WriteString("INIT:CONT OFF;*OPC?");   // Set single sweep mode and wait until finished...
            _ = instruments[SpecAn1].ReadString();
            instruments[SpecAn1].WriteString("CALC:MARK:PEAK:EXC 1 dB;*OPC?");    // Allows NEXT PEAK to select noise
            _ = instruments[SpecAn1].ReadString();
            instruments[SpecAn1].WriteString("CALC:MARK:PEAK:THR -110 dBm;*OPC?");
            _ = instruments[SpecAn1].ReadString();

            return startStop;
        }
        private void SetupSignalGenerator(Button b, decimal p)
        {
            int SigGen = instrumentLog.FindIndex(element => element == b.Tag.ToString());
            if (SigGen == -1)
            {
                MessageBox.Show("Did not find a connected Signal Generator");
            }

            // Set the Signal Generator power...
            instruments[SigGen].WriteString(":POW " + p);

            // Turn it on...
            instruments[SigGen].WriteString(":OUTP:STATE ON");
        }

        private decimal[] ReadMarker(FormattedIO488 specAn)
        {
            // Measure the marker frequency...
            specAn.WriteString("CALC:MARK1:X?");
            decimal markerFrequency = decimal.Parse(specAn.ReadString().Replace(@"\n", string.Empty), NumberStyles.Float) / 1000000;

            // Measure the marker power...
            specAn.WriteString("CALC:MARK1:Y?");
            decimal markerPower = decimal.Parse(specAn.ReadString().Replace(@"\n", string.Empty), NumberStyles.Float);
            return new decimal[] { markerFrequency, markerPower };
        }

        private void BSpurTestDBS_Click(object sender, EventArgs e)
        {
            // In DBS, LNB-A => Rx1, LNB-B => Rx2 and LNB-C => Rx 3. No other combinations are valid.

            char temperature = PreliminaryActions("DBS", out string test, out List<int> ports, out List<string> LNBs, out List<string> bands, out string serialNumber);
            ports = new List<int> { 1, 2, 3 };  // Over-ride to do all three ports
            LNBs = new List<string> { "A", "B", "C" };
            bands.Remove("High");   // Works whether present or not

            // Set up the progress bar...
            progressBar1.Value = 0;
            progressBar1.Maximum = ports.Count * 2;

            File.AppendAllText(resultsFile,
               "LNB".PadRight(4) + "Pol.".PadRight(6) + "Rx".PadRight(4) + "Band".PadRight(7) +
                   "Input_MHz".PadRight(17) +
                   (test == "In-band" ? "Wanted_MHz".PadRight(17) : string.Empty) +
                   "Spur_MHz".PadRight(17) +
                   "Spur_dBm".PadRight(17) +
                   (test == "In-band" ? "Spur_dBc" : string.Empty) + Environment.NewLine);


            // Set up the signal generator...

            Binstrument_Click(bSG1, new EventArgs());
            decimal power1 = radioButton1.Checked ? numSGPower1.Value : numSGPower2.Value;
            decimal fixtureLoss1 = numFixtureLoss1.Value;
            SetupSignalGenerator(bSG1, power1 + fixtureLoss1);
            int SigGen1 = instrumentLog.FindIndex(element => element == "SG1");

            foreach (int r in ports)
            {
                // Route the proper port to the spectrum analyser...
                if (automaticSwitching)
                {
                    int switchNo = IdentifySwitch(r);
                    for (int w = 1; w <= 3; w++)
                    {
                        if (w == r) RouteSignal(w, true);
                        else RouteSignal(w, false);
                    }

                }
                else MessageBox.Show("Connect Rx port " + r.ToString());

                string l = LNBs[r - 1];
                if (chIntervene.Checked)
                    MessageBox.Show("Attach LNB " + l);

                foreach (bool isUpper in new[] { false, true })
                {
                    string b = isUpper ? "Mid" : "Low";

                    // Set up the spectrum analyser...
                    int[] specAnSpan = SetupSpectrumAnalyser(test, cbRbw.SelectedItem.ToString(), b);
                    int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");


                    // Set the signal generator start/stop frequencies
                    decimal startF = isUpper ? 1650 : 950;
                    decimal stopF = isUpper ? 2150 : 1450;
                    decimal stepF = numStep.Value;

                    //Populate a 'known-spur' list...
                    List<decimal> knownSpurs = new List<decimal>();

                    knownSpurs.Add(1650m);
                    knownSpurs.Add(specAnSpan[0] + 1);
                    knownSpurs.Add(specAnSpan[1] - 1);

                    List<decimal[]> results = MeasureSpurs(SigGen1, SpecAn1, test, ref startF, ref stopF, ref stepF, knownSpurs);

                    string[] data;
                    if (test != "In-band") data = new string[] { results[0][0].ToString(),
                                results[1][0].ToString(), results[1][1].ToString("F1") };
                    else
                    {
                        string dBc_WorstSpur = (results[1][1] - results[2][1]).ToString("F1");
                        data = new string[] { results[0][0].ToString(),
                                    results[1][0].ToString(),
                                    results[2][0].ToString(), results[2][1].ToString("F1"), dBc_WorstSpur };
                    }

                    string dataRow = l.PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4) + b.PadRight(7);

                    foreach (string s in data)
                    {
                        dataRow += s.PadRight(17);
                    }

                    // Save results to a file...

                    File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                    progressBar1.Value += 1;
                }
            }
            File.AppendAllText(resultsFile, Environment.NewLine + "Finished at " + DateTime.Now);
            RenameResultsFile("DBS", ports, serialNumber, test, temperature);

            System.Media.SystemSounds.Beep.Play();
        }

        private void BDcsSpurTest_Click(object sender, EventArgs e)
        {
            char temperature = PreliminaryActions("DCS", out string test, out List<int> ports, out List<string> LNBs, out List<string> bands, out string serialNumber);

            File.AppendAllText(resultsFile,
                "LNB".PadRight(4) + "Pol.".PadRight(6) + "Rx".PadRight(4) +
                    "Input_MHz".PadRight(17) +
                    (test == "In-band" ? string.Empty : "Unwanted".PadRight(17)) +
                    (test == "In-band" ? "Worst_MHz".PadRight(17) : string.Empty) +
                    "Worst_dBm".PadRight(17) +
                    (test == "In-band" ? "Worst_dBc" : string.Empty) + Environment.NewLine);

            // Set each Rx to CHANNEL-STACKED and disable TDMA polling...
            foreach (int r in ports)
            {
                bool success = FSK.SetStackingMode(usb, r, "ChannelStacked");
                if (!success) MessageBox.Show("Failed to set C/S mode for Rx " + r.ToString());

                success = FSK.DisableTdmaPolling(usb, 'F', r, true);
                if (!success) MessageBox.Show("Failed to disable TDMA polling for Rx " + r.ToString());

                success = FSK.DisableTdmaPolling(usb, 'F', r, true);
                if (!success) MessageBox.Show("Failed to disable TDMA polling for Rx " + r.ToString());
                else MessageBox.Show("Successfully disabled TDMA polling for Rx " + r.ToString());
            }

            // Set up the signal generator...

            Binstrument_Click(bSG1, new EventArgs());
            decimal power1 = radioButton1.Checked ? numSGPower1.Value : numSGPower2.Value;
            decimal fixtureLoss1 = numFixtureLoss1.Value;
            SetupSignalGenerator(bSG1, power1 + fixtureLoss1);

            int SigGen1 = instrumentLog.FindIndex(element => element == "SG1");

            // Set up S/A
            SetupSpectrumAnalyser(test, cbRbw.SelectedItem.ToString());
            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");

            // START THE MEASUREMENTS...

            foreach (int r in ports)
            {
                // Route the proper port to the spectrum analyser...
                if (automaticSwitching)
                {
                    int switchNo = IdentifySwitch(r);
                    for (int w = 1; w <= 3; w++)
                    {
                        if (w == r) RouteSignal(w, true);
                        else RouteSignal(w, false);
                    }

                }
                else MessageBox.Show("Connect Rx port " + r.ToString());

                foreach (string l in LNBs)
                {
                    if (chIntervene.Checked)
                        MessageBox.Show("Attach LNB " + l);

                    // Allocate all channels to current Rx port...

                    for (int ch = 0; ch < 18; ch++)
                    {
                        _ = FSK.DeallocateChannel(usb, 'F', r, ch);
                        bool success = FSK.AllocateChannel(usb, 'F', r, ch);
                        if (!success) MessageBox.Show("Failed to allocate channel " + ch.ToString());
                    }

                    // read grid & write header to file...

                    List<int> grid = FSK.RequestUserBlockGrid(usb, 'F', r).ToList<int>();
                    List<decimal> knownSpurs = new List<decimal>();
                    for (int i = 0; i < 18; i++)
                    {
                        knownSpurs.Add(grid[0] + i * grid[1]);
                    }

                    foreach (bool isUpper in new[] { false, true })
                    {
                        // Set the signal generator start/stop frequencies

                        decimal startF = isUpper ? 1650 : 950;
                        decimal stopF = isUpper ? 2150 : 1450;
                        decimal stepF = numStep.Value;

                        // Set the progress bar up...
                        int noMeasurements = 1 + (int)Math.Floor((stopF - startF) / stepF);
                        progressBar1.Maximum = LNBs.Count * ports.Count * 2 * noMeasurements;

                        for (decimal f = startF; f < stopF; f += stepF)
                        {
                            // Set up the channel frequencies...
                            int KHz = Convert.ToInt32(1000 * f);
                            if (l == "C" && isUpper)
                                KHz += 28500;       // Neither understand nor like this. Must add 28.5 MHz!

                            for (int ch = 0; ch < 18; ch++)
                            {
                                bool success = FSK.Command38(usb, 'F', r, l, ch, KHz, isUpper, null);
                                if (!success) MessageBox.Show("Failed to allocate frequency to channel " + ch.ToString());
                            }
                            decimal dummyStop = f + 1;
                            List<decimal[]> results = MeasureSpurs(SigGen1, SpecAn1, test, ref f, ref dummyStop, ref stepF, knownSpurs);

                            string[] data;
                            if (test != "In-band") data = new string[] { results[0][0].ToString(),
                                results[1][0].ToString(), results[1][1].ToString("F1") };
                            else
                            {
                                string dBc_WorstSpur = (results[1][1] - results[2][1]).ToString("F1");
                                data = new string[] { results[0][0].ToString(),
                                    results[2][0].ToString(), results[2][1].ToString("F1"), dBc_WorstSpur };
                            }

                            string dataRow = l.PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4);

                            foreach (string s in data)
                            {
                                dataRow += s.PadRight(17);
                            }

                            // Save results to a file...

                            File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                            progressBar1.Value += 1;
                        }
                    }
                }
            }
            File.AppendAllText(resultsFile, Environment.NewLine + "Finished at " + DateTime.Now);
            RenameResultsFile("DCS", ports, serialNumber, test, temperature);
            System.Media.SystemSounds.Beep.Play();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var checkedBoxes = groupBox2.Controls.OfType<CheckBox>().Where(r => r.Checked);
            int rx = 1;

            foreach (CheckBox b in checkedBoxes)
            {
                rx = Int32.Parse(b.Tag.ToString());
                //if (FSK.TdmaPollingIsOn(usb, 'F', rx)) rtb1.AppendText("TDMA polling is Enabled on port " + rx.ToString() + Environment.NewLine);
                //else rtb1.AppendText("TDMA polling is Disabled on port " + rx.ToString() + Environment.NewLine);
            }
        }
        private void RouteSignal(int rx, bool on)
        {
            Binstrument_Click(bPickering, new EventArgs());
            string card = numCard.Value.ToString();
            string state = on ? "Close" : "Open";

            string switchNo = IdentifySwitch(rx).ToString();
            int pickering = instrumentLog.FindIndex(element => element == "Pickering");
            instruments[pickering].WriteString(state + " " + card + "," + switchNo);
        }

        private void Pickering_CheckedChanged(object sender, EventArgs e)
        {
            // Only respond to the newly checked button, not to unchecked ones... 
            if (!(sender as RadioButton).Checked) return;

            var rb = gbPickering.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked);
            int rx = int.Parse(rb.Tag.ToString());

            int switchNo = IdentifySwitch(rx);

            foreach (int r in new int[] { 1, 2, 3 })
            {
                if (r == switchNo) RouteSignal(switchNo, true);
                else RouteSignal(r, false);
            }
        }
        private int IdentifySwitch(int rx)
        {
            var nud = gbPickering.Controls.OfType<NumericUpDown>().FirstOrDefault(r => r.Tag.ToString() == rx.ToString());
            int switchNo = int.Parse(nud.Value.ToString());
            return switchNo;
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            TabPage selectedTab = (sender as TabControl).SelectedTab;
            switch (selectedTab.Name)
            {
                case "tpIM3":
                case "tpSpurs":
                    {
                        gbMeasurements.Parent = selectedTab;
                        gbMisc.Parent = selectedTab;
                        progressBar1.Parent = selectedTab;
                        break;
                    }
                case "tpMER":
                    {
                        gbMeasurements.Parent = selectedTab;
                        progressBar1.Parent = selectedTab;
                        break;
                    }
            }
        }

        private void bIMTbs_Click(object sender, EventArgs e)
        {
            char temperature = PreliminaryActions("IM3-TBS", out string test, out List<int> ports, out List<string> lnbs, out List<string> bands, out string serialNumber);
            test = "IM3-TBS";

            File.AppendAllText(resultsFile,
                "LNB".PadRight(4) + "Pol.".PadRight(6) + "Rx".PadRight(4) + "Band".PadRight(7) +
                    "Input_MHz".PadRight(17) +
                    "Wanted_MHz".PadRight(17) +
                    "LSB_MHz".PadRight(17) +
                    "LSB_dBm".PadRight(17) +
                    "USB_MHz".PadRight(17) +
                    "USB_dBm".PadRight(17) +
                    "LSB_dBc".PadRight(17) +
                    "USB_dBc" + Environment.NewLine);

            // Sort out the ProgressBar...
            progressBar1.Maximum = lnbs.Count * ports.Count * 2 * bands.Count;

            // Set each Rx to BANDSTACKED
            foreach (int r in ports)
            {
                bool success = FSK.SetStackingMode(usb, r, "BandStacked");
                if (!success) MessageBox.Show("Failed to set BS mode for Rx " + r.ToString());
            }

            // Set up the signal generators...

            Binstrument_Click(bSG1, new EventArgs());
            int SigGen1 = instrumentLog.FindIndex(element => element == "SG1");
            decimal power1 = radioButton1.Checked ? numSGPower1.Value : numSGPower2.Value;
            decimal fixtureLoss1 = numFixtureLoss1.Value;
            SetupSignalGenerator(bSG1, power1 + fixtureLoss1);
            decimal lowerTone = nudSgStart.Value;
            instruments[SigGen1].WriteString(":FREQ " + lowerTone.ToString() + "E6;*OPC?");
            _ = instruments[SigGen1].ReadString();
            Thread.Sleep(0);    // 'Yield to other threads'

            Binstrument_Click(bSG2, new EventArgs());
            int SigGen2 = instrumentLog.FindIndex(element => element == "SG2");
            decimal power2 = power1 + numSg2Power.Value;
            decimal fixtureLoss2 = numFixtureLoss2.Value;
            SetupSignalGenerator(bSG2, power2 + fixtureLoss2);
            decimal upperTone = nudSgStart.Value + numSG2DeltaF.Value;
            instruments[SigGen2].WriteString(":FREQ " + upperTone.ToString() + "E6;*OPC?");
            _ = instruments[SigGen2].ReadString();
            Thread.Sleep(0);    // 'Yield to other threads'

            string rbw = cbRbw.SelectedItem.ToString();

            foreach (int r in ports)
            {
                // Route the proper port to the spectrum analyser...

                try
                {
                    var rb = gbPickering.Controls.OfType<RadioButton>().FirstOrDefault(q => q.Tag.ToString() == r.ToString());
                    (rb as RadioButton).Checked = true;
                    Pickering_CheckedChanged((rb as RadioButton), new EventArgs());
                }
                catch
                {
                    MessageBox.Show("Connect Rx " + r.ToString());
                }

                // Define some variables for the results...
                decimal[] markerPeak;
                decimal[] markerIM3Lower = new decimal[2];
                decimal[] markerIM3Upper = new decimal[2];

                foreach (string l in lnbs)
                {
                    if (chIntervene.Checked)
                        MessageBox.Show("Attach LNB " + l);

                    foreach (bool isUpper in new[] { false, true })
                    {
                        foreach (string b in bands)
                        {
                            // Set up the Destacker
                            bool success = FSK.Command38(usb, 'D', r, l, null, null, isUpper, b);

                            // Set the signal generator frequencies...
                            lowerTone = isUpper ? 1900 : 1200;
                            instruments[SigGen1].WriteString(":FREQ " + lowerTone.ToString() + "E6;*OPC?");
                            _ = instruments[SigGen1].ReadString();
                            Thread.Sleep(0);    // 'Yield to other threads'

                            upperTone = lowerTone + numSG2DeltaF.Value;
                            instruments[SigGen2].WriteString(":FREQ " + upperTone.ToString() + "E6;*OPC?");
                            _ = instruments[SigGen2].ReadString();
                            Thread.Sleep(0);    // 'Yield to other threads'

                            // Calculate where to expect the signals
                            decimal expectedUpper = CalculateOutputSignalFrequency(upperTone, isUpper, b);
                            decimal expectedLower = CalculateOutputSignalFrequency(lowerTone, isUpper, b);
                            decimal deltaF = Math.Abs(expectedUpper - expectedLower);
                            decimal expectedLsb = expectedLower - deltaF;
                            decimal expectedUsb = expectedUpper + deltaF;

                            // Set up the spectrum analyser...
                            int centreFrequency = Convert.ToInt32((expectedUpper + expectedLower) / 2);
                            int span = (int)deltaF * 4;
                            int[] specAnSpan = SetupSpectrumAnalyser("IM3-TBS", rbw, b, new int[] { centreFrequency, span });
                            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");

                            // Do a sweep and wait for it to finish...
                            instruments[SpecAn1].WriteString("INIT:IMM;*OPC?");
                            _ = instruments[SpecAn1].ReadString();

                            // Find the 'expected lower' signal...
                            instruments[SpecAn1].WriteString("CALC:MARK1:MAX");
                            markerPeak = ReadMarker(instruments[SpecAn1]);   // MHz and dB
                            if (Math.Abs(markerPeak[0] - expectedLower) > deltaF / 2)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerPeak = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            decimal margin = 3;
                            // Find the LSB IM3 product...

                            while (Math.Abs(markerIM3Lower[0] - expectedLsb) > margin)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerIM3Lower = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            // Find the 'expected upper' signal...
                            instruments[SpecAn1].WriteString("CALC:MARK1:MAX");
                            markerPeak = ReadMarker(instruments[SpecAn1]);   // MHz and dB
                            if (Math.Abs(markerPeak[0] - expectedUpper) > deltaF / 2)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerPeak = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            // Find the USB IM3 product...

                            while (Math.Abs(markerIM3Upper[0] - expectedUsb) > margin)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerIM3Upper = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            string[] data = new string[] { lowerTone.ToString(), markerPeak[0].ToString("F1"), markerIM3Lower[0].ToString("F1"), markerIM3Lower[1].ToString("F1"), markerIM3Upper[0].ToString("F1"), markerIM3Upper[1].ToString("F1") };
                            string dataRow = l.PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4) + b.PadRight(7);

                            foreach (string s in data)
                            {
                                dataRow += s.PadRight(17);
                            }

                            dataRow += (markerPeak[1] - markerIM3Lower[1]).ToString("F1").PadRight(17);
                            dataRow += (markerPeak[1] - markerIM3Upper[1]).ToString("F1");

                            // Save results to a file...

                            File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                            progressBar1.Value += 1;
                        }
                    }
                }
            }
            File.AppendAllText(resultsFile, Environment.NewLine + "Finished at " + DateTime.Now);
            RenameResultsFile("IM3-TBS", ports, serialNumber, test, temperature);
            System.Media.SystemSounds.Beep.Play();
        }
        private void bIMDcs_Click(object sender, EventArgs e)
        {
            char temperature = PreliminaryActions("TBS", out string test, out List<int> ports, out List<string> lnbs, out List<string> bands, out string serialNumber);
            test = "IM3-DCS";

            File.AppendAllText(resultsFile,
                "LNB".PadRight(4) + "Pol.".PadRight(6) + "Rx".PadRight(4) +
                    "Input_MHz".PadRight(17) +
                    "Wanted_MHz".PadRight(17) +
                    "LSB_MHz".PadRight(17) +
                    "LSB_dBm".PadRight(17) +
                    "USB_MHz".PadRight(17) +
                    "USB_dBm".PadRight(17) +
                    "LSB_dBc".PadRight(17) +
                    "USB_dBc" + Environment.NewLine);

            // Define some variables for the results...
            decimal[] markerPeak;
            decimal[] markerIM3Lower = new decimal[2];
            decimal[] markerIM3Upper = new decimal[2];
            decimal deltaF = numSG2DeltaF.Value;

            // Sort out the ProgressBar...
            progressBar1.Maximum = lnbs.Count * ports.Count * 2 * bands.Count;

            // Set each Rx to CHANNEL-STACKED and disable TDMA polling...
            foreach (int r in ports)
            {
                bool success = FSK.SetStackingMode(usb, r, "ChannelStacked");
                if (!success) MessageBox.Show("Failed to set C/S mode for Rx " + r.ToString());

                success = FSK.DisableTdmaPolling(usb, 'F', r, true);
                if (!success) MessageBox.Show("Failed to disable TDMA polling for Rx " + r.ToString());

                success = FSK.DisableTdmaPolling(usb, 'F', r, true);
                if (!success) MessageBox.Show("Failed to disable TDMA polling for Rx " + r.ToString());
                else MessageBox.Show("Successfully disabled TDMA polling for Rx " + r.ToString());
            }

            // Set up the signal generators...

            Binstrument_Click(bSG1, new EventArgs());
            int SigGen1 = instrumentLog.FindIndex(element => element == "SG1");
            decimal power1 = radioButton1.Checked ? numSGPower1.Value : numSGPower2.Value;
            decimal fixtureLoss1 = numFixtureLoss1.Value;
            SetupSignalGenerator(bSG1, power1 + fixtureLoss1);
            decimal lowerTone = nudSgStart.Value;
            instruments[SigGen1].WriteString(":FREQ " + lowerTone.ToString() + "E6;*OPC?");
            _ = instruments[SigGen1].ReadString();
            Thread.Sleep(0);    // 'Yield to other threads'

            Binstrument_Click(bSG2, new EventArgs());
            int SigGen2 = instrumentLog.FindIndex(element => element == "SG2");
            decimal power2 = power1 + numSg2Power.Value;
            decimal fixtureLoss2 = numFixtureLoss2.Value;
            SetupSignalGenerator(bSG2, power2 + fixtureLoss2);
            decimal upperTone = nudSgStart.Value + deltaF;
            instruments[SigGen2].WriteString(":FREQ " + upperTone.ToString() + "E6;*OPC?");
            _ = instruments[SigGen2].ReadString();
            Thread.Sleep(0);    // 'Yield to other threads'

            Binstrument_Click(bSA, new EventArgs());
            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");
            string rbw = cbRbw.SelectedItem.ToString();

            foreach (int r in ports)
            {
                // Route the proper port to the spectrum analyser...
                if (automaticSwitching)
                {
                    int switchNo = IdentifySwitch(r);
                    for (int w = 1; w <= 3; w++)
                    {
                        if (w == r) RouteSignal(w, true);
                        else RouteSignal(w, false);
                    }
                }

                else MessageBox.Show("Connect Rx port " + r.ToString());

                // Allocate all channels to current Rx...

                for (int ch = 0; ch < 18; ch++)
                {
                    _ = FSK.DeallocateChannel(usb, 'F', r, ch);
                    bool success = FSK.AllocateChannel(usb, 'F', r, ch);
                    if (!success) MessageBox.Show("Failed to allocate channel " + ch.ToString());
                }

                foreach (string l in lnbs)
                {
                    if (chIntervene.Checked)
                        MessageBox.Show("Attach LNB " + l);

                    // read grid...

                    List<int> grid = FSK.RequestUserBlockGrid(usb, 'F', r).ToList<int>();
                    List<decimal> centreFrequencies = new List<decimal>();
                    int noUserBands = 16;
                    for (int i = 0; i < noUserBands; i++)
                    {
                        centreFrequencies.Add(grid[0] + i * grid[1]);
                    }

                    foreach (bool isUpper in new[] { false, true })
                    {
                        // Set the signal generator frequencies...
                        lowerTone = isUpper ? 1900 : 1200;
                        instruments[SigGen1].WriteString(":FREQ " + lowerTone.ToString() + "E6;*OPC?");
                        _ = instruments[SigGen1].ReadString();
                        Thread.Sleep(0);    // 'Yield to other threads'

                        upperTone = lowerTone + numSG2DeltaF.Value;
                        instruments[SigGen2].WriteString(":FREQ " + upperTone.ToString() + "E6;*OPC?");
                        _ = instruments[SigGen2].ReadString();
                        Thread.Sleep(0);    // 'Yield to other threads'


                        // Set the progress bar up...
                        progressBar1.Maximum = lnbs.Count * ports.Count * 2 * noUserBands;

                        // Set up the channel frequencies...
                        int KHz = Convert.ToInt32(1000 * (upperTone + lowerTone) / 2);
                        if (l == "C" && isUpper)
                            KHz += 28500;       // Neither understand nor like this. Must add 28.5 MHz!

                        for (int ch = 0; ch < 18; ch++)
                        {
                            bool success = FSK.Command38(usb, 'F', r, l, ch, KHz, isUpper, null);
                            if (!success) MessageBox.Show("Failed to allocate frequency to channel " + ch.ToString());
                        }

                        for (int f = grid[0]; f < grid[0] + 16 * grid[1]; f += grid[1])
                        {
                            // Calculate where to expect the signals
                            decimal expectedUpper = f + deltaF / 2;
                            decimal expectedLower = f - deltaF / 2;
                            decimal expectedLsb = expectedLower - deltaF;
                            decimal expectedUsb = expectedUpper + deltaF;

                            // Set up the spectrum analyser...
                            int[] specAnSpan = SetupSpectrumAnalyser("IM3-DCS", rbw, null, new int[] { f, grid[1] });

                            // Do a sweep and wait for it to finish...
                            instruments[SpecAn1].WriteString("INIT:IMM;*OPC?");
                            _ = instruments[SpecAn1].ReadString();


                            // Find the 'expected lower' signal...
                            instruments[SpecAn1].WriteString("CALC:MARK1:MAX");
                            markerPeak = ReadMarker(instruments[SpecAn1]);   // MHz and dB
                            if (Math.Abs(markerPeak[0] - expectedLower) > deltaF / 2)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerPeak = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            decimal margin = 3;
                            // Find the LSB IM3 product...

                            while (Math.Abs(markerIM3Lower[0] - expectedLsb) > margin)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerIM3Lower = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            // Find the 'expected upper' signal...
                            instruments[SpecAn1].WriteString("CALC:MARK1:MAX");
                            markerPeak = ReadMarker(instruments[SpecAn1]);   // MHz and dB
                            if (Math.Abs(markerPeak[0] - expectedUpper) > deltaF / 2)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerPeak = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            // Find the USB IM3 product...

                            while (Math.Abs(markerIM3Upper[0] - expectedUsb) > margin)
                            {
                                instruments[SpecAn1].WriteString("CALC:MARK1:MAX:NEXT");
                                markerIM3Upper = ReadMarker(instruments[SpecAn1]);  // MHz and dB
                            }

                            string[] data = new string[] { lowerTone.ToString(), markerPeak[0].ToString("F1"), markerIM3Lower[0].ToString("F1"), markerIM3Lower[1].ToString("F1"), markerIM3Upper[0].ToString("F1"), markerIM3Upper[1].ToString("F1") };
                            string dataRow = l.PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4);

                            foreach (string s in data)
                            {
                                dataRow += s.PadRight(17);
                            }

                            dataRow += (markerPeak[1] - markerIM3Lower[1]).ToString("F1").PadRight(17);
                            dataRow += (markerPeak[1] - markerIM3Upper[1]).ToString("F1");

                            // Save results to a file...

                            File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                            progressBar1.Value += 1;
                        }
                    }
                }
            }

            File.AppendAllText(resultsFile, Environment.NewLine + "Finished at " + DateTime.Now);
            RenameResultsFile("IM3-DCS", ports, serialNumber, test, temperature);
            System.Media.SystemSounds.Beep.Play();
        }
        private decimal CalculateOutputSignalFrequency(decimal frequency, bool isUpper, string band)
        {
            decimal answer = 0;
            switch (band)
            {
                case "Low":
                    {
                        answer = isUpper ? 3100 - frequency : frequency;
                        break;
                    }
                case "Mid":
                    {
                        //answer = isUpper ? frequency : frequency + 700;
                        answer = isUpper ? frequency : 3100 - frequency;
                        break;
                    }
                case "High":
                    {
                        answer = isUpper ? 4650 - frequency : frequency + 1550;
                        break;
                    }
            }
            return answer;
        }

        private void bMERTbs_Click(object sender, EventArgs e)
        {
            // Gather the basic parameters of the test...

            char temperature = PreliminaryActions("MER-TBS", out string test, out List<int> ports, out List<string> lnbs, out List<string> bands, out string serialNumber);
            string opticalPower = "-15dBm";

            File.AppendAllText(resultsFile,
                "PON_loss".PadRight(10)
                + "LNB".PadRight(4)
                + "Pol.".PadRight(6)
                + "Rx".PadRight(4)
                + "Band".PadRight(7)
                + "Input_MHz".PadRight(12)
                + "Output_MHz".PadRight(12)
                + "MER_dB" + Environment.NewLine);

            // Sort out the ProgressBar...
            progressBar1.Maximum = lnbs.Count * ports.Count * bands.Count;

            // Set each Rx to BANDSTACKED
            foreach (int r in ports)
            {
                bool success = FSK.SetStackingMode(usb, r, "BandStacked");
                if (!success) MessageBox.Show("Failed to set BS mode for Rx " + r.ToString());
            }

            // Prepare the spectrum analyser...

            Binstrument_Click(bSA, new EventArgs());
            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");

            int selectedTraceNo = Setup89601VSA();

            if (selectedTraceNo == 0)
            {
                rtbWhiteboard.AppendText("No suitable Trace found" + Environment.NewLine);
                return;
            }
            string traceName = "TRACE" + selectedTraceNo;

            // -------------------------  Setup the transponder frequencies  ------------------------------

            List<decimal> transponders = new List<decimal>();
            for (decimal tp = numFirstLower.Value; tp < 1450; tp += numLowerStep.Value)
            {
                transponders.Add(tp);
            }

            for (decimal tp = numFirstUpper.Value; tp < 2150; tp += numUpperStep.Value)
            {
                transponders.Add(tp);
            }

            progressBar1.Maximum *= transponders.Count;

            // Create a datatable for the results and bind it to the dgv

            System.Data.DataTable mers = new System.Data.DataTable();
            mers.Columns.Add("Frequency", typeof(decimal));
            mers.Columns.Add("MER(dB)", typeof(double));
            dgvMERs.DataSource = mers;
            dgvMERs.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvMERs.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            decimal specAnCentreFreq = 0;

            foreach (int r in ports)
            {
                // Route the proper port to the spectrum analyser...

                if (automaticSwitching)
                {
                    try
                    {
                        var rb = gbPickering.Controls.OfType<RadioButton>().FirstOrDefault(k => k.Tag.ToString() == r.ToString());
                        (rb as RadioButton).Checked = true;
                        Pickering_CheckedChanged((rb as RadioButton), new EventArgs());
                    }
                    catch
                    {
                        MessageBox.Show("Connect Rx " + r.ToString());
                    }
                }
                else MessageBox.Show("Connect Rx " + r.ToString());

                foreach (string l in lnbs)
                {
                    if (chIntervene.Checked)
                        MessageBox.Show("Attach LNB " + l);


                    foreach (bool isUpper in new[] { false, true })
                    {
                        foreach (string b in bands)
                        {
                            // Set up the Destacker
                            bool success = FSK.Command38(usb, 'D', r, l, null, null, isUpper, b);

                            rtbWhiteboard.AppendText("Measuring at port " + r.ToString() + ", from LNB " + l + (isUpper ? " Upper, " : " Lower, ") + "to " + b + Environment.NewLine);
                            rtbWhiteboard.Refresh();
                            rtbWhiteboard.Update();

                            //  ---------------------  Take the measurements  ---------------------------------------

                            foreach (decimal tp in transponders)
                            {
                                // Ignore frequencies out of band

                                if (isUpper && tp < numFirstUpper.Value)
                                    continue;
                                else if (!isUpper && tp > 1450)
                                    break;

                                // Calculate where the S/A needs to look...

                                specAnCentreFreq = CalculateDestackerFrequency(tp, isUpper, b);

                                string dataRow = opticalPower.PadRight(10) + l.PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4) + b.PadRight(7);

                                // Set the spectrum analyser frequency...
                                string message = string.Format(":FREQuency:CENTer {0} MHz", specAnCentreFreq);
                                instruments[SpecAn1].WriteString(message);

                                double merResult;
                                try
                                {
                                    PerformSweep(SpecAn1, 50, true);
                                    // Take a measurement...
                                    instruments[SpecAn1].WriteString(traceName + ":DATA:TABL?  'SigToNoise'");
                                    instruments[SpecAn1].WriteString(":DATA:TABL?  \"SigToNoise\"");
                                    merResult = double.Parse(instruments[SpecAn1].ReadString());
                                }
                                catch (COMException ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                    return;
                                }
                                mers.Rows.Add(tp.ToString(), Math.Round(merResult, 2));
                                dgvMERs.FirstDisplayedScrollingRowIndex = dgvMERs.RowCount - 1;
                                dgvMERs.Refresh();
                                dgvMERs.Update();

                                // Save results to a file...

                                dataRow += tp.ToString().PadRight(12) + specAnCentreFreq.ToString().PadRight(12) + Math.Round(merResult, 2);
                                File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                                progressBar1.Value += 1;
                            }
                        }
                    }
                    File.AppendAllText(resultsFile, "------------------------------------------------------------" + Environment.NewLine);
                }
            }
            RenameResultsFile("MER-TBS", ports, serialNumber, test, temperature);
            return;
        }

        private int Setup89601VSA()
        {
            Binstrument_Click(bSA, new EventArgs());
            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");

            // -----------------------------  Confirm the VSA process is running  ---------------------------------------------

            //instruments[SpecAn1].WriteString(":SYSTem:VSA:STARt;*OPC?");
            //string s = instruments[SpecAn1].ReadString();
            //rtbWhiteboard.AppendText("VSA returns " + s);

            instruments[SpecAn1].WriteString("INSTrument: SELect VSA89601");
            //instruments[SpecAn1].WriteString("INSTrument: SELect?");
            //string q = instruments[SpecAn1].ReadString();

            // --------------------  Set up the displays etc.  -----------------------------------------------------------

            instruments[SpecAn1].WriteString(":DISPlay:ENABle 1");
            instruments[SpecAn1].WriteString(":DISPlay:ENABle?");
            string q = instruments[SpecAn1].ReadString() == "1" ? "Display is on" : "Display is off";
            rtbWhiteboard.AppendText(q + Environment.NewLine);

            // Turn auto-cal off so that it doesn't run when it is not expected 
            // (which could result in timeouts when using SCPI without long enough timeout values).

            instruments[SpecAn1].WriteString(":CAL:AUTO 0");
            instruments[SpecAn1].WriteString(":CAL:AUTO?");
            q = instruments[SpecAn1].ReadString() == "1" ? "Auto-cal is on" : "Auto-cal is off";
            rtbWhiteboard.AppendText(q + Environment.NewLine);

            // ---------------------     Setup the Digital Demod Measurement   ------------------------------------

            instruments[SpecAn1].WriteString(":MEASure:CONFigure:DDEMod");
            instruments[SpecAn1].WriteString(":MEASure:CONFigure:DDEMod:NDEFault");
            instruments[SpecAn1].WriteString(":INITiate:DDEMod");

            instruments[SpecAn1].WriteString("INPut:DATA HW");   // Hardware rather than simulated input
            instruments[SpecAn1].WriteString(":INPut:DATA?");
            rtbWhiteboard.AppendText("Data is coming from " + instruments[SpecAn1].ReadString());

            instruments[SpecAn1].WriteString("DDEM:MOD PSK8");
            instruments[SpecAn1].WriteString("DDEM:MOD?");
            rtbWhiteboard.AppendText("Modulation is " + instruments[SpecAn1].ReadString());

            instruments[SpecAn1].WriteString(":DDEM:FILT:REF RECT");
            instruments[SpecAn1].WriteString(":DDEM:FILT?");
            rtbWhiteboard.AppendText("The filter is " + instruments[SpecAn1].ReadString());

            instruments[SpecAn1].WriteString(":DDEM:SRAT" + numSymbolRate.Value.ToString() + "MHZ");
            instruments[SpecAn1].WriteString(":DDEM:SRAT?");
            string reply = instruments[SpecAn1].ReadString().TrimEnd('\n').Replace("+", "");
            if (decimal.TryParse(reply, out decimal symbolRate))
            {
                symbolRate /= 1E6m;
                rtbWhiteboard.AppendText("The symbol rate is " + symbolRate + " MS/s" + Environment.NewLine);
            }
            else rtbWhiteboard.AppendText("The symbol rate is not readable" + Environment.NewLine);

            instruments[SpecAn1].WriteString(":FREQ:CENT?");
            decimal centre = decimal.Parse(instruments[SpecAn1].ReadString()) / 1E6m;
            rtbWhiteboard.AppendText("The centre frequency is " + centre.ToString() + " MHz" + Environment.NewLine);
            instruments[SpecAn1].WriteString(":FREQuency:SPAN 60 MHz");
            instruments[SpecAn1].WriteString(":FREQ:SPAN?");
            decimal span = decimal.Parse(instruments[SpecAn1].ReadString()) / 1E6m;
            rtbWhiteboard.AppendText("The frequency span is " + span.ToString() + " MHz" + Environment.NewLine);
            rtbWhiteboard.Refresh();
            rtbWhiteboard.Update();

            // ---------------------------   Seek the MER Trace  -----------------------------------------------------

            instruments[SpecAn1].WriteString("TRACe:COUNt?");
            int traceCount = int.Parse(instruments[SpecAn1].ReadString());
            int selectedTraceNo = 0;   // NB not a valid trace
            for (int t = 1; t <= traceCount; t++)
            {
                try
                {
                    instruments[SpecAn1].WriteString("TRACe" + t.ToString() + ":DATA:NAME?");
                    string tName = instruments[SpecAn1].ReadString().Replace("\n", string.Empty);

                    if (tName.Contains("Syms/Errs"))
                    {
                        selectedTraceNo = t;
                        rtbWhiteboard.AppendText("Suitable MER 'Trace' found: TRACE" + selectedTraceNo.ToString() + Environment.NewLine);
                        break;
                    }
                }
                catch
                {
                    continue;
                }

            }
            return selectedTraceNo;
        }

        private void bMERDcs_Click(object sender, EventArgs e)
        {
            // Gather the basic parameters of the test...

            char temperature = PreliminaryActions("MER-DCS", out string test, out List<int> ports, out List<string> lnbs, out List<string> bands, out string serialNumber);
            string opticalPower = "-15dBm";
            File.AppendAllText(resultsFile, "Optical input power  = " + opticalPower + Environment.NewLine);

            File.AppendAllText(resultsFile,
                "LNB".PadRight(4)
                + "Pol.".PadRight(6)
                + "Rx".PadRight(4)
                + "Input_MHz".PadRight(17)
                + "Output_MHz".PadRight(17)
                + "MER_dB" + Environment.NewLine);

            // Set each Rx to CHANNEL-STACKED and disable TDMA polling...
            foreach (int r in ports)
            {
                bool success = FSK.SetStackingMode(usb, r, "ChannelStacked");
                if (!success) MessageBox.Show("Failed to set C/S mode for Rx " + r.ToString());

                success = FSK.DisableTdmaPolling(usb, 'F', r, true);
                if (!success) MessageBox.Show("Failed to disable TDMA polling for Rx " + r.ToString());

                success = FSK.DisableTdmaPolling(usb, 'F', r, true);
                if (!success) MessageBox.Show("Failed to disable TDMA polling for Rx " + r.ToString());
                else MessageBox.Show("Successfully disabled TDMA polling for Rx " + r.ToString());
            }

            // Prepare the spectrum analyser...

            Binstrument_Click(bSA, new EventArgs());
            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");

            int selectedTraceNo = Setup89601VSA();

            if (selectedTraceNo == 0)
            {
                rtbWhiteboard.AppendText("No suitable Trace found" + Environment.NewLine);
                return;
            }
            string traceName = "TRACE" + selectedTraceNo;

            // -------------------------  Setup the transponder frequencies  ------------------------------

            List<decimal> lowerTransponders = new List<decimal>();
            List<decimal> upperTransponders = new List<decimal>();
            for (decimal tp = numFirstLower.Value; tp < 1450; tp += numLowerStep.Value)
            {
                lowerTransponders.Add(tp);
            }

            for (decimal tp = numFirstUpper.Value; tp < 2150; tp += numUpperStep.Value)
            {
                upperTransponders.Add(tp);
            }

            // read grid...

            List<int> grid = FSK.RequestUserBlockGrid(usb, 'F', ports[0]).ToList<int>();
            List<decimal> centreFrequencies = new List<decimal>();
            int noUserBands = 16;
            for (int i = 0; i < noUserBands; i++)
            {
                centreFrequencies.Add(grid[0] + i * grid[1]);
            }

            // Create a datatable for the results and bind it to the dgv

            System.Data.DataTable mers = new System.Data.DataTable();
            mers.Columns.Add("Frequency", typeof(decimal));
            mers.Columns.Add("MER(dB)", typeof(double));
            dgvMERs.DataSource = mers;
            dgvMERs.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvMERs.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            // Set the progress bar max...
            progressBar1.Maximum = lnbs.Count * ports.Count * (lowerTransponders.Count + upperTransponders.Count) * noUserBands;

            foreach (int r in ports)
            {
                // Route the proper port to the spectrum analyser...

                try
                {
                    var rb = gbPickering.Controls.OfType<RadioButton>().FirstOrDefault(k => k.Tag.ToString() == r.ToString());
                    (rb as RadioButton).Checked = true;
                    Pickering_CheckedChanged((rb as RadioButton), new EventArgs());
                }
                catch
                {
                    MessageBox.Show("Connect Rx " + r.ToString());
                }

                // Allocate all channels to current Rx...

                for (int ch = 0; ch < 18; ch++)
                {
                    _ = FSK.DeallocateChannel(usb, 'F', r, ch);
                    bool success = FSK.AllocateChannel(usb, 'F', r, ch);
                    if (!success) MessageBox.Show("Failed to allocate channel " + ch.ToString());
                }

                foreach (string l in lnbs)
                {
                    if (chIntervene.Checked)
                        MessageBox.Show("Attach LNB " + l);

                    foreach (bool isUpper in new[] { false, true })
                    {
                        foreach (decimal inputFrequency in (isUpper ? upperTransponders : lowerTransponders))
                        {
                            // Send the current tp to every UB...

                            int KHz = Convert.ToInt32(1000 * inputFrequency);
                            if (l == "C" && isUpper)
                                KHz += 28500;       // Neither understand nor like this. Must add 28.5 MHz!

                            for (int ch = 0; ch < 18; ch++)
                            {
                                bool success = FSK.Command38(usb, 'F', r, l, ch, KHz, isUpper, null);
                                if (!success) MessageBox.Show("Failed to allocate frequency to channel " + ch.ToString());
                            }

                            rtbWhiteboard.AppendText("Measuring at port " + r.ToString() + ", from LNB " + l + (isUpper ? " Upper, " : " Lower, ") + Environment.NewLine);
                            rtbWhiteboard.Refresh();
                            rtbWhiteboard.Update();
                            rtbWhiteboard.SelectionStart = rtbWhiteboard.Text.Length;
                            rtbWhiteboard.ScrollToCaret();

                            foreach (decimal ub in centreFrequencies)
                            {
                                //  ---------------------  Take the measurements  ---------------------------------------

                                string dataRow = l.PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4);

                                string message = string.Format(":FREQuency:CENTer {0} MHz", ub);
                                instruments[SpecAn1].WriteString(message);
                                instruments[SpecAn1].WriteString("*OPC?");    // Prevents further action until operation has completed
                                _ = instruments[SpecAn1].ReadString();

                                message = ":INPUT ANALOG:RANGE:AUTO";
                                instruments[SpecAn1].WriteString(message);
                                instruments[SpecAn1].WriteString("*OPC?");    // Prevents further action until operation has completed
                                _ = instruments[SpecAn1].ReadString();

                                double[] temp = new double[(int)numAverages.Value];
                                try
                                {
                                    for (int av = 0; av < (int)numAverages.Value; av++)
                                    {
                                        Thread.Sleep(100);
                                        instruments[SpecAn1].WriteString(traceName + ":DATA:TABL?  'SigToNoise'");
                                        instruments[SpecAn1].WriteString(":DATA:TABL?  \"SigToNoise\"");
                                        temp[av] = double.Parse(instruments[SpecAn1].ReadString());
                                    }
                                }
                                catch (COMException ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                    return;
                                }
                                mers.Rows.Add(ub.ToString(), Math.Round(temp.Average(), 2));
                                dgvMERs.Refresh();
                                dgvMERs.Update();
                                dgvMERs.FirstDisplayedScrollingRowIndex = dgvMERs.RowCount - 1;

                                // Save results to a file...

                                dataRow += ub.ToString().PadRight(17) + inputFrequency.ToString().PadRight(17) + Math.Round(temp.Average(), 2);
                                File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                                progressBar1.Value += 1;
                            }
                        }
                    }
                }
            }
            File.AppendAllText(resultsFile, Environment.NewLine + "Finished at " + DateTime.Now);
            RenameResultsFile("MER-DCS", ports, serialNumber, test, temperature);
            System.Media.SystemSounds.Beep.Play();
        }

        private void BMerDbs_Click(object sender, EventArgs e)
        {
            // Gather the basic parameters of the test...

            char temperature = PreliminaryActions("MER-DBS", out string test, out List<int> ports, out List<string> lnbs, out List<string> bands, out string serialNumber);
            string opticalPower = "-15dBm";
            File.AppendAllText(resultsFile, "Optical input power  = " + opticalPower + Environment.NewLine);

            File.AppendAllText(resultsFile,
                "LNB".PadRight(4)
                + "Pol.".PadRight(6)
                + "Rx".PadRight(4)
                + "Input_MHz".PadRight(17)
                + "Output_MHz".PadRight(17)
                + "MER_dB" + Environment.NewLine);

            // Sort out the ProgressBar...
            progressBar1.Maximum = ports.Count * bands.Count;

            // Prepare the spectrum analyser...

            Binstrument_Click(bSA, new EventArgs());
            int SpecAn1 = instrumentLog.FindIndex(element => element == "SA1");

            int selectedTraceNo = Setup89601VSA();

            if (selectedTraceNo == 0)
            {
                rtbWhiteboard.AppendText("No suitable Trace found" + Environment.NewLine);
                return;
            }
            string traceName = "TRACE" + selectedTraceNo;

            // -------------------------  Setup the transponder frequencies  ------------------------------

            List<decimal> transponders = new List<decimal>();
            for (decimal tp = numFirstLower.Value; tp < 1450; tp += numLowerStep.Value)
            {
                transponders.Add(tp);
            }

            for (decimal tp = numFirstUpper.Value; tp < 2150; tp += numUpperStep.Value)
            {
                transponders.Add(tp);
            }

            progressBar1.Maximum *= transponders.Count;

            // Create a datatable for the results and bind it to the dgv

            System.Data.DataTable mers = new System.Data.DataTable();
            mers.Columns.Add("Frequency", typeof(decimal));
            mers.Columns.Add("MER(dB)", typeof(double));
            dgvMERs.DataSource = mers;
            dgvMERs.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvMERs.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            foreach (int r in new int[] { 1, 2, 3 })
            {
                // Route the proper port to the spectrum analyser...

                try
                {
                    var rb = gbPickering.Controls.OfType<RadioButton>().FirstOrDefault(k => k.Tag.ToString() == r.ToString());
                    (rb as RadioButton).Checked = true;
                    Pickering_CheckedChanged((rb as RadioButton), new EventArgs());
                }
                catch
                {
                    MessageBox.Show("Connect Rx " + r.ToString());
                }
                MessageBox.Show("Attach LNB " + r.ToString());

                foreach (bool isUpper in new[] { false, true })
                {
                    string b = isUpper ? "Mid" : "Low";
                    rtbWhiteboard.AppendText("Measuring at port " + r.ToString() + (isUpper ? " Upper, " : " Lower, ") + "to " + b + Environment.NewLine);
                    rtbWhiteboard.Refresh();
                    rtbWhiteboard.Update();

                    //  ---------------------  Take the measurements  ---------------------------------------

                    foreach (decimal tp in transponders)
                    {
                        string dataRow = opticalPower.PadRight(10) + r.ToString().PadRight(4) + (isUpper ? "Upper" : "Lower").PadRight(6) + r.ToString().PadRight(4) + b.PadRight(7);

                        string message = string.Format(":FREQuency:CENTer {0} MHz", tp);
                        instruments[SpecAn1].WriteString(message);
                        instruments[SpecAn1].WriteString("*OPC?");    // Prevents further action until operation has completed
                        string discard = instruments[SpecAn1].ReadString();

                        double[] temp = new double[(int)numAverages.Value];
                        try
                        {
                            for (int av = 0; av < (int)numAverages.Value; av++)
                            {
                                Thread.Sleep(100);
                                instruments[SpecAn1].WriteString(traceName + ":DATA:TABL?  'SigToNoise'");
                                instruments[SpecAn1].WriteString(":DATA:TABL?  \"SigToNoise\"");
                                temp[av] = double.Parse(instruments[SpecAn1].ReadString());
                            }
                        }
                        catch (COMException ex)
                        {
                            MessageBox.Show(ex.ToString());
                            return;
                        }
                        mers.Rows.Add(tp.ToString(), Math.Round(temp.Average(), 2));
                        dgvMERs.Refresh();
                        dgvMERs.Update();

                        // Save results to a file...

                        dataRow += tp.ToString().PadRight(17) + Math.Round(temp.Average(), 2);
                        File.AppendAllText(resultsFile, dataRow + Environment.NewLine);
                        progressBar1.Value += 1;
                    }
                }
                File.AppendAllText(resultsFile, "-----------------------------------------------------" + Environment.NewLine);

            }
            RenameResultsFile("MER-DBS", ports, serialNumber, test, temperature);
            System.Media.SystemSounds.Beep.Play();
            return;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            automaticSwitching = checkBox1.Checked ? false : true;
            foreach (Control ch in gbPickering.Controls)
            {
                ch.Enabled = automaticSwitching;
            }
            (sender as CheckBox).Enabled = true;
        }
        private decimal CalculateDestackerFrequency(decimal input, bool isUpper, string band)
        {
            decimal Lower_Low = 0;
            decimal Lower_Mid = 705;
            decimal Lower_High = 1550;
            decimal Upper_Low = -725;
            decimal Upper_Mid = 0;
            decimal Upper_High = 825;

            decimal shift = 0;
            switch (band)
            {
                case "Low":
                    {
                        shift = isUpper ? Upper_Low : Lower_Low;
                        break;
                    }
                case "Mid":
                    {
                        shift = isUpper ? Upper_Mid : Lower_Mid;
                        break;
                    }
                case "High":
                    {
                        shift = isUpper ? Upper_High : Lower_High;
                        break;
                    }
            }
            return input + shift;
        }
        private void CheckStatus(int instrument, out bool measurementDone, out bool autoRanging, out bool calibrating, out bool notPaused)
        {
            measurementDone = false;
            autoRanging = false;
            calibrating = false;
            notPaused = false;

            for (int i = 0; i < 100; i++)
            {
                instruments[instrument].WriteString("MEAS:STAT?");
                int status = int.Parse(instruments[instrument].ReadString());
                if (status < 0) // it is autoranging
                {
                    autoRanging = true;
                    continue;
                }
                if ((status & 0x1) == 0x1)
                {
                    measurementDone = true;
                    break;
                }
                if ((status & 0x1) == 0x2)
                {
                    calibrating = true;
                    MessageBox.Show("Calibrating. Wait for it to finish.");
                }
                if ((status & 0x1073741824) == 0x1073741824)
                    notPaused = true;
                continue;
            }

            //Ranging(-2147483648) A measurement autorange is in progress.
            //MeasurementDone(1) Measurement is done.
            //Calibrating(2) Calibration is in progress.
            //Acquiring(4) The measurement is acquiring data
            //Settling(8) The measurement is acquiring the settle data(before the measurement data).
            //WaitPreTrigger(16) The measurement is acquiring pre - trigger data.
            // WaitTrigger(32) The measurement is waiting for the trigger event.
            //ReadingData (64) The measurement is reading the acquisition data from the hardware.
            //Recording(128) The measurement is recording data.
            //AverageComplete(256) The measurement average is complete.
            //SyncNotFound(512) No sync is found by the(demodulation) measurement.
            //PulseNotFound(1024) No pulse is found by the(demodulation) measurement.
            //CalibrationNeeded(4096) Calibration is needed.
            //CalibrationWarmUp(8192) Calibration is warming up.
            //ExternalReferenceLock(16384) The hardware is locked to the external reference.
            //InternalReferenceLock(32768) The hardware is locked to the internal reference.
            //GapData(65536) There is a gap between the last scan and the current scan of data.
            //EndData(131072) Recording playback has reached the end of the data.
            //TestFail(262144) A test(e.g.LimitTest) has failed.
            //WaitReferenceLock(524288) The measurement is waiting for a frequency reference lock.
            //Visible(1048576) The application window is visible
            //AdcOverload(2097152) The measurement hardware input channel has over-ranged.
            //StaleData(536870912) Restarting measurement ascquisition because the data is stale(too old).
            //Measuring(1073741824) - The measurement is not paused(i.e.it is running).
        }
        private void PerformSweep(int instrument, int noAverages, bool continuous)
        {
            // Set single-sweep mode
            instruments[instrument].WriteString("INIT:CONT ON");

            // Enable and set AVERAGES...
            instruments[instrument].WriteString(":SENSE:AVERAGE:STATE 1");
            instruments[instrument].WriteString(":SENSE:AVERAGE:COUNT " + (int)numAverages.Value);
            instruments[instrument].WriteString(":SENSe]:AVERage: REPeat 0");   // just the once

            // Autorange the spectrum analyser...this performs the sweep
            instruments[instrument].WriteString(":ANAL:RANGE:AUTO");
            CheckStatus(instrument, out bool measurementDone, out bool autoRanging, out bool calibrating, out bool notPaused);
        }
    }
}

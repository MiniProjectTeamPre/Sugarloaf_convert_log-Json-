using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Text;
using System.Drawing;
using System.Threading;
using Newtonsoft.Json;
using Spire.Xls;

namespace Sugarloaf_convert_log {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }
        private string pathInput = "Input";
        private string pathOutput = "Output";
        private string pathBackUp = "D:\\SugarloafConvertLogBackUpFile";

        private void Form1_Load(object sender, EventArgs e) {
            if (!Directory.Exists(pathInput)) Directory.CreateDirectory(pathInput);
            if (!Directory.Exists(pathOutput)) Directory.CreateDirectory(pathOutput);
            if (!Directory.Exists(pathBackUp)) Directory.CreateDirectory(pathBackUp);
            Application.Idle += Application_Idle;
        }

        Workbook workbook = new Workbook();
        Worksheet worksheet;
        private string DataSummary = "";
        private string FindResultFormHead(JsonConvert_ data, string head, string head2 = "lmdcasklcecednclwsdcsdf") {
            string result = "-";
            int index = 0;
            head = head.Replace(" ", string.Empty);
            head2 = head2.Replace(" ", string.Empty);

            foreach (JsonConvert_.ResultString_ dataSup in data.ResultString) {
                if (dataSup.Description.Replace(" ", string.Empty).Contains(head) ||
                    dataSup.Description.Replace(" ", string.Empty).Contains(head2)) {
                    result = dataSup.Measured;
                    break;
                }

                index++;
            }

            if (index < data.ResultString.Count) {
                data.ResultString.RemoveAt(index);
            }
            return result;
        }
        private string FindResultFormStep(JsonConvert_ data, string step, string step2 = "1212.23123.142") {
            string result = "-";
            int index = 0;
            step = step.Replace(" ", string.Empty);
            step2 = step2.Replace(" ", string.Empty);

            foreach (JsonConvert_.ResultString_ dataSup in data.ResultString) {
                if (dataSup.Step.Replace(" ", string.Empty) == step || 
                    dataSup.Step.Replace(" ", string.Empty) == step2 ) {
                    result = dataSup.Measured;
                    break;
                }

                index++;
            }

            if (index < data.ResultString.Count) {
                data.ResultString.RemoveAt(index);
            }
            return result;
        }
        private int FineResultFromHeadConfig(List<string> header, string nameHead) {
            int result = 0;

            for (int loop = 0; loop < header.Count; loop++) {
                if (header[loop] != nameHead) {
                    continue;
                }

                result = loop;
                break;
            }

            return result;
        }
        private void Application_Idle(object sender, EventArgs e) {
            if (!flag_running) return;
            List<string> file_data = new List<string>();
            try {
                string[] zxc = Directory.GetFiles(pathInput);
                file_data = zxc.ToList<string>();
            } catch { }
            string s = "";
            for (int i = 0; i < 5; i++) {
                try { s = file_data[i].Replace(pathInput + "\\", ""); } catch { break; }
                string file_name = s.Replace(".xlsx", ".csv");
                if (!File.Exists(pathOutput + "\\" + file_name)) {
                    StreamWriter swOut_ = new StreamWriter(pathOutput + "\\" + file_name, true);
                    string ghf = "";
                    foreach (string zzxc in head_all) {
                        ghf += zzxc + ",";
                    }
                    swOut_.WriteLine(ghf);
                    swOut_.Close();
                }
                StreamWriter swOut = new StreamWriter(pathOutput + "\\" + file_name, true);

                try {
                    workbook.LoadFromFile(file_data[i]);
                } catch {
                    MessageBox.Show("_ปิด excel ก่อน");
                    return;
                }
                worksheet = workbook.Worksheets["Sheet1"];
                List<string> result = new List<string>();
                List<string> serialCustomer = new List<string>();
                bool flagColumn = false;
                for (int row = 2; row < 99999; row++) {
                    try {
                        if (!worksheet.GetText(row, 1).Contains(",")) {
                            result.Add(worksheet.GetText(row, 2));
                            serialCustomer.Add(worksheet.GetText(row, 1));
                            flagColumn = true;

                        } else {
                            result.Add(worksheet.GetText(row, 1));
                            flagColumn = false;
                        }
                    } catch {
                        if (flagColumn) {
                            result.Add(worksheet.GetText(row, 2));
                        } else {
                            result.Add(worksheet.GetText(row, 1));
                        }
                    }

                    if (result[result.Count - 1] == null) {
                        result.RemoveAt(result.Count - 1);
                        break;
                    }
                }

                DataSummary = string.Empty;
                for (int loop = 0; loop < result.Count(); loop++) {

                    try {
                        if (result[loop].Contains("{\"Date\"")) {
                            ProcessJson(result[loop], serialCustomer[loop]);

                        } else {
                            ProcessNormal(result[loop], serialCustomer[loop]);
                        }
                    } catch {
                        MessageBox.Show("Error Row = " + (loop + 2) + " !!!!");
                        return;
                    }
                    
                    while (true) {
                        try {
                            swOut.WriteLine(DataSummary);
                            break;
                        } catch { MessageBox.Show("_กรุณาปิด log file csv ก่อน"); }
                    }

                }
                swOut.Close();
                bool removeFile = false;
                try {
                    File.Move(file_data[i], file_data[i].Replace(pathInput, pathBackUp));
                    removeFile = true;
                } catch { }
                if (!removeFile) {
                    File.Delete(file_data[i].Replace(pathInput, pathBackUp));
                    File.Move(file_data[i], file_data[i].Replace(pathInput, pathBackUp));
                }
            }
            Thread.Sleep(2000);
        }
        private void ProcessJson(string data, string serialCustomer) {
            string[] resultSup = data.Split(',');
            bool useSerialCus = false;
            if (!resultSup[0].Contains("{\"Date"))
            {
                DataSummary = data.Replace(resultSup[0] + ",", string.Empty);
            }
            else
            {
                DataSummary = data;
                useSerialCus = true;
            }

            DataSummary = DataSummary.Replace("},]", "}]");
            JsonConvert_ dataGet = JsonConvert.DeserializeObject<JsonConvert_>(DataSummary);
            if (useSerialCus)
            {
                DataSummary = serialCustomer + ",";
            }
            else
            {
                DataSummary = resultSup[0] + ",";
            }
            DataSummary += FindResultFormHead(dataGet, "program firmware") + ",";
            DataSummary += FindResultFormHead(dataGet, "measure voltage bat") + ",";
            DataSummary += FindResultFormHead(dataGet, "measure voltage vcc") + ",";
            DataSummary += FindResultFormHead(dataGet, "[running_current_test]") + ",";
            DataSummary += "'" + FindResultFormHead(dataGet, "read uc15 fm version") + ",";
            DataSummary += "'" + FindResultFormHead(dataGet, "read uc15 IMEI") + ",";
            DataSummary += "'" + FindResultFormHead(dataGet, "read sim card ICCID") + ",";
            DataSummary += FindResultFormHead(dataGet, "crystal y1 [ppm]") + ",";
            DataSummary += FindResultFormHead(dataGet, "crystal y1 [khz]") + ",";
            DataSummary += FindResultFormHead(dataGet, "measure rtc oscillator") + ",";
            DataSummary += FindResultFormHead(dataGet, "[standby_current_test]") + ",";
            DataSummary += FindResultFormHead(dataGet, "[sleep_current_test]") + ",";
            DataSummary += dataGet.Failure + "-,";
            DataSummary += dataGet.FinalResult + ",";
            string[] dateSup = dataGet.Date.Split('/');
            string dateTimeSup = dateSup[2] + "." + dateSup[1] + "." + dateSup[0] + " ";
            DataSummary += dateTimeSup;
            DataSummary += dataGet.Time + ",";
            DataSummary += comboBox2.Text + ",";
            DataSummary += dataGet.SN + ",";
            DataSummary += dataGet.LoginID + ",";
            int secondsEnd = Convert.ToInt32(TimeSpan.Parse(dataGet.Time).TotalSeconds);
            int secondsIng = Convert.ToInt32(dataGet.TestTime);
            int secondsOld = secondsEnd - secondsIng;
            DataSummary += dateTimeSup + convert2time(secondsOld) + ",";
            DataSummary += dateTimeSup + dataGet.Time + ",";
            DataSummary += "'" + convert2time(secondsIng) + ",";
            DataSummary += FindResultFormHead(dataGet, "antenna’s rssi") + ",";
            DataSummary += FindResultFormHead(dataGet, "start switch test on") + ",";
            DataSummary += FindResultFormHead(dataGet, "sttus switch test on") + ",";
            DataSummary += FindResultFormHead(dataGet, "uart ack test") + ",";
            DataSummary += FindResultFormStep(dataGet, "4.17") + ",";
            DataSummary += FindResultFormStep(dataGet, "4.16") + ",";
            DataSummary += FindResultFormHead(dataGet, "temperature in IC") + ",";
            DataSummary += FindResultFormHead(dataGet, "eeprom test") + ",";
            DataSummary += FindResultFormHead(dataGet, "check led red on") + ",";
            DataSummary += FindResultFormHead(dataGet, "check led green on") + ",";
            DataSummary += FindResultFormHead(dataGet, "check jumper config");
        }
        private void ProcessNormal(string data, string serialCustomer) {
            Workbook workbookSup = new Workbook();
            Worksheet worksheetSup;
            if (!data.Contains(serialCustomer))
            {
                data = serialCustomer + "," + data;
            }
            string[] dataSup = data.Split(',');

            if(dataSup.Length < 2) {
                return;
            }

            try {
                workbookSup.LoadFromFile("HeaderLog.xlsx");
            } catch {
                MessageBox.Show("_ปิด excel ก่อน");
                return;
            }

            worksheetSup = workbookSup.Worksheets["Sheet1"];
            List<string> header = new List<string>();
            for (int column = 1; column < 99999; column++) {
                header.Add(worksheetSup.GetText(1, column));

                if (header[header.Count - 1] == null) {
                    header.RemoveAt(header.Count - 1);
                    break;
                }
            }

            DataSummary = dataSup[0] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Firmware_CRC32")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Battery_Volt_Test")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Vcc_Volt_Test")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Running_Curr_Test")] + ",";
            DataSummary += "'" + dataSup[FineResultFromHeadConfig(header, "Modem_FW_Test")] + ",";
            DataSummary += "'" + dataSup[FineResultFromHeadConfig(header, "Modem_IMEI_Test")] + ",";
            DataSummary += "'" + dataSup[FineResultFromHeadConfig(header, "SIM_ICCID_Test")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Crystal_Test_ppm")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Crystal_Test_kHz")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Crystal_Test_Hz")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Standby_Curr_Test")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Sleep_Curr_Test")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Fail")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Final_Result")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "DATE_TIME")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "TESTER_ID")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "TEAM_SN")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Operator")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Test_Start_Time")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Test_Finish_Time")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Test_Total_Time")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Antenna_RSSI")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Check_SW1_Test")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Check_SW2_Test")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Processor_Functional")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Measure_Light_Sensor_Dark")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Measure_Light_Sensor_Light")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Measure_Temp_Sensor")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Check_Memory_EEPROM")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Led1_On")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Led2_On")] + ",";
            DataSummary += dataSup[FineResultFromHeadConfig(header, "Check_Jumper")].ToString();
        }
        private TestResult string2json(string input) {
            TestResult result = new TestResult();
            string[] split_ResultString = input.Replace("\"ResultString\":[", "฿").Split('฿');
            List<string> values = new List<string>();
            List<string> keys = new List<string>();
            string pattern = @"\""(?<key>[^\""]+)\""\:\""?(?<value>[^\"",}]+)\""?\,?";
            foreach (Match m in Regex.Matches(split_ResultString[0] + "}", pattern)) {
                if (m.Success) {
                    values.Add(m.Groups["value"].Value);
                    //keys.Add(m.Groups["key"].Value);
                }
            }
            result.Date = values[0];
            result.Time = values[1];
            result.LoginID = values[2];
            result.VersionSW = values[3];
            result.VersionFW = values[4];
            result.VersionSpec = values[5];
            result.TestTime = values[6];
            result.LoadIn = values[7];
            result.Mode = values[8];
            result.Result = values[9];
            result.SN = values[10];
            try { result.Failure = values[11]; } catch { result.Failure = ""; }

            List<ResultStepDetail> resultString = new List<ResultStepDetail>();
            split_ResultString[1] = split_ResultString[1].Replace(",]}", "");
            string[] step = split_ResultString[1].Replace("},{", "฿").Split('฿');
            step[0] = step[0] + "}";
            step[step.Count() - 1] = "{" + step[step.Count() - 1];
            for (int h = 1; h < step.Count() - 1; h++) {
                step[h] = "{" + step[h] + "}";
            }
            for (int h = 0; h < step.Count(); h++) {
                values.Clear();
                foreach (Match m in Regex.Matches(step[h], pattern)) {
                    if (m.Success) {
                        values.Add(m.Groups["value"].Value);
                    }
                }
                while (values.Count < 5) { values.Add(""); }
                resultString.Add(new ResultStepDetail() { Step = values[0], Description = values[1], Tolerance = values[2], Measured = values[3], Result = values[4] });
            }
            result.ResultString = resultString;

            return result;
        }

        private static void DelaymS(int mS) {
            Stopwatch stopwatchDelaymS = new Stopwatch();
            stopwatchDelaymS.Restart();
            while (mS > stopwatchDelaymS.ElapsedMilliseconds) {
                if (!stopwatchDelaymS.IsRunning) stopwatchDelaymS.Start();
                Application.DoEvents();
                Thread.Sleep(50);
            }
            stopwatchDelaymS.Stop();
        }
        private string convert2time(int testTime) {
            int testTime_hh = 0;
            int testTime_mm = testTime / 60;
            int testTime_ss = testTime % 60;
            if (testTime_mm > 59) {
                testTime_hh = testTime_mm / 60;
                testTime_mm = testTime_mm % 60;
            }
            return testTime_hh.ToString("00") + ":" + testTime_mm.ToString("00") + ":" + testTime_ss.ToString("00");
        }

        private bool flag_running = false;
        private void timer1_Tick(object sender, EventArgs e) {
            timer1.Enabled = false;
            if (!flag_running) { timer1.Enabled = true; return; }
            this.BackgroundImage = Properties.Resources.file_01;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_02;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_03;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_04;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_00;
            DelaymS(1000);
            timer1.Enabled = true;
        }
        List<string> head_all = new List<string>();
        private void button1_Click(object sender, EventArgs e) {
            //if (comboBox1.Text == null || comboBox1.Text == "") { MessageBox.Show("_กรุณาเลือก FG ก่อน"); return; }
            if (button1.Text == "RUN") {
                button1.Text = "STOP";
                button1.BackColor = Color.Red;
                flag_running = true;
            } else {
                button1.Text = "RUN";
                button1.BackColor = Color.Aqua;
                flag_running = false;
            }

            head_all.Clear();
            head_all.Add("Column1");
            head_all.Add("Firmware_CRC32");
            head_all.Add("Battery_Volt_Test");
            head_all.Add("Vcc_Volt_Test");
            head_all.Add("Running_Curr_Test");
            head_all.Add("Modem_FW_Test");
            head_all.Add("Modem_IMEI_Test");
            head_all.Add("SIM_ICCID_Test");
            head_all.Add("Crystal_Test_ppm");
            head_all.Add("Crystal_Test_kHz");
            head_all.Add("Crystal_Test_Hz");
            head_all.Add("Standby_Curr_Test");
            head_all.Add("Sleep_Curr_Test");
            head_all.Add("Fail");
            head_all.Add("Final_Result");
            head_all.Add("DATE_TIME");
            head_all.Add("TESTER_ID");
            head_all.Add("TEAM_SN");
            head_all.Add("Operator");
            head_all.Add("Test_Start_Time");
            head_all.Add("Test_Finish_Time");
            head_all.Add("Test_Total_Time");
            head_all.Add("Antenna_RSSI");
            head_all.Add("Check_SW1_Test");
            head_all.Add("Check_SW2_Test");
            head_all.Add("Processor_Functional");
            head_all.Add("Measure_Light_Sensor_Dark");
            head_all.Add("Measure_Light_Sensor_Light");
            head_all.Add("Measure_Temp_Sensor");
            head_all.Add("Check_Memory_EEPROM");
            head_all.Add("Led1_On");
            head_all.Add("Led2_On");
            head_all.Add("Check_Jumper");
        }
    }

    public class TestResult {
        public string Date { get; set; }
        public string Time { get; set; }
        public string LoginID { get; set; }
        public string VersionSW { get; set; }
        public string VersionFW { get; set; }
        public string VersionSpec { get; set; }
        public string TestTime { get; set; }
        public string LoadIn { get; set; }
        public string Mode { get; set; }
        public string Result { get; set; }
        public string SN { get; set; }
        public string Failure { get; set; }
        public List<ResultStepDetail> ResultString { get; set; }
    }
    public class ResultStepDetail {
        public string Step { get; set; }
        public string Description { get; set; }
        public string Tolerance { get; set; }
        public string Measured { get; set; }
        public string Result { get; set; }
    }

    public class JsonConvert_ {
        public string Date { get; set; }
        public string Time { get; set; }
        public string LoginID { get; set; }
        public string SWVersion { get; set; }
        public string FWVersion { get; set; }
        public string SpecVersion { get; set; }
        public string TestTime { get; set; }
        public string LoadInOut { get; set; }
        public string Mode { get; set; }
        public string FinalResult { get; set; }
        public string SN { get; set; }
        public object Failure { get; set; }
        public List<ResultString_> ResultString { get; set; }

        public JsonConvert_() {
            Date = string.Empty;
            Time = string.Empty;
            LoginID = string.Empty;
            SWVersion = string.Empty;
            FWVersion = string.Empty;
            SpecVersion = string.Empty;
            TestTime = string.Empty;
            LoadInOut = string.Empty;
            Mode = string.Empty;
            FinalResult = string.Empty;
            SN = string.Empty;
            Failure = string.Empty;
            ResultString = new List<ResultString_>();
        }
        public class ResultString_ {
            public string Step { get; set; }
            public string Description { get; set; }
            public string Tolerance { get; set; }
            public string Measured { get; set; }
            public string Result { get; set; }

            public ResultString_() {
                Step = string.Empty;
                Description = string.Empty;
                Tolerance = string.Empty;
                Measured = string.Empty;
                Result = string.Empty;
            }
        }
    }
}

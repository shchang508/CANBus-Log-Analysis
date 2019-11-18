using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace CANLog
{
    public partial class Form1 : MaterialForm
    {
        string logPathParameters;
        public Form1(string args)
        {
            InitializeComponent();
            logPathParameters = args;

            // Create a material theme manager and add the form to manage (this)
            MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;

            // Configure color schema
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE
            );

            // Button design
            this.button_Decode.FlatAppearance.BorderColor = Color.FromArgb(16, 145, 232);
            this.button_Decode.FlatAppearance.BorderSize = 3;
            this.button_Decode.BackColor = System.Drawing.Color.FromArgb(16, 145, 232);

            this.button_ExcelGenerate.FlatAppearance.BorderColor = Color.FromArgb(16, 145, 232);
            this.button_ExcelGenerate.FlatAppearance.BorderSize = 3;
            this.button_ExcelGenerate.BackColor = System.Drawing.Color.FromArgb(16, 145, 232);

            this.button_TxtToCSV.FlatAppearance.BorderColor = Color.FromArgb(16, 145, 232);
            this.button_TxtToCSV.FlatAppearance.BorderSize = 3;
            this.button_TxtToCSV.BackColor = System.Drawing.Color.FromArgb(16, 145, 232);

            this.button_GraphGenerate.FlatAppearance.BorderColor = Color.FromArgb(16, 145, 232);
            this.button_GraphGenerate.FlatAppearance.BorderSize = 3;
            this.button_GraphGenerate.BackColor = System.Drawing.Color.FromArgb(16, 145, 232);
        }

        // Select log file from computer 
        private void button_SelectFileClick(object sender, EventArgs e)
        {
            OpenFileDialog path = new OpenFileDialog();
            path.FileName = "Select a text file";
            path.InitialDirectory = logPathParameters;
            path.Filter = "Text files (*.txt)|*.txt";
            path.Title = "Open text file";
            path.ShowDialog();
            this.textBox_Path.Text = path.FileName;
        }

        // File path and name
        string docPath = string.Empty;
        string fileName = string.Empty;
        string txtFileName = string.Empty;

        //Raw data for CANBus log
        string line = string.Empty;
        string command = string.Empty;
        string line_CAN_ID = string.Empty;
        string line_CAN_DATA = string.Empty;

        //Raw data for Power Supply log
        string linePS = string.Empty;
        string commandPS = string.Empty;
        string line_PS_Command = string.Empty;

        // Command and response 
        List<string> commandList = new List<string>();
        List<string> responseList = new List<string>();
        string[] commandSplit = new string[] { "", "" };
        string result = string.Empty;
        string[] responseSplit = new string[] { };
        string responseResult = string.Empty;
        string frontTyreResponse = string.Empty;
        string rearTyreResponse = string.Empty;

        // CMD_A
        string Power = string.Empty;
        string OilPressureLight = string.Empty;
        string FuelLight = string.Empty;
        string ABSLight = string.Empty;
        string WaterTempLight = string.Empty;
        string MaintenanceLight = string.Empty;
        string FrontTirePressure = string.Empty;
        string RearTirePressure = string.Empty;
        string CMD_A_decode_result = string.Empty;

        // CMD_B
        string Speed = string.Empty;
        string RPM = string.Empty;
        string FuelConsumption = string.Empty;
        string CMD_B_decode_result = string.Empty;

        // CMD_C
        string WaterTemp = string.Empty;
        string TempSign = string.Empty;
        string RoomTemp = string.Empty;
        string Fuel = string.Empty;
        string Battery = string.Empty;
        string CMD_C_decode_result = string.Empty;

        // CMD_D
        string Mileage = string.Empty;
        string MaxSpeed = string.Empty;
        string AverageSpeed = string.Empty;
        string CMD_D_decode_result = string.Empty;

        // CMD_E
        List<string> ABS_dtc_List = new List<string>();
        string CMD_E_decode_result = string.Empty;
        List<string> absCommandList = new List<string>();
        string absTmp = string.Empty;
        string absResponse = string.Empty;
        string[] absResponseSplit = new string[] { };
        string[] absToArray = new string[] { };

        // CMD_F
        List<string> OBD_dtc_List = new List<string>();
        string CMD_F_decode_result = string.Empty;
        List<string> obdCommandList = new List<string>();
        string obdTmp = string.Empty;
        string obdResponse = string.Empty;
        string[] obdResponseSplit = new string[] { };
        string[] obdToArray = new string[] { };

        // Convert hex to text in CAN log
        private void button_Decode_Click(object sender, EventArgs e)
        {
            if (textBox_Path.Text != "")
            {
                List<string> readContent = new List<string>(); //For saving every line 
                StreamReader sr = new StreamReader(textBox_Path.Text);
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine();
                    if (!string.IsNullOrEmpty(line))
                    {
                        readContent.Add(line);
                    }
                }
                sr.Close();

                //Location of test result files 
                //docPath = @"D:\TestResult\"; 
                docPath = textBox_Path.Text;
                /*if (Directory.Exists(docPath))
                {
                    //Diretory exists
                }
                else
                {
                    //Create a new directory
                    Directory.CreateDirectory(docPath);
                }*/


                //fileName = string.Format("TestResult_{0}", DateTime.Now.ToString("yyyyMMddHHmmss")); //File name of test result file 
                fileName = string.Format("TestResult_{0}", Path.GetFileName(docPath).Substring(Path.GetFileName(docPath).IndexOf("_") + 1, 14));
                txtFileName = fileName + ".txt";
                StreamWriter sw = new StreamWriter(Path.Combine(Path.GetDirectoryName(docPath), txtFileName));

                //foreach (var content in readContent.Select((value, index) => new { value, index }))
                foreach (string content in readContent)
                {
                    //Get commands from CAN log
                    List<string> commandCheckList = new List<string> { "_HEX", "_ascii", "_K_ABS", "_K_OBD", "_K_SEND", "_K_CLEAR", "_Temperature", "_FuelDisplay", "_WaterTemp" };
                    if (commandCheckList.Any(content.Contains)) //Check if this line is a command
                    {
                        string[] commandTimestamp = content.Split('_');
                        string[] commandSplit = content.Split(',');
                        command = commandSplit[9].Trim();
                        commandList.Add(command);
                        sw.WriteLine(commandTimestamp[0].Trim() + " Command: " + command);
                    }

                    line_CAN_ID = content.Substring(content.IndexOf("ID:") + 3, 3).Trim(); //ID
                    line_CAN_DATA = content.Substring(content.IndexOf("Data: ") + 6).Trim(); //DATA
                    string[] hexValuesSplit = line_CAN_DATA.Split(' '); //Number of DLC

                    switch (line_CAN_ID)
                    {
                        case "0x1": //CMD_A
                            string hexToBianry = Convert.ToString(Convert.ToInt32(hexValuesSplit[0], 16), 2).PadLeft(8, '0'); //Convert from hex to binary
                            char[] bianryToChar = hexToBianry.ToCharArray();
                            hexToBianry = new string(bianryToChar);

                            Power = (hexToBianry.Substring(7, 1) != "0") ? "Power = ON" : "Power = OFF";
                            OilPressureLight = (hexToBianry.Substring(6, 1) != "0") ? "Oil pressure light = ON" : "Oil pressure light = OFF";
                            FuelLight = (hexToBianry.Substring(5, 1) != "0") ? "Fuel light = ON" : "Fuel light = OFF";
                            ABSLight = (hexToBianry.Substring(4, 1) != "0") ? "ABS Light = ON" : "ABS Light = OFF";
                            WaterTempLight = (hexToBianry.Substring(3, 1) != "0") ? "Water temp light = ON" : "Water temp light = OFF";
                            MaintenanceLight = (hexToBianry.Substring(2, 1) != "0") ? "Maintenance light = ON" : "Maintenance Light = OFF";
                            FrontTirePressure = (hexToBianry.Substring(1, 1) != "0") ? "Front tyre pressure light = ON" : "Front tyre pressure light = OFF";
                            RearTirePressure = (hexToBianry.Substring(0, 1) != "0") ? "Rear tyre pressure light = ON" : "Rear tyre pressure light = OFF";

                            CMD_A_decode_result = string.Format("---Result: {0} , {1} , {2} , {3} , {4} , {5} , {6} , {7}"
                                , Power, OilPressureLight, FuelLight, ABSLight, WaterTempLight, MaintenanceLight, FrontTirePressure, RearTirePressure);
                            Console.WriteLine(CMD_A_decode_result);
                            sw.WriteLine(content + CMD_A_decode_result);
                            break;

                        case "0x2": //CMD_B
                                    //Speed
                            Speed = Convert.ToString(Convert.ToInt32(hexValuesSplit[0], 16));
                            int rpm_value = Convert.ToInt32(hexValuesSplit[1], 16); //Convert from hex to dec.

                            //RPM
                            RPM = Convert.ToString(rpm_value * 256 + Convert.ToInt32(hexValuesSplit[2], 16));

                            //Fuel Consumption
                            int fuelConsumption_value = Convert.ToInt32(hexValuesSplit[3], 16);
                            FuelConsumption = Convert.ToString(fuelConsumption_value * 256 + Convert.ToInt32(hexValuesSplit[4], 16));
                            if ((FuelConsumption == "65534") || (Convert.ToInt32(Speed) < 9))
                            {
                                FuelConsumption = "--.-";
                            }
                            else
                            {
                                float f_fuelConsumption = ((float)fuelConsumption_value) / 10;
                                //FuelConsumption = f_fuelConsumption.ToString(); //in Km/L
                                if (fuelConsumption_value > 0)
                                {
                                    int temp = 100000 / fuelConsumption_value;
                                    FuelConsumption = (((float)temp) / 100).ToString();
                                }
                            }

                            CMD_B_decode_result = string.Format("---Result: Speed = {0}Km/h , RPM = {1}Km/h , Fuel Consumption = {2}L/100Km",
                                      Speed, RPM, FuelConsumption);
                            sw.WriteLine(content + CMD_B_decode_result);
                            break;

                        case "0x3": //CMD_C
                                    //Water Temperature
                            int waterTemp_value = Convert.ToInt32(hexValuesSplit[0], 16);
                            if (waterTemp_value <= 8)
                            {
                                WaterTemp = Convert.ToString(Convert.ToInt32(hexValuesSplit[0], 16));
                            }
                            else
                            {
                                WaterTemp = "Out of range";
                            }

                            //Room Temperature
                            if (hexValuesSplit[1] == "ff" && hexValuesSplit[2] == "fe")
                            {
                                RoomTemp = "Temperature sensor failed";
                            }
                            else
                            {
                                RoomTemp = Convert.ToString(Convert.ToInt32(hexValuesSplit[2], 16));
                            }

                            //Temperature sign
                            if (hexValuesSplit[1] == "0")
                            {
                                RoomTemp = RoomTemp + "°C";
                            }
                            else if (hexValuesSplit[1] == "1")
                            {
                                RoomTemp = "-" + RoomTemp + "°C";
                            }

                            //Fuel
                            int fuel_value = Convert.ToInt32(hexValuesSplit[3], 16);
                            if (fuel_value <= 160)
                            {
                                float f_Fuel = ((float)fuel_value) / 10;
                                Fuel = string.Format("{0}L", f_Fuel.ToString());
                            }
                            else
                            {
                                Fuel = "Out Of range";
                            }

                            //Battery
                            int battery_value = Convert.ToInt32(hexValuesSplit[4], 16);
                            float f_Battery;
                            f_Battery = ((float)battery_value) / 10;
                            Battery = string.Format("{0}V", f_Battery.ToString());

                            CMD_C_decode_result = string.Format("---Result: Water temperature = {0} , Room temperature = {1} , Fuel = {2} , Battery = {3}",
                                        WaterTemp, RoomTemp, Fuel, Battery);
                            sw.WriteLine(content + CMD_C_decode_result);
                            break;

                        case "0x4": // CMD_D
                            Mileage = Convert.ToString(Convert.ToInt32(hexValuesSplit[0] + hexValuesSplit[1] + hexValuesSplit[2], 16));
                            MaxSpeed = Convert.ToString(Convert.ToInt32(hexValuesSplit[3], 16)) != "ff" ? Convert.ToString(Convert.ToInt32(hexValuesSplit[3], 16)) : "--";
                            AverageSpeed = Convert.ToString(Convert.ToInt32(hexValuesSplit[4], 16)) != "ff" ? Convert.ToString(Convert.ToInt32(hexValuesSplit[4], 16)) : "--";

                            CMD_D_decode_result = string.Format("---Result: Mileage = {0}Km , Max speed = {1}Km , Average speed = {2}Km/h", Mileage, MaxSpeed, AverageSpeed);
                            sw.WriteLine(content + CMD_D_decode_result);
                            break;

                        case "0x5": //CMD_E
                            int dataIndexABS = 0;
                            for (dataIndexABS = 0; dataIndexABS < hexValuesSplit.Length; dataIndexABS++)
                            {
                                string hToB = Convert.ToString(Convert.ToInt32(hexValuesSplit[dataIndexABS], 16), 2).PadLeft(8, '0');
                                char[] binaryReverse = hToB.ToCharArray();
                                Array.Reverse(binaryReverse);
                                hToB = new string(binaryReverse);

                                if (dataIndexABS == 0)
                                {
                                    for (int n = 0; n < 8; n++)
                                    {
                                        if (hToB.Substring(n, 1) != "0")
                                        {
                                            DTC_ABS(dataIndexABS, n);
                                        }
                                    }
                                }
                                else
                                {
                                    for (int b = 0; b < 6; b++)
                                    {
                                        if (hToB.Substring(b, 1) != "0")
                                        {
                                            DTC_ABS(dataIndexABS, b);
                                        }
                                    }
                                }
                            }

                            if (ABS_dtc_List.Count() > 0)
                            {
                                string tmpStr = string.Join(" , ", ABS_dtc_List.ToArray());
                                CMD_E_decode_result = string.Format("---ABS error: {0}", tmpStr);
                            }
                            else
                            {
                                CMD_E_decode_result = "---ABS error: None";
                            }
                            sw.WriteLine(content + CMD_E_decode_result);
                            ABS_dtc_List.Clear();
                            break;

                        case "0x6": //CMD_F
                            int dataIndexOBD = 0;
                            for (dataIndexOBD = 0; dataIndexOBD < hexValuesSplit.Length; dataIndexOBD++)
                            {
                                string hToB = Convert.ToString(Convert.ToInt32(hexValuesSplit[dataIndexOBD], 16), 2).PadLeft(8, '0');
                                char[] binaryReverse = hToB.ToCharArray();
                                Array.Reverse(binaryReverse);
                                hToB = new string(binaryReverse);

                                if (dataIndexOBD == 6)
                                {
                                    for (int n = 0; n < 3; n++)
                                    {
                                        if (hToB.Substring(n, 1) != "0")
                                        {
                                            DTC_OBD(dataIndexOBD, n);
                                        }
                                    }
                                }
                                else
                                {
                                    for (int b = 0; b < 8; b++)
                                    {
                                        if (hToB.Substring(b, 1) != "0")
                                        {
                                            DTC_OBD(dataIndexOBD, b);
                                        }
                                    }
                                }
                            }

                            if (OBD_dtc_List.Count() > 0)
                            {
                                string tmpStr = string.Join(" , ", OBD_dtc_List.ToArray());
                                CMD_F_decode_result = string.Format("---OBD error: {0}", tmpStr);
                            }
                            else
                            {
                                CMD_F_decode_result = "---OBD error: None";
                            }
                            sw.WriteLine(content + CMD_F_decode_result);
                            OBD_dtc_List.Clear();
                            break;

                        default:
                            break;
                    }
                }
                sw.Close();
                MessageBox.Show("Test log decoding is finished.", "Message");
            }
            else
            {
                MessageBox.Show("Please select a valid path!", "Error");
            }
        }

        // Generate report after analysing
        private void button_Generate_Click(object sender, EventArgs e)
        {
            if (fileName != "")
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Test report");

                    // Format for all cells (border & font)
                    int rowCount = commandList.Count + 2;
                    var rangeAll = worksheet.Range("B2:F" + rowCount);
                    var caontent = worksheet.Range("B3:F" + rowCount);
                    rangeAll.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    rangeAll.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    rangeAll.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    rangeAll.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    rangeAll.Style.Alignment.WrapText = true;
                    caontent.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Style.Font.FontSize = 12;
                    worksheet.Column(2).Width = 35;
                    worksheet.Column(3).Width = 37;
                    worksheet.Column(4).Width = 35;
                    worksheet.Column(5).Width = 40;
                    worksheet.Column(6).Width = 15;

                    // Format for headers
                    var header = worksheet.Range("B2:F2");
                    header.Style.Font.Bold = true;
                    header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    header.Style.Fill.BackgroundColor = XLColor.CadetBlue;
                    worksheet.Row(2).Height = 24;

                    // Headers
                    worksheet.Cells("B2").Value = "Timestamp of Command";
                    worksheet.Cells("C2").Value = "Command";
                    worksheet.Cells("D2").Value = "Timestamp of Response";
                    worksheet.Cells("E2").Value = "Response";
                    worksheet.Cells("F2").Value = "Result";

                    // Result
                    worksheet.Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Command column
                    for (int j = 0; j < commandList.Count; j++)
                    {
                        var commandColumn = worksheet.Cell(j + 3, 3);
                        commandColumn.Value = commandList[j];
                        commandColumn.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                    }

                    // Response column
                    StreamReader txt_sr = new StreamReader(Path.Combine(Path.GetDirectoryName(docPath) + "\\" + txtFileName));
                    string txt_line = string.Empty;
                    List<string> readTxt = new List<string>(); //For saving every line in txt 
                    while (!txt_sr.EndOfStream)
                    {
                        txt_line = txt_sr.ReadLine();
                        if (!string.IsNullOrEmpty(txt_line))
                        {
                            readTxt.Add(txt_line);
                        }
                    }
                    txt_sr.Close();

                    #region ---Analyse log---
                    bool commandIsFound = false;
                    bool kLineIsSend = false;
                    bool kLineIsClear = false;
                    int i = 0;
                    DateTime startDate = new DateTime();
                    foreach (string txt in readTxt)
                    {
                        string timestamp = txt.Substring(txt.IndexOf("[") + 1, txt.IndexOf("]") - txt.IndexOf("[") - 1).Trim(); //Get timestamp in txt file
                        string timestampForExcel = txt.Substring(txt.IndexOf("["), txt.IndexOf("]") - txt.IndexOf("[") + 1).Trim(); //Timestamp of command in Excel to show milliseconds
                        string pattern = "yyyy/MM/dd HH:mm:ss.fff";
                        DateTime endDate = DateTime.ParseExact(timestamp, pattern, null);

                        List<string> cmdEList = new List<string>() { "Control unit failure", "Valve Relay Fault", "Front Inlet valve failure", "Rear Inlet valve failure", "Front Outlet valve failure",
                        "Rear Outlet valve failure", "Battery Voltage fault (Over-Voltage)", "Battery Voltage fault (Under-voltage)", "Pump Motor Failure", "Front WSS ohmic failure", "Rear WSS ohmic failure",
                        "Front WSS plausibility failure", "Rear WSS plausibility failure", "WSS generic failure" };

                        List<string> cmdFList = new List<string>() { "Front DX vehicle speed sensor/signal(vehicle is equipped with NST node)", "Tyres Pressure Lamp", "ASR Lamp Command", "Pressure Sensor", "Air Temperature Sensor", "Engine Temperature Sensor", "Throttle Position Potentiometer Sensor", "Oxygen Sensor 1", "Oxygen Sensor 1 Heater Circuit",
                            "Oxygen Sensor 2", "Oxygen Sensor 2 Heater Circuit", "Cylinder 1 Injector Command", "Cylinder 2 Injector Command", "Engine OverTemperature State Recognition", "Fuel Pump / Load Relay Command", "Engine speed sensor - Electric", "Engine speed sensor - Functional", "Ignition coil 1 circuit",
                            "Ignition coil 2 circuit", "Secondary Air Valve", "Cooling fan relay command", "Vehicle speed sensor/signal", "Front vehicle speed sensor/signal from ABS", "Idle control / stepper motor control", "Engine Start Button", "Battery Voltage", "EEPROM error(Flash emul.)", "RAM error", "ROM error (Flash)", "Microprocessor error",
                            "Alternator \"Boost\" command diagnosis PIN 2", "Alternator \"Boost\" command diagnosis PIN 31", "MIL Lamp command","Engine Overtemperature Lamp Command", "Engine start enable lamp command", "Headlights Off relay command", "Light Feedback Signal Error State", "Tip Over", "Data buffer full and triggered by special events",
                            "Wheel spoke learn", "Rear vehicle speed sensor / signal from ABS", "Electric water pump command", "\"Bus Off\" CAN line diagnosis(NCM)", "\"Mute Node\" CAN line diagnosis(NCM)", "CAN line diagnosis \"ABS Node\"", "CAN line diagnosis \"UIN Node\"", "CAN line diagnosis \"NST node\"", "Dashboard Node Absent", "Immobilizer" };

                        if (txt.Contains("Command"))
                        {
                            worksheet.Cell(i + 3, 2).Value = timestampForExcel;
                            i++;
                            startDate = DateTime.ParseExact(timestamp, pattern, null); //Convert from string to DateTime
                            Console.WriteLine("Timestamp for Command: " + timestampForExcel);
                            string CommandIndex = "Command: ";
                            command = txt.Substring(txt.IndexOf(CommandIndex) + CommandIndex.Length).Trim(); //Get command only
                            commandIsFound = true;

                            //For KLine (send multiple commands at one time)
                            if (cmdEList.Contains(command, StringComparer.OrdinalIgnoreCase))
                            {
                                absCommandList.Add(command);
                            }
                            else if (cmdFList.Contains(command, StringComparer.OrdinalIgnoreCase))
                            {
                                obdCommandList.Add(command);
                            }
                            else if (command == "Kline SEND") //_K_SEND
                            {
                                kLineIsSend = true;
                                kLineIsClear = false;
                            }
                            else if (command == "Kline CLEAR") //_K_CLEAR
                            {
                                kLineIsClear = true;
                                kLineIsSend = false;
                                absCommandList.Clear();
                                obdCommandList.Clear();
                            }
                        }
                        else
                        {
                            endDate = DateTime.ParseExact(timestamp, pattern, null);
                            var ts = endDate - startDate;
                            char newLine = (char)10;
                            var responseTimeColumn = worksheet.Cell(i + 2, 4);
                            var responseColumn = worksheet.Cell(i + 2, 5);
                            var resultColumn = worksheet.Cell(i + 2, 6);

                            if (txt.Contains("Result: "))
                            {
                                commandSplit = command.Split('=').Select(x => x.Trim()).ToArray(); //Trim items in an array;
                                result = txt.Substring(txt.IndexOf("Result: ") + 8).Trim();
                                responseSplit = result.Split(new char[2] { ',', '=' }).Select(x => x.Trim()).ToArray(); //Trim items in an array
                                responseResult = responseSplit[Array.FindIndex(responseSplit, t => t.Equals(commandSplit[0], StringComparison.OrdinalIgnoreCase)) + 1].Trim();
                                frontTyreResponse = responseSplit[Array.FindIndex(responseSplit, f => f.Contains("Front tyre")) + 1].Trim();
                                rearTyreResponse = responseSplit[Array.FindIndex(responseSplit, f => f.Contains("Rear tyre")) + 1].Trim();
                            }
                            else if (txt.Contains("ABS error: "))
                            {
                                absResponse = txt.Substring(txt.IndexOf("ABS error: ") + 11).Trim();
                                absResponseSplit = absResponse.Split(',').Select(x => x.Trim()).ToArray(); //Trim ABS error response

                            }
                            else if (txt.Contains("OBD error: "))
                            {
                                obdResponse = txt.Substring(txt.IndexOf("OBD error: ") + 11).Trim();
                                obdResponseSplit = obdResponse.Split(',').Select(x => x.Trim()).ToArray(); //Trim OBD error response
                            }

                            List<string> cmdAList = new List<string>() { "Oil pressure light", "Fuel light", "ABS light", "Water temp light", "Maintenance light" };
                            List<string> cmdBList = new List<string>() { "Speed", "RPM", "Fuel consumption" };
                            //List<string> fTyreList = new List<string>() { "Front tyre over-pressure", "Front tyre low-pressure", "Front tyre is leaking", "Front tyre sensor error", "Low voltage of Front tyre sensor" };
                            //List<string> rTyreList = new List<string>() { "Rear tyre over-pressure", "Rear tyre low-pressure", "Rear tyre is leaking", "Rear tyre sensor error", "Low voltage of Rear tyre sensor" };

                            int absDifference = absToArray.Except(absResponseSplit, StringComparer.OrdinalIgnoreCase).Count();
                            int obdDifference = obdToArray.Except(obdResponseSplit, StringComparer.OrdinalIgnoreCase).Count();

                            if (commandIsFound)
                            {
                                #region ---CMD_A commands except for Power & Tyres---
                                if (cmdAList.Contains(commandSplit[0], StringComparer.OrdinalIgnoreCase)) //CMD_A commands except for Power & Tyres
                                {
                                    if (ts.TotalMilliseconds < 1000)
                                    {
                                        if (Array.FindIndex(responseSplit, t => t.Equals(commandSplit[0], StringComparison.OrdinalIgnoreCase)) != -1)
                                        {
                                            responseTimeColumn.Value = timestampForExcel; //Timestamp of response in Excel to show milliseconds
                                            responseColumn.Value = string.Format("{0} = {1}", commandSplit[0], responseResult);  //Response column except for Power
                                            if (commandSplit[1] == responseResult)
                                            {
                                                resultColumn.Value = "PASS"; //Result column for ABS
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                            }
                                            else
                                            {
                                                resultColumn.Value = "FAIL";
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        commandIsFound = false;
                                    }
                                }
                                #endregion

                                #region ---Front tyre--- 
                                else if (commandSplit[0].IndexOf("Front tyre", StringComparison.OrdinalIgnoreCase) != -1) //Front tyre 
                                {
                                    if (ts.TotalMilliseconds < 5000)
                                    {
                                        if (Array.FindIndex(responseSplit, f => f.Contains("Front tyre")) != -1)
                                        {
                                            responseTimeColumn.Value = timestampForExcel;
                                            responseColumn.Value = string.Format("{0} = {1}", "Front Tyre Pressure Light", frontTyreResponse);  //Response column except for Front tyre
                                            if (commandSplit[1] == frontTyreResponse)
                                            {
                                                resultColumn.Value = "PASS"; //Result column for Front tyre
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                            }
                                            else
                                            {
                                                resultColumn.Value = "FAIL";
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //worksheet.Cell(i + 2, 6).Value = "Time out of range";  //Response column except for Power
                                        commandIsFound = false;
                                    }
                                }
                                #endregion

                                #region ---Rear tyre--- 
                                else if (commandSplit[0].IndexOf("Rear tyre", StringComparison.OrdinalIgnoreCase) != -1) //Rear tyre 
                                {
                                    if (ts.TotalMilliseconds < 5000)
                                    {
                                        if (Array.FindIndex(responseSplit, f => f.Contains("Rear tyre")) != -1)
                                        {
                                            responseTimeColumn.Value = timestampForExcel;
                                            responseColumn.Value = string.Format("{0} = {1}", "Rear Tyre Pressure Light", rearTyreResponse);  //Response column except for Front tyre
                                            if (commandSplit[1] == rearTyreResponse)
                                            {
                                                resultColumn.Value = "PASS"; //Result column for Front tyre
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                            }
                                            else
                                            {
                                                resultColumn.Value = "FAIL";
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        commandIsFound = false;
                                    }
                                }
                                #endregion

                                #region ---Power---
                                else if (commandSplit[0] == "Power") //Power 
                                {
                                    if (ts.TotalMilliseconds < 5000) //Power ON needs more time
                                    {
                                        if (Array.IndexOf(responseSplit, commandSplit[0]) != -1)
                                        {
                                            responseTimeColumn.Value = timestampForExcel; //Timestamp of response in Excel to show milliseconds
                                            responseColumn.Value = string.Format("{0} = {1}", "Power", responseResult);  //Response column for Power
                                            if (commandSplit[1] == responseResult)
                                            {
                                                resultColumn.Value = "PASS"; //Result column for Power
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                            }
                                            else
                                            {
                                                resultColumn.Value = "FAIL";
                                                resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                            }

                                        }
                                    }
                                    else
                                    {
                                        commandIsFound = false;
                                    }
                                }
                                #endregion

                                #region ---CMD_B---
                                else if (cmdBList.Contains(commandSplit[0], StringComparer.OrdinalIgnoreCase)) //CMD_B
                                {
                                    int responseValue = 0;
                                    if (ts.TotalMilliseconds < 1000)
                                    {
                                        if (Array.IndexOf(responseSplit, commandSplit[0]) != -1)
                                        {
                                            responseTimeColumn.Value = timestampForExcel; //Timestamp of response in Excel to show milliseconds
                                            switch (commandSplit[0])
                                            {
                                                case "Speed":
                                                    responseValue = Convert.ToInt32(responseSplit[1].Substring(0, responseSplit[1].IndexOf("Km/h")));
                                                    if (commandSplit[1] == "0")
                                                    {
                                                        responseColumn.Value = string.Format("{0} = {1}", "Speed", responseResult);  //Response column for Speed = 0
                                                        if (responseValue == 0)
                                                        {
                                                            resultColumn.Value = "PASS"; //Result column for Speed
                                                            resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                        }
                                                        else
                                                        {
                                                            resultColumn.Value = "FAIL";
                                                            resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                        }
                                                    }
                                                    else if (commandSplit[1] == "10")
                                                    {
                                                        responseColumn.Value = string.Format("{0} = {1}", "Speed", responseResult);  //Response column for Speed = 1
                                                        if (responseValue >= 11 && responseValue <= 11)
                                                        {
                                                            resultColumn.Value = "PASS"; //Result column for Speed
                                                            resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                        }
                                                        else
                                                        {
                                                            resultColumn.Value = "FAIL";
                                                            resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        int speedHigh = Convert.ToInt32(commandSplit[1].Substring(commandSplit[1].IndexOf("~") + 1));
                                                        int speedLow = Convert.ToInt32(commandSplit[1].Substring(0, commandSplit[1].IndexOf("~")));
                                                        responseColumn.Value = string.Format("{0} = {1}", "Speed", responseResult);  //Response column for Speed
                                                        if (responseValue >= speedLow && responseValue <= speedHigh)
                                                        {
                                                            resultColumn.Value = "PASS"; //Result column for Speed
                                                            resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                        }
                                                        else
                                                        {
                                                            resultColumn.Value = "FAIL";
                                                            resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                        }
                                                    }
                                                    break;
                                                case "RPM":
                                                    int intCommandValue = Convert.ToInt32(commandSplit[1]);
                                                    responseValue = Convert.ToInt32(responseSplit[3].Substring(0, responseSplit[3].IndexOf("Km/h")));
                                                    responseColumn.Value = string.Format("{0} = {1}", "RPM", responseResult); //Response column for RPM
                                                    if (responseValue >= intCommandValue - 300 && responseValue <= intCommandValue + 300)
                                                    {
                                                        resultColumn.Value = "PASS"; //Result column for RPM
                                                        resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                    }
                                                    else
                                                    {
                                                        resultColumn.Value = "FAIL";
                                                        resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        commandIsFound = false;
                                    }
                                }
                                #endregion

                                #region---CMD_E---

                                else if (txt.Contains("ABS error")) //ABS error codes
                                {
                                    if (kLineIsSend && absCommandList.Count() > 0)
                                    {
                                        absToArray = absCommandList.ToArray();
                                        if (kLineIsClear != true)
                                        {
                                            absTmp = string.Join(" , ", absResponseSplit);
                                            if (obdCommandList.Count() > 0)
                                            {
                                                responseColumn.Value = string.Format("ABS error: " + absTmp + newLine + "OBD error: " + obdTmp);
                                                if (absToArray.Except(absResponseSplit, StringComparer.OrdinalIgnoreCase).Count() == 0 && obdToArray.Except(obdResponseSplit, StringComparer.OrdinalIgnoreCase).Count() == 0)
                                                {
                                                    resultColumn.Value = "PASS";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                }
                                                else
                                                {
                                                    resultColumn.Value = "FAIL";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                }
                                            }
                                            else
                                            {
                                                responseTimeColumn.Value = timestampForExcel;
                                                responseColumn.Value = string.Format("ABS error = {0}", absTmp);
                                                if (absToArray.Except(absResponseSplit, StringComparer.OrdinalIgnoreCase).Count() == 0) //check two arrays are same
                                                {
                                                    resultColumn.Value = "PASS";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                }
                                                else
                                                {
                                                    resultColumn.Value = "FAIL";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            absCommandList.Clear();
                                            commandIsFound = false;
                                        }
                                    }
                                }
                                #endregion

                                #region ---CMD_F---
                                else if (txt.Contains("OBD error")) //OBD error codes
                                {
                                    if (kLineIsSend && obdCommandList.Count() > 0)
                                    {
                                        obdToArray = obdCommandList.ToArray();
                                        if (kLineIsClear != true)
                                        {
                                            obdTmp = string.Join(" , ", obdResponseSplit);
                                            if (absCommandList.Count() > 0) //ABS & OBD commands
                                            {
                                                responseColumn.Value = string.Format("ABS error: " + absTmp + newLine + "OBD error: " + obdTmp);
                                                if (absToArray.Except(absResponseSplit, StringComparer.OrdinalIgnoreCase).Count() == 0 && obdToArray.Except(obdResponseSplit, StringComparer.OrdinalIgnoreCase).Count() == 0)
                                                {
                                                    resultColumn.Value = "PASS";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                }
                                                else
                                                {
                                                    resultColumn.Value = "FAIL";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                }
                                            }
                                            else //OBD command only
                                            {
                                                responseTimeColumn.Value = timestampForExcel;
                                                responseColumn.Value = string.Format("OBD error = {0}", obdTmp);

                                                if (obdToArray.Except(obdResponseSplit, StringComparer.OrdinalIgnoreCase).Count() == 0) //check two arrays are same
                                                {
                                                    resultColumn.Value = "PASS";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Green;
                                                }
                                                else
                                                {
                                                    resultColumn.Value = "FAIL";
                                                    resultColumn.Style.Fill.BackgroundColor = XLColor.Red;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            obdCommandList.Clear();
                                            commandIsFound = false;
                                        }
                                    }
                                }
                                #endregion
                            }

                        }
                    }
                    #endregion

                    workbook.SaveAs(Path.GetDirectoryName(docPath) + "\\" + fileName + ".xlsx");
                    workbook.Dispose();
                    worksheet = null;
                    MessageBox.Show("Report is generated.", "Message");
                    textBox_Path.Clear();
                }
            }
            else
            {
                MessageBox.Show("Please decode a log file first!", "Error");
            }
        }

        private void DTC_ABS(int data, int bit)
        {
            if (data == 0)
            {
                switch (bit)
                {
                    case 0:
                        ABS_dtc_List.Add("Control Unit Failure");
                        break;
                    case 1:
                        ABS_dtc_List.Add("Valve Replay Fault");
                        break;
                    case 2:
                        ABS_dtc_List.Add("Front Inlet valve failure");
                        break;
                    case 3:
                        ABS_dtc_List.Add("Rear Inlet valve failure");
                        break;
                    case 4:
                        ABS_dtc_List.Add("Front Outlet valve failure");
                        break;
                    case 5:
                        ABS_dtc_List.Add("Rear Outlet valve failure");
                        break;
                    case 6:
                        ABS_dtc_List.Add("Battery Voltage fault (Over-Voltage)");
                        break;
                    case 7:
                        ABS_dtc_List.Add("Battery Voltage fault (Under-voltage)");
                        break;
                    default:
                        break;
                }
            }
            else
            {
                switch (bit)
                {
                    case 0:
                        ABS_dtc_List.Add("Pump Motor Failure");
                        break;
                    case 1:
                        ABS_dtc_List.Add("Front WSS ohmic failure");
                        break;
                    case 2:
                        ABS_dtc_List.Add("Rear WSS ohmic failure");
                        break;
                    case 3:
                        ABS_dtc_List.Add("Front WSS plausibility failure");
                        break;
                    case 4:
                        ABS_dtc_List.Add("Rear WSS plausibility failure");
                        break;
                    case 5:
                        ABS_dtc_List.Add("WSS generic failure");
                        break;
                    default:
                        break;
                }
            }
        }

        private void DTC_OBD(int data, int bit)
        {
            switch (data)
            {
                case 0:
                    switch (bit)
                    {
                        case 0:
                            OBD_dtc_List.Add("Front DX vehicle speed sensor/signal (vehicle is equipped with NST node)");
                            break;
                        case 1:
                            OBD_dtc_List.Add("Tyres Pressure Lamp");
                            break;
                        case 2:
                            OBD_dtc_List.Add("ASR Lamp Command");
                            break;
                        case 3:
                            OBD_dtc_List.Add("Pressure Sensor");
                            break;
                        case 4:
                            OBD_dtc_List.Add("Air Temperature Sensor");
                            break;
                        case 5:
                            OBD_dtc_List.Add("Engine Temperature Sensor");
                            break;
                        case 6:
                            OBD_dtc_List.Add("Throttle Position Potentiometer Sensor");
                            break;
                        case 7:
                            OBD_dtc_List.Add("Oxygen Sensor 1");
                            break;
                        default:
                            break;
                    }
                    break;
                case 1:
                    switch (bit)
                    {
                        case 0:
                            OBD_dtc_List.Add("Oxygen Sensor 1 Heater Circuit");
                            break;
                        case 1:
                            OBD_dtc_List.Add("Oxygen Sensor 2");
                            break;
                        case 2:
                            OBD_dtc_List.Add("Oxygen Sensor 2 Heater Circuit");
                            break;
                        case 3:
                            OBD_dtc_List.Add("Cylinder 1 Injector Command");
                            break;
                        case 4:
                            OBD_dtc_List.Add("Cylinder 2 Injector Command");
                            break;
                        case 5:
                            OBD_dtc_List.Add("Engine OverTemperature State Recognition");
                            break;
                        case 6:
                            OBD_dtc_List.Add("Fuel Pump / Load Relay Command");
                            break;
                        case 7:
                            OBD_dtc_List.Add("Engine speed sensor - Electric");
                            break;
                        default:
                            break;
                    }
                    break;
                case 2:
                    switch (bit)
                    {
                        case 0:
                            OBD_dtc_List.Add("Engine speed sensor - Functional");
                            break;
                        case 1:
                            OBD_dtc_List.Add("Ignition coil 1 circuit");
                            break;
                        case 2:
                            OBD_dtc_List.Add("Ignition coil 2 circuit");
                            break;
                        case 3:
                            OBD_dtc_List.Add("Secondary Air Valve");
                            break;
                        case 4:
                            OBD_dtc_List.Add("Cooling fan relay command");
                            break;
                        case 5:
                            OBD_dtc_List.Add("Vehicle speed sensor/signal");
                            break;
                        case 6:
                            OBD_dtc_List.Add("Front vehicle speed sensor/signal from ABS");
                            break;
                        case 7:
                            OBD_dtc_List.Add("Idle control / stepper motor control");
                            break;
                        default:
                            break;
                    }
                    break;
                case 3:
                    switch (bit)
                    {
                        case 0:
                            OBD_dtc_List.Add("Engine Start Button");
                            break;
                        case 1:
                            OBD_dtc_List.Add("Battery Voltage");
                            break;
                        case 2:
                            OBD_dtc_List.Add("EEPROM error (Flash emul.)");
                            break;
                        case 3:
                            OBD_dtc_List.Add("RAM error");
                            break;
                        case 4:
                            OBD_dtc_List.Add("ROM error (Flash)");
                            break;
                        case 5:
                            OBD_dtc_List.Add("Microprocessor error");
                            break;
                        case 6:
                            OBD_dtc_List.Add("Alternator \"Boost\" command diagnosis PIN 2");
                            break;
                        case 7:
                            OBD_dtc_List.Add("Alternator \"Boost\" command diagnosis PIN 31");
                            break;
                        default:
                            break;
                    }
                    break;
                case 4:
                    switch (bit)
                    {
                        case 0:
                            OBD_dtc_List.Add("MIL Lamp command");
                            break;
                        case 1:
                            OBD_dtc_List.Add("Engine Overtemperature Lamp Command");
                            break;
                        case 2:
                            OBD_dtc_List.Add("Engine start enable lamp command");
                            break;
                        case 3:
                            OBD_dtc_List.Add("Headlights Off relay command");
                            break;
                        case 4:
                            OBD_dtc_List.Add("Light Feedback Signal Error State");
                            break;
                        case 5:
                            OBD_dtc_List.Add("Tip Over");
                            break;
                        case 6:
                            OBD_dtc_List.Add("Data buffer full and triggered by special events");
                            break;
                        case 7:
                            OBD_dtc_List.Add("Wheel spoke learn");
                            break;
                        default:
                            break;
                    }
                    break;
                case 5:
                    switch (bit)
                    {
                        case 0:
                            OBD_dtc_List.Add("Rear vehicle speed sensor / signal from ABS");
                            break;
                        case 1:
                            OBD_dtc_List.Add("Electric water pump command");
                            break;
                        case 2:
                            OBD_dtc_List.Add("\"Bus Off\" CAN line diagnosis (NCM)");
                            break;
                        case 3:
                            OBD_dtc_List.Add("\"Mute Node\" CAN line diagnosis (NCM)");
                            break;
                        case 4:
                            OBD_dtc_List.Add("CAN line diagnosis \"ABS Node\"");
                            break;
                        case 5:
                            OBD_dtc_List.Add("CAN line diagnosis \"UIN Node\"");
                            break;
                        case 6:
                            OBD_dtc_List.Add("CAN line diagnosis \"NST node\"");
                            break;
                        case 7:
                            OBD_dtc_List.Add("Dashboard Node Absent");
                            break;
                        default:
                            break;
                    }
                    break;
                case 6:
                    switch (bit)
                    {
                        case 0:
                            OBD_dtc_List.Add("Immobilizer");
                            break;
                        case 1:
                            OBD_dtc_List.Add("Immobilizer");
                            break;
                        default:
                            break;
                    }
                    break;
                default:
                    break;
            }
        }

        private void button_TxtToCSV_Click(object sender, EventArgs e)
        {
            if (textBox_Path.Text != "")
            {
                List<string> readContent = new List<string>(); //For saving every line 
                StreamReader sr = new StreamReader(textBox_Path.Text);
                while (!sr.EndOfStream)
                {
                    linePS = sr.ReadLine();
                    if (!string.IsNullOrEmpty(line))
                    {
                        readContent.Add(line);
                    }
                }
                sr.Close();

                docPath = textBox_Path.Text;
                fileName = string.Format("CsvForGraph_{0}", Path.GetFileName(docPath).Substring(Path.GetFileName(docPath).IndexOf("_") + 1, 14));
                txtFileName = fileName + ".csv";
                StreamWriter sw = new StreamWriter(Path.Combine(Path.GetDirectoryName(docPath), txtFileName));

                foreach (string content in readContent)
                {
                    //Get command from log
                    string command_PowerSupply = "VSET1:";
                    if (content.Contains(command_PowerSupply)) //Check if this line is a Power Supply command
                    {
                        string[] commandSplit = content.Split(',');
                        commandPS = commandSplit[5].Substring(content.IndexOf(command_PowerSupply) + 6, 3).Trim();
                        //sw.WriteLine("Power Supply Command: " + commandPS);
                        Console.WriteLine("Power Supply Command: " + commandPS);
                    }
                }
            }
        }
    }
}





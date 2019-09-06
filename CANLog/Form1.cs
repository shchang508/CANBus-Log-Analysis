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

namespace CANLog
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Select CAN log file from computer 
        private void button_SelectFileClick(object sender, EventArgs e)
        {
            OpenFileDialog path = new OpenFileDialog();
            path.FileName = "Select a text file";
            path.Filter = "Text files (*.txt)|*.txt";
            path.Title = "Open text file";
            path.ShowDialog();
            this.textBox_Path.Text = path.FileName;
        }

        List<string> commandList = new List<string>();
        List<string> responseList = new List<string>();


        //File path and name
        string docPath = string.Empty;
        string fileName = string.Empty;
        string txtFileName = string.Empty;

        //Raw data
        string line = string.Empty;
        string command = string.Empty;
        string line_CAN_ID = string.Empty;
        string line_CAN_DATA = string.Empty;

        //CMD_A
        string Power = string.Empty;
        string OilPressureLight = string.Empty;
        string FuelLight = string.Empty;
        string ABSLight = string.Empty;
        string WaterTempLight = string.Empty;
        string MaintenanceLight = string.Empty;
        string FrontTirePressure = string.Empty;
        string RearTirePressure = string.Empty;
        string CMD_A_decode_result = string.Empty;

        //CMD_B
        string Speed = string.Empty;
        string RPM = string.Empty;
        string FuelConsumption = string.Empty;
        string CMD_B_decode_result = string.Empty;

        //CMD_C
        string WaterTemp = string.Empty;
        string TempSign = string.Empty;
        string RoomTemp = string.Empty;
        string Fuel = string.Empty;
        string Battery = string.Empty;
        string CMD_C_decode_result = string.Empty;

        //CMD_D
        string Mileage = string.Empty;
        string MaxSpeed = string.Empty;
        string AverageSpeed = string.Empty;
        string CMD_D_decode_result = string.Empty;

        //Convert hex to text in CAN log
        private void button_Decode_Click(object sender, EventArgs e)
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

            docPath = @"D:\TestResult\"; ; //Location of test result files 
            if (Directory.Exists(docPath))
            {
                //Diretory exists
            }
            else
            {
                //Create a new directory
                Directory.CreateDirectory(docPath);
            }
            fileName = string.Format("TestResult_{0}", DateTime.Now.ToString("yyyyMMddHHmmss")); //File name of test result file 
            txtFileName = fileName + ".txt";
            StreamWriter sw = new StreamWriter(Path.Combine(docPath, txtFileName));


            //foreach (var content in readContent.Select((value, index) => new { value, index }))
            foreach (string content in readContent)
            {
                //Separate command from CAN log
                if (!content.Contains("ID"))
                {
                    string[] commandTimestamp = content.Split('_');
                    string[] commandSplit = content.Split(',');
                    command = commandSplit[10].Trim();
                    commandList.Add(command);
                    sw.WriteLine(commandTimestamp[0].Trim() + " Command: " + command);
                }

                line_CAN_ID = content.Substring(content.IndexOf("ID:") + 3, 3).Trim(); //ID
                line_CAN_DATA = content.Substring(content.IndexOf("Data: ") + 6).Trim(); //DATA
                string[] hexValuesSplit = line_CAN_DATA.Split(' ');

                switch (line_CAN_ID)
                {
                    case "0x1": //CMD_A
                        string hexToBianry = Convert.ToString(Convert.ToInt32(hexValuesSplit[0], 16), 2).PadLeft(8, '0'); //Convert from hex to binary
                        char[] bianry_reverse = hexToBianry.ToCharArray();
                        hexToBianry = new string(bianry_reverse);

                        Power = (hexToBianry.Substring(7, 1) != "0") ? "Power = ON" : "Power = OFF";
                        OilPressureLight = (hexToBianry.Substring(6, 1) != "0") ? "Oil Pressure Light = ON" : "Oil Pressure Light = OFF";
                        FuelLight = (hexToBianry.Substring(5, 1) != "0") ? "Fuel Light = ON" : "Fuel Light = OFF";
                        ABSLight = (hexToBianry.Substring(4, 1) != "0") ? "ABS Light = ON" : "ABS Light = OFF";
                        WaterTempLight = (hexToBianry.Substring(3, 1) != "0") ? "Water Temp Light = ON" : "Water Temp Light = OFF";
                        MaintenanceLight = (hexToBianry.Substring(2, 1) != "0") ? "Maintenance Light = ON" : "Maintenance Light = OFF";
                        FrontTirePressure = (hexToBianry.Substring(1, 1) != "0") ? "Front Tire Pressure = ON" : "Front Tire Pressure = OFF";
                        RearTirePressure = (hexToBianry.Substring(0, 1) != "0") ? "Rear Tire Pressure = ON" : "Rear Tire Pressure = OFF";

                        CMD_A_decode_result = string.Format("---Result: {0} , {1} , {2} , {3} , {4} , {5} , {6} , {7}"
                            , Power, OilPressureLight, FuelLight, ABSLight, WaterTempLight, MaintenanceLight, FrontTirePressure, RearTirePressure);
                        Console.WriteLine(CMD_A_decode_result);
                        sw.WriteLine(content + CMD_A_decode_result);
                        break;

                    case "0x2": //CMD_B
                        //Speed
                        Speed = Convert.ToString(Convert.ToInt32(hexValuesSplit[0], 16));
                        int rpm_value = Convert.ToInt32(hexValuesSplit[1], 16); // Convert from hex to dec.

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
                            Fuel = "Out Of Range";
                        }

                        //Battery
                        int battery_value = Convert.ToInt32(hexValuesSplit[4], 16);
                        float f_Battery;
                        f_Battery = ((float)battery_value) / 10;
                        Battery = string.Format("{0}V", f_Battery.ToString());

                        CMD_C_decode_result = string.Format("---Result: Water Temperature = {0} , Room Temperature = {1} , Fuel = {2} , Battery = {3}",
                                    WaterTemp, RoomTemp, Fuel, Battery);
                        sw.WriteLine(content + CMD_C_decode_result);
                        break;

                    case "0x4": // CMD_D
                        Mileage = Convert.ToString(Convert.ToInt32(hexValuesSplit[0] + hexValuesSplit[1] + hexValuesSplit[2], 16));
                        MaxSpeed = Convert.ToString(Convert.ToInt32(hexValuesSplit[3], 16)) != "ff" ? Convert.ToString(Convert.ToInt32(hexValuesSplit[3], 16)) : "--";
                        AverageSpeed = Convert.ToString(Convert.ToInt32(hexValuesSplit[4], 16)) != "ff" ? Convert.ToString(Convert.ToInt32(hexValuesSplit[4], 16)) : "--";

                        CMD_D_decode_result = string.Format("---Result: Mileage = {0}Km , Max Speed = {1}Km , Average Speed = {2}Km/h", Mileage, MaxSpeed, AverageSpeed);
                        sw.WriteLine(content + CMD_D_decode_result);
                        break;

                    case "0x5": //CMD_E

                        break;

                    case "0x6": //CMD_F

                        break;

                    default:
                        break;
                }
            }
            sw.Close();
            MessageBox.Show("Test log decoding is finished.", "Message");
        }

        //Generate report after analysing
        private void button_Generate_Click(object sender, EventArgs e)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test report");

                //Format for all cells (border & font)
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
                worksheet.Column(3).Width = 30;
                worksheet.Column(4).Width = 35;
                worksheet.Column(5).Width = 35;
                worksheet.Column(6).Width = 15;

                //Format for headers
                var header = worksheet.Range("B2:F2");
                header.Style.Font.Bold = true;
                header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                header.Style.Fill.BackgroundColor = XLColor.CadetBlue;
                worksheet.Row(2).Height = 24;

                //Headers
                worksheet.Cells("B2").Value = "Timestamp of Command";
                worksheet.Cells("C2").Value = "Command";
                worksheet.Cells("D2").Value = "Timestamp of Response";
                worksheet.Cells("E2").Value = "Response";
                worksheet.Cells("F2").Value = "Result";

                //Result
                worksheet.Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                List<string> readTxt = new List<string>(); //For saving every line in txt 

                for (int j = 0; j < commandList.Count; j++)
                {
                    //Command column
                    worksheet.Cell(j + 3, 3).Value = commandList[j];
                    worksheet.Cell(j + 3, 3).Style.Fill.BackgroundColor = XLColor.AliceBlue;
                }

                //Response column
                StreamReader txt_sr = new StreamReader(docPath + txtFileName);
                string txt_line = string.Empty;
                while (!txt_sr.EndOfStream)
                {
                    txt_line = txt_sr.ReadLine();
                    if (!string.IsNullOrEmpty(txt_line))
                    {
                        readTxt.Add(txt_line);
                    }
                }
                txt_sr.Close();

                #region ---Analyse command---
                bool commandIsFound = false;
                string CommandTemp = string.Empty;
                int i = 0;
                DateTime startDate = new DateTime();
                foreach (string txt in readTxt)
                {
                    string timestamp = txt.Substring(txt.IndexOf("[") + 1, txt.IndexOf("]") - txt.IndexOf("[") - 1).Trim(); //Get timestamp in txt file
                    string timestampForExcel = txt.Substring(txt.IndexOf("["), txt.IndexOf("]") - txt.IndexOf("[") + 1).Trim(); //Timestamp of command in Excel to show milliseconds
                    string pattern = "yyyy/MM/dd HH:mm:ss.fff";

                    DateTime endDate = DateTime.ParseExact(timestamp, pattern, null);
                    if (txt.Contains("Command"))
                    {
                        worksheet.Cell(i + 3, 2).Value = timestampForExcel;
                        i++;
                        startDate = DateTime.ParseExact(timestamp, pattern, null); //Convert from string to DateTime
                        Console.WriteLine("Timestamp for Command: " + timestampForExcel);
                        string CommandIndex = "Command: ";
                        CommandTemp = txt.Substring(txt.IndexOf(CommandIndex) + CommandIndex.Length).Trim();
                        commandIsFound = true;
                    }
                    else
                    {
                        endDate = DateTime.ParseExact(timestamp, pattern, null);
                        var ts = endDate - startDate;

                        string[] commandSplit = CommandTemp.Split('=').Select(x => x.Trim()).ToArray(); //Trim items in an array;
                        string result = txt.Substring(txt.IndexOf("Result: ") + 8).Trim();
                        string[] responseSplit = result.Split(new char[2] { ',', '=' }).Select(x => x.Trim()).ToArray(); //Trim items in an array
                        string responseResult = responseSplit[Array.IndexOf(responseSplit, commandSplit[0]) + 1].Trim();

                        List<string> cmdAList = new List<string>() { "Oil Pressure Light", "Fuel Light", "ABS Light", "Water Temp Light", "Maintenance Light", "Front Tire Pressure", "Rear Tire Pressure" };
                        List<string> cmdBList = new List<string>() { "Speed", "RPM", "Fuel Consumption" };

                        var resultColumn = worksheet.Cell(i + 2, 6);

                        if (commandIsFound)
                        {
                            if (cmdAList.Contains(commandSplit[0], StringComparer.OrdinalIgnoreCase)) //CMD_A commands except for Power
                            {
                                if (ts.TotalMilliseconds < 1000)
                                {
                                    if (Array.IndexOf(responseSplit, commandSplit[0]) != -1)
                                    {
                                        worksheet.Cell(i + 2, 4).Value = timestampForExcel; //Timestamp of response in Excel to show milliseconds
                                        worksheet.Cell(i + 2, 5).Value = string.Format("{0} = {1}", commandSplit[0], responseResult);  //Response column except for Power
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
                            else if (commandSplit[0] == "Power") //Power 
                            {
                                if (ts.TotalMilliseconds < 5000) //Power ON needs more time
                                {
                                    if (Array.IndexOf(responseSplit, commandSplit[0]) != -1)
                                    {
                                        worksheet.Cell(i + 2, 4).Value = timestampForExcel; //Timestamp of response in Excel to show milliseconds
                                        worksheet.Cell(i + 2, 5).Value = string.Format("{0} = {1}", "Power", responseResult);  //Response column for Power
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
                            else if (cmdBList.Contains(commandSplit[0], StringComparer.OrdinalIgnoreCase)) //CMD_B
                            {
                                int responseValue = 0;
                                if (ts.TotalMilliseconds < 1000)
                                {
                                    if (Array.IndexOf(responseSplit, commandSplit[0]) != -1)
                                    {
                                        worksheet.Cell(i + 2, 4).Value = timestampForExcel; //Timestamp of response in Excel to show milliseconds
                                        switch (commandSplit[0])
                                        {
                                            case "Speed":
                                                responseValue = Convert.ToInt32(responseSplit[1].Substring(0, responseSplit[1].IndexOf("Km/h")));
                                                if (commandSplit[1] == "0")
                                                {
                                                    worksheet.Cell(i + 2, 5).Value = string.Format("{0} = {1}", "Speed", responseResult);  //Response column for Speed
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
                                                    worksheet.Cell(i + 2, 5).Value = string.Format("{0} = {1}", "Speed", responseResult);  //Response column for Speed
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
                                                    worksheet.Cell(i + 2, 5).Value = string.Format("{0} = {1}", "Speed", responseResult);  //Response column for Speed
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
                                                worksheet.Cell(i + 2, 5).Value = string.Format("{0} = {1}", "RPM", responseResult); //Response column for RPM
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
                        }

                    }
                }
                #endregion

                workbook.SaveAs(docPath + fileName + ".xlsx");
                workbook.Dispose();
                worksheet = null;
                MessageBox.Show("Report is generated.", "Message");
            }
        }
    }
}



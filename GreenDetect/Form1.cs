using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Collections;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.MapProviders;
using GMap.NET;

namespace ComputerToArduino
{
    public partial class Form1 : Form

    {
        bool isConnected = false;
        String[] ports;
        SerialPort port;
        string[][] gen_settings = new string[102][];
        string[][] mod_settings = new string[61][];
        bool [][] mod_alarm = new bool[61][];
        bool[][] mod_alarm_mem = new bool[61][];
        bool first_reading = false;
        bool gen_modified = false;
        bool mod_modified = false;
        bool init = false;
        string[] raw1 = new string[61];
        string[] raw2 = new string[61];
        string[] raw3 = new string[61];
        string[] raw4 = new string[61];
        string[] value1 = new string[61];
        string[] value2 = new string[61];
        string[] value3 = new string[61];
        string[] value4 = new string[61];
        GMap.NET.WindowsForms.GMapOverlay[] markFlag1 = new GMap.NET.WindowsForms.GMapOverlay[61];
        GMap.NET.WindowsForms.GMapOverlay[] markFlag2 = new GMap.NET.WindowsForms.GMapOverlay[61];
        int read_count = 0;
        int timeout_count = 0;
        bool timeout_error = false;
        bool timeout_error_mem = false;
        bool new_alarm = false;
        StringBuilder sblog = new StringBuilder();

        Bitmap flag;
        public Form1()
        {
            InitializeComponent();
            
            getAvailableComPorts();
            timer1.Enabled = true;
            init = true;

            for (int w = 1; w <= 60; w++)
            {
                markFlag1[w] = new GMap.NET.WindowsForms.GMapOverlay();
                markFlag2[w] = new GMap.NET.WindowsForms.GMapOverlay();             
            }

         



            global.sen_ban.Add(sensorBanner0);
            global.sen_ban.Add(sensorBanner1);
            global.sen_ban.Add(sensorBanner2);
            global.sen_ban.Add(sensorBanner3);
            global.sen_ban.Add(sensorBanner4);
            global.sen_ban.Add(sensorBanner5);
            global.sen_ban.Add(sensorBanner6);
            global.sen_ban.Add(sensorBanner7);
            global.sen_ban.Add(sensorBanner8);
            global.sen_ban.Add(sensorBanner9);
            global.sen_ban.Add(sensorBanner10);
            global.sen_ban.Add(sensorBanner11);
            global.sen_ban.Add(sensorBanner12);
            global.sen_ban.Add(sensorBanner13);
            global.sen_ban.Add(sensorBanner14);
            global.sen_ban.Add(sensorBanner15);
            global.sen_ban.Add(sensorBanner16);
            global.sen_ban.Add(sensorBanner17);
            global.sen_ban.Add(sensorBanner18);
            global.sen_ban.Add(sensorBanner19);
            global.sen_ban.Add(sensorBanner20);
            global.sen_ban.Add(sensorBanner21);
            global.sen_ban.Add(sensorBanner22);
            global.sen_ban.Add(sensorBanner23);
            global.sen_ban.Add(sensorBanner24);
            global.sen_ban.Add(sensorBanner25);
            global.sen_ban.Add(sensorBanner26);
            global.sen_ban.Add(sensorBanner27);
            global.sen_ban.Add(sensorBanner28);
            global.sen_ban.Add(sensorBanner29);
            global.sen_ban.Add(sensorBanner30);
            global.sen_ban.Add(sensorBanner31);
            global.sen_ban.Add(sensorBanner32);
            global.sen_ban.Add(sensorBanner33);
            global.sen_ban.Add(sensorBanner34);
            global.sen_ban.Add(sensorBanner35);
            global.sen_ban.Add(sensorBanner36);
            global.sen_ban.Add(sensorBanner37);
            global.sen_ban.Add(sensorBanner38);
            global.sen_ban.Add(sensorBanner39);
            global.sen_ban.Add(sensorBanner40);
            global.sen_ban.Add(sensorBanner41);
            global.sen_ban.Add(sensorBanner42);
            global.sen_ban.Add(sensorBanner43);
            global.sen_ban.Add(sensorBanner44);
            global.sen_ban.Add(sensorBanner45);
            global.sen_ban.Add(sensorBanner46);
            global.sen_ban.Add(sensorBanner47);
            global.sen_ban.Add(sensorBanner48);
            global.sen_ban.Add(sensorBanner49);
            global.sen_ban.Add(sensorBanner50);
            global.sen_ban.Add(sensorBanner51);
            global.sen_ban.Add(sensorBanner52);
            global.sen_ban.Add(sensorBanner53);
            global.sen_ban.Add(sensorBanner54);
            global.sen_ban.Add(sensorBanner55);
            global.sen_ban.Add(sensorBanner56);
            global.sen_ban.Add(sensorBanner57);
            global.sen_ban.Add(sensorBanner58);
            global.sen_ban.Add(sensorBanner59);
            global.sen_ban.Add(sensorBanner60);

            

            for (int i = 1; i <= 9; i++)
            {
                global.sen_ban[i].address = "#0" + i.ToString();
                global.sen_ban[i].Visible = false;
                
            }

            for (int i = 10; i <= 60; i++)
            {
                global.sen_ban[i].address = "#" + i.ToString();
                global.sen_ban[i].Visible = false;
            }

            for (int i = 0; i <= 60; i++)
            {
                raw1[i] = "0";
                raw2[i] = "0";
                raw3[i] = "0";
                raw4[i] = "0";
                value1[i] = "0";
                value2[i] = "0";
                value3[i] = "0";
                value4[i] = "0";
                mod_alarm[i] = new bool[] { false, false, false, false, false, false, false, false, false, false, false, false, false, false };
                mod_alarm_mem [i] = new bool[] { true, false, false, false, false, false, false, false, false, false, false, false, false, false };
            }


            //General settings reading
            string doc_path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var path = doc_path + "\\GreenDetect\\general_settings.csv";            
            using (TextFieldParser csvReader = new TextFieldParser(path))
            {
                csvReader.CommentTokens = new string[] { "#" };
                csvReader.SetDelimiters(new string[] { ";" });
                csvReader.HasFieldsEnclosedInQuotes = true;

                // Skip the row with the column names
               // csvReader.ReadLine();
                
                int x = 0;

                while (!csvReader.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvReader.ReadFields();
                   // gen_settings [0] = new string[] { "numbers", "type", "description", "f1_en", "f1_size", "f1_unit", "f1_high", "f1_low", "f1_scale", "f2_en", "f2_size", "f2_unit", "f2_high", "f2_low", "f2_scale", "f3_en", "f3_size", "f3_unit", "f3_high", "f3_low", "f3_scale" };

                    
                    gen_settings [x] = new string [] { fields [0], fields[1] , fields[2], fields[3], fields[4], fields[5], fields[6], fields[7], fields[8], fields[9], fields[10], fields[11], fields[12], fields[13], fields[14], fields[15], fields[16], fields[17], fields[18], fields[19], fields[20] };
                   
                    x++;
                }
            }


            

            // Sensor Type Settings starting values
            comboBoxTypeSet.Items.Clear();
             int z = int.Parse(gen_settings[2][0]);
             for (int y = 0; y <= z; y++)
            {
                 comboBoxTypeSet.Items.Add(y);
            }
            comboBoxTypeSet.SelectedIndex = 0;  
            groupBoxSet1.Visible = bool.Parse(gen_settings[100][3]);
            groupBoxSet2.Visible = bool.Parse(gen_settings[100][9]);    
            groupBoxSet3.Visible = bool.Parse(gen_settings[100][15]);
            
            labelTipeDescription.Text = gen_settings[100][2];
                        
            textBoxF1size.Text = gen_settings[100][4];
            textBoxF1unit.Text = gen_settings[100][5];
            textBoxF1high.Text = gen_settings[100][6];
            textBoxF1low.Text = gen_settings[100][7];
            textBoxF1maxScale.Text = gen_settings[100][8];

            textBoxF2size.Text = gen_settings[100][10];
            textBoxF2unit.Text = gen_settings[100][11];
            textBoxF2high.Text = gen_settings[100][12];
            textBoxF2low.Text = gen_settings[100][13];
            textBoxF2maxScale.Text = gen_settings[100][14];

            textBoxF3size.Text = gen_settings[100][16];
            textBoxF3unit.Text = gen_settings[100][17];
            textBoxF3high.Text = gen_settings[100][18];
            textBoxF3low.Text = gen_settings[100][19];
            textBoxF3maxScale.Text = gen_settings[100][20];

            
            //Module settings reading
            string doc_path1 = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var path1 = doc_path1 + "\\GreenDetect\\module_settings.csv";
            using (TextFieldParser csvReader1 = new TextFieldParser(path1))
            {
                csvReader1.CommentTokens = new string[] { "#" };
                csvReader1.SetDelimiters(new string[] { ";" });
                csvReader1.HasFieldsEnclosedInQuotes = true;

                 //Skip the row with the column names
                 csvReader1.ReadLine();

                int x = 0;

                while (!csvReader1.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvReader1.ReadFields();
                    mod_settings[x] = new string[] { fields[0], fields[1], fields[2], fields[3], fields[4], fields[5], fields[6], fields[7] };

                    x++;
                }
            }

            // Module Settings starting value
            comboBoxAddress.Items.Clear();
            int h = int.Parse(gen_settings[1][0]);
            for (int y = 1; y <= h; y++)
            {
                comboBoxAddress.Items.Add(y);
            }
            comboBoxAddress.SelectedIndex = 0;

            comboBoxTypeMod.Items.Clear();
            int j = int.Parse(gen_settings[2][0]);
            for (int y = 0; y <= j; y++)
            {
                comboBoxTypeMod.Items.Add(y);
            }

            int type_sel_mod = int.Parse(mod_settings[1][1]);
            comboBoxTypeMod.SelectedIndex = type_sel_mod;
            int type_sel_gen = type_sel_mod;
            if (type_sel_mod == 0)
            {
                type_sel_gen = 100;
            }
            
            groupBoxF1mod.Visible = bool.Parse(gen_settings[type_sel_gen][3]);
            groupBoxF2mod.Visible = bool.Parse(gen_settings[type_sel_gen][9]);
            groupBoxF3mod.Visible = bool.Parse(gen_settings[type_sel_gen][15]);

            textBoxSocOffset.Text = mod_settings[1][2];

            textBoxF1sizeMod.Text = gen_settings[type_sel_gen][4];
            textBoxF1unitMod.Text = gen_settings[type_sel_gen][5];
            textBoxF1Offset.Text = mod_settings[1][3];

            textBoxF2sizeMod.Text = gen_settings[type_sel_gen][10];
            textBoxF2unitMod.Text = gen_settings[type_sel_gen][11];
            textBoxF2Offset.Text = mod_settings[1][4];

            textBoxF3sizeMod.Text = gen_settings[type_sel_gen][16];
            textBoxF3unitMod.Text = gen_settings[type_sel_gen][17];
            textBoxF3Offset.Text = mod_settings[1][5];

            //Address setting selection
            int temp_item = int.Parse(gen_settings[1][0]) - 1;
            comboBox2.SelectedIndex = temp_item;

            //Banner properties initialization            
            for (int i = 1; i <= int.Parse(gen_settings[1][0]); i++)
            {
                //group visualization
                int type_i = int.Parse(mod_settings[i][1]);
                if (type_i == 0)
                {
                    type_i = 100;
                }
                global.sen_ban[i].group2 = bool.Parse(gen_settings[type_i][3]);
                global.sen_ban[i].group3 = bool.Parse(gen_settings[type_i][9]);
                global.sen_ban[i].group4 = bool.Parse(gen_settings[type_i][15]);

                //Title
                global.sen_ban[i].title2 = gen_settings[type_i][4];
                global.sen_ban[i].title3 = gen_settings[type_i][10];
                global.sen_ban[i].title4 = gen_settings[type_i][16];

                //Unit
                global.sen_ban[i].unit2 = gen_settings[type_i][5];
                global.sen_ban[i].unit3 = gen_settings[type_i][11];
                global.sen_ban[i].unit4 = gen_settings[type_i][17];

            }


            gMap1.ShowCenter = false;
            //mainMap.MapProvider = GMap.NET.MapProviders.GoogleMapProvider.Instance;
            gMap1.MapProvider = GMap.NET.MapProviders.OpenStreetMapProvider.Instance;

            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerAndCache;
            gMap1.CacheLocation = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\GreenDetect\\cache";
           
            //GMap.NET.MapProviders.GMapProvider.WebProxy = null;
            double lat1 = Convert.ToDouble(gen_settings[3][0]);
            double longt1 = Convert.ToDouble(gen_settings[4][0]);
            gMap1.Position = new GMap.NET.PointLatLng(lat1, longt1);
            gMap1.DragButton = MouseButtons.Left;
            gMap1.Bearing = int.Parse(gen_settings[5][0]);
            gMap1.MinZoom = 1;
            gMap1.MaxZoom = 19;
            gMap1.Zoom = int.Parse(gen_settings[6][0]);

            gMap2.ShowCenter = false;
            //mainMap.MapProvider = GMap.NET.MapProviders.GoogleMapProvider.Instance;
            gMap2.MapProvider = GMap.NET.MapProviders.OpenStreetMapProvider.Instance;
            gMap1.CacheLocation = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\GreenDetect\\cache";
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerAndCache;

            //GMap.NET.MapProviders.GMapProvider.WebProxy = null;
            double lat2 = Convert.ToDouble(gen_settings[7][0]);
            double longt2 = Convert.ToDouble(gen_settings[8][0]);      
            gMap2.Position = new GMap.NET.PointLatLng(lat2, longt2);
            gMap2.DragButton = MouseButtons.Left;
            gMap2.Bearing = int.Parse(gen_settings[9][0]);
            gMap2.MinZoom = 1;
            gMap2.MaxZoom = 19;            
            gMap2.Zoom = int.Parse(gen_settings[10][0]);


            //Alarms list reading
            string doc_path3 = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var path3 = doc_path3 + "\\GreenDetect\\alarms.csv";
            using (TextFieldParser csvReader1 = new TextFieldParser(path3))
            {
                csvReader1.CommentTokens = new string[] { "#" };
                csvReader1.SetDelimiters(new string[] { ";" });
                csvReader1.HasFieldsEnclosedInQuotes = true;

                //Skip the row with the column names
                csvReader1.ReadLine();

                while (!csvReader1.EndOfData)
                {

                    string[] fields = csvReader1.ReadFields();
                    ListViewItem item = new ListViewItem(fields);
                    string status_check = item.SubItems[2].Text;
                    status_check = item.SubItems[2].Text;
                    listViewAlarm.Items.Add(item);
                    if (status_check == "ON")
                    {
                        item.ForeColor = Color.IndianRed;
                    }
                    if (status_check == "OFF")
                    {
                        item.ForeColor = Color.Green;
                    }
                    

                }

            }

            //log reading
            string doc_path4 = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var path4 = doc_path4 + "\\GreenDetect\\log\\" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_log.csv";
          
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(path4))
                {
                    csvReader.CommentTokens = new string[] { "#" };
                    csvReader.SetDelimiters(new string[] { ";" });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    // Skip the row with the column names
                    // csvReader.ReadLine();

                

                    while (!csvReader.EndOfData)
                    {                    
                        sblog.AppendLine(csvReader.ReadLine());
                    }
                }
            }
            catch (FileNotFoundException e)
            {
                int mod_num = int.Parse(gen_settings[1][0]);
                int column = 1 + (mod_num * 4);
                string strSeperator = ";";
                // string[] line = new string[] { "Date and time", "Supercap voltage", "value 1", "value 2", "value 3" };
                string[] line = new string [column];
                line[0] = "Date and time";
                int k = 1;
                for (int i = 1; i <= mod_num; i++)
                {
                    string sens = "S." + i.ToString(); 
                    line[k] = sens + " Cap Volt";
                    k++;
                    line[k] = sens + " Val1";
                    k++;
                    line[k] = sens + " Val2";
                    k++;
                    line[k] = sens + " Val3";
                    k++;
                }
                sblog.AppendLine(string.Join(strSeperator, line));
            }




            if (!isConnected)
            {
                for (int i = 1; i <= 60; i++)
                 {
                        global.sen_ban[i].data1 = "####";
                        global.sen_ban[i].data2 = "####";
                        global.sen_ban[i].data3 = "####";
                        global.sen_ban[i].data4 = "####";              
                }
                textBoxRaw1.Text = "####";
                textBoxRaw2.Text = "####";
                textBoxRaw3.Text = "####";
                textBoxRaw4.Text = "####";
                textBoxSocValue.Text = "####";
                textBoxF1Value.Text = "####";
                textBoxF2Value.Text = "####";
                textBoxF3Value.Text = "####";
            }


                   
            
            





            init = false;
        }



        private void button1_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                connectToArduino();
            } else
            {
                disconnectFromArduino();
            }
        }

        void getAvailableComPorts()
        {
            ports = SerialPort.GetPortNames();
            comboBox1.Items.Clear();
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
                Console.WriteLine(port);
                if (ports[0] != null)
                {
                    comboBox1.SelectedItem = ports[0];
                }
            }
        }

        private void connectToArduino()
        {
            
            string selectedPort = comboBox1.GetItemText(comboBox1.SelectedItem);
            port = new SerialPort(selectedPort, 115200, Parity.None, 8, StopBits.One);
            port.ReadTimeout = 30000;
            port.Open();
            port.DiscardInBuffer();
            button1.Text = "Disconnect";            
            first_reading = true;
            groupBox7.BackColor = Color.Green;
            isConnected = true;
            label50.Visible = true;
            labelReadTime.Visible = true;
        }

        

       

        private void disconnectFromArduino()
        {
            isConnected = false;
            port.Close();
            button1.Text = "Connect";            
            groupBox7.BackColor = Color.IndianRed;
            labelReadTime.Visible = false;
            label50.Visible = false;
            first_reading = false;
            refreshMarkers();

            for (int i = 1; i <= 60; i++)
            {
                global.sen_ban[i].data1 = "####";
                global.sen_ban[i].data2 = "####";
                global.sen_ban[i].data3 = "####";
                global.sen_ban[i].data4 = "####";
            }
            textBoxRaw1.Text = "####";
            textBoxRaw2.Text = "####";
            textBoxRaw3.Text = "####";
            textBoxRaw4.Text = "####";
            textBoxSocValue.Text = "####";
            textBoxF1Value.Text = "####";
            textBoxF2Value.Text = "####";
            textBoxF3Value.Text = "####";

            
        }

       

       

        private void flagBuilding(Image img, string text)
        {
            RectangleF rectf = new RectangleF(0, 0, img.Width, img.Height);
            Graphics g = Graphics.FromImage(img);
            StringFormat format = new StringFormat()
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(text, new Font("Verdana", 11), Brushes.White,
            rectf, format);
        }


        private void refreshMarkers()
        {
            for (int i = 1; i <= int.Parse(gen_settings[1][0]); i++)
            {
                markFlag1[i].Markers.Clear();
                markFlag2[i].Markers.Clear();

                string doc_path2 = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Bitmap flag = (Bitmap)Image.FromFile(doc_path2 + "\\GreenDetect\\map flags\\red.png");

                flagBuilding(flag, i.ToString());
                double flag_lat = Convert.ToDouble(mod_settings[i][6]);
                double flag_longt = Convert.ToDouble(mod_settings[i][7]);
                GMap.NET.WindowsForms.GMapMarker mark1 =
                new GMap.NET.WindowsForms.Markers.GMarkerGoogle(new GMap.NET.PointLatLng(flag_lat, flag_longt), flag);
                markFlag1[i].Markers.Add(mark1);
                gMap1.Overlays.Add(markFlag1[i]);
                gMap1.Zoom--;
                gMap1.Zoom++;

                GMap.NET.WindowsForms.GMapMarker mark2 =
                new GMap.NET.WindowsForms.Markers.GMarkerGoogle(new GMap.NET.PointLatLng(flag_lat, flag_longt), flag);
                markFlag2[i].Markers.Add(mark2);
                mark2.Tag = i.ToString();
                gMap2.Overlays.Add(markFlag2[i]);
                gMap2.Zoom--;
                gMap2.Zoom++;

                mod_alarm_mem[i][0] = true;
            }

        }

        
        private void dataRecord()
        {
            string doc_path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string strFilePath = doc_path + "\\GreenDetect\\log\\" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_log.csv";
            string strSeperator = ";";

            int mod_num = int.Parse(gen_settings[1][0]);
            int column = 1 + (mod_num * 4);

            // string[] line = new string[] { "Date and time", "Supercap voltage", "value 1", "value 2", "value 3" };
            string[] line = new string[column];
            line[0] = DateTime.Now.ToString();
            int k = 1;
            for (int i = 1; i <= mod_num; i++)
            {

                line[k] = value1[i];
                k++;
                line[k] = value2[i];
                k++;
                line[k] = value3[i];
                k++;
                line[k] = value4[i];
                k++;
            }
            sblog.AppendLine(string.Join(strSeperator, line));

            // Create and write the csv file
            File.WriteAllText(strFilePath, sblog.ToString());
        }


        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
            if (isConnected)
            {
                
                global.serial_data = port.ReadLine();

                if ((global.serial_data.Contains("&")) || (global.serial_data.Contains("#")))

                {                
                    
                    //raw data
                    if (global.serial_data.Contains("#"))
                    {
                        string index = global.serial_data.Remove(0, 1);
                        global.index = int.Parse(index);                            
                    }

                    if (global.serial_data.Contains("&1"))
                    {
                        string data_str = global.serial_data.Remove(0, 2);
                        raw1 [global.index] = data_str;
                    }

                    if (global.serial_data.Contains("&2"))
                    {
                        string data_str = global.serial_data.Remove(0, 2);
                        raw2 [global.index] = data_str;
                    }

                    if (global.serial_data.Contains("&3"))
                    {
                        string data_str = global.serial_data.Remove(0, 2);
                        raw3[global.index] = data_str;
                    }

                    if (global.serial_data.Contains("&4"))
                    {
                        string data_str = global.serial_data.Remove(0, 2);
                        raw4[global.index] = data_str;
                    }

                    if (global.serial_data.Contains("&5"))
                    {
                        string data_str = global.serial_data.Remove(0, 2);
                        read_count = int.Parse (data_str);
                        dataRecord();


                        if (first_reading)
                        {
                            global.index = 0;   
                            first_reading = false;
                            
                        }
                    }


                    //Supercap voltage
                    int data_byte = int.Parse(raw1[global.index]);
                    int data_byte_soc = data_byte & 63;
                    int soc_scaled = 1500 + (data_byte_soc * 25);
                    int soc_compensed = soc_scaled + int.Parse(mod_settings[global.index][2]);
                    value1[global.index] = soc_compensed.ToString();                  


                    //test
                    int data_byte_test = data_byte & 128;
                    
                    if (data_byte_test == 128)
                    {
                        global.sen_ban[global.index].BackColor = Color.DarkGreen;
                        mod_alarm[global.index][12] = true;

                    }
                    else
                    {
                        global.sen_ban[global.index].BackColor = Color.Black;
                        mod_alarm[global.index][12] = false;
                    }


                    int type_input = int.Parse(mod_settings [global.index] [1]);

                    // Sensor type 0 = repeater
                    if (type_input == 0)
                    {
                        //gestione repeater
                    }

                    // Sensor type 1 = SR01 ultrasonic sensor
                    if (type_input == 1)
                    {
                        int val1 = int.Parse(raw2[global.index]) ;
                        int val2 = int.Parse(raw3[global.index]) << 8;
                        int result = val1 | val2;
                        int result_compensed = result + int.Parse(mod_settings[global.index][3]);
                        value2 [global.index] = result_compensed.ToString();                       
                    }


                    // Sensor type 2 = HTU21D+IR temperature-Humidity-IR
                    if (type_input == 2)
                    {
                        //temperature
                        int val1 = int.Parse(raw2[global.index]);
                        int val2 = int.Parse(raw3[global.index]) << 8;
                        int result = (val1 | val2) / 10;
                        int result_compensed = result + int.Parse(mod_settings[global.index][3]);
                        value2[global.index] = result_compensed.ToString();

                        //Humidity
                        int val3 = int.Parse(raw4[global.index]);
                        int val3_compensed = val3 + int.Parse(mod_settings[global.index][4]);
                        value3[global.index] = val3_compensed.ToString();

                        //IR
                        int data_IR = int.Parse(raw1[global.index]) & 64;

                        if (data_IR == 64)
                        {
                            value4[global.index] = "TRUE";
                        }
                        else
                        {
                            value4[global.index] = "false";
                        }

                    }

                    // Sensor type 3 = DS18B20 temperature
                    if (type_input == 3)
                    {
                        int val1 = int.Parse(raw2[global.index]);
                        int val2 = int.Parse(raw3[global.index]) << 8;
                        int result = (val1 | val2);
                        double result_real = (double)result / 10.0;
                        double result_compensed = result_real + double.Parse(mod_settings[global.index][3]);
                        value2[global.index] = result_compensed.ToString();

                    }

                    // Sensor type 4 = SEN-13322 Sparkfun soil moisture
                    if (type_input == 4)
                    {
                        int val1 = int.Parse(raw2[global.index]);
                        int val2 = int.Parse(raw3[global.index]) << 8;
                        int result_point = val1 | val2;
                        int result_perc = (result_point * 100) / int.Parse (gen_settings[4][8]);
                        if (result_perc > 100)
                        {
                            result_perc = 100;
                        }

                        value2[global.index] = result_perc.ToString();
                    }


                    //banner update
                    global.sen_ban[global.index].data1 = value1[global.index];
                    global.sen_ban[global.index].data2 = value2[global.index];
                    global.sen_ban[global.index].data3 = value3[global.index];
                    global.sen_ban[global.index].data4 = value4[global.index];
                    int mapBannerAddrInt = int.Parse(global.mapBannerAddr);
                    sensorBannerMap.data1 = value1[mapBannerAddrInt];
                    sensorBannerMap.data2 = value2[mapBannerAddrInt];
                    sensorBannerMap.data3 = value3[mapBannerAddrInt];
                    sensorBannerMap.data4 = value4[mapBannerAddrInt];
                    

                    //Settings page update
                    if (comboBoxAddress.Text == global.index.ToString())
                    {
                        textBoxRaw1.Text = raw1[global.index];
                        textBoxRaw2.Text = raw2[global.index];
                        textBoxRaw3.Text = raw3[global.index];
                        textBoxRaw4.Text = raw4[global.index];

                        textBoxSocValue.Text = value1[global.index];
                        textBoxF1Value.Text = value2[global.index];
                        textBoxF2Value.Text = value3[global.index];
                        textBoxF3Value.Text = value4[global.index];

                    }

                    
                    if (!first_reading && (global.index > 0))
                    {

                        //module error

                        if ((int.Parse(raw1[global.index]) == 0) && (int.Parse(raw2[global.index]) == 0) && (int.Parse(raw3[global.index]) == 0) && (int.Parse(raw3[global.index]) == 0))
                        {
                            mod_alarm[global.index][13] = true;                          
                        }
                        else
                        {
                            mod_alarm[global.index][13] = false;
                        }

                        if (mod_alarm[global.index][13] != mod_alarm_mem[global.index][13])
                        {
                            if (mod_alarm[global.index][13])
                            {
                                string alarmText = "Sensor #" + global.index.ToString() + " module error";
                                ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                listViewAlarm.Items.Insert(0, item);
                                item.ForeColor = Color.IndianRed;
                                mod_alarm_mem[global.index][13] = true;
                                new_alarm = true;
                            }

                            else
                            {
                                string alarmText = "Sensor #" + global.index.ToString() + " module error";
                                ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                listViewAlarm.Items.Insert(0, item);
                                item.ForeColor = Color.Green;
                                mod_alarm_mem[global.index][13] = false;
                            }
                        }

                        if (!mod_alarm[global.index][13])
                        {
                            //High Supercap voltage alarm
                            if (int.Parse(value1[global.index]) >= 2800)
                            {
                                mod_alarm[global.index][1] = true;
                            }
                            else
                            {
                                mod_alarm[global.index][1] = false;
                            }
                            if (mod_alarm[global.index][1] != mod_alarm_mem[global.index][1])
                            {
                                if (mod_alarm[global.index][1])
                                {
                                    string alarmText = "Sensor #" + global.index.ToString() + " High Supercapacitor voltage";
                                    ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                    listViewAlarm.Items.Insert(0, item);
                                    item.ForeColor = Color.IndianRed;
                                    mod_alarm_mem[global.index][1] = true;
                                    new_alarm = true;
                                }

                                else
                                {
                                    string alarmText = "Sensor #" + global.index.ToString() + " High Supercapacitor voltage";
                                    ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                    listViewAlarm.Items.Insert(0, item);
                                    item.ForeColor = Color.Green;
                                    mod_alarm_mem[global.index][1] = false;
                                }
                            }


                            //Low Supercap voltage alarm
                            if (int.Parse(value1[global.index]) <= 1950)
                            {
                                mod_alarm[global.index][2] = true;
                            }
                            else
                            {
                                mod_alarm[global.index][2] = false;
                            }
                            if (mod_alarm[global.index][2] != mod_alarm_mem[global.index][2])
                            {
                                if (mod_alarm[global.index][2])
                                {
                                    string alarmText = "Sensor #" + global.index.ToString() + " Low Supercapacitor voltage";
                                    ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                    listViewAlarm.Items.Insert(0, item);
                                    item.ForeColor = Color.IndianRed;
                                    mod_alarm_mem[global.index][2] = true;
                                    new_alarm = true;
                                }

                                else
                                {
                                    string alarmText = "Sensor #" + global.index.ToString() + " Low Supercapacitor voltage";
                                    ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                    listViewAlarm.Items.Insert(0, item);
                                    item.ForeColor = Color.Green;
                                    mod_alarm_mem[global.index][2] = false;
                                }
                            }

                            //Test Button alarm
                            if (mod_alarm[global.index][12] != mod_alarm_mem[global.index][12])
                            {
                                if (mod_alarm[global.index][12])
                                {
                                    string alarmText = "Sensor #" + global.index.ToString() + " test button pressed";
                                    ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                    listViewAlarm.Items.Insert(0, item);
                                    item.ForeColor = Color.IndianRed;
                                    mod_alarm_mem[global.index][12] = true;
                                    new_alarm = true;
                                }

                                else
                                {
                                    string alarmText = "Sensor #" + global.index.ToString() + " test button pressed";
                                    ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                    listViewAlarm.Items.Insert(0, item);
                                    item.ForeColor = Color.Green;
                                    mod_alarm_mem[global.index][12] = false;
                                }
                            }

                        }



                        //other fields alarms

                        int sensor_type = int.Parse(mod_settings[global.index][1]);                                         
                        
                      
                        if (sensor_type >0 && sensor_type<=99 && !mod_alarm[global.index][13])
                        {
                            //first field alarms
                            if  (bool.Parse(gen_settings[sensor_type][3]))
                            {
                                if (gen_settings[sensor_type][5] != "B") // if field is not boolean
                                {

                                    //High first field alarm
                                    if (Convert.ToDouble(value2[global.index]) >= Convert.ToDouble(gen_settings[sensor_type][6]))
                                    {
                                        mod_alarm[global.index][3] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][3] = false;
                                    }
                                    if (mod_alarm[global.index][3] != mod_alarm_mem[global.index][3])
                                    {
                                        if (mod_alarm[global.index][3])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " High " + gen_settings[sensor_type][4];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][3] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " High " + gen_settings[sensor_type][4];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][3] = false;
                                        }
                                    }

                                    //Low first field alarm
                                    if (Convert.ToDouble(value2[global.index]) <= Convert.ToDouble(gen_settings[sensor_type][7]))
                                    {
                                        mod_alarm[global.index][4] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][4] = false;
                                    }
                                    if (mod_alarm[global.index][4] != mod_alarm_mem[global.index][4])
                                    {
                                        if (mod_alarm[global.index][4])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " Low " + gen_settings[sensor_type][4];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][4] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " Low " + gen_settings[sensor_type][4];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][4] = false;
                                        }
                                    }
                                }
                                else //if field is boolean
                                {
                                    if (bool.Parse(value2[global.index]))
                                    {
                                        mod_alarm[global.index][5] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][5] = false;
                                    }
                                    
                                    if (mod_alarm[global.index][5] != mod_alarm_mem[global.index][5])
                                    {
                                        if (mod_alarm[global.index][5])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " " + gen_settings[sensor_type][4] + " is TRUE";
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][5] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " " + gen_settings[sensor_type][4] + " is TRUE";
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][5] = false;
                                        }
                                    }

                                }

                            }

                            //second field alarms
                            if (bool.Parse(gen_settings[sensor_type][9]))
                            {
                                if (gen_settings[sensor_type][11] != "B") // if field is not boolean
                                {
                                    //High second field alarm
                                    if (Convert.ToDouble(value3[global.index]) >= Convert.ToDouble(gen_settings[sensor_type][12]))
                                    {
                                        mod_alarm[global.index][6] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][6] = false;
                                    }
                                    if (mod_alarm[global.index][6] != mod_alarm_mem[global.index][6])
                                    {
                                        if (mod_alarm[global.index][6])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " High " + gen_settings[sensor_type][10];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][6] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " High " + gen_settings[sensor_type][10];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][6] = false;
                                        }
                                    }

                                    //Low second field alarm
                                    if (Convert.ToDouble(value3[global.index]) <= Convert.ToDouble(gen_settings[sensor_type][13]))
                                    {
                                        mod_alarm[global.index][7] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][7] = false;
                                    }
                                    if (mod_alarm[global.index][7] != mod_alarm_mem[global.index][7])
                                    {
                                        if (mod_alarm[global.index][7])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " Low " + gen_settings[sensor_type][10];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][7] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " Low " + gen_settings[sensor_type][10];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][7] = false;
                                        }
                                    }

                                }
                                else //if field is boolean
                                {
                                    if (bool.Parse(value3[global.index]))
                                    {
                                        mod_alarm[global.index][8] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][8] = false;
                                    }
                                    if (mod_alarm[global.index][8] != mod_alarm_mem[global.index][8])
                                    {
                                        if (mod_alarm[global.index][8])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " " + gen_settings[sensor_type][10] + " is TRUE";
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][8] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " " + gen_settings[sensor_type][10] + " is TRUE";
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][8] = false;
                                        }
                                    }

                                }



                            }

                            //third field alarms
                            if (bool.Parse(gen_settings[sensor_type][15]))
                            {
                                if (gen_settings[sensor_type][17] != "B") // if field is not boolean
                                {
                                    //High third field alarm
                                    if (Convert.ToDouble(value4[global.index]) >= Convert.ToDouble(gen_settings[sensor_type][18]))
                                    {
                                        mod_alarm[global.index][9] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][9] = false;
                                    }
                                    if (mod_alarm[global.index][9] != mod_alarm_mem[global.index][9])
                                    {
                                        if (mod_alarm[global.index][9])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " High " + gen_settings[sensor_type][16];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][9] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " High " + gen_settings[sensor_type][16];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][9] = false;
                                        }
                                    }

                                    //Low third field alarm
                                    if (Convert.ToDouble(value4[global.index]) <= Convert.ToDouble(gen_settings[sensor_type][19]))
                                    {
                                        mod_alarm[global.index][10] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][10] = false;
                                    }
                                    if (mod_alarm[global.index][10] != mod_alarm_mem[global.index][10])
                                    {
                                        if (mod_alarm[global.index][10])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " Low " + gen_settings[sensor_type][16];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][10] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " Low " + gen_settings[sensor_type][16];
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][10] = false;
                                        }
                                    }

                                }
                                else //if field is boolean
                                {
                                    if (bool.Parse(value4[global.index]))
                                    {
                                        mod_alarm[global.index][11] = true;
                                    }
                                    else
                                    {
                                        mod_alarm[global.index][11] = false;
                                    }
                                    if (mod_alarm[global.index][11] != mod_alarm_mem[global.index][11])
                                    {
                                        if (mod_alarm[global.index][11])
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " " + gen_settings[sensor_type][16] + " is TRUE";
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.IndianRed;
                                            mod_alarm_mem[global.index][11] = true;
                                            new_alarm = true;
                                        }

                                        else
                                        {
                                            string alarmText = "Sensor #" + global.index.ToString() + " " + gen_settings[sensor_type][16] + " is TRUE";
                                            ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                                            listViewAlarm.Items.Insert(0, item);
                                            item.ForeColor = Color.Green;
                                            mod_alarm_mem[global.index][11] = false;
                                        }
                                    }

                                }



                            }



                        }
                       
                        //module cumulative alarm
                        mod_alarm[global.index][0] = false;

                        for (int i = 1; i <= 13; i++)
                        {
                            mod_alarm[global.index][0] = mod_alarm[global.index][0] | mod_alarm[global.index][i];
                        }


                        if (mod_alarm[global.index][0] != mod_alarm_mem[global.index][0])
                        {
                            markFlag2[global.index].Markers.Clear();                      
                            string doc_path2 = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                            Bitmap flag;
                            if (mod_alarm[global.index][0])
                            {
                                flag = (Bitmap)Image.FromFile(doc_path2 + "\\GreenDetect\\map flags\\red.png");
                                mod_alarm_mem[global.index][0] = true;
                            }

                            else
                            {
                                flag = (Bitmap)Image.FromFile(doc_path2 + "\\GreenDetect\\map flags\\green.png");
                                mod_alarm_mem[global.index][0] = false;
                            }

                            flagBuilding(flag, global.index.ToString());
                            double flag_lat = Convert.ToDouble(mod_settings[global.index][6]);
                            double flag_longt = Convert.ToDouble(mod_settings[global.index][7]);
                            GMap.NET.WindowsForms.GMapMarker mark2 =
                            new GMap.NET.WindowsForms.Markers.GMarkerGoogle(new GMap.NET.PointLatLng(flag_lat, flag_longt), flag);
                            markFlag2[global.index].Markers.Add(mark2);
                            mark2.Tag = global.index.ToString();
                            gMap2.Overlays.Add(markFlag2[global.index]);
                            gMap2.Zoom--;
                            gMap2.Zoom++;
                        

                        }

                        
                    }


                }
                else
                {
                    port.DiscardInBuffer();

                }

                bool cumul_alarm = false;

                for (int a = 1; a <= int.Parse(gen_settings[1][0]); a++)
                {
                    cumul_alarm = cumul_alarm | mod_alarm[a][0];
                }

                cumul_alarm = cumul_alarm | timeout_error;

                if (cumul_alarm)
                {
                    buttonAlarm.Visible = true;
                }
                else
                {
                    buttonAlarm.Visible = false;
                    new_alarm = false;
                }

            }

        }





        private void comboBoxTypeSet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            int k = int.Parse (comboBoxTypeSet.Text);
            
            if (k == 0)
            {
                k = 100;
            }
            if ((k > 0) && (k <= 100))
            {                  
                  
                groupBoxSet1.Visible = bool.Parse(gen_settings[k][3]);
                groupBoxSet2.Visible = bool.Parse(gen_settings[k][9]);
                groupBoxSet3.Visible = bool.Parse(gen_settings[k][15]);

                labelTipeDescription.Text = gen_settings[k][2];

                textBoxF1size.Text = gen_settings[k][4];
                textBoxF1unit.Text = gen_settings[k][5];
                textBoxF1high.Text = gen_settings[k][6];
                textBoxF1low.Text = gen_settings[k][7];
                textBoxF1maxScale.Text = gen_settings[k][8];

                textBoxF2size.Text = gen_settings[k][10];
                textBoxF2unit.Text = gen_settings[k][11];
                textBoxF2high.Text = gen_settings[k][12];
                textBoxF2low.Text = gen_settings[k][13];
                textBoxF2maxScale.Text = gen_settings[k][14];

                textBoxF3size.Text = gen_settings[k][16];
                textBoxF3unit.Text = gen_settings[k][17];
                textBoxF3high.Text = gen_settings[k][18];
                textBoxF3low.Text = gen_settings[k][19];
                textBoxF3maxScale.Text = gen_settings[k][20];
            }
        }

        private void comboBoxTypeMod_SelectedIndexChanged(object sender, EventArgs e)
        {
            int addr_sel_mod = int.Parse(comboBoxAddress.Text);
            mod_settings[addr_sel_mod][1] = comboBoxTypeMod.Text;     

            int type_sel_mod = int.Parse(mod_settings[addr_sel_mod][1]);

           

            int type_sel_gen = type_sel_mod;
            if (type_sel_mod == 0)
            {
                type_sel_gen = 100;
            }

            //settings fields update
            groupBoxF1mod.Visible = bool.Parse(gen_settings[type_sel_gen][3]);           
            groupBoxF2mod.Visible = bool.Parse(gen_settings[type_sel_gen][9]);            
            groupBoxF3mod.Visible = bool.Parse(gen_settings[type_sel_gen][15]);            

            textBoxSocOffset.Text = mod_settings[addr_sel_mod][2];

            textBoxF1sizeMod.Text = gen_settings[type_sel_gen][4];
            textBoxF1unitMod.Text = gen_settings[type_sel_gen][5];
            textBoxF1Offset.Text = mod_settings[addr_sel_mod][3];

            textBoxF2sizeMod.Text = gen_settings[type_sel_gen][10];
            textBoxF2unitMod.Text = gen_settings[type_sel_gen][11];
            textBoxF2Offset.Text = mod_settings[addr_sel_mod][4];

            textBoxF3sizeMod.Text = gen_settings[type_sel_gen][16];
            textBoxF3unitMod.Text = gen_settings[type_sel_gen][17];
            textBoxF3Offset.Text = mod_settings[addr_sel_mod][5];



            // banner properties update
            //Group
            global.sen_ban[addr_sel_mod].group2 = bool.Parse(gen_settings[type_sel_gen][3]);
            global.sen_ban[addr_sel_mod].group3 = bool.Parse(gen_settings[type_sel_gen][9]);
            global.sen_ban[addr_sel_mod].group4 = bool.Parse(gen_settings[type_sel_gen][15]);
            //Title
            global.sen_ban[addr_sel_mod].title2 = gen_settings[type_sel_gen][4];
            global.sen_ban[addr_sel_mod].title3 = gen_settings[type_sel_gen][10];
            global.sen_ban[addr_sel_mod].title4 = gen_settings[type_sel_gen][16];

            //Unit
            global.sen_ban[addr_sel_mod].unit2 = gen_settings[type_sel_gen][5];
            global.sen_ban[addr_sel_mod].unit3 = gen_settings[type_sel_gen][11];
            global.sen_ban[addr_sel_mod].unit4 = gen_settings[type_sel_gen][17];



            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                mod_modified = true;
            }

        }

        private void comboBoxAddress_SelectedIndexChanged(object sender, EventArgs e)
        {
            int addr_sel_mod = int.Parse(comboBoxAddress.Text);
            comboBoxTypeMod.Text = mod_settings[addr_sel_mod][1];    
            
            int type_sel_mod = int.Parse(mod_settings[addr_sel_mod][1]);

            int type_sel_gen = type_sel_mod;
            if (type_sel_mod == 0)
            {
                type_sel_gen = 100;
            }

            textBoxRaw1.Text = raw1[addr_sel_mod];
            textBoxRaw2.Text = raw2[addr_sel_mod];
            textBoxRaw3.Text = raw3[addr_sel_mod];
            textBoxRaw4.Text = raw4[addr_sel_mod];
            textBoxSocValue.Text = value1[addr_sel_mod];
            textBoxF1Value.Text = value2[addr_sel_mod];
            textBoxF2Value.Text = value3[addr_sel_mod];
            textBoxF3Value.Text = value4[addr_sel_mod];

            groupBoxF1mod.Visible = bool.Parse(gen_settings[type_sel_gen][3]);
            groupBoxF2mod.Visible = bool.Parse(gen_settings[type_sel_gen][9]);
            groupBoxF3mod.Visible = bool.Parse(gen_settings[type_sel_gen][15]);

            textBoxSocOffset.Text = mod_settings[addr_sel_mod][2];

            textBoxF1sizeMod.Text = gen_settings[type_sel_gen][4];
            textBoxF1unitMod.Text = gen_settings[type_sel_gen][5];
            textBoxF1Offset.Text = mod_settings[addr_sel_mod][3];

            textBoxF2sizeMod.Text = gen_settings[type_sel_gen][10];
            textBoxF2unitMod.Text = gen_settings[type_sel_gen][11];
            textBoxF2Offset.Text = mod_settings[addr_sel_mod][4];

            textBoxF3sizeMod.Text = gen_settings[type_sel_gen][16];
            textBoxF3unitMod.Text = gen_settings[type_sel_gen][17];
            textBoxF3Offset.Text = mod_settings[addr_sel_mod][5];

            textBoxLat.Text = mod_settings[addr_sel_mod][6];
            textBoxLongt.Text = mod_settings[addr_sel_mod][7];



            double lat = Convert.ToDouble(textBoxLat.Text);
            double longt = Convert.ToDouble(textBoxLongt.Text);
            
            
       

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            gen_settings[1][0] = comboBox2.Text;
            comboBoxAddress.Items.Clear();
            int h = int.Parse(gen_settings[1][0]);
            for (int y = 1; y <= h; y++)
            {
                comboBoxAddress.Items.Add(y);
            }
            comboBoxAddress.SelectedIndex = 0;

            global.k = int.Parse(comboBox2.Text);

            for (int i = 1; i <= global.k; i++)
            {
                global.sen_ban[i].Visible = true;
            }

            if (global.k < 60)
            {
                for (int i = (global.k + 1); i <= 60; i++)
                {
                    global.sen_ban[i].Visible = false;
                }
            }


            //Banners properties update            
            for (int i = 1; i <= int.Parse(gen_settings[1][0]); i++)
            {
                //group visualization
                int type_i = int.Parse(mod_settings[i][1]);
                if (type_i == 0)
                {
                    type_i = 100;
                }
                global.sen_ban[i].group2 = bool.Parse(gen_settings[type_i][3]);
                global.sen_ban[i].group3 = bool.Parse(gen_settings[type_i][9]);
                global.sen_ban[i].group4 = bool.Parse(gen_settings[type_i][15]);

                //Title
                global.sen_ban[i].title2 = gen_settings[type_i][4];
                global.sen_ban[i].title3 = gen_settings[type_i][10];
                global.sen_ban[i].title4 = gen_settings[type_i][16];

                //Unit
                global.sen_ban[i].unit2 = gen_settings[type_i][5];
                global.sen_ban[i].unit3 = gen_settings[type_i][11];
                global.sen_ban[i].unit4 = gen_settings[type_i][17];

            }

            //maps markers flag update

            refreshMarkers();

            for (int i = int.Parse(gen_settings[1][0]) + 1; i <= 60; i++)
            {
                               
                markFlag1[i].Markers.Clear();
                markFlag2[i].Markers.Clear();
                
            }


            if (!init) 
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF1high_TextChanged(object sender, EventArgs e)
        {
            
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][6] = textBoxF1high.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }

        }

        private void textBoxF1low_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][7] = textBoxF1low.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF1maxScale_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][8] = textBoxF1maxScale.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF2high_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][12] = textBoxF2high.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF2low_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][13] = textBoxF2low.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF2maxScale_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][14] = textBoxF2maxScale.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF3high_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][18] = textBoxF3high.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF3low_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][19] = textBoxF3low.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void textBoxF3maxScale_TextChanged(object sender, EventArgs e)
        {
            int type_sel_gen = comboBoxTypeSet.SelectedIndex;
            if (type_sel_gen == 0)
            {
                type_sel_gen = 100;
            }
            gen_settings[type_sel_gen][20] = textBoxF3maxScale.Text;
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                gen_modified = true;
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            
            if (gen_modified)
            {
                string doc_path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string strFilePath = doc_path + "\\GreenDetect\\general_settings.csv";
                string strSeperator = ";";
                StringBuilder sbOutput = new StringBuilder();         


                for (int i = 0; i <= 100; i++)
                {
                    sbOutput.AppendLine(string.Join(strSeperator, gen_settings [i]));
                }
                
                // Create and write the csv file
                File.WriteAllText(strFilePath, sbOutput.ToString());


                // To append more lines to the csv file
                //File.AppendAllText(strFilePath, sbOutput.ToString());
            }

            if (mod_modified)
            {
                string doc_path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string strFilePath = doc_path + "\\GreenDetect\\module_settings.csv";
                string strSeperator = ";";
                StringBuilder sbOutput = new StringBuilder();
                string[] first_line = new string[] { "address", "type", "offset1", "offset2", "offset3", "offset4", "lat", "longt" };

                sbOutput.AppendLine(string.Join(strSeperator, first_line));

                for (int i = 0; i <= 60; i++)
                {
                    sbOutput.AppendLine(string.Join(strSeperator, mod_settings[i]));
                }

                // Create and write the csv file
                File.WriteAllText(strFilePath, sbOutput.ToString());


                // To append more lines to the csv file
                //File.AppendAllText(strFilePath, sbOutput.ToString());
            }

            buttonSave.Enabled = false;
            buttonSave.BackColor = Color.WhiteSmoke;

        }

        private void textBoxSocOffset_TextChanged(object sender, EventArgs e)
        {
            int addr_sel_mod = int.Parse(comboBoxAddress.Text);            
            mod_settings[addr_sel_mod][2] = textBoxSocOffset.Text;  
            
            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                mod_modified = true;
            }
        }

        private void textBoxF1Offset_TextChanged(object sender, EventArgs e)
        {

            int addr_sel_mod = int.Parse(comboBoxAddress.Text);
            mod_settings[addr_sel_mod][3] = textBoxF1Offset.Text;

            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                mod_modified = true;
            }
        }

        private void textBoxF2Offset_TextChanged(object sender, EventArgs e)
        {
            int addr_sel_mod = int.Parse(comboBoxAddress.Text);
            mod_settings[addr_sel_mod][4] = textBoxF2Offset.Text;

            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                mod_modified = true;
            }
        }

        private void textBoxF3Offset_TextChanged(object sender, EventArgs e)
        {
            int addr_sel_mod = int.Parse(comboBoxAddress.Text);
            mod_settings[addr_sel_mod][5] = textBoxF3Offset.Text;

            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                mod_modified = true;
            }
        }

       
        private void timer2_Tick_1(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                getAvailableComPorts();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
           
        }

       

        private void gMap1_MouseMove_1(object sender, MouseEventArgs e)
        {
            base.OnMouseMove(e);

            double X = gMap1.FromLocalToLatLng(e.X, e.Y).Lng;
            double Y = gMap1.FromLocalToLatLng(e.X, e.Y).Lat;

            double x_tr = Math.Round(X, 7);
            double y_tr = Math.Round(Y, 7); 


            string longitude = x_tr.ToString();
            string latitude = y_tr.ToString();
            labelLongt1.Text = longitude;
            LabelLat1.Text = latitude;
        }

        private void gMap1_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            
           // gMap1.Overlays.Remove(markFlag1);
            markFlag1[1].Markers.Clear();
            gMap1.Zoom--;
            gMap1.Zoom++;
            



        }

        private void button4_Click(object sender, EventArgs e)
        {
            markFlag1[2].Markers.Clear();
            gMap1.Zoom--;
            gMap1.Zoom++;
        }

        private void textBoxLat_TextChanged(object sender, EventArgs e)
        {
            int addr_sel_mod = int.Parse(comboBoxAddress.Text);
            mod_settings[addr_sel_mod][6] = textBoxLat.Text;
            refreshMarkers();

            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                mod_modified = true;
            }
        }

        private void textBoxLongt_TextChanged(object sender, EventArgs e)
        {
            int addr_sel_mod = int.Parse(comboBoxAddress.Text);
            mod_settings[addr_sel_mod][7] = textBoxLongt.Text;
            refreshMarkers();

            if (!init)
            {
                buttonSave.Enabled = true;
                buttonSave.BackColor = Color.Yellow;
                mod_modified = true;
            }
        }

        private void gMap1_DoubleClick(object sender, EventArgs e)
        {
            textBoxLat.Text = LabelLat1.Text;
            textBoxLongt.Text = labelLongt1.Text;   
        }

       

        private void button9_Click(object sender, EventArgs e)
        {
            gMap1.Bearing -= 5;
            gen_settings[5][0] = gMap1.Bearing.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            gMap1.Bearing += 5;
            gen_settings[5][0] = gMap1.Bearing.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            gMap1.Zoom++;
            gen_settings[6][0] = gMap1.Zoom.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            gMap1.Zoom--;
            gen_settings[6][0] = gMap1.Zoom.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            gMap2.Bearing -= 5;
            gen_settings[9][0] = gMap2.Bearing.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;

        }

        private void button11_Click(object sender, EventArgs e)
        {
            gMap2.Bearing += 5;
            gen_settings[9][0] = gMap2.Bearing.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            gMap2.Zoom++;
            gen_settings[10][0] = gMap2.Zoom.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            gMap2.Zoom--;
            gen_settings[10][0] = gMap2.Zoom.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

       

        private void gMap1_LocationChanged(object sender, EventArgs e)
        {
            
        }

        private void gMap1_OnMapDrag()
        {
            double map1Lat = gMap1.Position.Lat;
            double map1Lat_tr = Math.Round(map1Lat, 7);
            double map1Longt = gMap1.Position.Lng;
            double map1Longt_tr = Math.Round(map1Longt, 7);
            gen_settings[3][0] = map1Lat_tr.ToString();
            gen_settings[4][0] = map1Longt_tr.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
            
        }

        private void gMap2_OnMapDrag()
        {
            double map2Lat = gMap2.Position.Lat;
            double map2Lat_tr = Math.Round(map2Lat, 7);
            double map2Longt = gMap2.Position.Lng;
            double map2Longt_tr = Math.Round(map2Longt, 7);
            gen_settings[7][0] = map2Lat_tr.ToString();
            gen_settings[8][0] = map2Longt_tr.ToString();
            buttonSave.Enabled = true;
            buttonSave.BackColor = Color.Yellow;
            gen_modified = true;
        }

        private void gMap2_OnMarkerDoubleClick(GMapMarker item, MouseEventArgs e)
        {
            sensorBannerMap.Visible = true;
            global.mapBannerAddr = item.Tag.ToString();
            int mapBannerAddrInt = int.Parse(global.mapBannerAddr);
            if (mapBannerAddrInt < 10)
            {
                sensorBannerMap.address = "#0" + global.mapBannerAddr;
            }
            else
            {
                sensorBannerMap.address = "#" + global.mapBannerAddr;
            }
            
            int type_i = int.Parse(mod_settings[mapBannerAddrInt][1]);
            if (type_i == 0)
            {
                type_i = 100;
            }
            

            sensorBannerMap.group2 = bool.Parse(gen_settings[type_i][3]);
            sensorBannerMap.group3 = bool.Parse(gen_settings[type_i][9]);
            sensorBannerMap.group4 = bool.Parse(gen_settings[type_i][15]);
            sensorBannerMap.title2 = gen_settings[type_i][4];
            sensorBannerMap.title3 = gen_settings[type_i][10];
            sensorBannerMap.title4 = gen_settings[type_i][16];
            sensorBannerMap.unit2 = gen_settings[type_i][5];
            sensorBannerMap.unit3 = gen_settings[type_i][11];
            sensorBannerMap.unit4 = gen_settings[type_i][17];

            if (!isConnected)
            {
                sensorBannerMap.data1 = "####";
                sensorBannerMap.data2 = "####";
                sensorBannerMap.data3 = "####";
                sensorBannerMap.data4 = "####";
            }
            else
            {
                sensorBannerMap.data1 = value1[mapBannerAddrInt];
                sensorBannerMap.data2 = value2[mapBannerAddrInt];
                sensorBannerMap.data3 = value3[mapBannerAddrInt];
                sensorBannerMap.data4 = value4[mapBannerAddrInt];
            }

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if (isConnected && (read_count>0))
            {
                if (label50.Text == "Timeout")
                {
                    label50.Text = "Next";
                    label50.ForeColor = Color.Yellow;
                    labelReadTime.ForeColor = Color.Yellow;
                    timeout_count = 0;
                }
                read_count--;
                labelReadTime.Text = read_count.ToString();
                
            }

            if (isConnected && (read_count == 0))
            {
                if (label50.Text == "Next") 
                {
                    label50.Text = "Timeout";
                    label50.ForeColor = Color.Pink;
                    labelReadTime.ForeColor = Color.Pink;
                    timeout_count = 0;
                }
                timeout_count++;               
                labelReadTime.Text = timeout_count.ToString();

                if (timeout_count > 130)
                {
                    timeout_error = true;
                }
                else
                {
                    timeout_error = false;
                }

                if ( timeout_error != timeout_error_mem)
                {
                    if (timeout_error)
                    {
                        string alarmText = "Last sensor or Gateway Communication error";
                        ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "ON" });
                        listViewAlarm.Items.Insert(0, item);
                        item.ForeColor = Color.IndianRed;
                        refreshMarkers();
                        timeout_error_mem = true;
                    }

                    else
                    {
                        string alarmText = "Last sensor or Gateway Communication error";
                        ListViewItem item = new ListViewItem(new[] { DateTime.Now.ToString(), alarmText, "OFF" });
                        listViewAlarm.Items.Insert(0, item);
                        item.ForeColor = Color.Green;
                        timeout_error_mem = false;
                    }                
                    

                }


            }


            if (isConnected)

            {
                if (new_alarm)

                {
                    System.Media.SystemSounds.Hand.Play();

                    if (buttonAlarm.BackColor == Color.Red)
                    {
                        buttonAlarm.BackColor = Color.WhiteSmoke;
                    }
                   else
                    {
                        buttonAlarm.BackColor = Color.Red;
                    }

                }

                else
                {
                    buttonAlarm.BackColor = Color.Red;
                }


            }


        }

        private void buttonAlarm_Click(object sender, EventArgs e)
        {
            new_alarm = false;            
        }
   

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (listViewAlarm.Items.Count > 0)
            {
                string doc_path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string strFilePath = doc_path + "\\GreenDetect\\alarms.csv";
                string strSeperator = ";";
                StringBuilder sbOutput = new StringBuilder();
                string[] line = new string[] { "Date and time", "Alarm", "Status" };
                sbOutput.AppendLine(string.Join(strSeperator, line));

                foreach (ListViewItem item in listViewAlarm.Items)

                {
                    line = new string[] { item.SubItems[0].Text, item.SubItems[1].Text, item.SubItems[2].Text };
                    sbOutput.AppendLine(string.Join(strSeperator, line));
                }

                // Create and write the csv file
                File.WriteAllText(strFilePath, sbOutput.ToString());
            }
        }

       
    }



    public class global
    {
       public static int ciao;
        //public static ArrayList sen_ban = new ArrayList();
        public static List<GreenDetect.SensorBanner> sen_ban = new List<GreenDetect.SensorBanner>();
        public static int k;
        public static bool read_enable = true;
        public static string serial_data;
        public static int index;
        public static string mapBannerAddr= "1";


    }

   


}
/*string pippo = port.ReadLine();
listBox1.Items.Add(pippo);

 global.sen_ban[15].data2 = global.ciao.ToString();
     global.sen_ban[45].title3 = global.ciao.ToString();
     global.sen_ban[30].title2 = textBox3.Text;
     global.sen_ban[59].group4 = !global.sen_ban[59].group4;*/

/*int pluto = int.Parse (pippo);
   int topolino = 1500 + (pluto * 25);
   string paperoga = topolino.ToString();*/
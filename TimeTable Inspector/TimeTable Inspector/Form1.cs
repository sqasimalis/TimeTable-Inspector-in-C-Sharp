/*
 *  Author  :   Syed Qasim Ali Shah
 *  Email   :   syedqasimali311@gmail.com
*/

using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;   // Use to call Marshal.ReleaseComObject(xlApp) etc.

//  Added a Reference "Microsoft.Office.Interop.Excel" to work with MS Excel Files

namespace TimeTable_Inspector
{
    public partial class Form1 : Form
    {
        //  Structure to save the Slots of Time Table
        public struct slot
        {
            public string day;
            public string time;
            public string batch;
            public string course;
            public string instructor;
            public string room;
        }

        const int noOfSlots = 120;                          //  Constant Variable used to define the Length of array of TimeTable's Slots
        slot[] TTSlot = new slot[noOfSlots];                //  Objects of Structure to save TimeTable's Slots
        static int index = 0;                               //  Variable used to Operate Loops of TimeTable's Slots
        List<string> instructorList = new List<string>();   //  List to save the repeating names of Instructors one time
        List<string> batchList = new List<string>();        //  List to save the Batch IDs
        List<string> roomList = new List<string>();         //  List to save the repeating names of Rooms one time

        //  Lists to Share Data with Other Form
        List<string> dl = new List<string>();
        List<string> tl = new List<string>();
        List<string> bl = new List<string>();
        List<string> cl = new List<string>();
        List<string> il = new List<string>();
        List<string> rl = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
            Microsoft.Office.Interop.Excel.Range range;

            var folderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            xlWorkbook = xlApp.Workbooks.Open(folderPath + @"\TimeTable.xlsx");

            //  Code to read data from MS Excel file of TimeTable
            for (int sheet = 1; sheet <= 9; sheet++)
            {
                xlWorksheet = xlApp.Worksheets.Item[sheet];
                range = xlWorksheet.Cells;

                for (int x = 3; x <= 15; x = x + 3)     //  Dealing Rows
                {
                    for (int y = 3; y <= 8; y++)    //  Dealing Columns
                    {
                        if (range.Item[x, y].value != null)
                        {
                            TTSlot[index].day = range.Item[x, 1].value;
                            TTSlot[index].time = range.Item[1, y].value;
                            TTSlot[index].batch = range.Item[1, 2].value;
                            TTSlot[index].course = range.Item[x - 1, y].value;
                            TTSlot[index].instructor = range.Item[x, y].value;
                            TTSlot[index].room = range.Item[x + 1, y].value;
                            index++;
                        }
                    }
                }

                batchList.Add(range.Item[1, 2].value);  // Code to Add Batch IDs in a list
            }

            xlWorkbook.Save();
            xlWorkbook.Close();
            xlApp.Quit();

            //Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);

            //  Code to Find the number of total Instructors
            int instructorLoopController = 0;
            while (TTSlot[instructorLoopController].instructor != null)
            {
                instructorLoopController++;
            }

            //  Code to Add repeating Names of Instructors one time in a list
            for (int i = 0; i < instructorLoopController - 1; i++)
            {
                if (instructorList.Contains(TTSlot[i].instructor) == false)
                {
                    instructorList.Add(TTSlot[i].instructor);
                }
            }

            //  Code to Add Instructors Names in Combo Box
            instructorList.Sort();    //  Code to sort the list in aescending order
            for (int i = 0; i < instructorList.Count; i++)
            {
                comboBox1.Items.Add(instructorList[i]);
            }

            //  Code to Add Batch IDs in Combo Box
            batchList.Sort();    //  Code to sort the list in aescending order
            for (int i = 0; i < batchList.Count; i++)
            {
                comboBox2.Items.Add(batchList[i]);
            }

            //  Code to Find the number of Total Rooms
            int roomLoopController = 0;
            while (TTSlot[roomLoopController].room != null)
            {
                roomLoopController++;
            }

            //  Code to Add repeating names of Rooms one time in a list
            for (int i = 0; i < roomLoopController - 1; i++)
            {
                if (roomList.Contains(TTSlot[i].room) == false)
                {
                    roomList.Add(TTSlot[i].room);
                }
            }

            //  Code to Add Rooms names in Combo Box
            roomList.Sort();    //  Code to sort the list in aescending order
            for (int i = 0; i < roomList.Count; i++)
            {
                comboBox3.Items.Add(roomList[i]);
            }

            //  Adding Data to Lists (for other Form)
            for (int i = 0; i < noOfSlots; i++)
            {
                dl.Add(TTSlot[i].day);
                tl.Add(TTSlot[i].time);
                bl.Add(TTSlot[i].batch);
                cl.Add(TTSlot[i].course);
                il.Add(TTSlot[i].instructor);
                rl.Add(TTSlot[i].room);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            comboBox2.Text = null;  //  Code to Clear Text of other Combo Box
            comboBox3.Text = null;  //  Code to Clear Text of other Combo Box
            
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Please select the Instructor Name (from given options).");
            }
            else
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
                Microsoft.Office.Interop.Excel.Range range;

                var folderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
                xlWorkbook = xlApp.Workbooks.Open(folderPath + @"\Instructor TT.xlsx");
                xlWorksheet = xlApp.Worksheets.Item[1];
                range = xlWorksheet.Cells;

                //  Code to Clear Excel Sheet
                for (int a = 5; a <= 19; a++)
                {
                    for (int b = 3; b <= 8; b++)
                    {
                        range.Item[a, b].value = null;
                    }
                }

                int x = 0;
                int y = 0;

                //  Code to Write Time Table on Excel Sheet
                for (int localIndex = 0; localIndex < noOfSlots; localIndex++)
                {
                    if (comboBox1.SelectedItem.ToString() == TTSlot[localIndex].instructor)
                    {
                        if (TTSlot[localIndex].day == "Monday")
                            x = 6;
                        else if (TTSlot[localIndex].day == "Tuesday")
                            x = 9;
                        else if (TTSlot[localIndex].day == "Wednesday")
                            x = 12;
                        else if (TTSlot[localIndex].day == "Thursday")
                            x = 15;
                        else if (TTSlot[localIndex].day == "Friday")
                            x = 18;

                        if (TTSlot[localIndex].time == "08:30 - 10:00")
                            y = 3;
                        else if (TTSlot[localIndex].time == "10:00 - 11:30")
                            y = 4;
                        else if (TTSlot[localIndex].time == "11:30 - 13:00")
                            y = 5;
                        else if (TTSlot[localIndex].time == "13:30 - 15:00")
                            y = 6;
                        else if (TTSlot[localIndex].time == "15:00 - 16:30")
                            y = 7;
                        else if (TTSlot[localIndex].time == "16:30 - 18:00")
                            y = 8;

                        range.Item[2, 5].value = TTSlot[localIndex].instructor;
                        range.Item[x - 1, y].value = TTSlot[localIndex].batch;
                        range.Item[x, y].value = TTSlot[localIndex].course;
                        range.Item[x + 1, y].value = TTSlot[localIndex].room;
                    }
                }

                showTimeTable(range);   // Code to Show Time Table

                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.Text = null;  //  Code to Clear Text of other Combo Box
            comboBox3.Text = null;  //  Code to Clear Text of other Combo Box
            
            if (comboBox2.SelectedItem == null)
            {
                MessageBox.Show("Please select the Batch ID (from given options).");
            }
            else
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
                Microsoft.Office.Interop.Excel.Range range;

                var folderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
                xlWorkbook = xlApp.Workbooks.Open(folderPath + @"\Batch TT.xlsx");
                xlWorksheet = xlApp.Worksheets.Item[1];
                range = xlWorksheet.Cells;

                //  Code to Clear Excel Sheet
                for (int a = 5; a <= 19; a++)
                {
                    for (int b = 3; b <= 8; b++)
                    {
                        range.Item[a, b].value = null;
                    }
                }

                int x = 0;
                int y = 0;

                //  Code to Write Time Table on Excel Sheet
                for (int localIndex = 0; localIndex < noOfSlots; localIndex++)
                {
                    if (comboBox2.SelectedItem.ToString() == TTSlot[localIndex].batch)
                    {
                        if (TTSlot[localIndex].day == "Monday")
                            x = 6;
                        else if (TTSlot[localIndex].day == "Tuesday")
                            x = 9;
                        else if (TTSlot[localIndex].day == "Wednesday")
                            x = 12;
                        else if (TTSlot[localIndex].day == "Thursday")
                            x = 15;
                        else if (TTSlot[localIndex].day == "Friday")
                            x = 18;

                        if (TTSlot[localIndex].time == "08:30 - 10:00")
                            y = 3;
                        else if (TTSlot[localIndex].time == "10:00 - 11:30")
                            y = 4;
                        else if (TTSlot[localIndex].time == "11:30 - 13:00")
                            y = 5;
                        else if (TTSlot[localIndex].time == "13:30 - 15:00")
                            y = 6;
                        else if (TTSlot[localIndex].time == "15:00 - 16:30")
                            y = 7;
                        else if (TTSlot[localIndex].time == "16:30 - 18:00")
                            y = 8;

                        range.Item[2, 5].value = TTSlot[localIndex].batch;
                        range.Item[x - 1, y].value = TTSlot[localIndex].course;
                        range.Item[x, y].value = TTSlot[localIndex].instructor;
                        range.Item[x + 1, y].value = TTSlot[localIndex].room;
                    }
                }

                showTimeTable(range);   // Code to Show Time Table

                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboBox1.Text = null;  //  Code to Clear Text of other Combo Box
            comboBox2.Text = null;  //  Code to Clear Text of other Combo Box
            
            if (comboBox3.SelectedItem == null)
            {
                MessageBox.Show("Please select the Room ID (from given options).");
            }
            else
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
                Microsoft.Office.Interop.Excel.Range range;

                var folderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
                xlWorkbook = xlApp.Workbooks.Open(folderPath + @"\Room TT.xlsx");
                xlWorksheet = xlApp.Worksheets.Item[1];
                range = xlWorksheet.Cells;

                //  Code to Clear Excel Sheet
                for (int a = 5; a <= 19; a++)
                {
                    for (int b = 3; b <= 8; b++)
                    {
                        range.Item[a, b].value = null;
                    }
                }

                int x = 0;
                int y = 0;

                //  Code to Write Time Table on Excel Sheet
                for (int localIndex = 0; localIndex < noOfSlots; localIndex++)
                {
                    if (comboBox3.SelectedItem.ToString() == TTSlot[localIndex].room)
                    {
                        if (TTSlot[localIndex].day == "Monday")
                            x = 6;
                        else if (TTSlot[localIndex].day == "Tuesday")
                            x = 9;
                        else if (TTSlot[localIndex].day == "Wednesday")
                            x = 12;
                        else if (TTSlot[localIndex].day == "Thursday")
                            x = 15;
                        else if (TTSlot[localIndex].day == "Friday")
                            x = 18;

                        if (TTSlot[localIndex].time == "08:30 - 10:00")
                            y = 3;
                        else if (TTSlot[localIndex].time == "10:00 - 11:30")
                            y = 4;
                        else if (TTSlot[localIndex].time == "11:30 - 13:00")
                            y = 5;
                        else if (TTSlot[localIndex].time == "13:30 - 15:00")
                            y = 6;
                        else if (TTSlot[localIndex].time == "15:00 - 16:30")
                            y = 7;
                        else if (TTSlot[localIndex].time == "16:30 - 18:00")
                            y = 8;

                        range.Item[2, 5].value = TTSlot[localIndex].room;
                        range.Item[x - 1, y].value = TTSlot[localIndex].batch;
                        range.Item[x, y].value = TTSlot[localIndex].course;
                        range.Item[x + 1, y].value = TTSlot[localIndex].instructor;
                    }
                }

                showTimeTable(range);   // Code to Show Time Table

                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        //  Code to Show Time Table
        private void showTimeTable(Microsoft.Office.Interop.Excel.Range range)
        {
            label6.Text = range.Item[2, 5].value;   //  Show Title (Time Table of)
            label19.Text = label22.Text = label25.Text = label28.Text = label31.Text = range.Item[5, 2].value;  // Show Title (Batch)
            label20.Text = label23.Text = label26.Text = label29.Text = label32.Text = range.Item[6, 2].value;  // Show Title (Course)
            label21.Text = label24.Text = label27.Text = label30.Text = label33.Text = range.Item[7, 2].value;  // Show Title (Room)
            label34.Text = range.Item[5, 3].value;  // Show Column of (08:30 - 10:00)
            label35.Text = range.Item[6, 3].value;
            label36.Text = range.Item[7, 3].value;
            label37.Text = range.Item[8, 3].value;
            label38.Text = range.Item[9, 3].value;
            label39.Text = range.Item[10, 3].value;
            label40.Text = range.Item[11, 3].value;
            label41.Text = range.Item[12, 3].value;
            label42.Text = range.Item[13, 3].value;
            label43.Text = range.Item[14, 3].value;
            label44.Text = range.Item[15, 3].value;
            label45.Text = range.Item[16, 3].value;
            label46.Text = range.Item[17, 3].value;
            label47.Text = range.Item[18, 3].value;
            label48.Text = range.Item[19, 3].value;
            label49.Text = range.Item[5, 4].value;  // Show Column of (10:00 - 11:30)
            label50.Text = range.Item[6, 4].value;
            label51.Text = range.Item[7, 4].value;
            label52.Text = range.Item[8, 4].value;
            label53.Text = range.Item[9, 4].value;
            label54.Text = range.Item[10, 4].value;
            label55.Text = range.Item[11, 4].value;
            label56.Text = range.Item[12, 4].value;
            label57.Text = range.Item[13, 4].value;
            label58.Text = range.Item[14, 4].value;
            label59.Text = range.Item[15, 4].value;
            label60.Text = range.Item[16, 4].value;
            label61.Text = range.Item[17, 4].value;
            label62.Text = range.Item[18, 4].value;
            label63.Text = range.Item[19, 4].value;
            label64.Text = range.Item[5, 5].value;  // Show Column of (11:30 - 13:00)
            label65.Text = range.Item[6, 5].value;
            label66.Text = range.Item[7, 5].value;
            label67.Text = range.Item[8, 5].value;
            label68.Text = range.Item[9, 5].value;
            label69.Text = range.Item[10, 5].value;
            label70.Text = range.Item[11, 5].value;
            label71.Text = range.Item[12, 5].value;
            label72.Text = range.Item[13, 5].value;
            label73.Text = range.Item[14, 5].value;
            label74.Text = range.Item[15, 5].value;
            label75.Text = range.Item[16, 5].value;
            label76.Text = range.Item[17, 5].value;
            label77.Text = range.Item[18, 5].value;
            label78.Text = range.Item[19, 5].value;
            label79.Text = range.Item[5, 6].value;  // Show Column of (13:30 - 15:00)
            label80.Text = range.Item[6, 6].value;
            label81.Text = range.Item[7, 6].value;
            label82.Text = range.Item[8, 6].value;
            label83.Text = range.Item[9, 6].value;
            label84.Text = range.Item[10, 6].value;
            label85.Text = range.Item[11, 6].value;
            label86.Text = range.Item[12, 6].value;
            label87.Text = range.Item[13, 6].value;
            label88.Text = range.Item[14, 6].value;
            label89.Text = range.Item[15, 6].value;
            label90.Text = range.Item[16, 6].value;
            label91.Text = range.Item[17, 6].value;
            label92.Text = range.Item[18, 6].value;
            label93.Text = range.Item[19, 6].value;
            label94.Text = range.Item[5, 7].value;  // Show Column of (15:00 - 16:30)
            label95.Text = range.Item[6, 7].value;
            label96.Text = range.Item[7, 7].value;
            label97.Text = range.Item[8, 7].value;
            label98.Text = range.Item[9, 7].value;
            label99.Text = range.Item[10, 7].value;
            label100.Text = range.Item[11, 7].value;
            label101.Text = range.Item[12, 7].value;
            label102.Text = range.Item[13, 7].value;
            label103.Text = range.Item[14, 7].value;
            label104.Text = range.Item[15, 7].value;
            label105.Text = range.Item[16, 7].value;
            label106.Text = range.Item[17, 7].value;
            label107.Text = range.Item[18, 7].value;
            label108.Text = range.Item[19, 7].value;
            label109.Text = range.Item[5, 8].value;  // Show Column of (16:30 - 18:00)
            label110.Text = range.Item[6, 8].value;
            label111.Text = range.Item[7, 8].value;
            label112.Text = range.Item[8, 8].value;
            label113.Text = range.Item[9, 8].value;
            label114.Text = range.Item[10, 8].value;
            label115.Text = range.Item[11, 8].value;
            label116.Text = range.Item[12, 8].value;
            label117.Text = range.Item[13, 8].value;
            label118.Text = range.Item[14, 8].value;
            label119.Text = range.Item[15, 8].value;
            label120.Text = range.Item[16, 8].value;
            label121.Text = range.Item[17, 8].value;
            label122.Text = range.Item[18, 8].value;
            label123.Text = range.Item[19, 8].value;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //  Code to pass lists to other Form
            using (Form2 f = new Form2())
            {
                f.instructorList.AddRange(instructorList);
                f.batchList.AddRange(batchList);
                f.roomList.AddRange(roomList);

                f.dl.AddRange(dl);
                f.tl.AddRange(tl);
                f.bl.AddRange(bl);
                f.cl.AddRange(cl);
                f.il.AddRange(il);
                f.rl.AddRange(rl);

                f.ShowDialog();
            }
        }
    }
}
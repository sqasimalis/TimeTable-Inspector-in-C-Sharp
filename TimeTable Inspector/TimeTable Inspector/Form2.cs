/*
 *  Author  :   Syed Qasim Ali Shah
 *  Email   :   syedqasimali311@gmail.com
*/

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TimeTable_Inspector
{
    public partial class Form2 : Form
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

        //  Structure to save the Slots by Room
        public struct roomSlot
        {
            public string time;
            public string room;
            public bool availability;
        }
        
        const int noOfSlots = 120;                                  //  Constant Variable used to define the Length of array of TimeTable's Slots
        public slot[] TTSlot = new slot[noOfSlots];                 //  Objects of Structure to save TimeTable's Slots
        public List<string> instructorList = new List<string>();    //  List to save the repeating names of Instructors one time
        public List<string> batchList = new List<string>();         //  List to save the Batch IDs
        public List<string> roomList = new List<string>();          //  List to save the repeating names of Rooms one time
        List<string> dayList = new List<string>();                  //  List to save Days
        roomSlot[] rSlot = new roomSlot[60];
        string[] row = new string[2];                               //  String to show data on DataGridView

        //  Lists to Save Data from Other Form
        public List<string> dl = new List<string>();
        public List<string> tl = new List<string>();
        public List<string> bl = new List<string>();
        public List<string> cl = new List<string>();
        public List<string> il = new List<string>();
        public List<string> rl = new List<string>();
        
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
           //  Code to Add Instructors Names in Combo Box
            for (int i = 0; i < instructorList.Count; i++)
            {
                comboBox1.Items.Add(instructorList[i]);
            }

            //  Code to Add Batch IDs in Combo Box
            for (int i = 0; i < batchList.Count; i++)
            {
                comboBox2.Items.Add(batchList[i]);
            }
            
            //  Code to Save Days in List
            dayList.Add("Monday");
            dayList.Add("Tuesday");
            dayList.Add("Wednesday");
            dayList.Add("Thursday");
            dayList.Add("Friday");

            //  Code to Add Days in Combo Box
            for (int i = 0; i < dayList.Count; i++)
            {
                comboBox3.Items.Add(dayList[i]);
            }

            //  Code to Initialize Array of Room Slots
            for (int i = 0; i < 60; i++)
            {
                if (i < 6)
                    rSlot[i].room = roomList[0];
                else if (i < 12)
                    rSlot[i].room = roomList[1];
                else if (i < 18)
                    rSlot[i].room = roomList[2];
                else if (i < 24)
                    rSlot[i].room = roomList[3];
                else if (i < 30)
                    rSlot[i].room = roomList[4];
                else if (i < 36)
                    rSlot[i].room = roomList[5];
                else if (i < 42)
                    rSlot[i].room = roomList[6];
                else if (i < 48)
                    rSlot[i].room = roomList[7];
                else if (i < 54)
                    rSlot[i].room = roomList[8];
                else if (i < 60)
                    rSlot[i].room = roomList[9];

                if (i == 0 || i == 6 || i == 12 || i == 18 || i == 24 || i == 30 || i == 36 || i == 42 || i == 48 || i == 54)
                    rSlot[i].time = "08:30 - 10:00";
                else if (i == 1 || i == 7 || i == 13 || i == 19 || i == 25 || i == 31 || i == 37 || i == 43 || i == 49 || i == 55)
                    rSlot[i].time = "10:00 - 11:30";
                else if (i == 2 || i == 8 || i == 14 || i == 20 || i == 26 || i == 32 || i == 38 || i == 44 || i == 50 || i == 56)
                    rSlot[i].time = "11:30 - 13:00";
                else if (i == 3 || i == 9 || i == 15 || i == 21 || i == 27 || i == 33 || i == 39 || i == 45 || i == 51 || i == 57)
                    rSlot[i].time = "13:30 - 15:00";
                else if (i == 4 || i == 10 || i == 16 || i == 22 || i == 28 || i == 34 || i == 40 || i == 46 || i == 52 || i == 58)
                    rSlot[i].time = "15:00 - 16:30";
                else if (i == 5 || i == 11 || i == 17 || i == 23 || i == 29 || i == 35 || i == 41 || i == 47 || i == 53 || i == 59)
                    rSlot[i].time = "16:30 - 18:00";

                rSlot[i].availability = true;
            }

            //  Adding data to Array from other Form
            for (int i = 0; i < noOfSlots; i++)
            {
                TTSlot[i].day = dl[i];
                TTSlot[i].time = tl[i];
                TTSlot[i].batch = bl[i];
                TTSlot[i].course = cl[i];
                TTSlot[i].instructor = il[i];
                TTSlot[i].room = rl[i];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear(); //  Code to clear Data Grid View if it is already filled

            //  Code to make all slots available on next Button Click
            for (int i = 0; i < 60; i++)
            {
                rSlot[i].availability = true;
            }

            if (comboBox1.SelectedItem == null || comboBox2.SelectedItem == null || comboBox3.SelectedItem == null)
            {
                MessageBox.Show("Please select the appropriate value (from given options) for all available fields.");
            }
            else
            {
                //  Code to overwrite the availablity of rooms
                for (int index = 0; index < noOfSlots; index++)
                {
                    if (TTSlot[index].day == comboBox3.SelectedItem.ToString())
                    {
                        for (int a = 0; a < 60; a++)
                        {
                            if (rSlot[a].time == TTSlot[index].time && rSlot[a].room == TTSlot[index].room)
                                rSlot[a].availability = false;
                            if (TTSlot[index].instructor == comboBox1.SelectedItem.ToString())
                            {
                                if (rSlot[a].time == TTSlot[index].time)
                                    rSlot[a].availability = false;
                            }
                            if (TTSlot[index].batch == comboBox2.SelectedItem.ToString())
                            {
                                if (rSlot[a].time == TTSlot[index].time)
                                    rSlot[a].availability = false;
                            }
                        }
                    }
                }

                //  Code to show those rooms with the time on which those are free
                for (int index = 0; index < 60; index++)
                {
                    if (rSlot[index].availability == true)
                    {
                        row = new string[] { rSlot[index].time, rSlot[index].room };
                        dataGridView1.Rows.Add(row);
                    }
                }
            }
        }
    }
}
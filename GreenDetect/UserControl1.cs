using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace GreenDetect
{
    public partial class SensorBanner : UserControl
    {
     

        public string address
        {
            get { return _address.Text; }
            set { _address.Text = value; }
        }


        public string data1
        {
            get { return _data1.Text; }
            set { _data1.Text = value; }
        }


        //group2 objects
        public bool group2
        {
            get { return _group2.Visible; }
            set { _group2.Visible = value; }
        }

        public string title2
        {
            get { return _group2.Text; }
            set { _group2.Text = value; }
        }

        public string data2
        {
            get { return _data2.Text; }
            set { _data2.Text = value; }
        }

        public string unit2
        {
            get { return _unit2.Text; }
            set { _unit2.Text = value; }
        }

        //group3 objects
        public bool group3
        {
            get { return _group3.Visible; }
            set { _group3.Visible = value; }
        }

        public string title3
        {
            get { return _group3.Text; }
            set { _group3.Text = value; }
        }

        public string data3
        {
            get { return _data3.Text; }
            set { _data3.Text = value; }
        }

        public string unit3
        {
            get { return _unit3.Text; }
            set { _unit3.Text = value; }
        }

        //group4 objects
        public bool group4
        {
            get { return _group4.Visible; }
            set { _group4.Visible = value; }
        }

        public string title4
        {
            get { return _group4.Text; }
            set { _group4.Text = value; }
        }

        public string data4
        {
            get { return _data4.Text; }
            set { _data4.Text = value; }
        }

        public string unit4
        {
            get { return _unit4.Text; }
            set { _unit4.Text = value; }
        }



        public SensorBanner()
        {
            InitializeComponent();
        }

        
    }
}

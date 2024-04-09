using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DAL;

using LSSERVICEPROVIDERLib;



namespace TestForm
{
    public partial class Form1 : Form
    {
        private List<string> persons1;
        private List<string> persons2;
        private List<string> persons3;

        public Form1()
        {
            InitializeComponent();
            //     radGridView1.Columns[4].DataEditFormatString="{0:hh.mm.ss}";
            //var ctl =
            //    new AutorizeSdgReport.AutorizeSdgReportCtl();//(new DB.MockDataLayer()));
            //var or = new revenewSameryByClinet.revenewSameryByClinet();
            //or.PreDisplay();

            //     var rbc = new GetLayoutByClient.Form1();
            //  rbc.Show();

            //or.Show();
            //this.Controls.Add(or);
            //return;
            //           var dal = new MockDataLayer();
            //         dal.Connect();
            //      PrintWorkshhet.PrintWorkshhetCls a=new PrintWorkshhetCls();
            //       var s = dal.GetClientByID(41);
            //


            //var rbc =
            //    new ResultsByClientCls();
            // rbc.DEBUG = true;
            // rbc.PreDisplay();
            // rbc.GetInitData(4234.ToString());
            var assoc =
                new FoodResultEntry.FoodResultEntry();
            assoc.SetParameters("a");
            assoc.PreDisplay();
            assoc.DEBUG = true;
            assoc.Show();
            assoc.Dock = DockStyle.Fill;
            this.Controls.Add(assoc);
       //     assoc.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;

                 return;
            persons1 = new List<string>();
            persons2 = new List<string>();
            for (int i = 0; i < 10; i++)
            {
                if (i > 4)
                {
                    persons1.Add(i + " aa");
                }
                else
                {
                    persons2.Add(i + " aa");


                }
            }
            //var ctl = new Order_cls();
            //Order_cls.DEBUG = true;
            //ctl.PreDisplay();
            //persons3 = persons2;
            //radDropDownList1.DataSource = persons3;
            //ctl.RightToLeft = RightToLeft.Yes;
            //    this.Controls.Add(ctl);


        }

        private void radGridView1_UserDeletingRow(object sender, Telerik.WinControls.UI.GridViewRowCancelEventArgs e)
        {

        }

        private void radGridView1_UserDeletedRow(object sender, Telerik.WinControls.UI.GridViewRowEventArgs e)
        {

        }

        private void radGridView1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            persons3 = persons1;
            radDropDownList1.DataSource = persons3;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }

    public class Person
    {

        public string Name { get; set; }

        public int Id { get; set; }

        public Person(string name, int id)
        {
            Name = name;
            Id = id;
        }
    }
}

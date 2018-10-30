using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LoadDeathsGUI
{

    public partial class DeathDivorceForm1 : Form
    {


        public DeathDivorceForm1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;
            toolStripStatusLabel1.Text = "Ready";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //String Length 
            int iSSNlength = textBox1.Text.Length;
            CheckSSN(iSSNlength);
            EnableSubmit();

        }

        private void button1_Click(object sender, EventArgs e)
        {

            Global.sTheSSN = textBox1.Text;

            button1.Enabled = false;
            toolStripStatusLabel1.Text = "Loading SQL...";
            toolStripProgressBar1.ProgressBar.Style = ProgressBarStyle.Marquee;
            toolStripProgressBar1.ProgressBar.MarqueeAnimationSpeed = 30;

            //LoadSQL.OpenSqlConnection(Submit.sTheSSN);

            if (backgroundWorker1.IsBusy != true)
            {
                Console.WriteLine("Starting BackGround Worker...");
                backgroundWorker1.RunWorkerAsync();
            }

        }

        public void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Get The Text of the selected option
            string sTheSelectedText = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
            CheckPrimaryDependant(sTheSelectedText);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

            //convert time to yyyyMMdd format 
            string sSelectedDate = dateTimePicker1.Value.ToString("yyyyMMdd");
            Console.WriteLine(sSelectedDate);

            Global.bIsDateSelected = true;
            Global.sTheSelectedDate = sSelectedDate;

            EnableSubmit();
           

        }

        public bool EnableSubmit()
        {
            if (Global.bIsSSNFilled == true && Global.bIsDateSelected == true && Global.bIsPrimaryDependentSelected == true && Global.isDeath == true || Global.isDivorce == true)
            {
                button1.Enabled = true;
                return true;
            }
            else
            {
                button1.Enabled = false;
                return false;
            }


        }

        public bool CheckSSN(int iSSNlength)
        {
            //If SSN is less then 9 digits then turn the text red.
            if (iSSNlength == 9)
            {
                Global.bIsSSNFilled = true;
                textBox1.ForeColor = Color.Black;
                return true;
            }
            else
            {
                Global.bIsSSNFilled = false;
                textBox1.ForeColor = Color.Red;
                return false;

            }

           
        }

        public void CheckPrimaryDependant(string sTheSelectedText)
        {
            if (sTheSelectedText != "")
            {
                Global.bIsPrimaryDependentSelected = true;
            }

            if (sTheSelectedText == "Primary")
            {
                Global.isPrimary = true;
                Console.WriteLine("Primary");
               
            }
            else
            {
                Global.isPrimary = false;
                Console.WriteLine("Dependent");
               
            }

            if (EnableSubmit())
            {
                button1.Enabled = true;
            }

            Console.WriteLine(sTheSelectedText);
        }

        private void toolStripProgressBar1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            LoadSQL.OpenSqlConnection(Global.sTheSSN);

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            Console.WriteLine("BackGround Worker Complete...");
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                // Next, handle the case where the user canceled 
                // the operation.
                // Note that due to a race condition in 
                // the DoWork event handler, the Cancelled
                // flag may not have been set, even though
                // CancelAsync was called.
                toolStripStatusLabel1.Text = "Canceled";
            }
            else
            {
                // Finally, handle the case where the operation succeeded.
                //MessageBox.Show(e.Result.ToString());

                //Clear Data After Submit
                ClearFormData();

            }

            toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
            toolStripProgressBar1.MarqueeAnimationSpeed = 0;
            toolStripStatusLabel1.Text = "Ready...";
           

        }

        public void ClearFormData()
        {
            //Clear the contents of the segments array
            if (Global.aCampgainSegments != null)
            {
                Global.aCampgainSegments.Clear();
            }
            

            //Reset the SSN Text field
            textBox1.Text = string.Empty;
            Global.sTheSSN = string.Empty;
            Global.bIsSSNFilled = false;
            
            //Reset Death or Divorce
            deathComboBox.ResetText();
            Global.isDivorce = false;
            Global.isDeath = false;

            //Reset The death date field
            dateTimePicker1.ResetText();
            Global.sTheSelectedDate = string.Empty;

            // Reset the primary Primary/Dependent combobox
            comboBox1.ResetText();
            Global.isPrimary = false;
            Global.bIsPrimaryDependentSelected = false;

            //Disable the submit so we can start over
            button1.Enabled = false;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //For some reason the combobox defaults to death when nothing is selected
            if (deathComboBox.Text == "Death")
            {
                Global.isDeath = true;
                Global.isDivorce = false;

                Console.WriteLine("Death");
            }
            else if (deathComboBox.Text == "Divorce")
            {
                Global.isDivorce = true;
                Global.isDeath = false;

                Console.WriteLine("Divorce");

            }

            EnableSubmit();

        }
    }
}

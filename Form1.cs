using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ivi.Visa.Interop;

namespace SGenerator_Pico
{
    public partial class Form1 : Form
    {
        DCPSG4042 psu = new DCPSG4042();
        internal FormattedIO488 Instrument;
        public Form1()
        {
            InitializeComponent();
            //string IDN;
            //psu.Connect("TCPIP0::192.168.1.44::inst0::INSTR", out IDN);
        }
        public class DCPSG4042
        {
            internal ResourceManager rm;
            internal FormattedIO488 Instrument;
            internal string Mod = "TMSL";
            internal string ResourceAddress;
            internal string[] errors;

            public bool Connect(string ResourceAddress_in, out string IDN)
            {
                IDN = "";
                rm = new ResourceManager();
                Instrument = new FormattedIO488();
                string dil = "";
                ResourceAddress = ResourceAddress_in;
                try
                {
                    Instrument.IO = (IMessage)rm.Open(ResourceAddress, AccessMode.NO_LOCK, 0, "");

                    //Instrument.WriteString("SYST:LANG?", true);         // geçerli dil ayarını sorgula
                    //dil = Instrument.ReadString();

                    //  if (Mod == "TMSL")
                    // {
                    // if (dil.Contains("SCPI"))
                    // {
                    //Instrument.WriteString("*RST", true);
                    //Instrument.WriteString("*CLS", true);
                    // GetErrors(out errors);
                    Instrument.WriteString("*IDN?", true);
                    IDN = Instrument.ReadString();


                    //Instrument.WriteString("*IDN?", true);
                    //Instrument.WriteString("c1:outp on", true);
                    //Thread.Sleep(2000);
                    //Instrument.WriteString("c1:outp off", true);

                    // DCPSOutputONOFF(false);
                    // PSU_Connection_Status = "Passed";
                    // PSU_Conn = "True";
                    return true;
                    //}
                    // else
                    // {
                    //    PSU_Conn = "False";
                    //     return false;
                    //  }
                    //   }
                    //  else
                    //      return false;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 bağlantı hatası! (Adres = " + ResourceAddress + ")");
                    return false;
                }
            }
            public bool Ch1OutputOpen()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:outp on", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Ch1 output çıkışı sağlanamadı.");
                    return false;
                }
            }
            public bool Ch1OutputClose()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:outp off", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Ch1 output çıkışı sağlanamadı.");
                    return false;
                }
            }
            public bool Ch2OutputClose()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:outp off", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Ch1 output çıkışı sağlanamadı.");
                    return false;
                }
            }
            public bool Ch2OutputOpen()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:outp on", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Ch1 output çıkışı sağlanamadı.");
                    return false;
                }
            }


            public bool ReadID(out string DeviceInfo)
            {
                DeviceInfo = "";
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("*IDN?", true);
                    DeviceInfo = Instrument.ReadString();
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show((IWin32Window)exp, "G40-42 ID sorgulama hatası!");
                    return false;
                }
            }


            public bool SetFrequency1(double Frequency)
            {
                if (Frequency<=0.009 || Frequency<=0.1)
                {
                    try
                    {
                        if (Mod == "TMSL")
                            Instrument.WriteString("c1:bswv frq," + string.Format("{0:F2}", Frequency*1000+"E-3") + "khz", true);
                        return true;
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("G40-42 Frekans ayarlama hatası! (Frequency = " + Frequency + ")");
                        return false;
                    }
                }
                
                else
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv frq," + string.Format("{0:F2}", Frequency ) + "kHz", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Frekans ayarlama hatası! (Frequency = " + Frequency + ")");
                    return false;
                }
            }
            public bool SetFrequency2(decimal Frequency)
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv frq," + string.Format("{0:F2}", Frequency ) + "kHz", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Frekans ayarlama hatası! (Frequency = " + Frequency + ")");
                    return false;
                }
            }
            public bool SetSine1()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv wvtp, " + string.Format("{0:F2}", "SINE"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 sine ayarlama hatası! ");
                    return false;
                }
            }
            public bool SetSine2()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv wvtp, " + string.Format("{0:F2}", "SINE"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 sine ayarlama hatası! ");
                    return false;
                }
            }
            public bool SetSquare1()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv wvtp, " + string.Format("{0:F2}", "SQUARE"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Square ayarlama hatası! ");
                    return false;
                }
            }
            public bool SetSquare2()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv wvtp, " + string.Format("{0:F2}", "SQUARE"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Square ayarlama hatası! ");
                    return false;
                }
            }
            public bool SetRamp1()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv wvtp, " + string.Format("{0:F2}", "RAMP"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Ramp ayarlama hatası! ");
                    return false;
                }
            }
            public bool SetRamp2()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv wvtp, " + string.Format("{0:F2}", "RAMP"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Ramp ayarlama hatası! ");
                    return false;
                }
            }
            public bool SetPulse1()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv wvtp, " + string.Format("{0:F2}", "PULSE"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Pulse ayarlama hatası! ");
                    return false;
                }
            }
            public bool SetPulse2()
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv wvtp, " + string.Format("{0:F2}", "PULSE"), true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Pulse ayarlama hatası! ");
                    return false;
                }
            }

            public bool SetVolt1(double Voltage)
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv amp, " + string.Format("{0:F2}", Voltage) + "V", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 gerilim ayarlama hatası! (Voltage = " + Voltage + ")");
                    return false;
                }
            }
            public bool SetVolt2(double Voltage)
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv amp, " + string.Format("{0:F2}", Voltage) + "V", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 gerilim ayarlama hatası! (Voltage = " + Voltage + ")");
                    return false;
                }
            }
            public bool SetOffset1(double Offset)
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv ofst, " + string.Format("{0:F2}", Offset) + "V", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 offset ayarlama hatası! (Voltage = " + Offset + ")");
                    return false;
                }
            }
            public bool SetOffset2(double Offset)
            {
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv ofst, " + string.Format("{0:F2}", Offset) + "V", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 offset ayarlama hatası! (Voltage = " + Offset + ")");
                    return false;
                }
            }
            public bool SetPhase1(double Phase)
            {
                if (Phase < 360)
                {
                    try
                    {
                        if (Mod == "TMSL")
                            Instrument.WriteString("c1:bswv phse, " + string.Format("{0:F2}", Phase) + "V", true);
                        return true;
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("G40-42 Phase ayarlama hatası! (Voltage = " + Phase + ")");
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show("Phase 360'dan küçük olmalıdır.");
                    return false;
                }

                //try
                //{
                //    if (Mod == "TMSL")
                //        Instrument.WriteString("bswv phase, " + string.Format("{0:F2}", Phase) + "V", true);
                //    return true;
                //}
                //catch (Exception exp)
                //{
                //    MessageBox.Show("G40-42 Phase ayarlama hatası! (Voltage = " + Phase + ")");
                //    return false;
                //}
            }
            public bool SetPhase2(double Phase)
            {
                if (Phase < 360)
                {
                    try
                    {
                        if (Mod == "TMSL")
                            Instrument.WriteString("c2:bswv phse, " + string.Format("{0:F2}", Phase) + "V", true);
                        return true;
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("G40-42 Phase ayarlama hatası! (Voltage = " + Phase + ")");
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show("Phase 360'dan küçük olmalıdır.");
                    return false;
                }

                //try
                //{
                //    if (Mod == "TMSL")
                //        Instrument.WriteString("bswv phase, " + string.Format("{0:F2}", Phase) + "V", true);
                //    return true;
                //}
                //catch (Exception exp)
                //{
                //    MessageBox.Show("G40-42 Phase ayarlama hatası! (Voltage = " + Phase + ")");
                //    return false;
                //}
            }
            public bool SetHighLowLevel1(double HighLevel,double LowLevel)
            {
                
                    try
                    {
                        if (Mod == "TMSL")
                            Instrument.WriteString("c1:bswv hlev," + string.Format("{0:F2}", HighLevel)+","+"llev,"+ string.Format("{0:F2}", LowLevel) + "V", true);
                        return true;
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("G40-42 High level ayarlama hatası! (Voltage = " + HighLevel + ")");
                        return false;
                    }
                

                


            }
            public bool SetHighLevel2(double HighLevel)
            {

                
                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv hlev," + string.Format("{0:F2}", HighLevel) + "V", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 High level ayarlama hatası! (Voltage = " + HighLevel + ")");
                    return false;
                }


            }
            public bool SetLowLevel1(double LowLevel)
            {


                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c1:bswv llev, " + string.Format("{0:F2}", LowLevel) + "V", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Low level ayarlama hatası! (Voltage = " + LowLevel + ")");
                    return false;
                }


            }
            public bool SetLowLevel2(double LowLevel)
            {


                try
                {
                    if (Mod == "TMSL")
                        Instrument.WriteString("c2:bswv llev, " + string.Format("{0:F2}", LowLevel) + "V", true);
                    return true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("G40-42 Low level ayarlama hatası! (Voltage = " + LowLevel + ")");
                    return false;
                }


            }






            ~DCPSG4042()
            {
                try
                {
                    Instrument.IO.Close();
                    Marshal.ReleaseComObject(Instrument);
                    Marshal.ReleaseComObject(rm);
                }
                catch (Exception)
                {
                    throw;
                }

            }



        }
       

        private void btnPico_Click_1(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
            this.Hide();
        }

        private void btnCh1_Click(object sender, EventArgs e)
        {
            psu.Ch1OutputOpen();
            label2.Text = "CH1 output açıldı.";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            psu.Ch1OutputClose();
            label2.Text = "CH1 output kapatıldı.";
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            string IDN;
            psu.Connect("TCPIP0::192.168.1.44::inst0::INSTR", out IDN);
            string DeviceInfo;
            psu.ReadID(out DeviceInfo);
            string device = DeviceInfo.Substring(0, 8);
            label1.Text = device + " adlı cihaza bağlanıldı.";

        }

        private void button5_Click(object sender, EventArgs e)
        {
            double frekans = Convert.ToDouble(textBox2.Text);
            if (comboBox1.SelectedIndex==0)
            {
                frekans = frekans*1;
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                frekans = frekans * 1000;
            } 
            else if (comboBox1.SelectedIndex == 2)
            {
                frekans = frekans /1000 ;
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                frekans = frekans * 1000000;
            }
            
            double offset = Convert.ToDouble(textBox3.Text);
            double volt = Convert.ToDouble(textBox1.Text);
            double phase = Convert.ToDouble(textBox4.Text);
            double highlevel = Convert.ToDouble(textBox12.Text);
            double lowlevel = Convert.ToDouble(textBox11.Text);
           
            psu.SetFrequency1(frekans);
            psu.SetOffset1(offset);
            psu.SetVolt1(volt);

            if (highlevel!=0 || lowlevel!=0)
            {
                psu.SetHighLowLevel1(highlevel, lowlevel);
            }
            
            psu.SetPhase1(phase);
            
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            double volt = Convert.ToDouble(textBox6.Text);
            decimal frekans = Convert.ToDecimal(textBox8.Text);
            if (comboBox2.SelectedIndex == 0)
            {
                frekans = frekans * 1;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                frekans = frekans * 1000;
            }
            else if (comboBox2.SelectedIndex == 2)
            {
                frekans = frekans / 1000;
            }
            else if (comboBox2.SelectedIndex == 3)
            {
                frekans = frekans * 1000000;
            }
            
            double offset = Convert.ToDouble(textBox7.Text);
            double phase = Convert.ToDouble(textBox5.Text);                        
            double highlevel = Convert.ToDouble(textBox9.Text);
            if (highlevel!=0)
            {
                psu.SetHighLevel2(highlevel);

            }
            double lowlevel = Convert.ToDouble(textBox10.Text);
            if (lowlevel!=0)
            {
                psu.SetLowLevel2(lowlevel);
            }
            psu.SetFrequency2(frekans);
            psu.SetOffset2(offset);
            psu.SetVolt2(volt);
            psu.SetPhase2(phase);
            
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            psu.Ch2OutputOpen();
            label22.Text = "CH2 output açıldı.";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            psu.Ch2OutputClose();
            label22.Text = "CH2 output kapatıldı.";
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            psu.SetSine1();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            psu.SetSquare1();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            psu.SetPulse1();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            psu.SetRamp1();
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            psu.SetSine2();
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            psu.SetSquare2();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            psu.SetPulse2();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            psu.SetRamp2();
        }
        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox1.BackColor = Color.Black;
        }
        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
           pictureBox1.BackColor = Color.White;
        }
        private void pictureBox2_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox2.BackColor = Color.Black;
        }

        private void pictureBox2_MouseUp(object sender, MouseEventArgs e)
        {
            pictureBox2.BackColor = Color.White;
        }
        private void pictureBox3_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox3.BackColor = Color.Black;
        }

        private void pictureBox3_MouseUp(object sender, MouseEventArgs e)
        {
            pictureBox3.BackColor = Color.White;
        }
        private void pictureBox4_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox4.BackColor = Color.Black;
        }

        private void pictureBox4_MouseUp(object sender, MouseEventArgs e)
        {
            pictureBox4.BackColor = Color.White;
        }
        private void pictureBox5_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox5.BackColor = Color.Black;
        }

        private void pictureBox5_MouseUp(object sender, MouseEventArgs e)
        {
            pictureBox5.BackColor = Color.White;
        }
        private void pictureBox6_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox6.BackColor = Color.Black;
        }

        private void pictureBox6_MouseUp(object sender, MouseEventArgs e)
        {
            pictureBox6.BackColor = Color.White;
        }
        private void pictureBox7_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox7.BackColor = Color.Black;
        }

        private void pictureBox7_MouseUp(object sender, MouseEventArgs e)
        {
            pictureBox7.BackColor = Color.White;
        }
        private void pictureBox8_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox8.BackColor = Color.Black;
        }

        private void pictureBox8_MouseUp(object sender, MouseEventArgs e)
        {
            pictureBox8.BackColor = Color.White;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("Hz");
            comboBox1.Items.Add("kHz");
            comboBox1.Items.Add("mHz");
            comboBox1.Items.Add("Mhz");
            
            comboBox1.SelectedIndex = 1;


            comboBox2.Items.Add("Hz");
            comboBox2.Items.Add("kHz");
            comboBox2.Items.Add("mHz");
            comboBox2.Items.Add("Mhz");
            
            comboBox2.SelectedIndex = 1;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }
    }
}

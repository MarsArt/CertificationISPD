using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CertificationISPD
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            
        }
        
        private void button1_Click(object sender, EventArgs e)
        {

            int vis = 0;
            int avg = 0;
            int litel = 0;

            if (radioButton5.Checked) vis++;
            if (radioButton8.Checked) vis++;
            if (radioButton9.Checked) vis++;
            if (radioButton16.Checked) vis++;
            if (radioButton18.Checked) vis++;
            if (radioButton22.Checked) vis++;

            if (radioButton3.Checked) avg++;
            if (radioButton4.Checked) avg++;
            if (radioButton7.Checked) avg++;
            if (radioButton10.Checked) avg++;
            if (radioButton12.Checked) avg++;
            if (radioButton17.Checked) avg++;
            if (radioButton21.Checked) avg++;

            if (radioButton1.Checked) litel++;
            if (radioButton2.Checked) litel++;
            if (radioButton6.Checked) litel++;
            if (radioButton11.Checked) litel++;
            if (radioButton13.Checked) litel++;
            if (radioButton14.Checked) litel++;
            if (radioButton15.Checked) litel++;
            if (radioButton19.Checked) litel++;
            if (radioButton20.Checked) litel++;
            Form1 f1 = new Form1();
            string Zach = "Низкий уровень исходной защищенности";
            if (litel > 0)
            {
                Zach = "Низкий уровень исходной защищенности";
                Properties.Settings.Default.Y1 = 10;
            }
            else
            {

                if (((float)vis / 7) > 0.7)
                {
                    Zach = "Высокий уровень исходной защищенности";
                    Properties.Settings.Default.Y1 = 0;
                }
                else
                {
                    Properties.Settings.Default.Y1 = 5;
                    Zach = "Средний уровень исходной защищенности";
                }
            }
            textBox1.Text = Zach;
           
            MessageBox.Show("VIS: "+vis.ToString()+ "\n avg: "+avg.ToString()+"\n litel: "+ litel.ToString() +"\n %: "+ ((float)vis /7).ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                WordReportGenerator WR = new WordReportGenerator();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK || saveFileDialog1.FileName == "")
                {
                    WR.PathToSave = saveFileDialog1.FileName + ".docx";
                    if (File.Exists("IZ.docx"))
                    {
                        WR.template = "IZ.docx";
                        string terrazm = (radioButton1.Checked) ? radioButton1.Text :
                                     (radioButton2.Checked) ? radioButton2.Text :
                                     (radioButton3.Checked) ? radioButton3.Text :
                                     (radioButton4.Checked) ? radioButton4.Text :
                                     (radioButton5.Checked) ? radioButton5.Text : "";
                        WR.Add("@terrazm", terrazm);

                        string soed = (radioButton6.Checked) ? radioButton6.Text :
                                     (radioButton7.Checked) ? radioButton7.Text :
                                     (radioButton8.Checked) ? radioButton8.Text : "";
                        WR.Add("@soed", soed);

                        string operazi = (radioButton9.Checked) ? radioButton9.Text :
                                    (radioButton10.Checked) ? radioButton10.Text : "";
                        WR.Add("@operazi", operazi);

                        string dostyp = (radioButton11.Checked) ? radioButton11.Text :
                                    (radioButton12.Checked) ? radioButton12.Text : "";
                        WR.Add("@dostyp", dostyp);

                        string baza = (radioButton13.Checked) ? radioButton13.Text :
                                      (radioButton14.Checked) ? radioButton14.Text : "";
                        WR.Add("@baza", baza);

                        string yroven = (radioButton15.Checked) ? radioButton15.Text :
                                        (radioButton16.Checked) ? radioButton16.Text :
                                        (radioButton18.Checked) ? radioButton18.Text : "";
                        WR.Add("@yroven", yroven);

                        string obem = (radioButton17.Checked) ? radioButton17.Text :
                                            (radioButton19.Checked) ? radioButton19.Text :
                                            (radioButton20.Checked) ? radioButton20.Text : "";
                        WR.Add("@obem", obem);

                        WR.Add("@Yrzash", textBox1.Text);


                        WR.Generate();
                        MessageBox.Show("Отчет успешно создан! Путь документа: " + WR.PathToSave, "Отчет");
                    }
                }
                else MessageBox.Show("Вы не указали путь для отчета! Отчет не был создан", "Ошибка!!!");

            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}

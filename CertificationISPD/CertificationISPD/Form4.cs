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
using DocumentFormat.OpenXml;

namespace CertificationISPD
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();            
        }
       

        public List<TypeYgroz> listYgroz;
       
        private void Form4_Load(object sender, EventArgs e)
        {
            listYgroz = new List<TypeYgroz>();
            listYgroz.Add(new TypeYgroz(1, "угроза утечки акустической (речевой) информации"));
            listYgroz.Add(new TypeYgroz(2, "угроза утечки видовой информации"));
            listYgroz.Add(new TypeYgroz(3, "угроза утечки информации по каналу ПЭМИН"));
            listYgroz.Add(new TypeYgroz(4, "угроза НСД"));
            listYgroz.Add(new TypeYgroz(5, "угроза Анализа сетевого трафика» с перехватом передаваемой во внешние сети и принимаемой из внешних сетей информации"));
            listYgroz.Add(new TypeYgroz(6, "угроза сканирования, направленные на выявление типа операционной системы АРМ, открытых портов и служб, открытых соединений и др."));
            listYgroz.Add(new TypeYgroz(7, "угроза выявления паролей"));
            listYgroz.Add(new TypeYgroz(8, "угроза получения НСД путем подмены доверенного объекта"));
            listYgroz.Add(new TypeYgroz(9, "угроза типа «Отказ в обслуживании»"));
            listYgroz.Add(new TypeYgroz(10, "угроза удаленного запуска приложений"));
            listYgroz.Add(new TypeYgroz(11, "угроза внедрения по сети вредоносных программ"));
            listYgroz.Add(new TypeYgroz(12, "угрозы внедрения ложного объекта сети"));
            listYgroz.Add(new TypeYgroz(13, "угрозы навязывания ложного маршрута путем несанкционированного изменения маршрутно - адресных данных"));
            listYgroz.Add(new TypeYgroz(14, "угроза внедрения по сети вредоносных программ"));
            bindingSource1.DataSource = listYgroz;
            dataGridView1.DataSource = bindingSource1;
            bindingNavigator1.BindingSource = bindingSource1;
            fillzero();


        }

        private void button2_Click(object sender, EventArgs e)
        {
          
            try
            {
                listYgroz.Add(new TypeYgroz(listYgroz.Count+1, textBox1.Text));
                bindingSource1.DataSource = null;
                bindingSource1.DataSource = listYgroz;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows[1].Cells[11].Value = "sadsad";
            findMax(ref dataGridView1);
        }

        private void findMax(ref DataGridView DG) {

            if (Properties.Settings.Default.Y1 == -1)
            {
                MessageBox.Show("Необходимо заполнить акт классификации уровня защищенности");
                new Form3().ShowDialog();
            }
            else
            {

                try
                {
                    for (int i = 0; i < DG.RowCount - 1; i++)
                    {
                        int max = int.Parse(DG.Rows[i].Cells[2].Value.ToString());
                        for (int rowindex = 3; rowindex < 10; rowindex++)
                        {
                            if (int.Parse(DG.Rows[i].Cells[rowindex].Value.ToString()) > max) max = int.Parse(DG.Rows[i].Cells[rowindex].Value.ToString());
                        }
                        dataGridView1.Rows[i].Cells[11].Value = max.ToString();
                        float Y = (((float)Properties.Settings.Default.Y1 + (float)max) / 20);
                        dataGridView1.Rows[i].Cells[12].Value = Y.ToString();
                        string VR = "низкая";
                        if (Y >= 0 && Y <= 0.3) { VR = "Низкая"; }
                        if (Y > 0.3 && Y <= 0.6) { VR = "Средняя"; }
                        if (Y > 0.6 && Y <= 0.8) { VR = "Высокая"; }
                        if (Y > 0.8 && Y <= 1) { VR = "Очень высокая"; }
                        dataGridView1.Rows[i].Cells[13].Value =VR;
                        string AR = "неактуальная";
                        string OR = dataGridView1.Rows[i].Cells[13].Value.ToString();
                        if (VR == "Низкая") {
                            if (OR == "Низкая") AR = "Неактуальная";
                            if (OR == "Средняя") AR = "Неактуальная";
                            if (OR == "Высокая") AR = "Актуальная";
                        } else
                        if (VR == "Средняя") {
                            if (OR == "Низкая") AR = "Неактуальная";
                            if (OR == "Средняя") AR = "Актуальная";
                            if (OR == "Высокая") AR = "Актуальная";
                        }else
                        if (VR == "Высокая")
                        {
                            if (OR == "Низкая") AR = "Актуальная";
                            if (OR == "Средняя") AR = "Актуальная";
                            if (OR == "Высокая") AR = "Актуальная";
                        }else
                        if (VR == "Очень высокая")
                        {
                            if (OR == "Низкая") AR = "Актуальная";
                            if (OR == "Средняя") AR = "Актуальная";
                            if (OR == "Высокая") AR = "Актуальная";
                        }
                        dataGridView1.Rows[i].Cells[15].Value = VR;
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
        public void fillzero() {
            try
            {
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    for (int j = 2; j <= 10; j++)
                    {
                        dataGridView1.Rows[i].Cells[j].Value = "0";
                    }
                    dataGridView1.Rows[i].Cells[14].Value = "Низкая";
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void Form4_VisibleChanged(object sender, EventArgs e)
        {
         //   fillzero();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fillzero();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string res = "";
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
               res+=dataGridView1.Rows[i].Cells[0].Value.ToString()+" "+ dataGridView1.Rows[i].Cells[1].Value.ToString()+" "+ dataGridView1.Rows[i].Cells[12].Value.ToString()+" "+ dataGridView1.Rows[i].Cells[13].Value.ToString() +" "+ dataGridView1.Rows[i].Cells[14].Value.ToString()+ " " + dataGridView1.Rows[i].Cells[15].Value.ToString()+"\n";
            }
            WordReportGenerator WRG = new WordReportGenerator();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK || saveFileDialog1.FileName == "")
            {
                WRG.PathToSave = saveFileDialog1.FileName + ".docx";
                if (File.Exists("model.docx"))
                {
                    WRG.template = "model.docx";
                    WRG.Add("@ressssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssss", res);
                    WRG.Generate();
                    MessageBox.Show("Отчет успешно создан! Путь документа: " + WRG.PathToSave, "Отчет");
                }
                MessageBox.Show(res);
            }
        }

       
    }
}

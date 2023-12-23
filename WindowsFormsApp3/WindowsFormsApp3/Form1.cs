
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        static int id = 0;
        public Store store = new Store();
        Config config = new Config();
        private RepExcel excel = new RepExcel();
        public Form1()
        {
            InitializeComponent();
        }

        private void ChangeLanguage(string newLanguageString)
        {
            try
            {
                var resources = new ComponentResourceManager(typeof(Form1));
                CultureInfo newCultureInfo = new CultureInfo(newLanguageString);
                foreach (Control c in this.Controls)
                { resources.ApplyResources(c, c.Name, newCultureInfo); }
                resources.ApplyResources(this, "$this", newCultureInfo);
                foreach (var item in SS_Status.Items.Cast<ToolStripItem>().Where(item => (item is ToolStripStatusLabel) != false))
                {
                    resources.ApplyResources(item, item.Name, newCultureInfo);
                }
                TSDDB.Text = newCultureInfo.NativeName;
                SetCurrenLanguageButtonChecked();
            }
            catch
            {
                MessageBox.Show("Error 404!", "Error 1", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ExitandSave();
            }

        }

        private void SetCurrenLanguageButtonChecked()
        {
            foreach (ToolStripMenuItem languageButton in TSDDB.DropDownItems)
            {

                languageButton.Checked = (languageButton.Text == TSDDB.Text);
            }
        }


        private void Init()
        {

            try
            {
                config = new Config();
                config = Serializator.DeserializeXML<Config>(ref config, "config");
                if (config.TypeOfSerialization == "XML")
                {
                    store.goods = Serializator.DeserializeXML<List<Goods>>(ref store.goods, config.Filepath);
                }
                else
                {
                    if (config.TypeOfSerialization == "json")
                    {
                        store.goods = Serializator.DeserializeJson<List<Goods>>(ref store.goods, config.Filepath);
                    }
                    else
                    {
                        store.goods = Serializator.DeserializeBinary<List<Goods>>(ref store.goods, config.Filepath);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Error 405!", "Error 2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Exit();
            }

        }

        private void DataGridViewInput()
        {
            try
            {
                dataGridViewGoods.RowCount = store.goods.Count + 1;
                dataGridViewGoods.ColumnCount = 4;
                for (int i = 0; i < dataGridViewGoods.Rows.Count - 1; i++)
                {
                    dataGridViewGoods.Rows[i].Cells[0].Value = store.goods[i].name;
                    dataGridViewGoods.Rows[i].Cells[1].Value = store.goods[i].price;
                    dataGridViewGoods.Rows[i].Cells[2].Value = store.goods[i].count;
                    dataGridViewGoods.Rows[i].Cells[3].Value = store.goods[i].code;
                    id++;
                }
            }
            catch
            {
                MessageBox.Show("Error 406!", "Error 3", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Exit();
            }
        }

        private void Exit()
        {
            Application.Exit();
        }

        private void Save()
        {
            SerializeStore(store.goods);
            if (!FilePathIsEmpty())
            {
                config.Filepath = textBox4.Text;
            }
            try
            {
                Serializator.SerializeConfiguration(config);
            }
            catch
            {
                MessageBox.Show("Error 407!", "Error 4", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExitandSave()
        {
            Save();
            Exit();
        }

        private void ModuleOfSerializeStorage<T>(T obj, string filename)
        {
            try
            {
                if (config.TypeOfSerialization == "XML")
                {
                    Serializator.SerializeXML<T>(ref obj, filename);
                }
                else
                {
                    if (config.TypeOfSerialization == "json")
                    {
                        Serializator.SerializeJson<T>(ref obj, filename);
                    }
                    else
                    {
                        Serializator.SerializeBinary<T>(ref obj, filename);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Error 408!", "Error 5", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Exit();
            }
        }

        private bool FilePathIsEmpty()
        {
            if (textBox4.Text == "")
                return true;

            return false;
        }

        private void SerializeStore<T>(T obj)
        {
            if (FilePathIsEmpty())
                ModuleOfSerializeStorage(obj, config.Filepath);
            else
                ModuleOfSerializeStorage(obj, textBox4.Text);
            TSL.Visible = true;
            timer2.Start();
        }

        private void Add()
        {
            try
            {
                Regex regex1 = new Regex(@"^[A-Za-zа-яА-ЯЁё ]{2,256}$");
                Regex regex2 = new Regex(@"^[^0]{1}[0-9]*?$");
                Regex regex3 = new Regex(@"^[^0]{1}[0-9]*?$");
                Regex regex4 = new Regex(@"^[^0]{1}[0-9]*?$");
                Regex regex5 = new Regex(@"^[^0]{1}[0-9]*?$");

                if (regex1.IsMatch(textBox1.Text) && regex2.IsMatch(textBox2.Text) && regex3.IsMatch(textBox3.Text) && regex4.IsMatch(textBox5.Text))
                {
                    dataGridViewGoods.Rows.Add();
                    dataGridViewGoods.Rows[id].Cells[0].Value = textBox1.Text;
                    dataGridViewGoods.Rows[id].Cells[1].Value = textBox2.Text;
                    dataGridViewGoods.Rows[id].Cells[2].Value = textBox3.Text;
                    dataGridViewGoods.Rows[id].Cells[3].Value = textBox5.Text;
                    id++;
                    Goods addGoods = new Goods(textBox1.Text, Convert.ToDouble(textBox2.Text), Convert.ToInt32(textBox3.Text), Convert.ToInt32(textBox5.Text));
                    store.goods.Add(addGoods);
                    ClearFields();
                }
                else
                {
                    MessageBox.Show("Invalid data!", "Error 6", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch
            {
                MessageBox.Show("Error 409!", "Error 7", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ExitandSave();
            }
        }

        private void ClearFields()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            Init();
            DataGridViewInput();
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            SerializeStore(store.goods);
        }


        private void AddInfoButton_Click(object sender, EventArgs e)
        {
            Add();
        }

        private void SerializeButton_Click(object sender, EventArgs e)
        {
            SerializeStore(store.goods);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Save();
        }

        private void TextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 46)
                e.Handled = true;
        }

        private void ComboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            config.TypeOfSerialization = comboBox1.SelectedItem.ToString();
        }

        private void TSMI_English_Click(object sender, EventArgs e)
        {
            ChangeLanguage("en");
        }
        private void RussianToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeLanguage("ru");
        }

        private void Timer2_Tick(object sender, EventArgs e)
        {
            TSL.Visible = false;
            timer2.Stop();
        }

        private void dataGridViewGoods_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormReport frmReport = new FormReport();
            frmReport.ShowDialog();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            excel.CreateNewBook("test.xls");
            excel.OpenBook("test.xls");
            excel.SetValue("Saturn Data", "A3", "Hi", "string", true);
            excel.GetValue("Saturn Data", "A3");
            excel.HidenRow("Saturn Data", 4);
            excel.DisplayLine("Saturn Data", 4);
            excel.Save("test.xls");
            excel.SaveAs("test1.xls");
        }
    }
}

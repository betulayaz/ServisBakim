using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ServisBakim
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        BaglantiSinif bgl = new BaglantiSinif();
        SqlCommand command = new SqlCommand();

        void control()
        {
                SqlConnection connection = new SqlConnection(bgl.Address);
                connection.Open();
                SqlCommand command = new SqlCommand("select Max(Km) from ServisBakimOnarim where Plaka=@Plaka", connection);
                command.Parameters.AddWithValue("@Plaka", txt_plaque.Text);
                SqlDataReader dr = command.ExecuteReader();
          
                if (dr.Read())
                {
                    int datakm = dr.GetInt32(0);
                    int km = Convert.ToInt32(txt_km.Text);
                  
                    if (km <= datakm)
                    {
                        MessageBox.Show("Kayıt Yapılamaz Km Değeri Önceki Kayıtlardan Büyük Olmalıdır.");
                    }
                    else if (km > datakm)
                    {
                        connection.Close();
                        Save();
                    }
                }
                else
                {
                    connection.Close();
                    Save();
                }
                connection.Close();          
        }

        void Save()
        {
            try
            { 
                SqlConnection connection = new SqlConnection(bgl.Address);

                string typeofmaintenance = "";

                if (cb_periodic.Checked)

                    typeofmaintenance = typeofmaintenance + "," + cb_periodic.Text;

                if (cb_wintercare.Checked)
                    typeofmaintenance = typeofmaintenance + "," + cb_wintercare.Text;

                if (cb_summercare.Checked)
                    typeofmaintenance = typeofmaintenance + "," + cb_summercare.Text;

                if (cb_fault.Checked)
                    typeofmaintenance = typeofmaintenance + "," + cb_fault.Text;

                if (cb_tire.Checked)
                    typeofmaintenance = typeofmaintenance + "," + cb_tire.Text;

                if (cb_oilcare.Checked)
                    typeofmaintenance = typeofmaintenance + "," + cb_oilcare.Text;

                typeofmaintenance = typeofmaintenance.Substring(1);


                string faultoffailure = string.Empty;

                if (rdb_user.Checked)
                {
                    faultoffailure = "Kullanıcı";
                }
                else if (rdb_fuel.Checked)
                {
                    faultoffailure = "Yakıt";
                }
                else if (rdb_way.Checked)
                {
                    faultoffailure = "Yol Şartları";
                }
                else if (rdb_other.Checked)
                {
                    faultoffailure = "Diğer";
                }

                string insert = "INSERT INTO ServisBakimOnarim(Tarih,Plaka,Km,Birim,Marka,Model,Proje,Gelis,Cikis,Sofor,BakimOnarimTürü,Ariza,ArizaSebebi,Yapilanİsler,Malzeme,BakimOnarimMaliyeti,ServisteKaldigiSüre,Aciklama,BakimOnarimYapan,TeslimAlan,KayitYapan) values (@Tarih,@Plaka, @Km, @Birim,@Marka, @Model,@Proje,@Gelis,@Cikis,@Sofor,@BakimOnarimTürü,@Ariza,@ArizaSebebi,@Yapilanİsler,@Malzeme,@BakimOnarimMaliyeti,@ServisteKaldigiSüre,@Aciklama,@BakimOnarimYapan,@TeslimAlan,@KayitYapan)";
                SqlCommand command = new SqlCommand();
                command = new SqlCommand(insert, connection);
                connection.Open();

                command.Connection = connection;

                command.Parameters.AddWithValue("@Tarih", txt_date.Text);
                command.Parameters.AddWithValue("@Plaka", txt_plaque.Text);
                command.Parameters.AddWithValue("@Km", txt_km.Text);
                command.Parameters.AddWithValue("@Birim", txt_unit.Text);
                command.Parameters.AddWithValue("@Marka", txt_brand.Text);
                command.Parameters.AddWithValue("@Model", txt_model.Text);
                command.Parameters.AddWithValue("@Proje", txt_project.Text);
                command.Parameters.AddWithValue("@Gelis", txt_arrival.Text);
                command.Parameters.AddWithValue("@Cikis", txt_exit.Text);
                command.Parameters.AddWithValue("@Sofor", txt_driver.Text);
                command.Parameters.AddWithValue("@BakimOnarimTürü", typeofmaintenance);
                command.Parameters.AddWithValue("@Ariza", txt_fault.Text);
                command.Parameters.AddWithValue("@ArizaSebebi", faultoffailure);
                command.Parameters.AddWithValue("@Yapilanİsler", txt_works.Text);
                command.Parameters.AddWithValue("@Malzeme", txt_material.Text);
                command.Parameters.AddWithValue("@BakimOnarimMaliyeti", txt_cost.Text);

                command.Parameters.AddWithValue("@Aciklama", txt_note.Text);
                command.Parameters.AddWithValue("@BakimOnarimYapan", txt_maintainer.Text);
                command.Parameters.AddWithValue("@TeslimAlan", txt_deliveryarea.Text);
                command.Parameters.AddWithValue("@KayitYapan", txt_recording.Text);

                DateTime arrival = Convert.ToDateTime(txt_arrival.Text);
                DateTime exit = Convert.ToDateTime(txt_exit.Text);
                if (arrival < exit)
                {
                    TimeSpan Sonuc = exit - arrival;
                    txt_time.Text = Sonuc.TotalHours.ToString() + " saat";
                    command.Parameters.AddWithValue("@ServisteKaldigiSüre", txt_time.Text);

                }
                command.ExecuteNonQuery();
                MessageBox.Show("Kayıt İşlemi Gerçekleşti.");

                txt_date.Clear();
                txt_plaque.Clear();
                txt_km.Clear();
                txt_unit.Clear();
                txt_brand.Clear();
                txt_model.Clear();
                txt_project.Clear();
                txt_arrival.Clear();
                txt_exit.Clear();
                txt_driver.Clear();
                txt_fault.Clear();
                txt_works.Clear();
                txt_material.Clear();
                txt_cost.Clear();
                txt_note.Clear();
                txt_maintainer.Clear();
                txt_deliveryarea.Clear();
                txt_recording.Clear();
                txt_time.Clear();
                cb_periodic.Checked = false;
                cb_wintercare.Checked = false;
                cb_summercare.Checked = false;
                cb_fault.Checked = false;
                cb_tire.Checked = false;
                cb_oilcare.Checked = false;
                rdb_user.Checked = false;
                rdb_fuel.Checked = false;
                rdb_way.Checked = false;
                rdb_other.Checked = false;
            }
            catch
            {
                DateTime arrival = Convert.ToDateTime(txt_arrival.Text);
                DateTime exit = Convert.ToDateTime(txt_exit.Text);
                if (arrival>exit)
                {
                    MessageBox.Show("Servise Geliş Tarihi, Servisten Çıkış Tarihinden İleri Bir Tarih Olamaz.");
                }                
            }
        }
        private void button1_Click(object sender, EventArgs e)  //SAVE BUTTON
        {
            try
            { 
                control();
            }
            catch
            { 
                MessageBox.Show("İşlem Sırasında Hata Oluştu.Doldurulması Zorunlu Alanları Kontrol Ediniz.");
            }
        }   


        private void button2_Click_1(object sender, EventArgs e) //datagridview LİSTS
        {
            lists();
        }

        void lists()
        {
            SqlConnection connection = new SqlConnection(bgl.Address);
            connection.Open();
            string record = "SELECT * from ServisBakimOnarim";
            SqlCommand command = new SqlCommand(record, connection);
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            connection.Close();
        }

        private void button3_Click_1(object sender, EventArgs e)   //excel creating
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            int StartCol = 1;

            int StartRow = 1;

            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)
                    sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++; for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)
                            sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    }
                    catch
                    {
                        ;
                    }
                }
            }
        }

        private void button4_Click_1(object sender, EventArgs e)   //plaque search
        {
            SqlConnection baglanti = new SqlConnection(bgl.Address);
            baglanti.Open();
            DataTable tbl = new DataTable();
            string search, data;
            search = textBox1.Text;
            data = "Select * from ServisBakimOnarim where Plaka like '%" + textBox1.Text + "%'";
            SqlDataAdapter adptr = new SqlDataAdapter(data, baglanti);
            adptr.Fill(tbl);
            baglanti.Close();
            dataGridView1.DataSource = tbl;
        }

        private void button5_Click(object sender, EventArgs e)  //driver search
        {
            SqlConnection connection = new SqlConnection(bgl.Address);
            connection.Open();
            DataTable tbl = new DataTable();
            string search, data;
            search = textBox1.Text;
            data = "Select * from ServisBakimOnarim where Sofor like '%" + textBox1.Text + "%'";
            SqlDataAdapter adptr = new SqlDataAdapter(data, connection);
            adptr.Fill(tbl);
            connection.Close();
            dataGridView1.DataSource = tbl;
        }

        private void button6_Click(object sender, EventArgs e)  //delete line
        {
            DialogResult selection = new DialogResult();
            selection = MessageBox.Show("Seçili Satırı Silmek İstediğinizden Emin Misiniz", "SERVİS BAKIM ONARIM FORMLARI", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

            if (selection == DialogResult.Yes)
            {
                foreach (DataGridViewRow drow in dataGridView1.SelectedRows)  
                {
                    int No = Convert.ToInt32(drow.Cells[0].Value);
                    DeleteRecord(No);
                }
                MessageBox.Show("Kayıt Başarıyla Silindi.");
            }

            lists();
        }

        void DeleteRecord(int No)
        {
            SqlConnection connection = new SqlConnection(bgl.Address);
            string sql = "DELETE FROM ServisBakimOnarim WHERE No=@No";
            SqlCommand command = new SqlCommand(sql, connection);
            command.Parameters.AddWithValue("@No", No);
            connection.Open();
            command.ExecuteNonQuery();
            connection.Close();
        }

        DataTable refresh()   //refresh line
        {
            SqlConnection connection = new SqlConnection(bgl.Address);
            connection.Open();
            SqlDataAdapter da = new SqlDataAdapter("select * from ServisBakimOnarim", connection);
            DataTable table = new DataTable();
            da.Fill(table);
            connection.Close();
            return table;
        }
  
        private void button7_Click(object sender, EventArgs e)  //refresh
        {
            SqlConnection connection = new SqlConnection(bgl.Address);
            DialogResult selection = new DialogResult();
            selection = MessageBox.Show("Seçili Satırı Güncellemek İstediğinizden Emin Misiniz", "SERVİS BAKIM ONARIM FORMLARI", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

            if (selection == DialogResult.Yes)
            {
                string no, date, plaque, km, unit, brand, model, project, arrival, exit, driver, typeofmaintenance, fault, faultoffailure, works, material, cost, time, note, maintainer, deliveryarea, recording;

                no = dataGridView1.CurrentRow.Cells["No"].Value.ToString();
                date = dataGridView1.CurrentRow.Cells["Tarih"].Value.ToString();
                plaque = dataGridView1.CurrentRow.Cells["Plaka"].Value.ToString();
                km = dataGridView1.CurrentRow.Cells["Km"].Value.ToString();
                unit = dataGridView1.CurrentRow.Cells["Birim"].Value.ToString();
                brand = dataGridView1.CurrentRow.Cells["Marka"].Value.ToString();
                model = dataGridView1.CurrentRow.Cells["Model"].Value.ToString();
                project = dataGridView1.CurrentRow.Cells["Proje"].Value.ToString();
                arrival = dataGridView1.CurrentRow.Cells["Gelis"].Value.ToString();
                exit = dataGridView1.CurrentRow.Cells["Cikis"].Value.ToString();
                driver = dataGridView1.CurrentRow.Cells["Sofor"].Value.ToString();
                typeofmaintenance = dataGridView1.CurrentRow.Cells["BakimOnarimTürü"].Value.ToString();
                fault = dataGridView1.CurrentRow.Cells["Ariza"].Value.ToString();
                faultoffailure = dataGridView1.CurrentRow.Cells["ArizaSebebi"].Value.ToString();
                works = dataGridView1.CurrentRow.Cells["Yapilanİsler"].Value.ToString();
                material = dataGridView1.CurrentRow.Cells["Malzeme"].Value.ToString();
                cost = dataGridView1.CurrentRow.Cells["BakimOnarimMaliyeti"].Value.ToString();
                time = dataGridView1.CurrentRow.Cells["ServisteKaldigiSüre"].Value.ToString();
                note = dataGridView1.CurrentRow.Cells["Aciklama"].Value.ToString();
                maintainer = dataGridView1.CurrentRow.Cells["BakimOnarimYapan"].Value.ToString();
                deliveryarea = dataGridView1.CurrentRow.Cells["TeslimAlan"].Value.ToString();
                recording = dataGridView1.CurrentRow.Cells["KayitYapan"].Value.ToString();
                connection.Open();
                SqlCommand command = new SqlCommand("update ServisBakimOnarim set Tarih='" + date + "',Plaka='" + plaque + "',Km='" + km + "',Birim='" + unit + "',Marka='" + brand + "',Model='" + model + "',Proje='" + project + "',Gelis='" + arrival + "',Cikis='" + exit + "',Sofor='" + driver + "',BakimOnarimTürü='" + typeofmaintenance + "',Ariza='" + fault + "',ArizaSebebi='" + faultoffailure + "',Yapilanİsler='" + works + "',Malzeme='" + material + "',BakimOnarimMaliyeti='" + cost + "',ServisteKaldigiSüre='" + time + "',Aciklama='" + note + "',BakimOnarimYapan='" + maintainer + "',TeslimAlan='" + deliveryarea + "',KayitYapan='" + recording +"' where No='" + no + "'", connection);
                command.ExecuteNonQuery();
                connection.Close();
                dataGridView1.DataSource = refresh();
                MessageBox.Show("Güncelleme İşlemi Gerçekleştirildi.");
            }
        }

        private void button8_Click(object sender, EventArgs e) //date search
        {
            SqlConnection connection = new SqlConnection(bgl.Address);
            connection.Open();
            DataTable tbl = new DataTable();
            string search, data;
            search = textBox1.Text;
            data = "Select * from ServisBakimOnarim where Tarih like '%" + textBox1.Text + "%'";
            SqlDataAdapter adptr = new SqlDataAdapter(data, connection);
            adptr.Fill(tbl);
            connection.Close();
            dataGridView1.DataSource = tbl;
        }

        private void txt_km_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void txt_cost_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void txt_driver_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar)
                 && !char.IsSeparator(e.KeyChar);
        }

        private void txt_maintainer_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar)
                 && !char.IsSeparator(e.KeyChar);
        }

        private void txt_deliveryarea_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar)
                 && !char.IsSeparator(e.KeyChar);
        }

        private void txt_recording_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar)
                 && !char.IsSeparator(e.KeyChar);
        }

        private void txt_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar)
                 && !char.IsSeparator(e.KeyChar);
        }
    }
}
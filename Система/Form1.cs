using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using MySql.Data.MySqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Система
{
    public partial class Form1 : Form
    {
        private Application application;
        private Workbook workBook;
        private Worksheet worksheet;        
        MySqlConnection sqlConnection;
        MySqlDataReader sqlReader;
        Dictionary<int, string> list = new Dictionary<int, string>();
        Dictionary<int, string> list1 = new Dictionary<int, string>();
        int count1;
        int count2;
        int count3;
        int count4;
        int count5;
        double[] mass = new double[5];
        double[] mass1, massv1, massv2, massv3, massv4;
        double naveskaCount;
        bool change = true;
        public Form1()
        {
            InitializeComponent();
        }
        private async void Form1_Load(object sender, EventArgs e)
        {
            string connectionString = "server=localhost;user=root;database=systems;password=Root;";

            sqlConnection = new MySqlConnection(connectionString);
            
            await sqlConnection.OpenAsync();

            getList();
            getList2();
            comboBox1.SelectedIndex = 0;
            tabPage3.Text = comboBox1.Text;
            comboBox2.SelectedIndex = 0;
            tabPage1.Text = comboBox2.Text;

            addDatagrid1();
            addDatagrid4();
            addDatagrid2();           
            addDatagrid5();
            addDatagrid3();
                        
            forReadOnly();
            KeyPreview = true;

            grap();
       
        }
        private void addToExcel(DataGridView data, int columnCount, int number)
        {
            worksheet = (Worksheet)application.Worksheets[number];
            worksheet.Activate();

            string str = "ABCDEFGHIJKLMNOP";

            for (int i = 1; i < columnCount; i++)
            {
                worksheet.Range[str[i - 1] + "1"].Value = data.Columns[i].HeaderText;
                for (int j = 0; j < data.Rows.Count; j++)
                {
                    worksheet.Range[str[i - 1] + "" + (j + 2)].Value = data[i, j].Value;
                }
            }
            worksheet.Columns.AutoFit();
        }
        private void CloseExcel()
        {
            try
            {
                if (application != null)
                {
                    int excelProcessId = -1;
                    GetWindowThreadProcessId(application.Hwnd, ref excelProcessId);

                    Marshal.ReleaseComObject(worksheet);
                    workBook.Close();
                    Marshal.ReleaseComObject(workBook);
                    application.Quit();
                    Marshal.ReleaseComObject(application);

                    application = null;

                    try
                    {
                        Process process = Process.GetProcessById(excelProcessId);
                        process.Kill();
                    }
                    finally { }
                }
            }
            catch { }
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(int hWnd, ref int lpdwProcessId);

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
            CloseExcel();
        }
        private void forReadOnly()
        {
            for(int i =0;i<5;i++)
            dataGridView1[3, i].ReadOnly = true;
        }
        private void updateDB1(DataGridView data1)
        {
            change = false;
            int k = 0;  

            foreach (KeyValuePair<int, string> kvp in list1)
            {
                if (comboBox2.SelectedItem.Equals(kvp.Value))
                    k = kvp.Key;
            }

            MySqlCommand command1 = new MySqlCommand("DELETE FROM `состав шихты` WHERE `Код шихты` = " + k, sqlConnection);
            command1.Parameters.Clear();
            command1.ExecuteNonQuery();


            MySqlCommand command = new MySqlCommand("INSERT INTO `Состав шихты` (`Код состава`, Компонент, `Массовая доля`, Навеска, `Порядок добавления`, Комментарий, `Код шихты`) VALUES (@Код_состава,@Компонент,@Массовая_доля,@Навеска,@Порядок_добавления,@Комментарий, @Код_шихты)", sqlConnection);
           
            for (int i = 0; i < data1.Rows.Count-1; ++i)
            {
                command.Parameters.Clear();
                try
                {
                    int count = data1.Rows.Count;
                    try
                    {
                        command.Parameters.AddWithValue("Код_состава", data1[0, i].Value.ToString());
                    }
                    catch
                    {
                        command.Parameters.AddWithValue("Код_состава", ++count1);
                    }
                    command.Parameters.AddWithValue("Компонент", data1[1, i].Value.ToString());
                    data1[2, i].Value = data1[2, i].Value.ToString().Replace(",", ".");
                    command.Parameters.AddWithValue("Массовая_доля", data1[2, i].Value.ToString());
                    data1[3, i].Value = data1[3, i].Value.ToString().Replace(",", ".");
                    command.Parameters.AddWithValue("Навеска", data1[3, i].Value.ToString());
                    command.Parameters.AddWithValue("Порядок_добавления", data1[4, i].Value.ToString());
                    try
                    {
                        command.Parameters.AddWithValue("Комментарий", data1[5, i].Value.ToString());
                    }
                    catch
                    {
                        command.Parameters.AddWithValue("Комментарий", " ");
                    }
                    command.Parameters.AddWithValue("Код_шихты", k);
                    command.ExecuteNonQuery();
                }
                catch
                {
                    command.Parameters.Clear();
                    command.Parameters.AddWithValue("Код_состава", ++count1);                    
                    command.Parameters.AddWithValue("Компонент", data1[1, i].Value.ToString());
                    data1[2, i].Value = data1[2, i].Value.ToString().Replace(",", ".");
                    command.Parameters.AddWithValue("Массовая_доля", data1[2, i].Value.ToString());
                    data1[3, i].Value = data1[3, i].Value.ToString().Replace(",", ".");
                    command.Parameters.AddWithValue("Навеска", data1[3, i].Value.ToString());
                    command.Parameters.AddWithValue("Порядок_добавления", data1[4, i].Value.ToString());
                    try
                    {
                        command.Parameters.AddWithValue("Комментарий", data1[5, i].Value.ToString());
                    }
                    catch
                    {
                        command.Parameters.AddWithValue("Комментарий", " ");
                    }
                    command.Parameters.AddWithValue("Код_шихты", k);
                    command.ExecuteNonQuery();
                }
            }
            change = true;
        }
        private void updateDB2(DataGridView data1)
        {

            MySqlCommand command = new MySqlCommand("INSERT INTO `Вид образцов`(`Код образца`, Название, Размер) VALUES(@Код_образца, @Название, @Размер) ON DUPLICATE KEY UPDATE Название = @Название, Размер = @Размер", sqlConnection);
                   
            for (int i = 0; i < data1.Rows.Count - 1; ++i)
            {
                command.Parameters.Clear();
                try
                {
                    command.Parameters.AddWithValue("Код_образца", data1[0, i].Value.ToString());
                }
                catch
                {
                    command.Parameters.AddWithValue("Код_образца", ++count4);
                    data1[0, i].Value = count4;
                }
                command.Parameters.AddWithValue("Название", data1[1, i].Value.ToString());
                command.Parameters.AddWithValue("Размер", Convert.ToDouble(data1[2, i].Value));

                command.ExecuteNonQuery();
            }
            
        }
        private void updateDB3(DataGridView data1)
        {  
            int k = 0;          

            foreach (KeyValuePair<int, string> kvp in list)
            {
                if (comboBox1.SelectedItem.Equals(kvp.Value))
                    k = kvp.Key;
            }
           
            MySqlCommand command1 = new MySqlCommand("DELETE FROM спекание WHERE `Код спекания` = " + k, sqlConnection);
            command1.Parameters.Clear();
            command1.ExecuteNonQuery();

            MySqlCommand command = new MySqlCommand("INSERT INTO спекание (`Код операции`, `Точка выдержки`, `Скорость нагрева`, `Начальная температура`, `Конечная температура`, `Код спекания`) VALUES (@Код_операции, @Точка_выдержки, @Скорость_нагрева, @Начальная_температура, @Конечная_температура, @Код_спекания)", sqlConnection);

            for (int i = 0; i < data1.Rows.Count - 1; ++i)
            {
                command.Parameters.Clear();
                try
                {
                    try
                    {
                        command.Parameters.AddWithValue("Код_операции", data1[0, i].Value.ToString());
                    }
                    catch
                    {
                        command.Parameters.AddWithValue("Код_операции", ++count5);
                    }
                    command.Parameters.AddWithValue("Точка_выдержки", data1[1, i].Value.ToString());
                    command.Parameters.AddWithValue("Скорость_нагрева", data1[2, i].Value.ToString());
                    command.Parameters.AddWithValue("Начальная_температура", data1[3, i].Value.ToString());
                    command.Parameters.AddWithValue("Конечная_температура", data1[4, i].Value.ToString());
                    command.Parameters.AddWithValue("Код_спекания", k);
                    command.ExecuteNonQuery();
                }
                catch
                {
                    command.Parameters.Clear(); 
                    command.Parameters.AddWithValue("Код_операции", ++count5);
                    command.Parameters.AddWithValue("Точка_выдержки", data1[1, i].Value.ToString());
                    command.Parameters.AddWithValue("Скорость_нагрева", data1[2, i].Value.ToString());
                    command.Parameters.AddWithValue("Начальная_температура", data1[3, i].Value.ToString());
                    command.Parameters.AddWithValue("Конечная_температура", data1[4, i].Value.ToString());
                    command.Parameters.AddWithValue("Код_спекания", k);
                    command.ExecuteNonQuery();
                }
            }
        }
        private void updateDB4(DataGridView data1)
        {


            MySqlCommand command = new MySqlCommand("INSERT INTO `Режимы прессования` (`Код рецепта`, `Рабочее давление`, `Время выдержки`, `Способ прессования`) VALUES (@Код_рецепта, @Рабочее_давление, @Время_выдержки, @Способ_прессования) ON DUPLICATE KEY UPDATE `Рабочее давление`=@Рабочее_давление, `Время выдержки`= @Время_выдержки, `Способ прессования`= @Способ_прессования", sqlConnection);

            for (int i = 0; i < data1.Rows.Count - 1; ++i)
            {
                command.Parameters.Clear();
                try
                {
                    command.Parameters.AddWithValue("Код_рецепта", data1[0, i].Value.ToString());
                }
                catch
                {
                    command.Parameters.AddWithValue("Код_рецепта", ++count2);
                    data1[0, i].Value = count2;
                }
                command.Parameters.AddWithValue("Рабочее_давление", data1[1, i].Value.ToString());
                command.Parameters.AddWithValue("Время_выдержки", data1[2, i].Value.ToString());
                try
                {
                    command.Parameters.AddWithValue("Способ_прессования", data1[3, i].Value.ToString());
                }
                catch
                {
                    command.Parameters.AddWithValue("Способ_прессования", "");
                }
                command.ExecuteNonQuery();
                
            }
        }
        private void updateDB5(DataGridView data1)
        {
            MySqlCommand command = new MySqlCommand("INSERT INTO журнал (`Код записи`, Дата, Время, `Код операции`, `Код шихты`, `Код образца`, `Код рецепта`, Комментарий) VALUES (@Код_записи, @Дата, @Время, @Код_операции, @Код_шихты, @Код_образца, @Код_рецепта, @Комментарий) " +
                                                    "ON DUPLICATE KEY UPDATE Дата = @Дата, Время = @Время, `Код операции` = @Код_операции, `Код шихты` = @Код_шихты, `Код образца` = @Код_образца, `Код рецепта` = @Код_рецепта, Комментарий = @Комментарий", sqlConnection);

            for (int i = 0; i < data1.Rows.Count - 1; ++i)
            {
                command.Parameters.Clear();
                try
                {
                    command.Parameters.AddWithValue("Код_записи", data1[0, i].Value.ToString());
                }
                catch
                {
                    command.Parameters.AddWithValue("Код_записи", ++count3);
                    data1[0, i].Value = count3;
                }
               
                string rem = data1[1, i].Value.ToString().Remove(5, 5);
                string date = data1[1, i].Value.ToString().Remove(0, 6) + "-" + rem.Remove(0, 3) + "-" + rem.Remove(3, 2);

                command.Parameters.AddWithValue("Дата", date);
                command.Parameters.AddWithValue("Время", data1[2, i].Value.ToString());
                command.Parameters.AddWithValue("Код_операции", data1[11, i].Value.ToString());
                command.Parameters.AddWithValue("Код_шихты", data1[12, i].Value.ToString());
                command.Parameters.AddWithValue("Код_образца", data1[13, i].Value.ToString());
                command.Parameters.AddWithValue("Код_рецепта", data1[14, i].Value.ToString());
                try
                {
                    command.Parameters.AddWithValue("Комментарий", data1[10, i].Value.ToString());
                }
                catch
                {
                    command.Parameters.AddWithValue("Комментарий", "");
                }
                command.ExecuteNonQuery();
            }
        }           
        private async void addDatagrid1()
        {
            change = false;
            int k = 0;
            foreach (KeyValuePair<int, string> kvp in list1)
            {
                if (comboBox2.SelectedItem.Equals(kvp.Value))
                    k = kvp.Key;
            }

            MySqlCommand command1 = new MySqlCommand("SELECT `Код состава` FROM `состав шихты`", sqlConnection);

            sqlReader = null;
            try
            {
                sqlReader = command1.ExecuteReader();

                while (await sqlReader.ReadAsync())
                {
                    count1 = Convert.ToInt32(sqlReader["Код состава"]);
                }
            }
            catch
            {

            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }

            sqlReader = null;
            
            change = true;
        }
        private async void addDatagrid2()
        {

            MySqlCommand command = new MySqlCommand("SELECT * FROM `Режимы прессования`", sqlConnection);

            List <string[]> data = new List<string[]>();

            sqlReader = null;
                                        
            try
            {                
                sqlReader = command.ExecuteReader();

                while (await sqlReader.ReadAsync())
                {
                    data.Add(new string[5]);

                    data[data.Count - 1][0] = Convert.ToString(sqlReader["Код рецепта"]); 
                    data[data.Count - 1][1] = Convert.ToString(sqlReader["Рабочее давление"]);
                    data[data.Count - 1][2] = Convert.ToString(sqlReader["Время выдержки"]);
                    data[data.Count - 1][3] = Convert.ToString(sqlReader["Способ прессования"]);
                }                 
                foreach (string[] s in data)
                    dataGridView2.Rows.Add(s);

                int rowsCount = dataGridView2.Rows.Count - 2;
                count2 = Convert.ToInt32(dataGridView2[0, rowsCount].Value);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();

            }
        }
        private async void addDatagrid4()
        {
            MySqlCommand command = new MySqlCommand("SELECT * FROM `Вид образцов`", sqlConnection);
           
            List<string[]> data = new List<string[]>();

            sqlReader = null;

            try
            {                
                sqlReader = command.ExecuteReader();

                while (await sqlReader.ReadAsync())
                {
                    data.Add(new string[3]);
                    data[data.Count - 1][0] = Convert.ToString(sqlReader["Код образца"]);
                    data[data.Count - 1][1] = Convert.ToString(sqlReader["Название"]);
                    data[data.Count - 1][2] = Convert.ToString(sqlReader["Размер"]);
                }

                foreach (string[] s in data)
                {
                    dataGridView4.Rows.Add(s);
                }

                int rowsCount = dataGridView4.Rows.Count - 2;
                count4 = Convert.ToInt32(dataGridView4[0, rowsCount].Value);

            }                
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();

            }
            mass1 = new double[dataGridView4.Rows.Count];
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                mass1[i] = Convert.ToDouble(dataGridView4[2, i].Value);
            }
            
        }       
        private async void addDatagrid5()
        {

            int k = 0;
            foreach (KeyValuePair<int, string> kvp in list)
            {
                if (comboBox1.SelectedItem.Equals(kvp.Value))
                    k = kvp.Key;
            }

            MySqlCommand command1 = new MySqlCommand("SELECT `Код операции` FROM спекание", sqlConnection);

            sqlReader = null;
            try
            {
                sqlReader = command1.ExecuteReader();

                while (await sqlReader.ReadAsync())
                {
                    count5 = Convert.ToInt32(sqlReader["Код операции"]);
                }
            }
            catch
            {

            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
                        
            MySqlCommand command = new MySqlCommand("SELECT * FROM спекание WHERE `Код спекания` = " + k, sqlConnection);

            List<string[]> data = new List<string[]>();
            dataGridView5.Rows.Clear();
            sqlReader = null;

            try
            {
                sqlReader = command.ExecuteReader();               

                while (await sqlReader.ReadAsync())
                {
                    data.Add(new string[5]);
                    data[data.Count - 1][0] = Convert.ToString(sqlReader["Код операции"]);
                    data[data.Count - 1][1] = Convert.ToString(sqlReader["Точка выдержки"]);
                    data[data.Count - 1][2] = Convert.ToString(sqlReader["Скорость нагрева"]);
                    data[data.Count - 1][3] = Convert.ToString(sqlReader["Начальная температура"]);
                    data[data.Count - 1][4] = Convert.ToString(sqlReader["Конечная температура"]);
                }

                foreach (string[] s in data)
                    dataGridView5.Rows.Add(s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();

            }
        }
        private async void addDatagrid3()
        {
            MySqlCommand command1 = new MySqlCommand("SELECT * FROM `список спеканий`;", sqlConnection);
            try
            {
                sqlReader = command1.ExecuteReader();
                спекание_журнал.Items.Clear();
                while (sqlReader.Read())
                {
                    спекание_журнал.Items.Add(Convert.ToString(sqlReader["Имя спекания"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }

            MySqlCommand command2 = new MySqlCommand("SELECT * FROM `список шихты`;", sqlConnection);
            try
            {
                sqlReader = command2.ExecuteReader();
                Шихта_журнал.Items.Clear();
                while (sqlReader.Read())
                {
                    Шихта_журнал.Items.Add(Convert.ToString(sqlReader["Название состава"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            MySqlCommand command = new MySqlCommand("SELECT * FROM журнал, `вид образцов`, `список шихты`, `режимы прессования`, `список спеканий` WHERE `журнал`.`Код образца`=`вид образцов`.`Код образца` AND `журнал`.`Код шихты`=`список шихты`.`Код шихты` AND `журнал`.`Код рецепта` = `режимы прессования`.`Код рецепта` AND `журнал`.`Код операции` = `список спеканий`.`Код спекания`;", sqlConnection);

            List<string[]> data = new List<string[]>();

            sqlReader = null;

            try
            {
                sqlReader = command.ExecuteReader();

                Вид_образца_журнал.Items.Clear();
                for (int i = 0; i < dataGridView4.Rows.Count - 1; ++i)
                {
                    Вид_образца_журнал.Items.Add(dataGridView4[1, i].Value.ToString());                    
                }

                Рабочее_давление_журнал.Items.Clear();
                for (int i = 0; i < dataGridView2.Rows.Count - 1; ++i)
                {
                    Рабочее_давление_журнал.Items.Add(dataGridView2[1, i].Value.ToString());
                }

                while (await sqlReader.ReadAsync())
                {
                    data.Add(new string[15]);
                    data[data.Count - 1][0] = Convert.ToString(sqlReader["Код записи"]);
                    data[data.Count - 1][1] = Convert.ToString(sqlReader["Дата"]).Remove(10,8);
                    data[data.Count - 1][2] = Convert.ToString(sqlReader["Время"]);
                    data[data.Count - 1][3] = Convert.ToString(sqlReader["Название"]);
                    data[data.Count - 1][4] = Convert.ToString(sqlReader["Размер"]);
                    data[data.Count - 1][5] = Convert.ToString(sqlReader["Название состава"]);                    
                    data[data.Count - 1][6] = Convert.ToString(sqlReader["Имя спекания"]);
                    data[data.Count - 1][7] = Convert.ToString(sqlReader["Рабочее давление"]);
                    data[data.Count - 1][8] = Convert.ToString(sqlReader["Время выдержки"]);
                    data[data.Count - 1][9] = Convert.ToString(sqlReader["Способ прессования"]);
                    data[data.Count - 1][10] = Convert.ToString(sqlReader["Комментарий"]);

                    data[data.Count - 1][11] = Convert.ToString(sqlReader["Код операции"]);
                    data[data.Count - 1][12] = Convert.ToString(sqlReader["Код шихты"]);
                    data[data.Count - 1][13] = Convert.ToString(sqlReader["Код образца"]);
                    data[data.Count - 1][14] = Convert.ToString(sqlReader["Код рецепта"]);
                }


                foreach (string[] s in data)
                    dataGridView3.Rows.Add(s);              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }

            MySqlCommand command3 = new MySqlCommand("SELECT * FROM журнал", sqlConnection);

            sqlReader = null;
            try
            {
                sqlReader = command3.ExecuteReader();

                while (sqlReader.Read())
                {
                    if (count3 < Convert.ToInt32(sqlReader["Код записи"]))
                        count3 = Convert.ToInt32(sqlReader["Код записи"]);
                }
            }
            catch
            {

            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }
        private void getList()
        {
            MySqlCommand command = new MySqlCommand("SELECT * FROM `список спеканий`;", sqlConnection);

            list.Clear();

            try
            {
                sqlReader = command.ExecuteReader();
                while (sqlReader.Read())
                {
                    list.Add(Convert.ToInt32(sqlReader["Код спекания"]), Convert.ToString(sqlReader["Имя спекания"]));
                }
                comboBox1.Items.Clear();
                Dictionary<int, string>.ValueCollection valueColl = list.Values;
                foreach (string s in valueColl)
                {
                    comboBox1.Items.Add(s);
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }
        private void getList2()
        {
            MySqlCommand command = new MySqlCommand("SELECT * FROM `список шихты`;", sqlConnection);

            list1.Clear();

            try
            {
                sqlReader = command.ExecuteReader();
                while (sqlReader.Read())
                {
                    list1.Add(Convert.ToInt32(sqlReader["Код шихты"]), Convert.ToString(sqlReader["Название состава"]));
                }
                comboBox2.Items.Clear();
                Dictionary<int, string>.ValueCollection valueColl = list1.Values;
                foreach (string s in valueColl)
                {
                    comboBox2.Items.Add(s);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(dataGridView5);
            form2.Show();
        }
        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 3) 
                for (int i = 0; i < dataGridView4.Rows.Count - 1; ++i)
                {
                    if ((dataGridView3[e.ColumnIndex, e.RowIndex].Value).Equals(dataGridView4[1, i].Value))
                    {
                            dataGridView3[4, e.RowIndex].Value = dataGridView4[2, i].Value;
                            dataGridView3[13, e.RowIndex].Value = dataGridView4[0, i].Value;
                            break;
                    }
                }
                if (e.ColumnIndex == 5)
                    foreach (KeyValuePair<int, string> kvp in list1)
                    {
                        if ((dataGridView3[e.ColumnIndex, e.RowIndex].Value).Equals(kvp.Value))
                        {                           
                            dataGridView3[12, e.RowIndex].Value = kvp.Key;
                            break;
                        }
                    }
                if (e.ColumnIndex == 6)
                    foreach (KeyValuePair<int, string> kvp in list)
                    {
                        if ((dataGridView3[e.ColumnIndex, e.RowIndex].Value).Equals(kvp.Value))
                            dataGridView3[11, e.RowIndex].Value = kvp.Key;
                    }
                if (e.ColumnIndex == 7)
                    for (int i = 0; i < dataGridView2.Rows.Count - 1; ++i)
                    {
                        if ((dataGridView3[e.ColumnIndex, e.RowIndex].Value).Equals(dataGridView2[1, i].Value))
                        {
                            dataGridView3[8, e.RowIndex].Value = dataGridView2[2, i].Value;
                            dataGridView3[9, e.RowIndex].Value = dataGridView2[3, i].Value;
                            dataGridView3[14, e.RowIndex].Value = dataGridView2[0, i].Value;
                            break;
                        }
                    }
            }
            catch { }
        }
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                updateDB1(dataGridView1);
                updateDB2(dataGridView4);
                updateDB4(dataGridView2);
                updateDB3(dataGridView5);
                updateDB5(dataGridView3);
                dataGridView3.Rows.Clear();
                addDatagrid3();
            }
            catch(MySqlException) { MessageBox.Show("Введен неправильный формат времени","Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (ArgumentOutOfRangeException) { MessageBox.Show("Введен неправильный формат даты", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception) { MessageBox.Show("Данные имеют не верный формат.","Ошибка БД", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 1)
                    dataGridView3[1, e.RowIndex].Value = DateTime.Today.ToString("d");

                if (e.ColumnIndex == 2)
                    dataGridView3[2, e.RowIndex].Value = DateTime.Now.ToLongTimeString();
            }
            catch { }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int k = 0;
            foreach (KeyValuePair<int, string> kvp in list)
            {
                if (comboBox1.SelectedItem.Equals(kvp.Value))
                    k = kvp.Key;
            }

            MySqlCommand command = new MySqlCommand("SELECT * FROM спекание WHERE `Код спекания` = " + k, sqlConnection);

            List<string[]> data = new List<string[]>();

            sqlReader.Close();
            sqlReader = null;

            dataGridView5.Rows.Clear();
            
            try
            {
                sqlReader = command.ExecuteReader();

                while (sqlReader.Read())
                {
                    data.Add(new string[5]);
                    data[data.Count - 1][0] = Convert.ToString(sqlReader["Код операции"]);
                    data[data.Count - 1][1] = Convert.ToString(sqlReader["Точка выдержки"]);
                    data[data.Count - 1][2] = Convert.ToString(sqlReader["Скорость нагрева"]);
                    data[data.Count - 1][3] = Convert.ToString(sqlReader["Начальная температура"]);
                    data[data.Count - 1][4] = Convert.ToString(sqlReader["Конечная температура"]);
                }

                foreach (string[] s in data)
                    dataGridView5.Rows.Add(s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();

            }

            tabPage3.Text = comboBox1.Text;

            massv1 = new double[1000];
            massv2 = new double[1000];
            massv3 = new double[1000];
            massv4 = new double[1000];

            for (int i=0; i< dataGridView5.Rows.Count - 1; i++)
            {
                massv1[i] = Convert.ToDouble(dataGridView5[1, i].Value);
                massv2[i] = Convert.ToDouble(dataGridView5[2, i].Value);
                massv3[i] = Convert.ToDouble(dataGridView5[3, i].Value);
                massv4[i] = Convert.ToDouble(dataGridView5[4, i].Value);
            }
            grap();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            updateDB3(dataGridView5);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                bool l = false;
                if (comboBox1.Items.Contains(textBox1.Text))
                {
                    l = true;
                }
                if (l == false)
                {
                    MySqlCommand command1 = new MySqlCommand("INSERT INTO `список спеканий` (`Код спекания`, `Имя спекания`) VALUES (@Код_спекания, @Имя_спекания)", sqlConnection);

                    sqlReader = null;

                    bool b = true;
                    int k = 0;
                    foreach (KeyValuePair<int, string> kvp in list)
                    {
                        if (k < kvp.Key)
                            k = kvp.Key;
                        if (textBox1.Text.Equals(kvp.Value))
                        {
                            b = false;
                        }
                    }

                    if (b)
                    {
                        try
                        {
                            command1.Parameters.Clear();
                            command1.Parameters.AddWithValue("Код_спекания", ++k);
                            command1.Parameters.AddWithValue("Имя_спекания", textBox1.Text);
                            command1.ExecuteNonQuery();

                            MySqlCommand command = new MySqlCommand("INSERT INTO спекание (`Код операции`, `Точка выдержки`, `Скорость нагрева`, `Начальная температура`, `Конечная температура`, `Код спекания`) VALUES (@Код_операции, @Точка_выдержки, @Скорость_нагрева, @Начальная_температура, @Конечная_температура, @Код_спекания)", sqlConnection);

                            for (int i = 0; i < dataGridView5.Rows.Count - 1; ++i)
                            {
                                command.Parameters.Clear();
                                try
                                {
                                    try
                                    {
                                        command.Parameters.AddWithValue("Код_операции", dataGridView5[0, i].Value.ToString());
                                    }
                                    catch
                                    {
                                        command.Parameters.AddWithValue("Код_операции", ++count5);
                                    }
                                    command.Parameters.AddWithValue("Точка_выдержки", dataGridView5[1, i].Value.ToString());
                                    command.Parameters.AddWithValue("Скорость_нагрева", dataGridView5[2, i].Value.ToString());
                                    command.Parameters.AddWithValue("Начальная_температура", dataGridView5[3, i].Value.ToString());
                                    command.Parameters.AddWithValue("Конечная_температура", dataGridView5[4, i].Value.ToString());
                                    command.Parameters.AddWithValue("Код_спекания", k);
                                    command.ExecuteNonQuery();
                                }
                                catch
                                {
                                    command.Parameters.Clear();
                                    command.Parameters.AddWithValue("Код_операции", ++count5);
                                    command.Parameters.AddWithValue("Точка_выдержки", dataGridView5[1, i].Value.ToString());
                                    command.Parameters.AddWithValue("Скорость_нагрева", dataGridView5[2, i].Value.ToString());
                                    command.Parameters.AddWithValue("Начальная_температура", dataGridView5[3, i].Value.ToString());
                                    command.Parameters.AddWithValue("Конечная_температура", dataGridView5[4, i].Value.ToString());
                                    command.Parameters.AddWithValue("Код_спекания", k);
                                    command.ExecuteNonQuery();
                                }
                            }
                            comboBox1.Items.Clear();
                            getList();
                            comboBox1.Text = textBox1.Text;

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else MessageBox.Show("Такое название уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else MessageBox.Show("Введите название спекания.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private void сохранитьВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            application = new Application
            {
                DisplayAlerts = false
            };

            workBook = application.Workbooks.Add(Type.Missing);

            addToExcel(dataGridView1, 6, 1);
            addToExcel(dataGridView4, 3, 2);
            addToExcel(dataGridView5, 5, 3);
            addToExcel(dataGridView2, 4, 4);
            addToExcel(dataGridView3, 11, 5);

            application.Visible = true;
            TopMost = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            updateDB1(dataGridView1);
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            change = false;
            int k = 0;
            foreach (KeyValuePair<int, string> kvp in list1)
            {
                if (comboBox2.SelectedItem.Equals(kvp.Value))
                    k = kvp.Key;
            }

            MySqlCommand command = new MySqlCommand("SELECT * FROM `состав шихты` WHERE `Код шихты` = " + k, sqlConnection);

            List<string[]> data = new List<string[]>();

            sqlReader.Close();
            sqlReader = null;

            dataGridView1.Rows.Clear();

            try
            {
                sqlReader = command.ExecuteReader();

                while (sqlReader.Read())
                {
                    data.Add(new string[7]);
                    data[data.Count - 1][0] = Convert.ToString(sqlReader["Код состава"]);
                    data[data.Count - 1][1] = Convert.ToString(sqlReader["Компонент"]);
                    data[data.Count - 1][2] = Convert.ToString(sqlReader["Массовая доля"]);
                    data[data.Count - 1][3] = Convert.ToString(sqlReader["Навеска"]);
                    data[data.Count - 1][4] = Convert.ToString(sqlReader["Порядок добавления"]);
                    data[data.Count - 1][5] = Convert.ToString(sqlReader["Комментарий"]);
                    data[data.Count - 1][6] = Convert.ToString(sqlReader["Код шихты"]);
                }

                foreach (string[] s in data)
                    dataGridView1.Rows.Add(s);

                dataGridView1.Rows.Add(1);
                dataGridView1[2, 5].Value = "         Масса:";                
                dataGridView1[2, 5].ReadOnly = true;
                dataGridView1[3, 5].Value = Convert.ToDouble(dataGridView1[3, 0].Value) + Convert.ToDouble(dataGridView1[3,1].Value) + Convert.ToDouble(dataGridView1[3, 2].Value) + Convert.ToDouble(dataGridView1[3, 3].Value) + Convert.ToDouble(dataGridView1[3, 4].Value);
                naveskaCount = Convert.ToDouble(dataGridView1[3, 5].Value);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }

            tabPage1.Text = comboBox2.Text;

            change = true;

            for (int i = 0; i< 5; i++)
            {
                mass[i] = Convert.ToDouble(dataGridView1[2, i].Value);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                change = false;

                bool l = false;
                if (comboBox2.Items.Contains(textBox2.Text))
                {
                    l = true;
                }
                if (!l)
                {
                    MySqlCommand command1 = new MySqlCommand("INSERT INTO `список шихты` (`Код шихты`, `Название состава`) VALUES (@Код_шихты, @Название_состава)", sqlConnection);

                    sqlReader = null;

                    bool b = true;
                    int k = 0;
                    foreach (KeyValuePair<int, string> kvp in list1)
                    {
                        if (k < kvp.Key)
                            k = kvp.Key;
                        if (textBox2.Text.Equals(kvp.Value))
                        {
                            b = false;
                        }
                    }

                    if (b)
                    {
                        try
                        {
                            command1.Parameters.Clear();
                            command1.Parameters.AddWithValue("Код_шихты", ++k);
                            command1.Parameters.AddWithValue("Название_состава", textBox2.Text);
                            command1.ExecuteNonQuery();

                            MySqlCommand command = new MySqlCommand("INSERT INTO `Состав шихты` (`Код состава`, Компонент, `Массовая доля`, Навеска, `Порядок добавления`, Комментарий, `Код шихты`) VALUES (@Код_состава,@Компонент,@Массовая_доля,@Навеска,@Порядок_добавления,@Комментарий, @Код_шихты)", sqlConnection);

                            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                            {
                                command.Parameters.Clear();
                                try
                                {
                                    try
                                    {
                                        command.Parameters.AddWithValue("Код_состава", dataGridView1[0, i].Value.ToString());
                                    }
                                    catch
                                    {
                                        command.Parameters.AddWithValue("Код_состава", ++count1);
                                    }
                                    command.Parameters.AddWithValue("Компонент", dataGridView1[1, i].Value.ToString());
                                    dataGridView1[2, i].Value = dataGridView1[2, i].Value.ToString().Replace(",", ".");
                                    command.Parameters.AddWithValue("Массовая_доля", dataGridView1[2, i].Value.ToString());
                                    dataGridView1[3, i].Value = dataGridView1[3, i].Value.ToString().Replace(",", ".");
                                    command.Parameters.AddWithValue("Навеска", dataGridView1[3, i].Value.ToString());
                                    command.Parameters.AddWithValue("Порядок_добавления", dataGridView1[4, i].Value.ToString());
                                    try
                                    {
                                        command.Parameters.AddWithValue("Комментарий", dataGridView1[5, i].Value.ToString());
                                    }
                                    catch
                                    {
                                        command.Parameters.AddWithValue("Комментарий", " ");
                                    }
                                    command.Parameters.AddWithValue("Код_шихты", k);
                                    command.ExecuteNonQuery();

                                }
                                catch
                                {
                                    command.Parameters.Clear();
                                    command.Parameters.AddWithValue("Код_состава", ++count1);
                                    command.Parameters.AddWithValue("Компонент", dataGridView1[1, i].Value.ToString());
                                    dataGridView1[2, i].Value = dataGridView1[2, i].Value.ToString().Replace(",", ".");
                                    command.Parameters.AddWithValue("Массовая_доля", dataGridView1[2, i].Value.ToString());
                                    dataGridView1[3, i].Value = dataGridView1[3, i].Value.ToString().Replace(",", ".");
                                    command.Parameters.AddWithValue("Навеска", dataGridView1[3, i].Value.ToString());
                                    command.Parameters.AddWithValue("Порядок_добавления", dataGridView1[4, i].Value.ToString());
                                    try
                                    {
                                        command.Parameters.AddWithValue("Комментарий", dataGridView1[5, i].Value.ToString());
                                    }
                                    catch
                                    {
                                        command.Parameters.AddWithValue("Комментарий", " ");
                                    }
                                    command.Parameters.AddWithValue("Код_шихты", k);
                                    command.ExecuteNonQuery();
                                }
                            }
                            comboBox2.Items.Clear();
                            getList2();
                            comboBox2.Text = textBox2.Text;
                        }
                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        change = true;
                    }
                }
                else MessageBox.Show("Такое название уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else MessageBox.Show("Введите название состава шихты.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (change)
            {
                if (e.ColumnIndex == 3 && e.RowIndex == 5)
                {
                    try
                    {
                        naveskaCount = Convert.ToDouble(dataGridView1[3, 5].Value);

                        double r1, r2, r3, r4, r5;
                        r1 = Convert.ToDouble(dataGridView1[2, 0].Value);
                        r2 = Convert.ToDouble(dataGridView1[2, 1].Value);
                        r3 = Convert.ToDouble(dataGridView1[2, 2].Value);
                        r4 = Convert.ToDouble(dataGridView1[2, 3].Value);
                        r5 = Convert.ToDouble(dataGridView1[2, 4].Value);

                        for (int i = 0; i < 5; i++)
                        {
                            dataGridView1[3, i].Value = Convert.ToDouble(dataGridView1[3, 5].Value) * Convert.ToDouble(dataGridView1[2, i].Value);
                        }

                        if (r1 + r2 + r3 + r4 < 1)
                        {
                            r5 = 1 - (r1 + r2 + r3 + r4);
                            dataGridView1[2, 4].Value = r5;
                        }
                        else
                        {
                            dataGridView1[2, e.RowIndex].Value = mass[e.RowIndex];
                            dataGridView1[2, 4].Value = r5;
                        }

                        for (int i = 0; i < 5; i++)
                        {
                            mass[i] = Convert.ToDouble(dataGridView1[2, i].Value);
                        }
                    }
                    catch (FormatException ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        dataGridView1[3, 5].Value = naveskaCount;
                    }
                }

                if (e.ColumnIndex == 2)
                {
                    try
                    {
                        double r1, r2, r3, r4, r5;
                        r1 = Convert.ToDouble(dataGridView1[2, 0].Value);
                        r2 = Convert.ToDouble(dataGridView1[2, 1].Value);
                        r3 = Convert.ToDouble(dataGridView1[2, 2].Value);
                        r4 = Convert.ToDouble(dataGridView1[2, 3].Value);
                        r5 = Convert.ToDouble(dataGridView1[2, 4].Value);

                        if (Convert.ToDouble(dataGridView1[2, e.RowIndex].Value) >= 1)
                        {
                            dataGridView1[2, e.RowIndex].Value = mass[e.RowIndex];
                        }
                                                
                        if (r1 + r2 + r3 + r4 < 1)
                        {
                            r5 = 1 - (r1 + r2 + r3 + r4);
                            dataGridView1[2, 4].Value = r5;
                        }
                        else
                        {
                            dataGridView1[2, e.RowIndex].Value = mass[e.RowIndex];
                            dataGridView1[2, 4].Value = r5;
                        }

                        for (int i = 0;i<5; i++)
                        {
                            dataGridView1[3, i].Value = Convert.ToDouble(dataGridView1[3, 5].Value) * Convert.ToDouble(dataGridView1[2, i].Value);
                        }
                        for (int i = 0; i < 5; i++)
                        {
                            mass[i] = Convert.ToDouble(dataGridView1[2, i].Value);
                        }
                        
                    }
                    catch (FormatException ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (e.ColumnIndex == 2)
                            dataGridView1[2, e.RowIndex].Value = mass[e.RowIndex];                            
                    }
                    catch
                    {

                    }
                    
                }
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox2.Items.Count > 1)
            {
                int k = 0;
                foreach (KeyValuePair<int, string> kvp in list1)
                {
                    if (comboBox2.SelectedItem.Equals(kvp.Value))
                        k = kvp.Key;
                }
                try
                {
                    MySqlCommand command1 = new MySqlCommand("DELETE FROM `список шихты` WHERE `Код шихты` = " + k, sqlConnection);
                    command1.Parameters.Clear();
                    command1.ExecuteNonQuery();
                    getList2();
                    comboBox2.SelectedIndex = 0;
                }
                catch { }
            }else
            {
                MessageBox.Show("Невозможно удалить состав (кол-во>=1)","Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count > 1)
            {
                int k = 0;
                foreach (KeyValuePair<int, string> kvp in list)
                {
                    if (comboBox1.SelectedItem.Equals(kvp.Value))
                        k = kvp.Key;
                }
                try
                {
                    MySqlCommand command1 = new MySqlCommand("DELETE FROM `список спеканий` WHERE `Код спекания` = " + k, sqlConnection);
                    command1.Parameters.Clear();
                    command1.ExecuteNonQuery();
                    getList();
                    comboBox1.SelectedIndex = 0;
                }
                catch { }
            }
            else
            {
                MessageBox.Show("Невозможно удалить спекание (кол-во>=1)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView4_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (change)
            {
                if (e.ColumnIndex == 2 && e.RowIndex != -1)
                {
                    try
                    {
                        mass1[e.RowIndex] = Convert.ToDouble(dataGridView4[2, e.RowIndex].Value);
                    }
                    catch (FormatException ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        dataGridView4[2, e.RowIndex].Value = mass1[e.RowIndex];
                    }
                }
            }
        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void размерToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (размерToolStripMenuItem.Checked)
            {
                размерToolStripMenuItem.Checked = false;
                Размер_журнал.Visible = false;
            }
            else
            {
                размерToolStripMenuItem.Checked = true;
                Размер_журнал.Visible = true;
            }
        }
        private void времяВыдержкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (времяВыдержкиToolStripMenuItem.Checked)
            {
                времяВыдержкиToolStripMenuItem.Checked = false;
                время_выдержки_журнал.Visible = false;
            }
            else
            {
                времяВыдержкиToolStripMenuItem.Checked = true;
                время_выдержки_журнал.Visible = true;
            }
        }
        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                try
                {
                    dataGridView5[1, e.RowIndex].Value = Convert.ToInt32(dataGridView5[1, e.RowIndex - 1].Value) + 1;
                }
                catch
                {
                    dataGridView5[1, e.RowIndex].Value = 1;
                }
            }
        }
        private void способПрессованияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (способПрессованияToolStripMenuItem.Checked)
            {
                способПрессованияToolStripMenuItem.Checked = false;
                Способ_прессования_журнал.Visible = false;
            }
            else
            {
                способПрессованияToolStripMenuItem.Checked = true;
                Способ_прессования_журнал.Visible = true;
            }
        }
        private void dataGridView5_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (change)
            {
                try
                {
                    grap();
                }
                catch
                {

                }

                for (int i = 0; i< dataGridView5.Rows.Count - 1; i++)
                {
                    dataGridView5[1, i].Value = i + 1;
                }

                for (int i = 1; i < dataGridView5.Rows.Count - 2; i++)
                {
                    dataGridView5[3, i].Value = dataGridView5[4, i-1].Value;
                }
            }
        }
        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(@"Справка.chm");
        }
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                Process.Start(@"Справка.chm");
            }
            if (e.Control && e.KeyCode == Keys.S)
            {
                сохранитьToolStripMenuItem.PerformClick();
            }

            if (e.Control && e.KeyCode == Keys.E)
            {
                сохранитьВExcelToolStripMenuItem.PerformClick();
            }
        }
        private void dataGridView5_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (change)
            {                                 
                if (e.RowIndex >= 0)
                {                      
                    try
                    {
                        if (e.ColumnIndex == 3 && Convert.ToDouble(dataGridView5[3, e.RowIndex].Value) > Convert.ToDouble(dataGridView5[4, e.RowIndex].Value))
                        {
                            dataGridView5[3, e.RowIndex].Value = dataGridView5[4, e.RowIndex-1].Value;
                        }

                        if (e.ColumnIndex == 4 && Convert.ToDouble(dataGridView5[4, e.RowIndex].Value) < Convert.ToDouble(dataGridView5[3, e.RowIndex].Value) && e.RowIndex > 0)
                        {
                            MessageBox.Show(e.ColumnIndex+""+e.RowIndex);
                            dataGridView5[4, e.RowIndex].Value = dataGridView5[3, e.RowIndex-1].Value;
                        }

                        if (e.ColumnIndex == 4 && Convert.ToDouble(dataGridView5[4, e.RowIndex].Value) > Convert.ToDouble(dataGridView5[4, e.RowIndex+1].Value) && dataGridView5[4, e.RowIndex + 1].Value != null)
                        {
                            dataGridView5[4, e.RowIndex].Value = massv4[e.RowIndex];
                        }

                        massv1[e.RowIndex] = Convert.ToDouble(dataGridView5[1, e.RowIndex].Value);
                        massv2[e.RowIndex] = Convert.ToDouble(dataGridView5[2, e.RowIndex].Value);
                        massv3[e.RowIndex] = Convert.ToDouble(dataGridView5[3, e.RowIndex].Value);
                        massv4[e.RowIndex] = Convert.ToDouble(dataGridView5[4, e.RowIndex].Value);

                    }
                    catch (FormatException)
                    {
                        dataGridView5[1, e.RowIndex].Value = massv1[e.RowIndex];
                        dataGridView5[2, e.RowIndex].Value = massv2[e.RowIndex];
                        dataGridView5[3, e.RowIndex].Value = massv3[e.RowIndex];
                        dataGridView5[4, e.RowIndex].Value = massv4[e.RowIndex];
                    }

                    
                    if (e.ColumnIndex == 3 && e.RowIndex == 0)
                        dataGridView5[3,0].Value = 0;

                    if (e.ColumnIndex == 3 && e.RowIndex > 0)
                    {
                        dataGridView5[4, e.RowIndex - 1].Value = dataGridView5[3, e.RowIndex].Value;
                    }
                    if (e.ColumnIndex == 4 && dataGridView5[3, e.RowIndex + 1].Value != null)
                    {
                        dataGridView5[3, e.RowIndex + 1].Value = dataGridView5[4, e.RowIndex].Value;
                    }
                    grap();
                }

            }
        }

        private void grap()
        {
            double Ymin = Convert.ToDouble(dataGridView5[3, 0].Value);
            double Ymax = Convert.ToDouble(dataGridView5[4, dataGridView5.Rows.Count - 1].Value);

            int count = dataGridView5.Rows.Count;
            for (int i = 0; i < dataGridView5.Rows.Count - 1; ++i)
            {
                if (Ymin > Convert.ToDouble(dataGridView5[3, i].Value))
                    Ymin = Convert.ToDouble(dataGridView5[3, i].Value);
                if (Ymax < Convert.ToDouble(dataGridView5[4, i].Value))
                    Ymax = Convert.ToDouble(dataGridView5[4, i].Value);
            }

            int massSize = 1;
            for (int i = 0; i < dataGridView5.Rows.Count - 1; ++i)
            {
                double pointY = Convert.ToDouble(dataGridView5[3, i].Value);
                double steps = (Convert.ToDouble(dataGridView5[4, i].Value) - Convert.ToDouble(dataGridView5[3, i].Value)) / Convert.ToDouble(dataGridView5[2, i].Value);
                int step = 0;
                while (step < steps)
                {
                    massSize++;
                    step++;
                }
            }

            double[] x = new double[massSize];
            double[] y = new double[massSize];

            int c = 1;
            int xc = 1;
            y[0] = Convert.ToDouble(dataGridView5[3, 0].Value);

            for (int i = 0; i < dataGridView5.Rows.Count - 1; ++i)
            {
                double pointY = Convert.ToDouble(dataGridView5[3, i].Value);
                double steps = (Convert.ToDouble(dataGridView5[4, i].Value) - Convert.ToDouble(dataGridView5[3, i].Value)) / Convert.ToDouble(dataGridView5[2, i].Value);
                int step = 0;
                while (step < steps)
                {
                    x[c] = xc;
                    y[c] = y[c - 1] + Convert.ToDouble(dataGridView5[2, i].Value);
                    c++;
                    step++;
                    xc++;
                }
            }

            chart1.ChartAreas[0].AxisX.Title = "время, мин.";
            chart1.ChartAreas[0].AxisY.Title = "температура, °С";
            chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Times New Roman", 14, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY.Minimum = Ymin;
            chart1.ChartAreas[0].AxisY.Maximum = Ymax;
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.ChartAreas[0].AxisX.MajorGrid.Interval = 1;
            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisY.Interval = 10;
            chart1.Series[0].BorderWidth = 3;
            chart1.Series[0].Points.DataBindXY(x, y);
        }
    }
}

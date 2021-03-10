using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace _12MConsulting
{
    public partial class Main : Form
    {
        public static SqlConnection connection = new SqlConnection("server = DESKTOP-MNAVLBM\\SQLEXPRESS;Initial Catalog = test;Integrated Security=True;");
        
        public Main()
        {
            InitializeComponent();
        }
        private void Main_Load(object sender, EventArgs e)
        {
            
            FillComboBox();

        }
        private void FillComboBox()
        {
            connection.Open();
            SqlDataAdapter da = new SqlDataAdapter("SELECT MalKodu FROM STK",connection);
            System.Data.DataTable dt = new System.Data.DataTable();
           
            da.Fill(dt);
           
            foreach(DataRow dr in dt.Rows)
            {
                comboBox2.Items.Add(dr["MalKodu"].ToString());
            }
            connection.Close();
        }

        public static int Stok;
        private void Display(string MalKodu, int baslangicTarihi, int bitisTarihi)
        {
            try {
                connection = new SqlConnection("server = DESKTOP-MNAVLBM\\SQLEXPRESS;Initial Catalog = test;Integrated Security=True;");
                connection.Open();
                string DropTable = "drop table ViewResult";
                SqlCommand cmd2 = new SqlCommand(DropTable, connection);
                cmd2.ExecuteNonQuery();
                string CreateTable = "if not exists(select * from sysobjects where name = 'ViewResult' and xtype = 'U')    create table ViewResult(SiraNo INT, IslemTur VARCHAR(10), EvrakNo VARCHAR(50), Tarih VARCHAR(16), GirisMiktar VARCHAR(10), CikisMiktar VARCHAR(10), Stok VARCHAR(10))";
                string selectQuery = "SELECT IslemTur, EvrakNo, Tarih, Miktar FROM STI WHERE @MalKodu=MalKodu AND @baslangicTarihi<=Tarih AND @bitisTarihi>Tarih ORDER BY Tarih";
                SqlCommand command = new SqlCommand(selectQuery, connection);
                command.Parameters.AddWithValue("@MalKodu", MalKodu);
                command.Parameters.AddWithValue("@baslangicTarihi", baslangicTarihi);
                command.Parameters.AddWithValue("@bitisTarihi", bitisTarihi);
                SqlCommand cmd = new SqlCommand(CreateTable, connection);
                
                cmd.ExecuteNonQuery();
                command.ExecuteNonQuery();
                
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = command;
                System.Data.DataTable ds = new System.Data.DataTable();
                da.Fill(ds);
                int i = 1;
                foreach(DataRow dr in ds.Rows)
                {
                    if(dr["IslemTur"].ToString() == "0")
                    {
                        string IslemTur = "Giriş";
                        string EvrakNo = dr["EvrakNo"].ToString();
                        int tarih = Convert.ToInt32(dr["Tarih"]);
                        int a = Convert.ToInt32(dr["Miktar"]);
                        Stok += a;
                        string CikisMiktar = "0";
                        string GirisMiktar = a.ToString();
                        string InsertColoumnValues = "INSERT INTO ViewResult (SiraNo,IslemTur,EvrakNo, Tarih, GirisMiktar, CikisMiktar, Stok) Values(@SiraNo, @IslemTur,@EvrakNo,CONVERT(VARCHAR(15), CAST(@Tarih - 2 AS datetime), 104), @GirisMiktar , @CikisMiktar, @Stok)";
                        SqlCommand command2 = new SqlCommand(InsertColoumnValues, connection);
                        command2.Parameters.AddWithValue("@Stok", Stok.ToString());
                        command2.Parameters.AddWithValue("@GirisMiktar", GirisMiktar);
                        command2.Parameters.AddWithValue("@CikisMiktar", CikisMiktar);
                        command2.Parameters.AddWithValue("@Tarih", tarih);
                        command2.Parameters.AddWithValue("@EvrakNo", EvrakNo);
                        command2.Parameters.AddWithValue("@IslemTur", IslemTur);
                        command2.Parameters.AddWithValue("@SiraNo", i.ToString());
                        i++;
                        command2.ExecuteNonQuery();
                    }
                    else
                    {
                        string IslemTur = "Çıkış";
                        string EvrakNo = dr["EvrakNo"].ToString();
                        int tarih = Convert.ToInt32(dr["Tarih"]);
                        int a = Convert.ToInt32(dr["Miktar"]);
                        Stok -= a;
                        string CikisMiktar = a.ToString();
                        string GirisMiktar = "0";
                        string AlterColoumnValues = "INSERT INTO ViewResult (SiraNo, IslemTur, EvrakNo,Tarih, CikisMiktar, GirisMiktar, Stok) Values(@SiraNo, @IslemTur, @EvrakNo,CONVERT(VARCHAR(15), CAST(@Tarih - 2 AS datetime), 104), @CikisMiktar, @GirisMiktar, @Stok)";
                        SqlCommand command3 = new SqlCommand(AlterColoumnValues, connection);
                        command3.Parameters.AddWithValue("@CikisMiktar", CikisMiktar);
                        command3.Parameters.AddWithValue("@GirisMiktar", GirisMiktar);
                        command3.Parameters.AddWithValue("@Stok", Stok.ToString());
                        command3.Parameters.AddWithValue("@Tarih", tarih);
                        command3.Parameters.AddWithValue("@EvrakNo", EvrakNo);
                        command3.Parameters.AddWithValue("@IslemTur", IslemTur);
                        command3.Parameters.AddWithValue("@SiraNo", i.ToString());
                        i++;
                        command3.ExecuteNonQuery();
                    }
                }
                Stok = 0;
                string ViewResult = "SELECT * FROM ViewResult";
                SqlCommand resultCommand = new SqlCommand(ViewResult, connection);
                SqlDataAdapter d = new SqlDataAdapter();
                d.SelectCommand = resultCommand;
                System.Data.DataTable dt = new System.Data.DataTable();
                d.Fill(dt);
                dataGridView1.DataSource = dt;
                connection.Close();
            }
            catch
            {
                throw;
            }

        }

        private void btnListele_Click(object sender, EventArgs e)
        {
            GetSelectedValue();
        }

        private void GetSelectedValue()
        {
            string MalKodu = comboBox2.SelectedItem.ToString();
            DateTime dtBaslangic = new DateTime();
            dtBaslangic = dateTimePicker1.Value.Date;
            DateTime dtBitis = new DateTime();
            dtBitis = dateTimePicker2.Value.Date;
            int baslangicTarihi = Convert.ToInt32(dtBaslangic.ToOADate());
            int bitisTarihi = Convert.ToInt32(dtBitis.ToOADate());
            Display(MalKodu, baslangicTarihi, bitisTarihi);
        }

        private void export_Click(object sender, EventArgs e)
        {
            _Application app = new Microsoft.Office.Interop.Excel.Application();
            _Workbook workbook = app.Workbooks.Add(Type.Missing);
            _Worksheet worksheet = null;
            app.Visible = false; 
            worksheet = workbook.Sheets["Sayfa1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Exported from gridview";
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            
            string spreadsheetName = "Export";
            app.DisplayAlerts = false;
            Dialog saveAsDialog = app.Dialogs[XlBuiltInDialog.xlDialogSaveAs];
            saveAsDialog.Show(spreadsheetName);

            workbook.Close(true);
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        Bitmap bitmap;
        private void yazdir_Click(object sender, EventArgs e)
        {
            int height = dataGridView1.Height;
            dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height * 2;
            bitmap = new Bitmap(dataGridView1.Width, dataGridView1.Height);
            dataGridView1.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, dataGridView1.Width, dataGridView1.Height));
            dataGridView1.Height = height;
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
        }
        private void PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(bitmap,0,0);
        }
    }
}

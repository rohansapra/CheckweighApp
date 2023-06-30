using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.BinaryFileFormat;
using ExcelLibrary.BinaryDrawingFormat;
using System.Configuration;
using Aspose.Cells;



namespace CheckWeigh
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

       
        private void btnExcel_Click(object sender, EventArgs e)
        {
            string Value = textBox1.Text;
            string mainconn = ConfigurationManager.ConnectionStrings["myconnection"].ConnectionString;
            SqlConnection sqlconn = new SqlConnection(mainconn);
            sqlconn.Open();
            SqlCommand cmdDataBase = new SqlCommand("SELECT [SORLotID],[LotStartDate],[LotEndDate] FROM [CheckWeigh].[dbo].[tblSORLot] where [ScheduleOLSNRelease] = @Value ", sqlconn);
            cmdDataBase.Parameters.AddWithValue("@Value", Value);
            SqlDataReader s = cmdDataBase.ExecuteReader();
            DateTime b=DateTime.Now;
            DateTime c=DateTime.Now;
            int a = 0;
            if (s.Read())
            {
                a = (Int32)s.GetValue(0);
               // Console.WriteLine(a);
                b = (DateTime)s.GetValue(1);
               // Console.WriteLine(b);
                c = (DateTime)s.GetValue(2);
               // Console.WriteLine(c);
            }
            s.Close();

            SqlCommand cmd = new SqlCommand("SELECT [SORLotShiftChangeID] FROM [CheckWeigh].[dbo].[tblSORLotShiftChange] where [SORLotID] = @a", sqlconn);
            cmd.Parameters.AddWithValue("@a", a);
            SqlDataReader sda = cmd.ExecuteReader();
            int d = 0;
            if (sda.Read())
            {
                d = (Int32)sda.GetValue(0);
               // Console.WriteLine(d);
            }
            sda.Close();

            if (radioButton1.Checked)
            {
                
                SqlCommand cm = new SqlCommand("SELECT [CapTorqueID],[SORLotShiftChangeID],convert(varchar, [DateTorqueMeasured], 120)as DateTorqueMeasured,[CapperNumber],convert(varchar, [SampleTorque],120)as SampleTorque,[OperatorEmployeeNumber],wm.[ShiftID] FROM [CheckWeigh].[dbo].[tblCapTorqueMeasure] wm inner join tblshifts s on s.ShiftID=wm.ShiftID where [SORLotShiftChangeID] = @d  and [DateTorqueMeasured] > @b and [DateTorqueMeasured] < @c order by SorLotShiftChangeID,CapTorqueID", sqlconn);
                cm.Parameters.AddWithValue("@d", d);
                cm.Parameters.AddWithValue("@b", b);
                cm.Parameters.AddWithValue("@c", c);

                
                SqlDataAdapter sdaa = new SqlDataAdapter();
                sdaa.SelectCommand = cm;
                DataTable dbdataset = new DataTable();
                sdaa.Fill(dbdataset);
                BindingSource bSource = new BindingSource();
                bSource.DataSource = dbdataset;
                sdaa.Update(dbdataset);

                
                DataSet ds = new DataSet("New_DataSet1");
                ds.Locale = System.Threading.Thread.CurrentThread.CurrentCulture;
                sdaa.Fill(dbdataset);
                ds.Tables.Add(dbdataset);
                ExcelLibrary.DataSetHelper.CreateWorkbook("MyExcelFile1.xls", ds);
               

            }

            else if(radioButton2.Checked)
            {
                  
                SqlCommand cm = new SqlCommand("SELECT [WeightID],[SORLotShiftChangeID],convert(varchar, [DateWeightMeasured], 120)as DateWeightMeasured,[StemNumber],convert(varchar, [SampleWeight],120)as SampleWeight,[OperatorEmployeeNumber],[ShiftID] FROM [CheckWeigh].[dbo].[tblWeightMeasure] where [SORLotShiftChangeID] = @d and DateWeightMeasured > @b and DateWeightMeasured < @c order by SorLotShiftChangeID,WeightID", sqlconn);
                cm.Parameters.AddWithValue("@d", d);
                cm.Parameters.AddWithValue("@b", b);
                cm.Parameters.AddWithValue("@c", c);

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cm;
                DataTable dbdataset = new DataTable();
                da.Fill(dbdataset);
                BindingSource bSource = new BindingSource();
                bSource.DataSource = dbdataset;
                da.Update(dbdataset);
                
                DataSet ds = new DataSet("New_DataSet2");
                ds.Locale = System.Threading.Thread.CurrentThread.CurrentCulture;
                da.Fill(dbdataset);
                ds.Tables.Add(dbdataset);
                ExcelLibrary.DataSetHelper.CreateWorkbook("MyExcelFile2.xls", ds);
                    
            }
            else 
            {
                   
                SqlCommand cm = new SqlCommand("SELECT [SealIntegrityID],[SORLotShiftChangeID],convert(varchar, [DateSealIntegrityMeasured], 120)as DateSealIntegrityMeasured,convert(varchar, [SealHeadNumber],120)as SealHeadNumber,[OperatorEmployeeNumber],[ShiftID] FROM [CheckWeigh].[dbo].[tblSealIntegrityMeasure] where [SORLotShiftChangeID] = @d and DateSealIntegrityMeasured > @b and DateSealIntegrityMeasured < @c order by SorLotShiftChangeID,SealIntegrityID", sqlconn);
                cm.Parameters.AddWithValue("@d", d);
                cm.Parameters.AddWithValue("@b", b);
                cm.Parameters.AddWithValue("@c", c);

                SqlDataAdapter ssda = new SqlDataAdapter();
                ssda.SelectCommand = cm;
                DataTable dbdataset = new DataTable();
                ssda.Fill(dbdataset);
                BindingSource bSource = new BindingSource();
                bSource.DataSource = dbdataset;
                ssda.Update(dbdataset);

                DataSet ds = new DataSet("New_DataSet3");
                ds.Locale = System.Threading.Thread.CurrentThread.CurrentCulture;
                ssda.Fill(dbdataset);
                ds.Tables.Add(dbdataset);
                ExcelLibrary.DataSetHelper.CreateWorkbook("MyExcelFile3.xls", ds);

            }
            sqlconn.Close();
        }


        private void btnPdf_Click(object sender, EventArgs e)
        {
            if(radioButton1.Checked)
                {
                Aspose.Cells.Workbook work = new Aspose.Cells.Workbook("D:\\Users\\RSapra3\\source\\repos\\Backup\\CheckWeigh\\bin\\Debug\\MyExcelFile1.xls");
                work.Save("D:\\Users\\RSapra3\\source\\repos\\Backup\\CheckWeigh\\bin\\Debug\\MyExcelFile1.pdf", SaveFormat.Pdf);
            }
            else if(radioButton2.Checked)
            {
                Aspose.Cells.Workbook work = new Aspose.Cells.Workbook("D:\\Users\\RSapra3\\source\\repos\\Backup\\CheckWeigh\\bin\\Debug\\MyExcelFile2.xls");
                work.Save("D:\\Users\\RSapra3\\source\\repos\\Backup\\CheckWeigh\\bin\\Debug\\MyExcelFile2.pdf", SaveFormat.Pdf);
            }
            else
            {
                Aspose.Cells.Workbook work = new Aspose.Cells.Workbook("D:\\Users\\RSapra3\\source\\repos\\Backup\\CheckWeigh\\bin\\Debug\\MyExcelFile3.xls");
                work.Save("D:\\Users\\RSapra3\\source\\repos\\Backup\\CheckWeigh\\bin\\Debug\\MyExcelFile3.pdf", SaveFormat.Pdf);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
       
        }

    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CsvHelper;

namespace primodvdprices
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //file load button
           
            string filePath;
            OpenFileDialog op = new OpenFileDialog();
            DialogResult result = op.ShowDialog();
            if (result == DialogResult.OK)
            {
                filePath = op.FileName;
                if (tabControl1.SelectedTab == tabControl1.TabPages[0])
                {
                    textBox1.Text = filePath;
                }
                else
                {
                    textBox4.Text = filePath;
                }
                
            }
        }
       
       
        string [] header()
        {
            string[] headers = { "Handle", "Title", "Body (HTML)", "Vendor", "Standard Product Type", "Custom Product Type", "Tags", "Published", "Option1 Name", 
                "Option1 Value", "Option2 Name", "Option2 Value", "Option3 Name", "Option3 Value", "Variant SKU", "Variant Grams", "Variant Inventory Tracker", 
                "Variant Inventory Policy", "Variant Fulfillment Service", "Variant Price", "Variant Compare At Price", "Variant Requires Shipping", 
                "Variant Taxable", "Variant Barcode", "Image Src", "Image Position", "Image Alt Text", "Gift Card", "SEO Title", "SEO Description", 
                "Google Shopping / Google Product Category", "Google Shopping / Gender", "Google Shopping / Age Group", "Google Shopping / MPN", 
                "Google Shopping / AdWords Grouping", "Google Shopping / AdWords Labels", "Google Shopping / Condition", "Google Shopping / Custom Product", 
                "Google Shopping / Custom Label 0", "Google Shopping / Custom Label 1", "Google Shopping / Custom Label 2", "Google Shopping / Custom Label 3", 
                "Google Shopping / Custom Label 4", "Variant Image", "Variant Weight Unit", "Variant Tax Code", "Cost per item", "Status" };
            return headers;
        }
      private void LargeMousepad()
        {

            string[] headers = header();
            DataTable dt = new DataTable();
            DataTable lastone = new DataTable();
            if (textBox1.Text != "")
            {
                dt = ReadExcelFile("Sheet1", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Choose an excel File!");
                return;
            }
            DataColumn col = dt.Columns[0];


            for (int i = 0; i < headers.Length; i++)
            {
                lastone.Columns.Add(headers[i]);
            }

            try
            {


                lastone.Rows.Add();
                lastone.Rows[0][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');
                if (lastone.Rows[0][0].ToString().IndexOf("Large") < 0)
                {
                    MessageBox.Show("Wrong File");
                    return;
                }
                string contitle2 = lastone.Rows[0][0].ToString().Replace('-', ' ');
                string url = contitle2;
                string title2 = contitle2;
                int index2 = title2.IndexOf("Mouse Pad") + 9;
                if (index2 > 9)
                {
                    title2 = title2.Substring(0, index2);
                }
                

                index2 = contitle2.IndexOf("By") + 3;
                if (index2 > 3)
                {
                    contitle2 = contitle2.Substring(index2, contitle2.Length - index2);
                }
                lastone.Rows[0][1] = title2;
                lastone.Rows[0][2] = "<p>Use this mouse pad decorate your home computer or use at work!</p>"+
                    "<p>It makes a great gift for yourself or someone else.</p>"+
                    "<p>Not only will our placemats keep your messes to a minimum, but they can also be used as a mouse pad.</p>"+
                    "<p>The shape is rectangle and  come in different sizes which are listed below. The image is sublimated to the surface so it won't fade,"+
                    "crack or peel. This is a high quality mouse pad that will take all the abuse your mouse can give it.</p>"+
                    "< p > (Also note that color may vary depending on your monitor settings.)</ p >\n"+
                     "< p > 10\" x 16\" x 1 / 8\"</p>\n"+
                    "< p > 12\" x 18\" x 1 / 8\"</p>\n"+
                    "< p > 14\" x 24\" x 1 / 8\"</p>\n"+
                    "< p > 18\" x 36\" x 1 / 8\"</p>";
                for (int i = 0; i < 4; i++)
                {
                    lastone.Rows.Add();
                    lastone.Rows[i][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');
                    
                   
                    lastone.Rows[i][16] = "shopify";
                    lastone.Rows[i][17] = "continue";
                    lastone.Rows[i][18] = "manual";
                    lastone.Rows[i][21] = "TRUE";
                    lastone.Rows[i][22] = "TRUE";
                    lastone.Rows[i][44] = "oz";
                }

                lastone.Rows[0][3] = contitle2;
                lastone.Rows[0][5] = "Large Mouse Pad";
                lastone.Rows[0][7] = "TRUE";
                lastone.Rows[0][8] = "Title";
                lastone.Rows[0][9] = "10 x 16";
                lastone.Rows[1][9] = "12 x 18";
                lastone.Rows[2][9] = "14 x 24";
                lastone.Rows[3][9] = "18 x 36";
                lastone.Rows[0][15] = 170.0971388;
                lastone.Rows[1][15] = 226.796185;
                lastone.Rows[2][15] = 340.1942775;
                lastone.Rows[3][15] = 453.59237;
                lastone.Rows[0][19] = 15;
                lastone.Rows[1][19] = 20;
                lastone.Rows[2][19] = 30;
                lastone.Rows[3][19] = 40;

                lastone.Rows[0][24] = "https://www.nurdtymedesigners.com/largemousepad/" + url + ".jpg";
                lastone.Rows[0][25] = 1;
                lastone.Rows[0][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                lastone.Rows[0][30] = "Electronics > Electronics Accessories > Computer Accessories > Mouse Pads";
                lastone.Rows[0][31] = "Unisex";
                lastone.Rows[0][32] = "Adult";
                lastone.Rows[0][36] = "New";
                lastone.Rows[0][37] = "TRUE";
                
                lastone.Rows[0][47] = "active";

                //automate

                for (int i = 1; i < dt.Rows.Count+1; i++)
                {
                    lastone.Rows.Add();
                    string handle = dt.Rows[i-1][0].ToString();
                    
                    string contitle = handle.Substring(0, handle.Length - 4).Replace('-', ' ');
                    string title = contitle;
                    url = contitle;
                    int index = title.IndexOf("Mouse Pad") + 9;
                    if (index > 9)
                    {
                        title = title.Substring(0, index);
                    }
                    lastone.Rows[4*i][1] = title;

                    index = contitle.IndexOf("By") + 3;
                    if (index > 3)
                    {
                        contitle = contitle.Substring(index, contitle.Length - index);
                    }
                    lastone.Rows[4*i ][2] = "<p>Use this mouse pad decorate your home computer or use at work!</p>" +
                    "<p>It makes a great gift for yourself or someone else.</p>" +
                    "<p>Not only will our placemats keep your messes to a minimum, but they can also be used as a mouse pad.</p>" +
                    "<p>The shape is rectangle and  come in different sizes which are listed below. The image is sublimated to the surface so it won't fade," +
                    "crack or peel. This is a high quality mouse pad that will take all the abuse your mouse can give it.</p>" +
                    "< p > (Also note that color may vary depending on your monitor settings.)</ p >\n" +
                     "< p > 10\" x 16\" x 1 / 8\"</p>\n" +
                    "< p > 12\" x 18\" x 1 / 8\"</p>\n" +
                    "< p > 14\" x 24\" x 1 / 8\"</p>\n" +
                    "< p > 18\" x 36\" x 1 / 8\"</p>";
                    for (int j = 4*i; j < 4*i+4; j++)
                    {
                        lastone.Rows.Add();
                        lastone.Rows[j][0] = handle.Substring(0, handle.Length - 4).Replace(' ', '-');
                        

                        lastone.Rows[j][16] = "shopify";
                        lastone.Rows[j][17] = "continue";
                        lastone.Rows[j][18] = "manual";
                        lastone.Rows[j][21] = "TRUE";
                        lastone.Rows[j][22] = "TRUE";
                        lastone.Rows[j][44] = "oz";
                    }

                    lastone.Rows[4*i][3] = contitle;
                    lastone.Rows[4*i][5] = "Large Mouse Pad";
                    lastone.Rows[4*i][7] = "TRUE";
                    lastone.Rows[4*i][8] = "Title";
                    lastone.Rows[4*i][9] = "10 x 16";
                    lastone.Rows[4*i+1][9] = "12 x 18";
                    lastone.Rows[4*i+2][9] = "14 x 24";
                    lastone.Rows[4*i+3][9] = "18 x 36";
                    lastone.Rows[4*i][15] = "170.0971388";
                    lastone.Rows[4*i+1][15] = "226.796185";
                    lastone.Rows[4*i+2][15] = "340.1942775";
                    lastone.Rows[4*i+3][15] = "453.59237";
                    lastone.Rows[4*i][19] = 15;
                    lastone.Rows[4*i+1][19] = 20;
                    lastone.Rows[4*i+2][19] = 30;
                    lastone.Rows[4*i+3][19] = 40;

                    lastone.Rows[4*i][24] = "https://www.nurdtymedesigners.com/largemousepad/" + url + ".jpg";
                    lastone.Rows[4*i][25] = 1;
                    lastone.Rows[4*i][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                    lastone.Rows[4*i][30] = "Electronics > Electronics Accessories > Computer Accessories > Mouse Pads";
                    lastone.Rows[4*i][31] = "Unisex";
                    lastone.Rows[4*i][32] = "Adult";
                    lastone.Rows[4*i][36] = "New";
                    lastone.Rows[4*i][37] = "TRUE";

                    lastone.Rows[4*i][47] = "active";

                }
    //            lastone = lastone.Rows
    //.Cast<DataRow>()
    //.Where(row => !row.ItemArray.All(field => field is DBNull ||
    //                                 string.IsNullOrWhiteSpace(field as string)))
    //.CopyToDataTable();
                dataGridView1.DataSource = lastone;
            }
            catch
            {
                MessageBox.Show("Are you sure to choose right xls file");
            }
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                try
                {
                    lastone = add_Tags(lastone, ReadTags("Sheet1", textBox2.Text));
                }
               catch { }
            }
            // write to csv file
            ToCsvFile(lastone,textBox1.Text);


        }
        private void MousePad()
        {
            string[] headers = header();
            DataTable dt = new DataTable();
            DataTable lastone = new DataTable();
            if (textBox1.Text != "")
            {
                dt = ReadExcelFile("Sheet1", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Choose an excel File!");
                return;
            }
            DataColumn col = dt.Columns[0];


            for (int i = 0; i < headers.Length; i++)
            {
                lastone.Columns.Add(headers[i]);
            }

            try
            {


                lastone.Rows.Add();
                lastone.Rows[0][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');

                string contitle2 = lastone.Rows[0][0].ToString().Replace('-', ' ');
                string url = contitle2;
                string title2 = contitle2;
                int index2 = title2.IndexOf("Mouse Pad") + 9;
                if (index2 > 9)
                {
                    title2 = title2.Substring(0, index2);
                }
                lastone.Rows[0][1] = title2;

                index2 = contitle2.IndexOf("By") + 3;
                if (index2 > 3)
                {
                    contitle2 = contitle2.Substring(index2, contitle2.Length - index2);
                }
                lastone.Rows[0][2] = "<p>Use this mouse pad decorate your home computer or use at work!</p>" +
                    "< p > It makes a great gift for yourself or someone else.</ p >" +
       "< p > The size of this rectangle shaped mouse pad is 7.75\" x 9.25\" x 1 / 4\", (this is not a skinny, flimsy mouse pad), with open cell rubber backing to prevent movement when your mouse gets wild. The image is sublimated to the surface so it won't fade, crack or peel. This is a high quality mouse pad that will take all the abuse your mouse can give it.</p>" +
            "< p > (Also note that color may vary depending on your monitor settings.)</ p > ";
                lastone.Rows[0][3] = contitle2;
                lastone.Rows[0][5] = "Mouse Pad";
                lastone.Rows[0][7] = "TRUE";
                lastone.Rows[0][8] = "Title";
                lastone.Rows[0][9] = "Default Title";
                lastone.Rows[0][15] = "112.9899";
                lastone.Rows[0][16] = "shopify";
                lastone.Rows[0][17] = "continue";
                lastone.Rows[0][18] = "manual";
                lastone.Rows[0][19] = 9.99;
                lastone.Rows[0][21] = "TRUE";
                lastone.Rows[0][22] = "TRUE";
                lastone.Rows[0][24] = "https://www.nurdtymedesigners.com/mousepad/" + url + ".jpg";
                lastone.Rows[0][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else.";
                lastone.Rows[0][30] = "1993";
                lastone.Rows[0][31] = "Unisex";
                lastone.Rows[0][32] = "Adult";
                lastone.Rows[0][36] = "New";
                lastone.Rows[0][37] = "TRUE";
                lastone.Rows[0][44] = "lb";
                lastone.Rows[0][47] = "active";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    lastone.Rows.Add();
                    string handle = dt.Rows[i][0].ToString();
                    lastone.Rows[i + 1][0] = handle.Substring(0, handle.Length - 4).Replace(' ', '-');
                    string contitle = lastone.Rows[i + 1][0].ToString().Replace('-', ' ');
                    string title = contitle;
                    url = contitle;
                    int index = title.IndexOf("Mouse Pad") + 9;
                    if (index > 9)
                    {
                        title = title.Substring(0, index);
                    }
                    lastone.Rows[i + 1][1] = title;

                    index = contitle.IndexOf("By") + 3;
                    if (index > 3)
                    {
                        contitle = contitle.Substring(index, contitle.Length - index);
                    }
                    lastone.Rows[i + 1][2] = "<p>Use this mouse pad decorate your home computer or use at work!</p>" +
                    "< p > It makes a great gift for yourself or someone else.</ p >" +
       "< p > The size of this rectangle shaped mouse pad is 7.75\" x 9.25\" x 1 / 4\", (this is not a skinny, flimsy mouse pad), with open cell rubber backing to prevent movement when your mouse gets wild. The image is sublimated to the surface so it won't fade, crack or peel. This is a high quality mouse pad that will take all the abuse your mouse can give it.</p>" +
            "< p > (Also note that color may vary depending on your monitor settings.)</ p > ";
                    lastone.Rows[i + 1][3] = contitle;
                    lastone.Rows[i + 1][5] = "MousePad";
                    lastone.Rows[i + 1][7] = "TRUE";
                    lastone.Rows[i + 1][8] = "Title";
                    lastone.Rows[i + 1][9] = "Default Title";
                    lastone.Rows[i + 1][15] = "112.9899";
                    lastone.Rows[i + 1][16] = "shopify";
                    lastone.Rows[i + 1][17] = "continue";
                    lastone.Rows[i + 1][18] = "manual";
                    lastone.Rows[i + 1][19] = 9.99;
                    lastone.Rows[i + 1][21] = "TRUE";
                    lastone.Rows[i + 1][22] = "TRUE";
                    lastone.Rows[i + 1][24] = "https://www.nurdtymedesigners.com/mousepad/" + url + ".jpg";
                    lastone.Rows[i + 1][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else.";
                    lastone.Rows[i + 1][30] = "1993";
                    lastone.Rows[i + 1][31] = "Unisex";
                    lastone.Rows[i + 1][32] = "Adult";
                    lastone.Rows[i + 1][36] = "New";
                    lastone.Rows[i + 1][37] = "TRUE";
                    lastone.Rows[i + 1][44] = "lb";
                    lastone.Rows[i + 1][47] = "active";

                }
                dataGridView1.DataSource = lastone;
            }
            catch
            {
                MessageBox.Show("Are you sure to choose right xls file");
            }
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                lastone = add_Tags(lastone, ReadTags("Sheet1", textBox2.Text));
            }
            // write to csv file
            ToCsvFile(lastone,textBox1.Text);

        }
        private void ToCsvFile(DataTable lastone,string path2)
        {
            try
            {
                string dte = Path.GetFileName(path2);
                string path = new FileInfo(path2).Directory.FullName + "/" + dte.Replace(".","byVbalpic")+ ".csv";
                using (var textWriter = File.CreateText(path))
                {
                    using (CsvWriter csv = new CsvWriter(textWriter, System.Globalization.CultureInfo.CurrentCulture))
                    {

                        // Write columns
                        foreach (DataColumn column in lastone.Columns)
                            csv.WriteField(column.ColumnName);
                        csv.NextRecord();
                        // Write row values
                        foreach (DataRow row in lastone.Rows)
                        {
                            for (var i = 0; i < lastone.Columns.Count; i++)
                                csv.WriteField(row[i]);
                            csv.NextRecord();
                        }
                    }
                }
                MessageBox.Show("You have successfully exported the file.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // converting button
            if (radioMousePad.Checked == true)
            {
                MousePad();
            }
            else if (radioLargeMousepad.Checked == true)
            {
                LargeMousepad();
            }
            else if (radioCarSeat.Checked == true)
            {
                Carseat();
            }
            else if (radioCuttingBoard.Checked == true)
            {
                CuttingBoard();
            }
            else if (radioHoodedBlanket.Checked == true)
            {
                HoodedBlanket();
            }
            else if (radioTumblers.Checked == true)
            {
                Tumblers();
            }
            else if (radioCoasters.Checked == true)
            {
                Coasters();
            }






        }

        private void Coasters()
        {

            string[] headers = header();
            DataTable dt = new DataTable();
            DataTable lastone = new DataTable();
            if (textBox1.Text != "")
            {
                dt = ReadExcelFile("Sheet1", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Choose an excel File!");
                return;
            }
            DataColumn col = dt.Columns[0];


            for (int i = 0; i < headers.Length; i++)
            {
                lastone.Columns.Add(headers[i]);
            }

            try
            {


                lastone.Rows.Add();
                lastone.Rows[0][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');

                if (lastone.Rows[0][0].ToString().ToLower().IndexOf("coaster") < 0)
                {
                    MessageBox.Show("Wrong File");
                    return;
                }
                string contitle2 = lastone.Rows[0][0].ToString().Replace('-', ' ');
                string url = contitle2;
                string title2 = contitle2;
                int index2 = title2.ToLower().IndexOf("coasters") + 8;
                if (index2 > 7)
                {
                    title2 = title2.Substring(0, index2);
                }
                lastone.Rows[0][1] = title2;

                index2 = contitle2.IndexOf("By") + 3;
                if (index2 > 3)
                {
                    contitle2 = contitle2.Substring(index2, contitle2.Length - index2);
                }
                lastone.Rows[0][2] = "Our Premium Beverage Coasters feature full color prints on a glossy hardboard top. Coasters are backed " +
                    "with medium-density fibrewood (MDF) cork and measure 3.75\" x 3.75\". Available as a pack of four or six ";

                for (int i = 0; i < 2; i++)
                {
                    lastone.Rows.Add();
                    lastone.Rows[i][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');

                }


                lastone.Rows[0][3] = contitle2;
                lastone.Rows[0][4] = "Home & Garden > Kitchen & Dining > Barware > Coasters";
               
                lastone.Rows[0][7] = "TRUE";
                lastone.Rows[0][8] = "Pack";
                lastone.Rows[0][9] = 4;
                lastone.Rows[1][9] = 6;



                lastone.Rows[0][15] = 453;
                lastone.Rows[1][15] = 453;
                lastone.Rows[0][16] = "shopify";
                lastone.Rows[1][16] = "shopify";
                lastone.Rows[0][17] = "continue";
                lastone.Rows[1][17] = "continue";
                lastone.Rows[0][18] = "manual";
                lastone.Rows[1][18] = "manual";
                lastone.Rows[0][19] = 24.99;
                lastone.Rows[1][19] = 29.99;
                lastone.Rows[0][21] = "True";
                lastone.Rows[1][21] = "True";
                lastone.Rows[0][22] = "True";
                lastone.Rows[1][22] = "True";


                lastone.Rows[0][24] = "https://www.nurdtymedesigners.com/coasters/" + url + ".jpg";
                //lastone.Rows[1][24] = "https://www.nurdtymedesigners.com/coasters/" + lastone.Rows[0][1] + " No Background By " + lastone.Rows[0][3] + ".jpg";
                lastone.Rows[0][25] = 1;
                
                lastone.Rows[0][27] = "FALSE";
                //lastone.Rows[0][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                lastone.Rows[0][30] = "Home & Garden > Kitchen & Dining > Barware > Coasters";
                lastone.Rows[0][31] = "Unisex";
                lastone.Rows[0][32] = "adult";
                lastone.Rows[0][36] = "New";
                lastone.Rows[0][37] = "TRUE";
                lastone.Rows[0][44] = "lb";
                lastone.Rows[1][44] = "lb";
                lastone.Rows[0][47] = "active";

                //--------------------------------------------------------------

                for (int i = 1; i < dt.Rows.Count + 1; i++)
                {
                    lastone.Rows.Add();
                    string handle = dt.Rows[i - 1][0].ToString();

                    string contitle = handle.Substring(0, handle.Length - 4).Replace('-', ' ');
                    string title = contitle;
                    url = contitle;
                    int index = title.ToLower().IndexOf("coasters") + 8;
                    if (index > 7)
                    {
                        title = title.Substring(0, index);
                    }
                    lastone.Rows[2 * i][1] = title;

                    index = contitle.IndexOf("By") + 3;
                    if (index > 3)
                    {
                        contitle = contitle.Substring(index, contitle.Length - index);
                    }
                    lastone.Rows[2 * i][2] = "Our Premium Beverage Coasters feature full color prints on a glossy hardboard top. Coasters are backed " +
                    "with medium-density fibrewood (MDF) cork and measure 3.75\" x 3.75\". Available as a pack of four or six ";



                    for (int j = 2 * i; j < 2 * i + 2; j++)
                    {
                        lastone.Rows.Add();
                        lastone.Rows[j][0] = handle.Substring(0, handle.Length - 4).Replace(' ', '-');


                    }


                    lastone.Rows[2 * i][3] = contitle;
                    lastone.Rows[2 * i][4] = "Home & Garden > Kitchen & Dining > Barware > Coasters";
                    lastone.Rows[2 * i][7] = "TRUE";
                    lastone.Rows[2 * i][8] = "Pack";
                    lastone.Rows[2 * i][9] = 4;
                    lastone.Rows[2 * i+1][9] = 6;



                    lastone.Rows[2 * i][15] = 453;
                    lastone.Rows[2 * i+1][15] = 453;
                    lastone.Rows[2 * i][16] = "shopify";
                    lastone.Rows[2 * i+1][16] = "shopify";
                    lastone.Rows[2 * i][17] = "continue";
                    lastone.Rows[2 * i+1][17] = "continue";
                    lastone.Rows[2 * i][18] = "manual";
                    lastone.Rows[2 * i+1][18] = "manual";
                    lastone.Rows[2 * i][19] = 24.99;
                    lastone.Rows[2 * i+1][19] = 29.99;
                    lastone.Rows[2 * i][21] = "True";
                    lastone.Rows[2 * i+1][21] = "True";
                    lastone.Rows[2 * i][22] = "True";
                    lastone.Rows[2 * i + 1][22] = "True";


                    lastone.Rows[2 * i][24] = "https://www.nurdtymedesigners.com/coasters/" + url + ".jpg";
                    //lastone.Rows[2 * i + 1][24] = "https://www.nurdtymedesigners.com/coasters/" + lastone.Rows[2 * i][1] + " No Background By " + lastone.Rows[2 * i][3] + ".jpg";
                    lastone.Rows[2 * i][25] = 1;
                  
                    lastone.Rows[2 * i][27] = "FALSE";
                    //lastone.Rows[2*i][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                    lastone.Rows[2 * i][30] = "Home & Garden > Kitchen & Dining > Barware > Coasters";
                    lastone.Rows[2 * i][31] = "Unisex";
                    lastone.Rows[2 * i][32] = "adult";
                    lastone.Rows[2 * i][36] = "New";
                    lastone.Rows[2 * i][37] = "TRUE";
                    lastone.Rows[2 * i][44] = "lb";
                    lastone.Rows[2 * i+1][44] = "lb";
                    lastone.Rows[2 * i][47] = "active";

                }

            }
            catch
            {
                MessageBox.Show("Are you sure to choose right xls file");
            }
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                try
                {
                    lastone = add_Tags(lastone, ReadTags("Sheet1", textBox2.Text));
                }
                catch { }
            }
            // write to csv file
            dataGridView1.DataSource = lastone;
            ToCsvFile(lastone, textBox1.Text);

        }

        private void Tumblers()
        {
            string[] headers = header();
            DataTable dt = new DataTable();
            DataTable lastone = new DataTable();
            if (textBox1.Text != "")
            {
                dt = ReadExcelFile("Sheet1", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Choose an excel File!");
                return;
            }
            DataColumn col = dt.Columns[0];


            for (int i = 0; i < headers.Length; i++)
            {
                lastone.Columns.Add(headers[i]);
            }

            try
            {


                lastone.Rows.Add();
                lastone.Rows[0][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');

                if (lastone.Rows[0][0].ToString().IndexOf("Tumbler") < 0)
                {
                    MessageBox.Show("Wrong File");
                    return;
                }
                string contitle2 = lastone.Rows[0][0].ToString().Replace('-', ' ');
                string url = contitle2;
                string title2 = contitle2;
                int index2 = title2.IndexOf("Tumbler") + 7;
                if (index2 > 7)
                {
                    title2 = title2.Substring(0, index2);
                }
                lastone.Rows[0][1] = title2;

                index2 = contitle2.IndexOf("By") + 3;
                if (index2 > 3)
                {
                    contitle2 = contitle2.Substring(index2, contitle2.Length - index2);
                }
                lastone.Rows[0][2] = "<p>This is a double-walled, stainless steel tumbler with a 20 oz. capacity.Â </p>\n"+
                "< p >< span > This stainless steel tumbler is created using sublimation.The images are infused onto the tumbler using heat.No epoxy is " +
                "necessary </ span ></ p >\n< p >< span > These images do not fade if cared for correctly.</ span ></ p >\n"+
                "< ul >\n< li >\n< p >< span > Vacuum insulated: 20 oz.tumblers are double - walled and vacuum insulated, which keeps your favorite beverage" +
                " hot or cold for hours </ span ></ p >\n</ li >\n< li >< p >< span > Sealable lid and plastic straw included </ span ></ p ></ li >"+
              "\n< li >< p >< span > BPA - free lid - The eco - friendly lid is completely BPA - free; silica gasket seals to achieve maximum spill - " +
              "proof capability.The splash proofÂ sliding cover, strawÂ </ span > friendly </ p ></ li >"+
              "\n< li > It will keep hot drinks warm for up to 6 hours and cold beverages will stay that way for more than 12 hours.Put ice in it in the" +
              " evening and it will still be there the next morning! </ li >\n</ ul >\n< p > Disclaimer: This listing is for only 1 tumbler.Image with" +
              " poster is not included.It's just an example of the image wrapped around the tumbler.</p>\n< p > Â </ p > ";

                for (int i = 0; i < 2; i++)
                {
                    lastone.Rows.Add();
                    lastone.Rows[i][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');

                }


                lastone.Rows[0][3] = contitle2;
                lastone.Rows[0][5] = "20oz Tumbler - AOP";
                lastone.Rows[0][7] = "TRUE";
                lastone.Rows[0][8] = "Title";
                lastone.Rows[0][9] = "Default Title";
               


                lastone.Rows[0][15] = 454.0006031;
                lastone.Rows[0][16] = "shopify";
                lastone.Rows[0][17] = "continue";
                lastone.Rows[0][18] = "manual";
                lastone.Rows[0][19] = 19.99;
                lastone.Rows[0][21] = "True";


                lastone.Rows[0][24] = "https://www.nurdtymedesigners.com/Tumblers/" + url + ".jpg";
                lastone.Rows[1][24] = "https://www.nurdtymedesigners.com/Tumblers/" + lastone.Rows[0][1] + " No Background By " + lastone.Rows[0][3] + ".jpg";
                lastone.Rows[0][25] = 1;
                lastone.Rows[1][25] = 2;
                lastone.Rows[0][27] = "FALSE";
                //lastone.Rows[0][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                lastone.Rows[0][30] = "Home & Garden > Kitchen & Dining > Tableware > Drinkware > Tumblers";
                lastone.Rows[0][31] = "Unisex";
                lastone.Rows[0][32] = "adult";
                lastone.Rows[0][36] = "New";
                lastone.Rows[0][37] = "TRUE";
                lastone.Rows[0][44] = "lb";
                lastone.Rows[0][47] = "active";

                //--------------------------------------------------------------

                for (int i = 1; i < dt.Rows.Count + 1; i++)
                {
                    lastone.Rows.Add();
                    string handle = dt.Rows[i - 1][0].ToString();

                    string contitle = handle.Substring(0, handle.Length - 4).Replace('-', ' ');
                    string title = contitle;
                    url = contitle;
                    int index = title.IndexOf("Tumbler") + 7;
                    if (index > 7)
                    {
                        title = title.Substring(0, index);
                    }
                    lastone.Rows[2 * i][1] = title;

                    index = contitle.IndexOf("By") + 3;
                    if (index > 3)
                    {
                        contitle = contitle.Substring(index, contitle.Length - index);
                    }
                    lastone.Rows[2 * i][2] = "<p>This is a double-walled, stainless steel tumbler with a 20 oz. capacity.Â </p>\n" +
                "< p >< span > This stainless steel tumbler is created using sublimation.The images are infused onto the tumbler using heat.No epoxy is " +
                "necessary </ span ></ p >\n< p >< span > These images do not fade if cared for correctly.</ span ></ p >\n" +
                "< ul >\n< li >\n< p >< span > Vacuum insulated: 20 oz.tumblers are double - walled and vacuum insulated, which keeps your favorite beverage" +
                " hot or cold for hours </ span ></ p >\n</ li >\n< li >< p >< span > Sealable lid and plastic straw included </ span ></ p ></ li >" +
              "\n< li >< p >< span > BPA - free lid - The eco - friendly lid is completely BPA - free; silica gasket seals to achieve maximum spill - " +
              "proof capability.The splash proofÂ sliding cover, strawÂ </ span > friendly </ p ></ li >" +
              "\n< li > It will keep hot drinks warm for up to 6 hours and cold beverages will stay that way for more than 12 hours.Put ice in it in the" +
              " evening and it will still be there the next morning! </ li >\n</ ul >\n< p > Disclaimer: This listing is for only 1 tumbler.Image with" +
              " poster is not included.It's just an example of the image wrapped around the tumbler.</p>\n< p > Â </ p > ";



                    for (int j = 2 * i; j < 2 * i + 2; j++)
                    {
                        lastone.Rows.Add();
                        lastone.Rows[j][0] = handle.Substring(0, handle.Length - 4).Replace(' ', '-');


                    }


                    lastone.Rows[2*i][3] = contitle;
                    lastone.Rows[2*i][5] = "20oz Tumbler - AOP";
                    lastone.Rows[2*i][7] = "TRUE";
                    lastone.Rows[2*i][8] = "Title";
                    lastone.Rows[2*i][9] = "Default Title";



                    lastone.Rows[2*i][15] = 454.0006031;
                    lastone.Rows[2*i][16] = "shopify";
                    lastone.Rows[2*i][17] = "continue";
                    lastone.Rows[2*i][18] = "manual";
                    lastone.Rows[2*i][19] = 19.99;
                    lastone.Rows[2*i][21] = "True";


                    lastone.Rows[2*i][24] = "https://www.nurdtymedesigners.com/Tumblers/" + url + ".jpg";
                    lastone.Rows[2*i+1][24] = "https://www.nurdtymedesigners.com/Tumblers/" + lastone.Rows[2*i][1] + " No Background By " + lastone.Rows[2*i][3] + ".jpg";
                    lastone.Rows[2*i][25] = 1;
                    lastone.Rows[2*i+1][25] = 2;
                    lastone.Rows[2*i][27] = "FALSE";
                    //lastone.Rows[2*i][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                    lastone.Rows[2*i][30] = "Home & Garden > Kitchen & Dining > Tableware > Drinkware > Tumblers";
                    lastone.Rows[2*i][31] = "Unisex";
                    lastone.Rows[2*i][32] = "adult";
                    lastone.Rows[2*i][36] = "New";
                    lastone.Rows[2*i][37] = "TRUE";
                    lastone.Rows[2*i][44] = "lb";
                    lastone.Rows[2*i][47] = "active";

                }
                
            }
            catch
            {
                MessageBox.Show("Are you sure to choose right xls file");
            }
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                try
                {
                    lastone = add_Tags(lastone, ReadTags("Sheet1", textBox2.Text));
                }
                catch { }
            }
            // write to csv file
            dataGridView1.DataSource = lastone;
            ToCsvFile(lastone,textBox1.Text);
        }

        private void CuttingBoard()
        {
            string[] headers = header();
            DataTable dt = new DataTable();
            DataTable lastone = new DataTable();
            if (textBox1.Text != "")
            {
                dt = ReadExcelFile("Sheet1", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Choose an excel File!");
                return;
            }
            DataColumn col = dt.Columns[0];


            for (int i = 0; i < headers.Length; i++)
            {
                lastone.Columns.Add(headers[i]);
            }

            try
            {


                lastone.Rows.Add();
                lastone.Rows[0][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');
               
                if (lastone.Rows[0][0].ToString().IndexOf("Cutting-Board") < 0)
                {
                    MessageBox.Show("Wrong File");
                    return;
                }
                string contitle2 = lastone.Rows[0][0].ToString().Replace('-', ' ');
                string url = contitle2;
                string title2 = contitle2;
                int index2 = title2.IndexOf("Cutting Board") + 13;
                if (index2 > 13)
                {
                    title2 = title2.Substring(0, index2);
                }
                lastone.Rows[0][1] = title2;

                index2 = contitle2.IndexOf("By") + 3;
                if (index2 > 3)
                {
                    contitle2 = contitle2.Substring(index2, contitle2.Length - index2);
                }
                lastone.Rows[0][2] = "<ul data-mce-fragment='1'>\n< li data - mce - fragment = '1' >\n< span data - mce - fragment = '1' > Geek up your kitchen" +
                    " with a cutting board.</ span >\n</ li >\n< li data - mce - fragment = '1' >\n< span data - mce - fragment = '1' >" +
                    " These also make great wedding or housewarming gifts! </ span >\n</ li >"+
                    "\n< li data - mce - fragment = '1' >< span data - mce - fragment = '1' > Makes a great accent piece.</ span ></ li >" +
                    "\n< li data - mce - fragment = '1' > Can also be used as a serving tray</ li >\n</ ul >\n"+
                    "< p data - mce - fragment = '1' > Geek up your kitchen with a cutting board.These also make great wedding or housewarming gifts!" +
                    "Cutting boards are made of textured, tempered glass with rubber feet to keep them from sliding around on your counter.Dishwasher safe." +
                    " </ p >\n< p data - mce - fragment = '1' >< br data - mce - fragment = '1' > Hand washing recommended.< br data - mce - fragment = '1' >" +
                    " Rubber feet included with every cutting board.< br data - mce - fragment = '1' ></ p > ";

                for (int i = 0; i < 2; i++)
                {
                    lastone.Rows.Add();
                    lastone.Rows[i][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');

                    
                    lastone.Rows[i][16] = "shopify";
                    lastone.Rows[i][17] = "deny";
                    lastone.Rows[i][18] = "manual";
                    lastone.Rows[i][21] = "TRUE";
                    lastone.Rows[i][22] = "TRUE";
                    lastone.Rows[i][44] = "lb";
                }


                lastone.Rows[0][3] = contitle2;
                lastone.Rows[0][5] = "Cutting Boards";
                lastone.Rows[0][7] = "TRUE";
                lastone.Rows[0][8] = "Size";
                lastone.Rows[0][9] = "8 x 11";
                lastone.Rows[1][9] = "11.25 x 15.5";
              

                lastone.Rows[0][15] = 453.59237;
                lastone.Rows[1][15] = 1360.78;
            

                lastone.Rows[0][19] = 25;
                lastone.Rows[1][19] = 35;
               

                lastone.Rows[0][24] = "https://www.nurdtymedesigners.com/cuttingboards/" + url + ".jpg";
                lastone.Rows[1][24] = "https://www.nurdtymedesigners.com/cuttingboards/" + lastone.Rows[0][1]+ " No Background By " + lastone.Rows[0][3]+ ".jpg";
                lastone.Rows[0][25] = 1;
                lastone.Rows[1][25] = 2;
                lastone.Rows[0][27] = "FALSE";
                //lastone.Rows[0][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                lastone.Rows[0][30] = 666;
                lastone.Rows[0][31] = "Unisex";
                lastone.Rows[0][32] = "adult";
                lastone.Rows[0][36] = "New";
                lastone.Rows[0][37] = "FALSE";

                lastone.Rows[0][47] = "active";

                //--------------------------------------------------------------

                for (int i = 1; i < dt.Rows.Count+1; i++)
                {
                    lastone.Rows.Add();
                    string handle = dt.Rows[i - 1][0].ToString();

                    string contitle = handle.Substring(0, handle.Length - 4).Replace('-', ' ');
                    string title = contitle;
                    url = contitle;
                    int index = title.IndexOf("Cutting Board") + 13;
                    if (index > 13)
                    {
                        title = title.Substring(0, index);
                    }
                    lastone.Rows[2 * i][1] = title;

                    index = contitle.IndexOf("By") + 3;
                    if (index > 3)
                    {
                        contitle = contitle.Substring(index, contitle.Length - index);
                    }
                    lastone.Rows[2*i][2] = "<ul data-mce-fragment='1'>\n< li data - mce - fragment = '1' >\n< span data - mce - fragment = '1' > Geek up your kitchen" +
                    " with a cutting board.</ span >\n</ li >\n< li data - mce - fragment = '1' >\n< span data - mce - fragment = '1' >" +
                    " These also make great wedding or housewarming gifts! </ span >\n</ li >" +
                    "\n< li data - mce - fragment = '1' >< span data - mce - fragment = '1' > Makes a great accent piece.</ span ></ li >" +
                    "\n< li data - mce - fragment = '1' > Can also be used as a serving tray</ li >\n</ ul >\n" +
                    "< p data - mce - fragment = '1' > Geek up your kitchen with a cutting board.These also make great wedding or housewarming gifts!" +
                    "Cutting boards are made of textured, tempered glass with rubber feet to keep them from sliding around on your counter.Dishwasher safe." +
                    " </ p >\n< p data - mce - fragment = '1' >< br data - mce - fragment = '1' > Hand washing recommended.< br data - mce - fragment = '1' >" +
                    " Rubber feet included with every cutting board.< br data - mce - fragment = '1' ></ p > ";


                    for (int j = 2 * i; j < 2 * i + 2; j++)
                    {
                        lastone.Rows.Add();
                        lastone.Rows[j][0] = handle.Substring(0, handle.Length - 4).Replace(' ', '-');


                        lastone.Rows[j][16] = "shopify";
                        lastone.Rows[j][17] = "deny";
                        lastone.Rows[j][18] = "manual";
                        lastone.Rows[j][21] = "TRUE";
                        lastone.Rows[j][22] = "TRUE";
                        lastone.Rows[j][44] = "lb";
                    }


                    lastone.Rows[2*i][3] = contitle;
                    lastone.Rows[2*i][5] = "Cutting Boards";
                    lastone.Rows[2*i][7] = "TRUE";
                    lastone.Rows[2*i][8] = "Size";
                    lastone.Rows[2*i][9] = "8 x 11";
                    lastone.Rows[2 * i + 1][9] = "11.25 x 15.5";


                    lastone.Rows[2*i][15] = 453.59237;
                    lastone.Rows[2 * i + 1][15] = 1360.78;


                    lastone.Rows[2*i][19] = 25;
                    lastone.Rows[2*i+1][19] = 35;


                    lastone.Rows[2*i][24] = "https://www.nurdtymedesigners.com/cuttingboards/" + url + ".jpg";
                    lastone.Rows[2*i+1][24] = "https://www.nurdtymedesigners.com/cuttingboards/" + lastone.Rows[2*i][1] + " No Background By " + lastone.Rows[2*i][3] + ".jpg";
                    lastone.Rows[2*i][25] = 1;
                    lastone.Rows[2*i+1][25] = 2;
                    lastone.Rows[2*i][27] = "FALSE";
                    //lastone.Rows[2*i][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                    lastone.Rows[2*i][30] = 666;
                    lastone.Rows[2*i][31] = "Unisex";
                    lastone.Rows[2*i][32] = "adult";
                    lastone.Rows[2*i][36] = "New";
                    lastone.Rows[2*i][37] = "FALSE";

                    lastone.Rows[2*i][47] = "active";

                }
                dataGridView1.DataSource = lastone;
            }
            catch
            {
                MessageBox.Show("Are you sure to choose right xls file");
            }
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                try
                {
                    lastone = add_Tags(lastone, ReadTags("Sheet1", textBox2.Text));
                }
                catch { }
            }
            // write to csv file
            ToCsvFile(lastone,textBox1.Text);
        }

        private void HoodedBlanket()
        {
            string[] headers = header();
            DataTable dt = new DataTable();
            DataTable lastone = new DataTable();
            if (textBox1.Text != "")
            {
                dt = ReadExcelFile("Sheet1", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Choose an excel File!");
                return;
            }
            DataColumn col = dt.Columns[0];


            for (int i = 0; i < headers.Length; i++)
            {
                lastone.Columns.Add(headers[i]);
            }

            try
            {


                lastone.Rows.Add();
                lastone.Rows[0][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');
                if (lastone.Rows[0][0].ToString().IndexOf("Hooded-Blanket") < 0)
                {
                    MessageBox.Show("Wrong File");
                    return;
                }
                string contitle2 = lastone.Rows[0][0].ToString().Replace('-', ' ');
                string url = contitle2;
                string title2 = contitle2;
                int index2 = title2.IndexOf("Hooded Blanket") + 14;
                if (index2 > 14)
                {
                    title2 = title2.Substring(0, index2);
                }
                lastone.Rows[0][1] = title2;

                index2 = contitle2.IndexOf("By") + 3;
                if (index2 > 3)
                {
                    contitle2 = contitle2.Substring(index2, contitle2.Length - index2);
                }
                lastone.Rows[0][2] = "<p>This hooded blanket is crafted from an outer silky smooth micro-mink polyester face. For the inside you can choose"+
                    "from an ultra soft microfiber fleece lining, or a premium plush 100% polyester sherpa lining.</p>\n< p >< strong >" +
                    " **Images cut out on the mockup will be correctly sized on the actual product.</ strong ></ p >\n< p >" +
                    " Each hooded blanket is individually printed, cut and sewn to ensure a flawless graphic with no imperfections.</ p >\n"+
                "< p > Outer Fabric: Micro - mink 100 % polyester \n< br > Lining: Ultra Soft Microfiber Fleece / Premium Plush 100 % Polyester Sherpa<br>\n" +
                " Ultra Soft handfeel<br>  High definition printing colors< br > Printed using Sublimation -Each one is uniquely crafted just for you! </ p >\n" +
              "< p > Machine wash cold cycle </ p >\n< p > Because it's handmade for you, these hooded blankets require 6-8 business days before " +
              "they are shipped. Orders placed before midnight will be included in the following day's batch for manufacturing.</ p > ";

                for (int i = 0; i < 4; i++)
                {
                    lastone.Rows.Add();
                    lastone.Rows[i][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');

                    lastone.Rows[i][15] = "0";
                    lastone.Rows[i][16] = "shopify";
                    lastone.Rows[i][17] = "continue";
                    lastone.Rows[i][18] = "subliminator";
                    lastone.Rows[i][21] = "TRUE";
                    lastone.Rows[i][22] = "TRUE";
                    lastone.Rows[i][44] = "oz";
                }


                lastone.Rows[0][3] = contitle2;
                lastone.Rows[0][5] = "Hooded Blanket - AOP";
                lastone.Rows[0][7] = "TRUE";
                lastone.Rows[0][8] = "Size";
                lastone.Rows[0][9] = "Adult";
                lastone.Rows[1][9] = "Adult";
                lastone.Rows[2][9] = "Youth";
                lastone.Rows[3][9] = "Youth";
                lastone.Rows[0][10] = "Type";
                lastone.Rows[0][11] = "Premium Sherpa";
                lastone.Rows[1][11] = "Micro Fleece";
                lastone.Rows[2][11] = "Premium Sherpa";
                lastone.Rows[3][11] = "Micro Fleece";

                lastone.Rows[0][19] = 65;
                lastone.Rows[1][19] = 60;
                lastone.Rows[2][19] = 55;
                lastone.Rows[3][19] = 50;

                lastone.Rows[0][24] = "https://www.nurdtymedesigners.com/hoodedblanket/" + url + ".jpg";
                lastone.Rows[0][25] = 1;
                lastone.Rows[0][27] = "False";
                //lastone.Rows[0][30] = "Electronics > Electronics Accessories > Computer Accessories > Mouse Pads";
                //lastone.Rows[0][31] = "Unisex";
                //lastone.Rows[0][32] = "Adult";
                //lastone.Rows[0][36] = "New";
                //lastone.Rows[0][37] = "TRUE";

                lastone.Rows[0][47] = "active";

                for (int i = 1; i < dt.Rows.Count+1; i++)
                {
                    lastone.Rows.Add();
                    string handle = dt.Rows[i - 1][0].ToString();

                    string contitle = handle.Substring(0, handle.Length - 4).Replace('-', ' ');
                    string title = contitle;
                    url = contitle;
                    int index = title.IndexOf("Hooded Blanket") + 14;
                    if (index > 9)
                    {
                        title = title.Substring(0, index);
                    }
                    lastone.Rows[4 * i][1] = title;

                    index = contitle.IndexOf("By") + 3;
                    if (index > 3)
                    {
                        contitle = contitle.Substring(index, contitle.Length - index);
                    }
                    lastone.Rows[4*i][2] = "<p>This hooded blanket is crafted from an outer silky smooth micro-mink polyester face. For the inside you can choose" +
                    "from an ultra soft microfiber fleece lining, or a premium plush 100% polyester sherpa lining.</p>\n< p >< strong >" +
                    " **Images cut out on the mockup will be correctly sized on the actual product.</ strong ></ p >\n< p >" +
                    " Each hooded blanket is individually printed, cut and sewn to ensure a flawless graphic with no imperfections.</ p >\n" +
                "< p > Outer Fabric: Micro - mink 100 % polyester \n< br > Lining: Ultra Soft Microfiber Fleece / Premium Plush 100 % Polyester Sherpa<br>\n" +
                " Ultra Soft handfeel<br>  High definition printing colors< br > Printed using Sublimation -Each one is uniquely crafted just for you! </ p >\n" +
              "< p > Machine wash cold cycle </ p >\n< p > Because it's handmade for you, these hooded blankets require 6-8 business days before " +
              "they are shipped. Orders placed before midnight will be included in the following day's batch for manufacturing.</ p > ";


                    for (int j = 4 * i; j < 4 * i + 4; j++)
                    {
                        lastone.Rows.Add();
                        lastone.Rows[j][0] = handle.Substring(0, handle.Length - 4).Replace(' ', '-');

                        lastone.Rows[j][15] = "0";
                        lastone.Rows[j][16] = "shopify";
                        lastone.Rows[j][17] = "continue";
                        lastone.Rows[j][18] = "subliminator";
                        lastone.Rows[j][21] = "TRUE";
                        lastone.Rows[j][22] = "TRUE";
                        lastone.Rows[j][44] = "oz";
                    }


                    lastone.Rows[4*i][3] = contitle;
                    lastone.Rows[4*i][5] = "Hooded Blanket - AOP";
                    lastone.Rows[4*i][7] = "TRUE";
                    lastone.Rows[4*i][8] = "Size";

                    lastone.Rows[4*i][9] = "Adult";
                    lastone.Rows[4*i+1][9] = "Adult";
                    lastone.Rows[4 * i + 2][9] = "Youth";
                    lastone.Rows[4 * i + 3][9] = "Youth";
                    lastone.Rows[4*i][10] = "Type";
                    lastone.Rows[4*i][11] = "Premium Sherpa";
                    lastone.Rows[4 * i+1][11] = "Micro Fleece";
                    lastone.Rows[4 * i + 2][11] = "Premium Sherpa";
                    lastone.Rows[4 * i + 3][11] = "Micro Fleece";

                    lastone.Rows[4*i][19] = 65;
                    lastone.Rows[4 * i + 1][19] = 60;
                    lastone.Rows[4 * i + 2][19] = 55;
                    lastone.Rows[4 * i + 3][19] = 50;

                    lastone.Rows[4*i][24] = "https://www.nurdtymedesigners.com/hoodedblanket/" + url + ".jpg";
                    lastone.Rows[4*i][25] = 1;
                    lastone.Rows[4*i][27] = "False";
                    //lastone.Rows[0][29] = "Use this mouse pad decorate your home computer or use at work! It makes a great gift for yourself or someone else. Not only will our placemats keep";
                    //lastone.Rows[0][30] = "Electronics > Electronics Accessories > Computer Accessories > Mouse Pads";
                    //lastone.Rows[0][31] = "Unisex";
                    //lastone.Rows[0][32] = "Adult";
                    //lastone.Rows[0][36] = "New";
                    //lastone.Rows[0][37] = "TRUE";

                    lastone.Rows[4 * i][47] = "active";

                }
                dataGridView1.DataSource = lastone;
            }
            catch
            {
                MessageBox.Show("Are you sure to choose right xls file");
            }

            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                try
                {
                    lastone = add_Tags(lastone, ReadTags("Sheet1", textBox2.Text));
                }
                catch { }
            }
            // write to csv file
            ToCsvFile(lastone,textBox1.Text);
        }

        private void Carseat()
        {
            string[] headers = header();
            DataTable dt = new DataTable();
            DataTable lastone = new DataTable();
            if (textBox1.Text != "")
            {
                dt = ReadExcelFile("Sheet1", textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Choose an excel File!");
                return;
            }
            DataColumn col = dt.Columns[0];


            for (int i = 0; i < headers.Length; i++)
            {
                lastone.Columns.Add(headers[i]);
            }

            try
            {


                lastone.Rows.Add();
                lastone.Rows[0][0] = col.ColumnName.Substring(0, col.ColumnName.Length - 4).Replace(' ', '-');
                if (lastone.Rows[0][0].ToString().IndexOf("Car-Seat") < 0)
                {
                    MessageBox.Show("Wrong File");
                    return;
                }
                string contitle2 = lastone.Rows[0][0].ToString().Replace('-', ' ');
                string url = contitle2;
                string title2 = contitle2;
                int index2 = title2.IndexOf("Car Seat Covers") +15 ;
                if (index2 > 15)
                {
                    title2 = title2.Substring(0, index2);
                }
                lastone.Rows[0][1] = title2;

                index2 = contitle2.IndexOf("By") + 3;
                if (index2 > 3)
                {
                    contitle2 = contitle2.Substring(index2, contitle2.Length - index2);
                }
                lastone.Rows[0][2] = "<style>\n.sbl - size - table {\nborder - collapse: collapse;\n padding: 0;\nmargin: 0 0 20px;\nwidth: 100 %;\nfont - size: 14px;\n" +
                    "text - align: center; \n }\n .sbl - size - table th {\nfont - weight: 500;\n}\n.sbl - size - table td,.sbl - size - table th {\n" +
                "padding: 8px 0;\nborder: 1px solid #e5e9f2;\n color: #3e3f42;\n text - shadow: 1px 1px 1px #fff;\ntext -align: center;\n}\n" +
                ".sbl - size - table th: first - child,.sbl - size - table td: first - child {\n text - align: left;\n padding: 8px 5px 8px 15px;\n width: 103px;\n}" +
                " .sbl - size - guide - container {\nwidth: 100 %;\ntext - align: center;\nmargin - bottom: 20px;\nmargin - top: 20px;\n}\n" +
                ".sbl - size - guide - container img {\nmax - width: 200px;\n;margin: auto;\n}\n</ style >\n" +
                "< div class='sbl-description'>\n< div class='subl-product-description' style='max-width: 100%'>\n" +
                "< p>Keep your car seats clean from spills, stains, tearing and fading, while adding your own personal touch and style to your car seats. </p>\n" +
                "<p>\nâ€¢ Fabric: 100% Microfiber Polyester<br>\nâ€¢ Quick and easy installation on most car and SUV bucket style seats<br>\n" +
                "â€¢ No tools required for installation<br>\nâ€¢ Not for use on seats with integrated airbags, seatbelts and armrests.<br>\n" +
                "â€¢ High definition printing colors<br>\nâ€¢ Printed, cut, and hand-sewn by our in-house team<br>\n</p>\n</div>\n< br>" +
                "Because itâ€™s handmade for you, these car seats covers require 6-8 business days before they are shipped. Orders placed before midnight " +
                "will be included in the following day's batch for manufacturing.\n< br>\n< div class='sbl-size-guide-container'>" +
                "<img src = 'https://static.subliminator.com/shops/images/size-guides/seat-cover.png' class='sbl-size-guide-image' alt='Eleven Car Seat Covers'>\n" +
                "</div>\n</div>";
                lastone.Rows[0][3] = contitle2;
                lastone.Rows[0][5] = "Car Seat Cover AOP";
                lastone.Rows[0][7] = "TRUE";
                lastone.Rows[0][8] = "Size";
                lastone.Rows[0][9] = "One size";
                lastone.Rows[0][15] = "0";
                lastone.Rows[0][16] = "shopify";
                lastone.Rows[0][17] = "continue";
                lastone.Rows[0][18] = "subliminator";
                lastone.Rows[0][19] = 64.99;
                lastone.Rows[0][21] = "TRUE";
                lastone.Rows[0][22] = "TRUE";
                lastone.Rows[0][24] = "https://www.nurdtymedesigners.com/carseatcovers/" + url + ".jpg";
                lastone.Rows[0][25] = "1";//added
        //removed
                lastone.Rows[0][30] = "2495";
                lastone.Rows[0][31] = "Unisex";
                lastone.Rows[0][32] = "Adult";
                lastone.Rows[0][36] = "New";
                lastone.Rows[0][37] = "TRUE";
                lastone.Rows[0][44] = "oz";
                lastone.Rows[0][47] = "active";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    lastone.Rows.Add();
                    string handle = dt.Rows[i][0].ToString();
                    lastone.Rows[i + 1][0] = handle.Substring(0, handle.Length - 4).Replace(' ', '-');
                    string contitle = lastone.Rows[i + 1][0].ToString().Replace('-', ' ');
                    string title = contitle;
                    url = contitle;
                    int index = title.IndexOf("Car Seat Covers") + 15;
                    if (index > 15)
                    {
                        title = title.Substring(0, index);
                    }
                    lastone.Rows[i + 1][1] = title;

                    index = contitle.IndexOf("By") + 3;
                    if (index > 3)
                    {
                        contitle = contitle.Substring(index, contitle.Length - index);
                    }
                    lastone.Rows[i + 1][2] = "<style>\n.sbl - size - table {\nborder - collapse: collapse;\n padding: 0;\nmargin: 0 0 20px;\nwidth: 100 %;\nfont - size: 14px;\n" +
                    "text - align: center; \n }\n .sbl - size - table th {\nfont - weight: 500;\n}\n.sbl - size - table td,.sbl - size - table th {\n" +
                "padding: 8px 0;\nborder: 1px solid #e5e9f2;\n color: #3e3f42;\n text - shadow: 1px 1px 1px #fff;\ntext -align: center;\n}\n" +
                ".sbl - size - table th: first - child,.sbl - size - table td: first - child {\n text - align: left;\n padding: 8px 5px 8px 15px;\n width: 103px;\n}" +
                " .sbl - size - guide - container {\nwidth: 100 %;\ntext - align: center;\nmargin - bottom: 20px;\nmargin - top: 20px;\n}\n" +
                ".sbl - size - guide - container img {\nmax - width: 200px;\n;margin: auto;\n}\n</ style >\n" +
                "< div class='sbl-description'>\n< div class='subl-product-description' style='max-width: 100%'>\n" +
                "< p>Keep your car seats clean from spills, stains, tearing and fading, while adding your own personal touch and style to your car seats. </p>\n" +
                "<p>\nâ€¢ Fabric: 100% Microfiber Polyester<br>\nâ€¢ Quick and easy installation on most car and SUV bucket style seats<br>\n" +
                "â€¢ No tools required for installation<br>\nâ€¢ Not for use on seats with integrated airbags, seatbelts and armrests.<br>\n" +
                "â€¢ High definition printing colors<br>\nâ€¢ Printed, cut, and hand-sewn by our in-house team<br>\n</p>\n</div>\n< br>" +
                "Because itâ€™s handmade for you, these car seats covers require 6-8 business days before they are shipped. Orders placed before midnight " +
                "will be included in the following day's batch for manufacturing.\n< br>\n< div class='sbl-size-guide-container'>" +
                "<img src = 'https://static.subliminator.com/shops/images/size-guides/seat-cover.png' class='sbl-size-guide-image' alt='Eleven Car Seat Covers'>\n" +
                "</div>\n</div>";
                    lastone.Rows[i+1][3] = contitle;
                    lastone.Rows[i+1][5] = "Car Seat Cover AOP";
                    lastone.Rows[i+1][7] = "TRUE";
                    lastone.Rows[i+1][8] = "Size";
                    lastone.Rows[i+1][9] = "One size";
                    lastone.Rows[i+1][15] = "0";
                    lastone.Rows[i+1][16] = "shopify";
                    lastone.Rows[i+1][17] = "continue";
                    lastone.Rows[i+1][18] = "subliminator";
                    lastone.Rows[i+1][19] = 64.99;
                    lastone.Rows[i+1][21] = "TRUE";
                    lastone.Rows[i+1][22] = "TRUE";
                    lastone.Rows[i+1][24] = "https://www.nurdtymedesigners.com/carseatcovers/" + url + ".jpg";
                    lastone.Rows[i+1][25] = "1";//added
                                              //removed
                    lastone.Rows[i+1][30] = "2495";
                    lastone.Rows[i+1][31] = "Unisex";
                    lastone.Rows[i+1][32] = "Adult";
                    lastone.Rows[i+1][36] = "New";
                    lastone.Rows[i+1][37] = "TRUE";
                    lastone.Rows[i+1][44] = "oz";
                    lastone.Rows[i+1][47] = "active";

                }
                dataGridView1.DataSource = lastone;
            }
            catch
            {
                MessageBox.Show("Are you sure to choose right xls file");
            }
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                try
                {
                    lastone = add_Tags(lastone, ReadTags("Sheet1", textBox2.Text));
                }
                catch { }
            }
            // write to csv file
            ToCsvFile(lastone,textBox1.Text);
        }

        private DataTable ReadExcelFile(string sheetName, string path)
        {
            DataTable dt = new DataTable();
            using (OleDbConnection conn = new OleDbConnection())
            {
                
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                {
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                }
                    
                else if (fileExtension == ".xlsx")
                {
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                }
                else
                {
                    MessageBox.Show("Please Choose an EXCEL file! Not " + fileExtension + " File");
                    return null;
                }
                   
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        
                        da.Fill(dt);
                       
                    }
                }
            }
            return dt;
        }
         private DataTable ReadTags(string sheetName, string path)
        {
            ////string[] headers = { "A", "B" };
            DataTable dt = new DataTable();
            //for (int i = 0; i < headers.Length; i++)
            //{
            //    dt.Columns.Add(headers[i]);
            //}
            using (OleDbConnection conn = new OleDbConnection())
            {

                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                {
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties=" +
                        "'Excel 8.0;HDR=NO;'";
                }

                else if (fileExtension == ".xlsx" || fileExtension==".csv")
                {
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties=" +
                        "'Excel 12.0 Xml;HDR=NO;'";
                }
                else
                {
                    MessageBox.Show("Please Choose an EXCEL file! Not " + fileExtension + " File");
                    return null;
                }
                
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;

                        da.Fill(dt);

                    }
                }
            }
            return dt;
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //file load button

            string filePath;
            OpenFileDialog op = new OpenFileDialog();
            DialogResult result = op.ShowDialog();
            if (result == DialogResult.OK)
            {
                filePath = op.FileName;
                if (tabControl1.SelectedTab == tabControl1.TabPages[0])
                {
                    textBox2.Text = filePath;
                }
                else
                {
                    textBox3.Text = filePath;
                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filePath;
            OpenFileDialog op = new OpenFileDialog();
            DialogResult result = op.ShowDialog();
            if (result == DialogResult.OK)
            {
                filePath = op.FileName;
              
                    textBox4.Text = filePath;
               

            }
        }
        private DataTable ReadCsv( string path)
        {
            string sConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.GetDirectoryName(path)+ ";Extended " +
                "Properties=\"Text;HDR=Yes;FMT=Delimited\"";

            string sConnectionString2 = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(path) + ";Extended " +
                "Properties=\"Text;HDR=Yes;FMT=Delimited\"";
            DataTable dt = new DataTable();
            try
            {
                OleDbConnection oleConn = new OleDbConnection(sConnectionString2);
                string sSQLQuery = "Select * From [" + Path.GetFileName(path) + "]";
                OleDbCommand oleCommand = new OleDbCommand(sSQLQuery, oleConn);
                OleDbDataAdapter oleAdapt = new OleDbDataAdapter(oleCommand);
               
                oleAdapt.Fill(dt);
            }
            catch
            {
                OleDbConnection oleConn = new OleDbConnection(sConnectionString);
                string sSQLQuery = "Select * From [" + Path.GetFileName(path) + "]";
                OleDbCommand oleCommand = new OleDbCommand(sSQLQuery, oleConn);
                OleDbDataAdapter oleAdapt = new OleDbDataAdapter(oleCommand);
               
                oleAdapt.Fill(dt);
            }
            
            dt = dt.Rows
                 .Cast<DataRow>()
                 .Where(row => !row.ItemArray.All(f => f is DBNull))
                 .CopyToDataTable();
            return dt;

        }

        private void button_AddTags_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox3.Text) || string.IsNullOrEmpty(textBox4.Text))
            {
                MessageBox.Show("Please Select source data and the csv file");
                return;
            }

            DataTable sourceTags = ReadTags("Sheet1", textBox3.Text);

            DataTable HTags = ReadCsv(textBox4.Text);
            DataTable dt = add_Tags(HTags, sourceTags);
            dataGridView2.DataSource = dt;
           ToCsvFile( dt,textBox4.Text);
        }

        private DataTable add_Tags(DataTable lastone, DataTable dt)
        {
            string kindword;
            if (tabControl1.SelectedTab == tabPage2)
            {
                var checkedButton = tabControl1.TabPages[0].Controls.OfType<RadioButton>()
                                      .FirstOrDefault(r => r.Checked);
                kindword = checkedButton.Text.Substring(0,checkedButton.Text.Length-1);
            }
            else
            {
                kindword = lastone.Rows[0][5].ToString().Substring(0, lastone.Rows[0][5].ToString().Length);
                if (kindword.IndexOf("AOP") > -1)
                {
                    kindword = kindword.Substring(0, kindword.Length - 6);
                }
            }
            List<string> names = new List<string> { "Large Mouse Pad", "Mouse Pad","MousePad","Mousepad", "Car Seat", 
                "Cutting Board", "Hooded Blanket", "Tumbler", "Coaster" };
            kindword = names.FirstOrDefault(s => kindword.IndexOf(s) != -1);
            if (kindword == "MousePad" || kindword=="Mousepad") { kindword = "Mouse Pad"; }
            
            for(int i=0;i<lastone.Rows.Count;i++)
            {
                string searchingWord = lastone.Rows[i].ItemArray[1].ToString();
                
                int a = searchingWord.IndexOf(kindword);
                if (a < 1) { continue; }
                searchingWord = searchingWord.Substring(0, a-1);
                try
                {
                   var x = (from s in dt.AsEnumerable()
                                where searchingWord.IndexOf(s.ItemArray[0].ToString()) > -1
                                select (s.ItemArray[1].ToString())).FirstOrDefault();
                    if (x == null)
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(lastone.Rows[i][6].ToString()))
                    {
                        lastone.Rows[i][6] =  x;
                    }
                    else
                    {
                        lastone.Rows[i][6] = lastone.Rows[i][6].ToString() + "," + x;
                    }
                   
                }
                catch
                {
                    continue;
                }
               
                           
            }

            return lastone;
        }


    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;   

namespace WindowsFormsApplication1
{
    
    
    
    public partial class Form1 : Form
    {


        string filename = null;   //initalize filename and message 
        string message = "intro.txt";
        

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

         DirectoryInfo dinfo = new DirectoryInfo(@"c:\\TrueBearingClientList");  //initalize two listboxes subject to written directory
         FileInfo[] Files = dinfo.GetFiles("*.txt");
         foreach (FileInfo file in Files)
    {
        listBox1.Items.Add(file.Name);
    }

         DirectoryInfo dinfo_1 = new DirectoryInfo(@"c:\\TBWI\\Prospects");
         FileInfo[] Files_1 = dinfo_1.GetFiles("*.txt");
         foreach (FileInfo file in Files_1)
         {
             listBox2.Items.Add(file.Name);
         }
        
        }








        private void button1_Click(object sender, EventArgs e)
        {
            if (filename == null)     
            {

            }
            
            else{
                grid1.Rows.Clear();
           


            string directory = string.Format("c:\\TrueBearingClientList\\{0}", filename);
            StreamReader salesdata = new StreamReader(directory);
            string linex;
            int r = 0;
            while ((linex = salesdata.ReadLine()) != null)
            {
                grid1.Rows.Add();
                for (int i = 0; i < 9; i++)
                {
                    string[] fieldsx = linex.Split(',');
                    
                    grid1.Rows[r].Cells[i+1].Value = Convert.ToString(fieldsx[i]); 

                }
                r++;

            }

            salesdata.Close();

}

            string edit_filename = filename.Replace(".txt", "");                 //create backup of data loaded
            string directory_1 = string.Format("c:\\TrueBearingClientList\\{0}_backup.txt", edit_filename);
            if (!(File.Exists(directory_1)))
            {

                FileStream fs = new FileStream(directory_1, FileMode.CreateNew, FileAccess.Write);
                StreamWriter tbwi = new StreamWriter(fs);

                int row = grid1.RowCount - 1;
                for (int i = 0; i < row; i++)
                {
                    int j = 1;
                    for (j = 1; j < 10; j++)
                    {
                        tbwi.Write(grid1.Rows[i].Cells[j].Value);
                        if (j < 9)
                        {
                            tbwi.Write(",");
                        }
                        if (j == 9)
                        {
                            tbwi.WriteLine();
                        }



                    }
                    j = 1;
                }
                tbwi.Close();


            }
        }

        private void button2_Click(object sender, EventArgs e) //updates changes to current file
        {
            string directory = string.Format("c:\\TrueBearingClientList\\{0}", filename);
            File.Delete(directory);
            FileStream fs = new FileStream(directory, FileMode.CreateNew, FileAccess.Write);
            StreamWriter tbwi = new StreamWriter(fs);

            int r = grid1.RowCount - 1;
            for (int i = 0; i < r; i++)
            {
                int j = 1;
                for (j = 1; j < 10; j++)
                {
                    tbwi.Write(grid1.Rows[i].Cells[j].Value);
                    if (j < 9)
                    {
                        tbwi.Write(",");
                    }
                    if (j == 9)
                    {
                        tbwi.WriteLine();
                    }
                    


                }
                j = 1;
            }
            tbwi.Close();

        }

  

        private void grid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)           //send message
        {
            
            string text = string.Format("c:\\TBWI\\Prospects\\{0}", message);
            string msg = File.ReadAllText(text);
            Outlook.Application app = new Outlook.Application();
           
           

            for (int s = 0; s < grid1.RowCount; s++)
            {
                
                
                int c = Convert.ToInt32(grid1.Rows[s].Cells[0].Value); 
                
                if (c == 1)
                {
                    string company = Convert.ToString(grid1.Rows[s].Cells[1].Value);
                    string contact = Convert.ToString(grid1.Rows[s].Cells[4].Value);
                    string address = company + " - " + contact;
                  
                    //listBox2.Items.Add(Convert.ToString(grid1.Rows[s].Cells[1].Value));
                    Outlook.MailItem mailitem = app.CreateItem(Outlook.OlItemType.olMailItem);
                    mailitem.Subject = "C/P Performance Reports, Route Advisories & Performance Monitoring";
                    mailitem.Body = address + msg;
                   
                    mailitem.To = Convert.ToString(grid1.Rows[s].Cells[6].Value);
                    mailitem.Send();


                }


            }



        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
            string filename_1 = listBox1.GetItemText(listBox1.SelectedItem);
            filename = filename_1.Replace(" ", "");
           
        }


        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {


            string message_1 = listBox2.GetItemText(listBox2.SelectedItem);
            message = message_1.Replace(" ", "");
            
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

       

       

       
    
    
    }
}

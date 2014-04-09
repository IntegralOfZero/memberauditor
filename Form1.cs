using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Web;
using System.Net;
using System.IO;

namespace MemberAuditor
{
    public partial class Form1 : Form
    {
        OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;" + @"data source=Blue Phoenix Members Database.mdb");


        //the adapter is specific to each database type, but their purpose is to present the 
        //data structures as generic DataSet objects, so it provides a layer of abstraction
        //between the database and this application
        OleDbDataAdapter adapter = new OleDbDataAdapter();

        DataSet ds = new DataSet("Members DataSet"); //parameter is a name for the dataset
        //Datasets store results from queries

        OleDbCommand command;
        OleDbCommandBuilder cb;

        DataTable dt;

        public Form1()
        {
            InitializeComponent();

            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Connection Failed");
            }

            //create a query (Events is the table from which to query from, and apparently this query
            //takes all data in the table
            command = new OleDbCommand("SELECT * from Members", conn);
            //attach it to the adapter
            adapter.SelectCommand = command;
            cb = new OleDbCommandBuilder(adapter);

            adapter.Fill(ds, "Members"); //connecting the dataset with the data from query
            //2nd parameter is the name of the table (name and table created here) contained within 
            //ds that gets filled with the data brought by the adapter's command

            dt = ds.Tables["Members"]; //numerical index or string access works
            // Console.WriteLine(ds.Tables[0].ToString());
        }

        HashSet<string> membersInDB = new HashSet<string>(); //contains names of members in the database
        HashSet<string> membersInSite = new HashSet<string>(); //in Jagex's clanmates section of clan website
        
        Dictionary<string, string> ranks = new Dictionary<string, string>(); //store member name and their rank

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (DataRow dr in dt.Rows)
            {
                membersInDB.Add((string)dr["UserName"]);
            }

        }

        private void btnLeft_Click(object sender, EventArgs e)
        {
           //see who has left by seeing who is in the database but no longer in the site

            HashSet<string> membersLeft = new HashSet<string>(membersInDB); //copy from the specified set
            membersLeft.ExceptWith(membersInSite);

            foreach(string str in membersLeft)
            {
                listBox1.Items.Add(str); //add only no matches
            }

            MessageBox.Show(listBox1.Items.Count.ToString());
        }

        private void btnInit_Click(object sender, EventArgs e)
        {
            //number of pages to read from site
            int num = Int32.Parse(txtNum.Text);

            //download each webpage from the rs clan site's clanmates section and read it, then write each to a separate file
            //the URL must be updated when a new page is added on the site
            for (int i = 1; i <= num; i++)
            {
                string page = "http://services.runescape.com/m=clan-hiscores/c=xYkesDp8PQ0/members.ws?clanId=22042&pageSize=15&ranking=-1&pageNum=" + i;
                WebRequest webRequest = WebRequest.Create(page);
                WebResponse webResponse = webRequest.GetResponse();
                StreamReader sr = new StreamReader(webResponse.GetResponseStream());

                FileStream fs = new FileStream(i + ".txt", FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(sr.ReadToEnd());
                sw.Close();
                fs.Close();
            }

            //reading each page that was downloaded
            string line; //initialize once instead of each loop
            string lineTwo = ""; //holds the name
            string lineThree; //holds the rank

            for (int i = 1; i <= num; i++)
            {
                StreamReader sr = new StreamReader(i + ".txt");

                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine();


                    if (line.Contains("<span class=\"name\">"))
                    {
                        int end = line.IndexOf("</span>");
                        lineTwo = line.Substring(19, end - 19);
                        byte[] lineTwoBytes = UTF8Encoding.UTF8.GetBytes(lineTwo); //must convert encoding due to certain chars (like space) not showing properly
                        Encoding.Convert(Encoding.UTF8, Encoding.ASCII,lineTwoBytes);
                        lineTwo = Encoding.ASCII.GetString(lineTwoBytes);
                        lineTwo = lineTwo.Replace("???", " "); //UTF8 black diamond -> ASCII triple ? -> space
                        membersInSite.Add(lineTwo);
                        //MessageBox.Show(lineTwo);

                       
                    }
                    
                    if (line.Contains("<span class=\"clanRank\">"))
                    {
                        int end = line.IndexOf("</span>");
                        lineThree = line.Substring(23, end - 23);
                       // MessageBox.Show(lineTwo + " " + lineThree);

                        if (!lineTwo.Equals("Name")) //Name is listed twice in our program
                        {
                            ranks.Add(lineTwo, lineThree);
                        }                    
                    }
                    
                }
               
            }
        }

        private void btnJoined_Click(object sender, EventArgs e)
        {
            //see who has joined by seeing which names in the site are not in the database

            HashSet<string> membersJoined = new HashSet<string>(membersInSite); //copy from the specified set
            membersJoined.ExceptWith(membersInDB);

            foreach (string str in membersJoined)
            {
                listBox2.Items.Add(str); //add only no matches
            }
        }

        private void btnRank_Click(object sender, EventArgs e)
        {
            foreach(KeyValuePair<string, string> mem in ranks)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["UserName"].Equals(mem.Key))
                    {
                       // MessageBox.Show(mem.Key + " " + mem.Value + dr["Rank"]);
                        dr["Rank"] = ReturnRank(mem.Value);
                    }
                }
            }

            adapter.Update(dt);
        }

        private string ReturnRank(string value)
        {
            if (value.Equals("Recruit"))
            {
                return "7   Recruit";
            }
            else if (value.Equals("Corporal"))
            {
                return "6   Corporal";
            }
            else if (value.Equals("Sergeant"))
            {
                return "5   Sergeant";
            }
            else if (value.Equals("Lieutenant"))
            {
                return "4   Lieutenant";
            }
            else if (value.Equals("Captain"))
            {
                return "3   Captain";
            }
            else if (value.Equals("General"))
            {
                return "2   General";
            }
            else if (value.Equals("Admin"))
            {
                return "1-6   Admin";
            }
            else if (value.Equals("Organiser"))
            {
                return "1-5   Organiser";
            }
            else if (value.Equals("Coordinator"))
            {
                return "1-4   Coordinator";
            }
            else if (value.Equals("Overseer"))
            {
                return "1-3   Overseer";
            }
            else if (value.Equals("Deputy Owner"))
            {
                return "1-2   Deputy Owner";
            }
            else if (value.Equals("Owner"))
            {
                return "1   Owner";
            }
            else
            {
                return "";
            }
        }

        private void cmsRightClick_Opening(object sender, CancelEventArgs e)
        {

        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string content = "";
            foreach(var item in listBox1.SelectedItems)
            {
                content += item.ToString() + "\r\n";
            }
            Clipboard.SetText(content);
        }

        private void cmsRightClick2_Opening(object sender, CancelEventArgs e)
        {

        }

        private void copyRightToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            string content = "";
            foreach (var item in listBox2.SelectedItems)
            {
                content += item.ToString() + "\r\n";
            }
            Clipboard.SetText(content);
        }
    }
}

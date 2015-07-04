using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Data.SqlClient;
//using MySql.Data.MySqlClient;
using System.Configuration;
using System.Diagnostics;
using System.Timers;
using System.Data.SqlClient;
namespace samidtime
{
    public partial class Form1 : Form
    {
        KeyboardHook k_hook;
       private SerialPort Sp = new SerialPort();
       public delegate void HandleInterfaceUpdataDelegate(string text); //委托，此为重点 
       public string mysqlcon = ConfigurationManager.AppSettings["mysql"];
       public String ID;
       public String timeS;
       public DateTime timeE;
       System.Timers.Timer timer = new System.Timers.Timer();
       public int flag = 0;
        public Form1()
        {
            InitializeComponent();
            k_hook = new KeyboardHook();
            k_hook.KeyDownEvent += new KeyEventHandler(hook_KeyDown);//钩住键按下
           
            button2.Enabled = false;

            timer.Elapsed += new ElapsedEventHandler(OnTimedEvent);





        }

        private void hook_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (int)Keys.F5 || e.KeyValue == (int)Keys.F6 || e.KeyValue == (int)Keys.F7)
            {

                
                SqlConnection  conn = new SqlConnection (mysqlcon);

                String sql = "select min(id) as newid ,timeS as times from samidtime where status='N'  group by times";
                string sql222 = "SELECT MAX(id) AS Expr1, ID1 FROM yiliangang where status='N' GROUP BY ID1 ORDER BY Expr1 DESC ";
                try
                {
                    SqlCommand con = new SqlCommand(sql, conn);
                    //SqlCommand con222 = new SqlCommand(sql222, conn);
                    conn.Open();
                    SqlDataReader DataReader = con.ExecuteReader();
                    //SqlDataReader DataReader222 = con222.ExecuteReader();
                    
                    
                    if (DataReader.HasRows)
                    {
                        DataReader.Read();
                        int id = DataReader.GetInt32(0);
                        ID = id.ToString();
                        timeS = DataReader.GetString(1);
                        conn.Close();

                        SqlConnection conn222 = new SqlConnection(mysqlcon);
                        SqlCommand con222 = new SqlCommand(sql222, conn222);
                        conn222.Open();
                        SqlDataReader DataReader222 = con222.ExecuteReader();
                        


                        if (DataReader222.HasRows)
                        {
                            DataReader222.Read();
                            int id222 = DataReader222.GetInt32(0);
                            string ID222 = id222.ToString();
                            string bianhao = DataReader222.GetString(1);
                            conn222.Close();

                            //if (!DataReader.IsDBNull(0) && !DataReader.IsDBNull(1))
                            //{


                                //ID = DataReader.GetString(0);
                                


                                //SqlConnection conn222 = new SqlConnection(mysqlcon);
                                //SqlCommand con222 = new SqlCommand(sql222, conn222);
                                //conn222.Open();
                                // SqlDataReader DataReader222 = con222.ExecuteReader();
                                // DataReader222.Read();
                               // DataReader222.Read();
                                
                                //conn222.Close();
                                

                                //计算时间
                                timeE = DateTime.Now;
                                TimeSpan timeAll = timeE.Subtract(Convert.ToDateTime(timeS));
                                //MessageBox.Show(timeS + ";" + timeE + ";" + timeAll);


                                String sql1 = "update  samidtime set timeE=' " + timeE.ToString("yyyy-MM-dd HH:mm:ss") + "',status='Y',sampleId='" + bianhao + "', timeAll='" + Convert.ToDouble(timeAll.TotalMinutes).ToString("0.0") + "' where id ='" + ID + "'";
                                String sql2 = "update  yiliangang set status='Y' where id ='" + ID222 + "'";
                                //String sql3 = "update  yiliangang set status='Y' where status='N'";
                                SqlConnection connU = new SqlConnection(mysqlcon); ;
                                connU.Open();
                                SqlCommand conU = new SqlCommand(sql1, connU);
                                SqlCommand conU2 = new SqlCommand(sql2, connU);
                                //SqlCommand conU3 = new SqlCommand(sql3, connU);
                                conU.ExecuteNonQuery();
                                conU2.ExecuteNonQuery();
                                //conU3.ExecuteNonQuery();
                                connU.Close();
                                refresh();   //刷新数据显示

                            }
                            else
                            {
                                //timeE = DateTime.Now;
                                //TimeSpan timeAll = timeE.Subtract(Convert.ToDateTime(timeS));
                                //string bianhao = "作废";
                                //String sql333 = "update  samidtime set timeE=' " + timeE.ToString("yyyy-MM-dd HH:mm:ss") + "',status='S',sampleId='" + bianhao + "', timeAll='" + Convert.ToDouble(timeAll.TotalMinutes).ToString("0.0") + "' where id ='" + ID + "'";
                                //SqlConnection conn333 = new SqlConnection(mysqlcon); ;
                                //conn333.Open();
                                //SqlCommand con333 = new SqlCommand(sql333, conn333);
                                //con333.ExecuteNonQuery();
                                //conn333.Close();
                                //refresh(); 
                            }


                        }

                    }
               // }
                catch
                {
                   
                }
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            k_hook.Stop();
            Sp.Close();
        }
        //运行
        private void button1_Click(object sender, EventArgs e)
        {
            if ((comboBox1.Text.Trim() != "") && (comboBox2.Text != ""))
             {
                 Sp.PortName = comboBox1.Text.Trim();
                 Sp.BaudRate = Convert.ToInt32(comboBox2.Text.Trim());
                 Sp.Parity = Parity.None;
                 Sp.StopBits = StopBits.One;
                 Sp.DataReceived += new SerialDataReceivedEventHandler(Sp_DataReceived);

                 Sp.ReceivedBytesThreshold = 1;
                try
                 {
                     Sp.Open();
                     
                     k_hook.Start();
                     button2.Enabled = true;
                     button1.Enabled = false;
                    label4.Text="运行中";
                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                 
                 //   String text = "01 03 01 0A 00 01 A5 F4";  //MD82
                    timer.Interval = 50;
                    timer.Start();       
                }
                catch
                 {
                     MessageBox.Show("端口" + comboBox1.Text.Trim() + "打开失败！");
                }
             }
            else
            {
                 MessageBox.Show("请选择正确的端口号和波特率！");

            }

        }
        //定时器任务
        void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
           
            byte[] b1 = { 0x01, 0x03, 0x01, 0x04, 0x00, 0x04, 0x04, 0x34 };
            //   String text = "01 03 01 04 00 04 04 34";  //MD82
           
            Sp.Write(b1, 0, b1.Length);//向端口写数据
        }
        //接收串口数据
        //01 03 08 00 00 00 00 00 00 00 00 85 17     无按钮按下
        //01 03 08 00 01 00 00 00 00 00 00 85 17     DI0按下   4   2#炉
        //01 03 08 00 00 00 00 00 01 00 00 85 17     DI2按下   8   1#炉
        



        private void Sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
              timer.Close();

              System.Threading.Thread.Sleep(100);

              byte[] ReDatas = new byte[Sp.BytesToRead];
              Sp.Read(ReDatas, 0, ReDatas.Length);

              if (ReDatas[4].ToString().Equals("1") || ReDatas[6].ToString().Equals("1") || ReDatas[8].ToString().Equals("1"))
              {
                 
                  if (flag == 0)
                  {
                      if (ReDatas[6].ToString().Equals("1"))
                      {
                          SqlConnection  conn = new SqlConnection (mysqlcon); ;
                          conn.Open();
                          String sql = "insert into samidtime (timeS,status,date1,station) values ( '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','N','" + DateTime.Now.ToString("yyyyMMdd") + "','3#')";
                          Debug.WriteLine(sql);
                          SqlCommand con = new SqlCommand(sql, conn);
                          con.ExecuteNonQuery();
                          conn.Close();
                          
                      }
                      if (ReDatas[4].ToString().Equals("1"))
                      {
                          SqlConnection  conn = new SqlConnection (mysqlcon); ;
                          conn.Open();
                          String sql = "insert into samidtime (timeS,status,date1,station) values ( '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','N','" + DateTime.Now.ToString("yyyyMMdd") + "','2#')";
                          SqlCommand con = new SqlCommand(sql, conn);
                          con.ExecuteNonQuery();
                          conn.Close();
                         
                      }
                      if (ReDatas[8].ToString().Equals("1"))
                      {
                          SqlConnection  conn = new SqlConnection (mysqlcon); ;
                          conn.Open();
                          String sql = "insert into samidtime (timeS,status,date1,station) values ( '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','N','" + DateTime.Now.ToString("yyyyMMdd") + "','1#')";
                          SqlCommand con = new SqlCommand(sql, conn);
                          con.ExecuteNonQuery();
                          conn.Close();
                          
                      }

                      flag = 1;
                  }
                  }
              else
              {
                  flag = 0;
              }
            timer.Start();
            
        }

        //暂停
        private void button2_Click(object sender, EventArgs e)
        {
            Sp.Close();
            k_hook.Stop();
            timer.Close();
            label4.Text = "未运行";
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = false;
        }
        //查询
        private void button3_Click(object sender, EventArgs e)
        {
            String date1 = this.dateTimePicker1.Text.ToString().Trim(); //日期起
            String date2 = this.dateTimePicker2.Text.ToString().Trim(); //日期
            String sql = "select date1,sampleId,timeS,timeE,timeAll,status,station from samidtime where date1 >= '" + date1 + "' and date1<='" + date2 + "' order by timeS desc";
            
            
            SqlConnection  conn = new SqlConnection (mysqlcon); ;

            //查询 
            SqlDataAdapter sad = new SqlDataAdapter(sql, conn);//创建查询器
            DataSet ds = new DataSet();//创建结果集
            sad.Fill(ds);//将结果集填入
            this.dataGridView1.DataSource = ds.Tables[0];//搜索获取结果集中第一个表，指定
        }

        //刷新
        private void refresh()
        {
            String date1 = this.dateTimePicker1.Text.ToString().Trim(); //日期起
            String date2 = this.dateTimePicker2.Text.ToString().Trim(); //日期
            String sql = "select date1,sampleId,timeS,timeE,timeAll,status,station from samidtime where date1 >= '" + date1 + "' and date1<='" + date2 + "' order by timeS desc";


            SqlConnection  conn = new SqlConnection (mysqlcon); ;

            //查询 
            SqlDataAdapter sad = new SqlDataAdapter(sql, conn);//创建查询器
            DataSet ds = new DataSet();//创建结果集
            sad.Fill(ds);//将结果集填入
            this.dataGridView1.DataSource = ds.Tables[0];//搜索获取结果集中第一个表，指定
        
        
        }
        //修改表yiliangang的status置为Y，（默认值为N）  （更新状态）
        private void button4_Click(object sender, EventArgs e)
        {

            String sql444 = "update  yiliangang set status='Y' where status='N'";
            SqlConnection conn444 = new SqlConnection(mysqlcon); ;
            conn444.Open();
            SqlCommand con444 = new SqlCommand(sql444, conn444);
            con444.ExecuteNonQuery();
            conn444.Close();
            
        }
        //作废
        private void button5_Click(object sender, EventArgs e)
        {
            SqlConnection conn555 = new SqlConnection(mysqlcon);

            String sql = "select min(id) as newid ,timeS as times from samidtime where status='N'  group by times";

            SqlCommand con = new SqlCommand(sql, conn555);
                    //SqlCommand con222 = new SqlCommand(sql222, conn);
            conn555.Open();
            SqlDataReader DataReader = con.ExecuteReader();
                    //SqlDataReader DataReader222 = con222.ExecuteReader();


            if (DataReader.HasRows)
            {

                DataReader.Read();
                int id = DataReader.GetInt32(0);
                ID = id.ToString();
                
                timeS = DataReader.GetString(1);

                
                timeE = DateTime.Now;
                TimeSpan timeAll = timeE.Subtract(Convert.ToDateTime(timeS));
                conn555.Close();


                string bianhao = "作废";
                String sql666 = "update  samidtime set timeE=' " + timeE.ToString("yyyy-MM-dd HH:mm:ss") + "',status='S',sampleId='" + bianhao + "', timeAll='" + Convert.ToDouble(timeAll.TotalMinutes).ToString("0.0") + "' where id ='" + ID + "'";
                
                SqlConnection conn666 = new SqlConnection(mysqlcon);
                SqlCommand con666 = new SqlCommand(sql666, conn666);
                conn666.Open();
                con666.ExecuteNonQuery();
                conn666.Close();
                refresh();
            }
            else
            { 
            MessageBox.Show("没有需要作废的样品");
            }

        }
    }
}

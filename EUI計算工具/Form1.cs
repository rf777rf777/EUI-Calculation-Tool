using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;


namespace EUI計算工具
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds2 = new DataSet(); //關連表用的DataSet
        DataSet ds4 = new DataSet();
        SqlConnection cn = new SqlConnection();
       
        string string1 = "SELECT "
                        +"房間資料表.編號,房間資料表.名稱,ROUND(房間資料表.面積,3) AS 房間面積,"
                        + "convert(Decimal(38,3),ROUND((((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/1000.0),3)) AS 耗電量度KW_H" + ","
                        + "convert(Decimal(38,3),ROUND((((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/1000.0/房間資料表.面積),3)) AS EUI "
                        +"FROM (SELECT DISTINCT 編號,名稱,面積,年使用週數,週使用時數,空調使用週數 FROM 房間) AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號) AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 "
                        +"WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號";

        string string3 = "SELECT A1.部門 AS 房間類型,"
                        +"convert(Decimal(20,3),ROUND(A1.面積,3)) AS 類型面積,"
                        +"convert(Decimal(20,3),ROUND(A1.用電容計算之耗電量,3)) AS 耗電量,"
                        +"convert(Decimal(20,3),ROUND((SUM(A2.用電容計算之耗電量)),3))AS 累計耗電量,"
                        +"convert(Decimal(20,3),ROUND((A1.用電容計算之耗電量/A3.耗電量*100),3)) AS 耗電百分比,"
                        +"convert(Decimal(20,3),ROUND((SUM(A2.用電容計算之耗電量)/A3.耗電量*100),3)) AS 累計百分比,"
                        +"convert(Decimal(20,3),ROUND(A1.EUI,3)) AS EUI "
                        +"FROM (SELECT 房間資料表.部門 AS 部門,SUM(房間資料表.面積) AS 面積,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0 AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000.0 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門)AS A1 ,(SELECT 房間資料表.部門 AS 部門,SUM(房間資料表.面積) AS 面積,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000.0 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門)AS A2,(SELECT SUM(((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/1000.00) AS 耗電量 FROM(SELECT DISTINCT 編號,名稱,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表 ,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing  ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 "
                        +"WHERE 房間資料表.編號=燈具類型資料.編號  AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號)AS A3 WHERE a1.用電容計算之耗電量 <= a2.用電容計算之耗電量 or (a1.用電容計算之耗電量=a2.用電容計算之耗電量 and a1.部門 = a2.部門) GROUP BY  A1.部門,A1.用電容計算之耗電量,(A1.用電容計算之耗電量)/(A3.耗電量)*100,A3.耗電量,A1.EUI,A1.面積 ";
       
        string string7 = "ORDER BY A1.用電容計算之耗電量 DESC";


        string ROOMstr_s = "(SELECT DISTINCT 編號,名稱,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表"; 
        //房間暫存資料表
        string LIGHT_s = "(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料";
        //燈具類型佔存資料表

        string PLUMBING_s = "(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料";
        //衛工裝置類型暫存資料表
        string MACHINE_s = "(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料";
        //機械設備類型暫存資料表
        string DATA_s = "(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料";
        //資料裝置類型暫存資料表
        string FURNITURE_s = "(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料";
        //家具類型暫存資料表
        string WHERE_s = "WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號";
        //合併條件

        private void Form1_Load(object sender, EventArgs e)
        {
       
        }
        private void button1_Click(object sender, EventArgs e)
        {
             string cnStr = "Server=" + textBox1.Text + ";database=" + textBox2.Text+ ";Integrated Security=True;"; //SYASHIN\\SQLEXPRESS 兩條斜線
 
            if (textBox1.Text == "" | textBox2.Text == "")
            {
               MessageBox.Show("請輸入您要連結的資料庫名稱", "目前狀態",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        
            else
            {
                try
                {
                          cn.ConnectionString = cnStr;
                          cn.Open();
                          if (cn.State == ConnectionState.Open)
                          {
                              MessageBox.Show("「" + cn.Database + "」已連接", "目前狀態", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                              button1.Text = "已連接";//cn.Database;
                              //button1.Font = new Font(FontFamily.GenericSansSerif,6F, FontStyle.Bold);  字型設定
                              button1.Enabled = false;
                              button2.Enabled = true;
                              button3.Enabled = true;
                              button4.Enabled = true;
                              button5.Enabled = true;
                              button6.Enabled = true;
                              button7.Enabled = true;
                              
                              textBox1.Enabled = false;
                              textBox2.Enabled = false;
                              textBox3.Enabled = true;

                              button8.Enabled = true;
                              button9.Enabled = true;
                              checkBox1.Enabled = true;
                              checkBox2.Enabled = true;
                              radioButton1.Enabled = true;
                              radioButton2.Enabled = true;

//-----------------------------------------各式資料表查看----------------------------------------------------//

                              SqlDataAdapter da1 = new SqlDataAdapter
                              ("SELECT 樓層.名稱 AS 所在樓層,位置,衛工裝置.名稱,規格容量W "
                              +"FROM 衛工裝置,樓層 "
                              +"WHERE 衛工裝置.樓層=樓層.ID "
                              +"ORDER BY 位置 ASC", cn);                             
                              da1.Fill(ds, "飲水機(衛工裝置)");                            
  
                              SqlDataAdapter da3 = new SqlDataAdapter
                              ("SELECT 樓層.名稱 AS 所在樓層,位置,燈具.名稱,規格,規格容量W "
                              +"FROM 燈具,樓層 "
                              +"WHERE 燈具.樓層=樓層.ID "
                              +"ORDER BY 位置 ASC", cn);                              
                              da3.Fill(ds, "燈具");

                              SqlDataAdapter da4 = new SqlDataAdapter
                              ("SELECT 樓層.名稱 AS 所在樓層,位置,機械設備.名稱,規格容量W "
                              +"FROM 機械設備,樓層 "
                              +"WHERE 機械設備.樓層=樓層.ID "
                              +"ORDER BY 位置 ASC", cn);
                              da4.Fill(ds, "機械設備");

                              SqlDataAdapter da6 = new SqlDataAdapter
                              ("SELECT 樓層.名稱 AS 所在樓層,位置,資料裝置.名稱,規格容量W "
                              +"FROM 資料裝置,樓層 "
                              +"WHERE 資料裝置.樓層=樓層.ID "
                              +"ORDER BY 位置 ASC", cn);
                              da6.Fill(ds, "資料裝置");

                              SqlDataAdapter da7 = new SqlDataAdapter
                              ("SELECT 樓層.名稱 AS 所在樓層,位置,家具.名稱,規格容量W "
                              +"FROM 家具,樓層 "
                              +"WHERE 家具.樓層=樓層.ID "
                              +"ORDER BY 位置 ASC", cn);
                              da7.Fill(ds, "家具");
                                             
                              SqlDataAdapter da8 = new SqlDataAdapter
                              ("SELECT 樓層.名稱 AS 所在樓層,編號,房間.名稱,部門 AS 房間類型,ROUND(面積,3) AS 房間面積,年使用週數,週使用時數,空調使用週數 "
                              +"FROM 房間,樓層 "
                              +"WHERE 房間.樓層=樓層.ID "
                              +"ORDER BY 編號 ASC", cn);
                              da8.Fill(ds, "房間");
                             
                              comboBox1.Text = "請選擇資料表";
                                                                                                 
                              for (int i = 0; i < ds.Tables.Count; i++)                              
                              {                                                                        
                                  comboBox1.Items.Add(ds.Tables[i].TableName);                                                                                                                  
                              }
             
//-----------------------------------------各式資料表查看(END)----------------------------------------------------//     
                                                          
                              SqlDataAdapter da10 = new SqlDataAdapter(string1, cn);
                              da10.Fill(ds1, "房間各別統計");
                              
                          
                             
                              SqlDataAdapter da12 = new SqlDataAdapter
                              ("SELECT "
                              +"convert(Decimal(38,3),ROUND((SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0) )/1000.0),3)) AS 總耗電量度KW_H" + ","
                              +"ROUND(SUM(房間資料表.面積),3) AS 總樓地板面積"+","
                              +"convert(Decimal(38,3),ROUND((SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0) )/1000.0/SUM (房間資料表.面積)),3)) AS 總體EUI "
                              +"FROM" + ROOMstr_s + "," + LIGHT_s + "," + PLUMBING_s + "," + MACHINE_s + "," + DATA_s + "," + FURNITURE_s + " " + WHERE_s, cn);                       
                              da12.Fill(ds1, "總耗電量統計");
                              
                              string string4 = "HAVING SUM(A2.用電容計算之耗電量)/A3.耗電量*100<=100 ORDER BY A1.用電容計算之耗電量 DESC"; //HAVING子句可以消除
                              SqlDataAdapter da13 = new SqlDataAdapter(string3+" "+string4, cn);
                              da13.Fill(ds1, "耗電量與EUI統計");
                           
                              SqlDataAdapter daCATEGORY = new SqlDataAdapter
                              ("SELECT DISTINCT 部門 FROM 房間", cn);
                              daCATEGORY.Fill(ds2, "空間類型");

                              SqlDataAdapter daROOM = new SqlDataAdapter
                              ("SELECT 編號,名稱,部門,年使用週數,週使用時數,空調使用週數 FROM 房間", cn);
                              daROOM.Fill(ds2, "房間列表");

                              SqlDataAdapter daFacility = new SqlDataAdapter
                              ("SELECT 位置,名稱,規格容量W,耗電量 FROM(SELECT 位置,燈具.名稱,規格容量W,規格容量W*年使用週數*週使用時數 AS 耗電量 FROM 燈具,房間 WHERE 燈具.位置=房間.編號 UNION ALL SELECT 位置,資料裝置.名稱,規格容量W,規格容量W*年使用週數*週使用時數 AS 耗電量 FROM 資料裝置,房間 WHERE 資料裝置.位置=房間.編號 UNION ALL SELECT 位置,機械設備.名稱,規格容量W,規格容量W*週使用時數*空調使用週數 AS 耗電量 FROM 機械設備,房間 WHERE 機械設備.位置=房間.編號 UNION ALL SELECT 位置,名稱,規格容量W,規格容量W*年使用時數hr AS 耗電量 FROM 衛工裝置 UNION ALL SELECT 位置,家具.名稱,規格容量W,規格容量W*年使用時數hr AS 耗電量 FROM 家具)AS A1", cn);
                              daFacility.Fill(ds2, "設備列表");
                          
                             
                                  DataRelation datarelation1 = new DataRelation("FK_房間列表_空間類型", ds2.Tables["空間類型"].Columns["部門"], ds2.Tables["房間列表"].Columns["部門"]);
                                  DataRelation datarelation2 = new DataRelation("FK_設備列表_房間列表", ds2.Tables["房間列表"].Columns["編號"], ds2.Tables["設備列表"].Columns["位置"]);

                                  ds2.Relations.AddRange(new DataRelation[] { datarelation1, datarelation2 });

                                  dataGridView4.DataSource = ds2;          
                                  dataGridView4.DataMember = "空間類型";
                                  dataGridView4.Dock = DockStyle.Fill;

                                   
                                                       
                                 
                          }
                     
                  }
                  catch
                  {
                      if (cn.State == ConnectionState.Open)
                      {
                          MessageBox.Show("資料庫沒有對應的資料表，請檢查您的伺服器與資料庫名稱是否輸入正確", "目前狀態", MessageBoxButtons.OK, MessageBoxIcon.Question);
                      }
                      else
                      {
                          MessageBox.Show("沒有這個資料庫，請檢查您的伺服器與資料庫名稱是否輸入正確", "目前狀態", MessageBoxButtons.OK, MessageBoxIcon.Question);
                      }
                  }
             }
           
        }
        private void button2_Click(object sender, EventArgs e)
        {
            cn.Close();
            if (cn.State == ConnectionState.Closed)
            {
                MessageBox.Show("資料庫已關閉", "目前狀態", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                button1.Text = "連接資料庫";
                button1.Enabled = true;
                button2.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                button7.Enabled = false;

                textBox1.Enabled = true;
                textBox2.Enabled = true;
                textBox3.Enabled = false;
                button8.Enabled = false;
                button9.Enabled = false;
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                label4.Visible = false;
                label5.Visible = false;
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                label10.Visible = false;
                label11.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                comboBox1.Text = "";
                comboBox1.Items.Clear();
                textBox5.Text = "";
                textBox7.Text = "";
                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();
                button3.Enabled = false;
                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();

                dataGridView4.DataSource = null;
                dataGridView4.Rows.Clear();          

                dataGridView5.DataSource = null;
                dataGridView5.Rows.Clear();
                

                dataGridView6.DataSource = null;
                dataGridView6.Rows.Clear();

                dataGridView7.DataSource = null;
                dataGridView7.Rows.Clear();
                
                dataGridView8.DataSource = null;
                dataGridView8.Rows.Clear();
                
                ds.Reset();
                ds1.Reset();
                ds2.Reset();
                ds3.Reset();
             
            }  
        }

        private void button3_Click(object sender, EventArgs e)
        {         
            
               
                dataGridView1.DataSource = ds.Tables[comboBox1.Text];           
        }


        private void button5_Click(object sender, EventArgs e)
        {             
                dataGridView2.DataSource = ds1.Tables["房間各別統計"];           
        }
      
       

        private void button6_Click(object sender, EventArgs e)
        { 
           
                if (textBox3.Text == "")
                {
                    MessageBox.Show("請輸入房間代號", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    string searchROOM = " AND 房間資料表.編號 = '" + textBox3.Text + "'";                  
                    SqlDataAdapter searchROOMCON = new SqlDataAdapter(string1 + " " + searchROOM, cn);
                    searchROOMCON.Fill(ds2, "房間查詢結果");
                    dataGridView2.DataSource = ds2.Tables["房間查詢結果"];
                }
                                                  
          
        }

      

     

        private void button7_Click(object sender, EventArgs e)
        {
            
            dataGridView2.DataSource = ds1.Tables["總耗電量統計"];
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                dataGridView3.DataSource = ds1.Tables["耗電量與EUI統計"];
                SqlCommand cmdcount = new SqlCommand("SELECT  COUNT(DISTINCT 部門) FROM 房間", cn);
                SqlCommand cmdEUIMAX = new SqlCommand("SELECT A1.部門 FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing  ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND  房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1 WHERE A1.EUI=(SELECT MAX(A1.EUI) FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1)", cn);
                SqlCommand cmdEUIMIN = new SqlCommand("SELECT A1.部門 FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing  ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND  房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1 WHERE A1.EUI=(SELECT MIN(A1.EUI) FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1)", cn);
                SqlCommand cmdPOWERMAX = new SqlCommand("SELECT A1.部門 FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing  ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND  房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1 WHERE A1.用電容計算之耗電量=(SELECT MAX(A1.用電容計算之耗電量) FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1)", cn);
                SqlCommand cmdPOWERMIN = new SqlCommand("SELECT A1.部門 FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing  ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND  房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1 WHERE A1.用電容計算之耗電量=(SELECT MIN(A1.用電容計算之耗電量) FROM(SELECT 房間資料表.部門 AS 部門,(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0)))/1000.0  AS 用電容計算之耗電量,SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/SUM(房間資料表.面積)/1000 AS EUI FROM(SELECT DISTINCT 編號,名稱,部門,面積,年使用週數,週使用時數,空調使用週數 FROM 房間)AS 房間資料表,(SELECT DISTINCT ROOM.編號,COUNT(LIGHT.名稱) AS 燈具數量,SUM(LIGHT.規格容量W) AS 燈具總電容 FROM 房間 AS ROOM FULL OUTER JOIN 燈具 AS LIGHT ON ROOM.編號=LIGHT.位置 GROUP BY ROOM.編號)AS 燈具類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Plumbing.名稱) AS 衛工裝置數量,SUM(Plumbing.規格容量W*Plumbing.年使用時數hr) AS 衛工裝置總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 衛工裝置 AS Plumbing ON Plumbing.位置=ROOM.編號 GROUP BY ROOM.編號)AS 衛工裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Machine.名稱) AS 機械設備數量,SUM(Machine.規格容量W) AS 機械設備總電容 FROM 房間 AS ROOM FULL OUTER JOIN 機械設備 AS Machine ON Machine.位置=ROOM.編號 GROUP BY ROOM.編號)AS 機械設備類型資料,(SELECT DISTINCT ROOM.編號,COUNT(DATA.名稱) AS 資料裝置數量,SUM(DATA.規格容量W) AS 資料裝置總電容 FROM 房間 AS ROOM FULL OUTER JOIN 資料裝置 AS DATA ON DATA.位置=ROOM.編號 GROUP BY ROOM.編號)AS 資料裝置類型資料,(SELECT DISTINCT ROOM.編號,COUNT(Furniture.名稱) AS 家具數量,SUM(Furniture.規格容量W*Furniture.年使用時數hr) AS 家具總耗電 FROM 房間 AS ROOM FULL OUTER JOIN 家具 AS Furniture ON Furniture.位置=ROOM.編號 GROUP BY ROOM.編號)AS 家具類型資料 WHERE 房間資料表.編號=燈具類型資料.編號 AND 房間資料表.編號=衛工裝置類型資料.編號 AND 房間資料表.編號=機械設備類型資料.編號 AND 房間資料表.編號=資料裝置類型資料.編號 AND 房間資料表.編號=家具類型資料.編號 GROUP BY 房間資料表.部門) AS A1)", cn);

                label4.Visible = true;
                label5.Visible = true;
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                label9.Visible = true;
                label10.Visible = true;
                label11.Visible = true;
                label12.Visible = true;
                label13.Visible = true;

                label13.Text = cmdcount.ExecuteScalar().ToString();
                label9.Text = cmdEUIMAX.ExecuteScalar().ToString();
                label11.Text = cmdEUIMIN.ExecuteScalar().ToString();
                label5.Text = cmdPOWERMAX.ExecuteScalar().ToString();
                label7.Text = cmdPOWERMIN.ExecuteScalar().ToString();
            }
            else if (radioButton2.Checked)
            {
                if (textBox4.Text == "")
                {
                    MessageBox.Show("請輸入房間類型", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    string searchCATEGORY = " HAVING A1.部門= '" + textBox4.Text + "'";
                    SqlDataAdapter searchCATEGORYCON = new SqlDataAdapter(string3 + " " + searchCATEGORY, cn);
                    searchCATEGORYCON.Fill(ds2, "類型查詢結果");
                    dataGridView3.DataSource = ds2.Tables["類型查詢結果"];
                }
            }
            else 
            {
                MessageBox.Show("請選擇全部類型或輸入要查詢的房間類別後按下按鈕", "請選取一篩選方式", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox5.Enabled = true;
            }
            else 
            {
                textBox5.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                textBox7.Enabled = true;
            }
            else
            {
                textBox7.Enabled = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ds3.Reset();
            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();

            if (checkBox1.Checked && checkBox2.Checked && (textBox5.Text != "" || textBox7.Text != ""))
            {
                string string6 = textBox5.Text;
                string string8 = textBox7.Text;
                string string10 = "HAVING SUM(A2.用電容計算之耗電量)/A3.耗電量*100<='"+string6+"' AND A1.EUI>='" + string8 + "'" +string7;
                SqlDataAdapter da16 = new SqlDataAdapter(string3 + " " + string10, cn);
                da16.Fill(ds3, "累計耗電與EUI篩選");
                dataGridView3.DataSource = ds3.Tables["累計耗電與EUI篩選"];
            }

            else if (checkBox1.Checked && textBox5.Text !="")
            { 
                string string6= textBox5.Text;
                string string5="HAVING SUM(A2.用電容計算之耗電量)/A3.耗電量*100<='"+string6+"'"+" "+string7;
                SqlDataAdapter da14 = new SqlDataAdapter(string3+" "+string5, cn);
                da14.Fill(ds3, "累計耗電篩選");                    
                dataGridView3.DataSource = ds3.Tables["累計耗電篩選"];
            }

            else if (checkBox2.Checked && textBox7.Text != "")
            {
                string string8 = textBox7.Text;
                string string9 = "HAVING  A1.EUI>='" + string8 + "'" + " " + string7;
                SqlDataAdapter da15 = new SqlDataAdapter(string3 + " " + string9, cn);
                da15.Fill(ds3, "EUI篩選");
                dataGridView3.DataSource = ds3.Tables["EUI篩選"];
            }
            else
            {
                MessageBox.Show("請至少選取一篩選方式，並輸入條件", "請給定篩選範圍", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
            
        {  
            dataGridView5.DataMember = "空間類型.FK_房間列表_空間類型";
            dataGridView5.DataSource = ds2;
          
            dataGridView5.Dock = DockStyle.Fill;
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        { 
            dataGridView6.DataMember = "空間類型.FK_房間列表_空間類型.FK_設備列表_房間列表";
            dataGridView6.DataSource = ds2;
           
            dataGridView6.Dock = DockStyle.Fill;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            textBox4.Enabled = true;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            textBox4.Enabled = false;
            textBox4.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SqlDataAdapter structure = new SqlDataAdapter("SELECT "
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(資料裝置類型資料.資料裝置總電容,0)*週使用時數*年使用週數)/1000.0),3)) AS 資料裝置耗電量" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(燈具類型資料.燈具總電容,0)*週使用時數*年使用週數)/1000.0),3)) AS 燈具耗電量" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*空調使用週數)/1000.0),3)) AS 機械設備耗電量" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(家具類型資料.家具總耗電,0))/1000.0),3)) AS 家具耗電量" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(衛工裝置類型資料.衛工裝置總耗電,0))/1000.0),3)) AS 衛工裝置耗電量 "
                              +"FROM" + ROOMstr_s + "," + LIGHT_s + "," + PLUMBING_s + "," + MACHINE_s + "," + DATA_s + "," + FURNITURE_s + " " + WHERE_s, cn);                      
            structure.Fill(ds2, "用電分布結構");  
            dataGridView7.DataSource = ds2.Tables["用電分布結構"];

            SqlDataAdapter structurepercent = new SqlDataAdapter("SELECT "
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(資料裝置類型資料.資料裝置總電容,0)*週使用時數*年使用週數)/1000.0/(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0) )/1000.0)*100),3)) AS 資料裝置耗電百分比" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(燈具類型資料.燈具總電容,0)*週使用時數*年使用週數)/1000.0/(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0) )/1000.0)*100),3)) AS 燈具耗電百分比" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*空調使用週數)/1000.0/(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0))/1000.0)*100),3)) AS 機械設備耗電百分比" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(家具類型資料.家具總耗電,0))/1000.0/(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0) )/1000.0)*100),3)) AS 家具耗電百分比" + ","
                              + "convert(Decimal(38,3),ROUND((SUM(ISNULL(衛工裝置類型資料.衛工裝置總耗電,0))/1000.0/(SUM((ISNULL(燈具類型資料.燈具總電容,0)+ISNULL(資料裝置類型資料.資料裝置總電容,0))*週使用時數*年使用週數+ISNULL(機械設備類型資料.機械設備總電容,0)*週使用時數*ISNULL(空調使用週數,0)+ISNULL(衛工裝置類型資料.衛工裝置總耗電,0)+ISNULL(家具類型資料.家具總耗電,0) )/1000.0)*100),3)) AS 衛工裝置耗電百分比 "
                              + "FROM" + ROOMstr_s + "," + LIGHT_s + "," + PLUMBING_s + "," + MACHINE_s + "," + DATA_s + "," + FURNITURE_s + " " + WHERE_s, cn);
            structurepercent.Fill(ds2, "用電結構百分比");
            dataGridView8.DataSource = ds2.Tables["用電結構百分比"];



        }

    }
}

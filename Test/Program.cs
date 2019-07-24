using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using UnityExcel;

namespace Test
{
    class Program
    {
        public class Data
        {
            private string _姓名;
            private int _账号;
            private bool _性别;

            public string 姓名 { get => this._姓名; set => this._姓名 = value; }
            public int 账号 { get => this._账号; set => this._账号 = value; }
            public bool 性别 { get => this._性别; set => this._性别 = value; }
        }
         
        [STAThread]
        static void Main(string[] args)
        {

            List<Data> datas = new List<Data>();
            //datas.Add(new Data()
            //{
            //    姓名 = "adsa",
            //    性别 = true,
            //    账号 = 158943
            //});
            //datas.Add(new Data()
            //{
            //    姓名 = "ddd",
            //    性别 = true,
            //    账号 = 23334
            //});
            //datas.Add(new Data()
            //{
            //    姓名 = "wwww",
            //    性别 = false,
            //    账号 = 1589134215
            //});

            //Excel.SaveDialog<Data>(datas);


            //Excel.Save<Data>( Directory.GetCurrentDirectory() + "/aa.csv", datas);

            //datas = Excel.Read<Data>(typeof(Data), Directory.GetCurrentDirectory() + "/aa.csv");

            datas  =Excel.ReadDialog<Data>(typeof(Data));
            foreach (var item in datas)
            {
                Console.WriteLine(item.姓名 + " | " + item.性别 + " | " + item.账号);
            }
            Console.WriteLine("保存成功");
            Console.Read();
            //UnityExcel.Excel.Read<>();
        }
    }
}

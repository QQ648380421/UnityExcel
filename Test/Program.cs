using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UnityExcel;
using Newtonsoft.Json;
namespace Test
{
    class Program
    {

        public class Data
        {
     
            private string column1;
            private int column2;
            private float column3;
            private bool column4;
            private long column5;
            public string Column1 { get => column1; set => column1 = value; }
            [ExcelHead("int类型")]
            public int Column2 { get => column2; set => column2 = value; }
            [ExcelHead("float类型")]
            public float Column3 { get => column3; set => column3 = value; }
            [ExcelHead("bool类型")]
            public bool Column4 { get => column4; set => column4 = value; }
            [ExcelHead("long类型")]
            public long Column5 { get => column5; set => column5 = value; }
            public override string ToString()
            {
                string str = JsonConvert.SerializeObject(this);
                return str;
            }
        }
        [STAThread]
        static void Main(string[] args)
        {
            SaveDatas();

            ReadDatas();
            Console.Read();
        }

        private static void ReadDatas()
        {
            //读取数据
            List<Data> datas = Excel.ReadDialog<Data>(typeof(Data)); 
            foreach (var item in datas)
            {
               
                Console.WriteLine(item);
            }

        }

        private static void SaveDatas()
        {
            //保存数据
            List<Data> datas = new List<Data>();
            for (int i = 0; i < 30; i++)
            {
                datas.Add(new Data()
                {
                    Column1 = "string类型：" + i,
                    Column2 = i,
                    Column3 = i * 0.01f,
                    Column4 = i / 2 == 0,
                    Column5 = i * 1000000

                });
            }
            Excel.SaveDialog(datas);
        }
    }
}

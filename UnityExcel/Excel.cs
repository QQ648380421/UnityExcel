using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace UnityExcel
{
    public class Excel
    {
        /// <summary>
        /// 保存编码
        /// </summary>
        public static Encoding Encoding= Encoding.UTF8;

        /// <summary>
        /// 打开文件窗口
        /// </summary>
        /// <returns></returns>
        private static string OpenPath()
        { 
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.RestoreDirectory = true;
            fileDialog.Filter = "表格|*.csv;";
            if (fileDialog.ShowDialog() != DialogResult.OK) return null;
            if (!File.Exists(fileDialog.FileName)) return null;
            return fileDialog.FileName;
        }

        /// <summary>
        /// 保存数据，带文件选择器
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="datas"></param>
        public static void SaveDialog<T>(List<T> datas)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.RestoreDirectory = true;
            fileDialog.Filter = "表格|*.csv;";
            if (fileDialog.ShowDialog() != DialogResult.OK) return;
            Save(fileDialog.FileName, datas);
            MessageBox.Show("保存成功！", "温馨提示"); 
        }

        public static void Save<T>(string Path,List<T> datas) {
            List<object> Datas = new List<object>();
            foreach (var item in datas)
            {
                Datas.Add(item);
            }
            string SaveData=""; 
            for (int i = 0; i < Datas.Count; i++)
            {
                var obj = Datas[i];
                var ObjType = obj.GetType();
                var pros = ObjType.GetProperties();
                if (pros.Length==0)
                {
                    throw new Exception("没有找到任何一个get和set属性，请将字段封装为属性在进行保存！");
                }
                if (i==0)
                {
                    AddHead(ref SaveData, pros);
                }
                AddRow(ref SaveData, pros, obj);
            }
           var bytes = Encoding.GetBytes(SaveData);
            File.WriteAllBytes(Path, bytes); 
        }

        /// <summary>
        /// 添加第一行表头
        /// </summary>
        /// <param name="saveData"></param>
        /// <param name="pros"></param>
        private static void AddHead(ref string saveData, PropertyInfo[] pros)
        {
            for (int i = 0; i < pros.Length; i++)
            {
                if (i!=0)
                {
                    saveData += ",";
                }
                saveData += pros[i].Name;
            }
            saveData += Environment.NewLine;
        }
        /// <summary>
        /// 添加一行数据
        /// </summary>
        /// <param name="saveData"></param>
        /// <param name="pros"></param>
        /// <param name="obj"></param>
        private static void AddRow(ref string saveData, PropertyInfo[] pros,object obj)
        {
            for (int i = 0; i < pros.Length; i++)
            {
                if (i != 0)
                {
                    saveData += ",";
                }
                saveData += pros[i].GetValue(obj,null);
            }
            saveData += Environment.NewLine;
        }

        private class Column
        {
          public  int Id;
            public string Name;
        }

        /// <summary>
        /// 从excel文件中读取，并转化为类型
        /// </summary>
        /// <param name="classType">转化的类型</param>
        /// <param name="filePath">表格文件</param>
        /// <returns></returns>
        public static List<T> Read<T>(Type classType,string filePath) {
           var bytes = File.ReadAllBytes(filePath); 
            string content = Encoding.GetString(bytes);
            string[] rows = Regex.Split(content, Environment.NewLine, RegexOptions.IgnoreCase);   List<T> datas = new List<T>();
            List<Column> columns = GetHeadColumns(rows);

            for (int i = 0; i < rows.Length; i++)
            {
                var row = rows[i];
                if (string.IsNullOrEmpty(row)) continue;
                string[] cells = GetCells(row);
                if (i==0)
                {//表头
                    if (!IsClassData(row, classType)) return null;  
                }
                else
                {//数据
                    datas.Add((T)GetRowData(classType, row, columns)); 
                }
            }  
            return datas; 
        }

        /// <summary>
        /// 弹出选择文件窗口，并读取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="classType"></param>
        /// <returns></returns>
        public static List<T> ReadDialog<T>(Type classType) { 
            return Read<T>(classType,OpenPath());
        }


        /// <summary>
        /// 获取表头列表
        /// </summary>
        private static List<Column> GetHeadColumns(string[] rows)
        {
            var columns = GetCells(rows[0]);
            List<Column> columnDatas = new List<Column>();
            for (int i = 0; i < columns.Length; i++)
            {
                columnDatas.Add(new Column() { 
                    Id = i,
                    Name = columns[i]
                }); 
            }
            return columnDatas;
        }

        private static string[] GetCells(string row)
        {
            return row.Split(',');
        }

        /// <summary>
        /// 获取每一行的数据
        /// </summary>
        /// <param name="classType"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static object GetRowData(Type classType, string row, List<Column> columns)
        {
            var cells = GetCells(row);
            var newobj = classType.Assembly.CreateInstance(classType.FullName);
            var objType = newobj.GetType();
           var pros = objType.GetProperties();
            for (int i = 0; i < columns.Count; i++)
            {
                var column = columns[i];
                var pro = pros.FirstOrDefault(p => p.Name == column.Name);
                pro.SetValue(newobj, Convert.ChangeType(cells[column.Id], pro.PropertyType),null);
            }
            return newobj;
        }

        /// <summary>
        /// 验证表头数据
        /// </summary>
        /// <param name="row"></param>
        /// <param name="classType"></param>
        /// <returns></returns>
        private static bool IsClassData(string row, Type classType)
        {
            var columns = row.Split(',');
            var pros = classType.GetProperties();
            foreach (var item in columns)
            {
               var column = pros.FirstOrDefault(p=>p.Name== item);
                if (column == null) return false; 
            }
            return true;
        }
    }
}

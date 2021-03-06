﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions; 
namespace UnityExcel
{
    /// <summary>
    /// 如果编译后无法打开，请导入以下DLL
    /// I18N.CJK.dll    I18N.dll
    /// </summary>
    public class Excel
    {
        /// <summary>
        /// 保存编码
        /// </summary>
        public static Encoding Encoding= Encoding.GetEncoding("GB2312");
         
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
            //List<byte> bytes = new List<byte>();
            //bytes.AddRange(new byte[] { (byte)0xEF, (byte)0xBB, (byte)0xBF });
            //bytes.AddRange(Encoding.GetBytes(SaveData));
            //File.WriteAllBytes(Path, bytes.ToArray());

             
            File.WriteAllBytes(Path, Encoding.GetBytes(SaveData));
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
                var item = pros[i]; 
              var attribute = GetCustonAttribute(item); 
                if (attribute != null)
                {
                    saveData += attribute.Name;
                }
                else
                {
                    saveData += pros[i].Name;
                }
              
            }
            saveData += Environment.NewLine;
        }
        /// <summary>
        /// 获取自定义属性
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private static ExcelHead GetCustonAttribute(PropertyInfo item)
        {
            var attribute = item.GetCustomAttributes(typeof(ExcelHead), true);
            if (attribute!=null && attribute.Length>0)
            {
                return (ExcelHead)attribute[0];
            }
            return null;
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
            List<T> datas = new List<T>();
            if (string.IsNullOrEmpty(filePath)) return datas;
            var bytes = File.ReadAllBytes(filePath);
            string content = Encoding.GetString(bytes);
            string[] rows = Regex.Split(content, Environment.NewLine, RegexOptions.IgnoreCase);
       
            List<Column> columns = GetHeadColumns(rows);

            for (int i = 0; i < rows.Length; i++)
            {
                var row = rows[i];
                if (string.IsNullOrEmpty(row)) continue;
                string[] cells = GetCells(row);
                if (i==0)
                {//表头
                    if (!IsClassData(row, classType))
                        throw new Exception("表头不匹配，读取失败！");
                }
                else
                {//数据
                    datas.Add((T)GetRowData(classType, row, columns)); 
                }
            }  
            return datas; 
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
                PropertyInfo proInfo=null;
                foreach (var item in pros)
                {
                    var head = GetCustonAttribute(item);
                    if (head==null)
                    {
                        if (item.Name== column.Name)
                        {
                            proInfo = item;
                            break;
                        }
                    }
                    else
                    {
                        if (head.Name == column.Name)
                        {
                            proInfo = item;
                            break;
                        }
                    }
                }
                if (proInfo!=null)
                {
                    proInfo.SetValue(newobj, Convert.ChangeType(cells[column.Id], proInfo.PropertyType), null);
                }
          
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
            if (columns.Length != pros.Length) return false;

            for (int i = 0; i < columns.Length; i++)
            {
                bool flag = false;
                var columns_item = columns[i];
                var pros_item = pros[i]; 
                var head = GetCustonAttribute(pros_item);
                if (head == null)
                {
                    if (pros_item.Name == columns_item)
                    {
                        flag = true; 
                    }
                }
                else
                {
                    if (head.Name == columns_item)
                    {
                        flag = true; 
                    }
                }
                if (flag == false)
                {
                    return false;
                }
            }
             
            return true;
        }
    }
}

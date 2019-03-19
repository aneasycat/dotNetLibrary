using System;
using Microsoft.Office.Interop.Excel;

namespace AnEasyCat.Office
{
    public class Excel
    {
        Application excel;
        Workbooks wbs;
        Workbook wb;
        string _File = "";

        public string FilePath { get; set; } = @"C:\AnEasyCat";
        /// <summary>
        /// 文档名称
        /// </summary>
        public string Name { get; set; }
        public Sheets Sheets => new Sheets { wb=wb};
        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <param name="file">文件路径</param>
        public void Open(string file)
        {
            _File = file;
            excel = new Application();
            wbs = excel.Workbooks;
            wb = wbs.Add(file);

            int last = _File.LastIndexOf(@"\");
            Name = file.Substring(last+1);
            FilePath = file.Substring(0, last);
        }
        /// <summary>
        /// 新建Excel文件
        /// </summary>
        /// <param name="name">文档名称</param>
        public void Create(string name)
        {
            excel = new Application();
            wbs = excel.Workbooks;
            wb = wbs.Add();

            Name = name+".xls";
            if (System.IO.Directory.Exists(FilePath) == false)
                System.IO.Directory.CreateDirectory(FilePath);
            _File = FilePath +@"\"+ Name;
        }
        /// <summary>
        /// 保存Excel文件
        /// </summary>
        public void Save()
        {
            excel.AlertBeforeOverwriting = false;
            excel.DisplayAlerts = false;
            wb.Author = "AnEasyCat";
            wb.SaveAs(_File,Type.Missing,Type.Missing,Type.Missing,Type.Missing, Type.Missing,XlSaveAsAccessMode.xlNoChange,Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }
        /// <summary>
        /// 关闭工作
        /// </summary>
        public void Close()
        {
            wb.Close();
            wbs.Close();
            excel.Quit();
        }
    }
    #region 类
    public class Sheets
    {
        public Workbook wb;
        public Sheet this[int index] => new Sheet { _ws = wb.Sheets.Item[index] };

        public class Sheet
        {
            public Worksheet _ws;
            public string Name
            {
                get
                {
                    return _ws.Name;
                }
                set
                {
                    _ws.Name = value;
                }
            }
            public Rows Rows
            {
                get
                {
                    return new Rows { _rows = _ws.Rows };
                }
            }
        }
        public class Rows
        {
            public Range _rows;
            public int Count
            {
                get { return _rows.Row; }
            }
            public Row this[int index]
            {
                get
                {
                    return new Row { _range=_rows.Rows[index] };
                }
            }
        }
        public class Row:Cell
        {
            public double Hight
            {
                get { return _range.RowHeight; }
                set { _range.RowHeight = value; }
            }
            public Cells Cells
            {
                get
                {
                    return new Cells { _cells = _range.Cells };
                }
            }
            /// <summary>
            /// 在有数据的单元格后追加值
            /// </summary>
            /// <param name="value"></param>
            public void Append(dynamic value)
            {
                bool c = true;
                
                for (int i = 1; c; i++)
                {
                    Range r = _range.Cells[i];
                    if (r.Value2 == null)
                    {
                        r.Value2 = value;
                        c = false;
                    }
                }
                
            }
        }
        public class Cells
        {
            //public Worksheet _cells;
            public Range _cells;
            public Cell this[int index]
            {
                get
                {

                    return new Cell { _range=_cells[index]};
                }
            }
        }
        public class Cell
        {
            public Range _range;
            public dynamic Value
            {
                get
                {
                    return Convert.ToString(_range.Value2);
                }
                set
                {
                    _range.Value2 = value;
                }
            }
            public bool FontBold
            {
                get
                {
                    return _range.Font.Bold;
                }
                set
                {
                    _range.Font.Bold = value;
                }
            }
            public int FontSize
            {
                get
                {
                    return _range.Font.Size;
                }
                set
                {
                    _range.Font.Size = value;
                }
            }
            public System.Drawing.Color FontColor
            {
                get
                {
                    return System.Drawing.Color.FromArgb((int)_range.Font.Color);
                }
                set
                {
                    _range.Font.Color = value.ToArgb();
                }
            }
            public System.Drawing.Color Color
            {
                get
                {
                    return System.Drawing.Color.FromArgb((int)_range.Interior.Color);
                }
                set
                {
                    _range.Interior.Color = value.ToArgb();
                }
            }
        }
    }
    #endregion
}

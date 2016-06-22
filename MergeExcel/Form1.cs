using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;


namespace MergeExcel
{
    public partial class Form1 : Form
    {
        private string inputFoldPath = null;
        private string outputFileName = null;

        private IWorkbook workbook = null;
        private FileStream fs = null;
        private List<string> allFiles = null;
        private List<DataTable> allTable = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                inputFoldPath = dialog.SelectedPath;
                textBox1.Text = inputFoldPath;
                sheetHeaderInit();
                allFiles = excelFileCollector();
                readFiles();
            }
        }
        private void readFiles() {
            foreach (string fileName in allFiles)
            {
                listBox1.Items.Add(fileName);
            }
        }
        /// <summary>
        /// 对Excel表格表头进行标准初始化
        /// </summary>
        private void sheetHeaderInit() {
            DataTable sheetTable1 = new DataTable();
            DataTable sheetTable2 = new DataTable();
            DataTable sheetTable3 = new DataTable();
            DataTable sheetTable4 = new DataTable();
            DataTable sheetTable5 = new DataTable();
            DataTable sheetTable6 = new DataTable();
            DataTable sheetTable7 = new DataTable();
            DataTable sheetTable8 = new DataTable();
            DataTable sheetTable9 = new DataTable();
            allTable = new List<DataTable>();
            #region
            //表1表头初始化
            for (int i = 0; i < 14; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable1.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "区域划分"; break;
                    case 2: temp = "支行名称"; break;
                    case 3: temp = "机构名称（机构名称全称）"; break;
                    case 4: temp = "机构类型（按规范术语要求填写）"; break;
                    case 5: temp = "机构风险防护级别"; break;
                    case 6: temp = "详细地址"; break;
                    case 7: temp = "机构安防负责人"; break;
                    case 8: temp = "员工号"; break;
                    case 9: temp = "负责人固定电话"; break;
                    case 10: temp = "负责人移动电话"; break;
                    case 11: temp = "负责人办公邮箱"; break;
                    case 12: temp = "网点CAD点位图（标有设备编号）"; break;
                    case 13: temp = "备注"; break;
                }
                sheetTable1.Columns[i].ColumnName = temp;
            }
            //表2表头初始化
            for (int i = 0; i < 6; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable2.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "安防业务使用的线路"; break;
                    case 3: temp = "安防业务最高可用带宽"; break;
                    case 4: temp = "安防业务最低可用带宽"; break;
                    case 5: temp = "备注"; break;
                }
                sheetTable2.Columns[i].ColumnName = temp;
            }

            //表3表头初始化
            for (int i = 0; i < 22; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable3.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "机构内序号"; break;
                    case 3: temp = "设备类型"; break;
                    case 4: temp = "品牌"; break;
                    case 5: temp = "型号"; break;
                    case 6: temp = "运行状况"; break;
                    case 7: temp = "软件版本"; break;
                    case 8: temp = "视频通道数(已用/共计）"; break;
                    case 9: temp = "编码协议(H.264/MPEG4)"; break;
                    case 10: temp = "编码格式(CIF/4CIF/D1)"; break;
                    case 11: temp = "单路码率上限(对应编码格式)"; break;
                    case 12: temp = "报警输入通道数(已用/共计）"; break;
                    case 13: temp = "报警输出通道数(已用/共计）"; break;
                    case 14: temp = "硬盘参数（数量/总容量/接口）"; break;
                    case 15: temp = "IP地址/掩码/网关"; break;
                    case 16: temp = "MAC地址"; break;
                    case 17: temp = "远程访问端口号"; break;
                    case 18: temp = "超级用户/密码"; break;
                    case 19: temp = "厂家联系方式(姓名/电话/邮箱）"; break;
                    case 20: temp = "维护商联系方式(姓名/电话/邮箱）"; break;
                    case 21: temp = "备注"; break;
                }
                sheetTable3.Columns[i].ColumnName = temp;
            }
            //表4表头初始化
            for (int i = 0; i < 19; i++)//for test
            {
                DataColumn dc = new DataColumn();
                sheetTable4.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "机构内序号"; break;
                    case 3: temp = "品牌"; break;
                    case 4: temp = "型号"; break;
                    case 5: temp = "IP网络模块品牌"; break;
                    case 6: temp = "IP网络模块型号"; break;
                    case 7: temp = "运行状况"; break;
                    case 8: temp = "软件版本"; break;
                    case 9: temp = "报警输入通道数(已用/共计）"; break;
                    case 10: temp = "报警输出通道数(已用/共计）"; break;
                    case 11: temp = "IP地址/掩码/网关"; break;
                    case 12: temp = "MAC地址"; break;
                    case 13: temp = "远程访问端口号"; break;
                    case 14: temp = "超级用户/密码"; break;
                    case 15: temp = "厂家联系方式(姓名/电话/邮箱）"; break;
                    case 16: temp = "维护商联系方式(姓名/电话/邮箱）"; break;
                    case 17: temp = "备注"; break;
                    case 18: temp = "fault-tolerance"; break;
                }
                sheetTable4.Columns[i].ColumnName = temp;
            }
            //表5表头初始化
            for (int i = 0; i < 18; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable5.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "硬盘录像机名称"; break;
                    case 3: temp = "硬盘录像机IP地址（与表3中IP地址一致）"; break;
                    case 4: temp = "视频通道号"; break;
                    case 5: temp = "摄像机在CAD点位图上的设备编号"; break;
                    case 6: temp = "录像方式"; break;
                    case 7: temp = "安装部位（按规范术语要求填写）"; break;
                    case 8: temp = "设备类型（枪机/球机/…）"; break;
                    case 9: temp = "特殊功能（普通/宽动态/红外/…）"; break;
                    case 10: temp = "品牌"; break;
                    case 11: temp = "型号"; break;
                    case 12: temp = "云台支持"; break;
                    case 13: temp = "云台协议"; break;
                    case 14: temp = "音频编码协议（无/G.711/G.729/…）"; break;
                    case 15: temp = "摄像机厂家联系方式(姓名/电话/邮箱）"; break;
                    case 16: temp = "摄像机维护商联系方式(姓名/电话/邮箱）"; break;
                    case 17: temp = "备注"; break;
                    //case 17: temp = ""; break;
                }
                sheetTable5.Columns[i].ColumnName = temp;
            }
            //表6表头初始化
            for (int i = 0; i < 20; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable6.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "硬盘录像机名称"; break;
                    case 3: temp = "硬盘录像机IP地址（与表3中IP地址一致）"; break;
                    case 4: temp = "视频通道号"; break;
                    case 5: temp = "摄像机在CAD点位图上的设备编号"; break;
                    case 6: temp = "安装部位（按规范术语要求填写）"; break;
                    case 7: temp = "报警探测器（按规范术语要求填写）"; break;
                    case 8: temp = "报警类型"; break;
                    case 9: temp = "报警探测器品牌"; break;
                    case 10: temp = "报警探测器型号"; break;
                    case 11: temp = "是否有防拆报警"; break;
                    case 12: temp = "防拆防区号"; break;
                    case 13: temp = "报警方式"; break;
                    case 14: temp = "布防时间段"; break;
                    case 15: temp = "需要联动的硬盘录像机IP地址"; break;
                    case 16: temp = "需要联动的视频通道号"; break;
                    case 17: temp = "报警探测器厂家联系方式(姓名/电话/邮箱）"; break;
                    case 18: temp = "报警探测器维护商联系方式(姓名/电话/邮箱）"; break;
                    case 19: temp = "备注"; break;
                }
                sheetTable6.Columns[i].ColumnName = temp;
            }
            //表7表头初始化
            for (int i = 0; i < 20; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable7.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "报警主机名称"; break;
                    case 3: temp = "报警主机IP地址（与表4中IP地址一致）"; break;
                    case 4: temp = "报警通道号"; break;
                    case 5: temp = "报警探测器在CAD点位图上的设备编号"; break;
                    case 6: temp = "安装部位（按规范术语要求填写）"; break;
                    case 7: temp = "报警探测器（按规范术语要求填写）"; break;
                    case 8: temp = "报警类型"; break;
                    case 9: temp = "报警探测器品牌"; break;
                    case 10: temp = "报警探测器型号"; break;
                    case 11: temp = "是否有防拆报警"; break;
                    case 12: temp = "防拆防区号"; break;
                    case 13: temp = "报警方式（常开/常闭）"; break;
                    case 14: temp = "布防时间段"; break;
                    case 15: temp = "需要联动的硬盘录像机IP地址"; break;
                    case 16: temp = "需要联动的视频通道号"; break;
                    case 17: temp = "报警探测器厂家联系方式(姓名/电话/邮箱）"; break;
                    case 18: temp = "报警探测器维护商联系方式(姓名/电话/邮箱）"; break;
                    case 19: temp = "备注"; break;
                }
                sheetTable7.Columns[i].ColumnName = temp;
            }

            //表8表头初始化
            for (int i = 0; i < 16; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable8.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "有无出入口控制系统"; break;
                    case 3: temp = "门禁设备品牌"; break;
                    case 4: temp = "门禁设备型号"; break;
                    case 5: temp = "门禁在CAD点位图上的设备编号"; break;
                    case 6: temp = "安装部位"; break;
                    case 7: temp = "有无门禁控制器"; break;
                    case 8: temp = "通讯接口"; break;
                    case 9: temp = "IP地址/掩码/网关"; break;
                    case 10: temp = "用户名/密码"; break;
                    case 11: temp = "需要联动的硬盘录像机IP地址"; break;
                    case 12: temp = "需要联动的视频通道号"; break;
                    case 13: temp = "厂家联系方式(姓名/电话/邮箱）"; break;
                    case 14: temp = "维护商联系方式(姓名/电话/邮箱）"; break;
                    case 15: temp = "备注"; break;
                }
                sheetTable8.Columns[i].ColumnName = temp;
            }
            //表9表头初始化
            for (int i = 0; i < 14; i++)
            {
                DataColumn dc = new DataColumn();
                sheetTable9.Columns.Add(dc);
                string temp = null;
                switch (i)
                {
                    case 0: temp = "序号"; break;
                    case 1: temp = "机构名称（与表1中机构名称必须完全一致）"; break;
                    case 2: temp = "有无对讲系统"; break;
                    case 3: temp = "对讲设备品牌"; break;
                    case 4: temp = "对讲设备型号"; break;
                    case 5: temp = "对讲设备在CAD点位图上的设备编号"; break;
                    case 6: temp = "安装部位"; break;
                    case 7: temp = "IP地址/掩码/网关"; break;
                    case 8: temp = "用户名/密码"; break;
                    case 9: temp = "需要联动的硬盘录像机IP地址"; break;
                    case 10: temp = "需要联动的视频通道号"; break;
                    case 11: temp = "厂家联系方式(姓名/电话/邮箱）"; break;
                    case 12: temp = "维护商联系方式(姓名/电话/邮箱）"; break;
                    case 13: temp = "备注"; break;
                }
                sheetTable9.Columns[i].ColumnName = temp;
            }
            #endregion
            allTable.Add(sheetTable1);
            allTable.Add(sheetTable2);
            allTable.Add(sheetTable3);
            allTable.Add(sheetTable4);
            allTable.Add(sheetTable5);
            allTable.Add(sheetTable6);
            allTable.Add(sheetTable7);
            allTable.Add(sheetTable8);
            allTable.Add(sheetTable9);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            saveExcel();

        }
        private void saveExcel() {
            foreach (string fileName in allFiles)
            {
                outputFileName = outputFileName + fileName + "_";
                DataTable tempTable;
                for (int i = 0; i < 9; i++)
                {
                    tempTable = ExcelToDataTable(inputFoldPath + "\\" + fileName, i, true);
                    foreach (DataRow dr in tempTable.Rows)
                    {
                        if (dr.ItemArray[1].ToString() != "")
                        {
                            allTable[i].Rows.Add(dr.ItemArray);
                        }
                    }
                }
            }

            DataTableToExcel(inputFoldPath + "\\" + outputFileName + "1.xls", allTable, true);
            MessageBox.Show("Done!");
        }
        /// <summary>
        /// 读取路径下文件的名字
        /// </summary>
        /// <returns>返回名字列表供后续使用</returns>
        private List<string> excelFileCollector()
        {
            List<string> collector = new List<string>();
            DirectoryInfo folder = new DirectoryInfo(inputFoldPath);
            foreach (FileInfo nextFile in folder.GetFiles())
            {
                collector.Add(nextFile.Name);
            }
            return collector;
        }
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string fileName, int sheetPosition, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                sheet = workbook.GetSheetAt(sheetPosition);
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    //int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数
                    int cellCount = 0;
                    for (int i = 0; i < firstRow.LastCellNum;i++ ) {
                        if (firstRow.GetCell(i).CellType!=CellType.Blank)
                        {
                            cellCount++;
                        }
                     }
                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    int MergedCount = sheet.NumMergedRegions;
                    for (int i = MergedCount - 1; i >= 0; i--)
                    {
                        //sheet.RemoveMergedRegion(i);
                        var cellrange = sheet.GetMergedRegion(i);
                        for (int row = cellrange.FirstRow; row <= cellrange.LastRow; row++)
                        {
                            sheet.GetRow(row).GetCell(cellrange.FirstColumn).SetCellType(CellType.String);
                            sheet.GetRow(row).GetCell(cellrange.FirstColumn).SetCellValue(sheet.GetRow(cellrange.FirstRow).GetCell(cellrange.FirstColumn).ToString());
                        }
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {

                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　
                        if (row.ZeroHeight) continue;//删除隐藏行

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            {
                                row.GetCell(j).SetCellType(CellType.String);
                                dataRow[j] = row.GetCell(j).StringCellValue;
                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                fs.Close();
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                DialogResult re = MessageBox.Show("如果你打开了将要合并的Excel表格，请关闭！ 如果都是关闭的，你还看到了这个对话框，这个表格就太不规范了，最好发给我！", "警告！");
                if (re == DialogResult.OK)
                {
                    System.Environment.Exit(0);
                }
                return null;
            }
        }
        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(string outPutName, List<DataTable> dataList, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            string sheetName=null;
            string all = null;

            fs = new FileStream(outPutName, FileMode.Append, FileAccess.Write);
            if (outPutName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (outPutName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();
            for (int m = 0; m < dataList.Count;m++ )
            {
                //for (int k = 0; k < dataList[m].Rows.Count; k++) {
                //    for (int p = 1; p < dataList[m].Columns.Count; p++) {
                //        if (dataList[m].Rows[k][p].ToString() == "")
                //        {
                //            Console.WriteLine();
                //        }
                //        all += dataList[m].Rows[k][p].ToString();
                //        if (all=="") {
                //            dataList[m].Rows.Remove(dataList[m].Rows[k]);
                //        }
                //    }
                //}
                    switch (m)
                    {
                        case 0: sheetName = "表1机构信息调查表"; break;
                        case 1: sheetName = "表2机构网络情况调查表"; break;
                        case 2: sheetName = "表3机构硬盘录像机设备调查表"; break;
                        case 3: sheetName = "表4机构报警主机设备调查表"; break;
                        case 4: sheetName = "表5摄像机与硬盘录像机通道对应关系统计表"; break;
                        case 5: sheetName = "表6报警器、硬盘录像机、联动摄像机对应关系调查表"; break;
                        case 6: sheetName = "表7报警器、报警主机、联动摄像机对应关系调查表"; break;
                        case 7: sheetName = "表8机构门禁设备调查表"; break;
                        case 8: sheetName = "表9机构对讲设备调查表"; break;
                    }
                DataTable data = dataList[m];
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);

                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
            }
            try
            {
                

                workbook.Write(fs); //写入到excel
                fs.Close();
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            inputFoldPath = null;
            outputFileName = null;
            workbook = null;
            fs = null;
            allFiles = null;
            allTable = null;
            listBox1.Items.Clear();
        }
    }
}

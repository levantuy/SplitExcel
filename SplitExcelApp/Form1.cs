using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using DevExpress.UserSkins;
using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.Spreadsheet;

namespace SplitExcelApp
{
    public partial class Form1 : RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
            InitSkinGallery();

        }

        void InitSkinGallery()
        {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }

        private void btnSplitExcel_ItemClick(object sender, ItemClickEventArgs e)
        {
            var workbook = spreadsheetControl.Document;
            var worksheet = spreadsheetControl.Document.Worksheets.ActiveWorksheet;
            Range usedRange = worksheet.GetUsedRange();
            var dataTable = new DataTable("DataInput");
            dataTable.Columns.Add("Name");
            dataTable.Columns.Add("HPQ2");
            dataTable.Columns.Add("PaymentDate");
            dataTable.Columns.Add("PaymentType");
            dataTable.Columns.Add("HP");
            dataTable.Columns.Add("Class");
            dataTable.Columns.Add("Day");
            dataTable.Columns.Add("Note");
            for (int i = 2; i <= usedRange.BottomRowIndex; i++)
            {
                if (!string.IsNullOrEmpty(worksheet.Cells["F" + i.ToString()].Value.ToString()))
                {
                    // add table
                    DataRow dataRow = dataTable.NewRow();
                    dataRow["Name"] = worksheet.Cells["F" + i.ToString()].Value;
                    dataRow["HPQ2"] = worksheet.Cells["K" + i.ToString()].Value;
                    dataRow["PaymentDate"] = worksheet.Cells["N" + i.ToString()].Value;
                    dataRow["PaymentType"] = worksheet.Cells["O" + i.ToString()].Value;
                    dataRow["HP"] = worksheet.Cells["M" + i.ToString()].Value;
                    dataRow["Class"] = worksheet.Cells["G" + i.ToString()].Value;
                    dataRow["Day"] = worksheet.Cells["H" + i.ToString()].Value;
                    dataRow["Note"] = worksheet.Cells["C" + i.ToString()].Value;
                    dataTable.Rows.Add(dataRow);
                }
            }
            spreadsheetControl.Document.LoadDocument(@"templates\temp.xlsx");                      
            IWorkbook workbookNew = spreadsheetControl.Document;
            var dataTableSort = (from DataRow dRow in dataTable.Rows select dRow).OrderBy(x => SplitFullName(x["Name"].ToString(), true)).ThenBy(x => SplitFullName(x["Name"].ToString(), false));

            var distinctRows = (from DataRow dRow in dataTable.Rows
                                select new { ClassName = dRow["Class"].ToString(), ClassDay = dRow["Day"].ToString() }).Distinct().OrderBy(x => x.ClassName);
            var sheetErrors = workbookNew.Worksheets.Add("Errors");
            workbook.Worksheets[sheetErrors.Name].CopyFrom(workbook.Worksheets["tempSheet"]);
            var errorCount = 5;
            foreach (var info in distinctRows)
            {
                var sheet = workbookNew.Worksheets.Add();                
                sheet.Name = info.ClassName + " " + ConvertToEnglish(info.ClassDay);
                workbook.Worksheets[sheet.Name].CopyFrom(workbook.Worksheets["tempSheet"]);
                workbook.Worksheets[sheet.Name].Cells["C2"].Value = "LỚP: " + info.ClassName;
                var count = 5;                
                foreach (DataRow row in dataTableSort)
                {
                    if (row["Class"].ToString() == info.ClassName && row["Day"].ToString() == info.ClassDay)
                    {
                        try
                        {
                            count++;
                            sheet.Rows.Insert(count);
                            sheet.Rows[count].CopyFrom(sheet.Rows[count + 1], PasteSpecial.All);
                            sheet.Rows[count][1].Value = (count - 5).ToString();
                            sheet.Rows[count][2].Value = row[0].ToString();
                            sheet.Rows[count][3].Value = row[1].ToString().Length > 0 ? Convert.ToDouble(row[1]) : 0;
                            sheet.Rows[count][4].Value = row[2].ToString().Length > 6 ? Convert.ToDateTime(row[2]).ToString("dd/MM/yyyy") : row[2].ToString();
                            sheet.Rows[count][5].Value = row[3].ToString();
                            sheet.Rows[count][6].Value = row[4].ToString().Length > 0 ? Convert.ToDouble(row[4]) : 0;
                            sheet.Cells["Q" + (count + 1).ToString()].Value = row[7].ToString();
                        }
                        catch(Exception ex)
                        {
                            sheet.Rows.Remove(count);
                            count--;
                            errorCount++;
                            sheetErrors.Rows.Insert(errorCount);
                            sheetErrors.Rows[errorCount].CopyFrom(sheetErrors.Rows[errorCount + 1], PasteSpecial.All);
                            sheetErrors.Rows[errorCount][1].Value = (errorCount - 5).ToString();
                            sheetErrors.Rows[errorCount][2].Value = row[0].ToString();
                            sheetErrors.Rows[errorCount][3].Value = row[1].ToString();
                            sheetErrors.Rows[errorCount][4].Value = row[2].ToString();
                            sheetErrors.Rows[errorCount][5].Value = row[3].ToString();
                            sheetErrors.Rows[errorCount][6].Value = row[4].ToString();
                            sheetErrors.Cells["Q" + (errorCount + 1).ToString()].Value = row[7].ToString();
                            sheetErrors.Cells["R" + (errorCount + 1).ToString()].Value = ex.Message;
                        }                        
                    }
                }
                //sheet.Rows.Remove(count + 1);
            }
            workbookNew.Worksheets.RemoveAt(0);
        }

        public string ConvertToEnglish(string vietnamese)
        {
            switch (vietnamese)
            {
                case "Thứ 2":
                    return "MON";
                case "Thứ 3":
                    return "TUE";
                case "Thứ 4":
                    return "WED";
                case "Thứ 5":
                    return "THU";
                case "Thứ 6":
                    return "FRI";
                case "Thứ 7":
                    return "SAT";
                case "Chủ nhật":
                    return "SUN";
                default:
                    return vietnamese;
            }
        }

        public string SplitFullName(string fullName, bool isFirstName)
        {
            var list = fullName.Split(' ');
            if (isFirstName) return list[list.Count() - 1];
            else
            {
                string lastName = string.Empty;
                foreach(var info in list)
                {
                    if (info != list[list.Count() - 1])
                        lastName += " " + info;
                }
                if (lastName.Length > 0)
                    lastName = lastName.Substring(1, lastName.Length - 1);
                return lastName;
            }            
        }

        public string ConvertDateToString(string strDate)
        {
            try
            {
                return strDate.ToString().Length > 6 ?
                            Convert.ToDateTime(strDate).ToString("dd/MM/yyyy") : strDate;
            }
            catch
            {
                return strDate;
            }
        }
    }
}

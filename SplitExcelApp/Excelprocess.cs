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
    public partial class Excelprocess : RibbonForm
    {
        private DataTable LookupTable;

        public Excelprocess()
        {
            InitializeComponent();
            InitSkinGallery();
            FillLookupTable();            
        }

        void InitSkinGallery()
        {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }

        private void FillLookupTable()
        {
            const int RowCount = 4;
            LookupTable = new DataTable();
            LookupTable.Columns.Add("ID");
            LookupTable.Columns.Add("Name");
            DataRow Row;

            for (int i = 1; i <= RowCount; i++)
            {
                Row = LookupTable.NewRow();
                Row["ID"] = i;
                Row["Name"] = "Qúy " + i.ToString();
                LookupTable.Rows.Add(Row);
            }
            LookupTable.AcceptChanges();
            repositoryItemLookUpEdit1.DataSource = LookupTable;
            repositoryItemLookUpEdit1.ShowHeader = false;
            repositoryItemLookUpEdit1.ValueMember = "ID";
            repositoryItemLookUpEdit1.DisplayMember = "Name";
        }

        private void btnSplitExcel_ItemClick(object sender, ItemClickEventArgs e)
        {
            var firstDateOfQuarter = DateOfQuarter(true);
            var lastDateOfQuarter = DateOfQuarter(false);

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

            var dataErrors = new DataTable("DataError");
            dataErrors.Columns.Add("Name");
            dataErrors.Columns.Add("HPQ2");
            dataErrors.Columns.Add("PaymentDate");
            dataErrors.Columns.Add("PaymentType");
            dataErrors.Columns.Add("HP");
            dataErrors.Columns.Add("Class");
            dataErrors.Columns.Add("Day");
            dataErrors.Columns.Add("Note");
            dataErrors.Columns.Add("ErrorMessage");

            for (int i = 2; i <= usedRange.BottomRowIndex; i++)
            {
                if (!string.IsNullOrEmpty(worksheet.Cells["F" + i.ToString()].Value.ToString()))
                {
                    try
                    {
                        var hpq2 = Convert.ToDouble(worksheet.Cells["K" + i.ToString()].Value.ToObject());
                        var hp = Convert.ToDouble(worksheet.Cells["M" + i.ToString()].Value.ToObject());
                        var paymentDate = Convert.ToDateTime(worksheet.Cells["N" + i.ToString()].Value.ToObject());
                        if (paymentDate >= firstDateOfQuarter && paymentDate <= lastDateOfQuarter)
                        {
                            // add table
                            DataRow dataRow = dataTable.NewRow();
                            dataRow["Name"] = worksheet.Cells["F" + i.ToString()].Value;
                            dataRow["HPQ2"] = hpq2;
                            dataRow["PaymentDate"] = paymentDate;
                            dataRow["PaymentType"] = worksheet.Cells["O" + i.ToString()].Value;
                            dataRow["HP"] = hp;
                            dataRow["Class"] = worksheet.Cells["G" + i.ToString()].Value;
                            dataRow["Day"] = worksheet.Cells["H" + i.ToString()].Value;
                            dataRow["Note"] = worksheet.Cells["C" + i.ToString()].Value;
                            dataTable.Rows.Add(dataRow);
                        }                        
                    }
                    catch(Exception ex)
                    {
                        DataRow dataRow = dataErrors.NewRow();
                        dataRow["Name"] = worksheet.Cells["F" + i.ToString()].Value;
                        dataRow["HPQ2"] = worksheet.Cells["K" + i.ToString()].Value;
                        dataRow["PaymentDate"] = worksheet.Cells["N" + i.ToString()].Value;
                        dataRow["PaymentType"] = worksheet.Cells["O" + i.ToString()].Value;
                        dataRow["HP"] = worksheet.Cells["M" + i.ToString()].Value;
                        dataRow["Class"] = worksheet.Cells["G" + i.ToString()].Value;
                        dataRow["Day"] = worksheet.Cells["H" + i.ToString()].Value;
                        dataRow["Note"] = worksheet.Cells["C" + i.ToString()].Value;
                        dataRow["ErrorMessage"] = ex.Message;
                        dataErrors.Rows.Add(dataRow);                        
                    }                  
                }
            }

            spreadsheetControl.Document.LoadDocument(@"templates\temp.xlsx");            
            IWorkbook workbookNew = spreadsheetControl.Document;
            var sheetErrors = workbookNew.Worksheets.Add("Errors");
            var sheetTotal = workbookNew.Worksheets["TỔNG"];
            workbook.Worksheets[sheetErrors.Name].CopyFrom(workbook.Worksheets["tempSheet"]);
            var errorCount = 5;
            foreach(DataRow row in dataErrors.Rows)
            {
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
                sheetErrors.Cells["R" + (errorCount + 1).ToString()].Value = row["ErrorMessage"].ToString();
            }
                    
            // filter by conditions
            //var datafiltered = (from DataRow dRow in dataTable.Rows select dRow).Where(x => (DateTime)x["Day"] >= firstDateOfQuarter && Convert.ToDateTime(x["Day"]) <= lastDateOfQuarter);
            // sort data
            var dataTableSort = (from DataRow dRow in dataTable.Rows select dRow).OrderBy(x => SplitFullName(x["Name"].ToString(), true)).ThenBy(x => SplitFullName(x["Name"].ToString(), false));
            // get sheets
            var distinctRows = (from DataRow dRow in dataTable.Rows
                                select new { ClassName = dRow["Class"].ToString(), ClassDay = dRow["Day"].ToString() }).Distinct().OrderBy(x => x.ClassName);
          
            foreach (var info in distinctRows)
            {
                var sheet = workbookNew.Worksheets.Add();                
                sheet.Name = info.ClassName + " " + ConvertToEnglish(info.ClassDay);
                workbook.Worksheets[sheet.Name].CopyFrom(workbook.Worksheets["tempSheet"]);
                workbook.Worksheets[sheet.Name].Cells["C2"].Value = "LỚP: " + info.ClassName;
                workbook.Worksheets[sheet.Name].Cells["C3"].Value = info.ClassDay.ToUpper() + "- TH";
                var title = "TỔNG THU HỌC PHÍ LỚP " + info.ClassName +
                    " (" + info.ClassDay + ") QUÝ " + barEditQuarter.EditValue.ToString() + "/" +
                    Convert.ToDateTime(barEditYear.EditValue).ToString("yyyy") + ":  TỪ " + firstDateOfQuarter.ToString("dd/MM/yyyy")
                    + " ĐẾN HẾT " + lastDateOfQuarter.ToString("dd/MM/yyyy");
                workbook.Worksheets[sheet.Name].Cells["E2"].Value = title.ToUpper();
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

                // write sheet total      
                var usedTotal = sheetTotal.GetUsedRange();
                var rowIndexTotal = 0;
                for (int rowIndex = 4; rowIndex <= usedTotal.BottomRowIndex; rowIndex++)
                {
                    if (sheetTotal.Cells["B" + rowIndex.ToString()].Value.ToString().ToUpper() == info.ClassDay.ToUpper())
                    {
                        rowIndexTotal = rowIndex;
                        break;
                    }                        
                }
                if(rowIndexTotal > 0)
                {
                    var rowIndex = count + 8;
                    sheetTotal.Rows.Insert(rowIndexTotal);
                    sheetTotal.Rows[rowIndexTotal].CopyFrom(sheetTotal.Rows[rowIndexTotal - 1], PasteSpecial.All);
                    sheetTotal.Cells["C" + rowIndexTotal.ToString()].Value = info.ClassName.ToUpper();
                    sheetTotal.Cells["D" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$D$" + rowIndex.ToString());
                    sheetTotal.Cells["E" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$E$" + rowIndex.ToString());
                    sheetTotal.Cells["F" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$G$" + rowIndex.ToString());
                    sheetTotal.Cells["G" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$H$" + rowIndex.ToString());
                    sheetTotal.Cells["H" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$I$" + rowIndex.ToString());
                    sheetTotal.Cells["I" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$J$" + rowIndex.ToString());
                    sheetTotal.Cells["J" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$K$" + rowIndex.ToString());
                    sheetTotal.Cells["K" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$L$" + rowIndex.ToString());
                    sheetTotal.Cells["L" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$M$" + rowIndex.ToString());
                    sheetTotal.Cells["M" + rowIndexTotal.ToString()].SetValueFromText("='" + sheet.Name + "'!$N$" + rowIndex.ToString());
                }
            }
            workbookNew.Worksheets.RemoveAt(0);
            // remove line have class value is null
            var usedTotal1 = sheetTotal.GetUsedRange();
            for (int rowIndex = 4; rowIndex <= 200; rowIndex++)
            {
                if (string.IsNullOrEmpty(sheetTotal.Cells["C" + rowIndex.ToString()].Value.ToString()) &&
                    sheetTotal.Cells["A" + rowIndex.ToString()].Value.ToString().Trim().ToUpper() != "TỔNG")
                {
                    sheetTotal.Rows.Remove(rowIndex);
                }
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"D:\";  
            saveFileDialog1.Title = "Khai báo file lưu kết quả";
            //saveFileDialog1.CheckFileExists = true;
            //saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.DefaultExt = "xlsx";
            //saveFileDialog1.Filter = "Text files (*.xlsx)|*.xls|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                workbookNew.SaveDocument(saveFileDialog1.FileName);
                spreadsheetControl.Document.LoadDocument(saveFileDialog1.FileName);
            }            
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

        public DateTime DateOfQuarter(bool isType)
        {
            DateTime date = (DateTime)barEditYear.EditValue;
            //int quarterNumber = (date.Month - 1) / 3 + 1;
            int quarterNumber = Convert.ToInt32(barEditQuarter.EditValue);            
            DateTime firstDayOfQuarter = new DateTime(date.Year, (quarterNumber - 1) * 3 + 1, 1,0,0,0);
            DateTime lastDayOfQuarter = firstDayOfQuarter.AddMonths(3).AddDays(-1).Add(DateTime.MaxValue.TimeOfDay);
            if (isType) return firstDayOfQuarter;
            else return lastDayOfQuarter;
        }
    }
}

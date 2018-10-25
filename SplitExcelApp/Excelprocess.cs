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
using System.Text.RegularExpressions;

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
            try
            {
                if (barEditQuarter.EditValue == null || barEditYear.EditValue == null)
                {
                    MessageBox.Show("Bạn phải chọn điều kiện năm, quý muốn xử lý!", "Thông báo lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

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
                dataTable.Columns.Add("IsError");
                dataTable.Columns.Add("DebitQ1");
                dataTable.Columns.Add("DebitQ2");

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
                            //if (paymentDate >= firstDateOfQuarter && paymentDate <= lastDateOfQuarter)
                            //{

                            //}
                            // add table
                            DataRow dataRow = dataTable.NewRow();
                            dataRow["Name"] = worksheet.Cells["F" + i.ToString()].Value;
                            dataRow["HPQ2"] = hpq2;
                            dataRow["PaymentDate"] = paymentDate;
                            dataRow["PaymentType"] = worksheet.Cells["O" + i.ToString()].Value;
                            dataRow["HP"] = hp;
                            dataRow["Class"] = worksheet.Cells["G" + i.ToString()].Value;
                            dataRow["Day"] = worksheet.Cells["H" + i.ToString()].Value;
                            dataRow["Note"] = worksheet.Cells["Q" + i.ToString()].Value.ToString() + " " + worksheet.Cells["Q" + i.ToString()].Value.ToString() + " " + worksheet.Cells["Q" + i.ToString()].Value.ToString();
                            dataRow["IsError"] = false;
                            dataRow["DebitQ1"] = worksheet.Cells["I" + i.ToString()].Value;
                            dataRow["DebitQ2"] = worksheet.Cells["J" + i.ToString()].Value;                            
                            dataTable.Rows.Add(dataRow);
                        }
                        catch
                        {
                            DataRow dataRow = dataTable.NewRow();
                            dataRow["Name"] = worksheet.Cells["F" + i.ToString()].Value;
                            dataRow["HPQ2"] = worksheet.Cells["K" + i.ToString()].Value;
                            dataRow["PaymentDate"] = worksheet.Cells["N" + i.ToString()].Value;
                            dataRow["PaymentType"] = worksheet.Cells["O" + i.ToString()].Value;
                            dataRow["HP"] = worksheet.Cells["M" + i.ToString()].Value;
                            dataRow["Class"] = worksheet.Cells["G" + i.ToString()].Value;
                            dataRow["Day"] = worksheet.Cells["H" + i.ToString()].Value;
                            dataRow["Note"] = worksheet.Cells["Q" + i.ToString()].Value.ToString() + " " + worksheet.Cells["Q" + i.ToString()].Value.ToString() + " " + worksheet.Cells["Q" + i.ToString()].Value.ToString();
                            dataRow["IsError"] = true;
                            dataRow["DebitQ1"] = worksheet.Cells["I" + i.ToString()].Value;
                            dataRow["DebitQ2"] = worksheet.Cells["J" + i.ToString()].Value;
                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }

                spreadsheetControl.Document.LoadDocument(@"templates\temp.xlsx");
                IWorkbook workbookNew = spreadsheetControl.Document;
                //var sheetErrors = workbookNew.Worksheets.Add("Errors");
                var sheetTotal = workbookNew.Worksheets["TỔNG"];
                sheetTotal.Cells["A1"].Value = "DOANH THU QUÝ " + barEditQuarter.EditValue.ToString() + "/" + Convert.ToDateTime(barEditYear.EditValue).ToString("yyyy");
                //workbook.Worksheets[sheetErrors.Name].CopyFrom(workbook.Worksheets["tempSheet"]);
                //var errorCount = 5;
                //foreach(DataRow row in dataErrors.Rows)
                //{
                //    errorCount++;
                //    sheetErrors.Rows.Insert(errorCount);
                //    sheetErrors.Rows[errorCount].CopyFrom(sheetErrors.Rows[errorCount + 1], PasteSpecial.All);
                //    sheetErrors.Rows[errorCount][1].Value = (errorCount - 5).ToString();
                //    sheetErrors.Rows[errorCount][2].Value = row["Name"].ToString();
                //    sheetErrors.Rows[errorCount][3].Value = row["HP"].ToString();
                //    sheetErrors.Rows[errorCount][4].Value = row["PaymentDate"].ToString();
                //    sheetErrors.Rows[errorCount][5].Value = row["PaymentType"].ToString();
                //    sheetErrors.Rows[errorCount][6].Value = row["HPQ2"].ToString();
                //    sheetErrors.Cells["Q" + (errorCount + 1).ToString()].Value = row[7].ToString();
                //    sheetErrors.Cells["R" + (errorCount + 1).ToString()].Value = row["ErrorMessage"].ToString();
                //}

                // filter by conditions
                //var datafiltered = (from DataRow dRow in dataTable.Rows select dRow).Where(x => (DateTime)x["Day"] >= firstDateOfQuarter && Convert.ToDateTime(x["Day"]) <= lastDateOfQuarter);
                // sort data
                var dataTableSort = (from DataRow dRow in dataTable.Rows select dRow).OrderBy(x => SplitFullName(x["Name"].ToString(), true)).ThenBy(x => SplitFullName(x["Name"].ToString(), false));
                // get sheets
                var distinctRows = (from DataRow dRow in dataTable.Rows
                                    select new { ClassName = dRow["Class"].ToString(), ClassDay = dRow["Day"].ToString() }).Distinct().OrderBy(x => x.ClassDay).ThenBy(x=> x.ClassName);

                foreach (var info in distinctRows)
                {
                    var sheet = workbookNew.Worksheets.Add();
                    if(!ValidateSheetName(info.ClassName + " " + ConvertToEnglish(info.ClassDay)))
                    {
                        MessageBox.Show("Thông báo lỗi", "Tên sheet không đúng định dạng cho phép của Excel: " + info.ClassName + " " + ConvertToEnglish(info.ClassDay), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
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
                            count++;
                            sheet.Rows.Insert(count);
                            sheet.Rows[count].CopyFrom(sheet.Rows[count + 1], PasteSpecial.All);
                            sheet.Rows[count][1].Value = (count - 5).ToString();
                            sheet.Rows[count][2].Value = row["Name"].ToString();
                            sheet.Rows[count][3].SetValue(ConvertToNumber(row["HP"]));
                            sheet.Rows[count][4].Value = ConvertDateToString(row["PaymentDate"]);
                            sheet.Rows[count][5].Value = row["PaymentType"].ToString();
                            //sheet.Rows[count][6].Value = row["HPQ2"].ToString().Length > 0 ? Convert.ToDouble(row["HPQ2"]) : 0;
                            var strCount = (count + 1).ToString();
                            sheet.Rows[count][6].SetValueFromText("=D" + strCount + "-(H" + strCount + "+L" + strCount + "+M" + strCount
                                + "+N" + strCount + "+O" + strCount + ")+(J" + strCount + "+K" + strCount + ")");
                            sheet.Cells["P" + (count + 1).ToString()].Value = row[7].ToString();
                            sheet.Rows[count]["H"].SetValue(ConvertToNumber(row["DebitQ1"]));
                            sheet.Rows[count]["J"].SetValue(ConvertToNumber(row["DebitQ2"]));
                            if (Convert.ToBoolean(row["IsError"]))
                            {
                                Range range = sheet.Rows[count].GetRangeWithAbsoluteReference();
                                Formatting rangeFormatting = range.BeginUpdateFormatting();
                                rangeFormatting.Font.Color = Color.Red;
                                rangeFormatting.Fill.BackgroundColor = Color.Yellow;                                
                                range.EndUpdateFormatting(rangeFormatting);
                            }
                        }
                    }

                    Application.DoEvents();
                    // if counting less paramerter_row_total than rows                    
                    while (count < (Convert.ToInt32(barEditRows.EditValue == null ? 30 : barEditRows.EditValue) + 5))
                    {                        
                        count++;
                        sheet.Rows.Insert(count);
                        sheet.Rows[count].CopyFrom(sheet.Rows[count + 1], PasteSpecial.All);
                        sheet.Rows[count][1].Value = (count - 5).ToString();
                        sheet.Rows[count][2].Value = string.Empty;
                        sheet.Rows[count][3].Value = 0;
                        sheet.Rows[count][4].Value = string.Empty;
                        sheet.Rows[count][5].Value = string.Empty;
                        //sheet.Rows[count][6].Value = row["HPQ2"].ToString().Length > 0 ? Convert.ToDouble(row["HPQ2"]) : 0;
                        var strCount = (count + 1).ToString();
                        sheet.Rows[count][6].SetValueFromText("=D" + strCount + "-(H" + strCount + "+L" + strCount + "+M" + strCount
                            + "+N" + strCount + "+O" + strCount + ")+(J" + strCount + "+K" + strCount + ")");
                        sheet.Cells["Q" + (count + 1).ToString()].Value = string.Empty;
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
                    if (rowIndexTotal > 0)
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
                        if (sheetTotal.Cells["B" + (rowIndexTotal + 1).ToString()].Value == sheetTotal.Cells["B" + rowIndexTotal.ToString()].Value)
                            sheetTotal.MergeCells(sheetTotal.Range["B" + (rowIndexTotal + 1).ToString() + ":" + "B" + rowIndexTotal.ToString()]);
                    }
                }
                workbookNew.Worksheets.RemoveAt(0);

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
            catch(Exception ex)
            {
                MessageBox.Show("Thông báo lỗi", "Đã có lỗi xảy ra trong quá trình xử lý: " + ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        public string ConvertDateToString(object strDate)
        {
            try
            {
                return Convert.ToDateTime(strDate).ToString("dd/MM/yyyy");
            }
            catch
            {
                return strDate.ToString();
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

        public bool ValidateSheetName(string wsName)
        {
            Match m = Regex.Match(wsName, @"[\[/\?\]\*]");
            return (m.Success || (string.IsNullOrEmpty(wsName)) || (wsName.Length > 31)) ? false : true;
        }

        public object ConvertToNumber(object numberIn)
        {
            try
            {
                if (string.IsNullOrEmpty(numberIn.ToString().Trim()))
                    return Convert.ToDouble(0);
                return Convert.ToDouble(numberIn);
            }
            catch
            {
                return numberIn.ToString();
            }
        }
    }
}

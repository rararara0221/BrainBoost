using CsvHelper;
using Microsoft.AspNetCore.Mvc;
using BrainBoost.Models;
using QuestAI.Parameter;
using BrainBoost.Services;
using System.Globalization;
using System;
using System.Data;
using System.IO;
using Microsoft.AspNetCore.Http;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace BrainBoost.Controllers
{
    [Route("BrainBoost/[controller]")]
    [ApiController]
    public class FileController : Controller
    {
        #region 呼叫Service
        private readonly QuestionsDBService QuestionService;

        public FileController(QuestionsDBService questionService)
        {
            QuestionService = questionService;
        }
        #endregion

        // 讀取 是非題Excel檔案
        [HttpPost("[Action]")]
        public IActionResult UploadExcel_Tfq(IFormFile file)
        {
            #region 檔案處理

            // 上傳的文件
            Stream stream = file.OpenReadStream();

            // 儲存Excel的資料
            DataTable dataTable = new DataTable();

            // 讀取or處理Excel文件
            IWorkbook wb;

            // 工作表
            ISheet sheet;

            // 標頭
            IRow headerRow;

            // 欄數
            int cellCount;

            try
            {
                // excel版本(.xlsx)
                if (file.FileName.ToUpper().EndsWith("XLSX"))
                    wb = new XSSFWorkbook(stream);
                // excel版本(.xls)
                else
                    wb = new HSSFWorkbook(stream);

                // 取第一個工作表
                sheet = wb.GetSheetAt(0);

                // 此工作表的第一列
                headerRow = sheet.GetRow(0);

                // 計算欄位數
                cellCount = headerRow.LastCellNum;

                // 讀取標題列，將抓到的值放進DataTable做完欄位名稱
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    dataTable.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue));

                int column = 1; //計算每一列讀到第幾個欄位

                // 略過標題列，處理到最後一列
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    // 取目前的列
                    IRow row = sheet.GetRow(i);

                    // 若該列的第一個欄位無資料，break跳出
                    if (string.IsNullOrEmpty(row.Cells[0].ToString().Trim())) break;
                    
                    // 宣告DataRow
                    DataRow dataRow = dataTable.NewRow();

                    // 宣告ICell，獲取單元格的資訊
                    ICell cell;

                    try
                    {
                        // 依先前取得，依每一列的欄位數，逐一設定欄位內容
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            // 計算每一列讀到第幾個欄位（For錯誤訊息）
                            column = j + 1;

                            // 設定cell為目前第j欄位
                            cell = row.GetCell(j);

                            // 若cell有值
                            if (cell != null)
                            {
                                // 判斷資料格式
                                switch (cell.CellType)
                                {
                                    // 字串
                                    case NPOI.SS.UserModel.CellType.String:
                                        if (cell.StringCellValue != null)
                                            // 設定dataRow第j欄位的值，cell以字串型態取值
                                            dataRow[j] = cell.StringCellValue;
                                        else
                                            dataRow[j] = "";
                                        break;

                                    // 數字
                                    case NPOI.SS.UserModel.CellType.Numeric:
                                        // 日期
                                        if (HSSFDateUtil.IsCellDateFormatted(cell))
                                            // 設定dataRow第j欄位的值，cell以日期格式取值
                                            dataRow[j] = DateTime.FromOADate(cell.NumericCellValue).ToString("yyyy/MM/dd HH:mm");
                                        else
                                            // 非日期格式
                                            dataRow[j] = cell.NumericCellValue;
                                        break;

                                    // 布林值
                                    case NPOI.SS.UserModel.CellType.Boolean:
                                        // 設定dataRow第j欄位的值，cell以布林型態取值
                                        dataRow[j] = cell.BooleanCellValue;
                                        break;

                                    //空值
                                    case NPOI.SS.UserModel.CellType.Blank:
                                        dataRow[j] = "";
                                        break;

                                    // 預設
                                    default:
                                        dataRow[j] = cell.StringCellValue;
                                        break;
                                }
                            }
                        }
                        // DataTable加入dataRow
                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception e)
                    {
                        //錯誤訊息
                        throw new Exception("第 " + i + "列，資料格式有誤:\r\r" + e.ToString());
                    }
                }


            }
            catch (Exception e)
            {
                return BadRequest(e.Message.ToString());
            }
            finally
            {
                //釋放資源
                sheet = null;
                wb = null;
                stream.Dispose();
                stream.Close();
            }
            #endregion

            // 將dataTable資料匯入資料庫
            foreach (DataRow dataRow in dataTable.Rows)
            {
                QuestionList question = new QuestionList();

                question.QuestionData = new Question()
                {
                    type_id = 1,
                    question_content = dataRow["Question"].ToString()
                };

                question.AnswerData = new Answer()
                {
                    option_content = dataRow["Answer"].ToString(),
                    question_parse = dataRow["Parse"].ToString()
                };

                try
                {
                    QuestionService.InsertQuestion(question);
                }
                catch (Exception e)
                {
                    return BadRequest($"發生錯誤:  {e}");
                }
            }
            return Ok("匯入成功");    
        }
    
        // 讀取 選擇題Excel檔案
        [HttpPost("[Action]")]
        public IActionResult UploadExcel_Mcq(IFormFile file)
        {
            #region 檔案處理

            // 上傳的文件
            Stream stream = file.OpenReadStream();

            // 儲存Excel的資料
            DataTable dataTable = new DataTable();

            // 讀取or處理Excel文件
            IWorkbook wb;

            // 工作表
            ISheet sheet;

            // 標頭
            IRow headerRow;

            // 欄數
            int cellCount;

            try
            {
                // excel版本(.xlsx)
                if (file.FileName.ToUpper().EndsWith("XLSX"))
                    wb = new XSSFWorkbook(stream);
                // excel版本(.xls)
                else
                    wb = new HSSFWorkbook(stream);

                // 取第一個工作表
                sheet = wb.GetSheetAt(0);

                // 此工作表的第一列
                headerRow = sheet.GetRow(0);

                // 計算欄位數
                cellCount = headerRow.LastCellNum;

                // 讀取標題列，將抓到的值放進DataTable做完欄位名稱
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    dataTable.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue));

                int column = 1; //計算每一列讀到第幾個欄位

                // 略過標題列，處理到最後一列
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    // 取目前的列
                    IRow row = sheet.GetRow(i);

                    // 若該列的第一個欄位無資料，break跳出
                    if (string.IsNullOrEmpty(row.Cells[0].ToString().Trim())) break;
                    
                    // 宣告DataRow
                    DataRow dataRow = dataTable.NewRow();

                    // 宣告ICell，獲取單元格的資訊
                    ICell cell;

                    try
                    {
                        // 依先前取得，依每一列的欄位數，逐一設定欄位內容
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            // 計算每一列讀到第幾個欄位（For錯誤訊息）
                            column = j + 1;

                            // 設定cell為目前第j欄位
                            cell = row.GetCell(j);

                            // 若cell有值
                            if (cell != null)
                            {
                                // 判斷資料格式
                                switch (cell.CellType)
                                {
                                    // 字串
                                    case NPOI.SS.UserModel.CellType.String:
                                        if (cell.StringCellValue != null)
                                            // 設定dataRow第j欄位的值，cell以字串型態取值
                                            dataRow[j] = cell.StringCellValue;
                                        else
                                            dataRow[j] = "";
                                        break;

                                    // 數字
                                    case NPOI.SS.UserModel.CellType.Numeric:
                                        // 日期
                                        if (HSSFDateUtil.IsCellDateFormatted(cell))
                                            // 設定dataRow第j欄位的值，cell以日期格式取值
                                            dataRow[j] = DateTime.FromOADate(cell.NumericCellValue).ToString("yyyy/MM/dd HH:mm");
                                        else
                                            // 非日期格式
                                            dataRow[j] = cell.NumericCellValue;
                                        break;

                                    // 布林值
                                    case NPOI.SS.UserModel.CellType.Boolean:
                                        // 設定dataRow第j欄位的值，cell以布林型態取值
                                        dataRow[j] = cell.BooleanCellValue;
                                        break;

                                    //空值
                                    case NPOI.SS.UserModel.CellType.Blank:
                                        dataRow[j] = "";
                                        break;

                                    // 預設
                                    default:
                                        dataRow[j] = cell.StringCellValue;
                                        break;
                                }
                            }
                        }
                        // DataTable加入dataRow
                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception e)
                    {
                        //錯誤訊息
                        throw new Exception("第 " + i + "列，資料格式有誤:\r\r" + e.ToString());
                    }
                }


            }
            catch (Exception e)
            {
                return BadRequest(e.Message.ToString());
            }
            finally
            {
                //釋放資源
                sheet = null;
                wb = null;
                stream.Dispose();
                stream.Close();
            }
            #endregion

            // 將dataTable資料匯入資料庫
            foreach (DataRow dataRow in dataTable.Rows)
            {
                QuestionList question = new QuestionList();

                question.QuestionData = new Question()
                {
                    type_id = 2,
                    question_content = dataRow["Question"].ToString()
                };

                question.Options = new List<string>
                {
                    dataRow["OptionA"].ToString(),
                    dataRow["OptionB"].ToString(),
                    dataRow["OptionC"].ToString(),
                    dataRow["OptionD"].ToString()
                };

                question.AnswerData = new Answer()
                {
                    option_content = dataRow["Answer"].ToString(),
                    question_parse = dataRow["Parse"].ToString()
                };

                try
                {
                    QuestionService.InsertQuestion(question);
                }
                catch (Exception e)
                {
                    return BadRequest($"發生錯誤:  {e}");
                }
            }
            return Ok("匯入成功");    
        }
    
        // 讀取 填充題Excel檔案
        [HttpPost("[Action]")]
        public IActionResult UploadExcel_Fq(IFormFile file)
        {
            #region 檔案處理

            // 上傳的文件
            Stream stream = file.OpenReadStream();

            // 儲存Excel的資料
            DataTable dataTable = new DataTable();

            // 讀取or處理Excel文件
            IWorkbook wb;

            // 工作表
            ISheet sheet;

            // 標頭
            IRow headerRow;

            // 欄數
            int cellCount;

            try
            {
                // excel版本(.xlsx)
                if (file.FileName.ToUpper().EndsWith("XLSX"))
                    wb = new XSSFWorkbook(stream);
                // excel版本(.xls)
                else
                    wb = new HSSFWorkbook(stream);

                // 取第一個工作表
                sheet = wb.GetSheetAt(0);

                // 此工作表的第一列
                headerRow = sheet.GetRow(0);

                // 計算欄位數
                cellCount = headerRow.LastCellNum;

                // 讀取標題列，將抓到的值放進DataTable做完欄位名稱
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    dataTable.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue));

                int column = 1; //計算每一列讀到第幾個欄位

                // 略過標題列，處理到最後一列
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    // 取目前的列
                    IRow row = sheet.GetRow(i);

                    // 若該列的第一個欄位無資料，break跳出
                    if (string.IsNullOrEmpty(row.Cells[0].ToString().Trim())) break;
                    
                    // 宣告DataRow
                    DataRow dataRow = dataTable.NewRow();

                    // 宣告ICell，獲取單元格的資訊
                    ICell cell;

                    try
                    {
                        // 依先前取得，依每一列的欄位數，逐一設定欄位內容
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            // 計算每一列讀到第幾個欄位（For錯誤訊息）
                            column = j + 1;

                            // 設定cell為目前第j欄位
                            cell = row.GetCell(j);

                            // 若cell有值
                            if (cell != null)
                            {
                                // 判斷資料格式
                                switch (cell.CellType)
                                {
                                    // 字串
                                    case NPOI.SS.UserModel.CellType.String:
                                        if (cell.StringCellValue != null)
                                            // 設定dataRow第j欄位的值，cell以字串型態取值
                                            dataRow[j] = cell.StringCellValue;
                                        else
                                            dataRow[j] = "";
                                        break;

                                    // 數字
                                    case NPOI.SS.UserModel.CellType.Numeric:
                                        // 日期
                                        if (HSSFDateUtil.IsCellDateFormatted(cell))
                                            // 設定dataRow第j欄位的值，cell以日期格式取值
                                            dataRow[j] = DateTime.FromOADate(cell.NumericCellValue).ToString("yyyy/MM/dd HH:mm");
                                        else
                                            // 非日期格式
                                            dataRow[j] = cell.NumericCellValue;
                                        break;

                                    // 布林值
                                    case NPOI.SS.UserModel.CellType.Boolean:
                                        // 設定dataRow第j欄位的值，cell以布林型態取值
                                        dataRow[j] = cell.BooleanCellValue;
                                        break;

                                    //空值
                                    case NPOI.SS.UserModel.CellType.Blank:
                                        dataRow[j] = "";
                                        break;

                                    // 預設
                                    default:
                                        dataRow[j] = cell.StringCellValue;
                                        break;
                                }
                            }
                        }
                        // DataTable加入dataRow
                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception e)
                    {
                        //錯誤訊息
                        throw new Exception("第 " + i + "列，資料格式有誤:\r\r" + e.ToString());
                    }
                }


            }
            catch (Exception e)
            {
                return BadRequest(e.Message.ToString());
            }
            finally
            {
                //釋放資源
                sheet = null;
                wb = null;
                stream.Dispose();
                stream.Close();
            }
            #endregion

            // 將dataTable資料匯入資料庫
            foreach (DataRow dataRow in dataTable.Rows)
            {
                QuestionList question = new QuestionList();

                question.QuestionData = new Question()
                {
                    type_id = 3,
                    question_content = dataRow["Question"].ToString()
                };

                question.AnswerData = new Answer()
                {
                    option_content = dataRow["Answer"].ToString(),
                    question_parse = dataRow["Parse"].ToString()
                };

                try
                {
                    QuestionService.InsertQuestion(question);
                }
                catch (Exception e)
                {
                    return BadRequest($"發生錯誤:  {e}");
                }
            }
            return Ok("匯入成功");    
        }
    }
}

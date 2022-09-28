using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Table;
using RetrieveExcelData.Models;
using System.Data;
using System.Diagnostics;

namespace RetrieveExcelData.Controllers
{
    public class HomeController : Controller
    {
        private IWebHostEnvironment _hostEnv;
        const string number = "Number";

        public static CellList cellList = new CellList { };
        public static DataTable extractedTable = new DataTable();
        public static List<Cells> inputcells = new List<Cells> { };
        public static List<OutputCells> outputcells = new List<OutputCells> { };

        public HomeController(IWebHostEnvironment hostEnv)
        {
            _hostEnv = hostEnv;
        }
        
        [HttpGet]
        public async Task<IActionResult> Index()
        {
            try
            {
                var file = await GetFile();
                if (file != null)
                   cellList.ImageMessage = $"1 file uploaded";
            }
            catch (Exception ex)
            {
                FileError(ex);
            }
            
            return View(cellList);
        }

        [HttpPost]
        public async Task<IActionResult> Extract(CellList model)
        {
            if (model.OutputData.Count() == 0)
            {
                ModelState.AddModelError("", "Please add output cells to the form to define what cells you need to extract.");
                return View("Index", cellList);
            }
            try
            {
                var updatedmodel = await RetrieveCells(model);
                cellList = updatedmodel;

                for (int i = 0; i < model.OutputData.Count(); i++)
                {
                    var output = model.OutputData[i];

                    try
                    {
                        var extractedmodel = ExtractRange(output.OutputSheetName, output.Range);

                        if (extractedmodel.DataTable != null)
                        {
                            var dataTable = new RangeTables();
                            dataTable.SheetName = output.OutputSheetName;
                            dataTable.CellAddress = output.Range;
                            dataTable.ExtractedList = extractedmodel.DataTable;

                            cellList.ExtractedTables.Add(dataTable);
                        }
                        if (extractedmodel.DataCell != null)
                        {
                            var dataCell = new SingleCellsTable();
                            dataCell.SheetName = output.OutputSheetName;
                            dataCell.CellAddress = output.Range;
                            dataCell.CellValue = extractedmodel.DataCell;

                            cellList.ExtractedCells.Add(dataCell);
                        }
                    }
                    catch (Exception ex)
                    {
                        var error = "";
                        if (ex.Message.Contains("Sequence contains no matching element"))
                            error = "Please ensure the sheet names you have entered are correct.";

                        if (ex.Message.Contains("SkipNumberOfRowsEnd was out of range"))
                            error = "Please ensure the range you have entered is correct. Range should be alphanumberic with a colon separator e.g. A1:D10. " +
                                   "The number on the right side of the range should be equal to or greater than the number on the left side.";

                        if (ex.Message.Contains("Error saving file"))
                            error = $@"Please ensure you close any sheet named {output.OutputSheetName} you may have opened.";

                        if (ex.Message.Contains("Duplicate column name"))
                            error = $@"Please ensure your output range returns a header and that the header row has no duplicates.";

                        if (ex.Message.Contains("Value cannot be null. (Parameter 'DataColumnName')"))
                            error = $@"Please ensure no cell or cells in the first row of the output range is empty.";
                        

                        ModelState.AddModelError("", string.IsNullOrWhiteSpace(error) ? ex.Message : error);

                        return View("Index", cellList);
                    }
                }

                return RedirectToAction("Display", cellList);
            }
            catch (Exception ex)
            {
                FileError(ex);
            }

            return View("Index", cellList);
        }

        [HttpPost]
        public IActionResult AddInputCell(CellList model)
        {
            updateFormData(model);
            inputcells.Add(new Cells());

            return RedirectToAction("Index", cellList);
        }

        [HttpPost]
        public IActionResult AddOutputCell(CellList model)
        {
            updateFormData(model);
            outputcells.Add(new OutputCells());

            return RedirectToAction("Index", cellList);
        }

        
        [HttpPost]
        public IActionResult ClearInputTable(CellList model)
        {
            updateFormData(model);
            inputcells = new List<Cells> { };
            cellList.InputData = inputcells;

            return RedirectToAction("Index", cellList);
        }

        [HttpPost]
        public IActionResult ClearOutputTable(CellList model)
        {
            updateFormData(model);
            outputcells = new List<OutputCells> { };
            cellList.OutputData = outputcells;

            return RedirectToAction("Index", cellList);
        }

        [HttpPost]
        public IActionResult DeleteInput(CellList model, string deleteinput)
        {
            if (deleteinput != null)
            {
                int i = Convert.ToInt32(deleteinput);
                model.InputData.RemoveAt(i);

                updateFormData(model);
            }

            return RedirectToAction("Index", cellList);
        }

        [HttpPost]
        public IActionResult DeleteOutput(CellList model, string deleteoutput)
        {
            if (deleteoutput != null)
            {
                int i = Convert.ToInt32(deleteoutput);
                model.OutputData.RemoveAt(i);

                updateFormData(model);
            }

            return RedirectToAction("Index", cellList);
        }

        [HttpPost]
        public async Task<IActionResult> UploadFiles()
        {
            var file = Request.Form.Files.FirstOrDefault();

            string message = "";
            cellList.ImageMessage = message;

            if (file == null)
            {
                message = $"upload an excel file.";
                cellList.ImageMessage = message;

                return Json(message);
            }

            string fileName = Path.GetFileName(file.FileName);
            string uploadedfileName = _hostEnv.WebRootPath + $@"\UploadedFiles\{file.FileName}";
            string excelPath = Path.GetFullPath(uploadedfileName);

            FileInfo fileInfo = new FileInfo(excelPath);
            HttpContext.Session.SetString("filename", file.FileName);

            cellList.FileName = file.FileName;
            cellList.File = fileInfo;

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    // create and save a copy of the excel file.
                    using (FileStream fs = System.IO.File.Create(uploadedfileName))
                    {
                        file.CopyTo(fs);
                        fs.Flush();
                    }
                }

                message = $"1 file uploaded";
                cellList.ImageMessage = message;

                return Json(message);
            }
            catch (Exception ex)
            {
                if (!fileName.EndsWith(".xlsx") || !fileName.EndsWith(".xls"))
                {
                    message = $"only excel files are allowed.";
                    cellList.ImageMessage = message;

                    return Json(message);
                }

                cellList.ImageMessage = ex.Message;
                return Json(ex.Message);
            }
        }


        private async Task<CellList> RetrieveCells(CellList model)
        {
            var file = await GetFile();

            // Creating and use an instance of ExcelPackage

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                for (int i = 0; i < model.InputData.Count(); i++)
                {
                    var inputsheet = model.InputData[i].SheetName;
                    if (!string.IsNullOrWhiteSpace(inputsheet))
                    {
                        ExcelWorksheet inputws = excelPackage.Workbook.Worksheets.First(x => x.Name.ToLower() == model.InputData[i].SheetName.ToLower());

                        if (inputws == null)
                            throw new ArgumentNullException(nameof(Cells.SheetName), $@"{model.InputData[i].SheetName} worksheet does not exist.");

                        if (!string.IsNullOrWhiteSpace(model.InputData[i].CellAddress))
                            inputws.Cells[model.InputData[i].CellAddress].Value = model.InputData[i].DataType == number ? SetCelllValueNumber(model.InputData[i].DataType, model.InputData[i].CellValue)
                                : SetCelllValueString(model.InputData[i].CellValue, model.InputData[i].CellAddress);

                    }
                }
                excelPackage.Workbook.Calculate();
                excelPackage.Save();
            }

            return model;
        }

        private ExtractObject ExtractRange(string extractSheetName, string rangeAddress)
        {
            if (string.IsNullOrWhiteSpace(extractSheetName) || string.IsNullOrWhiteSpace(rangeAddress))
                throw new ArgumentNullException(nameof(OutputCells.OutputSheetName), $@"Output worksheet name and range must be provided.");

            var file = GetFile();

            using (var excelPackage = new ExcelPackage(file.Result))
            {
                var worksheet = excelPackage.Workbook.Worksheets.First(x => x.Name.ToLower() == extractSheetName.ToLower());
                if (worksheet == null)
                    throw new ArgumentNullException(nameof(OutputCells.OutputSheetName), $@"{extractSheetName} worksheet does not exist.");

                string[] cells = rangeAddress.Split(':');

                var data = new ExtractObject();

                if (cells.Count() > 1)
                {
                    var leftside = cells[0].Replace(" ", "");
                    var rightside = cells[1].Replace(" ", "");

                    ExcelCellAddress start = new ExcelCellAddress(leftside);
                    ExcelCellAddress end = new ExcelCellAddress(rightside);

                    //ExcelWorkSheet.Cells[FromRow, FromCol, ToRow, ToCol]

                    ExcelRange range = worksheet.Cells[start.Row, start.Column, end.Row, end.Column];

                    var columnStartNumber = worksheet.Cells[leftside].Start.Column;
                    var columnEndNumber = worksheet.Cells[rightside].Start.Column;

                    var rowStartNumber = worksheet.Cells[leftside].Start.Row;
                    var rowEndNumber = worksheet.Cells[rightside].Start.Row;

                    //loop through rows and columns
                    for (int row = rowStartNumber; row <= rowEndNumber; row++)
                    {
                        for (int col = columnStartNumber; col <= columnEndNumber; col++)
                        {
                            //check if the cell is empty or not
                            if (worksheet.Cells[row, col].Value == null)
                            {
                                worksheet.Cells[row, col].Value = "";
                            }
                        }
                    }

                    range = worksheet.Cells[start.Row, start.Column, end.Row, end.Column];

                    if (end.Row == start.Row)
                    {

                        data.DataTable = GetDatatableFromExcel( worksheet, false, range, columnEndNumber, rowEndNumber);
                    }
                    else
                    {
                        data.DataTable = range.ToDataTable();
                    }
                   
                    range.Calculate();
                    excelPackage.Save();
                }
                else
                {
                    var cell = worksheet.Cells[rangeAddress].Value;
                    if (cell != null)
                    {
                        data.DataCell = cell.ToString();
                    }
                    else
                    {
                        data.DataCell = "";
                    }
                }

                return data;
            }
        }

        private static DataTable GetDatatableFromExcel(ExcelWorksheet worksheet, bool hasHeader, ExcelRange range, int columnEndNumber, int rowEndNumber)
        {

            DataTable tbl = new DataTable();

            //range = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
            //rowEndNumber = worksheet.Dimension.End.Row
            //columnEndNumber = worksheet.Dimension.End.Column;

            foreach (var firstRowCell in range)
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (int rowNum = startRow; rowNum <= rowEndNumber; rowNum++)
            {
                var wsRow = worksheet.Cells[rowNum, 1, rowNum, columnEndNumber];
                DataRow row = tbl.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }
            return tbl;
        }

        private async Task<FileInfo> GetFile()
        {
            string[] filePaths = Directory.GetFiles(_hostEnv.WebRootPath + @"\UploadedFiles\", "*.xlsx",
                                           SearchOption.TopDirectoryOnly);

            var filename = HttpContext.Session.GetString("filename");

            FileInfo excelfile;

            if (string.IsNullOrEmpty(filename))
                throw new ArgumentNullException("", $@"Can not find file, please upload excel fine.");

            if (filePaths.ToList().Count>0)
            {
                excelfile = new FileInfo(filePaths.FirstOrDefault(x => x.Contains(filename)));
            }
            else
            {
                string uploadedfileName = _hostEnv.WebRootPath + $@"\UploadedFiles\{filename}";
                string excelPath = Path.GetFullPath(uploadedfileName);
                excelfile = new FileInfo(excelPath);
            }

            // set the license to noncommercial to use it for testing / demo purposes.=
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            cellList.FileName = filename;
            cellList.File = excelfile;

            return excelfile;
        }

        private static string ConvertDataTableToHTML(DataTable dt)
        {
            const string quote = "\"";
            string html = "<table" + " " + "class=" + quote + "table table-responsive table-hover table-sm" + quote + " " + ">";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<td>" + dt.Columns[i].ColumnName + "</td>";
            html += "</tr>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }

        private double SetCelllValueNumber(string datatype, string cellvalue)
        {
            if (datatype == number && string.IsNullOrWhiteSpace(cellvalue))
            {
                return double.Parse("0");
            }
            var parsednumber = double.Parse(cellvalue);
            return parsednumber;
        }

        private string SetCelllValueString(string cellvalue, string cellAddress)
        {
            if (string.IsNullOrWhiteSpace(cellAddress))
            {
                return cellAddress;
            }

            return cellvalue ?? string.Empty;
        }

        private CellList updateFormData(CellList model)
        {
            inputcells = model.InputData;
            cellList.InputData = inputcells;

            outputcells = model.OutputData;
            cellList.OutputData = outputcells;

            return model;
        }

        private IActionResult FileError(Exception ex)
        {
            var error = "";
            if (ex.Message.Contains("Can not find file, please upload excel fine"))
                error = "Can not find file, please upload excel fine.";

            ModelState.AddModelError("", string.IsNullOrWhiteSpace(error) ? ex.Message : error);
            return View("Index", cellList);
        }


        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult Display()
        {
            return View(cellList);
        }

        public IActionResult EPPlus()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public class ExtractObject
        {
            public DataTable? DataTable { get; set; }
            public string DataCell { get; set; }
        }
    }
}
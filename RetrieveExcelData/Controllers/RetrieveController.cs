using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using RetrieveExcelData.Models;

namespace RetrieveExcelData.Controllers
{
    public class RetrieveController : Controller
    {
        private IWebHostEnvironment _hostEnv;
        const string number = "Number";

        public static CellList cellList = new CellList { };
        public static List<Cells> inputcells = new List<Cells> { };
        public static List<OutputCells> outputCells = new List<OutputCells> { };

        public RetrieveController(IWebHostEnvironment hostEnv)
        {
            _hostEnv = hostEnv;
        }

      
        public IActionResult RetrieveIndex()
        {
            ViewBag.InputList = cellList.InputData;
            ViewBag.OutputList = cellList.OutputData;


            return View(cellList);
        }

        public async Task<IActionResult> Retrieve(CellList cellList)
        {
            if (!ModelState.IsValid)
            {
                ViewBag.InputList = cellList.InputData;
                ViewBag.OutputList = cellList.OutputData;
                return View("RetrieveIndex", ModelState);
            }

            inputcells.AddRange(cellList.InputData);
            cellList.InputData = inputcells;

            outputCells.AddRange(cellList.OutputData);
            cellList.OutputData = outputCells;

            await RetrieveCells(cellList);

            return RedirectToAction("RetrieveIndex", cellList);
        }

        [HttpPost]
        public IActionResult UploadFiles()
        {
            long size = 0;
            var file = Request.Form.Files.FirstOrDefault();

            string message = "";

            if (file == null)
            {
                message = $"upload an excel file.";
                ModelState.AddModelError("File", message);
                return View();
            }

            string fileName = Path.GetFileName(file.FileName);
            string uploadedfileName = _hostEnv.WebRootPath + $@"\UploadedFiles\{file.FileName}";
            string excelPath = Path.GetFullPath(uploadedfileName);


            FileInfo fileInfo = new FileInfo(excelPath);

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    // create and save a copy of the excel file.
                    using (FileStream fs = System.IO.File.Create(uploadedfileName))
                    {
                        file.CopyTo(fs);
                        fs.Flush();
                    }

                }

                message = $"1 file and {size} bytes uploaded";
                return Json(message);
            }
            catch (Exception ex)
            {
                if (!fileName.EndsWith(".xlsx") || !fileName.EndsWith(".xls"))
                {
                    message = $"only excel files are allowed.";
                    ModelState.AddModelError("File", message);
                    return View();
                }
                ModelState.AddModelError("File", ex.Message);
                return View();
            }
        }

        public async Task<CellList> RetrieveCells(CellList model)
        {
            string[] filePaths = Directory.GetFiles(_hostEnv.WebRootPath + @"\UploadedFiles\", "*.xlsx",
                                           SearchOption.TopDirectoryOnly);

            var excelfile = new FileInfo(filePaths.FirstOrDefault());


            // set the license to noncommercial to use it for testing / demo purposes.=
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Creating and use an instance of ExcelPackage

            using (ExcelPackage excelPackage = new ExcelPackage(excelfile))
            {
                for (int i = 0; i < model.InputData.Count(); i++)
                {
                    var inputsheet = model.InputData[i].SheetName;
                    if (!string.IsNullOrWhiteSpace(inputsheet))
                    {
                        ExcelWorksheet inputws = excelPackage.Workbook.Worksheets.First(x => x.Name.ToLower() == model.InputData[i].SheetName.ToLower());

                        if (!string.IsNullOrWhiteSpace(model.InputData[i].CellAddress))
                            inputws.Cells[model.InputData[i].CellAddress].Value = model.InputData[i].DataType == number ? SetCelllValueNumber(model.InputData[i].DataType, model.InputData[i].CellValue)
                                : SetCelllValueString(model.InputData[i].CellValue, model.InputData[i].CellAddress);

                        inputws.Calculate();
                    }
                }
                for (int i = 0; i < model.OutputData.Count(); i++)
                {
                    var outputsheet = model.OutputData[i].OutputSheetName;
                    if (!string.IsNullOrWhiteSpace(outputsheet))
                    {
                        ExcelWorksheet outputws = excelPackage.Workbook.Worksheets.First(x => x.Name.ToLower() == model.OutputData[i].OutputSheetName.ToLower());

                        if (!string.IsNullOrWhiteSpace(model.OutputData[i].CellAddress))
                            model.OutputData[i].CellValue = Convert.ToString(outputws.Cells[model.OutputData[i].CellAddress].Value);
                    }
                }
            }

            return model;
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

        public IActionResult EPPlus()
        {
            return View();
        }
    }
}

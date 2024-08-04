using excelproj.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;

namespace exccel2.Controllers
{
    public class HomeController : Controller
    {
        /*
      
         */
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {   //giao dien tinh luong co ban
            return View();
        }
        public IActionResult IfIndex()
        {   //giao dien IF co ban
            return View();
        }
        public IActionResult IfAdvanced()
        {   //giao dien IF nang cao
            return View();
        }
        
        public IActionResult Vlookup()
        {   //giao dien Vlookup
            return View();
        }
        public IActionResult Hlookup()
        {   //giao dien Hlookup
            return View();
        }
        public IActionResult Privacy()
        {
            return View();
        }
       
       

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
           if (file != null && file.Length > 0)
            {
                var data = new List<Exceldatamodel>();

                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets.First();
                        var rowCount = worksheet.Dimension.Rows;
                        // thay co the gioi han so hang lai, vd: var rowCount=5;
                        var titles = new List<String>();
                        
                        for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                        {
                            var title = worksheet.Cells[1, i].Value.ToString();
                            titles.Add(title);
                        }
                        ViewBag.Titles = titles;
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var stt = worksheet.Cells[row, 1].Value?.ToString();
                            var slr = worksheet.Cells[row, 6].Value?.ToString();
                            if (int.TryParse(stt, out int column1Int) &&
                                int.TryParse(slr, out int column2Int))
                            {
                                data.Add(new Exceldatamodel
                                {
                                    stt = column1Int,
                                    name = worksheet.Cells[row, 2].Value?.ToString(),
                                    id = worksheet.Cells[row, 3].Value?.ToString(),
                                    sex = worksheet.Cells[row, 4].Value?.ToString(),
                                    department = worksheet.Cells[row, 5].Value?.ToString(),
                                    salary = column2Int,
                                    netsalary = 0
                                });
                            }
                        }
                    }
                }
                return View("Index", data);
            }

            return View("Index");
        }
        [HttpPost]
        public IActionResult Upload2(IFormFile file1)
        {
            _logger.LogInformation(file1.Length.ToString());
            if (file1 != null && file1.Length > 0)
            {
                var data = new List<exceldatamodel2>();

                using (var stream = new MemoryStream())
                {
                    file1.CopyTo(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets.First();
                        var rowCount = worksheet.Dimension.Rows;
                        var titles = new List<String>();
                        for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                        {
                            var title = worksheet.Cells[1, i].Value.ToString();
                            titles.Add(title);
                        }
                        ViewBag.Titles = titles;
                        
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var stt = worksheet.Cells[row, 1].Value?.ToString();
                            var sale = worksheet.Cells[row, 4].Value?.ToString();
                            if (int.TryParse(stt, out int column1Int) &&
                                int.TryParse(sale, out int column2Int))
                            {
                                data.Add(new exceldatamodel2
                                {
                                    stt = column1Int,
                                    mach = worksheet.Cells[row,2].Value.ToString(),
                                    zone = worksheet.Cells[row, 3].Value.ToString(),
                                    sale = column2Int,
                                    rank = ""
                                });
                            }
                        }
                    }
                }
                
                return View("IfIndex", data);
            }

            return View("Ifindex");
        }
        
        [HttpPost]
        public IActionResult Upload3(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var data = new List<exceldatamodel3>();

                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets.First();
                        var rowCount = worksheet.Dimension.Rows;
                        var titles = new List<String>();
                        for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                        {
                            var title = worksheet.Cells[1, i].Value.ToString();
                            titles.Add(title);
                        }
                        ViewBag.Titles = titles;
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var a1 = worksheet.Cells[row, 2].Value?.ToString();
                            var a2 = worksheet.Cells[row, 3].Value?.ToString();
                            var a3 = worksheet.Cells[row, 4].Value?.ToString();
                            var a4 = worksheet.Cells[row, 5].Value?.ToString();
                            var a5 = worksheet.Cells[row, 6].Value?.ToString();

                            if (int.TryParse(a1, out int column1Int) &&
                                int.TryParse(a2, out int column2Int) &&
                                int.TryParse(a3, out int column3Int)&&
                                int.TryParse(a4, out int column4Int) &&
                                int.TryParse(a5, out int column5Int))
                            {
                                data.Add(new exceldatamodel3
                                {
                                    Name = worksheet.Cells[row,1].Value.ToString(),
                                    diem1=column1Int,
                                    diem2=column2Int,
                                    diem3=column3Int,
                                    diem4=column4Int,
                                    diem5=column5Int,
                                    danhgia=""
                                });
                            }
                        }
                    }
                }
                return View("IfAdvanced", data);
            }

            return View("IfAdvanced");
        }
        [HttpPost]
        public IActionResult Upload4(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var data = new List<exceldatamodel4>();

                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets.First();
                        var rowCount = worksheet.Dimension.Rows;
                        var titles = new List<String>();
                        for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                        {   if (i == 5) titles.Add("");
                            else
                            {
                                var title = worksheet.Cells[1, i].Value.ToString();
                                titles.Add(title);
                            }
                        }
                        
                        ViewBag.Titles = titles;
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var a1 = worksheet.Cells[row, 1].Value?.ToString();
                            var a2 = worksheet.Cells[row, 3].Value?.ToString();
                            var a3 = worksheet.Cells[row,6].Value?.ToString();
                            

                            if (int.TryParse(a1, out int column1Int) &&
                                float.TryParse(a2, out float column2Int)
                                )
                            {
                                float.TryParse(a3, out float column3Int);
                                data.Add(new exceldatamodel4
                                {
                                    stt = column1Int,
                                    name = worksheet.Cells[row, 2].Value?.ToString(),
                                    score = column2Int,
                                    rank = "",
                                    empty = "",
                                    score2 = column3Int,
                                    rank2 = worksheet.Cells[row,7].Value?.ToString()

                                }); ;
                                
                            }
                        }
                    }
                }
                return View("Vlookup", data);
            }

            return View("Vlookup");
        }
        [HttpPost]
        public IActionResult Upload5(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var data = new List<exceldatamodel4>();

                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    using (var package = new ExcelPackage(stream))
                    {                          

                        var worksheet = package.Workbook.Worksheets.First();
                        var rowCount = worksheet.Dimension.Rows;

                        var titles = new List<String> { "ho ten", "diem trung binh","xep hang" };
                        

                        ViewBag.Titles = titles;
                        for (int row = 2; row <= rowCount-2; row++)
                        {
                            
                            var a1 = worksheet.Cells[row, 2].Value?.ToString();
                            if (float.TryParse(a1, out float column1Int))                                
                            {
                                
                                data.Add(new exceldatamodel4
                                {
                                    name = worksheet.Cells[row,1].Value.ToString(),
                                    score= column1Int,
                                    rank="",

                                }); 
                                
                            }
                        }
                        for (int col = 2; col <= 5; col++)
                        {
                            data.Add(new exceldatamodel4
                            {
                                score2 = float.Parse(worksheet.Cells[rowCount-1, col].Value?.ToString()),
                                rank2 = worksheet.Cells[rowCount, col].Value?.ToString()
                            }) ;
                            
                        }
                    }
                }
                return View("Hlookup", data);
            }

            return View("Hlookup");
        }
        
        public IActionResult tessting()
        {
            return View();
        }
       
       

        [HttpPost]
        public IActionResult Upload1(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var filePath = SaveFile(file);
                var data = ProcessFile(filePath);

                // Store file path in session for later use
                HttpContext.Session.SetString("UploadedFilePath", filePath);

                return PartialView("_DataTable", data); // Return partial view with data
            }

            return BadRequest("No file uploaded.");
        }

        [HttpPost]
        public IActionResult ExecuteFormula([FromBody] FormulaModel model)
        {
            var filePath = HttpContext.Session.GetString("UploadedFilePath");
            if (string.IsNullOrEmpty(filePath))
            {
                return BadRequest("No file uploaded.");
            }

            var data = ProcessFile(filePath);

            
            var valid = ValidateFormula(model.Formula, data);
            var c2 = 0;
            var c1 = "";
            var c5 = "";
            if (model.Formula.Split("*").Length == 2) {
            c1 = model.Formula.Split("*")[0];
            c5 = c1[1].ToString();
                _logger.LogInformation(c5);
            c2 =int.Parse(model.Formula.Split("*")[1]);}
            // cach 1 : dung validate de kiem tra input thuoc dang gi r duplicate process file 
            // cach 2 : check if formula string is correct roi load ket qua dung only( JS & ajax )
            switch (valid)
            { //git push

                case 1:
                    
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets.First();
                        worksheet.Cells[2, 7].Formula = $"{c1}*{c2}";
                       
                        for (int i = int.Parse(c5); i <= 11; i++) {
                            var c3=$"{c1[0]}{i}";
                            
                        worksheet.Cells[i, 7].Formula = $"{c3}*{c2}";
                        }
                        package.Workbook.Calculate();
                        package.Save();
                        data = ProcessFile(filePath);

                    }
                    break;
                case 2:
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets.First();
                        worksheet.Cells[2,8].Formula = $"sumif(e:e,\"Kinh Doanh\",g:g)";
                        package.Workbook.Calculate();
                        package.Save();
                        data = ProcessFile(filePath);

                    }
                    break;
            }
           

            return PartialView("_DataTable", data); 
        }

        private string SaveFile(IFormFile file)
        {
            var uploads = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");
            if (!Directory.Exists(uploads))
            {
                Directory.CreateDirectory(uploads);
            }

            var filePath = Path.Combine(uploads, file.FileName);
            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                file.CopyTo(fileStream);
            }

            return filePath;
        }

        private List<Exceldatamodel> ProcessFile(string filePath)
        {
            var data = new List<Exceldatamodel>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var stt = worksheet.Cells[row, 1].Value?.ToString();
                    var slr = worksheet.Cells[row, 6].Value?.ToString();

                    if (int.TryParse(stt, out int column1Int) &&
                        int.TryParse(slr, out int column2Int))
                    {
                        int.TryParse(worksheet.Cells[row, 7].Value?.ToString(), out int column3Int);
                        int.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out int column4Int);
                        data.Add(new Exceldatamodel
                        {
                            stt = column1Int,
                            name = worksheet.Cells[row, 2].Value?.ToString(),
                            id = worksheet.Cells[row, 3].Value?.ToString(),
                            sex = worksheet.Cells[row, 4].Value?.ToString(),
                            department = worksheet.Cells[row, 5].Value?.ToString(),
                            salary = column2Int,
                            netsalary = column3Int,
                            slr2=column4Int
                        });
                    }
                }
            }

            return data;
        }

        private int ValidateFormula(string formula, List<Exceldatamodel> data)
        {
            if (string.IsNullOrEmpty(formula))
                return 0;
            if (formula.Split('*').Length == 2)
            {
                
                return 1;
            }
            else if (formula.ToLower().Contains("sumif"))
            {
                return 2;
            }
            else return 0;
                   } 
    }

}
public class FormulaModel
{
    public string? Formula { get; set; }
}

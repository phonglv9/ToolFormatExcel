using OfficeOpenXml;

namespace ToolFormatExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnFormat_Click(object sender, EventArgs e)
        {
            // Đường dẫn đến tệp Excel mẫu của bạn
            string templatePath = @"C:\Users\lovan\Downloads\2022-2023 CN\1A.xlsx";

            // Thư mục chứa các tệp dữ liệu của bạn
            string dataDirectory = @"C:\Users\lovan\Downloads\2022-2023 CN"; // Cập nhật với đường dẫn thư mục dữ liệu của bạn

            // Thư mục đầu ra mà bạn muốn lưu trữ các tệp Excel đã được định dạng
            string outputDirectory = @"C:\Users\lovan\Downloads\ThuMucXuat"; // Cập nhật với đường dẫn thư mục đầu ra của bạn

            // Đảm bảo thư mục đầu ra tồn tại
            Directory.CreateDirectory(outputDirectory);

            // Danh sách các tệp dữ liệu cần xử lý
            List<string> dataFiles = Directory.GetFiles(dataDirectory, "*.xlsx").ToList();


            // Load định dạng từ tệp Excel mẫu
            FileInfo templateFile = new FileInfo(templatePath);
            using (var templatePackage = new ExcelPackage(templateFile))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet templateWorksheet = templatePackage.Workbook.Worksheets[0]; // Chỉ mục của bảng tính trong tệp mẫu

                // Lặp qua mỗi tệp dữ liệu và áp dụng định dạng từ tệp mẫu
                //foreach (string dataFile in dataFiles)
                //{
                //    FileInfo dataFileInfo = new FileInfo(dataFile);

                //    using (var dataPackage = new ExcelPackage(dataFileInfo))
                //    {
                //        ExcelWorksheet dataWorksheet = dataPackage.Workbook.Worksheets[0]; // Chỉ mục của bảng tính trong tệp dữ liệu

                //        // Sao chép định dạng từ tệp mẫu và áp dụng cho tệp dữ liệu
                //        for (int row = 1; row <= templateWorksheet.Dimension.End.Row; row++)
                //        {
                //            for (int col = 1; col <= templateWorksheet.Dimension.End.Column; col++)
                //            {
                //                ExcelRangeBase templateCell = templateWorksheet.Cells[row, col];
                //                ExcelRangeBase dataCell = dataWorksheet.Cells[row, col];

                //                // Sao chép định dạng (kiểu dáng) từ tệp mẫu và áp dụng cho tệp dữ liệu
                //                dataCell.Style.Font.Bold = templateCell.Style.Font.Bold;
                //                dataCell.Style.Numberformat.Format = templateCell.Style.Numberformat.Format;
                //                dataCell.Style.Fill.PatternType = templateCell.Style.Fill.PatternType;
                //                // Sao chép định dạng cột từ tệp mẫu và áp dụng cho tệp dữ liệu
                //                dataWorksheet.Column(col).Width = templateWorksheet.Column(col).Width;
                //                dataWorksheet.Column(col).Style.Font.Size = templateWorksheet.Column(col).Style.Font.Size;
                //                dataWorksheet.Column(col).Style.Font.Name = templateWorksheet.Column(col).Style.Font.Name;
                //                //// Chuyển đổi màu nền từ ExcelColor sang Color
                //                //dataCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(templateCell.Style.Fill.BackgroundColor.Rgb));
                //            }
                //        }

                //        // Lưu tệp dữ liệu đã được định dạng vào thư mục đầu ra
                //        string outputFileName = Path.Combine(outputDirectory, dataFileInfo.Name);
                //        dataPackage.SaveAs(new FileInfo(outputFileName));
                //    }
                //}\

                // Lặp qua mỗi tệp dữ liệu và áp dụng định dạng từ tệp mẫu
                foreach (string dataFile in dataFiles)
                {
                    FileInfo dataFileInfo = new FileInfo(dataFile);

                    using (var dataPackage = new ExcelPackage(dataFileInfo))
                    {
                        ExcelWorksheet dataWorksheet = dataPackage.Workbook.Worksheets[0]; // Chỉ mục của bảng tính trong tệp dữ liệu

                        // Sao chép định dạng cột (bao gồm cả kiểu chữ) từ tệp mẫu và áp dụng cho tệp dữ liệu
                        for (int col = 1; col <= templateWorksheet.Dimension.End.Column; col++)
                        {
                            var templateColumnStyle = templateWorksheet.Column(col).Style;
                            dataWorksheet.Column(col).Width = templateWorksheet.Column(col).Width;

                            for (int row = 1; row <= dataWorksheet.Dimension.End.Row; row++)
                            {
                                var templateCell = templateWorksheet.Cells[row, col];
                                var dataCell = dataWorksheet.Cells[row, col];

                                // Sao chép tất cả thuộc tính định dạng từ tệp mẫu sang tệp dữ liệu
                                dataCell.Style.Font.Name = templateCell.Style.Font.Name;
                                dataCell.Style.Font.Size = templateCell.Style.Font.Size;
                                dataCell.Style.Font.Bold = templateCell.Style.Font.Bold;
                                dataCell.Style.Font.Italic = templateCell.Style.Font.Italic;
                                dataCell.Style.Font.UnderLine = templateCell.Style.Font.UnderLine;
                                dataCell.Style.Font.Strike = templateCell.Style.Font.Strike;                              
                                dataCell.Style.Fill.PatternType = templateCell.Style.Fill.PatternType;
                                dataCell.Style.Fill.PatternType = templateCell.Style.Fill.PatternType;
                                //dataCell.Style.Border.Top.Style = templateCell.Style.Border.Top.Style;                         
                                //dataCell.Style.Border.Left.Style = templateCell.Style.Border.Left.Style;                            
                                //dataCell.Style.Border.Right.Style = templateCell.Style.Border.Right.Style;                              
                                //dataCell.Style.Border.Bottom.Style = templateCell.Style.Border.Bottom.Style;
                                dataCell.Style.TextRotation = templateCell.Style.TextRotation;

                            }
                        }

                        // Lưu tệp dữ liệu đã được định dạng vào thư mục đầu ra
                        string outputFileName = Path.Combine(outputDirectory, dataFileInfo.Name);
                        dataPackage.SaveAs(new FileInfo(outputFileName));
                    }
                }
            }

            Console.WriteLine("Tất cả các tệp Excel đã được định dạng và đã được lưu.");
        }

        private void btnFileTemplate_Click(object sender, EventArgs e)
        {

        }
    }
}
using System.Drawing;
using System.Text.Json.Serialization;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.SystemDrawing.Image;
using System.Reflection;


class Program
{
	public class Part
	{
		[JsonPropertyName("brickLinkId")]
		public string BrickLinkId { get; set; } = "";

		[JsonPropertyName("name")]
		public string Name { get; set; } = "";

		[JsonPropertyName("imgUrl")]
		public string ImgUrl { get; set; } = "";

		[JsonPropertyName("imgPath")]
		public string ImgPath { get; set; } = "";

		[JsonPropertyName("amountNeeded")]
		public int AmountNeeded { get; set; } = 0;

		[JsonPropertyName("amountFound")]
		public int AmountFound { get; set; } = 0;
	}

	public class Set
	{
		[JsonPropertyName("name")]
		public string Name { get; set; } = "";

		[JsonPropertyName("id")]
		public string Id { get; set; } = "";

		[JsonPropertyName("parts")]
		public Part[] Parts { get; set; } = new Part[] { };
	}

	static void Main(string[] args)
	{
		string assemblyPath = Assembly.GetExecutingAssembly().Location;
		string directoryPath = Path.GetDirectoryName(assemblyPath) ?? "";

		if (directoryPath == "")
		{
			Console.Error.WriteLine("Failed to get directory path.");
			return;
		}

		directoryPath = directoryPath.Substring(0, directoryPath.LastIndexOf("bin\\Debug\\net8.0"));

		string jsonString = File.ReadAllText(directoryPath + "..\\set.json");
		Set set = System.Text.Json.JsonSerializer.Deserialize<Set>(jsonString)!;

		var excel = new ExcelPackage();

		if (File.Exists($"{directoryPath}..\\sets\\{set.Name}.xlsm"))
		{
			var existingExcel = new ExcelPackage(new FileInfo($"{directoryPath}..\\sets\\{set.Name}.xlsm"));
			var existingSetSheet = existingExcel.Workbook.Worksheets[0];

			for (int index = 2; index < existingSetSheet.Rows.Count() + 1; index++)
			{
				string name = existingSetSheet.Cells[index, 7].Value?.ToString() ?? "";
				if (name == "") continue;

				int found = int.Parse(existingSetSheet.Cells[index, 2]?.Value.ToString() ?? "0");

				// find if there is a part with the same name and if so set the found value of the part to the one from the spreadsheet
				if (found != 0)
				{
					foreach (var part in set.Parts)
					{
						if (name == part.Name)
						{
							part.AmountFound = found;
						}
					}
				}
			}
		}

		excel.Settings.ImageSettings.PrimaryImageHandler = new SystemDrawingImageHandler();

		var workbook = excel.Workbook;
		var setSheet = workbook.Worksheets.Add("Set");
		setSheet.DefaultColWidth = 16;
		setSheet.DefaultRowHeight = 48;
		setSheet.Columns[7].Width = 48;
		setSheet.View.FreezePanes(2, 1);

		var settingsSheet = workbook.Worksheets.Add("Settings");
		settingsSheet.DefaultColWidth = 16;
		settingsSheet.DefaultRowHeight = 48;
		settingsSheet.View.FreezePanes(2, 1);


		// Styles
		var normalStyle = workbook.Styles.NamedStyles[0];
		normalStyle.Style.Font.Name = "Arial";
		normalStyle.Style.Font.Color.SetColor(ColorTranslator.FromHtml("#000000"));
		normalStyle.Style.Font.Size = 14;
		normalStyle.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
		normalStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
		normalStyle.Style.ShrinkToFit = true;
		normalStyle.Style.WrapText = true;
		// border
		normalStyle.Style.Border.Top.Style = ExcelBorderStyle.Thin;
		normalStyle.Style.Border.Top.Color.SetColor(ColorTranslator.FromHtml("#c4c4c4"));
		normalStyle.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
		normalStyle.Style.Border.Bottom.Color.SetColor(ColorTranslator.FromHtml("#c4c4c4"));
		normalStyle.Style.Border.Left.Style = ExcelBorderStyle.Thin;
		normalStyle.Style.Border.Left.Color.SetColor(ColorTranslator.FromHtml("#c4c4c4"));
		normalStyle.Style.Border.Right.Style = ExcelBorderStyle.Thin;
		normalStyle.Style.Border.Right.Color.SetColor(ColorTranslator.FromHtml("#c4c4c4"));


		var headerStyle = workbook.Styles.CreateNamedStyle("header", normalStyle.Style);
		headerStyle.Style.Font.Bold = true;
		headerStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
		headerStyle.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#b6ddfd"));

		var doneStyle = workbook.Styles.CreateNamedStyle("done", normalStyle.Style);
		doneStyle.Style.Font.Bold = true;

		var foundStyle = workbook.Styles.CreateNamedStyle("found", normalStyle.Style);
		foundStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
		foundStyle.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#c7e1b9"));

		var infoStyle = workbook.Styles.CreateNamedStyle("info", normalStyle.Style);
		infoStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
		infoStyle.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#eeeeee"));

		var helperStyle = workbook.Styles.CreateNamedStyle("helper", normalStyle.Style);
		helperStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
		helperStyle.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#e5d994"));

		var totalRows = set.Parts.Length + 1;

		// apply styles
		setSheet.Cells[$"A1:I1"].StyleName = "header";
		setSheet.Cells[$"B2:B{totalRows}"].StyleName = "found";
		setSheet.Cells[$"C2:E{totalRows}"].StyleName = "info";
		setSheet.Cells[$"G2:G{totalRows}"].StyleName = "info";
		setSheet.Cells[$"H2:I{totalRows}"].StyleName = "helper";
		setSheet.Columns[8, 9].Hidden = true;

		// conditional formatting
		var zeroFormat = setSheet.ConditionalFormatting.AddExpression($"A2:A{totalRows}");
		zeroFormat.Formula = "=C2=D2";
		zeroFormat.Style.Fill.PatternType = ExcelFillStyle.Solid;
		zeroFormat.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#f13b29"));

		var incompleteFormat = setSheet.ConditionalFormatting.AddExpression($"A2:A{totalRows}");
		incompleteFormat.Formula = "=AND(C2>0,C2<D2)";
		incompleteFormat.Style.Fill.PatternType = ExcelFillStyle.Solid;
		incompleteFormat.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffe30c"));

		var completeFormat = setSheet.ConditionalFormatting.AddExpression($"A2:A{totalRows}");
		completeFormat.Formula = "=C2=0";
		completeFormat.Style.Fill.PatternType = ExcelFillStyle.Solid;
		completeFormat.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#55c44b"));

		var extraFormat = setSheet.ConditionalFormatting.AddExpression($"A2:A{totalRows}");
		extraFormat.Formula = "=C2<0";
		extraFormat.Style.Fill.PatternType = ExcelFillStyle.Solid;
		extraFormat.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffba3f"));

		setSheet.Cells[1, 1].Value = "Done";
		setSheet.Cells[1, 2].Value = "Amount Found";
		setSheet.Cells[1, 3].Value = "Amount Left";
		setSheet.Cells[1, 4].Value = "Amount Needed";
		setSheet.Cells[1, 5].Value = "BrickLink ID";
		setSheet.Cells[1, 6].Value = "Image";
		setSheet.Cells[1, 7].Value = "Name";
		setSheet.Cells[1, 8].Value = "Order";
		setSheet.Cells[1, 9].Value = "Complete";

		setSheet.Cells[$"E2:E{set.Parts.Length + 1}"].Style.Numberformat.Format = "@";

		for (int partIndex = 0; partIndex < set.Parts.Length; partIndex++)
		{
			var part = set.Parts[partIndex];
			int rowIndex = partIndex + 2;

			setSheet.Cells[rowIndex, 1].Formula = $"IF(C{rowIndex}=0,\"✔\",IF(C{rowIndex}<0,\" ^ \",IF(C{rowIndex}=D{rowIndex},\"✖\",\"—\")))";
			setSheet.Cells[rowIndex, 2].Value = part.AmountFound;
			setSheet.Cells[rowIndex, 3].Formula = $"D{rowIndex}-B{rowIndex}";
			setSheet.Cells[rowIndex, 4].Value = part.AmountNeeded;
			setSheet.Cells[rowIndex, 5].Value = part.BrickLinkId;

			// image
			if (!string.IsNullOrEmpty(part.ImgPath))
			{
				var imgPath = $"{directoryPath}../{part.ImgPath}";

				if (File.Exists(imgPath))
				{
					var imageFile = new FileInfo(imgPath);
					if (imageFile.Extension.ToLower() == ".jpg")
					{
						int dpi = 96;
						decimal rowHeight = (decimal)(setSheet.Rows[rowIndex].Height * dpi / 72);
						decimal columnWidth = (decimal)(setSheet.Columns[6].Width * 8 * dpi / 72);

						var image = setSheet.Drawings.AddPicture($"{partIndex}", imageFile);
						// decimal aspectRatio = (decimal)(image.Size.Width) / (decimal)(image.Size.Height);
						decimal aspectRatio = (decimal)image.Size.Width / (decimal)image.Size.Height;

						int imageHeightDifference = 1;
						int imageHeight = (int)(rowHeight - imageHeightDifference);
						int imageWidth = (int)(imageHeight * aspectRatio);

						image.ChangeCellAnchor(eEditAs.OneCell);
						image.SetSize(imageWidth, imageHeight);

						image.SetPosition(rowIndex - 1, imageHeightDifference, 5, (int)(columnWidth / 2 - imageWidth / 2));
					}
					else
					{
						Console.WriteLine($"Unsupported image format: {imageFile.Extension}");
					}
				}
				else
				{
					Console.WriteLine($"Image file does not exist: {imgPath}");
				}
			}

			setSheet.Cells[rowIndex, 7].Value = part.Name;
			setSheet.Cells[rowIndex, 8].Value = partIndex + 1;
			setSheet.Cells[rowIndex, 9].Formula = $"=IF(C{rowIndex}<=0, 1, 0)";
		}

		settingsSheet.Columns[1].Width = 20;
		settingsSheet.Rows[1].Height = 48;
		settingsSheet.Rows[2].Height = 55;
		settingsSheet.Rows[3, 6].Height = 48;

		settingsSheet.Cells["A1"].StyleName = "header";
		settingsSheet.Cells["A2:A5"].StyleName = "info";
		settingsSheet.Cells["A6"].StyleName = "found";

		settingsSheet.Cells["A1"].Value = "Sort Type";
		settingsSheet.Cells["A2"].Value = "1: Sort by Incomplete, then by Amount Left, then by Set Order";
		settingsSheet.Cells["A3"].Value = "2: Sort by Incomplete, then by Set Order";
		settingsSheet.Cells["A4"].Value = "3: Sort by Set Order";
		settingsSheet.Cells["A5"].Value = "4: Sort by Part Name";
		settingsSheet.Cells["A6"].Value = 2;

		settingsSheet.Rows[2].Height = 55;

		workbook.CreateVBAProject();

		var vbaFile = new StreamReader(directoryPath + "script.vba");

		workbook.CodeModule.Code = vbaFile.ReadToEnd();

		// excel.SaveAs($"../{set.Name}.xlsx");
		excel.SaveAs($"{directoryPath}../sets/{set.Name}.xlsm");
	}
}
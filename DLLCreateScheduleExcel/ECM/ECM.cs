using ClosedXML.Excel;
using DLLCreateScheduleExcel.Data;
using DLLCreateScheduleExcel.Entities;
using DLLCreateScheduleExcel.Services;
using GO.ECM.ScriptDotNet.Atributos;
using GO.ECM.ScriptDotNet.ProdutoEspecifico;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DLLCreateScheduleExcel.ECM
{
	[ClasseScript]
	public class ECM : ScriptEcm
	{
		#region Ctor

		public ECM()
		{
		}

		public ECM(object obj)
		{
		}

		#endregion Ctor

		#region Metodo

		[MetodoScript]
		public void GerarXls()
		{
			var documentId = base.RetornaContexto().Documento.DocumentoId;

			//Connection with DataBase
			var dataBase = new ConnectionDataBase();
			dataBase.Connect();
			string json = dataBase.Query("select resultado from GRID_DLL where id_versao = " + documentId);

			//Get the Object Array
			var gridObject = Welcome.FromJson(json.ToString());

			Console.WriteLine("Creating excel file...");

			var interval = new IntervalDates(DateTime.Parse(gridObject[0].DataInicio), DateTime.Parse(gridObject[0].DataFim));
			var intervalDates = new TimelineRange();
			List<string> daysList = intervalDates.TimelineDaysList(interval);
			List<string> monthsList = intervalDates.TimeLineMonthsList(interval);

			var db = new ConnectionDataBase();
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add("Cronograma");
			Console.WriteLine(ws.GetType());
			//Report Title
			ws.Cell("A1").Value = "Cronograma";
			ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
			ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
			ws.Cell("A1").Style.Font.Bold = true;

			ws.Range("A1:G3").Merge();

			//Report Header
			ws.Cell("A4").Value = "Item";
			ws.Cell("B4").Value = "Nome da Atividade";
			ws.Cell("C4").Value = "Responsável";
			ws.Cell("D4").Value = "Data Início";
			ws.Cell("E4").Value = "Data Fim";
			ws.Cell("F4").Value = "Dias Úteis";
			ws.Cell("G4").Value = "Realizado";

			//Report Grid Body
			int countIndex = 0;
			int linesLength = gridObject[0].Linhas.Count() + 5;
			for (int i = 5; i < linesLength; i += 2)
			{
				var backgroundColor = XLColor.PowderBlue;
				if (countIndex % 2 == 1) backgroundColor = XLColor.White;
				else backgroundColor = XLColor.FromHtml("#dde4ff");

				//insert data into table
				var startDate = gridObject[0].Linhas[i - 5].Tr[3].Td[0].Valor;
				var endDate = gridObject[0].Linhas[i - 5].Tr[4].Td[0].Valor;
				var accomplished = gridObject[0].Linhas[i - 5].Tr[6].Td[0].Valor;
				/*item*/
				ws.Cell($"A{i}").Value = gridObject[0].Linhas[i - 5].Tr[0].Td[0].Valor;
				ws.Range(i, 1, (i + 1), 1).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;
				/*Atividade*/
				ws.Cell($"B{i}").Value = gridObject[0].Linhas[i - 5].Tr[1].Td[0].Valor;
				ws.Range(i, 2, (i + 1), 2).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;
				/*Responsavel*/
				ws.Cell($"C{i}").Value = gridObject[0].Linhas[i - 5].Tr[2].Td[0].Valor;
				ws.Range(i, 3, (i + 1), 3).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;
				/*DataInicio*/
				ws.Cell($"D{i}").Value = startDate;
				ws.Range(i, 4, (i + 1), 4).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;
				/*DataFim*/
				ws.Cell($"E{i}").Value = endDate;
				ws.Range(i, 5, (i + 1), 5).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;
				/*Realizado*/
				ws.Cell($"G{i}").Value = accomplished;
				ws.Range(i, 7, (i + 1), 7).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;

				ws.Row(i).Height = 25;
				ws.Row(i + 1).Height = 10;
				countIndex++;
				//worksheet, timeline days list, expected stard, expected end, timeline row, background color
				/*color Timeline expected*/
				int workDays = intervalDates.colorTimeLine(ws, daysList, startDate, endDate, i, "#8ebbff");
				/*color timeline accomplished*/
				intervalDates.colorTimeLine(ws, daysList, startDate, accomplished, (i + 1), "#7eff77");

				/*D.Uteis*/
				ws.Cell($"F{i}").Value = workDays;
				ws.Range(i, 6, (i + 1), 6).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;
			}

			//Report Timeline Body
			int count = 8;
			int posMonth = 8;
			int indexMonthPrevious = 0;
			int indexYearPrevious = 0;
			int colSpanYear = 1;
			int posYear = 8;
			int posYearPrevious = 0;
			foreach (var d in daysList)
			{
				string[] month = new string[13] { "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" };

				var date = DateTime.Parse(d);
				int indexMonth = (int)date.Month;
				int year = (int)date.Year;

				if (indexMonth != indexMonthPrevious)
				{
					List<string> currentMonth = monthsList.FindAll(months =>
					{
						return months == month[indexMonth] + year;
					});
					int colSpan = currentMonth.Count();
					ws.Cell(2, posMonth).Value = month[indexMonth];
					int positionColSpan = posMonth + colSpan - 1;
					if (colSpan != 1) ws.Range(2, posMonth, 2, positionColSpan).Merge().Style.Fill.BackgroundColor = XLColor.FromHtml("#dee3f9");
					else ws.Cell(2, posMonth).Style.Fill.BackgroundColor = XLColor.FromHtml("#dee3f9");
					posMonth += colSpan;
					indexMonthPrevious = indexMonth;
				}

				if (year != indexYearPrevious)
				{
					posYearPrevious = posYear + 1;
					posYear = posYear != 8 ? (posYear + colSpanYear) : posYear;
					int positionColSpan = posYear + colSpanYear - 1;
					indexYearPrevious = year;

					if (colSpanYear != 1) ws.Range(1, posYearPrevious, 1, posYear - 1).Merge().Style.Fill.BackgroundColor = XLColor.FromHtml("#cfd8fc");
					else ws.Cell(1, posYear).Style.Fill.BackgroundColor = XLColor.FromHtml("#cfd8fc");

					ws.Cell(1, posYearPrevious).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
					ws.Cell(1, posYear).Value = indexYearPrevious;
					posYear--;
				}
				colSpanYear++;
				string shortDay = d.Substring(0, 5);
				ws.Cell(3, count).Value = shortDay;
				ws.Cell(3, count).Style.Fill.BackgroundColor = XLColor.FromHtml("#e8ebf7");
				ws.Cell(3, count).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				ws.Cell(4, count).Style.Fill.BackgroundColor = XLColor.FromHtml("#f4f5f7");

				var dayWeek = date.DayOfWeek;
				CultureInfo brasil = new CultureInfo("pt-BR");
				string week = brasil.DateTimeFormat.DayNames[(int)dayWeek];

				ws.Cell(4, count).Value = week.ToString().ToUpper().Substring(0, 1);

				if (dayWeek.ToString() == "Saturday" || dayWeek.ToString() == "Sunday")
				{
					ws.Range(5, count, linesLength, count).Style.Fill.BackgroundColor = XLColor.FromHtml("#d1d1d1");
				}

				count++;
			}

			//Set merge the last year e put background it
			if (colSpanYear != 1) ws.Range(1, posYear + 1, 1, colSpanYear + 6).Merge().Style.Fill.BackgroundColor = XLColor.FromHtml("#e8ecff");
			else ws.Cell(1, posYear + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#e8ecff");
			ws.Cell(1, posYear + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

			//Filters and Create table-------------------------------------------------------------
			var range = ws.Range("A4:G" + (linesLength - 1));
			range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
			ws.Range("A4:G" + (linesLength)).Style.Border.RightBorder = XLBorderStyleValues.Thin;
			ws.Range("A4:G" + (linesLength)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
			range.Style.Border.OutsideBorderColor = XLColor.DimGray;
			range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
			range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
			// range.CreateTable();
			ws.Range("A4:G4").Style.Fill.BackgroundColor = XLColor.AliceBlue;

			//Fix the column size with column content
			var item = ws.ColumnsUsed();

			//Other table adjustments
			ws.Columns("4-7").Width = 11; //Define size Dates Grid Columns
			ws.Columns("8-" + count).Width = 7; //Define size Dates Schedule Columns
			ws.Rows(5, linesLength - 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
			ws.Rows(5, linesLength - 1).Style.Border.BottomBorderColor = XLColor.FromHtml("#d3d3d3");
			ws.Column(2).Width = 40; // Define size Column "Nome da Atividade"
			ws.Column(3).Width = 35; // Define size Column "Responsavel"
			ws.Column(1).Width = 5;//Define size Column "Item"
			ws.Row(4).Cells("1:" + (count - 1)).Style.Border.BottomBorder = XLBorderStyleValues.Thin; // set the border of header bottom
			ws.Row(4).Cells("1:" + (count - 1)).Style.Border.BottomBorderColor = XLColor.FromHtml("#bcbcbc"); // set header border color
			ws.Range(1, 1, linesLength, count - 1).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;// set the table OutsideBorder
			ws.Columns("A, D:G").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;//leaves selected cells centered
			ws.Column("B").Style.Alignment.WrapText = true;//Set wrap at  Column "Nome da Atividade"
			ws.Range(4, 8, 4, (count - 1)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; //set cells of the week centered
			ws.Range(3, 8, 3, (count - 1)).Style.NumberFormat.Format = "dd";//format as showed the dates
			ws.Range(3, 8, 3, (count - 1)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; //set cells of the date centered
			ws.SheetView.FreezeColumns(7);//freeze grid columns and let time line free

			//Salve file

			wb.SaveAs(@"D:\PROJETOS\TGESTIONA\SITE\Arquivos\test2.xlsx");

			//Release objects

			wb.Dispose();

			Console.WriteLine("Finish");
			Console.ReadKey();
		}

		#endregion Metodo
	}
}
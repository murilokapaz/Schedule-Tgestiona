using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DLLCreateScheduleExcel.Data;
using DLLCreateScheduleExcel.Entities;
using DLLCreateScheduleExcel.Services;
using Newtonsoft.Json;

namespace DLLCreateScheduleExcel
{
    class Program
    {
        static void Main(string[] args)
        {


            //Connection with DataBase
            var dataBase = new ConnectionDataBase();
            dataBase.Connect();
            string json = dataBase.Query("select resultado from GRID_DLL where id_versao = 10565;");



            //var jsonString = "[{\"nomeTabela\":\"cronograma\",\"colunas\":[{\"id\":0,\"valor\":\"item\",\"elemento\":\"input\",\"type\":\"number\"},{\"id\":1,\"valor\":\"Atividade\",\"elemento\":\"input\",\"type\":\"text\"},{\"id\":2,\"valor\":\"Responsavel\",\"elemento\":\"input\",\"type\":\"text\"},{\"id\":3,\"valor\":\"Início\",\"elemento\":\"input\",\"type\":\"date\"},{\"id\":4,\"valor\":\"Fim\",\"elemento\":\"input\",\"type\":\"date\"},{\"id\":5,\"valor\":\"D. Utéis\",\"elemento\":\"input\",\"type\":\"number\"},{\"id\":6,\"valor\":\"Real.\",\"elemento\":\"input\",\"type\":\"number\"},{\"id\":7,\"valor\":\"Ações\"}],\"linhas\":[{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"2\"}]},{\"td\":[{\"id\":1,\"valor\":\"Mesa de Apresentação Baseline \"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"3\"}]},{\"td\":[{\"id\":1,\"valor\":\"Publicação conforme 155 anatel\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"4\"}]},{\"td\":[{\"id\":1,\"valor\":\"Envio das Condições de Participação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"5\"}]},{\"td\":[{\"id\":1,\"valor\":\"Entrega dos documentos para Cadastro\"}]},{\"td\":[{\"id\":2,\"valor\":\"Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"6\"}]},{\"td\":[{\"id\":1,\"valor\":\"Análise e Comunicação da pré-qualificação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Juridico / Risco / Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"7\"}]},{\"td\":[{\"id\":1,\"valor\":\"Entrega documentação complementar para cadastro\"}]},{\"td\":[{\"id\":2,\"valor\":\"Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"8\"}]},{\"td\":[{\"id\":1,\"valor\":\"Analise da Documentação Complementar e definição da participação das Pretentendes\"}]},{\"td\":[{\"id\":2,\"valor\":\"Juridico / Risco / Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"9\"}]},{\"td\":[{\"id\":1,\"valor\":\"Definição das Empresas Participantes\"}]},{\"td\":[{\"id\":2,\"valor\":\"Alta Direção Telefonica\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"10\"}]},{\"td\":[{\"id\":1,\"valor\":\"Envio e recebimento do termo de Confidencialidade\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"11\"}]},{\"td\":[{\"id\":1,\"valor\":\"Envio da Solicitação de Propostas (RFP)\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"12\"}]},{\"td\":[{\"id\":1,\"valor\":\"Elaboração da Proposta\"}]},{\"td\":[{\"id\":2,\"valor\":\"Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"13\"}]},{\"td\":[{\"id\":1,\"valor\":\"Recebimento das Propostas\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"14\"}]},{\"td\":[{\"id\":1,\"valor\":\"Análise / Alinhamento / RV / Emissão de Laudo/Ditame\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras / Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"15\"}]},{\"td\":[{\"id\":1,\"valor\":\"Mesa preparatória\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"16\"}]},{\"td\":[{\"id\":1,\"valor\":\"Treinamento / Vista prévia / realização da Subasta\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"17\"}]},{\"td\":[{\"id\":1,\"valor\":\"Mesa Adjudicação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"18\"}]},{\"td\":[{\"id\":1,\"valor\":\"Comunicação de Adjudicação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras/Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"19\"}]},{\"td\":[{\"id\":1,\"valor\":\"Periodo de transição operacional\"}]},{\"td\":[{\"id\":2,\"valor\":\"Contratada/Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]}]}]";
            
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
            int linesLength = gridObject[0].Linhas.Count()+5;
            for(int i=5; i< linesLength; i+=2)
            {
                var backgroundColor= XLColor.PowderBlue;
                if (countIndex  % 2 == 1) backgroundColor = XLColor.White;
                else backgroundColor = XLColor.FromHtml("#dde4ff");

                //insert data into table

                /*item*/
                ws.Cell($"A{i}").Value = gridObject[0].Linhas[i-5].Tr[0].Td[0].Valor;
                               ws.Range(i, 1, (i+1), 1).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor;
                /*Atividade*/  ws.Cell($"B{i}").Value = gridObject[0].Linhas[i-5].Tr[1].Td[0].Valor;
                               ws.Range(i, 2, (i + 1), 2).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor; ;                            
                /*Responsavel*/ws.Cell($"C{i}").Value = gridObject[0].Linhas[i-5].Tr[2].Td[0].Valor;
                               ws.Range(i, 3, (i + 1), 3).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor; ;
                /*DataInicio*/ ws.Cell($"D{i}").Value = gridObject[0].Linhas[i-5].Tr[3].Td[0].Valor;
                               ws.Range(i, 4, (i + 1), 4).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor; ;
                /*DataFim*/    ws.Cell($"E{i}").Value = gridObject[0].Linhas[i-5].Tr[4].Td[0].Valor;
                               ws.Range(i, 5, (i + 1), 5).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor; ;
                /*D.Uteis*/    ws.Cell($"F{i}").Value = gridObject[0].Linhas[i-5].Tr[5].Td[0].Valor;
                               ws.Range(i, 6, (i + 1), 6).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor; ;
                /*Realizado*/  ws.Cell($"G{i}").Value = gridObject[0].Linhas[i-5].Tr[6].Td[0].Valor;
                               ws.Range(i, 7, (i + 1), 7).Column(1).Merge().Style.Fill.BackgroundColor = backgroundColor; ;

                ws.Row(i).Height = 25;
                ws.Row(i+1).Height = 10;
                countIndex++;

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
                int indexMonth = (int) date.Month;
                int year = (int)date.Year;
                

                if(indexMonth != indexMonthPrevious)
                {
                    List<string> currentMonth = monthsList.FindAll(months => { 
                        return months == month[indexMonth] + year;
                    });
                    int colSpan = currentMonth.Count();
                    ws.Cell(2, posMonth).Value = month[indexMonth];
                    int positionColSpan = posMonth+ colSpan-1;
                    if(colSpan != 1) ws.Range(2, posMonth, 2, positionColSpan).Merge().Style.Fill.BackgroundColor= XLColor.FromHtml("#dee3f9");
                    else ws.Cell(2, posMonth).Style.Fill.BackgroundColor = XLColor.FromHtml("#dee3f9");
                    posMonth += colSpan;
                    indexMonthPrevious = indexMonth;
                }

                if(year != indexYearPrevious)
                {
                    posYearPrevious = posYear+1;
                    posYear = posYear!= 8? (posYear + colSpanYear) : posYear;
                    int positionColSpan = posYear + colSpanYear-1;
                    indexYearPrevious = year;

                    if (colSpanYear != 1) ws.Range(1, posYearPrevious, 1, posYear-1).Merge().Style.Fill.BackgroundColor = XLColor.FromHtml("#cfd8fc");
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

                ws.Cell(4, count).Value = dayWeek;

                if (dayWeek.ToString() == "Saturday"|| dayWeek.ToString() == "Sunday")
                {
                    ws.Range(5, count, linesLength, count).Style.Fill.BackgroundColor = XLColor.FromHtml("#d1d1d1");
                }

                count++;
            }
            //Set merge the last year e put background it
            if (colSpanYear != 1) ws.Range(1, posYear+1, 1, colSpanYear+6).Merge().Style.Fill.BackgroundColor = XLColor.FromHtml("#e8ecff");
            else ws.Cell(1, posYear+1).Style.Fill.BackgroundColor = XLColor.FromHtml("#e8ecff");
            ws.Cell(1, posYear+1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;


            //Filters and Create table
            var range = ws.Range("A4:G"+(linesLength-1));
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


            ws.Columns("1-" + count).Width = 11; //Define size Dates Columns
            ws.Column(2).Width = 40; // Define size Column "Nome da Atividade"
            ws.Column(3).Width = 35; // Define size Column "Responsavel"
            ws.Column(1).Width = 5;//Define size Column "Item"
            ws.Row(4).Cells("1:"+(count-1)).Style.Border.BottomBorder = XLBorderStyleValues.Thin; // set the border of header bottom
            ws.Row(4).Cells("1:" + (count - 1)).Style.Border.BottomBorderColor = XLColor.FromHtml("#bcbcbc"); // set header border color
            ws.Range(1, 1, linesLength, count - 1).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;// set the table OutsideBorder 
            ws.Columns("A, D:G").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;//leaves selected cells centered
            ws.Column("B").Style.Alignment.WrapText = true;//Set wrap at  Column "Nome da Atividade"
            ws.Range(4, 8, 4, (count - 1)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.SheetView.FreezeColumns(7);
            
            //Salve file
            wb.SaveAs(@"C:\Users\murilo.paz.REDESPC\source\repos\ExcelSchedule\test2.xlsx");




            //Release objects

            wb.Dispose();          
            


            Console.WriteLine("Finish");
            Console.ReadKey();
            
            
        }


    }
}

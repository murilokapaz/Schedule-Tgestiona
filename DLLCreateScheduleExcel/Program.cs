using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DLLCreateScheduleExcel.Data;
using DLLCreateScheduleExcel.Entities;
using DLLCreateScheduleExcel.Services;

namespace DLLCreateScheduleExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Connection with DataBase
            var jsonString = "[{\"nomeTabela\":\"cronograma\",\"colunas\":[{\"id\":0,\"valor\":\"item\",\"elemento\":\"input\",\"type\":\"number\"},{\"id\":1,\"valor\":\"Atividade\",\"elemento\":\"input\",\"type\":\"text\"},{\"id\":2,\"valor\":\"Responsavel\",\"elemento\":\"input\",\"type\":\"text\"},{\"id\":3,\"valor\":\"Início\",\"elemento\":\"input\",\"type\":\"date\"},{\"id\":4,\"valor\":\"Fim\",\"elemento\":\"input\",\"type\":\"date\"},{\"id\":5,\"valor\":\"D. Utéis\",\"elemento\":\"input\",\"type\":\"number\"},{\"id\":6,\"valor\":\"Real.\",\"elemento\":\"input\",\"type\":\"number\"},{\"id\":7,\"valor\":\"Ações\"}],\"linhas\":[{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"2\"}]},{\"td\":[{\"id\":1,\"valor\":\"Mesa de Apresentação Baseline \"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"3\"}]},{\"td\":[{\"id\":1,\"valor\":\"Publicação conforme 155 anatel\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"4\"}]},{\"td\":[{\"id\":1,\"valor\":\"Envio das Condições de Participação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"5\"}]},{\"td\":[{\"id\":1,\"valor\":\"Entrega dos documentos para Cadastro\"}]},{\"td\":[{\"id\":2,\"valor\":\"Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"6\"}]},{\"td\":[{\"id\":1,\"valor\":\"Análise e Comunicação da pré-qualificação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Juridico / Risco / Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"7\"}]},{\"td\":[{\"id\":1,\"valor\":\"Entrega documentação complementar para cadastro\"}]},{\"td\":[{\"id\":2,\"valor\":\"Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"8\"}]},{\"td\":[{\"id\":1,\"valor\":\"Analise da Documentação Complementar e definição da participação das Pretentendes\"}]},{\"td\":[{\"id\":2,\"valor\":\"Juridico / Risco / Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"9\"}]},{\"td\":[{\"id\":1,\"valor\":\"Definição das Empresas Participantes\"}]},{\"td\":[{\"id\":2,\"valor\":\"Alta Direção Telefonica\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"10\"}]},{\"td\":[{\"id\":1,\"valor\":\"Envio e recebimento do termo de Confidencialidade\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"11\"}]},{\"td\":[{\"id\":1,\"valor\":\"Envio da Solicitação de Propostas (RFP)\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"12\"}]},{\"td\":[{\"id\":1,\"valor\":\"Elaboração da Proposta\"}]},{\"td\":[{\"id\":2,\"valor\":\"Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"13\"}]},{\"td\":[{\"id\":1,\"valor\":\"Recebimento das Propostas\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras Proponente\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"14\"}]},{\"td\":[{\"id\":1,\"valor\":\"Análise / Alinhamento / RV / Emissão de Laudo/Ditame\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras / Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"15\"}]},{\"td\":[{\"id\":1,\"valor\":\"Mesa preparatória\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"16\"}]},{\"td\":[{\"id\":1,\"valor\":\"Treinamento / Vista prévia / realização da Subasta\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"17\"}]},{\"td\":[{\"id\":1,\"valor\":\"Mesa Adjudicação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"18\"}]},{\"td\":[{\"id\":1,\"valor\":\"Comunicação de Adjudicação\"}]},{\"td\":[{\"id\":2,\"valor\":\"Compras/Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]},{\"tr\":[{\"td\":[{\"id\":0,\"valor\":\"19\"}]},{\"td\":[{\"id\":1,\"valor\":\"Periodo de transição operacional\"}]},{\"td\":[{\"id\":2,\"valor\":\"Contratada/Gestor\"}]},{\"td\":[{\"id\":3,\"valor\":\"\"}]},{\"td\":[{\"id\":4,\"valor\":\"\"}]},{\"td\":[{\"id\":5,\"valor\":\"\"}]},{\"td\":[{\"id\":6,\"valor\":\"\"}]},{\"td\":[{\"id\":7,\"valor\":\"\"}]}]}]}]";
            
            //Get the Object Array
            var gridObject = Welcome.FromJson(jsonString.ToString());


            
            Console.WriteLine("Creating excel file...");


            var interval = new IntervalDates(DateTime.Parse("31/01/2020"), DateTime.Parse("01/04/2020"));
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

            int linesLength = gridObject[0].Linhas.Count()+5;
            for(int i=5; i< linesLength; i++)
            {
                //insert data into table

                /*item*/       ws.Cell($"A{i}").Value = gridObject[0].Linhas[i-5].Tr[0].Td[0].Valor;
                /*Atividade*/  ws.Cell($"B{i}").Value = gridObject[0].Linhas[i-5].Tr[1].Td[0].Valor;
                /*Responsavel*/ws.Cell($"C{i}").Value = gridObject[0].Linhas[i-5].Tr[2].Td[0].Valor;
                /*DataInicio*/ ws.Cell($"D{i}").Value = gridObject[0].Linhas[i-5].Tr[3].Td[0].Valor;
                /*DataFim*/    ws.Cell($"E{i}").Value = gridObject[0].Linhas[i-5].Tr[4].Td[0].Valor;
                /*D.Uteis*/    ws.Cell($"F{i}").Value = gridObject[0].Linhas[i-5].Tr[5].Td[0].Valor;
                /*Realizado*/  ws.Cell($"G{i}").Value = gridObject[0].Linhas[i-5].Tr[6].Td[0].Valor;
            }

            //Report Timeline Body 
            int count = 8;
            int posMonth = 8;
            int indexMonthPrevious = 0;
            int indexYearPrevious = 0;
            int colSpanYear = 1;
            int posYear = 8;
            foreach (var d in daysList)
            {
                string[] month = new string[13] { "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" };

                var date = DateTime.Parse(d);
                int indexMonth = (int) date.Month;
                int year = (int)date.Year;
                

                if(indexMonth != indexMonthPrevious)
                {
                    List<string> currentMonth = monthsList.FindAll(months => months == month[indexMonth]);
                    int colSpan = currentMonth.Count();
                    ws.Cell(3, posMonth).Value = month[indexMonth];
                    int positionColSpan = posMonth+ colSpan;
                    ws.Range(3, posMonth, 3, positionColSpan).Merge();
                    posMonth += colSpan;
                    indexMonthPrevious = indexMonth;
                }

                if(year != indexYearPrevious)
                {
                    //int positionColSpan = posYear + colSpanYear;
                    //ws.Cell(1, posYear, 1, positionColSpan);
                }

                ws.Cell(4, count).Value = d;           
                count++;
            }


            //Filters and Create table
            var range = ws.Range("A4:G"+(linesLength-1));
            range.CreateTable();
             
            //Fix the column size with column content 
            ws.Columns("1-"+count).AdjustToContents();
            ws.SheetView.FreezeColumns(7);
            
            //Salve file
            wb.SaveAs(@"C:\Users\muril\source\repos\CreateExcelSchedule-master\test.xlsx");

            //Release objects
           
            wb.Dispose();          
            


            Console.WriteLine("Finish");
            Console.ReadKey();

            
        }


    }
}

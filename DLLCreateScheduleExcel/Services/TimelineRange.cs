using ClosedXML.Excel;
using DLLCreateScheduleExcel.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace DLLCreateScheduleExcel.Services
{
    internal class TimelineRange
    {
        public int ColSpan { get; set; }

        public List<string> TimelineDaysList(IntervalDates interval)
        {
            var startDay = interval.StartSchedule;
            TimeSpan diffDays = interval.EndSchedule.Subtract(startDay);
            int numberDays = (int)diffDays.TotalDays;

            List<string> listOfDays = new List<string>();

            var day = startDay;
            listOfDays.Add(day.ToShortDateString());
            for (int i = 0; i < numberDays; i++)
            {
                day = day.AddDays(1);
                listOfDays.Add(day.ToShortDateString());
            }
            return listOfDays;
        }

        public List<string> TimeLineMonthsList(IntervalDates interval)
        {
            string[] month = new string[13] { "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" };

            var startDate = interval.StartSchedule;
            TimeSpan diffDays = interval.EndSchedule.Subtract(startDate);
            int numberDays = (int)diffDays.TotalDays;

            List<string> listOfMonths = new List<string>();

            var day = startDate;
            listOfMonths.Add(month[day.Month] + day.Year);
            this.ColSpan = 0;
            for (int i = 0; i < numberDays; i++)
            {
                day = day.AddDays(1);
                listOfMonths.Add(month[day.Month] + day.Year);
            }

            return listOfMonths;
        }

        public int colorTimeLine(IXLWorksheet ws, List<string> daysList, string startDate, string endDate, int positionLineDate, string background)
        {
            DateTime startDateFormat;
            DateTime endDateFormat;
            try
            {
                startDateFormat = DateTime.ParseExact(startDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                endDateFormat = DateTime.ParseExact(endDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }

            var interval = new IntervalDates(startDateFormat, endDateFormat);

            List<string> datesToColorList = TimelineDaysList(interval);

            var XLbackground = XLColor.FromHtml(background);
            int column = 8;
            int workDays = 0;
            foreach (var dateToColor in datesToColorList)
            {
                foreach (var date in daysList)
                {
                    if (dateToColor == date)
                    {
                        var week = DateTime.Parse(date).DayOfWeek.ToString();
                        if (week != "Sunday" && week != "Saturday")
                        {
                            ws.Cell(positionLineDate, column).Style.Fill.BackgroundColor = XLbackground;
                            column++;
                            workDays++;
                            break;
                        }
                        else
                        {
                            column++;
                            break;
                        }
                    }
                }
            }
            return workDays;
        }
    }
}
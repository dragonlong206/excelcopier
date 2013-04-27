using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MicrosoftExcelCopier
{
    public static class DatetimeUtils
    {
        /// <summary>
        /// Check a day is last day of month
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static bool IsLastDayOfMonth(this DateTime date)
        {
            var currentMonth = date.Month;
            var monthOfDayPlusOne = (date.AddDays(1)).Month;

            // If add 1 day and month not change then it isn't the last day of month
            if (currentMonth.Equals(monthOfDayPlusOne))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Get first day of month
        /// </summary>
        /// <param name="date">Any day in month</param>
        /// <returns></returns>
        public static DateTime GetFirstDayOfMonth(this DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        /// <summary>
        /// Get last day of month
        /// </summary>
        /// <param name="date">Any day in month</param>
        /// <returns></returns>
        public static DateTime GetLastDayOfMonth(this DateTime date)
        {
            // Get first day of next month
            DateTime temp = date.AddMonths(1).GetFirstDayOfMonth();

            // then subtract 1 day
            DateTime lastDate = temp.AddDays(-1);
            return lastDate;
        }
    }
}

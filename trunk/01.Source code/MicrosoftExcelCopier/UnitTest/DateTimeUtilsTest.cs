using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using MicrosoftExcelCopier;

namespace UnitTest
{
    [TestFixture]
    public class DateTimeUtilsTest_IsLastDayOfMonth
    {
        #region IsLastDayOfMonth
        [Test]
        public void IsLastDayOfMonth31Days()
        {
            DateTime date = new DateTime(2013, 1, 31);
            Assert.AreEqual(true, date.IsLastDayOfMonth());
        }

        [Test]
        public void IsLastDayOfMonth30Days()
        {
            DateTime date = new DateTime(2013, 4, 30);
            Assert.AreEqual(true, date.IsLastDayOfMonth());
        }

        [Test]
        public void IsLastDayOfMonth28Days()
        {
            DateTime date = new DateTime(2013, 2, 28);
            Assert.AreEqual(true, date.IsLastDayOfMonth());
        }

        [Test]
        public void IsLastDayOfMonth29Days()
        {
            DateTime date = new DateTime(2012, 2, 29);
            Assert.AreEqual(true, date.IsLastDayOfMonth());
        }

        [Test]
        public void IsNotLastDayOfMonth()
        {
            DateTime date = new DateTime(2013, 2, 15);
            Assert.AreEqual(false, date.IsLastDayOfMonth());
        }
        #endregion
    }

    [TestFixture]
    public class DateTimeUtilsTest_GetFirstDayOfMonth
    {
        #region GetFirstDayOfMonth
        [Test]
        public void GetFirstDayOfMonth_FirstDay()
        {
            DateTime date = new DateTime(2013, 4, 1);
            DateTime firstDay = date.GetFirstDayOfMonth();
            Assert.AreEqual(2013, firstDay.Year);
            Assert.AreEqual(4, firstDay.Month);
            Assert.AreEqual(1, firstDay.Day);
        }

        [Test]
        public void GetFirstDayOfMonth_MiddleDay()
        {
            DateTime date = new DateTime(2013, 4, 12);
            DateTime firstDay = date.GetFirstDayOfMonth();
            Assert.AreEqual(2013, firstDay.Year);
            Assert.AreEqual(4, firstDay.Month);
            Assert.AreEqual(1, firstDay.Day);
        }

        [Test]
        public void GetFirstDayOfMonth_LastDay()
        {
            DateTime date = new DateTime(2013, 4, 30);
            DateTime firstDay = date.GetFirstDayOfMonth();
            Assert.AreEqual(2013, firstDay.Year);
            Assert.AreEqual(4, firstDay.Month);
            Assert.AreEqual(1, firstDay.Day);
        }
        #endregion
    }

    [TestFixture]
    public class DateTimeUtilsTest_GetLastDayOfMonth
    {
        #region GetLastDayOfMonth
        [Test]
        public void GetLastDayOfMonth_FirstDay_30DayMonth()
        {
            DateTime date = new DateTime(2013, 4, 1);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(4, lastDay.Month);
            Assert.AreEqual(30, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_MiddleDay_30DayMonth()
        {
            DateTime date = new DateTime(2013, 4, 12);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(4, lastDay.Month);
            Assert.AreEqual(30, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_LastDay_30DayMonth()
        {
            DateTime date = new DateTime(2013, 4, 30);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(4, lastDay.Month);
            Assert.AreEqual(30, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_FirstDay_31DayMonth()
        {
            DateTime date = new DateTime(2013, 1, 1);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(1, lastDay.Month);
            Assert.AreEqual(31, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_MiddleDay_31DayMonth()
        {
            DateTime date = new DateTime(2013, 1, 12);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(1, lastDay.Month);
            Assert.AreEqual(31, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_LastDay_31DayMonth()
        {
            DateTime date = new DateTime(2013, 1, 31);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(1, lastDay.Month);
            Assert.AreEqual(31, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_FirstDay_28DayMonth()
        {
            DateTime date = new DateTime(2013, 2, 1);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(2, lastDay.Month);
            Assert.AreEqual(28, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_MiddleDay_28DayMonth()
        {
            DateTime date = new DateTime(2013, 2, 12);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(2, lastDay.Month);
            Assert.AreEqual(28, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_LastDay_28DayMonth()
        {
            DateTime date = new DateTime(2013, 2, 28);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2013, lastDay.Year);
            Assert.AreEqual(2, lastDay.Month);
            Assert.AreEqual(28, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_FirstDay_29DayMonth()
        {
            DateTime date = new DateTime(2012, 2, 1);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2012, lastDay.Year);
            Assert.AreEqual(2, lastDay.Month);
            Assert.AreEqual(29, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_MiddleDay_29DayMonth()
        {
            DateTime date = new DateTime(2012, 2, 12);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2012, lastDay.Year);
            Assert.AreEqual(2, lastDay.Month);
            Assert.AreEqual(29, lastDay.Day);
        }

        [Test]
        public void GetLastDayOfMonth_LastDay_29DayMonth()
        {
            DateTime date = new DateTime(2012, 2, 29);
            DateTime lastDay = date.GetLastDayOfMonth();
            Assert.AreEqual(2012, lastDay.Year);
            Assert.AreEqual(2, lastDay.Month);
            Assert.AreEqual(29, lastDay.Day);
        }
        #endregion
    }
}

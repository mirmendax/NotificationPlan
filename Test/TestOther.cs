
using System;
using NUnit.Framework;
using NotificationPlan.Data;
using NotificationPlan.Models;

namespace Test
{
    public class TestOther
    {
        string[] foo = new string[]
        {
            string.Empty, "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь",
            "Декабрь"
        };
        [Test]
        public void TestGetMonthToString()
        {
            int i = 0;
            foreach (var item in foo)
            {
                Assert.AreEqual(Other.GetMonthToString(i), item);
                i++;
            }
        }

        [Test]
        public void TestGetMonthToInt()
        {
            
            for (int i = 0; i < 13; i++)
            {
                Assert.AreEqual(Other.GetMonthToInt(foo[i]), i);
            }
            
        }
    }
}
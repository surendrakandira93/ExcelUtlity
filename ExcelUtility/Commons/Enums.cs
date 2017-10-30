using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUtility
{
    public static partial class Enums
    {
        public enum PeriodType
        {
            /// <summary>
            /// The new
            /// </summary>
            New = 0,

            /// <summary>
            /// The april
            /// </summary>
            April = 1,

            /// <summary>
            /// The may
            /// </summary>
            May = 2,

            /// <summary>
            /// The june
            /// </summary>
            June = 3,

            /// <summary>
            /// The july
            /// </summary>
            July = 4,

            /// <summary>
            /// The august
            /// </summary>
            August = 5,

            /// <summary>
            /// The september
            /// </summary>
            September = 6,

            /// <summary>
            /// The october
            /// </summary>
            October = 7,

            /// <summary>
            /// The november
            /// </summary>
            November = 8,

            /// <summary>
            /// The december
            /// </summary>
            December = 9,

            /// <summary>
            /// The january
            /// </summary>
            January = 10,

            /// <summary>
            /// The february
            /// </summary>
            February = 11,

            /// <summary>
            /// The march
            /// </summary>
            March = 12,

            /// <summary>
            /// The q1
            /// </summary>
            Q1 = 13,

            /// <summary>
            /// The q2
            /// </summary>
            Q2 = 14,

            /// <summary>
            /// The q3
            /// </summary>
            Q3 = 15,

            /// <summary>
            /// The q4
            /// </summary>
            Q4 = 16
        }

        public enum FinancialYearText
        {
            /// <summary>
            /// The new
            /// </summary>
            New = 0,

            /// <summary>
            /// The f y201415
            /// </summary>
            FY201415 = 1,

            /// <summary>
            /// The f y201516
            /// </summary>
            FY201516 = 2,

            /// <summary>
            /// The f y201617
            /// </summary>
            FY201617 = 3,

            /// <summary>
            /// The f y201718
            /// </summary>
            FY201718 = 4,

            /// <summary>
            /// The f y201819
            /// </summary>
            FY201819 = 5,

            /// <summary>
            /// The f y201920
            /// </summary>
            FY201920 = 6,

            /// <summary>
            /// The f y202021
            /// </summary>
            FY202021 = 7
        }

        public enum CellType
        {
            Empty,
            Invalid,
            Exist,
            Duplicate,
            NotMatched
        }
    }
}
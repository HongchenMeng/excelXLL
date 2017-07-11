﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace excelXLL
{
    class Common
    {
        public static Microsoft.Office.Interop.Excel.Application xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        
        public static object RunMacro(object oApp, object[] oRunArgs)
        {
            return oApp.GetType().InvokeMember("Run",
                 System.Reflection.BindingFlags.Default |
                 System.Reflection.BindingFlags.InvokeMethod,
                 null, oApp, oRunArgs);
        }

        public static bool IsMissOrEmpty(object srcPara)
        {
            if (srcPara is ExcelMissing || string.IsNullOrEmpty(srcPara.ToString().Trim()))
            {
                return true;

            }
            else
            {
                return false;
            }
        }

        public static double TransNumberPara(object srcPara)
        {
            double number;
            NumberStyles style = NumberStyles.Any;
            CultureInfo culture = CultureInfo.CurrentCulture;
            if (srcPara is bool && (bool)srcPara == true)
            {
                return 1;
            }
            else if (double.TryParse(srcPara.ToString(), style, culture, out number))
            {
                return number;
            }
            else
            {
                return 0;
            }
        }


        public static bool TransBoolPara(object srcPara)
        {
            double number;
            NumberStyles style = NumberStyles.Any;
            CultureInfo culture = CultureInfo.CurrentCulture;

            if (srcPara is bool && (bool)srcPara == true)
            {
                return true;
            }
            else if (double.TryParse(srcPara.ToString(), style, culture, out number))
            {
                return true;
            }
            else if (srcPara is string && string.IsNullOrEmpty((string)srcPara) != true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static object ReturnDataArray(object[] srcArrData, string optAlignHorL)
        {

            int resultCount = srcArrData.Count();

            if (Common.IsMissOrEmpty(optAlignHorL) || optAlignHorL.Equals("H", StringComparison.CurrentCultureIgnoreCase) == false)
            {
                optAlignHorL = "L";
            }
            else
            {
                optAlignHorL = "H";
            }
            //直接用从下标为0开始的数组也可以
            if (optAlignHorL == "L")
            {
                object[,] resultArr = new object[resultCount, 1];
                for (int i = 0; i < resultCount; i++)
                {
                    resultArr[i, 0] = srcArrData[i];
                }
                return resultArr;
            }

            else
            {
                //int[] myLengthsArray = new int[2] {1, resultCount };
                //int[] myBoundsArray = new int[2] { 1, 1 };
                //Array resultArr = Array.CreateInstance(typeof(object), myLengthsArray, myBoundsArray);
                //for (int i = resultArr.GetLowerBound(0); i <= resultArr.GetUpperBound(0); i++)
                //    for (int j = resultArr.GetLowerBound(1); j <= resultArr.GetUpperBound(1); j++)
                //    {
                //        int[] myIndicesArray = new int[2] { i, j };
                //        resultArr.SetValue(subfolders[ilist], myIndicesArray);
                //        ilist++;
                //    }
                //横排时，直接用一维数组就可以识别到
                return srcArrData;
                //object[,] resultArr = new object[1, resultCount];
                //for (int i = 0; i < resultCount; i++)
                //{


                //}
                //return resultArr;
            }

        }
    }
}

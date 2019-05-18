using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace StockApplication
{
    public partial class StockRibbon
    {
        /// <summary>
        /// Variables
        /// Ideally All the Stock Info should Load from an API. 
        /// For the sample Assignment data purpose, added a Collection
        /// </summary>
        #region
        static Dictionary<string, double> StockDictionary = new Dictionary<string, double>()
        {
            {"GOOG", 256.85 },
            {"AAPL", 556.25 },
            {"GE", 206.83 },
            {"TSLA", 226.85 }
        };
        static Excel.Worksheet CurrentWorkSheet = null;
        #endregion

        /// <summary>
        /// Events
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        #region
        private void StockRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //Worksheet CurrenntSheet = Globals.ThisAddIn.GetActiveWorkSheet();
           
            //Header Style - Start
            Excel.Style HeaderStyle = ActiveSheet().Application.ActiveWorkbook.Styles.Add("HdrStyle");
            HeaderStyle.Font.Name = "Calibri";
            HeaderStyle.Font.Size = 12;
            HeaderStyle.Font.Bold = true;
            HeaderStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
            HeaderStyle.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Orange);

            Excel.Range HeaderRange = ActiveSheet().Range["A1", "B1"];

            ActiveSheet().Range["A1"].Value = "Stock Ticker";
            ActiveSheet().Range["B1"].Value = "Current Stock Value";
            //CurrenntSheet.Columns.AutoFit();
            HeaderRange.Style = "HdrStyle";
            //Header Style - End

            //Info Style - Start
            Excel.Style InfoStyle = ActiveSheet().Application.ActiveWorkbook.Styles.Add("InfoStyle");
            InfoStyle.Font.Name = "Calibri";
            InfoStyle.Font.Size = 11;
            InfoStyle.Font.Bold = true;
            InfoStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
            InfoStyle.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Orange);

            Excel.Range InfoRange = ActiveSheet().Range["A2", "B10"];
            fillStockItems();
            //CurrenntSheet.Columns.AutoFit();                  
            HeaderRange.Style = "InfoStyle";
            // Info Style -  End

            ActiveSheet().Columns.AutoFit();

        }
        
        private void btnAllStocks_Click(object sender, RibbonControlEventArgs e)
        {
            ClearStockPrices();

            int i = 2;
            foreach (string item in StockDictionary.Keys)
            {
                ActiveSheet().Range["B" + Convert.ToString(i)].Value = StockDictionary[item];
                i++;
            }

        }

        private void btnIndvStocks_Click(object sender, RibbonControlEventArgs e)
        {
            ActiveSheet().Range["A" + Convert.ToString(StockDictionary.Count + 2)].Value = string.Empty;
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            double stockVal = GetStockPrice(activeCell.Value);

            if (stockVal== 0)
            {
                ActiveSheet().Range["A" + Convert.ToString(StockDictionary.Count + 2)].Value = "Please select a Listed Company";
            }
            else
            {
                ActiveSheet().Range["B" + activeCell.Row.ToString()].Value = stockVal.ToString();
            }          

        }

        private void btnClear_Click(object sender, RibbonControlEventArgs e)
        {
            ClearStockPrices();
            
        }
        #endregion

        /// <summary>
        /// Private Metods
        /// </summary>
        /// <returns></returns>
        #region
        private Excel.Worksheet ActiveSheet()
        {
            return CurrentWorkSheet == null ? CurrentWorkSheet = Globals.ThisAddIn.GetActiveWorkSheet(): CurrentWorkSheet;
        }

        private double GetStockPrice(string CompanyName)
        {
            if (CompanyName == string.Empty || CompanyName.ToLower() == "stock ticker" || CompanyName.ToLower() == "current stock value")
                return 0;
            else
            {
                foreach (string item in StockDictionary.Keys)
                {
                    if (CompanyName == item)
                    {
                        return StockDictionary[item];
                    }                    
                }
                return 0;
            }
        }

        private void fillStockItems()
        {
            int i = 2;
            foreach (string item in StockDictionary.Keys)
            {
                ActiveSheet().Range["A" + Convert.ToString(i)].Value = item;
                i++;
            }
        }

        private void ClearStockPrices()
        {
            int i = 2;
            foreach (string item in StockDictionary.Keys)
            {
                ActiveSheet().Range["B" + Convert.ToString(i)].Value = string.Empty;
                i++;
            }
        }
        #endregion
    }
}

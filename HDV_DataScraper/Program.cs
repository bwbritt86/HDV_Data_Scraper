using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;

namespace HDV_DataScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            year2009();
            /*year2010();
            year2011();
            year2012();
            year2013();
            year2014();
            year2015();
            year2016();
            year2017();*/
        }

        static void year2009()
        {
            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"C:\Users\nikaido\Documents\College\Senior Project\Crime Stats\jun09.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2009";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            /*Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul09.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2009";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug09.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2009";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep09.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2009";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct09.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2009";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov09.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2009";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec09.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2009";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);*/
        }

        static void year2010()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan10.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January_2010";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb10.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February_2010";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar10.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March_2010";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr10.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April_2010";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may10.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May_2010";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun10.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2010";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul10.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2010";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug10.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2010";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep10.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2010";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct10.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2010";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov10.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2010";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec10.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2010";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }

        static void year2011()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan11.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January_2011";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb11.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February_2011";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar11.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March_2011";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr11.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April_2011";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may11.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May_2011";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun11.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2011";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul11.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2011";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug11.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2011";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep11.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2011";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct11.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2011";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov11.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2011";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec11.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2011";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }

        static void year2012()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan12.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January_2012";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb12.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February_2012";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar12.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March_2012";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr12.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April_2012";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may12.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May_2012";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun12.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2012";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul12.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2012";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug12.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2012";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep12.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2012";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct12.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2012";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov12.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2012";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec12.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2012";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }

        static void year2013()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan13.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January 2013";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb13.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February 2013";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar13.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March 2013";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr13.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April 2013";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may13.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May 2013";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun13.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June 2013";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul13.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July 2013";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug13.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August 2013";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep13.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September 2013";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct13.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October 2013";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov13.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November 2013";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec13.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December 2013";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }

        static void year2014()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan14.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January_2014";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb14.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February_2014";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar14.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March_2014";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr14.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April_2014";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may14.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May_2014";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun14.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2014";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul14.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2014";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug14.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2014";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep14.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2014";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct14.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2014";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov14.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2014";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec14.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2014";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }

        static void year2015()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan15.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January_2015";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb15.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February_2015";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar15.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March_2015";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr15.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April_2015";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may15.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May_2015";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun15.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2015";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul15.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2015";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug15.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2015";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep15.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2015";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct15.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2015";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov15.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2015";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec15.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2015";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }

        static void year2016()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan16.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January_2016";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb16.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February_2016";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar16.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March_2016";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr16.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April_2016";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may16.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May_2016";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun16.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2016";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul16.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2016";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug16.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2016";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep16.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2016";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct16.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2016";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov16.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2016";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec16.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2016";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }

        static void year2017()
        {
            Excel.Application JanXlApp = new Excel.Application();
            Excel.Workbook JanXlWkBk = JanXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jan17.xls");
            Excel._Worksheet JanXlWorksheet = JanXlWkBk.Sheets[1];
            Excel.Range JanXlRange = JanXlWorksheet.UsedRange;
            String JanTblName = "January_2017";
            insertIntoDB(JanXlApp, JanXlWkBk, JanXlWorksheet, JanXlRange, JanTblName);

            Excel.Application FebXlApp = new Excel.Application();
            Excel.Workbook FebXlWkBk = FebXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/feb17.xls");
            Excel._Worksheet FebXlWorksheet = FebXlWkBk.Sheets[1];
            Excel.Range FebXlRange = FebXlWorksheet.UsedRange;
            String FebTblName = "February_2017";
            insertIntoDB(FebXlApp, FebXlWkBk, FebXlWorksheet, FebXlRange, FebTblName);

            Excel.Application MarXlApp = new Excel.Application();
            Excel.Workbook MarXlWkBk = MarXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/mar17.xls");
            Excel._Worksheet MarXlWorksheet = MarXlWkBk.Sheets[1];
            Excel.Range MarXlRange = MarXlWorksheet.UsedRange;
            String MarTblName = "March_2017";
            insertIntoDB(MarXlApp, MarXlWkBk, MarXlWorksheet, MarXlRange, MarTblName);

            Excel.Application AprXlApp = new Excel.Application();
            Excel.Workbook AprXlWkBk = AprXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/apr17.xls");
            Excel._Worksheet AprXlWorksheet = AprXlWkBk.Sheets[1];
            Excel.Range AprXlRange = AprXlWorksheet.UsedRange;
            String AprTblName = "April_2017";
            insertIntoDB(AprXlApp, AprXlWkBk, AprXlWorksheet, AprXlRange, AprTblName);

            Excel.Application MayXlApp = new Excel.Application();
            Excel.Workbook MayXlWkBk = MayXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/may17.xls");
            Excel._Worksheet MayXlWorksheet = MayXlWkBk.Sheets[1];
            Excel.Range MayXlRange = MayXlWorksheet.UsedRange;
            String MayTblName = "May_2017";
            insertIntoDB(MayXlApp, MayXlWkBk, MayXlWorksheet, MayXlRange, MayTblName);

            Excel.Application JunXlApp = new Excel.Application();
            Excel.Workbook JunXlWkBk = JunXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jun17.xls");
            Excel._Worksheet JunXlWorksheet = JunXlWkBk.Sheets[1];
            Excel.Range JunXlRange = JunXlWorksheet.UsedRange;
            String JunTblName = "June_2017";
            insertIntoDB(JunXlApp, JunXlWkBk, JunXlWorksheet, JunXlRange, JunTblName);

            Excel.Application JulXlApp = new Excel.Application();
            Excel.Workbook JulXlWkBk = JulXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/jul17.xls");
            Excel._Worksheet JulXlWorksheet = JulXlWkBk.Sheets[1];
            Excel.Range JulXlRange = JulXlWorksheet.UsedRange;
            String JulTblName = "July_2017";
            insertIntoDB(JulXlApp, JulXlWkBk, JulXlWorksheet, JulXlRange, JulTblName);

            Excel.Application AugXlApp = new Excel.Application();
            Excel.Workbook AugXlWkBk = AugXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/aug17.xls");
            Excel._Worksheet AugXlWorksheet = AugXlWkBk.Sheets[1];
            Excel.Range AugXlRange = AugXlWorksheet.UsedRange;
            String AugTblName = "August_2017";
            insertIntoDB(AugXlApp, AugXlWkBk, AugXlWorksheet, AugXlRange, AugTblName);

            Excel.Application SepXlApp = new Excel.Application();
            Excel.Workbook SepXlWkBk = SepXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/sep17.xls");
            Excel._Worksheet SepXlWorksheet = SepXlWkBk.Sheets[1];
            Excel.Range SepXlRange = SepXlWorksheet.UsedRange;
            String SepTblName = "September_2017";
            insertIntoDB(SepXlApp, SepXlWkBk, SepXlWorksheet, SepXlRange, SepTblName);

            Excel.Application OctXlApp = new Excel.Application();
            Excel.Workbook OctXlWkBk = OctXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/oct17.xls");
            Excel._Worksheet OctXlWorksheet = OctXlWkBk.Sheets[1];
            Excel.Range OctXlRange = OctXlWorksheet.UsedRange;
            String OctTblName = "October_2017";
            insertIntoDB(OctXlApp, OctXlWkBk, OctXlWorksheet, OctXlRange, OctTblName);

            Excel.Application NovXlApp = new Excel.Application();
            Excel.Workbook NovXlWkBk = NovXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/nov17.xls");
            Excel._Worksheet NovXlWorksheet = NovXlWkBk.Sheets[1];
            Excel.Range NovXlRange = NovXlWorksheet.UsedRange;
            String NovTblName = "November_2017";
            insertIntoDB(NovXlApp, NovXlWkBk, NovXlWorksheet, NovXlRange, NovTblName);

            Excel.Application DecXlApp = new Excel.Application();
            Excel.Workbook DecXlWkBk = DecXlApp.Workbooks.Open(@"http://www.houstontx.gov/police/cs/xls/dec17.xls");
            Excel._Worksheet DecXlWorksheet = DecXlWkBk.Sheets[1];
            Excel.Range DecXlRange = DecXlWorksheet.UsedRange;
            String DecTblName = "December_2017";
            insertIntoDB(DecXlApp, DecXlWkBk, DecXlWorksheet, DecXlRange, DecTblName);
        }


        static void insertIntoDB(Excel.Application xlApp, Excel.Workbook xlWkBk, Excel._Worksheet xlWkSh, Excel.Range xlRange, String tblName)
        {
            //Set database connection string
            String conStr = "Data Source = dcm.uhcl.edu; Initial Catalog = c438818fa01g3; User ID = c438818fa01g3; Password = 6552950";

            //Create Table using tblName value
            using (SqlConnection conn = new SqlConnection(conStr))
            {
                //Create table command
                String createTbl = "Create TABLE " + tblName + "(Offense_Type char(50),Number char(50),Street char(50),Type char(50),Suffix char(10));";
                SqlCommand createCommand = new SqlCommand(createTbl, conn);

                //Create table in database
                try
                {
                    conn.Open();
                    createCommand.ExecuteNonQuery();
                    Console.WriteLine("Table " + tblName + " Created Successfully...");


                }
                finally
                {
                    conn.Close();
                }

                //Declare row and column counts
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int rowCounter = 0;
                int unkCount = 0;
                int counter = 0;

                //Begin looping through rows
                for (int i = 2; i <= rowCount; i++)
                {
                    //Declare variables for Excel Data
                    String offense = null;
                    String block = null;
                    String streetAddr = null;
                    String type = null;
                    String suffix = null;

                    //Begin looping through columns
                    for (int j = 1; j <= colCount; j++)
                    {
                        //Check if data in the cells
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            if (j == 4)
                                offense = (String)xlRange.Cells[i, j].Value2;
                            else if (j == 7)
                            {
                                if ((String)xlRange.Cells[i, j].Value2 == "unk")
                                {
                                    unkCount++;
                                    break; //break loop if block number is unknown (row does not get inerted into database)
                                }
                                else
                                    block = (String)xlRange.Cells[i, j].Value2;
                            }
                            else if (j == 8)
                                streetAddr = (String)xlRange.Cells[i, j].Value2;
                            else if (j == 9)
                            {
                                if ((String)xlRange.Cells[i, j].Value2 == "-")
                                    type = "";
                                else
                                    type = (String)xlRange.Cells[i, j].Value2;
                            }
                            else if (j == 10)
                            {
                                if ((String)xlRange.Cells[i, j].Value2 == "-")
                                    suffix = "";
                                else
                                    suffix = (String)xlRange.Cells[i, j].Value2;
                            }
                        }
                        counter++;
                        Console.WriteLine("counter =" + counter);
                    }

                    //Show the amount of rows skipped
                    Console.WriteLine("UNK Rows Skipped = " + unkCount);

                    //Do not insert row if any of the data is null
                    if (offense != null && block != null && streetAddr != null && type != null && suffix != null)
                    {
                        //Create insertion command
                        string insertRow = "INSERT INTO " + tblName + " (Offense_Type,Number,Street,Type,Suffix)"
                          + " VALUES ('" + offense + "','" + block + "','" + streetAddr + "','" + type + "','" + suffix + "');";
                        SqlCommand insertCommand = new SqlCommand(insertRow, conn);

                        //Insert into Database
                        try
                        {
                            conn.Open();
                            insertCommand.ExecuteNonQuery();
                        }
                        finally
                        {
                            conn.Close();
                        }
                    }
                    else if (offense == "" && block == "" && streetAddr == "" && type == "" && suffix == "")
                    {
                        Console.WriteLine("\nFinished populating [" + tblName + "]");
                        break;
                    }

                    rowCounter++;
                    Console.WriteLine("Row = " + rowCounter);
                }
    
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release objects
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWkSh);

            //close and release
            xlWkBk.Close();
            Marshal.ReleaseComObject(xlWkBk);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}

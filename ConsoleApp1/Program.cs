using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;

namespace ConsoleApp1
{
    class Program
    {
        static readonly string OriginalXLSXFile = "Pythia_Report_Original.xlsx";
        static string TemporaryXLSX = "Temporary_Pythia_File.xlsx";
        static readonly string FileNamePrefix = "CUDB_WIND_Q_Weekly_Report_For_";
        static readonly string FileNameSuffix = ".xlsx";
        static string DateOfReport = "";
        static string AdditionalRelativeExportPath = "new_reports";

        static void Main(string[] args)
        {

            string providedXlsxPath = System.Reflection.Assembly.GetEntryAssembly().Location; ;
            // providedXlsxPath = Path.GetDirectoryName(providedXlsxPath);

            providedXlsxPath = @"C:\Pythia_Excel_Report";

            try
            {
                // Check if temporary file exists
                if (File.Exists(Path.Combine(new string[] { providedXlsxPath, TemporaryXLSX })))
                {
                    File.Delete(Path.Combine(new string[] { providedXlsxPath, TemporaryXLSX }));
                }
            }
            catch (Exception e)
            {
                TemporaryXLSX = TemporaryXLSX + "_1";
                Console.WriteLine("e message " + e.Message);
            }
            // Copy Orig ExcelFile to temprorary Excel file
            File.Copy(Path.Combine(new string[] { providedXlsxPath, OriginalXLSXFile }), Path.Combine(new string[] { providedXlsxPath, TemporaryXLSX }));

            // Use Temporary Excel File
            UpdateExcelFile(Path.Combine(new string[] { providedXlsxPath, TemporaryXLSX }), args[0]);

            // Copy to Today File
            CopyFileToTodayFile(Path.Combine(new string[] { providedXlsxPath, TemporaryXLSX }));

            // Delete Temporary File
            DeleteTempFile(providedXlsxPath, TemporaryXLSX);

            // Console app
            System.Environment.Exit(1);

            /*
            // Exit Forms Application
            if (System.Windows.Forms.Application.MessageLoop)
            {
                // WinForms app
                System.Windows.Forms.Application.Exit();
            }
            else
            {
           
            }
            */
        }

        private static void DeleteTempFile(string providedXlsxPath, string TemporaryXLSX)
        {
            if (File.Exists(Path.Combine(new string[] { Path.GetDirectoryName(providedXlsxPath), TemporaryXLSX })))
            {
                File.Delete(Path.Combine(new string[] { Path.GetDirectoryName(providedXlsxPath), TemporaryXLSX }));
            }
        }
 
        // Copy Source File to -->> CUDB_WIND_Q_Weekly_Report_For_1_Αύγουστος_2018.xlsx
        private static void CopyFileToTodayFile(string filePath)
        {
            var culture = new System.Globalization.CultureInfo("en-UK");
            var day = culture.DateTimeFormat.GetDayName(DateTime.Today.DayOfWeek);

            if (day == "Friday")
            {
                AdditionalRelativeExportPath = Path.Combine("new_reports", "friday_only");
            }
            string currentDirectoryPath = Path.GetDirectoryName(filePath);

            // Check if destination file exists
            if (File.Exists(Path.Combine(new string[] { currentDirectoryPath, AdditionalRelativeExportPath, FileNamePrefix + DateOfReport + FileNameSuffix })))
            {
                File.Delete(Path.Combine(new string[] { currentDirectoryPath, AdditionalRelativeExportPath, FileNamePrefix + DateOfReport + FileNameSuffix }));
            }

            File.Copy(filePath, Path.Combine(new string[] { currentDirectoryPath, AdditionalRelativeExportPath, FileNamePrefix + DateOfReport + FileNameSuffix }));
        }

        private static void UpdateExcelFile(string xlsxFilePath, string dbName)
        {
            String weekOfDayFromDbName = dbName.Replace("CUDB_", "");

            MyExcel xl = new MyExcel(xlsxFilePath);
            Connect_To_Mysql sql = new Connect_To_Mysql("server=172.16.142.124;uid=pythia_ro;pwd=9esa8E2r9tWPhPB5M2mn;database=" + dbName + ";ConnectionTimeout=120;Port=3306;AllowUserVariables=true;DefaultCommandTimeout=120");
            MySqlConnection conn = sql.Connect_ToSQL();

            string ActiveSQLPred = " CSLOC <> 2 AND CSLOC <> 5 ";
            string SQL_Pred_1 = " CSLOC = 2 ";
            string SQL_Pred_2 = " VLRADD LIKE '1930%' ";
            string SQL_Pred_3 = " RVLRI = 1 ";
            string SQL_Pred_4 = " RVLRI = 0 ";
            string SQL_Pred_5 = " RVLRI = 2 ";
            string SQL_Pred_6 = " CSLOC = 5 ";
            string SQL_Pred_7 = " CSLOC = 4 ";
            // VLRADD LIKE '1930%' and RVLRI == 0 and CSLOC == 5
            // RVLRI == 2 and CSLOC == 5
            string SQL_Predicate_B = "";
            string ColumnPrefix = "";

            string sqlString = "";
            int WorkSheetNum = 0;

            // Open Excel File
            xl.openExcel();

            // Open MySQL Connection
            //MySqlConnection conn = Connect_ToSQL();

            // Per HLR WIND
            // Date
            sqlString = "SELECT DATE_FORMAT(CurrentDate,'%m/%d/%Y') as OutColumn FROM DateOfReport;";
            string date = sql.Exec_Query(conn, sqlString);
            xl.addSingleDataToExcel(1, "A", 2, date);
            xl.addSingleDataToExcel(1, "A", 18, date);

            DateTime dt = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            // Save Date also in this variable
            DateOfReport = dt.ToString("dd_MMMM_yyy");

            try
            {
                for (int i = 1; i <= 9; i++)
                {
                    WorkSheetNum = 1;
                    if (i == 1) { ColumnPrefix = "B"; SQL_Predicate_B = " STYPE = 5 "; }
                    if (i == 2) { ColumnPrefix = "C"; SQL_Predicate_B = " STYPE = 2 AND IMSI LIKE '20210%' "; }
                    if (i == 3) { ColumnPrefix = "D"; SQL_Predicate_B = " (STYPE = 9 OR STYPE = 12) "; }
                    if (i == 4) { ColumnPrefix = "E"; SQL_Predicate_B = " (STYPE = 8 OR STYPE = 14) "; }
                    if (i == 5) { ColumnPrefix = "F"; SQL_Predicate_B = " (STYPE = 11 OR STYPE = 15) "; }
                    if (i == 6) { ColumnPrefix = "G"; SQL_Predicate_B = " STYPE = 6 "; }
                    if (i == 7) { ColumnPrefix = "H"; SQL_Predicate_B = " STYPE = 3 "; }
                    if (i == 8) { ColumnPrefix = "I"; SQL_Predicate_B = " STYPE = 17 "; }
                    if (i == 9) { ColumnPrefix = "J"; SQL_Predicate_B = " STYPE = 19 "; }

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 3, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_1 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 4, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 5, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_3 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 6, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_4 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 7, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_5 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 8, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 9, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 10, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_3 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 11, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_4 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 12, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_5 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 13, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_7 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 14, sql.Exec_Query(conn, sqlString));

                }

                // Addition 24 Feb 2021 for WIND 5G Column
                // K3
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G'" + ";";
                xl.addSingleDataToExcel(1, "K", 3, sql.Exec_Query(conn, sqlString));

                // K4
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_1 + ";";
                xl.addSingleDataToExcel(1, "K", 4, sql.Exec_Query(conn, sqlString));

                // K5
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_2 + ";";
                xl.addSingleDataToExcel(1, "K", 5, sql.Exec_Query(conn, sqlString));

                // K6
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_3 + ";";
                xl.addSingleDataToExcel(1, "K", 6, sql.Exec_Query(conn, sqlString));

                // K7
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_4 + ";";
                xl.addSingleDataToExcel(1, "K", 7, sql.Exec_Query(conn, sqlString));

                // K8
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_5 + ";";
                xl.addSingleDataToExcel(1, "K", 8, sql.Exec_Query(conn, sqlString));

                // K9
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_6 + ";";
                xl.addSingleDataToExcel(1, "K", 9, sql.Exec_Query(conn, sqlString));

                // K10
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_6 + ";";
                xl.addSingleDataToExcel(1, "K", 10, sql.Exec_Query(conn, sqlString));

                // K11
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_3 + "AND" + SQL_Pred_6 + ";";
                xl.addSingleDataToExcel(1, "K", 11, sql.Exec_Query(conn, sqlString));

                // K12
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_4 + "AND" + SQL_Pred_6 + ";";
                xl.addSingleDataToExcel(1, "K", 12, sql.Exec_Query(conn, sqlString));

                // K13
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_5 + "AND" + SQL_Pred_6 + ";";
                xl.addSingleDataToExcel(1, "K", 13, sql.Exec_Query(conn, sqlString));

                // K14
                sqlString = "SELECT count(CS.MSISDN) as OutColumn from App_DS_XX_CSPS as CS INNER JOIN App_DS_XX_EPS_EpsStaInf as EP ON CS.ID = EP.ID WHERE EP.EpsProfileId = '5G' " + "AND" + SQL_Pred_7 + ";";
                xl.addSingleDataToExcel(1, "K", 14, sql.Exec_Query(conn, sqlString));

                // Addition of VoLTE per WIND - 01 Sep 2021
                // L3
                string sqlString_master_volte_wind = @"SELECT count(CS.IMSI) as OutColumn
                                from
                                (
                                select
                                SUBSTR(IFNULL(ImsAssocImpi, ''), 1, POSITION('@' IN ImsAssocImpi) - 1) as ImsIMSI
                                from
                                (
                                SELECT DISTINCT ImsAssocImpi
                                FROM CUDB_Assoc_" + weekOfDayFromDbName + @".IMPU
                                where SUBSTR(IFNULL(ImsAssocImpi, ''), POSITION('@' IN ImsAssocImpi) + 1, length(ImsAssocImpi)) = 'ims.mnc010.mcc202.3gppnetwork.org'
                                    AND SUBSTR(IFNULL(ImsAssocImpi, ''), 1 , POSITION('@' IN ImsAssocImpi) - 1) like '202%'
                                ) as a
                                ) as ImsData
                                INNER JOIN
                                CUDB_" + weekOfDayFromDbName + @".App_DS_XX_CSPS as CS
                                ON ImsData.ImsIMSI = CS.IMSI";

                xl.addSingleDataToExcel(1, "L", 3, sql.Exec_Query(conn, sqlString_master_volte_wind));

                // L4
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.CSLOC = 2" + ";";
                xl.addSingleDataToExcel(1, "L", 4, sql.Exec_Query(conn, sqlString));

                // L5
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.VLRADD LIKE '1930%'" + ";";
                xl.addSingleDataToExcel(1, "L", 5, sql.Exec_Query(conn, sqlString));

                // L6
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.RVLRI = 1" + ";";
                xl.addSingleDataToExcel(1, "L", 6, sql.Exec_Query(conn, sqlString));

                // L7
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.VLRADD LIKE '1930%' and CS.RVLRI = 0" + ";";
                xl.addSingleDataToExcel(1, "L", 7, sql.Exec_Query(conn, sqlString));

                // L8
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.RVLRI = 2" + ";";
                xl.addSingleDataToExcel(1, "L", 8, sql.Exec_Query(conn, sqlString));

                // L9
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "L", 9, sql.Exec_Query(conn, sqlString));

                // L10
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.VLRADD LIKE '1930%' and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "L", 10, sql.Exec_Query(conn, sqlString));

                // L11
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.VLRADD LIKE '1930%' and CS.RVLRI = 1 and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "L", 11, sql.Exec_Query(conn, sqlString));

                // L12 
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.VLRADD LIKE '1930%' and CS.RVLRI = 0 and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "L", 12, sql.Exec_Query(conn, sqlString));

                // L13 
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.RVLRI = 2 and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "L", 13, sql.Exec_Query(conn, sqlString));

                // L14
                sqlString = sqlString_master_volte_wind + " WHERE " + "CS.CSLOC = 4" + ";";
                xl.addSingleDataToExcel(1, "L", 14, sql.Exec_Query(conn, sqlString));

                // Addition of General VoLTE Calculation
                sqlString = @"SELECT 
                            SUBSTR(IFNULL(ImsAssocImpi, ''), POSITION('@' IN ImsAssocImpi) + 1, length(ImsAssocImpi)) as ImsAssocImpi_Substring,
                            count(SUBSTR(IFNULL(ImsAssocImpi, ''), POSITION('@' IN ImsAssocImpi) + 1, length(ImsAssocImpi))) as OutColumn
                            from
                            ( SELECT DISTINCT ImsAssocImpi FROM CUDB_Assoc_" + weekOfDayFromDbName + @".IMPU where SUBSTR(IFNULL(ImsAssocImpi, ''), POSITION('@' IN ImsAssocImpi) + 1, length(ImsAssocImpi))  = 'ims.mnc010.mcc202.3gppnetwork.org'
                               AND SUBSTR(IFNULL(ImsAssocImpi, ''), 1 , POSITION('@' IN ImsAssocImpi) - 1) like '202%'
                            ) as a
                            group by ImsAssocImpi_Substring;";
                xl.addSingleDataToExcel(1, "G", 19, sql.Exec_Query(conn, sqlString));


                // Per HLR Q
                for (int i = 1; i <= 3; i++)
                {
                    WorkSheetNum = 1;
                    if (i == 1) { ColumnPrefix = "B"; SQL_Predicate_B = " STYPE = 7 "; }
                    if (i == 2) { ColumnPrefix = "C"; SQL_Predicate_B = " STYPE = 2 and IMSI Like '20209%' "; }
                    if (i == 3) { ColumnPrefix = "D"; SQL_Predicate_B = " STYPE = 20 and IMSI Like '2021009%' "; }

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 19, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_1 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 20, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 21, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_3 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 22, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_4 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 23, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_5 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 24, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 25, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 26, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_3 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 27, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_2 + "AND" + SQL_Pred_4 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 28, sql.Exec_Query(conn, sqlString));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS WHERE" + SQL_Predicate_B + "AND" + SQL_Pred_5 + "AND" + SQL_Pred_6 + ";";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 29, sql.Exec_Query(conn, sqlString));

                }

                // Addition of VoLTE per Q - 01 Sep 2021
                // D19
                string sqlString_master_volte_q = @"SELECT count(CS.IMSI) as OutColumn
                                                    from
                                                    (
                                                    select 
                                                    SUBSTR(IFNULL(ImsAssocImpi, ''), 1, POSITION('@' IN ImsAssocImpi)-1) as ImsIMSI
                                                    from
                                                    (
                                                    SELECT DISTINCT ImsAssocImpi 
                                                    FROM CUDB_Assoc_" + weekOfDayFromDbName + @".IMPU
                                                    where SUBSTR(IFNULL(ImsAssocImpi, ''), POSITION('@' IN ImsAssocImpi) + 1, length(ImsAssocImpi))  = 'ims.mnc010.mcc202.3gppnetwork.org'
                                                        AND SUBSTR(IFNULL(ImsAssocImpi, ''), 1 , POSITION('@' IN ImsAssocImpi) - 1) like '202%'
                                                    ) as a
                                                    ) as ImsData
                                                    INNER JOIN
                                                    CUDB_" + weekOfDayFromDbName + @".App_DS_XX_CSPS as CS
                                                    ON ImsData.ImsIMSI = CS.IMSI
                                                    WHERE (CS.STYPE = 7 or (CS.STYPE = 2 and IMSI Like '20209%') or (CS.STYPE = 20 and IMSI Like '2021009%'))";

                xl.addSingleDataToExcel(1, "E", 19, sql.Exec_Query(conn, sqlString_master_volte_q));

                // D20
                sqlString = sqlString_master_volte_q + " AND " + "CS.CSLOC = 2" + ";";
                xl.addSingleDataToExcel(1, "E", 20, sql.Exec_Query(conn, sqlString));

                // D21
                sqlString = sqlString_master_volte_q + " AND " + "CS.VLRADD LIKE '1930%'" + ";";
                xl.addSingleDataToExcel(1, "E", 21, sql.Exec_Query(conn, sqlString));

                // D22
                sqlString = sqlString_master_volte_q + " AND " + "CS.RVLRI = 1" + ";";
                xl.addSingleDataToExcel(1, "E", 22, sql.Exec_Query(conn, sqlString));

                // D23
                sqlString = sqlString_master_volte_q + " AND " + "CS.VLRADD LIKE '1930%' and CS.RVLRI = 0" + ";";
                xl.addSingleDataToExcel(1, "E", 23, sql.Exec_Query(conn, sqlString));

                // D24
                sqlString = sqlString_master_volte_q + " AND " + "CS.RVLRI = 2" + ";";
                xl.addSingleDataToExcel(1, "E", 24, sql.Exec_Query(conn, sqlString));

                // D25
                sqlString = sqlString_master_volte_q + " AND " + "CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "E", 25, sql.Exec_Query(conn, sqlString));

                // D26
                sqlString = sqlString_master_volte_q + " AND " + "CS.VLRADD LIKE '1930%' and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "E", 26, sql.Exec_Query(conn, sqlString));

                // D27
                sqlString = sqlString_master_volte_q + " AND " + "CS.VLRADD LIKE '1930%' and CS.RVLRI = 1 and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "E", 27, sql.Exec_Query(conn, sqlString));

                // D28
                sqlString = sqlString_master_volte_q + " AND " + "CS.VLRADD LIKE '1930%' and CS.RVLRI = 0 and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "E", 28, sql.Exec_Query(conn, sqlString));

                // D29
                sqlString = sqlString_master_volte_q + " AND " + "CS.RVLRI = 2 and CS.CSLOC = 5" + ";";
                xl.addSingleDataToExcel(1, "E", 29, sql.Exec_Query(conn, sqlString));


                // Per VLR WIND
                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 5 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "A", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 5 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "B", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 2 AND IMSI LIKE '20210%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "D", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 2 AND IMSI LIKE '20210%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "E", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE (STYPE = 9 OR STYPE = 12) AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "G", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE (STYPE = 9 OR STYPE = 12) AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "H", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 8 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "J", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 8 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "K", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 11 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "M", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 11 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "N", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 6 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "P", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 6 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "Q", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 3 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "S", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 3 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "T", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 17 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "V", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 17 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "W", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 19 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "Y", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 19 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(2, "Z", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                // Per VLR Q
                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 2 AND IMSI LIKE '20209%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(3, "A", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 2 AND IMSI LIKE '20209%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(3, "B", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 7 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(3, "D", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 7 AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(3, "E", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 20 AND IMSI like '2021009%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(3, "G", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 20 AND IMSI like '2021009%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(3, "H", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                // Active Per VLR
                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 5 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "A", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 5 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "B", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 2 AND IMSI LIKE '20210%' " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "D", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 2 AND IMSI LIKE '20210%' " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "E", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE (STYPE = 9 OR STYPE = 12) " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "G", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE (STYPE = 9 OR STYPE = 12) " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "H", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 8 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "J", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 8 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "K", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 11 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "M", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 11 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "N", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 6 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "P", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 6 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "Q", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 3 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "S", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 3 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "T", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 17 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "V", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 17 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "W", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 19 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "Y", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 19 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(4, "Z", 3, sql.Exec_MultiLine_Query(conn, sqlString));


                // Active Per VLR Q
                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 2 " + "AND" + ActiveSQLPred + "AND" + " IMSI LIKE '20209%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(5, "A", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 2 " + "AND" + ActiveSQLPred + "AND" + " IMSI LIKE '20209%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(5, "B", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 7 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(5, "D", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 7 " + "AND" + ActiveSQLPred + "AND" + " VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(5, "E", 3, sql.Exec_MultiLine_Query(conn, sqlString));

                sqlString = "select VLRADD as COL1, COUNT(*) from App_DS_XX_CSPS WHERE STYPE = 7 " + "AND" + ActiveSQLPred + "AND" + " STYPE = 20 AND IMSI like '2021009%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(5, "G", 3, sql.Exec_MultiLine_Query(conn, sqlString));
                sqlString = "select VLRADD, COUNT(*) as COL1 from App_DS_XX_CSPS WHERE STYPE = 7 " + "AND" + ActiveSQLPred + "AND" + " STYPE = 20 AND IMSI like '2021009%' AND VLRADD IS NOT NULL GROUP BY VLRADD" + ";";
                xl.addMultipleDataToExcel(5, "H", 3, sql.Exec_MultiLine_Query(conn, sqlString));


                // USIM 
                WorkSheetNum = 6;
                ColumnPrefix = "B";

                //string USIM_Predicate = " IM.AKATYPE = 1 ";
                for (int i = 1; i <= 1; i++)
                {

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 3, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE CSLOC = '2' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 4, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE VLRADD LIKE '1930%' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 5, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE RVLRI = '1' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 6, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE VLRADD LIKE '1930%' AND RVLRI = 0 AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 7, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE RVLRI = '2' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 8, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE CSLOC = 5 AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 9, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE VLRADD LIKE '1930%' AND CSLOC = '5' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 10, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE VLRADD LIKE '1930%' AND RVLRI = '1' AND CSLOC = '5' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 11, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE VLRADD LIKE '1930%' AND RVLRI = '0' AND CSLOC = '5' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 12, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE RVLRI = '2' AND CSLOC = '5' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 13, sql.Exec_Query(conn, sqlString.ToString()));

                    sqlString = "select COUNT(MSISDN) as OutColumn from App_DS_XX_CSPS INNER JOIN App_DS_XX_IMSI ON App_DS_XX_CSPS.ID = App_DS_XX_IMSI.ID WHERE CSLOC = '4' AND App_DS_XX_IMSI.AKATYPE = '1';";
                    xl.addSingleDataToExcel(WorkSheetNum, ColumnPrefix, 14, sql.Exec_Query(conn, sqlString.ToString()));

                }
                // MessageBox.Show("Execution Completed!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                // Save & Close Excel File
                xl.closeExcel();
            }
    }
    }

    public class Connect_To_Mysql
    {
        string myConnectionString;
        MySqlConnection conn;

        public Connect_To_Mysql(string conn_string)
        {
            this.myConnectionString = conn_string;
        }

        public void Disconnect_From_Mysql()
        {
            conn.Close();
        }

        public MySqlConnection Connect_ToSQL()
        {
            conn = new MySql.Data.MySqlClient.MySqlConnection(myConnectionString);
            conn.Open();
            return conn;
        }

        public string Exec_Query(MySqlConnection conn, string query)
        {
            string returnString = "";
            MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(query, conn);
            var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                var someValue = reader["OutColumn"];
                //MessageBox.Show(someValue.ToString());
                returnString = someValue.ToString();
            }

            reader.Close();

            return returnString;
        }

        public string[] Exec_MultiLine_Query(MySqlConnection conn, string query)
        {
            List<string> myList1 = new List<string>();

            MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(query, conn);
            var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                myList1.Add(reader["COL1"].ToString());
            }

            reader.Close();
            string[] array = myList1.ToArray();
            return array;
        }
    }
    public class MyExcel
    {
        string excelFilePath;
        int rowNumber = 1;
        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet myExcelWorkSheet;

        public MyExcel(string myExcelPath)
        {
            this.excelFilePath = myExcelPath;
        }
        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }

        public int Rownumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }

        public void openExcel()
        {
            myExcelApplication = null;

            myExcelApplication = new Excel.Application(); // create Excel App
            myExcelApplication.DisplayAlerts = false; // turn off alerts


            myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks._Open(excelFilePath, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

            int numberOfWorkbooks = myExcelApplication.Workbooks.Count; // get number of workbooks (optional)

            //myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1]; // define in which worksheet, do you want to add data
            //myExcelWorkSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

            int numberOfSheets = myExcelWorkbook.Worksheets.Count; // get number of worksheets (optional)
        }

        public void addSingleDataToExcel(int workSheetNumber, string columnLetter, int rowNumber, string data)
        {
            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[workSheetNumber]; // define in which worksheet, do you want to add data
            myExcelWorkSheet.Cells[rowNumber, columnLetter] = data;

            //rowNumber++;  // if you put this method inside a loop, you should increase rownumber by one or wat ever is your logic
        }

        public void addMultipleDataToExcel(int workSheetNumber, string startColumnLetter, int startRowNumber, string[] data)
        {
            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[workSheetNumber]; // define in which worksheet, do you want to add data

            foreach (string row in data)
            {
                myExcelWorkSheet.Cells[startRowNumber, startColumnLetter] = row;
                startRowNumber++;
            }
        }


        public void closeExcel()
        {
            try
            {
                Console.WriteLine("excelFilePath = " + excelFilePath);
                myExcelWorkbook.SaveAs(excelFilePath, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value); // Save data in excel


                myExcelWorkbook.Close(true, excelFilePath, System.Reflection.Missing.Value); // close the worksheet


                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("e message " + e.Message);
                System.Diagnostics.Debug.WriteLine(e);
            }
            finally
            {
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                }

                if (myExcelWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(myExcelWorkSheet);
                    myExcelWorkSheet = null;
                }
                if (myExcelWorkbook != null)
                {
                    Marshal.FinalReleaseComObject(myExcelWorkbook);
                    myExcelWorkbook = null;
                }
                if (myExcelApplication != null)
                {
                    Marshal.FinalReleaseComObject(myExcelApplication);
                    myExcelApplication = null;
                }
            }
        }
    }
}

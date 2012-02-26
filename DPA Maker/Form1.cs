using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using Microsoft.Win32;

namespace DPA_Maker
{
    public partial class Form1 : Form
    {
        ////////////////////////////////////////////////////////////////////////
        //Creates the variables that are filled with the input from textboxes// 
        //////////////////////////////////////////////////////////////////////

        string fullName;
        string stdntNmbr;
        int grade;
        string peClass;
        string peTeacher;
        string sport;
        bool doYouPractice;
        bool doYouHaveGames;
        int howLongToSchool = 0;
        decimal howLongToSchoolDecimal;
        
        /////////////////////////////////////////////////////////////////////////////
        //Create input checker variables so text input can be checked for validity//
        ///////////////////////////////////////////////////////////////////////////

        bool fullNameGood = false;
        bool stdntNmbrGood = false;
        bool gradeGood = false;
        bool sportGood = false;
        bool doYouPracticeGood = false;
        bool doYouHaveGamesGood = false;
        bool howLongToSchoolGood = false;
        int practiceOptionsChosen = 0;
        int gameOptionChosen = 0;

        /////////////////////////////////////
        //Create Day of the week variables//
        ///////////////////////////////////

        string practiceMonday = "";
        string practiceTuesday = "";
        string practiceWednesday = "";
        string practiceThursday = "";
        string practiceFriday = "";

        string activitySaturday = "";
        string activitySunday = "";

        int minuteMonday;
        int minuteTuesday;
        int minuteWednesday;
        int minuteThursday;
        int minuteFriday;
        int minuteSaturday;
        int minuteSunday;
        int practiceTime = 90;



        public void checkInfo()
        {
            if (fullName == "")
            {
                label9.Show();
            }
            else
            {
                label9.Hide();
                fullNameGood = true;
            }

            if (stdntNmbr == "")
            {
                label10.Show();
            }
            else
            {
                label10.Hide();
                stdntNmbrGood = true;
            }

            if (sport == "")
            {
                label12.Show();
            }
            else
            {
                label12.Hide();
                sportGood = true;
            }
            if (practiceOptionsChosen == 0)
            {
                label13.Show();
            }
            else
            {
                label13.Hide();
                doYouPracticeGood = true;
            }
            if (gameOptionChosen == 0)
            {
                label15.Show();
            }
            else
            {
                label15.Hide();
                doYouHaveGamesGood = true;
            }

           if (howLongToSchoolDecimal == 0)
            {
                label14.Show();
                howLongToSchoolGood = false;
            }
            else
            {

                label14.Hide();
                howLongToSchoolGood = true;
            }

           if (grade == 0)
           {
               grade = 8;
           }
           else if (grade == 1)
           {
               grade = 9;
           }
           else if (grade == 2)
           {
               grade = 10;
           }
           else if (grade == 3)
           {
               grade = 11;
           }
           else if (grade == 4)
           {
               grade = 12;
           }

            //checks to see if excel is installed, an essential program that is software uses.
           RegistryKey key = Registry.ClassesRoot;
           RegistryKey excelKey = key.OpenSubKey("Excel.Application");
           bool excelInstalled = excelKey == null ? false : true;

           if (excelInstalled == false)
           {
               MessageBox.Show("Microsoft Excel not Installed.  Install Excel or leave me alone ;)", "No Excel!",
               MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
           }
            //Check to see if it's all goooooood!
           if (fullNameGood == true && stdntNmbrGood == true && sportGood == true && doYouPracticeGood == true && doYouHaveGamesGood == true && howLongToSchoolGood == true && excelInstalled == true)
            {   
                ///////////////////////////////////////////////////////////////////
                //Launches all the other functions if all the info is filled out//
                /////////////////////////////////////////////////////////////////

                //Figure out what day practice is on
                practiceDay();
                
                //Launches Loading Window
                //loadingscreen1 loadingScreen = new loadingscreen1();
                //loadingScreen.ShowDialog();
                
                //Figures out what day your main activity is on
                activityDay();

                //SHows the Progress Bar
                loadingscreen1 loadingScreen = new loadingscreen1();

                loadingScreen.Show();



                //Launches the function that fills in the entire excel spreadsheet
                fillInExcel();

            }
        }

        public void fillInExcel()
        {
            ////////////////////////////////////////////
            //Writes everything to the excel document//
            //////////////////////////////////////////

            Excel.Application myExcelApp;
            Excel.Workbooks myExcelWorkbooks;
            Excel.Workbook myExcelWorkbook;


            object misValue = System.Reflection.Missing.Value;


            myExcelApp = new Excel.Application();
            myExcelApp.Visible = true;
            myExcelWorkbooks = myExcelApp.Workbooks;
            //String fileName = @"C:\Term_2_ DPA.xls"; //set this to your file you want
            String fileName = Path.GetDirectoryName(Application.ExecutablePath) + @"\Term_2_ DPA.xls"; //set this to your file you want
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.ActiveSheet;

            //Writes Basic Info to excel file
            myExcelWorksheet.get_Range("D5", misValue).Formula = fullName;
            myExcelWorksheet.get_Range("D6", misValue).Formula = stdntNmbr;
            myExcelWorksheet.get_Range("D7", misValue).Formula = grade;
            myExcelWorksheet.get_Range("D8", misValue).Formula = peClass;
            myExcelWorksheet.get_Range("D9", misValue).Formula = peTeacher;

            //Fill in Activities
            myExcelWorksheet.get_Range("D17", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D20", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D23", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D26", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D29", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D32", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D35", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D39", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D42", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D45", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D48", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D51", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D54", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D57", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D61", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D64", misValue).Formula = "" + practiceMonday; //Winter Break Start
            myExcelWorksheet.get_Range("D67", misValue).Formula = "" + practiceTuesday;
            myExcelWorksheet.get_Range("D70", misValue).Formula = "" + practiceWednesday;
            myExcelWorksheet.get_Range("D73", misValue).Formula = "" + practiceThursday;
            myExcelWorksheet.get_Range("D76", misValue).Formula = "" + practiceFriday;
            myExcelWorksheet.get_Range("D79", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D97", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D100", misValue).Formula = "" + practiceMonday;
            myExcelWorksheet.get_Range("D103", misValue).Formula = "" + practiceTuesday;
            myExcelWorksheet.get_Range("D106", misValue).Formula = "" + practiceWednesday;
            myExcelWorksheet.get_Range("D109", misValue).Formula = "" + practiceThursday;
            myExcelWorksheet.get_Range("D112", misValue).Formula = "" + practiceFriday;
            myExcelWorksheet.get_Range("D115", misValue).Formula = "" + activitySaturday; //Winter Break End
            myExcelWorksheet.get_Range("D119", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D122", misValue).Formula = "" + practiceMonday; //Pro-D Day
            myExcelWorksheet.get_Range("D125", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D128", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D131", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D134", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D137", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D141", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D144", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D147", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D150", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D153", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D156", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D159", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D177", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D180", misValue).Formula = "" + practiceMonday; //Pro-D Day
            myExcelWorksheet.get_Range("D183", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D186", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D189", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D192", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D195", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D199", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D202", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D205", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D208", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D211", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D214", misValue).Formula = "" + practiceFriday; //Pro-D Day
            myExcelWorksheet.get_Range("D217", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D221", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D224", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D227", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D230", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D233", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D236", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D239", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D257", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D260", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D263", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D266", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D269", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D272", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D275", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D279", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D282", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D285", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D288", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D291", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D294", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D297", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D301", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D304", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D307", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D310", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D313", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D316", misValue).Formula = "" + practiceFriday; //Pro-D Day
            myExcelWorksheet.get_Range("D319", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D337", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D340", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D343", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D346", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D349", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D352", misValue).Formula = "Walk to School. " + practiceFriday;
            myExcelWorksheet.get_Range("D355", misValue).Formula = "" + activitySaturday;
            myExcelWorksheet.get_Range("D359", misValue).Formula = "" + activitySunday; //sunday
            myExcelWorksheet.get_Range("D362", misValue).Formula = "Walk to School. " + practiceMonday;
            myExcelWorksheet.get_Range("D365", misValue).Formula = "Walk to School. " + practiceTuesday;
            myExcelWorksheet.get_Range("D368", misValue).Formula = "Walk to School. " + practiceWednesday;
            myExcelWorksheet.get_Range("D371", misValue).Formula = "Walk to School. " + practiceThursday;
            myExcelWorksheet.get_Range("D374", misValue).Formula = "Walk to School. " + practiceFriday;

            //Fill in times
            myExcelWorksheet.get_Range("E17", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E20", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E23", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E26", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E29", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E32", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E35", misValue).Formula = "";
            myExcelWorksheet.get_Range("E39", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E42", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E45", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E48", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E51", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E54", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E57", misValue).Formula = "";
            myExcelWorksheet.get_Range("E61", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E64", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E67", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E70", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E73", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E76", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E79", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E97", misValue).Formula = ""; //sunday //Winter Break
            myExcelWorksheet.get_Range("E100", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E103", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E106", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E109", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E112", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E115", misValue).Formula = ""; //Winter Break
            myExcelWorksheet.get_Range("E119", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E122", misValue).Formula = "" + minuteMonday; //Pro D Day
            myExcelWorksheet.get_Range("E125", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E128", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E131", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E134", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E137", misValue).Formula = "";
            myExcelWorksheet.get_Range("E141", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E144", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E147", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E150", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E153", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E156", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E159", misValue).Formula = "";
            myExcelWorksheet.get_Range("E177", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E180", misValue).Formula = ""; //Pro-D Day
            myExcelWorksheet.get_Range("E183", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E186", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E189", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E192", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E195", misValue).Formula = "";
            myExcelWorksheet.get_Range("E199", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E202", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E205", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E208", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E211", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E214", misValue).Formula = ""; //Pro-D Day
            myExcelWorksheet.get_Range("E217", misValue).Formula = "";
            myExcelWorksheet.get_Range("E221", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E224", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E227", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E230", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E233", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E236", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E239", misValue).Formula = "";
            myExcelWorksheet.get_Range("E257", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E260", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E263", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E266", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E269", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E272", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E275", misValue).Formula = "";
            myExcelWorksheet.get_Range("E279", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E282", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E285", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E288", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E291", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E294", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E297", misValue).Formula = "";
            myExcelWorksheet.get_Range("E301", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E304", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E307", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E310", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E313", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E316", misValue).Formula = ""; //Pro-D Day
            myExcelWorksheet.get_Range("E319", misValue).Formula = "";
            myExcelWorksheet.get_Range("E337", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E340", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E343", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E346", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E349", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E352", misValue).Formula = "" + minuteFriday;
            myExcelWorksheet.get_Range("E355", misValue).Formula = "";
            myExcelWorksheet.get_Range("E359", misValue).Formula = ""; //sunday
            myExcelWorksheet.get_Range("E362", misValue).Formula = "" + minuteMonday;
            myExcelWorksheet.get_Range("E365", misValue).Formula = "" + minuteTuesday;
            myExcelWorksheet.get_Range("E368", misValue).Formula = "" + minuteWednesday;
            myExcelWorksheet.get_Range("E371", misValue).Formula = "" + minuteThursday;
            myExcelWorksheet.get_Range("E374", misValue).Formula = "" + minuteFriday;

            //Fill in Activity Type
            myExcelWorksheet.get_Range("F18", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F21", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F24", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F27", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F30", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F33", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F36", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F40", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F43", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F46", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F49", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F52", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F55", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F58", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F62", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F65", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F68", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F71", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F74", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F77", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F80", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F98", misValue).Formula = "x"; //sunday //Winter Break
            myExcelWorksheet.get_Range("F101", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F104", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F107", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F110", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F113", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F116", misValue).Formula = "x"; //Winter Break
            myExcelWorksheet.get_Range("F120", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F123", misValue).Formula = "x"; //Pro D Day
            myExcelWorksheet.get_Range("F126", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F129", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F132", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F135", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F138", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F142", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F145", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F148", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F151", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F154", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F157", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F160", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F178", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F181", misValue).Formula = "x"; //Pro-D Day
            myExcelWorksheet.get_Range("F184", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F187", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F190", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F193", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F196", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F200", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F203", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F206", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F209", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F212", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F215", misValue).Formula = "x"; //Pro-D Day
            myExcelWorksheet.get_Range("F218", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F222", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F225", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F228", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F231", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F234", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F237", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F240", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F258", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F261", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F264", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F267", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F270", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F273", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F276", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F280", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F283", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F286", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F289", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F292", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F295", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F298", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F302", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F305", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F308", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F311", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F314", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F317", misValue).Formula = "x"; //Pro-D Day
            myExcelWorksheet.get_Range("F320", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F338", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F341", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F344", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F347", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F350", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F353", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F356", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F360", misValue).Formula = "x"; //sunday
            myExcelWorksheet.get_Range("F363", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F366", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F369", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F372", misValue).Formula = "x";
            myExcelWorksheet.get_Range("F375", misValue).Formula = "x";
        }

        public void practiceDay()
        {
            Random random = new Random();
            int practiceDay = random.Next(1, 6);

            if (practiceDay == 1 && doYouPractice == true)
            {
                practiceMonday = sport + " Practice";
                practiceTuesday = "";
                practiceWednesday = "";
                practiceThursday = "";
                practiceFriday = "";
                
                minuteMonday = howLongToSchool + practiceTime;
                minuteTuesday = howLongToSchool;
                minuteWednesday = howLongToSchool;
                minuteThursday = howLongToSchool;
                minuteFriday = howLongToSchool;
            }
            else if (practiceDay == 2 && doYouPractice == true)
            {
                practiceTuesday = sport + " Practice";
                practiceMonday = "";
                practiceWednesday = "";
                practiceThursday = "";
                practiceFriday = "";
                
                minuteTuesday = howLongToSchool + practiceTime;
                minuteWednesday = howLongToSchool;
                minuteThursday = howLongToSchool;
                minuteFriday = howLongToSchool;
                minuteMonday = howLongToSchool;
            }
            else if (practiceDay == 3 && doYouPractice == true)
            {
                practiceWednesday = sport + " Practice";
                practiceMonday = "";
                practiceTuesday = "";
                practiceThursday = "";
                practiceFriday = "";
                
                minuteWednesday = howLongToSchool + practiceTime;
                minuteMonday = howLongToSchool;
                minuteTuesday = howLongToSchool;
                minuteThursday = howLongToSchool;
                minuteFriday = howLongToSchool;
            }
            else if (practiceDay == 4 && doYouPractice == true)
            {
                practiceThursday = sport + " Practice";
                practiceMonday = "";
                practiceTuesday = "";
                practiceWednesday = "";
                practiceFriday = "";
                
                minuteThursday = howLongToSchool + practiceTime;
                minuteMonday = howLongToSchool ;
                minuteTuesday = howLongToSchool;
                minuteWednesday = howLongToSchool;
                minuteFriday = howLongToSchool;
            }
            else if (practiceDay == 5 && doYouPractice == true)
            {
                practiceFriday = sport + " Practice";
                practiceMonday = "";
                practiceTuesday = "";
                practiceWednesday = "";
                practiceThursday = "";
                
                minuteFriday = howLongToSchool + practiceTime;
                minuteMonday = howLongToSchool;
                minuteTuesday = howLongToSchool;
                minuteWednesday = howLongToSchool;
                minuteThursday = howLongToSchool;
            }

            //They they don't have "practices" leave out teh work practice (DUH!!!!)

            if (practiceDay == 1 && doYouPractice == false)
            {
                practiceMonday = sport;
                practiceTuesday = "";
                practiceWednesday = "";
                practiceThursday = "";
                practiceFriday = "";

                minuteMonday = howLongToSchool + practiceTime;
                minuteTuesday = howLongToSchool;
                minuteWednesday = howLongToSchool;
                minuteThursday = howLongToSchool;
                minuteFriday = howLongToSchool;
            }
            else if (practiceDay == 2 && doYouPractice == false)
            {
                practiceTuesday = sport;
                practiceMonday = "";
                practiceWednesday = "";
                practiceThursday = "";
                practiceFriday = "";

                minuteTuesday = howLongToSchool + practiceTime;
                minuteWednesday = howLongToSchool;
                minuteThursday = howLongToSchool;
                minuteFriday = howLongToSchool;
                minuteMonday = howLongToSchool;
            }
            else if (practiceDay == 3 && doYouPractice == false)
            {
                practiceWednesday = sport;
                practiceMonday = "";
                practiceTuesday = "";
                practiceThursday = "";
                practiceFriday = "";

                minuteWednesday = howLongToSchool + practiceTime;
                minuteMonday = howLongToSchool;
                minuteTuesday = howLongToSchool;
                minuteThursday = howLongToSchool;
                minuteFriday = howLongToSchool;
            }
            else if (practiceDay == 4 && doYouPractice == false)
            {
                practiceThursday = sport;
                practiceMonday = "";
                practiceTuesday = "";
                practiceWednesday = "";
                practiceFriday = "";

                minuteThursday = howLongToSchool + practiceTime;
                minuteMonday = howLongToSchool;
                minuteTuesday = howLongToSchool;
                minuteWednesday = howLongToSchool;
                minuteFriday = howLongToSchool;
            }
            else if (practiceDay == 5 && doYouPractice == false)
            {
                practiceFriday = sport;
                practiceMonday = "";
                practiceTuesday = "";
                practiceWednesday = "";
                practiceThursday = "";

                minuteFriday = howLongToSchool + practiceTime;
                minuteMonday = howLongToSchool;
                minuteTuesday = howLongToSchool;
                minuteWednesday = howLongToSchool;
                minuteThursday = howLongToSchool;
            }

        }

        public void activityDay()
        {
            Random random = new Random();
            int activityDay = random.Next(1, 3);

            if (activityDay == 1)
            {
                if (doYouHaveGames == true)
                {
                    activitySaturday = sport + " Game";
                }
                else if (doYouHaveGames == false)
                {
                    activitySaturday = sport;
                }
                activitySunday = "";
            }
            else if (activityDay == 2)
            {
                if (doYouHaveGames == true)
                {
                    activitySaturday = sport + " Game";
                }
                else if (doYouHaveGames == false)
                {
                    activitySaturday = sport;
                }
                activitySunday = "";
            }
        }

        public void walkTime()
        {
            if(howLongToSchool <= 5){
                howLongToSchool = 10;
            }
            else if(howLongToSchool >= 35)
            {
                howLongToSchool = 30;
            }
            }
        

        public void addTime()
        {

        }
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            //////////////////////////////////////////
            //Adds data to the Basic Info Variables// 
            ////////////////////////////////////////
            
            fullName = textBox1.Text;
            stdntNmbr = textBox2.Text;
            grade = comboBox1.SelectedIndex;
            peClass = textBox3.Text;
            peTeacher = textBox4.Text;
            sport = textBox5.Text;
            howLongToSchoolDecimal = numericUpDown1.Value;

            howLongToSchool = (int)howLongToSchoolDecimal;


            if (radioButton1.Checked == true && radioButton2.Checked == false){
                doYouPractice = true;
                practiceOptionsChosen = 1;
            }
            else if (radioButton1.Checked == false && radioButton2.Checked == true)
            {
                doYouPractice = false;
                practiceOptionsChosen = 1;
            }

            if (radioButton3.Checked == true && radioButton4.Checked == false)
            {
                doYouHaveGames = true;
                gameOptionChosen = 1;
            }
            else if (radioButton3.Checked == false && radioButton4.Checked == true)
            {
                doYouHaveGames = false;
                gameOptionChosen = 1;
            }

            
            //Check variables for errors
            checkInfo();

           

            }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }
        
        private void button2_Click(object sender, EventArgs e)
        {

            AboutBox1 aboutBox = new AboutBox1();
            aboutBox.ShowDialog();
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            string practiceLabel = "Do you have " + textBox5.Text + " practices?";
            label7.Text = practiceLabel;

            if (textBox5.Text == "")
            {
                label7.Text = "Do you go to practices?";
            }

            string gameLabel = "Do you play " + textBox5.Text + " games?";
            label11.Text = gameLabel;

            if (textBox5.Text == "")
            {
                label11.Text = "Do you go to games?";
            }
            
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}

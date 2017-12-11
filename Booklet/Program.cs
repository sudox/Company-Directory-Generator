using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using TheArtOfDev.HtmlRenderer.PdfSharp;
using Excel = Microsoft.Office.Interop.Excel;

namespace Booklet
{
    class Program
    {
        public static List<string> departments = new List<string>();
        public static List<List<string>> names = new List<List<string>>();

        public static void getExcelData()
        {
            bool departmentName = true;
            Console.WriteLine("Loading names from book.xlsx...");
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Directory.GetCurrentDirectory() + "\\book.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            for (int i = 1; i <= rowCount; i++)
            {
                string movie = "";
                if(xlRange.Cells[i, 2].Value2 != null)
                {
                    movie = xlRange.Cells[i, 2].Value2.ToString();
                }
                if (departmentName)
                {
                    departments.Add(xlRange.Cells[i, 1].Value2.ToString());
                    departmentName = false;
                    i++;
                    i++;
                    List<string> temp = new List<string>();
                    if(xlRange.Cells[i, 2].Value2 != null)
                    {
                        movie = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    temp.Add(xlRange.Cells[i, 1].Value2.ToString() + " - " + movie);
                    names.Add(temp);
                    continue;
                }

                if (xlRange.Cells[i, 1].Value2 == null)
                {
                    departmentName = true;
                    continue;
                }
                //if (names[departments.Count - 1].Count == 8)
                //{
                //    departments.Add(departments.Last());
                //    List<string> temp = new List<string>();
                //    temp.Add(xlRange.Cells[i, 1].Value2.ToString() + " - " + xlRange.Cells[i, 2].Value2.ToString());
                //    names.Add(temp);
                //    continue;
                //}
                names[departments.Count - 1].Add(xlRange.Cells[i, 1].Value2.ToString() + " - " + movie);
            }
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Loaded.");
        }

        static void Main(string[] args)
        {
            getExcelData();
            //if (departments.Count % 2 == 0)
            //{
            //    for (int i = 0; i < departments.Count / 2; i++)
            //    {
            //        buildPage(departments[i], departments[departments.Count - i - 1], names[i], names[departments.Count - i - 1], "Page_" + (i + 1) + "_and_" + (departments.Count - i) + ".html");
            //    }
            //}
            //else
            //{
            //    for (int i = 0; i < departments.Count / 2; i++)
            //    {
            //        buildPage(departments[i], departments[departments.Count - i - 2], names[i], names[departments.Count - i - 2], "Page_" + (i + 1) + "_and_" + (departments.Count - i - 1) + ".html");
            //    }
            //    buildPage(departments.Last(), "", names[departments.Count - 1], new List<string>(), "Final Page.html");
            //}



            //List<string> pages = new List<string>();

            ////Title Page
            //pages.Add(File.ReadAllText("./title.html"));
            //for (int i = 0; i < departments.Count; i++)
            //{
            //    Console.WriteLine("Creating HTML Page " + (i + 1));
            //    if (names[i].Count > 8)
            //    {
            //        foreach (List<string> nameSubList in splitList(names[i]))
            //        {
            //            pages.Add(buildPdfHtml(departments[i], nameSubList));
            //        }
            //    }
            //    else
            //    {
            //        pages.Add(buildPdfHtml(departments[i], names[i]));
            //    }
            //}

            //Console.WriteLine("Done creating HTML Pages");

            //List<Process> processes = new List<Process>();
            //for (int i = 0; i < pages.Count; i++)
            //{
            //    Console.WriteLine("Creating PDF Page " + (i + 1));
            //    File.WriteAllText("./Page" + (i + 1) + ".html", pages[i]);
            //    ProcessStartInfo startInfo = new ProcessStartInfo();
            //    startInfo.FileName = "wkhtmltopdf.exe";
            //    startInfo.Arguments = "./Page" + (i + 1) + ".html" + " " + "./Page" + (i + 1) + ".pdf";

            //    processes.Add(Process.Start(startInfo));
            //}

            //foreach(var p in processes)
            //{
            //    p.WaitForExit();
            //}

            //Console.WriteLine("All PDF's Generated");
            //Console.WriteLine("Merging PDF's");
            //PdfDocument output = new PdfDocument();
            //for(int i = 0; i < pages.Count; i++)
            //{
            //    Console.WriteLine("Merging in Page " + (i + 1));
            //    using (PdfDocument temp = PdfReader.Open("Page" + (i + 1) + ".pdf", PdfDocumentOpenMode.Import))
            //    {
            //        output.AddPage(temp.Pages[0]);
            //    }
            //}

            //Console.WriteLine("Done. Outputing...");

            //output.Save("Booklet.pdf");

            ////Cleanup

            //Console.WriteLine("Booklet output, cleaning up");

            //for(int i = 0; i < pages.Count; i++)
            //{
            //    File.Delete("./Page" + (i + 1) + ".html");
            //    File.Delete("./Page" + (i + 1) + ".pdf");
            //}

            // Single page HTML testing
            //string html = "<html><head>\n";
            //html += "<style>.section{ clear: both; padding: 0px; margin: 0px; } .col { display: block; float:left; margin: 1% 0 1% 1.6%; } .col:first-child { margin-left: 0; } .span_2_of_2 { width:100%; text-align:center; } .span_1_of_2 { width: 49.2%; text-align:center; }  img { width:350px; height:250px; }</style>\n";
            //html += "</head>\n";
            //html += "<body style=\"zoom:85%\">\n";
            //for (int i = 0; i < departments.Count; i++)
            //{
            //    html += buildSingleHtml(departments[i], names[i]);
            //}
            //html += "</body></html>\n";
            //File.WriteAllText("singleHtmlTest.html", html);

            //ProcessStartInfo info = new ProcessStartInfo();
            //info.FileName = "wkhtmltopdf.exe";
            //info.Arguments = "singleHtmlTest.html test.pdf";

            //Process.Start(info);

            // Process Multi Department HTML Pages (only pass 8 or less pictures per group, never more than 2 departments)

            List<string> htmlPages = new List<string>();

            htmlPages.Add(File.ReadAllText("./title.html"));

            //Combine the name lists into a single list (this process can be optimized, when loading names load them in list, count names per department, store in another list)

            List<string> namesCombined = new List<string>();

            foreach (List<string> n in names)
            {
                foreach (string s in n)
                {
                    namesCombined.Add(s);
                }
            }

            //Build department index
            List<string> departmentsLoose = new List<string>();

            for (int i = 0; i < departments.Count; i++)
            {
                for (int j = 0; j < names[i].Count; j++)
                {
                    departmentsLoose.Add(departments[i]);
                }
            }

            //Build a list of 8 names from up to 2 departments

            while (namesCombined.Count > 0)
            {
                List<string> nameSet = new List<string>();
                List<string> departmentSet = new List<string>();
                int currentDepartmentCount = 0;
                string currentDepartment = "";

                for (int i = 0; i < 8; i++)
                {
                    if (namesCombined.Count <= 0)
                    {
                        break;
                    }
                    if (currentDepartment != departmentsLoose[0])
                    {
                        if (nameSet.Count % 2 != 0 && currentDepartmentCount > 0)
                        {
                            i++;
                            if (i >= 8)
                                break;
                        }
                        currentDepartment = departmentsLoose[0];
                        currentDepartmentCount++;
                        
                    }
                    if (currentDepartmentCount > 2)
                    {
                        break;
                    }

                    nameSet.Add(namesCombined[0]);
                    departmentSet.Add(departmentsLoose[0]);

                    namesCombined.RemoveAt(0);
                    departmentsLoose.RemoveAt(0);
                }

                htmlPages.Add(buildPdfHtmlMultiDepartment(departmentSet, nameSet));
            }

            Console.WriteLine("Done creating HTML Pages");

            List<Process> processes = new List<Process>();
            for (int i = 0; i < htmlPages.Count; i++)
            {
                Console.WriteLine("Creating PDF Page " + (i + 1));
                File.WriteAllText("./Page" + (i + 1) + ".html", htmlPages[i]);
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = "wkhtmltopdf.exe";
                startInfo.Arguments = "./Page" + (i + 1) + ".html" + " " + "./Page" + (i + 1) + ".pdf";

                processes.Add(Process.Start(startInfo));
            }

            foreach (var p in processes)
            {
                p.WaitForExit();
            }

            Console.WriteLine("All PDF's Generated");
            Console.WriteLine("Merging PDF's");
            PdfDocument output = new PdfDocument();
            for (int i = 0; i < htmlPages.Count; i++)
            {
                Console.WriteLine("Merging in Page " + (i + 1));
                using (PdfDocument temp = PdfReader.Open("Page" + (i + 1) + ".pdf", PdfDocumentOpenMode.Import))
                {
                    output.AddPage(temp.Pages[0]);
                }
            }

            Console.WriteLine("Done. Outputing...");

            output.Save("Booklet.pdf");

            //Cleanup

            Console.WriteLine("Booklet output, cleaning up");

            for (int i = 0; i < htmlPages.Count; i++)
            {
                File.Delete("./Page" + (i + 1) + ".html");
                File.Delete("./Page" + (i + 1) + ".pdf");
            }
        }

        public static void buildPage(string department1, string department2, List<string> names1, List<string> names2, string filename)
        {
            string s = "<html><head><link href=\"./booklet.css\" rel =\"stylesheet\" /></head>\n";
            s += "<body style=\"zoom:50%;\">\n";
            s += "<div class=\"section group\">\n";
            s += "<div class=\"col span_1_of_2\">\n";
            s += "<h1>" + department1 + "</h1>\n";
            s += "</div>\n";
            s += "<div class=\"col span_1_of_2\">\n";
            s += "<h1>" + department2 + "</h1>\n";
            s += "</div>\n";
            s += "</div>\n";

            for (int i = 0; i < 8; i++)
            {
                s += "<div class=\"section group\">\n";

                //for (int j = 0; i < 2; i++)
                //{
                if (names1.Count > i)
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "<img src=\"./" + department1 + "/" + names1[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names1[i] + "</h3>\n";
                    s += "</div>\n";
                }
                else
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "</div>\n";
                }
                i++;
                if (names1.Count > i)
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "<img src=\"./" + department1 + "/" + names1[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names1[i] + "</h3>\n";
                    s += "</div>\n";
                }
                else
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "</div>\n";
                }
                i--;
                if (names2.Count > i)
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "<img src=\"./" + department2 + "/" + names2[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names2[i] + "</h3>\n";
                    s += "</div>\n";
                }
                else
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "</div>\n";
                }
                i++;
                if (names2.Count > i)
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "<img src=\"./" + department2 + "/" + names2[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names2[i] + "</h3>\n";
                    s += "</div>\n";
                }
                else
                {
                    s += "<div class=\"col span_1_of_4\">\n";
                    s += "</div>\n";
                }
                //}
                s += "</div>\n";
            }
            s += "</body>\n</html>";

            File.WriteAllText(filename, s);
        }

        public static string buildPdfHtml(string department, List<string> names)
        {
            string s = "<html><head>\n";
            s += "<style>.section{ clear: both; padding: 0px; margin: 0px; } .col { display: block; float:left; margin: 1% 0 1% 1.6%; } .col:first-child { margin-left: 0; } .span_2_of_2 { width:100%; text-align:center; } .span_1_of_2 { width: 49.2%; text-align:center; }  img { width:350px; height:250px; }</style>\n";
            s += "</head>\n";
            s += "<body style=\"zoom:85%\">\n";
            s += "<div class=\"section group\">\n";
            s += "<div class=\"col span_2_of_2\">\n";
            s += "<h1>" + department + "</h1>\n";
            s += "</div>\n";
            s += "</div>\n";

            List<List<string>> nameSub = splitList(names);

            foreach (List<string> nameList in nameSub)
            {
                for (int i = 0; i < 8; i++)
                {
                    if (i >= nameList.Count)
                    {
                        if (i % 2 == 0)
                        {
                            s += "<div class=\"section group\">\n";
                            s += "<div class=\"col span_1_of_2\">\n";
                            s += "</div>";
                        }
                        else
                        {
                            s += "<div class=\"col span_1_of_2\">\n";
                            s += "</div>";
                            s += "</div>";
                        }
                        continue;
                    }
                    if (i % 2 == 0)
                    {
                        s += "<div class=\"section group\">\n";
                        s += "<div class=\"col span_1_of_2\">\n";
                        s += "<img src=\"./" + department + "/" + nameList[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                        s += "<h3>" + nameList[i] + "</h3>\n";
                        s += "</div>\n";
                    }
                    else
                    {
                        s += "<div class=\"col span_1_of_2\">\n";
                        s += "<img src=\"./" + department + "/" + nameList[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                        s += "<h3>" + nameList[i] + "</h3>\n";
                        s += "</div>\n";
                        s += "</div>\n";
                    }
                }
            }
            s += "</body>\n</html>";
            return s;
        }

        public static string buildSingleHtml(string department, List<string> names)
        {
            string s = "<div class=\"section group\">\n";
            s += "<div class=\"col span_2_of_2\">\n";
            s += "<h1>" + department + "</h1>\n";
            s += "</div>\n";
            s += "</div>\n";

            for (int i = 0; i < names.Count; i++)
            {
                if (i % 2 == 0)
                {
                    s += "<div class=\"section group\">\n";
                    s += "<div class=\"col span_1_of_2\">\n";
                    s += "<img src=\"./" + department + "/" + names[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names[i] + "</h3>\n";
                    s += "</div>\n";
                }
                else
                {
                    s += "<div class=\"col span_1_of_2\">\n";
                    s += "<img src=\"./" + department + "/" + names[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names[i] + "</h3>\n";
                    s += "</div>\n";
                    s += "</div>\n";
                }
            }

            //If an odd number of people in department add a blank to keep everything square
            if (names.Count % 2 != 0)
            {
                s += "<div class=\"col span_1_of_2\">\n";
                s += "</div>\n";
                s += "</div>\n";
            }

            return s;
        }

        public static string buildPdfHtmlMultiDepartment(List<string> departments, List<string> names)
        {
            bool columnTwo = false;
            string s = "<html><head>\n";
            s += "<style>.section{ clear: both; padding: 0px; margin: 0px; } .col { display: block; float:left; margin: 1% 0 1% 1.6%; } .col:first-child { margin-left: 0; } .span_2_of_2 { width:100%; text-align:center; } .span_1_of_2 { width: 49.2%; text-align:center; }  img { width:350px; height:250px; }</style>\n";
            s += "</head>\n";
            s += "<body style=\"zoom:85%\">\n";
            s += "<div class=\"section group\">\n";
            s += "<div class=\"col span_2_of_2\">\n";
            s += "<h1>" + departments[0] + "</h1>\n";
            s += "</div>\n";
            s += "</div>\n";

            string currentDepartment = departments[0];

            for (int i = 0; i < names.Count; i++)
            {
                if (departments[i] != currentDepartment)
                {
                    if (columnTwo)
                    {
                        s += "<div class=\"col span_1_of_2\">\n";
                        s += "</div>\n";
                        s += "</div>\n";
                        columnTwo = false;
                    }
                    s += "<div class=\"section group\">\n";
                    s += "<div class=\"col span_2_of_2\">\n";
                    s += "<h1>" + departments[i] + "</h1>\n";
                    s += "</div>\n";
                    s += "</div>\n";
                    currentDepartment = departments[i];
                }

                if (columnTwo)
                {
                    s += "<div class=\"col span_1_of_2\">\n";
                    s += "<img src=\"./" + departments[i] + "/" + names[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names[i] + "</h3>\n";
                    s += "</div>\n";
                    s += "</div>\n";
                    columnTwo = false;
                }
                else
                {
                    s += "<div class=\"section group\">\n";
                    s += "<div class=\"col span_1_of_2\">\n";
                    s += "<img src=\"./" + departments[i] + "/" + names[i].Split(',')[0].Split('-')[0].TrimEnd(' ') + ".jpg\" />";
                    s += "<h3>" + names[i] + "</h3>\n";
                    s += "</div>\n";
                    columnTwo = true;
                }
            }

            s += "</body></html>\n";

            return s;
        }

        public static List<List<string>> splitList(List<string> locations, int nSize = 8)
        {
            var list = new List<List<string>>();

            for (int i = 0; i < locations.Count; i += nSize)
            {
                list.Add(locations.GetRange(i, Math.Min(nSize, locations.Count - i)));
            }

            return list;
        }
    }
}

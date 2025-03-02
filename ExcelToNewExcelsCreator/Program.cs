using OfficeOpenXml;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;

namespace ExcelToNewExcelsCreator
{

    internal class Program
    {
        static void Main(string[] args)
        {
            //Package licence
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Paths
            string baseFilesDirectoryPath = GetPathToDirectory_Files() + @"\BaseFiles";
            string newFilesDirectoryPath = baseFilesDirectoryPath + @"\NewFiles";

            //////////Program//////////
            // 0. Ask User
            Console.WriteLine("Write below how many new files create");
            int amountOfFiles = AskUserAboutAmount();

            if (amountOfFiles == 0)
            {
                Console.WriteLine("You chose 0 new files");
                CloseApp();
            }

            Console.WriteLine("Write below how many carriers per day");
            int carriersPerDay = AskUserAboutAmount();
            if (carriersPerDay == 0)
            {
                Console.WriteLine("Zero? Are you made all of this in one day? Ok here we go");
            }

            // 1. Create new directory for files
            CreateNewDirectory(newFilesDirectoryPath);


            // 3.      //amount of values (for set array) to read form file

            ChangingFiles(carriersPerDay, amountOfFiles, baseFilesDirectoryPath, newFilesDirectoryPath);




            // 7. Info: where files are
            Console.WriteLine($"\nFiles saved in folder: {newFilesDirectoryPath}");
            Console.WriteLine("Press any key to close the window");
            Console.ReadKey();

            return;
        }
        enum ColorEnumMessage
        {
            yellow,
            red,
        }
        static void ColoredMesage(ColorEnumMessage color, string message)
        {

            switch (color)
            {
                case ColorEnumMessage.yellow:

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;

                case ColorEnumMessage.red:

                    Console.ForegroundColor = ConsoleColor.Red;
                    break;

                default:

                    Console.ResetColor();
                    break;
            }

            Console.WriteLine(message);
            Console.ResetColor();
        }
        static int[,] WitchCellsToChange()
        {
            int[,] cellPostitionXY = {
                { 6, 3  },  //Device_number
                { 7, 11 },  //IP (last octet) - Device_1
                { 7, 13 },  //IP (last octet) - Device_2
                { 6, 23 }   //Date
            };

            return cellPostitionXY;
        }
        static void CreateNewDirectory(string newDirectoryPath)
        {
            //New directory for files
            try
            {
                if (Directory.Exists(newDirectoryPath))
                {
                    Directory.Delete(newDirectoryPath, true); //true, give permision to delete directory and all content
                    Directory.CreateDirectory(newDirectoryPath);
                }
                else
                {
                    Directory.CreateDirectory(newDirectoryPath);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                Debug.WriteLine("Catch in CreateNewDirectory -> UnauthorizedAccessException");

                Console.WriteLine("No permition to delete directory:\n" + newDirectoryPath);
                CloseApp();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Can't delete or create new directory");
                Console.WriteLine($"Close all files from directory:\n + {newDirectoryPath}");
                Console.ResetColor();

                CloseApp();
            }
        }
        static string FindFileWithExtension(string directoryPath, string extension)
        {
            string[] path = null;
            try
            {
                path = Directory.GetFiles(directoryPath, extension);
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine("Directory not found:\n" + directoryPath);
                CloseApp();
            }
            catch 
            {
                Console.WriteLine("Problems with found file with extension: " + extension);
                CloseApp();
            }
            
            if (path.Length == 0)
            {
                CloseApp();
            }
            else
            {
                for (int i = 0; i < path.Length; i++)
                {
                    if ('~' != path[i][0]) //Checking if ther is temporary file"~"
                    {
                        return path[i];
                    }
                }
            }

            return "";
        }
        static string GetPathToDirectory_Files()
        {
            string path = new DirectoryInfo(".").FullName;
            int ile = path.IndexOf("bin") - 1;
            if (ile < 0)
            {
                Console.WriteLine("Director not fount, please contact with your IT department");
                CloseApp();
            }
            else
            {
                path = path.Substring(0, ile);
                path = path + @"\Files";
            }
            return path;
        }
        static void CloseApp([CallerMemberName] string methodName = "")
        {
            Debug.WriteLine($"!!!There was a problem in method: {methodName}");

            Console.WriteLine("\nPress Enter to close app");
            Console.ReadKey();
            Environment.Exit(0);
        }
        static int AskUserAboutAmount()
        {
            Console.Write("Amount: ");
            try
            {
                return Int32.Parse(Console.ReadLine());
            }
            catch (FormatException)
            {
                ColoredMesage(ColorEnumMessage.yellow, "Error: not valid integer value");
                CloseApp();
            }
            catch
            {
                ColoredMesage(ColorEnumMessage.red, "!!!Error durring reading value from keyboard");
                CloseApp();
            }

            return 0;
        }
        static void ChangingFiles(int carriersPerDay, int amountOfFiles, string directoryPath, string newDirectoryPath)
        {
            int[,] cellsPositionsXY = WitchCellsToChange();

            //Variable without values
            int amountOfValues = cellsPositionsXY.Length / 2;
            string[] value = new string[amountOfValues];    //amount of value to change

            string filePath = FindFileWithExtension(directoryPath, "*.xlsx");   //bellow try for this variable
            string fileName = filePath.Substring(directoryPath.Length + 1);     //file name (cut path)
            string newFilePath = newDirectoryPath + $@"\{fileName}";             //coppy file in new directory

            // 1. Check file name
            try
            {
                fileName.Substring(fileName.IndexOf("-") - 1, fileName.Length - fileName.IndexOf("-") + 1);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(
                    "Error" + "\n" +
                    "Check name of oryginal file. Should start like \"CAxxxx - ...\"" + "\n" +
                    "For example: \"CA0001 - RoDipE carrier precommissioning Check list v1.8\" "
                    );
                Console.ResetColor();

                //Open excel -> WorkSheet -> read one value and write to variable
                value[0] = new ExcelPackage(filePath).Workbook.Worksheets[2].GetValue(cellsPositionsXY[0, 0], cellsPositionsXY[0, 1]).ToString();
                fileName = $"CA{value[0]} - {fileName}";

                newFilePath = newDirectoryPath + $@"\{fileName}";
                Console.WriteLine("Changing name of first file on: " + fileName);
            }

            // 3.1. Info:
            Console.WriteLine($"\nCreated copy for: CA{value[0]}");

            // 2. Operate on excel file 
            File.Copy(filePath, newFilePath);                   //Copy Excel to new diercotry
            ExcelPackage excel = new ExcelPackage(newFilePath); //Create instance for excel
            ExcelWorksheet ws = excel.Workbook.Worksheets[2];   //Create instance for worksheet in excel

            int ipDevice_1;                                     // First part of IP for WLAN
            int ipDevice_2;                                      // First part of IP for PHC BOX
            string nameZeros;                               // First part of Name
            int nameNumber;                                 // Second part of Name

            //for (int i = 0; i < amountOfValues; i++)
            for (int i = 0; i < 4; i++)
            {
                value[i] = ws.GetValue(cellsPositionsXY[i, 0], cellsPositionsXY[i, 1]).ToString();   //GetValue() - takes value form cell

                //value[0] = 0001;          //Device_number
                //value[1] = 192;           //IP (last octet) - Device_1
                //value[2] = 102;           //IP (last octet) - Device_2
                //value[3] = dd.mm.yyyy;    //Date       
            }


            // 4. Prepare data
            // 4.1. Name for new file
            nameNumber = Int32.Parse(value[0]);
            nameZeros = value[0].Substring(0, (value[0].Length - 2));

            // 4.2. Conversion (casting) string to int
            ipDevice_1 = Int32.Parse(value[1]);
            ipDevice_2 = Int32.Parse(value[2]);

            // 4.3. Creating new variable "date" of type "DateTime". Parse string to DateTime
            DateTime.TryParse(value[3], out DateTime date);

            // 4.4. Edit file name
            Console.WriteLine(fileName);
            fileName = fileName.Substring(fileName.IndexOf("-") - 1, fileName.Length - fileName.IndexOf("-") + 1);

            if (carriersPerDay == 0)
            {
                date = date.AddDays(-1);
            }
            // 5. Changing data and create new excel
            for (int i = 1; i <= amountOfFiles; i++)
            {
                // 5.1. Change Name:
                if (nameNumber + i < 10) { value[0] = nameZeros + "0" + (nameNumber + i); }
                else { value[0] = nameZeros + (nameNumber + i); }

                // 5.2. Change and prepare IP:
                value[1] = (ipDevice_1 + i).ToString();   //WLAN
                value[2] = (ipDevice_2 + i).ToString();    //PHC

                // 5.3. Change date
                if ((i - 1) % carriersPerDay == 0)
                {
                    date = date.AddDays(1);
                }

                if (6 == (int)date.DayOfWeek)       //if Saturday
                {
                    date = date.AddDays(2);
                }
                else if (0 == (int)date.DayOfWeek)  //if Sunday 
                {
                    date = date.AddDays(1);
                }
                value[3] = date.ToString("d");      //"d" change to string but in data structure

                // 5.4. Write values in WorkSheet into correct cells (not saved yet)
                for (int j = 0; j < 4; j++)
                { ws.Cells[cellsPositionsXY[j, 0], cellsPositionsXY[j, 0]].Value = value[j]; }

                // 5.5. Save content to excel file  
                File.WriteAllBytes(newFilePath, excel.GetAsByteArray());

                // 5.6. Copy file and save with new name and changeds
                File.Copy(newFilePath, $@"{newDirectoryPath}\CA{value[0]}{fileName}");

                // 5.7. Info: new file created
                Console.WriteLine($"New excel for: CA{value[0]} is ready, with date: {date.ToShortDateString()}");
            };


            // 6. Clear all
            // 6.1. Close instance of excel (close file)
            excel.Dispose();                    //Clsoe excel instance

            // 6.2. Correcting the first copy
            File.Delete(newFilePath);           //Delete modified copy
            File.Copy(filePath, newFilePath);   //Create copy form orginal to new diercotry
        }
    }
}

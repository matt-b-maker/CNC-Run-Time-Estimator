using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using EPDM.Interop.epdm;
using OfficeOpenXml;
using SolidWorks.Interop.sldworks;

namespace WPFSWTry
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async void Button_Click(object sender, RoutedEventArgs e)
        {
            BoxBoi.Text = "";

            if (Prodnum.Text == "")
            {
                MessageBox.Show("You have to enter a dang PROD number, ya ding dang donkey.");
                return;
            }

            #region Define things (interfaces / beginning variables)

            //Get the progress bar going
            ProgBoi.Visibility = Visibility.Visible;
            Robot.Visibility = Visibility.Visible;
            ProgBoi.IsIndeterminate = true;

            #endregion

            string prodNum = Prodnum.Text;
            string userProdNum = "PROD-" + prodNum;

            #region Get Material Counts

            BoxBoi.Text += "\nGetting materials from PDM";
            //Get those sweet materials
            (var resultsMaterial, var cncFiles, var cncPath) = await Task.Run(() => GetMaterialList(prodNum));

            if (cncPath == null)
            {
                return;
            }

            List<OfficialBomItem> officialBomItems = new List<OfficialBomItem>();

            //Report number of material types found and populate official item object list with materials
            BoxBoi.Text += $"\n{resultsMaterial.Count} material types found";

            foreach (var materialType in resultsMaterial)
            {
                officialBomItems.Add(new OfficialBomItem(materialType.Description, materialType.Thickness, materialType.Quantity));
            }

            BoxBoi.Text += "\nCalculating run times of each CNC file.";

            #endregion

            #region Get the CNC Run Times

            //Get the estimated file run times if there are files in the pdm
            var fileRunTimes = await Task.Run(() => GetFileRunTimes(cncFiles, cncPath));

            if (fileRunTimes == null)
            {
                MessageBox.Show("You have to pull local copies of these files. Otherwise, this will never work for either of us. Go on and do that now");
                Close();
            }

            #endregion

            #region Make excel doc (Currently commented out)
            
            #region Generate the Excel Spreadsheet

            //Get the excel bom going now
            ExcelCreator excelBoi = new ExcelCreator(userProdNum, prodNum, officialBomItems, fileRunTimes);
            FileInfo file = MakeExcelDoc(excelBoi, prodNum, cncPath);

            try
            {
                if (fileRunTimes != null)
                {
                    await excelBoi.SaveExcelFile(excelBoi.BomItems, file);
                }
            }
            catch
            {
                MessageBox.Show("Something went wrong with the excel file");
            }

            BoxBoi.Text += "\n\nExcel document created at " + GetExcelPath(userProdNum) + "\\";

            ProgBoi.IsIndeterminate = false;
            Robot.Text = "DONE.";

            if ((bool)Checkboi.IsChecked)
            {
                excelBoi.OpenExcelFile(file);
            }

            #endregion

            #endregion

            foreach (var material in officialBomItems)
            {
                if (material.Quantity == 1)
                {
                    BoxBoi.Text += '\n' + $"{material.Quantity} sheet of {material.Description} {material.Thickness}";
                }
                else
                {
                    BoxBoi.Text += '\n' + $"{material.Quantity} sheets of {material.Description} {material.Thickness}";
                }
            }

            int totalSeconds = 0;

            foreach (var runTime in fileRunTimes)
            {
                totalSeconds += runTime.Seconds;
            }

            BoxBoi.Text += '\n' + totalSeconds.ToString() + " total seconds to run";

            TimeSpan t = TimeSpan.FromSeconds(totalSeconds);
            string totalRunTime = t.ToString(@"hh\:mm\:ss");
            BoxBoi.Text += '\n' + totalRunTime;

            ProgBoi.IsIndeterminate = false;
            Robot.Text = "DONE.";

            #region Tidy up the dang place

            #endregion
        }

        #region CNC Runtime Methods

        private List<RunTime> GetFileRunTimes(List<string> cncFiles, string cncPath)
        {
            List<int> runTimes = new List<int>();
            List<string> filePaths = new List<string>();
            List<string[]> files = new List<string[]>();

            //Now, do the PDM search
            IEdmVault21 CurrentVault = new EdmVault5() as IEdmVault21;
            //Log in to the PDM

            CurrentVault.LoginAuto("CreativeWorks", 0);

            //Set up search functionality
            IEdmSearchResult5 _searchResult;
            IEdmSearch9 _search = (IEdmSearch9)CurrentVault.CreateSearch2();
            _search.FindFolders = true;

            //Test search functionality
            string[] VarNames0 = { };
            _search.AddMultiVariableCondition(VarNames0, "@:"); // poVariableNames can be null
            _search.GetFirstResult();
            bool OriginatedFromCreateSearch2 = _search.GetSyntaxErrors() != null;

            //Define the search path
            //bool ExceptionEncountered = false;
            _search.Clear();
            _search.StartFolderID = CurrentVault.GetFolderFromPath(cncPath + "\\Programs").ID;
            //string folder = _search.StartFolderID.ToString();

            int testCount = 0;

            _search.FileName = cncFiles[0];
            _searchResult = _search.GetFirstResult();

            if (_searchResult != null)
            {
                filePaths.Add(_searchResult.Path);
                for (int i = 0; i < cncFiles.Count - 1; i++)
                {
                    _search.FileName = cncFiles[i + 1];
                    _searchResult = _search.GetFirstResult();
                    filePaths.Add(_searchResult.Path);
                    testCount++;
                }
            }
            
            try
            {
                foreach (var file in filePaths)
                {
                    files.Add(File.ReadAllLines(file));
                }
            }
            catch
            {
                filePaths = null;
            }

            List<RunTime> runTimeObjects = new List<RunTime>();

            if (filePaths == null)
            {
                runTimeObjects = null;
                return runTimeObjects;
            }

            runTimeObjects = CalculateRunTime(files, runTimeObjects, filePaths);

            return runTimeObjects;
        }

        private List<RunTime> CalculateRunTime(List<string[]> files, List<RunTime> runTimeObjects, List<string> filePaths)
        {
            //Point and move speed variables
            double moveSpeedXY = 0, moveSpeedZ = 0;
            /* jogSpeed will change once user input is accepted in the program*/
            double jogSpeed = 10;
            //These variable will represent the coordinates of the tool's previous position
            double x2 = 0, y2 = 0, z2 = 0;
            //This variable is to ensure that the first function in the conditional below will only happen once
            var zCount = 0;

            //Calculation Variables
            double runTime = 0;
            double totRunTime = 0;

            //Used to separate the strings
            char[] separators = new char[] { ',', ' ' };
            int nameCount = 0;

            //Go through each file
            foreach (string[] file in files)
            {
                //x and y will always start at zero, so at the beginning of each iteration (each file), they will be
                //set to zero
                double x1 = 0;
                double y1 = 0;

                //The beginning of shopbot files will always set the safe z height. Once the iterations find the 
                //line with safez in it, it will assign the value on the line to this variable
                double z1 = 0;
                foreach (string line in file)
                {
                    double distance;
                    if (line.StartsWith("MS"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        moveSpeedXY = Convert.ToDouble(subs[1]);
                        moveSpeedZ = Convert.ToDouble(subs[2]);
                        //Console.WriteLine($"{subs[0]} {subs[1]} {subs[2]}");
                    }
                    else if (line.StartsWith("JZ"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        if (zCount == 0)
                        {
                            z1 = Convert.ToDouble(subs[1]);
                            zCount++;
                        }
                        else
                        {
                            z2 = Convert.ToDouble(subs[1]);
                            if (z1 == z2)
                            {
                                continue;
                            }
                            else if (z1 > z2)
                            {
                                distance = Math.Abs(z1 - z2);
                                runTime += GetTime(distance, moveSpeedZ);
                                z1 = z2;
                            }
                            else if (z2 > z1)
                            {
                                distance = Math.Abs(z2 - z1);
                                runTime += GetTime(distance, moveSpeedZ);
                                z1 = z2;
                            }
                        }
                    }
                    else if (line.StartsWith("M3"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        //Console.WriteLine($"{subs[0]} {subs[1]} {subs[2]} {subs[3]}");
                        x2 = Convert.ToDouble(subs[1]);
                        y2 = Convert.ToDouble(subs[2]);
                        z2 = Convert.ToDouble(subs[3]);
                        if (IsZ(x1, y1, x2, y2))
                        {
                            distance = Math.Abs(z2 - z1);
                            runTime += GetTime(distance, moveSpeedZ);
                        }
                        else
                        {
                            distance = GetDistance(x1, y1, z1, x2, y2, z2);
                            runTime += GetTime(distance, moveSpeedXY);
                        }
                        x1 = x2;
                        y1 = y2;
                        z1 = z2;
                    }
                    else if (line.StartsWith("J3"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        //Console.WriteLine($"{subs[0]} {subs[1]} {subs[2]} {subs[3]}");
                        x2 = Convert.ToDouble(subs[1]);
                        y2 = Convert.ToDouble(subs[2]);
                        z2 = Convert.ToDouble(subs[3]);
                        if (IsZ(x1, y1, x2, y2))
                        {
                            distance = Math.Abs(z2 - z1);
                            runTime += GetTime(distance, moveSpeedZ);
                        }
                        else
                        {
                            distance = GetDistance(x1, y1, z1, x2, y2, z2);
                            runTime += GetTime(distance, moveSpeedXY);
                        }
                        x1 = x2;
                        y1 = y2;
                        z1 = z2;
                    }
                    else if (line.StartsWith("CG"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        //double startX = 0, startY = 0, endX = 0, endY = 0, xOffset = 0, yOffset = 0;
                        double startX = x1;
                        double startY = y1;
                        double endX = Convert.ToDouble(subs[1]);
                        double endY = Convert.ToDouble(subs[2]);
                        double xOffset = Convert.ToDouble(subs[3]);
                        //CG variables
                        double yOffset = Convert.ToDouble(subs[4]);

                        distance = GetArcLength(startX, startY, endX, endY, xOffset, yOffset);

                        runTime += GetTime(distance, moveSpeedXY);

                        x1 = endX;
                        y1 = endY;
                    }
                    else if (line.StartsWith("JH"))
                    {
                        x1 = 0;
                        y1 = 0;
                        distance = GetDistance(x1, y1, z1, x2, y2, z2);
                        runTime += GetTime(distance, moveSpeedXY);
                        x1 = x2;
                        y1 = y2;
                        z1 = z2;
                    }
                    else if (line.Contains("SafeZ"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        z1 = Convert.ToDouble(subs[subs.Length - 1]);
                    }
                    else if (line.StartsWith("END"))
                    {
                        break;
                    }
                }

                //Add current run time to total
                totRunTime += runTime;

                //Add to the list of objects to return
                runTimeObjects.Add(new RunTime(filePaths[nameCount], (int)runTime));

                //Sets it so the next name received for the current one being processed is one over
                nameCount++;

                //Resets runtime to 0
                runTime = 0;
            }

            return runTimeObjects;
        }

        static double GetArcLength(double x1, double y1, double x2, double y2, double xOffset, double yOffset)
        {
            double r, d, c1, c2, theta, arcLength;
            c1 = x1 + xOffset;
            c2 = y1 + yOffset;

            r = Math.Sqrt(Math.Pow((c1 - x1), 2) + Math.Pow((c2 - y1), 2));
            d = Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2));
            theta = Math.Acos(((2 * Math.Pow(r, 2)) - Math.Pow(d, 2)) / (2 * Math.Pow(r, 2)));
            arcLength = r * theta;

            return arcLength;
        }

        //For calculating time
        static double GetTime(double distance, double speed)
        {
            double time = (distance / speed);
            return Math.Abs(time);
        }

        //Get Angle between delta directions
        static double GetDistance(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            double distance = Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2) + Math.Pow((z2 - z1), 2));
            return Math.Abs(distance);
        }

        //To determine whether Z is the only coord to change
        static bool IsZ(double x1, double y1, double x2, double y2)
        {
            double x = x2 - x1;
            double y = y2 - y1;
            if (x == 0 && y == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region Main Methods 

        private (List<BomMaterialItem>, List<string>, string) GetMaterialList(string prodNum)
        {
            //Now, do the PDM search
            IEdmVault21 CurrentVault = new EdmVault5() as IEdmVault21;
            //Log in to the PDM

            CurrentVault.LoginAuto("CreativeWorks", 0);

            //Set up search functionality
            IEdmSearchResult5 _searchResult = null;
            IEdmSearch9 _search = (IEdmSearch9)CurrentVault.CreateSearch2();
            _search.FindFiles = true;

            //Define the search path
            bool ExceptionEncountered = false;
            bool multipleCncFiles = false;
            _search.Clear();
            _search.StartFolderID = CurrentVault.GetFolderFromPath("C:\\CreativeWorks").ID;
            //string folder = _search.StartFolderID.ToString();

            // Set up the Regex
            string cncSnipNamePattern = @"\\(PROD-)\d+ (CNC).SLDASM";
            string nestSnipNamePattern = @"\\(PROD-)\d+ (NEST).SLDASM";

            //Set up the lists to hold folders, files, and BomItem objects
            List<IEdmSearchResult5> prodCncFiles = new List<IEdmSearchResult5>();
            List<IEdmSearchResult5> shopBotFiles = new List<IEdmSearchResult5>();
            List<BomMaterialItem> bomMaterialItems = new List<BomMaterialItem>();
            List<string> cncFiles = new List<string>();
            string cncPath = null;

            //Run the search
            SearchForCNC(_search, prodNum);

            //Check for exception
            if (ExceptionEncountered)
            {
                _searchResult = null;
            }
            else
            {
                _searchResult = _search.GetFirstResult();
            }

            //Iterate through all found files and add them to the list of potentials
            if (_searchResult != null)
            {
                while (_searchResult != null)
                {
                    prodCncFiles.Add(_searchResult);
                    _searchResult = _search.GetNextResult();
                }
            }

            //Find the CNC folder in the search results and add all of the contained sbp files, if any, to a list
            if (prodCncFiles.Count > 0)
            {
                foreach (var file in prodCncFiles)
                {
                    if (file.Name.Contains("CNC") || file.Name.Contains("NEST"))
                    {
                        if (!multipleCncFiles)
                        {
                            cncPath = SnipPath(file.Path, cncSnipNamePattern, nestSnipNamePattern);

                            _search.Clear();
                            try
                            {
                                _search.StartFolderID = CurrentVault.GetFolderFromPath(cncPath).ID;
                            }
                            catch
                            {
                                Console.WriteLine("That's not a valid path");
                            }

                            _search.FileName = "sbp";
                            _searchResult = _search.GetFirstResult();

                            if (_searchResult == null)
                            {
                                MessageBox.Show("\nEither there are no ShopBot files in here or you didn't pull local copies of the files in the CNC folder. Do that and try again.");
                            }
                            else
                            {
                                while (_searchResult != null)
                                {
                                    if (!_searchResult.Path.Contains("Parts") && !_searchResult.Path.Contains("Recuts"))
                                    {
                                        shopBotFiles.Add(_searchResult);
                                        cncFiles.Add(_searchResult.Name);
                                        _searchResult = _search.GetNextResult();
                                    }
                                    else
                                    {
                                        _searchResult = _search.GetNextResult();
                                    }
                                }
                            }
                            multipleCncFiles = true;
                        }
                        else
                        {
                            MessageBox.Show("\n\nThere are multiple CNC files in here. Might want to check on that");
                        }
                    }
                }
            }


            //Check file names to determine which ones are part of a multiple part file
            //then begin populating the BomItem object list
            (string material, string thickness) = (null, null);
            (string tempMaterial, string tempThickness) = (null, null);
            int typeCount = 0;

            if (shopBotFiles.Count > 0)
            {
                foreach (var file in shopBotFiles)
                {
                    string pat = "[2-9].sbp";
                    Regex rgx = new Regex(pat);
                    bool isMultiPart = rgx.IsMatch(file.Name);
                    if (!isMultiPart)
                    {
                        (material, thickness) = GetDescription(file.Name);
                        if (material != tempMaterial || thickness != tempThickness)
                        {
                            typeCount++;
                            bomMaterialItems.Add(new BomMaterialItem(material, thickness));
                            bomMaterialItems[typeCount - 1].Quantity = 1;
                            tempMaterial = material;
                            tempThickness = thickness;
                        }
                        else
                        {
                            bomMaterialItems[typeCount - 1].Quantity++;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("\nThere were no shopbot files");
                cncPath = null;
                return (bomMaterialItems, cncFiles, cncPath);
            }

            return (bomMaterialItems, cncFiles, cncPath);

        }

        public string GetProdNumFromPath(string userProdNum) // For the SOLIDWORKS ALGO
        {
            string[] sections = userProdNum.Split('\\');
            string path;
            foreach (string section in sections)
            {
                if (section.Contains("PROD"))
                {
                    string pattern = @"PROD-";
                    path = Regex.Replace(section, pattern, "");
                    return path;
                }
            }
            path = "nope";
            return path;
        }

        public void RenameStuff(ModelDoc2 swModel) //For the SOLIDWORKS ALGO
        {
            Feature swFeature;
            swFeature = (Feature)swModel.FeatureByPositionReverse(3);
            swFeature.Name = "Good";

            swFeature = (Feature)swModel.FeatureByPositionReverse(2);
            swFeature.Name = "Job";

            swFeature = (Feature)swModel.FeatureByPositionReverse(1);
            swFeature.Name = "Brooo";
        }

        //This method separates the filename string into parts and gets the most relevant information out of it
        //There are a few exceptions that are dealt with using the if/else if stuff in there
        private static (string material, string thickness) GetDescription(string name)
        {
            string materialName = null;
            string thickness = null;
            string wordPattern = @"[A-Z,a-z]+";
            string thickness1 = @"1.sbp";
            string thickness2 = @"\d+$";

            Regex rgxName = new Regex(wordPattern);
            Regex rgxThick = new Regex(thickness2);

            MatchCollection words = rgxName.Matches(name);

            if (words[0].Value == "NO")
            {
                materialName += words[2];
            }
            else
            {
                materialName += words[0];
            }

            thickness = Regex.Replace(name, thickness1, "");
            thickness = rgxThick.Match(thickness).Value;

            switch (materialName)
            {
                case "Cheap":
                    materialName = "Cheap Plywood";
                    thickness = "18mm(.7\")";
                    break;
                case "bendaply":
                    materialName = "Bendaply";
                    thickness = "375";
                    break;
                case "Plywood":
                    materialName = "Bendaply";
                    thickness = "375";
                    break;
                case "Birch":
                    materialName = "Birch Plywood";
                    thickness = "18mm(.7\")";
                    break;
                case "mm":
                    materialName = "Laminate";
                    thickness = "3mm";
                    break;
                case "Diffusion":
                    materialName = "Diffusion Film";
                    thickness = "0325";
                    break;
                case "NO":
                    materialName = "Cheap plywood";
                    thickness = "18mm, check this. It was a no margin material";
                    break;
                default:
                    break;
            }

            return (materialName, thickness);
        }//For the PDM search ALGO

        //Searches for the first directory
        private static void SearchForCNC(IEdmSearch9 _search, string prodNum)
        {
            _search.FileName = "PROD-" + prodNum + " CNC.SLDASM";
        }//For the PDM search ALGO

        //Cuts path down to the folder 
        private static string SnipPath(string path, string cncPattern, string nestPattern)
        {
            if (path.Contains("CNC") && !path.Contains("NEST"))
            {
                path = Regex.Replace(path, cncPattern, "");
            }
            else if (path.Contains("NEST") && path.Contains("CNC"))
            {
                path = Regex.Replace(path, nestPattern, "");
            }
            return path;
        }//For the PDM search ALGO

        private string GetExcelPath(string fileName)
        {
            string pattern = @"2-CNC+";
            string path = Regex.Replace(fileName, pattern, "");
            path += "3-Build Sheet & BOM";
            return path;
        }

        private FileInfo MakeExcelDoc(ExcelCreator excelBoi, string prodNum, string fileName)
        {
            string path = GetExcelPath(fileName);
            FileInfo file = excelBoi.CreateExcelFile(path, prodNum);
            return file;
        }

        #endregion
    }
}


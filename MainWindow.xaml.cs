using System;
using System.Collections.Generic;
using System.Configuration;
using System.ComponentModel;
using System.IO;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Window = System.Windows.Window;
using Microsoft.Win32;
using System.Windows.Threading;

namespace Custom_Report_Debugger_2._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string filePath = null;
        public List<String> mergefieldsInDoc = new List<string>();
        public List<String> failedMergefields_Hierarchy = new List<string>();
        public List<String> failedMergefields_MissingBegin = new List<string>();
        public List<String> failedMergefields_MissingEnd = new List<string>();
        public List<String> failedMergefields_doNotExist = new List<string>();
        public string logFile = "";
        public MainWindow()
        {
            InitializeComponent();
        }

        //Load Button
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            mergefieldsInDoc.Clear();
            OpenFileDialog openFileDialog = new OpenFileDialog();                                                       //Opens File Explorer
            openFileDialog.Filter = "docx files (*.docx)|*.docx|All files (*.*)|*.*";                                   //Allows file drop down to include .docx and All files
            if (openFileDialog.ShowDialog() == true)                                                                    //If a document is selected then grab it's filepath, otherwise exit the method
            {
                filePath = openFileDialog.FileName;
            }
            else
            {
                return;
            }

            if (filePath.Contains(".docx") == true)                                                                     //Checks it's a .docx file
            {
                previewDescription.Text = "Loading, please wait...";
                previewDescription.Visibility = Visibility.Visible;
                Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);
                textToTest.Document.Blocks.Clear();                                                                     //Clears test data preview
                documentSelectionGuidance.Foreground = Brushes.DarkGreen;                                               //Prints a green message confirming successful doc selection
                documentSelectionGuidance.Content = filePath + " has been selected";
                Application app = new Application();                                                                    //Opens the document at that file path, and takes all contents and places it in a string
                Document doc = app.Documents.Open(filePath);
                string testDocContent = doc.Content.Text.Replace("  ", " ").Replace("   ", " ");
                doc.Close();
                app.Quit();

                var exceptionInFileUse = true;                
                (mergefieldsInDoc, exceptionInFileUse) = TestingFields.TestData(testDocContent);                        //Passes the full string through the method that converts the string to individual fields in a list
                if (exceptionInFileUse == true)
                {
                    previewDescription.Text = "There appears to have been an exception thrown whilst using this document";
                    previewDescription.Visibility = Visibility.Visible;
                }
                else
                {
                    previewDescription.Text = "The below mergefields were taken from the selected file:";
                    previewDescription.Visibility = Visibility.Visible;

                    foreach (var field in mergefieldsInDoc)                                                             //Displays every mergefield in the list
                    {
                        var example = field.Trim();
                        textToTest.AppendText(example);
                        textToTest.AppendText(Environment.NewLine);
                    }
                    SelectFile_Button.Content = "Change File";
                    confirmation_button.Visibility = Visibility.Visible;
                    formatting_description.Visibility = Visibility.Visible;
                }
            }
            else
            {
                documentSelectionGuidance.Foreground = Brushes.Red;                                                     //Errors if a non .docx file is used
                documentSelectionGuidance.Content = "This file is not compatible, please select a file in the .docx format";
            }
            return;
        }


        //Confirmation Button
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            selectAndLoad_Grid.Visibility = Visibility.Hidden;
            progress_Grid.Visibility = Visibility.Visible;
            processingStatus_Text.AppendText(DateTime.Now.ToString("HH'.'mm'.'ss") + ": Creating Log File");
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);

            //Creates a Log file - and folder if one doesn't exist
            var installLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Custom Report Debugger Logs";
            Directory.CreateDirectory(installLocation);
            logFile = installLocation + "\\" + DateTime.Now.ToString("dd-MM-yy HH'.'mm'.'ss") + ".txt";
            var log = File.Create(logFile);
            log.Close();

            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Log file created at: " + logFile);
            progress_bar.Value = 10;
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Beginning Test");
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Removing excluded fields listed in config");
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);

            //Loads excluded Fields and removes any test fields that match
            var mergefieldsToTest = ExistenceCheck.RemoveAdditionalSpacingAndUnneccearyFields(mergefieldsInDoc, logFile);
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Field exclusion complete");
            progress_bar.Value = 20;
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);

            //Loads legal Mergefields from config
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Loading legal mergefields from config in to MergeField database");
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);
            var (uniqueMergefields, duplicateMergefields) = ExistenceCheck.DefineLegalMergefields(ConfigurationManager.AppSettings.Get("DataDictionary").Split(',')); //Grabs config arguements for the legal mergefields so I don't have to hardcode
            File.AppendAllText(logFile, "The following fields were loaded as legal fields from the config file, if any fields are missing or incorrect, please add them to the config:" + Environment.NewLine);
            foreach (var field in uniqueMergefields)
            {
                File.AppendAllText(logFile, Environment.NewLine + field);
            }
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Database Loaded");
            progress_bar.Value = 30;
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);

            //Checks if Test Fields exist in the config
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Comparing submitted fields against the MergeField database");
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);
            failedMergefields_doNotExist = ExistenceCheck.CheckExistenceOfMergefield(mergefieldsToTest, uniqueMergefields, logFile);
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Comparison complete, a total of " + failedMergefields_doNotExist.Count + " MergeFields were found to not exist in the database");
            progress_bar.Value = 50;
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);

            //Checks for missing Begin or End groups
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Checking for missing begin and end groups");
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);
            (failedMergefields_MissingBegin, failedMergefields_MissingEnd) = BeginAndEndGroups.CheckAllBeginGroupsHaveEndGroups(mergefieldsToTest);
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Check complete, a total of " + failedMergefields_MissingBegin.Count + " begin groups and " + failedMergefields_MissingEnd.Count + " end groups appear to be missing");
            progress_bar.Value = 75;
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);

            //Checks Hierarchy
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Checking hierarchy against the database");
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);
            failedMergefields_Hierarchy = HierarchyChecker.HierarchyManagement(mergefieldsToTest, logFile);
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Check complete, a total of " + failedMergefields_Hierarchy.Count + " fields did not appear to match the expected Hierarchy");
            progress_bar.Value = 100;
            processingStatus_Text.AppendText(Environment.NewLine + DateTime.Now.ToString("HH'.'mm'.'ss") + ": Test Complete");

            doggo_pic.Opacity = 0.2;
            kitty_pic.Opacity = 0.2;
            hamster_pic.Opacity = 0.2;
            softwareRunning_Label.Content = "Checks Complete! Press continue to view the results";
            viewResults_Button.Visibility = Visibility.Visible;
            Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);
        }

        //Show Results
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            progress_Grid.Visibility = Visibility.Hidden;
            results_Grid.Visibility = Visibility.Visible;

            if (failedMergefields_doNotExist.Count == 0)
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "All Mergefields tested were found in the MergeField database!");
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "All tested fields could be found in the databse!" + Environment.NewLine);
            }
            else
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "The following tested fields could not be found in the MergeField database this could be down to a spelling error, capitalisation issue or spacing:" + Environment.NewLine);
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "The following fields could not be found within the loaded mergefield database, this could be down to a spelling error, capitalisation issue or spacing:" + Environment.NewLine + Environment.NewLine + "***IMPORTANT*** If any begin groups fail here or in the next steps, it will cause major issues later on, fix Begin and End Group issues and then re-run the program!" + Environment.NewLine);
                foreach (var field in failedMergefields_doNotExist)
                {
                    results_text.AppendText(Environment.NewLine + " - " + field);
                    File.AppendAllText(logFile, Environment.NewLine + " - " + field);
                }
            }
            if (failedMergefields_MissingEnd.Count == 0)
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "All tested mergefields had a corresponding end group!");
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "All tested mergefields had a corresponding end group!" + Environment.NewLine);
            }
            else
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "The following Begin Groups are missing their corresponding End Group:" + Environment.NewLine);
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "The Following Fields are missing their corresponding End Group: " + Environment.NewLine);
                foreach (var field in failedMergefields_MissingEnd)
                {
                    results_text.AppendText(Environment.NewLine + " - " + field);
                    File.AppendAllText(logFile, Environment.NewLine + " - " + field);
                }
            }
            if (failedMergefields_MissingBegin.Count == 0)
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "All tested mergefields had a corresponding Begin group!");
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "All tested mergefields had a corresponding Begin group!" + Environment.NewLine);
            }
            else
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "The following End Groups are missing their corresponding Begin Group:" + Environment.NewLine);
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "The Following Fields are missing their corresponding Begin Group: " + Environment.NewLine);
                foreach (var field in failedMergefields_MissingBegin)
                {
                    results_text.AppendText(Environment.NewLine + " - " + field);
                    File.AppendAllText(logFile, Environment.NewLine + " - " + field);
                }
            }
            if (failedMergefields_Hierarchy.Count == 0)
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "All tested mergefields appeared to follow the correct hierarchy!");
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "All tested mergefields appeared to follow the correct hierarchy!" + Environment.NewLine);
            }
            else
            {
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "The following fields did not appear to follow the correct hierarchy, for further details on what groups were open at the time, please check the log file:" + Environment.NewLine);
                File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine + "The following fields did not appear to follow the correct hierarchy, please check the fields here do not also appear in the above sections: " + Environment.NewLine);
                foreach (var field in failedMergefields_Hierarchy)
                {
                    results_text.AppendText(Environment.NewLine + " - " + field);
                    File.AppendAllText(logFile, Environment.NewLine + " - " + field);
                }
                results_text.AppendText(Environment.NewLine + Environment.NewLine + "Thank you for using the Custom Report Debugger, further information pertaining to this report can be found at:"+ Environment.NewLine + logFile);
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            results_Grid.Visibility = Visibility.Hidden;
            selectAndLoad_Grid.Visibility = Visibility.Visible;

            results_text.Document.Blocks.Clear();
            results_text.AppendText("Below are the full results of the various tests, please work section by section when fixing these, and re-run the checks in between. Sometimes errors found higher up in this list, will affect those further down. If you can't work out what the problem is, please speak with other members of the team to try and work it out. Please remember that this debugger is not an auto solver, it covers most issues that arise, however it is not perfect and may well miss errors, for example, it does not check that all groups opened inside tables are closed in the same table. However, if you identify something that it has missed, please let Tom know so that he can try to fix it.");
            processingStatus_Text.Document.Blocks.Clear();

            mergefieldsInDoc.Clear();
            failedMergefields_doNotExist.Clear();
            failedMergefields_Hierarchy.Clear();
            failedMergefields_MissingBegin.Clear();
            failedMergefields_MissingEnd.Clear();
            confirmation_button.Visibility = Visibility.Hidden;
            previewDescription.Visibility = Visibility.Hidden;
            formatting_description.Visibility = Visibility.Hidden;
            textToTest.Document.Blocks.Clear();
            documentSelectionGuidance.Content = "Please select a document in the.docx format to test.";
            SelectFile_Button.Content = "Select File";
            doggo_pic.Opacity = 1;
            kitty_pic.Opacity = 1;
            hamster_pic.Opacity = 1;
            softwareRunning_Label.Content = "Software running, please be patient, this should take 1 - 2 minutes...\nWhilst you wait, please enjoy some pictures of cute pets.";
            viewResults_Button.Visibility = Visibility.Hidden;


        }

        //Load Screen Help
        private void help_button_Click(object sender, RoutedEventArgs e)
        {
            selectAndLoad_Grid.Visibility = Visibility.Hidden;
            helpScreen_Grid.Visibility = Visibility.Visible;
            backToLoad_Button.Visibility = Visibility.Visible;
        }

        private void backToLoad_Button_Click_(object sender, RoutedEventArgs e)
        {
            selectAndLoad_Grid.Visibility = Visibility.Visible;
            helpScreen_Grid.Visibility = Visibility.Hidden;
            backToLoad_Button.Visibility = Visibility.Hidden;
        }

        private void HelpButtonClick_Results(object sender, RoutedEventArgs e)
        {
            results_Grid.Visibility = Visibility.Hidden;
            helpScreen_Grid.Visibility = Visibility.Visible;
            backToResults_Button.Visibility = Visibility.Visible;
        }

        private void backToResults_Button_Click_(object sender, RoutedEventArgs e)
        {
            results_Grid.Visibility = Visibility.Visible;
            helpScreen_Grid.Visibility = Visibility.Hidden;
            backToResults_Button.Visibility = Visibility.Hidden;
        }
    }
}




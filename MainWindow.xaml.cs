//-----------------------------------------------------
// <copyright file="Fields.cs" company="Lateral Thinking Solutions Inc.">
//  Copyright 2016 James N. Romach, All rights Reserved
// </copyright>
// <summary>Object to store field data</summary>
// Information provided by:
//
// http://geekswithblogs.net/TarunArora/archive/2011/08/21/tfs-sdk-work-item-history-visualizer-using-tfs-api.aspx
//
// http://blogs.msdn.com/b/dgorti/archive/2007/09/26/querying-on-workitem-links-through-the-api.aspx
//
// http://social.msdn.microsoft.com/Forums/vstudio/en-US/dae0ce70-e18a-44c9-a788-605909e2e88b/download-video-attached-to-bug-via-tfs-api?forum=vsmantest
//
// http://social.msdn.microsoft.com/Forums/vstudio/en-US/94cfc7ed-37d9-4c52-966b-b42a618cf20a/test-case-result-using-tfs-api?forum=vsmantest
//
// Solutions via http://codeplex.com 
//
//-----------------------------------------------------


using System;
using TfsDataReporting.Models;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.TeamFoundation.Server;
using Excel = Microsoft.Office.Interop.Excel;

namespace TfsDataReporting
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// Initializes a new instance of the MainWindow class
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// TFS project collection variable
        /// </summary>
        private TfsTeamProjectCollection tfs;
        
        /// <summary>
        /// Work item store variable
        /// </summary>
        private WorkItemStore workItemStore;

        /// <summary>
        /// TeamProject interface for test management items
        /// </summary>
        private ITestManagementTeamProject testManagementTeamProject;

        /// <summary>
        /// String to store the team project name
        /// </summary>
        private string teamProject;

        /// <summary>
        /// WorkItemCollection to store all the work items
        /// </summary>
        private WorkItemCollection workItemCollection;

        /// <summary>
        /// WorkItemCollection to store all the work items
        /// </summary>
        private string workItemCurrently = "Bug";  // Default 
        
        /// <summary>
        /// Background worker used for pulling in the TFS data when the team project gets loaded
        /// </summary>
        private BackgroundWorker bw = new BackgroundWorker();

        /// <summary>
        /// Background worker for the export so the main UI doesn't lock
        /// </summary>
        private BackgroundWorker exportBW = new BackgroundWorker();

        /// <summary>
        /// List to store all the fields from TFS
        /// </summary>
        private List<Fields> fieldList = new List<Fields>();

        /// <summary>
        /// List to store the selected fields
        /// </summary>
        private List<string> selectedFields = new List<string>();

        /// <summary>
        /// List to store all the fields from TFS
        /// </summary>
        private List<Fields> WorkItemTypesList = new List<Fields>();

        /// <summary>
        /// String to hold the file location
        /// </summary>
        private string fileLocation;

        /// <summary>
        /// String to hold the file name
        /// </summary>
        private string fileDate;

        /// <summary>
        /// Excel application
        /// </summary>
        private Excel.Application xlApp;

        /// <summary>
        /// Excel workbook
        /// </summary>
        private Excel.Workbook xlWorkBook;

        /// <summary>
        /// Excel worksheet
        /// </summary>
        private Excel.Worksheet xlWorkSheet;

        /// <summary>
        /// Missing value for Excel
        /// </summary>
        private object misValue = System.Reflection.Missing.Value;

        /// <summary>
        /// Chart range for Excel
        /// </summary>
        private Excel.Range chartRange;

        /// <summary>
        /// Row placement holder for when data is entered
        /// </summary>
        private int row = 2;

        /// <summary>
        /// Sheet number to use
        /// </summary>
        private int sheetno = 1;

        /// <summary>
        /// Default sheets
        /// </summary>
        private int defaultSheets;

        /// <summary>
        /// About page view
        /// </summary>
        private AboutPage aboutPage = new AboutPage();

        public MainWindow()
        {
            InitializeComponent();
            bw.DoWork += bw_DoWork;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;
            exportBW.DoWork += exportBW_DoWork;
            exportBW.RunWorkerCompleted += exportBW_RunWorkerCompleted;
            Closing += MainWindow_Closing;
        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;
            workItemCurrently = checkBox.Content.ToString();
        }


        private void InstructionTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }


        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            //MainWindow_Closing(sender, CancelButton);
            if (exportBW.IsBusy)
            {
                var answer = MessageBox.Show("Are you sure you want to exit?", "Notice", MessageBoxButton.YesNo);
                if (answer.Equals(MessageBoxResult.No))
                {
                    MessageBox.Show("Since you canceled, Excel still might be open with the workbook.", "Notice", MessageBoxButton.OK);
                    //e.Cancel = true;
                }
            }
        }

        private void MainWindow_Closing(object sender, Button cancelButton)
        {
            throw new NotImplementedException();
        }



        /// <summary>
        /// Event handler for handling when the main window is closing
        /// </summary>
        /// <param name="sender">Sender Object</param>
        /// <param name="e">Event for cancellation</param>
        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (exportBW.IsBusy)
            {
                var answer = MessageBox.Show("Are you sure you want to exit?", "Notice", MessageBoxButton.YesNo);
                if (answer.Equals(MessageBoxResult.No))
                {
                    MessageBox.Show("Since you canceled, Excel still might be open with the workbook.", "Notice", MessageBoxButton.OK);
                    e.Cancel = true;
                }
            }
        }

        /// <summary>
        /// Event handler when the background worker is completed
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event completed background worker</param>
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            EnableAllControls();

            TfsProjectComboBox.Items.Insert(0, teamProject);
            TfsProjectComboBox.Items.MoveCurrentToFirst();
            TfsProjectComboBox.SelectedIndex = TfsProjectComboBox.Items.CurrentPosition;

        }

        /// <summary>
        /// Background worker for the main page items
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event for the background worker</param>
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            TeamProjectPicker teamProjectPicker = (TeamProjectPicker)e.Argument;
            tfs = teamProjectPicker.SelectedTeamProjectCollection;
            workItemStore = tfs.GetService<WorkItemStore>();
            teamProject = teamProjectPicker.SelectedProjects[0].Name;



            ITestManagementService test_service = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));



            testManagementTeamProject = test_service.GetTeamProject(teamProjectPicker.SelectedProjects[0].Name);



//##################

            var workItemTypes = workItemStore.Projects
                       .Cast<Project>()
                       .SelectMany(p => p.WorkItemTypes.Cast<WorkItemType>())
                       .ToArray();
//##################

            TfsTeamProjectCollection projectCollection = teamProjectPicker.SelectedTeamProjectCollection;
            ProjectInfo info = teamProjectPicker.SelectedProjects[0];

            WorkItemStore store = projectCollection.GetService<WorkItemStore>();
            Project project = store.Projects[info.Name];

            //Category category = project.Categories.FirstOrDefault(c => c.ReferenceName == "Microsoft.RequirementCategory");
            Category category = project.Categories.FirstOrDefault(c => c.ReferenceName == "Microsoft.RequirementCategory");


            //var dddd = project.Categories["category"].WorkItemTypes;
            if (category != null)
            {
                IEnumerable<WorkItemType> types = category.WorkItemTypes;
            }
            //##################




            workItemCollection = workItemStore.Query(" SELECT [System.Id], [System.WorkItemType],[System.State], [System.AssignedTo], [System.Title] FROM WorkItems WHERE [System.TeamProject] = '" + teamProject + "' AND [System.WorkItemType] = '" + workItemCurrently + "' ORDER BY [System.WorkItemType], [System.Id]");
            //workItemCollection = workItemStore.Query(" SELECT [System.Id], [System.WorkItemType],[System.State], [System.AssignedTo], [System.Title] FROM WorkItems WHERE [System.TeamProject] = '" + teamProject + "' AND [System.WorkItemType] = 'Bug' ORDER BY [System.WorkItemType], [System.Id]");
            //workItemCollection = workItemStore.Query(" SELECT [System.Id], [System.WorkItemType],[System.State], [System.AssignedTo], [System.Title] FROM WorkItems WHERE [System.TeamProject] = '" + teamProject + "' AND [System.WorkItemType] = 'Test Case' ORDER BY [System.WorkItemType], [System.Id]");
            //workItemCollection = workItemStore.Query(" SELECT [System.Id], [System.WorkItemType],[System.State], [System.AssignedTo], [System.Title] FROM WorkItems WHERE [System.TeamProject] = '" + teamProject + "' AND [System.WorkItemType] = 'User Story' ORDER BY [System.WorkItemType], [System.Id]");



            if (workItemCollection.Count > 0)
            {
                LoadWorkItemsToListbox();
            }
        }

        /// <summary>
        /// When the tfsProject button is clicked, show the TeamProjectPicker
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event when item is clicked</param>
        private void GetTfsProjectButton_Click(object sender, RoutedEventArgs e)
        {
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();
            if (tpp.SelectedTeamProjectCollection != null)
            {
                DisableAllControls();
                bw.RunWorkerAsync(tpp);
            }
        }

        /// <summary>
        /// Loads the items into the listbox.
        /// Need to use a dispatcher because we are running on a different thread
        /// </summary>
        private void LoadWorkItemsToListbox()
        {
            WorkItem workItem = this.workItemCollection[0];
            fieldList.Clear();
            foreach (Field field in workItem.Fields)
            {
                var newField = new Fields(field.Name);
                fieldList.Add(newField);
            }

            fieldList.Sort((x, y) => x.FieldName.CompareTo(y.FieldName));

            Dispatcher.Invoke((Action)(() =>
            {
                DataItemSelectionListBox.DataContext = fieldList;
            }));
        }

        /// <summary>
        /// Removes the HTML tags from the field values
        /// </summary>
        /// <param name="text">String text from the values of the field</param>
        /// <returns>A string with the HTML tags removed</returns>
        private string removehtmltags(string text)
        {
            text = text.Replace("</P><P>", System.Environment.NewLine);
            text = text.Replace("&nbsp;", " ");
            text = Regex.Replace(text, "<.*?>", "");
            text = text.Replace("&#160;", "");
            return text;
        }

        /// <summary>
        /// Disables all the controls on the main window when the process has started
        /// </summary>
        private void DisableAllControls()
        {
            GetTfsProjectButton.IsEnabled = false;
            GetOutputFileLocationButton.IsEnabled = false;
            ExportDataItemsButton.IsEnabled = false;
            DataItemSelectionListBox.IsEnabled = false;
            ExportAttachmentsCheckBox.IsEnabled = false;
            AboutImage.IsEnabled = false;
        }

        /// <summary>
        /// Enables all the controls on the main window when the process has completed
        /// </summary>
        private void EnableAllControls()
        {
            GetTfsProjectButton.IsEnabled = true;
            GetOutputFileLocationButton.IsEnabled = true;
            ExportDataItemsButton.IsEnabled = true;
            DataItemSelectionListBox.IsEnabled = true;
            ExportAttachmentsCheckBox.IsEnabled = true;
            AboutImage.IsEnabled = true;
        }

        /// <summary>
        /// Method to handle when the file location button is clicked
        /// </summary>
        /// <param name="sender">Object sender</param>
        /// <param name="e">Event for when button is clicked</param>
        private void GetOutputFileLocationButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderDiag = new System.Windows.Forms.FolderBrowserDialog();
            folderDiag.ShowNewFolderButton = true;
            folderDiag.ShowDialog();
            fileLocation = folderDiag.SelectedPath;

            OutputFileLocationComboBox.Items.Insert(0, folderDiag.SelectedPath);
            OutputFileLocationComboBox.Items.MoveCurrentToFirst();
            OutputFileLocationComboBox.SelectedIndex = OutputFileLocationComboBox.Items.CurrentPosition;
        }

        /// <summary>
        /// Method to handle when the file location item selection is changed
        /// </summary>
        /// <param name="sender">Object sender</param>
        /// <param name="e">Event for when button is clicked</param>
        private void OutputFileLocationComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            fileLocation = OutputFileLocationComboBox.SelectedItem.ToString();
        }

        /// <summary>
        /// When the export button is clicked, do validation.
        /// If validation passes, then start the export
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event for when button is clicked</param>
        private void ExportDataItemsButton_Click(object sender, RoutedEventArgs e)
        {
            if (!TfsProjectComboBox.Text.Equals("") && !OutputFileLocationComboBox.Text.Equals("") && DataItemSelectionListBox.SelectedItems.Count > 0)
            {
                if (!isFileOpen())
                {
                    ExportProgressBar.Visibility = Visibility.Visible;
                    WorkItemsLabel.Visibility = Visibility.Visible;
                    ExportProgressBar.Maximum = this.workItemCollection.Count;
                    ExportProgressBar.Value = 0;
                    WorkItemsLabel.Content = "Work Item 0/" + this.workItemCollection.Count;
                    selectedFields.Clear();
                    foreach (Fields item in DataItemSelectionListBox.SelectedItems)
                    {
                        selectedFields.Add(item.FieldName);
                    }

                    DisableAllControls();
                    exportBW.RunWorkerAsync();
                }
                else
                {
                    MessageBox.Show("Please close the Excel document before proceeding.", "Notice", MessageBoxButton.OK);
                }
            }
            else
            {
                MessageBox.Show("Please select all required fields", "Notice", MessageBoxButton.OK);
            }
        }

        /// <summary>
        /// This method is to check whether the file is open or not
        /// </summary>
        /// <returns>Returns a boolean whether the file is open</returns>
        private bool isFileOpen()
        {
            FileStream stream = null;

            fileDate = DateTime.Now.ToString("yyyy.MM.dd.HHmmss");

            FileInfo fileInfo = new FileInfo(this.fileLocation + "\\" + this.teamProject + fileDate + ".xlsx");
            if (fileInfo.Exists)
            {
                try
                {
                    stream = fileInfo.Open(FileMode.Open, FileAccess.Read, FileShare.None);
                }
                catch
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Event handler when the export background worker is complete
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event when the background worker is complete</param>
        private void exportBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("The export is complete.", "Notice", MessageBoxButton.OK);
            EnableAllControls();
            DataItemSelectionListBox.SelectedIndex = -1;
            ExportProgressBar.Visibility = Visibility.Collapsed;
            WorkItemsLabel.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Background worker for the export
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event to start the doWork</param>
        private void exportBW_DoWork(object sender, DoWorkEventArgs e)
        {
            SetupExcelWorkbook();
            foreach (WorkItem workItem in this.workItemCollection)
            {
                bool isChecked = false;
                Dispatcher.Invoke((Action)(() =>
                {
                    WorkItemsLabel.Content = "Work Item " + (ExportProgressBar.Value + 1) + "/" + workItemCollection.Count;
                    ExportProgressBar.Value = ExportProgressBar.Value + 1;
                    isChecked = ExportAttachmentsCheckBox.IsChecked.Value;
                }));
                if (isChecked)
                {
                    ExportBugAttachments(workItem);
                    checkWorkItemForLinksInDescription(workItem);
                    checkForFileAttachmentsViaLinks(workItem);
                }

                EnterDataIntoExcel(workItem);
            }

            FormatExcelSheet();
            SaveExcelWorkbook();
            row = 2;
        }

        /// <summary>
        /// This method will format the Excel sheet for the top row and apply filters
        /// </summary>
        private void FormatExcelSheet()
        {
            chartRange = this.xlWorkSheet.get_Range("A1", ExcelColumnFromNumber(selectedFields.Count + 1) + row);
            chartRange.Columns.AutoFit();

            chartRange = xlWorkSheet.get_Range("A1", ExcelColumnFromNumber(selectedFields.Count + 1) + "1");
            chartRange.AutoFilter(1, misValue, Excel.XlAutoFilterOperator.xlAnd, misValue, true);
        }

        /// <summary>
        /// This method will enter all the data into Excel
        /// </summary>
        /// <param name="workItem">The work item used to enter in</param>
        private void EnterDataIntoExcel(WorkItem workItem)
        {
            int column = 1;
            foreach (string item in selectedFields)
            {
                if (item.Equals("History"))
                {
                    //This is to grab all history items
                    //This is commented out because of the amount of data that gets entered
                    //string historyString = "";

                    //foreach (Revision revision in workItem.Revisions)
                    //{
                    //    historyString = historyString + "\nRevision " + revision.Index + " : \n";
                    //    // Get value of fields in the work item revision
                    //    foreach (Field field in workItem.Fields)
                    //    {
                    //        if (field.Value != null)
                    //        {
                    //            if (!field.Value.Equals(""))
                    //            {

                    //                    historyString = historyString + field.Name + " : " + revision.Fields[field.Name].Value + "\n";

                    //            }
                    //        }
                    //        //Console.WriteLine(revision.Fields[field.Name].Value);
                    //    }
                    //}
                    //This is only for the comments
                    RevisionCollection revisionCollection = workItem.Revisions;
                    string historyString = "";
                    foreach (Revision rev in revisionCollection)
                    {
                        if (rev.Fields["History"].Value != null)
                        {
                            if (!rev.Fields["History"].Value.Equals(""))
                            {
                                historyString = historyString + rev.Fields["History"].Value + "\n";
                            }
                        }
                    }

                    historyString = removehtmltags(historyString);
                    xlWorkSheet.Cells[row, column] = historyString;
                }
                else
                {
                    xlWorkSheet.Cells[row, column] = removehtmltags(CheckForNull(workItem.Fields[item]));
                }

                column++;
            }

            var list = GetLinkedWorkItems(workItem);
            string combinedList = "";
            foreach (string listString in list)
            {
                combinedList = combinedList + listString + "\n";
            }

            xlWorkSheet.Cells[row, column] = combinedList;
            row++;
        }

        /// <summary>
        /// This method is to get all the linked work items
        /// </summary>
        /// <param name="workItem">Work item being used</param>
        /// <returns>A list of work item names and their ID</returns>
        private List<string> GetLinkedWorkItems(WorkItem workItem)
        {
            WorkItemLinkCollection workItemLinkCollection = workItem.WorkItemLinks;
            List<string> workItemList = new List<string>();
            foreach (WorkItemLink link in workItemLinkCollection)
            {
                var newWorkItem = workItemStore.GetWorkItem(link.TargetId);
                workItemList.Add(newWorkItem.Type.Name + " : " + newWorkItem.Id);
            }

            return workItemList;
        }

        /// <summary>
        /// This method will check for null values
        /// </summary>
        /// <param name="field">Field being sent to method</param>
        /// <returns>A blank string</returns>
        private string CheckForNull(Field field)
        {
            if (field.Value != null)
            {
                return field.Value.ToString();
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// This method will setup the Excel workbook
        /// </summary>
        private void SetupExcelWorkbook()
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            sheetno = 1;
            defaultSheets = xlWorkBook.Sheets.Count;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetno);
            xlWorkSheet.Name = workItemCurrently;

            int column = 1;
            foreach (string item in selectedFields)
            {
                xlWorkSheet.Cells[1, column] = item;
                column++;
            }

            xlWorkSheet.Cells[1, column] = "Linked Work Items";
            chartRange = xlWorkSheet.get_Range("A1", ExcelColumnFromNumber(selectedFields.Count + 1) + "1");
            chartRange.Font.Bold = true;
        }

        /// <summary>
        /// Changes the column form a number to a letter for Excel
        /// </summary>
        /// <param name="column">The integer for the column </param>
        /// <returns>A string for the letter</returns>
        private string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            return columnString;
        }

        /// <summary>
        /// This method will save the excel workbook
        /// </summary>
        private void SaveExcelWorkbook()
        {
            try
            {
                xlWorkBook.SaveAs(fileLocation + "\\" + teamProject + fileDate + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlApp);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);

                //MessageBox.Show("Test Cases exported successfully to specified file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
                if (ex.Message == "Cannot access '" + teamProject + fileDate + ".xlsx'.")
                {
                    //MessageBox.Show("File with same name exists in specified location", "File Exists", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //txtFileName.Text = "";
                }
                else
                {
                    //MessageBox.Show("Application has encountered Fatal Errro. \nPlease contact your System Administrator.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show("Application has encountered Fatal Error.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Handle objects to try and release them
        /// </summary>
        /// <param name="obj">Object to be sent over</param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// This method will take the work item to check to see if there are any file attachments in the links
        /// </summary>
        /// <param name="workItem">Work item being brought in</param>
        private void checkForFileAttachmentsViaLinks(WorkItem workItem)
        {
            LinkCollection links = workItem.Links;
            System.Net.WebClient request = new System.Net.WebClient();

            // NOTE: If you use custom credentials to authenticate with TFS then you would most likely
            //       want to use those same credentials here
            request.Credentials = System.Net.CredentialCache.DefaultCredentials;
            foreach (Link link in links)
            {
                if (!(link is ExternalLink))
                {
                    continue;
                }

                ExternalLink externalLink = (ExternalLink)link;
                var artifactUri = new Uri(externalLink.LinkedArtifactUri);
                ITestAttachment attachment;
                testManagementTeamProject.TestResults.FindByLink(artifactUri, out attachment);
                if (attachment != null)
                {
                    string fileLocationForDownload = fileLocation + "\\" + workItem.Id + " - " + attachment.Name;
                    request.DownloadFile(attachment.Uri.OriginalString, fileLocationForDownload);
                }
            }
        }

        /// <summary>
        /// This method will check for links to grab
        /// </summary>
        /// <param name="workItem">Work item being passed in</param>
        private void checkWorkItemForLinksInDescription(WorkItem workItem)
        {
            string reproSteps = workItem.Fields[13].Value.ToString();

            System.Net.WebClient request = new System.Net.WebClient();

            // NOTE: If you use custom credentials to authenticate with TFS then you would most likely
            //       want to use those same credentials here
            request.Credentials = System.Net.CredentialCache.DefaultCredentials;
            List<string> links = new List<string>();
            string regexImgSrc = @"<img[^>]*?src\s*=\s*[""']?([^'"" >]+?)[ '""][^>]*?>";
            MatchCollection matchesImgSrc = Regex.Matches(reproSteps, regexImgSrc, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (Match m in matchesImgSrc)
            {
                string href = m.Groups[1].Value;
                links.Add(href);
            }

            foreach (string link in links)
            {
                if (link.Contains("fileName"))
                {
                    string fileName = this.After(link, "fileName=");

                    Console.WriteLine("Attachment: '" + fileName);

                    string fileLocationForDownload = fileLocation + "\\" + workItem.Id + " - " + fileName;

                    // Save the attachment to a local file
                    request.DownloadFile(new Uri(link), fileLocationForDownload);
                }
            }
        }

        /// <summary>
        /// Gets the name of the file name
        /// <c>http://www.dotnetperls.com/between-before-after</c>
        /// </summary>
        /// <param name="value">String value being passed in</param>
        /// <param name="a">String to get file name after</param>
        /// <returns>The string after the value</returns>
        private string After(string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }

            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }

            return value.Substring(adjustedPosA);
        }

        /// <summary>
        /// Takes the work item being passed in and downloads the attachment to a specified file
        /// </summary>
        /// <param name="workItem">Brings in the workItem to be used</param>
        private void ExportBugAttachments(WorkItem workItem)
        {
            if (workItem.Attachments.Count > 0)
            {
                if (workItem != null)
                {
                    System.Net.WebClient request = new System.Net.WebClient();
                    request.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    foreach (Attachment attachment in workItem.Attachments)
                    {
                        // Display the name & size of the attachment
                        Console.WriteLine("Attachment: '" + attachment.Name + "' (" + attachment.Length.ToString() + " bytes)");

                        string fileLocationForDownload = fileLocation + "\\" + workItem.Id + " - " + attachment.Name;

                        // Save the attachment to a local file
                        request.DownloadFile(attachment.Uri, fileLocationForDownload);
                    }
                }
            }
        }

        /// <summary>
        /// Event handler when the about icon is tapped
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event for Mouse button up</param>
        private void AboutImage_MouseUp(object sender, MouseButtonEventArgs e)
        {
            this.aboutPage.ShowDialog();
        }

        private void DataItemSelectionListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

    }
}

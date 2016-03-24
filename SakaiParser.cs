///////////////////////////////////////////////////////////////////////////////
//
// AErobotSakai
// author: Alexander Yasko
//
// version 1.01 (20150110)
// version 1.02 (20150110)
// version 1.03 (20150111)
// version 1.04 (20150112)
// version 1.05 (20150228)
// version 1.06 (20150301)
// version 1.07 (20150306)
// version 1.08 (20150308)
// version 1.09 (20150309) Fixed issue of authorization using partial URL (aeperepelitsyn)
// version 1.10 (20150320) Added asynchronous group deleting
// version 1.11 (20150517) Added comments for methods. Fixied issue with Initialize(), GetAssignmentItems(), etc.
// version 1.12 (20150526) Added ability to grade an assignments (Alexander Yasko) and IsIdle() method (aeperepelitsyn)
// version 1.13 (20150529) Added possibility to set first name and last name for specified user by its id
// version 1.14 (20150601) Fixed issue with User tab for renaming users (Alexander Yasko) and with reloading of student information (aeperepelitsyn)
// version 1.15 (20160123) Fixed issue with Draft Assignments, Added method ReadWorksitesAsync (aeperepelitsyn)
// version 1.16 (20160125) Added ability to read all worksites, fixed issue with few worksites with the same name (moskalenkoBV)
// version 1.17 (20160322) Added event WorksitesReady for providing of ability to handle end of worksites parsing (moskalenkoBV)
// version 1.18 (20160324) Added events WorksiteSelected, AssignmentItemsReady, StudentsInformationReady (moskalenkoBV)
//
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SakaiParser
{
    public class Assignment
    {
        public Assignment(string title, string link, string status, string open, string due, string innew, string scale)
        {
            StudentInfosDictionary = new Dictionary<string, StudentInfo>();
            Link = link;
            Title = title;
            Status = status;
            Open = open;
            Due = due;
            InNew = innew;
            Scale = scale;
        }

        public Dictionary<string, StudentInfo> StudentInfosDictionary { get; set; }

        public string Link { get; set; }
        public string Title { get; set; }
        public string Status { get; set; }
        public string Open { get; set; }
        public string Due { get; set; }
        public string InNew { get; set; }
        public string Scale { get; set; }

        public static implicit operator String(Assignment a)
        {
            return a.Title;
        }

        public override string ToString()
        {
            return Title;
        }
    }

    public struct SubmittedFile
    {
        public string Name;
        public string Link;
    }

    public class StudentInfo
    {
        public StudentInfo(string name, string id, string submitted, string status, string grade, bool released, string gradeLink, bool attached)
        {
            Name = name;
            ID = id;
            FilesAttached = attached;
            Submitted = submitted;
            Status = status;
            Grade = grade;
            Released = released;
            GradeLink = gradeLink;
        }

        public List<SubmittedFile> SubmittedFiles { get; set; }
        public bool FilesAttached { get; set; }
        public string Submitted { get; set; }
        public string Status { get; set; }
        public string Grade { get; set; }
        public bool Released { get; set; }
        public string GradeLink { get; set; }
        public string Name { get; set; }
        public string ID { get; set; }

        public override string ToString()
        {
            return string.Format("{0} ({1})", Name, ID);
        }
    }

    public enum GroupEditorSectionTask
    {
        CreateNewGroup,
        DeleteExistingGroup,
        GetGroupList
    }

    public enum WebBrowserTask
    {
        Idle,
        Busy,
        LogIn,
        GetMembershipLink,
        ParseWorksites,
        ParseSelectedWorksite,
        GoToAssignments,
        GoToAdministrationWorkspaceUsersTab,
        ParseAssignments,
        LoadStudents,
        ReloadStudents,
        LoadStudentAttachments,
        OpenManageGroupsSection,
        LoadGroupsEditor,
        RenameStudent,
        ContinueStudentRenaming,
        SelectStudentToGrade,
        GradeStudent,
        AddNewGroup,
        SetNewFirstNameAndLastName
    }

    public enum SPExceptions
    {
        UnableToFindMembership,
        UnableToFindLinkToUsersTab,
        WorksiteNameAlreadyExist
    }                   
    public delegate void DelegateException(SPExceptions exception, String message);
    public delegate void DelegateWorksitesReady(String[] worksites);
    public delegate void DelegateWorksiteSelected(String worksiteName);
    public delegate void DelegateAssignmentItemsReady(String[] assignmentTitles);
    public delegate void DelegateStudentsInformationReady(String[] studentIDs); 

    class SakaiParser285
    {
        delegate void ConfirmDeleting();

        ConfirmDeleting confirmDeletingVoid;

        WebBrowser webBrowser;

        WebBrowserTask webBrowserTask;
        GroupEditorSectionTask groupSectionTask;

        Dictionary<string, string> dctWorksites;
        Dictionary<string, StudentInfo> dctStudentInfos;
        Dictionary<string, Assignment> dctAssignmentItems;

        string worksiteName;
        string linkToMembership;
        string linkToSiteEditor;
        string linkToManageGroupsSection;
        string linkToCreateNewGroupSection;

        private string linkToAssignments; // Link to Assignments page

        private string gradeStudentID;
        private String studentmark;

        int addingStudentsCount;

        string addingGroupName;
        string addingGroupDescription;

        string renamingStudentName;
        string renamingStudentID;
        string renamingStudentLastname;

        string[] addingStudentIDs;

        string deletingGroupName;

        int indexOfProcessingAssignment;
        int indexOfProcessingStudent;

        bool confidentLoad;
        bool attachmentPresent;

        bool assignmentsParsed;

        public event DelegateException SPException;
        private void SPExceptionProvider(SPExceptions exception)
        {
            String message;
            switch(exception)
            {
                case SPExceptions.UnableToFindMembership:
                    message = "Unable to find Membership link"; break;
                case SPExceptions.UnableToFindLinkToUsersTab:
                    message = "Unable to find link to Users Tab"; break;
                case SPExceptions.WorksiteNameAlreadyExist:
                    message = "Worksite name already exist"; break;
                default:
                    message = "Unknown exception"; break;
            }
            if (SPException != null)
            {
                SPException(exception, message);
            }
            else
            {
                throw new Exception(message);
            }
        }

        public event DelegateWorksitesReady WorksitesReady;
        private void WorksitesReadyProvider()
        {
            if (WorksitesReady != null)
            {
                WorksitesReady(GetWorksites());
            }
        }

        public event DelegateWorksiteSelected WorksiteSelected;
        private void WorksiteSelectedProvider()
        {
            if (WorksiteSelected != null)
            {
                WorksiteSelected(worksiteName);
            }
        }

        public event DelegateAssignmentItemsReady AssignmentItemsReady;
        private void AssignmentItemsReadyProvider()
        {
            if (AssignmentItemsReady != null)
            {
                AssignmentItemsReady(GetAssignmentItemNames());
            }
        }

        public event DelegateStudentsInformationReady StudentsInformationReady;
        private void StudentsInformationReadyProvider()
        {
            if (StudentsInformationReady != null)
            {
                StudentsInformationReady(GetStudentIDs());
            }
        }
        
        void ResetFields()
        {
            dctStudentInfos.Clear();
            dctAssignmentItems.Clear();
        }

        public SakaiParser285(WebBrowser webBrowser, string initialUrl, string userName, string password)
        {
            InitialUrl = initialUrl;
            UserName = userName;
            Password = password;
            confidentLoad = false;
            assignmentsParsed = false;
            dctWorksites = new Dictionary<string, string>();
            dctAssignmentItems = new Dictionary<string, Assignment>();
            dctStudentInfos = new Dictionary<string, StudentInfo>();
            worksiteName = "";
            linkToMembership = "";
            SPException = null;
            this.webBrowser = webBrowser;
            webBrowser.DocumentCompleted += webBrowser_DocumentCompleted;
            

            confirmDeletingVoid = new ConfirmDeleting(InvokeGroupDeleting);
        }

        public string InitialUrl { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        public bool IsIdle()
        {
            return webBrowserTask == WebBrowserTask.Idle;
        }

        /// <summary>
        /// Initializes program passing Login in navigating to specified worksite.
        /// </summary>
        /// <param name="worksiteName">The worksite</param>
        public void Initialize(string worksiteName)
        {
            this.worksiteName = worksiteName;
            Initialize();
            // Than login when document is completed
        }

        /// <summary>
        /// Initializes program passing Login in.
        /// </summary>
        public void Initialize()
        {
            webBrowserTask = WebBrowserTask.LogIn;
            webBrowser.Navigate(InitialUrl);
        }

        /// <summary>
        /// Function read worksite names for current authorized user.
        /// </summary>
        public void ReadWorksitesAsync()
        {
            bool renavigate = (webBrowserTask == WebBrowserTask.Idle);
            webBrowserTask = WebBrowserTask.GetMembershipLink;
            if (renavigate)
            {
                webBrowser.Navigate(webBrowser.Url);
            }
        }

        /// <summary>
        /// Selects the worksite name.
        /// The name should be existing.
        /// </summary>
        /// <param name="worksiteName">The name of the worksite</param>
        public void SelectWorksite(string worksiteName)
        {
            this.worksiteName = worksiteName;
            ResetFields();
            webBrowserTask = WebBrowserTask.ParseSelectedWorksite;
            webBrowser.Navigate(dctWorksites[worksiteName]);
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (!confidentLoad)
            {
                if (e == null) return;
                if (e.Url != webBrowser.Url) return; // If we have not reached destination URL
                if (String.IsNullOrWhiteSpace(webBrowser.DocumentTitle)) return; // Waiting for page
            }
            else // Web form has been reloaded, we know confidently
            {
                confidentLoad = false; // We don't want to be confident every time =)
            }

            switch (webBrowserTask)
            {
                case WebBrowserTask.Busy:
                    {
                        webBrowserTask = WebBrowserTask.Idle;
                        break;
                    }
                case WebBrowserTask.LogIn:
                    LogIn();
                    break;
                case WebBrowserTask.ParseWorksites:
                    HtmlElement worksiteTable = webBrowser.Document.Window.Frames[0].Document.GetElementById("currentSites");
                    HtmlElementCollection tableTDs = worksiteTable.GetElementsByTagName("td");
                    foreach (HtmlElement worksiteTD in tableTDs)
                    {
                        if (worksiteTD.GetAttribute("headers") == "worksite")
                        {
                            HtmlElement linkToWorksite = worksiteTD.GetElementsByTagName("a")[0];
                            string sLink = linkToWorksite.GetAttribute("href");
                            if (dctWorksites.ContainsKey(worksiteTD.InnerText))
                            {
                               //SPExceptionProvider(SPExceptions.WorksiteNameAlreadyExist);
                            }
                            else
                            {
                                dctWorksites.Add(worksiteTD.InnerText, sLink);
                            }
                        }
                    }
                    //Alternative implementation:
                    //HtmlElement worksiteTables = webBrowser.Document.Window.Frames[0].Document.GetElementById("currentSites");
                    //HtmlElementCollection curSites = worksiteTables.GetElementsByTagName("a");
                    //string sLink = "";
                    //foreach (HtmlElement workSiteName in curSites)
                    //{
                    //    if (workSiteName.GetAttribute("target") == "_top")
                    //    {
                    //        sLink = workSiteName.GetAttribute("href");
                    //        dctWorksites.Add(workSiteName.InnerText,sLink);
                    //    }
                    //}
                    // Worksites in dctWorksites
                    bool nextPageAvailable = false;
                    HtmlElementCollection buttonForms = webBrowser.Document.Window.Frames[0].Document.GetElementsByTagName("input");
                    foreach (HtmlElement worksiteTD in buttonForms)
                    {
                        if (worksiteTD.GetAttribute("name") == "eventSubmit_doList_next")
                        {
                            if (worksiteTD.GetAttribute("disabled") == "False")
                            {
                                nextPageAvailable = true;
                                break;
                            }
                        }
                    }
                    if (nextPageAvailable)
                    {
                        webBrowserTask = WebBrowserTask.ParseWorksites;
                        webBrowser.Navigate("javascript:window.frames[0].document.forms[4].elements[0].click()");
                        Task asyncTask = new Task(() =>
                        {
                            Thread.Sleep(500);
                        });

                        asyncTask.ContinueWith((a) =>
                        {
                            webBrowser.Navigate(webBrowser.Url);
                        }, TaskScheduler.FromCurrentSynchronizationContext());

                        asyncTask.Start();
                    }
                    else
                    {
                        webBrowserTask = WebBrowserTask.Idle;
                        WorksitesReadyProvider();
                    }
                    break;
                case WebBrowserTask.GetMembershipLink:

                    HtmlElementCollection linksCollection = webBrowser.Document.GetElementsByTagName("a");
                    foreach (HtmlElement link in linksCollection)
                        if (link.GetAttribute("className") == "icon-sakai-membership")
                            linkToMembership = link.GetAttribute("href");
                    if (linkToMembership == "")
                    {
                        SPExceptionProvider(SPExceptions.UnableToFindMembership);
                    }
                    webBrowserTask = WebBrowserTask.ParseWorksites;
                    dctWorksites.Clear();
                    webBrowser.Navigate(linkToMembership);
                    break;
                case WebBrowserTask.ParseSelectedWorksite:
                    // Now we are at HOME link, maybe

                    HtmlElementCollection links = webBrowser.Document.GetElementsByTagName("a");
                    foreach (HtmlElement link in links)
                        if (link.GetAttribute("className") == "icon-sakai-assignment-grades")
                            linkToAssignments = link.GetAttribute("href");

                    if (linkToAssignments == "") throw new Exception("Unable to find Assignments link");

                    // Now we are at Home of the WorkSite
                    // Parsing site editor link.

                    linkToSiteEditor = "";
                    links = webBrowser.Document.GetElementsByTagName("a");
                    foreach (HtmlElement link in links)
                        if (link.GetAttribute("className") == "icon-sakai-siteinfo")
                            linkToSiteEditor = link.GetAttribute("href");

                    //                    if (linkToSiteEditor == "") throw new Exception("Unable to find SiteEditor link");

                    // If it is okay, we have linkToSiteEditor
                    webBrowserTask = WebBrowserTask.Idle;
                    WorksiteSelectedProvider();
                    break;
                case WebBrowserTask.GoToAssignments:
                    //////////////////////////////////////////////////////////////////////////////////
                    /*webBrowserTask = WebBrowserTask.ParseAssignments;
                    webBrowser.Navigate(linkToAssignments);*/
                    //////////////////////////////////////////////////////////////////////////////////

                    break;
                case WebBrowserTask.ParseAssignments:
                    // To fix
                    HtmlElementCollection assignmentsTable = webBrowser.Document.Window.Frames[1].Document.Forms["listAssignmentsForm"].Document.GetElementsByTagName("table");
                    foreach (HtmlElement table in assignmentsTable)
                    {
                        if (table.GetAttribute("className") == "listHier lines nolines")
                        {
                            HtmlElementCollection trs = table.GetElementsByTagName("tr");
                            foreach (HtmlElement tr in trs)
                            {
                                HtmlElementCollection tds = tr.GetElementsByTagName("td");

                                string title = "";
                                string status = "";
                                string open = "";
                                string url = "";
                                string due = "";
                                string innew = "";
                                string scale = "";

                                foreach (HtmlElement td in tds)
                                {
                                    if (td.GetAttribute("headers") == "title")
                                    {
                                        HtmlElement link = td.GetElementsByTagName("a")[0];
                                        title = link.InnerText.Trim();
                                    }
                                    if (td.GetAttribute("headers") == "status")
                                    {
                                        status = td.InnerText;
                                    }
                                    if (td.GetAttribute("headers") == "openDate")
                                    {
                                        open = td.InnerText;
                                    }
                                    if (td.GetAttribute("headers") == "dueDate")
                                    {
                                        due = td.InnerText;
                                    }
                                    if (td.GetAttribute("headers") == "num_submissions")
                                    {
                                        HtmlElementCollection linksCount = td.GetElementsByTagName("a");
                                        if (linksCount.Count > 0)
                                        {
                                            HtmlElement link = linksCount[0];
                                            string outerHtml = link.OuterHtml;
                                            MatchCollection mathes = Regex.Matches(outerHtml,
                                                "window.location\\s*=\\s*('|\")(.*?)('|\")", RegexOptions.IgnoreCase);
                                            url = mathes[0].Groups[2].Value;

                                            innew = td.InnerText;
                                        }
                                    }
                                    if (td.GetAttribute("headers") == "maxgrade")
                                    {
                                        scale = td.InnerText;
                                    }
                                }
                                if (title != "" && status != "" && open != "" && url != "" && due != "" && innew != "" && scale != "")
                                    dctAssignmentItems.Add(title, new Assignment(title, url, status, open, due, innew, scale));
                            }
                        }
                    }
                    // Assignments are supposed to be parsed
                    assignmentsParsed = true;
                    webBrowserTask = WebBrowserTask.Idle;
                    AssignmentItemsReadyProvider();
                    break;
                // We have all assignments. Go to LoadStudents

                case WebBrowserTask.ReloadStudents:

                    ParseLoadingStudents(indexOfProcessingAssignment, false);
                    webBrowserTask = WebBrowserTask.Idle;
                    StudentsInformationReadyProvider();
                    break;
                case WebBrowserTask.LoadStudents:

                    // Finding students
                    ParseLoadingStudents(indexOfProcessingAssignment, false);

                    confidentLoad = true; // We are not women :)
                    indexOfProcessingAssignment++;
                    if (indexOfProcessingAssignment >= dctAssignmentItems.Count)
                    {
                        // We have all in data base
                        // We can start to download students attachments
                        // We will process each assignment and each submission of processing assignment
                        webBrowserTask = WebBrowserTask.LoadStudentAttachments;
                        indexOfProcessingAssignment = 0; // Reset indexers
                        indexOfProcessingStudent = 0; // Reset indexers
                        // Only student with attachments will ne processed
                        confidentLoad = true; // Here we also men

                        attachmentPresent = dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.FilesAttached;
                        if (attachmentPresent)
                            webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.GradeLink); // If you see it you are hero
                        else webBrowser_DocumentCompleted(webBrowser, null);
                        // Than the programe will be in LoadStudentAttachments switch case
                    }
                    else webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.Link);

                    break;
                case WebBrowserTask.LoadStudentAttachments:
                    // Indexers should be initialized before this case reached

                    if (attachmentPresent)
                    {
                        // We are here if student has attachment. The page with it is loaded
                        HtmlElementCollection uls = webBrowser.Document.Window.Frames[1].Document.Forms["gradeForm"].Document.GetElementsByTagName("ul");
                        foreach (HtmlElement ul in uls)
                        {
                            if (ul.GetAttribute("className") == "attachList indnt1")
                            {
                                HtmlElementCollection alinks = ul.GetElementsByTagName("a");
                                foreach (HtmlElement alink in alinks)
                                {
                                    string slink = alink.GetAttribute("href");
                                    string fileName = alink.InnerText;

                                    if (dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.SubmittedFiles == null)
                                    {
                                        dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.SubmittedFiles = new List<SubmittedFile>();
                                    }

                                    SubmittedFile file = new SubmittedFile();
                                    file.Link = slink;
                                    file.Name = fileName;
                                    dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.SubmittedFiles.Add(file);
                                }
                            }
                        }
                    }

                    indexOfProcessingStudent++;
                    if (indexOfProcessingStudent >= dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.Count)
                    {
                        indexOfProcessingAssignment++;
                        indexOfProcessingStudent = 0;
                        if (indexOfProcessingAssignment >= dctAssignmentItems.Count)
                        {
                            webBrowserTask = WebBrowserTask.Idle; // We are done with attachments loading. Lets have a rest
                            break;
                        }
                    }

                    attachmentPresent = dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.FilesAttached;
                    if (attachmentPresent == true)
                    {
                        // If the student attached files
                        confidentLoad = true; // I'm a man! too dooo do do
                        webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.GradeLink);
                    }
                    else
                        if (attachmentPresent == false)
                        {
                            confidentLoad = true; // Bypass call. Don't know whether it is necessary, but let it be
                            webBrowser_DocumentCompleted(webBrowser, e);
                        }

                    break;
                case WebBrowserTask.OpenManageGroupsSection:
                    // Todo: find link to Manage Groups section.

                    linkToManageGroupsSection = "";
                    HtmlElementCollection lisElementCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("li");
                    foreach (HtmlElement element in lisElementCollection)
                    {
                        if (element.GetAttribute("role") == "menuitem" && element.InnerHtml.Contains("doManageGroupHelper"))
                        {
                            MatchCollection matchCollection = Regex.Matches(element.InnerHtml,
                                @"(onclick)\s*=\s*""\s*(location)\s*=\s*'(?<link>.*)'\s*;");
                            linkToManageGroupsSection = matchCollection[0].Groups["link"].Value;
                        }
                    }

                    // linkToManageGroupsSection contains link to Manage Groups section. 
                    // We can navigate it right now.

                    webBrowser.Document.Window.Frames[1].Navigate(linkToManageGroupsSection);


                    //webBrowser.Navigate(linkToManageGroupsSection);
                    // LoadGroupsEditor is a page where we set name of group and add participants.
                    webBrowserTask = WebBrowserTask.LoadGroupsEditor;

                    confidentLoad = true; // The Frame URL of doesn't match Browser URL, beacuse Browser URL has nothing to do with frame.
                    //webBrowser_DocumentCompleted(webBrowser, e);

                    break;
                case WebBrowserTask.LoadGroupsEditor:
                    // Here we find the link to Create New Group section

                    switch (groupSectionTask)
                    {
                        case GroupEditorSectionTask.CreateNewGroup:
                            {
                                linkToCreateNewGroupSection = "";

                                HtmlElementCollection liElementCollection =
                                    webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("li");
                                foreach (
                                    HtmlElement htmlElement in
                                        liElementCollection.Cast<HtmlElement>()
                                            .Where(
                                                htmlElement => htmlElement.GetAttribute("className") == "firstToolBarItem"))
                                {
                                    if (htmlElement.InnerHtml.Contains("GroupEdit") &&
                                        htmlElement.InnerHtml.Contains("item_control"))
                                    {
                                        MatchCollection matchCollection = Regex.Matches(htmlElement.InnerHtml,
                                            "<[a|A].*href\\s*=\\s*[\"|'](?<link>.*)[\"|']\\s*>");
                                        linkToCreateNewGroupSection = matchCollection[0].Groups["link"].Value;
                                    }
                                }

                                if (linkToCreateNewGroupSection == "")
                                    throw new Exception("Link to Create New Group is not found.");

                                confidentLoad = true;
                                // The Frame URL of doesn't match Browser URL, beacuse Browser URL has nothing to do with frame.
                                webBrowser.Document.Window.Frames[1].Navigate(linkToCreateNewGroupSection);

                                // Browser is navigated to linkToCreateNewGroupSection
                                // Now lets add new group. 

                                webBrowserTask = WebBrowserTask.AddNewGroup;

                                break;
                            }
                        case GroupEditorSectionTask.DeleteExistingGroup:
                            {

                                webBrowserTask = WebBrowserTask.Idle;

                                HtmlElementCollection tablesCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("table");
                                foreach (HtmlElement element in tablesCollection)
                                {
                                    if (element.GetAttribute("className").Contains("listHier") ||
                                        element.GetAttribute("className").Contains("lines") ||
                                        element.GetAttribute("className").Contains("nolines") ||
                                        element.GetAttribute("className").Contains("centerLines"))
                                    {
                                        HtmlElementCollection trsCollection = element.GetElementsByTagName("tr");
                                        foreach (HtmlElement trElement in trsCollection)
                                        {
                                            if (trElement.InnerText.Contains(deletingGroupName))
                                            {
                                                if (trElement.GetElementsByTagName("span")[0].InnerText == deletingGroupName)
                                                {
                                                    HtmlElement inputToCheck = trElement.GetElementsByTagName("input")[0];
                                                    inputToCheck.InvokeMember("CLICK");
                                                }
                                            }
                                        }
                                    }
                                }

                                HtmlElement deleteButtonInputElement = webBrowser.Document.Window.Frames[1].Document.GetElementById("delete-groups");
                                deleteButtonInputElement.InvokeMember("CLICK");

                                new Thread(() =>
                                {
                                    Thread.Sleep(1000);
                                    Application.OpenForms["Form1"].Invoke(confirmDeletingVoid);
                                }).Start();

                                break;
                            }
                        case GroupEditorSectionTask.GetGroupList:
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }


                    break;
                case WebBrowserTask.AddNewGroup:

                    // Description adding is not implemented yet.

                    HtmlWindow addGroupFrame = webBrowser.Document.Window.Frames[1];

                    addGroupFrame.Document.GetElementById("group_title").SetAttribute("value", addingGroupName);

                    HtmlElement selectOprionsElement = addGroupFrame.Document.GetElementById("siteMembers-selection");
                    HtmlElementCollection options = selectOprionsElement.Document.GetElementsByTagName("option");

                    Dictionary<string, int> dctSelectionOptions = new Dictionary<string, int>();

                    for (int i = 0; i < options.Count; i++)
                    {
                        MatchCollection matches = Regex.Matches(options[i].OuterHtml, ">.*\\((?<id>.*)\\)\\s*<");
                        dctSelectionOptions.Add(
                            matches.Count != 0 ? matches[0].Groups["id"].Value : options[i].InnerText, i);
                    }

                    foreach (string id in addingStudentIDs)
                    {
                        addGroupFrame.Document.GetElementsByTagName("option")[dctSelectionOptions[id]].SetAttribute("selected", "selected");
                        addingStudentsCount++;
                    }

                    if (addingStudentIDs.Length != 0)
                        addGroupFrame.Document.GetElementById("siteMembers-selection").InvokeMember("ondblclick");

                    addGroupFrame.Document.GetElementById("save").InvokeMember("click");
                    //addGroupFrame.Document.GetElementById("pw").SetAttribute("value", Password);
                    //addGroupFrame.Navigate("javascript:document.forms[0].submit()");
                    //webBrowserTask = WebBrowserTask.GetMembershipLink;

                    webBrowserTask = WebBrowserTask.Idle;

                    break;
                case WebBrowserTask.SelectStudentToGrade:
                    // Current page contains the list of students.
                    // Let's find the link width "Grade" title

                    ParseLoadingStudents(0, true);

                    break;

                case WebBrowserTask.GradeStudent:

                    webBrowser.Document.Window.Frames[1].Document.GetElementById("grade").SetAttribute("value", studentmark);

                    HtmlElementCollection processBtn = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
                    foreach (HtmlElement element in processBtn)
                    {
                        if (element.GetAttribute("name") == "return")
                        {
                            element.InvokeMember("CLICK");
                            webBrowserTask = WebBrowserTask.Idle;
                            break;
                        }
                    }
                    webBrowserTask = WebBrowserTask.Idle;
                    break;
                case WebBrowserTask.GoToAdministrationWorkspaceUsersTab:
                    {
                        HtmlElementCollection hec = webBrowser.Document.GetElementsByTagName("a");
                        //foreach(String str in from HtmlElement he in hec where he.InnerText.Contains("Users") select he.GetAttribute("href"))
                        String str = "";
                        foreach (HtmlElement elem in hec)
                        {
                            if (elem.InnerText != null)
                            {
                                if (elem.InnerText.Contains("Users"))
                                {
                                    str = elem.GetAttribute("href");
                                    break;
                                }
                            }
                        }
                        if (!String.IsNullOrEmpty(str))
                        {
                            webBrowserTask = WebBrowserTask.Busy;
                            webBrowser.Navigate(str);
                        }
                        else
                        {
                            SPExceptionProvider(SPExceptions.UnableToFindLinkToUsersTab);
                            return;
                        }
                        break;
                    }
                case WebBrowserTask.RenameStudent:
                    {
                        HtmlElement inputSearch = webBrowser.Document.Window.Frames[1].Document.GetElementById("search");

                        if (inputSearch == null) throw new NullReferenceException("Input field is NULL.");

                        inputSearch.Document.GetElementById("search").SetAttribute("value", renamingStudentID);

                        HtmlElementCollection framesAs = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("a");

                        foreach (HtmlElement a in framesAs)
                        {
                            if (a.GetAttribute("title").Contains("Search"))
                            {
                                a.InvokeMember("CLICK");
                                break;
                            }
                        }

                        Task asyncTask = new Task(() =>
                        {
                            Thread.Sleep(500);
                        });

                        asyncTask.ContinueWith((a) =>
                        {
                            webBrowserTask = WebBrowserTask.ContinueStudentRenaming;
                            confidentLoad = true;
                            webBrowser_DocumentCompleted(webBrowser, null);

                        }, TaskScheduler.FromCurrentSynchronizationContext());

                        asyncTask.Start();

                        break;
                    }
                case WebBrowserTask.ContinueStudentRenaming:
                    {
                        HtmlElementCollection tds = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("td");

                        foreach (HtmlElement td in tds)
                        {
                            if (td.GetAttribute("headers").Contains("Eid") && td.InnerText.Contains(renamingStudentID))
                            {
                                HtmlElementCollection tdAs = td.GetElementsByTagName("a");

                                foreach (HtmlElement a in tdAs)
                                {
                                    string innerText = a.InnerText.Trim();
                                    if (innerText.Equals(renamingStudentID))
                                    {
                                        string href = a.GetAttribute("href");

                                        confidentLoad = true;
                                        webBrowserTask = WebBrowserTask.SetNewFirstNameAndLastName;
                                        webBrowser.Document.Window.Frames[1].Navigate(href);
                                    }
                                }


                            }
                        }

                        break;
                    }
                case WebBrowserTask.SetNewFirstNameAndLastName:
                    {
                        HtmlElementCollection inputCollection =
                            webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");

                        HtmlElement firstNameInput = webBrowser.Document.Window.Frames[1].Document.GetElementById("first-name");
                        HtmlElement lastNameInput = webBrowser.Document.Window.Frames[1].Document.GetElementById("last-name");

                        firstNameInput.SetAttribute("value", renamingStudentName);
                        lastNameInput.SetAttribute("value", renamingStudentLastname);

                        foreach (HtmlElement elem in from HtmlElement he in inputCollection where he.GetAttribute("name").Contains("eventSubmit_doSave") select he)
                        {
                            elem.InvokeMember("CLICK");
                        }

                        webBrowserTask = WebBrowserTask.Idle;
                        break;
                    }
                case WebBrowserTask.Idle:
                    break;
                default:
                    break;
            }
        }

        private void InvokeGroupDeleting()
        {
            HtmlElement deleteButtonInputElement2 = webBrowser.Document.Window.Frames[1].Document.GetElementById("delete-groups");
            deleteButtonInputElement2.InvokeMember("CLICK");
        }

        /// <summary>
        /// Creates new group with specified name and students to add
        /// </summary>
        /// <param name="studentIDs">Students IDs</param>
        /// <returns>The count of added students. Not implemented yet</returns>
        public int CreateNewGroup(string groupName, string[] studentIDs)
        {
            return CreateNewGroup(groupName, "", studentIDs);
        }

        /// <summary>
        /// Creates a new group with specified GroupName, Description, students IDs
        /// The method is Private, because Description adding is not implemented
        /// </summary>
        /// <param name="groupName">The name of creating group</param>
        /// <param name="groupDescription">The description of creating group</param>
        /// <param name="studentIDs">The student IDs to add in a new group</param>
        /// <returns>The count of added students. Not implemented yet</returns>
        private int CreateNewGroup(string groupName, string groupDescription, string[] studentIDs)
        {
            addingStudentsCount = 0;

            if (linkToSiteEditor == "") throw new Exception("Link to SiteEditor is empty. Can't create new group.");

            addingGroupName = groupName;
            addingGroupDescription = groupDescription;
            addingStudentIDs = studentIDs;

            groupSectionTask = GroupEditorSectionTask.CreateNewGroup;
            webBrowserTask = WebBrowserTask.OpenManageGroupsSection;
            webBrowser.Navigate(linkToSiteEditor);

            return addingStudentsCount;
        }

        public bool DeleteGroup(string groupName)
        {
            bool result = true;

            if (String.IsNullOrWhiteSpace(linkToSiteEditor)) throw new Exception("Link to SiteEditor is empty. Can't delete new group.");
            if (String.IsNullOrWhiteSpace(groupName)) throw new ArgumentNullException("The given GroupName is incorrect. Can't delete new group.");

            deletingGroupName = groupName;
            groupSectionTask = GroupEditorSectionTask.DeleteExistingGroup;
            webBrowserTask = WebBrowserTask.OpenManageGroupsSection;
            webBrowser.Navigate(linkToSiteEditor);

            return result;
        }


        public void DownloadStudentsAttachments()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Method works with html and gets all required information about students.
        /// Works with indexOfProcessingAssignment, it should be set.
        /// </summary>
        private StudentInfo[] ParseLoadingStudents(int assignmentIndex, bool gradecall)
        {
            HtmlElementCollection submissionTable = webBrowser.Document.Window.Frames[1].Document.Forms["listSubmissionsForm"].Document.GetElementsByTagName("table");

            foreach (HtmlElement table in submissionTable)
            {
                if (table.GetAttribute("className") == "listHier lines nolines")
                {
                    HtmlElementCollection trs = table.GetElementsByTagName("tr");
                    foreach (HtmlElement tr in trs)
                    {
                        HtmlElementCollection tds = tr.GetElementsByTagName("td");

                        string studentName = "";
                        string studentID = "";
                        string submitted = "";
                        string status = "";
                        string grade = "";
                        bool released = false;
                        bool attached = tr.InnerHtml.Contains("attachments.gif");
                        string gradeLink = "";

                        foreach (HtmlElement td in tds)
                        {
                            if (td.GetAttribute("headers") == "studentname")
                            {
                                string text = td.InnerText.Trim();
                                MatchCollection matches = Regex.Matches(text, @"\(([\w|\.|\s]*)\)"); // Gets student ID
                                studentID = matches[0].Groups[1].Value;
                                studentName = text.Substring(0, text.IndexOf(" ("));
                                gradeLink = td.GetElementsByTagName("a")[0].GetAttribute("href");
                            }
                            if (td.GetAttribute("headers") == "submitted")
                            {
                                submitted = td.InnerText;
                            }
                            if (td.GetAttribute("headers") == "status")
                            {
                                status = td.InnerText;
                            }
                            if (td.GetAttribute("headers") == "grade")
                            {
                                grade = td.InnerText;
                            }
                            if (td.GetAttribute("headers") == "gradereleased")
                            {
                                released = td.InnerHtml.Contains("checkon.gif");
                            }
                        }

                        if (gradecall)
                        {
                            if (studentID == gradeStudentID)
                            {
                                webBrowserTask = WebBrowserTask.GradeStudent;
                                confidentLoad = true;
                                webBrowser.Document.Window.Frames[1].Navigate(gradeLink);

                                break;
                            }
                        }
                        else
                        {
                            if (studentID != "" && studentName != "")
                            {
                                Dictionary<string, StudentInfo> sid = dctAssignmentItems.ElementAt(assignmentIndex).Value.StudentInfosDictionary;
                                StudentInfo si = new StudentInfo(studentName, studentID, submitted, status, grade, released, gradeLink, attached);
                                if (sid.ContainsKey(studentID))
                                {
                                    sid[studentID] = si;
                                }
                                else
                                {
                                    sid.Add(studentID, si);
                                }

                            }
                        }
                    }

                }
            }

            return gradecall? null: dctAssignmentItems.ElementAt(assignmentIndex).Value.StudentInfosDictionary.Values.ToArray();
        }

        public Dictionary<String, StudentInfo> GetStudentsInformation(string assignmentTitle)
        {
            return dctAssignmentItems[assignmentTitle].StudentInfosDictionary;
        }

        /// <summary>
        /// Returns an array of strings that contain student IDs.
        /// </summary>
        /// <param name="assignmentTitle">The title of assignment</param>
        /// <returns></returns>
        public String[] GetStudentIDs()
        {
            return dctAssignmentItems.ElementAt(0).Value.StudentInfosDictionary.Keys.ToList().OrderBy(q => q).ToArray();
        }

        public String[] GetStudentNames()
        {
            List<String> names = new List<String>();
            String[] studIDs = GetStudentIDs();
            Dictionary<String, StudentInfo> studentInfos = dctAssignmentItems.ElementAt(0).Value.StudentInfosDictionary;
            foreach (String studentID in studIDs)
            {
                names.Add(studentInfos[studentID].Name.Replace(",", ""));
            }
            return names.ToArray();
        }

        public String GetStudentIdByName(String name)
        {
            String studID = "";
            String[] studIDs = GetStudentIDs();
            Dictionary<String, StudentInfo> studentInfos = dctAssignmentItems.ElementAt(0).Value.StudentInfosDictionary;
            foreach (String studentID in studIDs)
            {
                if (studentInfos[studentID].Name.Replace(",", "") == name)
                {
                    studID = studentID;
                    break;
                }
            }
            return studID;
        }

        /// <summary>
        /// Parse students from assignment
        /// </summary>
        public void ParseStudents()
        {
            ParseStudents(dctAssignmentItems.Keys.ToArray()[0]);
        }
        /// <summary>
        /// Parse students from selected assignment
        /// </summary>
        /// <param name="assignment"></param>
        public void ParseStudents(String assignment)
        {
            indexOfProcessingAssignment = -1;
            String[] assignments = dctAssignmentItems.Keys.ToArray();
            for (int i = 0; i < assignments.Length; i++)
            {
                if (assignments[i] == assignment)
                {
                    indexOfProcessingAssignment = i;
                    break;
                }
            }
            webBrowserTask = WebBrowserTask.ReloadStudents;

            if (dctAssignmentItems.Count >= 1) // Load into the main frame students
            {
                dctStudentInfos.Clear();
                confidentLoad = true;
                webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.Link);
            }
        }

        public void UpdateAssignmentSubmissions(string assignmentTitle)
        {
            // We should find indexOfProcessingAssignment for assignmentTitle
            indexOfProcessingAssignment = -1;
            for (int i = 0; i < dctAssignmentItems.Count; i++ )
            {
                string keyTrimed = dctAssignmentItems.ElementAt(i).Key.Trim();
                if (keyTrimed == assignmentTitle)
                {
                    indexOfProcessingAssignment = i;
                    break;
                }
            }
            if (indexOfProcessingAssignment == -1)
                return;
            webBrowserTask = WebBrowserTask.ReloadStudents;

            if (dctAssignmentItems.Count >= 1) // Load students into the main frame
            {
                dctStudentInfos.Clear();
                confidentLoad = true;
                webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.Link);
            }
        }

        /// <summary>
        /// Gets the names of assignments in the specified Site (Worksite).
        /// </summary>
        /// <returns>List of Assignments</returns>
        public Assignment[] GetAssignmentItems()
        {
            return dctAssignmentItems.Values.ToArray();
        }

        public string[] GetAssignmentItemNames()
        {
            return dctAssignmentItems.Values.Select(x => x.Title).ToArray();
        }

        public void GradeStudent(string assignmentName, string studentID, String mark)
        {
            Assignment assignment = dctAssignmentItems[assignmentName];
            studentmark = mark;
            webBrowserTask = WebBrowserTask.SelectStudentToGrade;
            gradeStudentID = studentID;
            confidentLoad = true;
            webBrowser.Document.Window.Frames[1].Navigate(assignment.Link);
        }

        /// <summary>
        /// Starts asynchronous student attachments parsing process
        /// </summary>
        public void ParseStudentAttachmentsAsync()
        {

        }

        /// <summary>
        /// Opens Users tab. Web browser should be navigated to Administration Workspace
        /// </summary>
        public void OpenAdministrationWorkspaceUsersTab()
        {
            webBrowserTask = WebBrowserTask.GoToAdministrationWorkspaceUsersTab;
            confidentLoad = true;
            webBrowser_DocumentCompleted(webBrowser, null);
        }

        /// <summary>
        /// Renames User. Web browser should be navigated to the User tab in Administration Workspace
        /// </summary>
        public void RenameUser(string id, string firstName, string lastName)
        {
            renamingStudentLastname = lastName;
            renamingStudentName = firstName;
            renamingStudentID = id;

            webBrowserTask = WebBrowserTask.RenameStudent;
            confidentLoad = true;
            webBrowser_DocumentCompleted(webBrowser, null);
        }

        /// <summary>
        /// Starts asynchronous assignment items parsing process
        /// </summary>
        public void ParseAssignmentItems()
        {
            webBrowserTask = WebBrowserTask.ParseAssignments;
            webBrowser.Navigate(linkToAssignments);
        }

        public String[] GetWorksites()
        {
            return dctWorksites.Keys.ToArray();
        }

        public StudentInfo[] GetUserSubmissions(string assignmentName)
        {
            return dctAssignmentItems[assignmentName].StudentInfosDictionary.Values.ToArray();
        }

        private void LogIn()
        {
            HtmlWindow loginFrame = webBrowser.Document.Window.Frames[0];
            loginFrame.Document.GetElementById("eid").SetAttribute("value", UserName);
            loginFrame.Document.GetElementById("pw").SetAttribute("value", Password);
            loginFrame.Navigate("javascript:document.forms[0].submit()");
            ReadWorksitesAsync();
        }
    }
}

///////////////////////////////////////////////////////////////////////////////
//
// SakaiParser.cs
// https://github.com/aeperepelitsyn/SakaiParser
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
// version 1.19 (20160403) Added event TestsAndQuizzesReady and base class CourseTools for worksite elements (moskalenkoBV)
// version 1.20 (20160409) Added event StudentGraded that provide grade message (moskalenkoBV) and new related exception messages (vystorobets)
// version 1.21 (20160503) Added SetDelayOfTestDueDate that modifies due date of test and event DelayOfTestAssigned (moskalenkoBV)
// version 1.22 (20160530) Added method AddAssignmentItem that adds new assignment item and event AddNewAssignmentItem (moskalenkoBV)
// version 1.23 (20160612) Fixed issue with Drafts in Assignments (moskalenkoBV)
// version 1.24 (20160915) Added method CreateUser and base class SakaiWebParser. Fixed issue with Idle state (aeperepelitsyn)
// version 1.25 (20170125) Added methods RemoveParticipants and LogOut. Implemented adding of participants during creation of group (aeperepelitsyn)
// version 1.26 (20170426) Fixed issue with Drafts in Assignments. Added methods ReadSubmissions, WriteSubmissions and SetTaskAndCallAgain (aeperepelitsyn)
// version 1.27 (20170529) Fixed issue with SelectWorksite for worksite without Assignments. (aeperepelitsyn)
//
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SakaiParser
{
    public static class SakaiParserVersion
    {
        public const string Version = "version 1.27 (20170529)";
        public const string Product = "SakaiParser";
    }

    public class CourseTools
    {
        public string Title { get; set; }

        public static implicit operator String(CourseTools a)
        {
            return a.Title;
        }

        public override string ToString()
        {
            return Title;
        }
    }
    public class Assignment : CourseTools
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
            Draft = (status == "Draft");
        }

        public Dictionary<string, StudentInfo> StudentInfosDictionary { get; set; }

        public string Link { get; set; }

        public string Status { get; set; }
        public string Open { get; set; }
        public string Due { get; set; }
        public string InNew { get; set; }
        public string Scale { get; set; }
        public bool Draft { get; set; }

    }

    public class TestsAndQuizzes : CourseTools
    {
        public TestsAndQuizzes(string title)
        {
            Title = title;
        }

    }
    public struct SubmittedFile
    {
        public string Name;
        public string Link;
    }

    public class UserInfo
    {
        public UserInfo()
        {
            
        }
        public UserInfo(string userID, string userFirstName, string userLastName, string userEmail, string userPassword, string userRole)
        {
            ID = userID;
            FirstName = userFirstName;
            LastName = userLastName;
            Email = userEmail;
            Password = userPassword;
            Rple = userRole;
        }
        public string ID;
        public string FirstName;
        public string LastName;
        public string Email;
        public string Password;
        public string Rple;
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
        public string TutorComments { get; set; }
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
        AddAssignmentItems,
        CountinueAddAssignmentItems,
        AddAssignmentItemsResultMessage,
        GoToAdministrationWorkspaceUsersTab,
        ParseAssignments,
        ParseTestsAndQuizzes,
        OpenTestAndQuizzesSettings,
        SetTestAndQuizzesDueDate,
        LoadStudents,
        ReloadStudents,
        LoadStudentAttachments,
        ReadStudentSubmissions,
        ReadIndividualSubmission,
        WaitIndividualSubmission,
        ReadSubmission,
        OpenManageGroupsSection,
        LoadGroupsEditor,
        CreateUser,
        RenameStudent,
        ContinueStudentRenaming,
        SetNewUserInfo,
        SelectStudentToGrade,
        GradeStudent,
        GradeResultMessage,
        AddNewGroup,
        NewGroupAdded,
        AddParticipanUsernames,
        RemovingParticipants,
        RemovedParticipants,
        SetNewFirstNameAndLastName,
        Waiting
    }

    public enum SPExceptions
    {
        UnableToFindMembership,
        UnableToFindLinkToUsersTab,
        WorksiteNameAlreadyExist,
        GradeFrameWasntFound,
        GradeMessageWasntFound,
        GradeUnsuccessful,
        TwoAssignmentsWithTheSameName
    }               
    public delegate void DelegateException(SPExceptions exception, String message);
    public delegate void DelegateWorksitesReady(String[] worksites);
    public delegate void DelegateWorksiteSelected(String worksiteName);
    public delegate void DelegateAssignmentItemsReady(String[] assignmentTitles, bool[] assignmentDrafts);
    public delegate void DelegateStudentsInformationReady(String[] studentIDs); 
    public delegate void DelegateTestsAndQuizzesReady(String[] testsName);
    public delegate void DelegateStudentGraded(bool success, String releaseMessage);
    public delegate void DelegateDelayOfTestAssigned(String testName, uint delay);
    public delegate void DelegateAddNewAssignmentItem(String message);
    public delegate void DelegateNewGroupCreated(String message);
    public delegate void DelegateParticipantsRemoved(String[] removed);
    public delegate void DelegateSubmissionsReady(StudentInfo[] studentsInfo);

    public class SakaiWebParser
    {
        protected WebBrowser webBrowser;

        protected WebBrowserTask webBrowserTask, webBrowserNextTask = WebBrowserTask.Idle;

        protected bool confidentLoad;

        protected UserInfo creatingUser;

        protected virtual void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            
        }

        protected void SetTaskAndCallAgain(WebBrowserTask task)
        {
            confidentLoad = true;
            webBrowserTask = task;
            webBrowser_DocumentCompleted(webBrowser, null);
        }

        /// <summary>
        /// Creates User. Web browser should be navigated to the User tab in Administration Workspace
        /// </summary>
        public void CreateUser(string userId, string userFirstName, string userLastName, string userEmail, string userPassword, string userRole)
        {
            CreateUser(new UserInfo(userId, userFirstName, userLastName, userEmail, userPassword, userRole));
        }
        public void CreateUser(UserInfo userInfo)
        {
            creatingUser = userInfo;
            SetTaskAndCallAgain(WebBrowserTask.CreateUser);
        }
    }

    public class SakaiParser285 : SakaiWebParser
    {
        delegate void ConfirmDeleting();

        ConfirmDeleting confirmDeletingVoid;
        ConfirmDeleting confirmInvokeAssignment;
        ConfirmDeleting confirmAssignmentItemMessage;


        GroupEditorSectionTask groupSectionTask;

        Dictionary<string, string> dctWorksites;
        Dictionary<string, StudentInfo> dctStudentInfos;
        Dictionary<string, Assignment> dctAssignmentItems;
        Dictionary<string, TestsAndQuizzes> dctTestAndQuizzesItems;

        Queue<WebBrowserTask> spTasks;

        string worksiteName;
        string linkToMembership;
        string linkToManageGroupsSection;
        string linkToCreateNewGroupSection;

        private string linkToAssignments;           // Link to Assignments page
        private string linkToSiteEditor;            // Link to SiteEditor page
        private string linkToTestsAndQuizzes;       // Link to TestsAndQuizzes page

        private string AssignmentItemTitle;
        private string AssignmentItemDecription;
        private double AssignmentItemGrade;
        private string AssignmentItemOpenDay;
        private string AssignmentItemOpenMonth;
        private string AssignmentItemOpenYear;
        private string AssignmentItemDueDay;
        private string AssignmentItemDueMonth;
        private string AssignmentItemDueYear;
        private string AssignmentItemCloseDay;
        private string AssignmentItemCloseMonth;
        private string AssignmentItemCloseYear;

        private string AssignmentNoSuccessMessage;

        private string AssignmentName;
        private string[] AssignmentStudentIDs;
        private Dictionary<String, String> AssignmentUsersTutorComment;

        private String currentTestName;
        private uint currentTestDelay;

        private string gradeStudentID;
        private String studentmark;

        int addingStudentsCount;

        string addingGroupName;
        string addingGroupDescription;

        string renamingStudentName;
        string renamingStudentID;
        string renamingStudentLastname;

        private List<string> StudentIDs;
        string[] addingStudentIDs;
        string[] addingParticipantIDs;
        string[] removingParticipantIDs;

        string deletingGroupName;

        int indexOfProcessingAssignment;
        int indexOfProcessingStudent;

        
        bool attachmentPresent;

        //bool assignmentsParsed;

        public event DelegateException SPException;

        private void SPExceptionProvider(SPExceptions exception)
        {
            SPExceptionProvider(exception, String.Empty);
        }

        private void SPExceptionProvider(SPExceptions exception, String details)
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
                case SPExceptions.GradeFrameWasntFound:
                    message = "Error! Frame wasn't found"; break;
                case SPExceptions.GradeMessageWasntFound:
                    message = "Error! Message wasn't found"; break;
                case SPExceptions.GradeUnsuccessful:
                    message = "Unsuccessful grading"; break;
                case SPExceptions.TwoAssignmentsWithTheSameName:
                    message = "Two Assignments with the same name"; break;
                default:
                    message = "Unknown exception"; break;
            }
            message += ": " + details;
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
                AssignmentItemsReady(GetAssignmentItemNames(), GetAssignmentItemDrafts());
            }
        }

        public event DelegateTestsAndQuizzesReady TestsAndQuizzesReady;
        private void TestsAndQuizzesReadyProvider()
        {
            if (TestsAndQuizzesReady != null)
            {
                TestsAndQuizzesReady(GetPublishedTestNames());
            }
        }
        public event DelegateDelayOfTestAssigned DelayOfTestAssigned;
        private void DelayOfTestAssignedProvider()
        {
            if (DelayOfTestAssigned != null)
            {
                DelayOfTestAssigned(currentTestName, currentTestDelay);
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
        public event DelegateStudentGraded StudentGraded;
        private void StudentGradedProvider(bool success, String releaseMessage)
        {
            if (StudentGraded != null)
            {
                StudentGraded(success, releaseMessage);
            }
        }
        public event DelegateAddNewAssignmentItem AddNewAssignmentItem;
        private void AddNewAssignmentItemProvider(String message)
        {
            if (AddNewAssignmentItem != null)
            {
                AddNewAssignmentItem(message);
            }
        }
        public event DelegateNewGroupCreated NewGroupCreated;
        private void NewGroupCreatedProvider(String message)
        {
            if (NewGroupCreated != null)
            {
                NewGroupCreated(message);
            }
        }
        public event DelegateParticipantsRemoved ParticipantsRemoved;
        private void ParticipantsRemovedProvider(String[] removed)
        {
            if (ParticipantsRemoved != null)
            {
                ParticipantsRemoved(removed);
            }
        }
        public event DelegateSubmissionsReady SubmissionsReady;
        private void SubmissionsReadyProvider(StudentInfo[] studentsInfo)
        {
            if (SubmissionsReady != null)
            {
                SubmissionsReady(studentsInfo);
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
            //assignmentsParsed = false;
            dctWorksites = new Dictionary<string, string>();
            dctAssignmentItems = new Dictionary<string, Assignment>();
            dctStudentInfos = new Dictionary<string, StudentInfo>();
            dctTestAndQuizzesItems = new Dictionary<string, TestsAndQuizzes>();
            spTasks = new Queue<WebBrowserTask>();
            worksiteName = "";
            linkToMembership = "";  
            SPException = null;
            webBrowser.ScriptErrorsSuppressed = true;
            this.webBrowser = webBrowser;
            webBrowser.DocumentCompleted += webBrowser_DocumentCompleted;

            confirmDeletingVoid = new ConfirmDeleting(InvokeGroupDeleting);
            confirmInvokeAssignment = new ConfirmDeleting(InvokeAssignmentItemPost);
            confirmAssignmentItemMessage = new ConfirmDeleting(GetAssignmentItemMessage);

        }

        private string InitialUrl { get; set; }
        private string UserName { get; set; }
        private string Password { get; set; }

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
        public void ReadWorksites()
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

        protected override void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowserTask == WebBrowserTask.Busy) webBrowserTask = WebBrowserTask.Idle;
            // if (e.Url.AbsolutePath != (sender as WebBrowser).Url.AbsolutePath) return;
            // //The page is finished loading 
            // or condition
            // (webBrowser.ReadyState != WebBrowserReadyState.Complete)

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
                case WebBrowserTask.LogIn:
                    LogIn();
                    break;
                case WebBrowserTask.ParseWorksites:
                    HtmlElement worksiteTable = webBrowser.Document.Window.Frames[0].Document.GetElementById("currentSites");
                    if (worksiteTable != null)
                    {
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
                    // Now we are at Home of the WorkSite
                    linkToAssignments = "";
                    linkToSiteEditor = "";
                    linkToTestsAndQuizzes = "";
                    HtmlElementCollection links = webBrowser.Document.GetElementsByTagName("a");
                    foreach (HtmlElement link in links)
                    {
                        if (link.GetAttribute("className") == "icon-sakai-assignment-grades")
                        {
                            linkToAssignments = link.GetAttribute("href");
                        }
                        else if (link.GetAttribute("className") == "icon-sakai-siteinfo")
                        {
                            // Parsing site editor link.
                            linkToSiteEditor = link.GetAttribute("href");
                        }
                        else if (link.GetAttribute("className") == "icon-sakai-samigo")
                        {
                            linkToTestsAndQuizzes = link.GetAttribute("href");
                        }
                    }
                    foreach (HtmlElement link in links)
                    {
                        if (link.GetAttribute("className") == "icon-sakai-iframe-site" && linkToAssignments == "")
                        {
                            link.InvokeMember("CLICK");
                        }
                    }
                    if (linkToSiteEditor == "" && worksiteName != "Administration Workspace")
                    {
                        webBrowserTask = WebBrowserTask.ParseSelectedWorksite;
                        // if (linkToSiteEditor == "") throw new Exception("Unable to find SiteEditor link");
                        confidentLoad = true;
                    }
                    else
                    {
                        // If it is okay, we have linkToSiteEditor
                        //if (linkToAssignments == "") throw new Exception("Unable to find Assignments link");
                        webBrowserTask = WebBrowserTask.Idle;
                        WorksiteSelectedProvider();
                    }
                    break;
                case WebBrowserTask.GoToAssignments:
                    //////////////////////////////////////////////////////////////////////////////////
                    /*webBrowserTask = WebBrowserTask.ParseAssignments;
                    webBrowser.Navigate(linkToAssignments);*/
                    //////////////////////////////////////////////////////////////////////////////////

                    break;
                case WebBrowserTask.ParseAssignments:
                    // To fix
                    HtmlElementCollection assignmentsTable = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("table");//my god
                    if (assignmentsTable.Count > 1)
                    {
                        HtmlElementCollection kostls = webBrowser.Document.GetElementsByTagName("a");
                        foreach (HtmlElement kostl in kostls)
                        {
                            if (kostl.GetAttribute("title") == "Reset")
                            {
                                kostl.InvokeMember("CLICK");
                                confidentLoad = true;
                                break;
                            }
                        }
                    }
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
                                bool draft = false;

                                foreach (HtmlElement td in tds)
                                {
                                    if (td.GetAttribute("headers") == "title")
                                    {
                                        HtmlElement link = td.GetElementsByTagName("a")[0];
                                        title = link.InnerText.TrimEnd();
                                    }
                                    if (td.GetAttribute("headers") == "status")
                                    {
                                        status = td.InnerText.TrimEnd();
                                    }
                                    if (td.GetAttribute("headers") == "openDate")
                                    {
                                        open = td.InnerText.TrimEnd();
                                    }
                                    if (td.GetAttribute("headers") == "dueDate")
                                    {
                                        due = td.InnerText.TrimEnd();
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

                                            innew = td.InnerText.TrimEnd();
                                        }
                                        else
                                        {

                                        }
                                    }
                                    if (td.GetAttribute("headers") == "maxgrade")
                                    {
                                        scale = td.InnerText.TrimEnd();
                                    }
                                }

                                if (title != "" && status != "") //&& open != "" && url != "" && due != "" && innew != "" && scale != ""
                                {
                                    if (dctAssignmentItems.ContainsKey(title))
                                    {
                                        SPExceptionProvider(SPExceptions.TwoAssignmentsWithTheSameName, title);
                                    }
                                    else
                                    {
                                        dctAssignmentItems.Add(title, new Assignment(title, url, status, open, due, innew, scale));
                                    }
                                }
                            }
                        }
                    }
                    // Assignments are supposed to be parsed
                    //assignmentsParsed = true;
                    if (dctAssignmentItems.Count == 0)
                    {

                    }
                    webBrowserTask = WebBrowserTask.Idle;
                    AssignmentItemsReadyProvider();
                    break;
                // We have all assignments. Go to LoadStudents
                case WebBrowserTask.AddAssignmentItems:
                    webBrowserTask = WebBrowserTask.Waiting;
                    HtmlElementCollection AssignmentMenuLinks = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("a");
                    foreach (HtmlElement AssignmentMenuLink in AssignmentMenuLinks)
                    {
                        if (AssignmentMenuLink.GetAttribute("title") == "Add")
                        {
                            AssignmentMenuLink.InvokeMember("CLICK");
                        }
                    }
                    Task asyncTask1 = new Task(() =>
                        {
                            Thread.Sleep(1000);
                        });

                    asyncTask1.ContinueWith((a) =>
                    {
                        SetTaskAndCallAgain(WebBrowserTask.CountinueAddAssignmentItems);
                    }, TaskScheduler.FromCurrentSynchronizationContext());

                    asyncTask1.Start();
                    break;
                case WebBrowserTask.CountinueAddAssignmentItems:
                    webBrowserTask = WebBrowserTask.Waiting;
                    HtmlElementCollection asAreas = webBrowser.Document.Window.Frames[1].Document.Window.Frames[0].Document.Window.Frames[0].Document.GetElementsByTagName("body");
                    foreach (HtmlElement asArea in asAreas)
                    {
                        asArea.InnerHtml = AssignmentItemDecription;
                    }
                    HtmlElementCollection asSelects = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("select");
                    foreach (HtmlElement asSelect in asSelects)
                    {
                        if (asSelect.Id == "new_assignment_openmonth")
                        {
                            HtmlElementCollection openmonths = asSelect.Children;
                            foreach (HtmlElement openmonth in openmonths)
                            {
                                if (openmonth.GetAttribute("value") == (AssignmentItemOpenMonth).ToString())
                                {
                                    openmonth.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_openday")
                        {
                            HtmlElementCollection opendays = asSelect.Children;
                            foreach (HtmlElement openday in opendays)
                            {
                                if (openday.GetAttribute("value") == (AssignmentItemOpenDay).ToString())
                                {
                                    openday.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_openyear")
                        {
                            HtmlElementCollection openyears = asSelect.Children;
                            foreach (HtmlElement openyear in openyears)
                            {
                                if (openyear.GetAttribute("value") == (AssignmentItemOpenYear).ToString())
                                {
                                    openyear.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_duemonth")
                        {
                            HtmlElementCollection duemonths = asSelect.Children;
                            foreach (HtmlElement duemonth in duemonths)
                            {
                                if (duemonth.GetAttribute("value") == (AssignmentItemDueMonth).ToString())
                                {
                                    duemonth.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_closemonth")
                        {
                            HtmlElementCollection closemonths = asSelect.Children;
                            foreach (HtmlElement closemonth in closemonths)
                            {
                                if (closemonth.GetAttribute("value") == (AssignmentItemDueMonth).ToString())
                                {
                                    closemonth.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_dueday")
                        {
                            HtmlElementCollection opendays = asSelect.Children;
                            foreach (HtmlElement openday in opendays)
                            {
                                if (openday.GetAttribute("value") == (AssignmentItemDueYear).ToString())
                                {
                                    openday.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_closeday")
                        {
                            HtmlElementCollection closedays = asSelect.Children;
                            foreach (HtmlElement closeday in closedays)
                            {
                                if (closeday.GetAttribute("value") == (AssignmentItemDueYear).ToString())
                                {
                                    closeday.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_dueyear")
                        {
                            HtmlElementCollection dueyears = asSelect.Children;
                            foreach (HtmlElement dueyear in dueyears)
                            {
                                if (dueyear.GetAttribute("value") == (AssignmentItemDueYear).ToString())
                                {
                                    dueyear.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_closeyear")
                        {
                            HtmlElementCollection closeyears = asSelect.Children;
                            foreach (HtmlElement closeyear in closeyears)
                            {
                                if (closeyear.GetAttribute("value") == (AssignmentItemDueYear).ToString())
                                {
                                    closeyear.SetAttribute("selected", "selected");
                                }
                            }
                        }
                        if (asSelect.Id == "new_assignment_grade_type")
                        {
                            HtmlElementCollection gradeTypes = asSelect.Children;
                            foreach (HtmlElement gradeType in gradeTypes)
                            {
                                if (gradeType.GetAttribute("value") == "3")
                                {
                                    gradeType.SetAttribute("selected", "selected");
                                }
                            }
                            asSelect.InvokeMember("onchange");
                        }
                        HtmlElementCollection txts = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
                        foreach (HtmlElement asName in txts)
                        {
                            if (asName.GetAttribute("id") == "new_assignment_title")
                            {
                                asName.SetAttribute("value", AssignmentItemTitle);
                            }
                            if (asName.Id == "new_assignment_grade_points")
                            {
                                asName.SetAttribute("value", AssignmentItemGrade.ToString());
                            }
                        }
                    }
                    Task asyncTask2 = new Task(() =>
                        {
                            Thread.Sleep(1000);
                        });

                    asyncTask2.ContinueWith((a) =>
                    {
                        InvokeAssignmentItemPost();

                    }, TaskScheduler.FromCurrentSynchronizationContext());

                    asyncTask2.Start();
                    break;
                case WebBrowserTask.AddAssignmentItemsResultMessage:
                    webBrowserTask = WebBrowserTask.Waiting;
                    Task asyncTask4 = new Task(() =>
                        {
                            Thread.Sleep(1000);
                        });

                    asyncTask4.ContinueWith((a) =>
                    {

                        GetAssignmentItemMessage();

                    }, TaskScheduler.FromCurrentSynchronizationContext());

                    asyncTask4.Start();
                    webBrowserTask = WebBrowserTask.Idle;
                    break;
                case WebBrowserTask.ParseTestsAndQuizzes:
                    dctTestAndQuizzesItems.Clear();
                    int ii = 0;
                    HtmlElement form = webBrowser.Document.Window.Frames[1].Document.GetElementById("authorIndexForm");
                    HtmlElementCollection testsName = form.Document.GetElementsByTagName("td");
                    foreach (HtmlElement testName in testsName)
                    {
                        if (testName.GetAttribute("className") == "titlePub")
                        {
                            dctTestAndQuizzesItems.Add(ii.ToString(), new TestsAndQuizzes(testName.InnerText));
                            ii++;
                        }
                    }

                    webBrowserTask = WebBrowserTask.Idle;
                    TestsAndQuizzesReadyProvider();
                    break;
                case WebBrowserTask.OpenTestAndQuizzesSettings:
                    Regex reginactive = new Regex("inactivePublishedSelectAction");
                    Regex regactive = new Regex("publishedSelectAction");
                    Regex reg1 = new Regex("inactivePublishedSelectAction[0-9]+");
                    Regex reg2 = new Regex("publishedSelectAction[0-9]+");
                    int numLink = 0;
                    HtmlElement testAndQuizzesItem;
                    HtmlElementCollection testAndQuizzesLinks = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("a");
                    //HtmlElementCollection testAndQuizzesSelect;
                    HtmlElementCollection testAndQuizzesOptions;
                    String testAndQuizzesStr = "";
                    HtmlElement testAndQuizessForm = webBrowser.Document.Window.Frames[1].Document.GetElementById("authorIndexForm");
                    HtmlElementCollection testAndQuizzesItems = testAndQuizessForm.Document.GetElementsByTagName("td");
                    foreach (HtmlElement item in testAndQuizzesItems)
                    {
                        if (item.GetAttribute("className") == "titlePub")
                        {
                            if (item.InnerText == currentTestName)
                            {
                                testAndQuizzesItem = item.Parent;
                                testAndQuizzesItem = testAndQuizzesItem.FirstChild;
                                testAndQuizzesItem = testAndQuizzesItem.FirstChild;
                                testAndQuizzesOptions = testAndQuizzesItem.Children;
                                foreach (HtmlElement option in testAndQuizzesOptions)
                                {
                                    if (option.GetAttribute("value").ToString() == "settings_published")
                                    {
                                        option.SetAttribute("selected", "selected");
                                        if (reginactive.IsMatch(testAndQuizzesItem.Id))
                                        {
                                            testAndQuizzesStr = reg1.Replace(testAndQuizzesItem.Id, "inactivePublishedHiddenlink");
                                        }
                                        else
                                        {
                                            testAndQuizzesStr = reg2.Replace(testAndQuizzesItem.Id, "publishedHiddenlink");
                                        }
                                        foreach (HtmlElement link in testAndQuizzesLinks)
                                        {
                                            if (link.Id == testAndQuizzesStr)
                                            {
                                                webBrowser.Navigate("javascript:window.frames[1].document.links[" + numLink + "].click()");
                                            }
                                            else
                                            {
                                                numLink++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    webBrowserTask = WebBrowserTask.SetTestAndQuizzesDueDate;
                    confidentLoad = true;
                    break;
                case WebBrowserTask.SetTestAndQuizzesDueDate:
                    var culture = new CultureInfo("en-US");
                    DateTime localDate = DateTime.Now;
                    DateTime newdate = DateTime.Now;
                    TimeSpan timer = new TimeSpan(0, (int)currentTestDelay, 0);
                    newdate = newdate.Add(timer);
                    HtmlElementCollection dates = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
                    foreach (HtmlElement date in dates)
                    {
                        if (date.GetAttribute("id") == "assessmentSettingsAction:startDate")
                        {
                            date.SetAttribute("value", localDate.ToString(culture));
                        }
                        if (date.GetAttribute("id") == "assessmentSettingsAction:endDate")
                        {
                            date.SetAttribute("value", newdate.ToString(culture));
                        }
                    }
                    HtmlElement testSettings = webBrowser.Document.Window.Frames[1].Document.GetElementById("assessmentSettingsAction");
                    HtmlElementCollection testSettingsElems = testSettings.All;
                    foreach (HtmlElement elem in testSettingsElems)
                    {
                        if (elem.GetAttribute("value") == "Save Settings")
                        {
                            elem.InvokeMember("CLICK");
                        }
                    }
                    webBrowserTask = WebBrowserTask.Idle;
                    DelayOfTestAssignedProvider();
                    break;
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
                ////////////////////////////////////////////////////////////////
                case WebBrowserTask.ReadStudentSubmissions:
                    {
                        ParseLoadingStudents(AssignmentName, false);
                        StudentIDs = dctAssignmentItems[AssignmentName].StudentInfosDictionary.Keys.ToList();

                        if (AssignmentStudentIDs != null ? AssignmentStudentIDs.Length > 0 : false)
                        {
                            foreach (string studentId in StudentIDs)
                            {
                                if (!AssignmentStudentIDs.Contains(studentId))
                                {
                                    dctAssignmentItems[AssignmentName].StudentInfosDictionary.Remove(studentId);
                                }
                            }
                            StudentIDs = dctAssignmentItems[AssignmentName].StudentInfosDictionary.Keys.ToList();
                            if (AssignmentUsersTutorComment != null)
                            {
                                foreach (string studentId in AssignmentStudentIDs)
                                {
                                    if (!StudentIDs.Contains(studentId))
                                    {
                                        AssignmentUsersTutorComment.Remove(studentId);
                                    }
                                }
                            }
                            AssignmentStudentIDs = StudentIDs.ToArray();
                        }
                        SetTaskAndCallAgain(WebBrowserTask.ReadIndividualSubmission);
                        break;
                    }
                case WebBrowserTask.ReadIndividualSubmission:
                    {
                        if (StudentIDs.Count > 0)
                        {
                            webBrowserTask = WebBrowserTask.WaitIndividualSubmission;
                            confidentLoad = true;
                            webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems[AssignmentName].StudentInfosDictionary[StudentIDs[0]].GradeLink);
                        }
                        else
                        {
                            webBrowserTask = WebBrowserTask.Idle;
                            if (AssignmentUsersTutorComment != null)
                            {
                                AssignmentUsersTutorComment = null;
                            }
                            SubmissionsReady(dctAssignmentItems[AssignmentName].StudentInfosDictionary.Values.ToArray());
                        }
                        break;
                    }
                case WebBrowserTask.WaitIndividualSubmission:
                {
                    Task asyncTaskSubmission = new Task(() =>
                    {
                        Thread.Sleep(1000);
                    });

                    asyncTaskSubmission.ContinueWith((a) =>
                    {
                        SetTaskAndCallAgain(WebBrowserTask.ReadSubmission);
                    }, TaskScheduler.FromCurrentSynchronizationContext());

                    asyncTaskSubmission.Start();
                    break;
                }
                case WebBrowserTask.ReadSubmission:
                {
                    HtmlElementCollection textAreas = webBrowser.Document.Window.Frames[1].Document.Window.Frames[0].Document.Window.Frames[0].Document.GetElementsByTagName("body");
                    foreach (HtmlElement textArea in textAreas)
                    {
                        dctAssignmentItems[AssignmentName].StudentInfosDictionary[StudentIDs[0]].TutorComments = textArea.InnerText;
                        if (AssignmentUsersTutorComment != null)
                        {
                            textArea.InnerText = AssignmentUsersTutorComment[StudentIDs[0]];
                            SaveAndReturnToStudent();
                        }
                        else
                        {
                            StudentIDs.RemoveAt(0);
                            SetTaskAndCallAgain(WebBrowserTask.ReadIndividualSubmission);
                        }
                        break;
                    }
                    break;
                }
                ////////////////////////////////////////////////////////////////
                case WebBrowserTask.RemovingParticipants:
                    {
                        HtmlElementCollection lineParticipantsCollection =
                            webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("tbody");
                        List<String> idsSakaiCollection = new List<string>();
                        List<String> idsCollection = new List<string>();
                        foreach (HtmlElement lineParticipant in lineParticipantsCollection[1].Children)
                        {
                            String[] htmlParts = lineParticipant.InnerHtml.Split(new string[] { "<H5>", "<h5>", "</H5>", "</h5>" }, StringSplitOptions.RemoveEmptyEntries);
                            if (htmlParts.Length == 3)
                            {
                                String[] participantNames = htmlParts[1].Trim().Split(new char[] { ',', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                                String name = participantNames[participantNames.Length - 1].Trim();
                                if (removingParticipantIDs.Contains(name))
                                {
                                    String[] participantLineHTML = htmlParts[2].Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (String htmlStringPerts in participantLineHTML)
                                    {
                                        if (htmlStringPerts.StartsWith("id=remove_"))
                                        {
                                            idsSakaiCollection.Add(htmlStringPerts.Substring(3));
                                            idsCollection.Add(name);
                                        }
                                    }
                                }
                            }
                        }
                        for (int i = 0; i < idsSakaiCollection.Count; i++)
                        {
                            HtmlElement htmlCheckbox = webBrowser.Document.Window.Frames[1].Document.GetElementById(idsSakaiCollection[i]);
                            htmlCheckbox.InvokeMember("CLICK");
                        }
                        lineParticipantsCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
                        foreach (HtmlElement htmlInputElement in lineParticipantsCollection)
                        {
                            if (htmlInputElement.GetAttribute("accesskey") == "s")
                            {
                                removingParticipantIDs = idsCollection.ToArray();
                                htmlInputElement.InvokeMember("CLICK");
                                break;
                            }
                        }
                        confidentLoad = true;
                        webBrowserTask = WebBrowserTask.RemovedParticipants;
                        break;
                    }
                case WebBrowserTask.RemovedParticipants:
                    {
                        webBrowserTask = WebBrowserTask.Idle;
                        ParticipantsRemovedProvider(removingParticipantIDs);
                        removingParticipantIDs = new string[] { };
                        break;
                    }
                case WebBrowserTask.OpenManageGroupsSection:
                    // Todo: find link to Manage Groups section.
                    String[] IDs = GetStudentIDs();
                    ///////////////////////////////////
                    // GetParticipants
                    List<String> addingParticipants = new List<String>();
                    List<String> participantsID = new List<String>();
                    HtmlElementCollection lisParticipantsCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("h5");
                    foreach (HtmlElement htmlParticipantText in lisParticipantsCollection)
                    {
                        String[] participantNames = htmlParticipantText.InnerText.Trim().Split(new char[] { ',', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                        participantsID.Add(participantNames[participantNames.Length - 1].Trim());
                    }
                    for (int i = 0; i < addingStudentIDs.Length; i++)
                    {
                        if (!participantsID.Contains(addingStudentIDs[i]))
                        {
                            addingParticipants.Add(addingStudentIDs[i]);
                        }
                    }
                    addingParticipantIDs = addingParticipants.ToArray();
                    ///////////////////////////////////
                    linkToManageGroupsSection = "";
                    HtmlElementCollection lisElementCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("li");
                    foreach (HtmlElement element in lisElementCollection)
                    {
                        if (element.GetAttribute("role") == "menuitem" && element.InnerHtml.Contains(addingParticipantIDs.Length == 0 ? "doManageGroupHelper" : "doParticipantHelper"))
                        {
                            MatchCollection matchCollection = Regex.Matches(element.InnerHtml,
                                @"(onclick)\s*=\s*""\s*(location)\s*=\s*'(?<link>.*)'\s*;");
                            linkToManageGroupsSection = matchCollection[0].Groups["link"].Value;
                            break;
                        }
                    }

                    // linkToManageGroupsSection contains link to Manage Groups section. 
                    // We can navigate it right now.

                    webBrowser.Document.Window.Frames[1].Navigate(linkToManageGroupsSection);


                    //webBrowser.Navigate(linkToManageGroupsSection);
                    // LoadGroupsEditor is a page where we set name of group and add participants.
                    if (addingParticipantIDs.Length == 0)
                    {
                        webBrowserTask = WebBrowserTask.LoadGroupsEditor;
                    }
                    else
                    {
                        webBrowserTask = WebBrowserTask.AddParticipanUsernames;
                    }
                    confidentLoad = true; // The Frame URL of doesn't match Browser URL, beacuse Browser URL has nothing to do with frame.
                    //webBrowser_DocumentCompleted(webBrowser, e);

                    break;
                case WebBrowserTask.AddParticipanUsernames:
                    if (addingParticipantIDs.Length > 0)
                    {
                        HtmlElementCollection usernamesElementCollection =
                            webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("textarea");
                        String participantIDs = String.Empty;
                        foreach (String participantID in addingParticipantIDs)
                        {
                            participantIDs += participantID + "\n";
                        }
                        usernamesElementCollection[0].InnerText = participantIDs;
                        addingParticipantIDs = new string[] { };
                        HtmlElementCollection inputsCollection =
                            webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
                        foreach (HtmlElement inputElement in inputsCollection)
                        {
                            if (inputElement.Id == "content::continue")
                            {
                                webBrowser.ScriptErrorsSuppressed = true;
                                confidentLoad = true;
                                inputElement.InvokeMember("click");
                                break;
                            }
                        }
                    }
                    else
                    {
                        HtmlElementCollection inputsCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
                        foreach (HtmlElement htmlInputElement in inputsCollection)
                        {
                            if (htmlInputElement.Id == "content::role-row:2:role-select")
                            {
                                htmlInputElement.SetAttribute("checked", "checked");
                            }
                            else if (htmlInputElement.GetAttribute("accesskey") == "s")
                            {
                                //webBrowser.ScriptErrorsSuppressed = false;
                                confidentLoad = true;
                                if (inputsCollection.Count < 6)
                                {
                                    webBrowserTask = WebBrowserTask.OpenManageGroupsSection;
                                }
                                htmlInputElement.InvokeMember("click");
                                break;
                            }
                        }
                    }
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

                    confidentLoad = true;
                    webBrowserTask = WebBrowserTask.NewGroupAdded;
                    break;
                case WebBrowserTask.NewGroupAdded:
                    webBrowserTask = WebBrowserTask.Idle;
                    NewGroupCreated(addingGroupName);

                    break;
                case WebBrowserTask.SelectStudentToGrade:
                    // Current page contains the list of students.
                    // Let's find the link width "Grade" title

                    ParseLoadingStudents(0, true);

                    break;

                case WebBrowserTask.GradeStudent:
                    webBrowser.Document.Window.Frames[1].Document.GetElementById("grade").SetAttribute("value", studentmark);
                    SaveAndReturnToStudent();
                    break;
                case WebBrowserTask.GradeResultMessage:
                    {
                        bool success = false;
                        String releaseMessage = "";
                        HtmlElementCollection bRelease = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("div");
                        if (bRelease.Count != 0)
                        {
                            foreach (HtmlElement divs in bRelease)
                            {
                                if (divs.GetAttribute("className") == "success")
                                {
                                    success = true;
                                    releaseMessage = divs.InnerText;
                                    break;
                                }
                                else if (divs.GetAttribute("className") == "alertMessage")
                                {
                                    success = false;
                                    releaseMessage = divs.InnerText;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            SPExceptionProvider(SPExceptions.GradeFrameWasntFound);
                        }
                        if (releaseMessage == String.Empty)
                        {
                            SPExceptionProvider(SPExceptions.GradeMessageWasntFound);
                        }

                        if (StudentIDs != null ? StudentIDs.Count > 0 : false)
                        {
                            if (!success)
                            {
                                SPExceptionProvider(SPExceptions.GradeUnsuccessful, StudentIDs[0]);
                            }
                            else
                            {
                                dctAssignmentItems[AssignmentName].StudentInfosDictionary[StudentIDs[0]].TutorComments = AssignmentUsersTutorComment[StudentIDs[0]];
                            }
                            StudentIDs.RemoveAt(0);
                            SetTaskAndCallAgain(WebBrowserTask.ReadIndividualSubmission);
                        }
                        else
                        {
                            webBrowserTask = WebBrowserTask.Idle;
                            StudentGradedProvider(success, releaseMessage);
                        }
                        break;
                    }
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
                case WebBrowserTask.CreateUser:
                    {
                        //                        HtmlElement inputSearch = webBrowser.Document.Window.Frames[1].Document.GetElementById("search");

                        //                        if (inputSearch == null) throw new NullReferenceException("Input field is NULL.");

                        //                        inputSearch.Document.GetElementById("search").SetAttribute("value", renamingStudentID);

                        HtmlElementCollection framesAs = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("a");

                        foreach (HtmlElement htmlNewUser in framesAs)
                        {
                            if (htmlNewUser.GetAttribute("title").Contains("New User"))
                            {
                                htmlNewUser.InvokeMember("CLICK");
                                break;
                            }
                        }

                        Task asyncTask = new Task(() =>
                        {
                            Thread.Sleep(500);
                        });

                        asyncTask.ContinueWith((htmlNewUser) =>
                        {
                            SetTaskAndCallAgain(WebBrowserTask.SetNewUserInfo);
                        }, TaskScheduler.FromCurrentSynchronizationContext());

                        asyncTask.Start();

                        break;
                    }
                case WebBrowserTask.RenameStudent:
                    {
                        HtmlElement inputSearch = webBrowser.Document.Window.Frames[1].Document.GetElementById("search");

                        if (inputSearch == null) throw new NullReferenceException("Input field is NULL.");

                        inputSearch.Document.GetElementById("search").SetAttribute("value", renamingStudentID);

                        HtmlElementCollection framesAs = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("a");

                        foreach (HtmlElement htmlStudent in framesAs)
                        {
                            if (htmlStudent.GetAttribute("title").Contains("Search"))
                            {
                                htmlStudent.InvokeMember("CLICK");
                                break;
                            }
                        }

                        Task asyncTask = new Task(() =>
                        {
                            Thread.Sleep(500);
                        });

                        asyncTask.ContinueWith((htmlStudent) =>
                        {
                            SetTaskAndCallAgain(WebBrowserTask.ContinueStudentRenaming);
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
                case WebBrowserTask.SetNewUserInfo:
                    {
                        HtmlElementCollection inputCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");

                        HtmlElement userIDInput = webBrowser.Document.Window.Frames[1].Document.GetElementById("eid");
                        HtmlElement firstNameInput = webBrowser.Document.Window.Frames[1].Document.GetElementById("first-name");
                        HtmlElement lastNameInput = webBrowser.Document.Window.Frames[1].Document.GetElementById("last-name");
                        HtmlElement emailInput = webBrowser.Document.Window.Frames[1].Document.GetElementById("email");
                        HtmlElement pwInput = webBrowser.Document.Window.Frames[1].Document.GetElementById("pw");
                        HtmlElement pw0Input = webBrowser.Document.Window.Frames[1].Document.GetElementById("pw0");

                        userIDInput.SetAttribute("value", creatingUser.ID);
                        firstNameInput.SetAttribute("value", creatingUser.FirstName);
                        lastNameInput.SetAttribute("value", creatingUser.LastName);
                        emailInput.SetAttribute("value", creatingUser.Email);
                        pwInput.SetAttribute("value", creatingUser.Password);
                        pw0Input.SetAttribute("value", creatingUser.Password);

                        HtmlElementCollection selectCollection = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("select");
                        foreach (HtmlElement htmlSelect in selectCollection)
                        {
                            HtmlElementCollection optionsCollection = htmlSelect.GetElementsByTagName("option");
                            foreach (HtmlElement option in optionsCollection)
                            {
                                if (option.InnerText == creatingUser.Rple)
                                {
                                    option.SetAttribute("selected", "selected");
                                    break;
                                }
                            }
                        }

                        foreach (HtmlElement elem in from HtmlElement he in inputCollection where he.GetAttribute("name").Contains("eventSubmit_doSave") select he)
                        {
                            elem.InvokeMember("CLICK");
                        }

                        confidentLoad = true;
                        webBrowserTask = WebBrowserTask.Busy;
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

                        //                        webBrowserNextTask = 
                        webBrowserTask = WebBrowserTask.Busy;
                        break;
                    }
                case WebBrowserTask.Idle:
                    break;
                case WebBrowserTask.Waiting:
                    break;
                default:
                    break;
            }
        }

        private void SaveAndReturnToStudent()
        {
            HtmlElementCollection processBtn = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
            foreach (HtmlElement element in processBtn)
            {
                if (element.GetAttribute("name") == "return")
                {
                    webBrowserTask = WebBrowserTask.GradeResultMessage;
                    confidentLoad = true;
                    element.InvokeMember("CLICK");   
                    break;
                }
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

        /// <summary>
        /// Removes participants from current Worksite with available SiteEditor tool.
        /// </summary>
        /// <param name="idsForRemove">The array of IDs to remove.</param>
        public void RemoveParticipants(String[] idsForRemove)
        {
            removingParticipantIDs = idsForRemove;
            webBrowserTask = WebBrowserTask.RemovingParticipants;
            webBrowser.Navigate(linkToSiteEditor);
        }

        private void DownloadStudentsAttachments()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Method works with html and gets all required information about students.
        /// Works with indexOfProcessingAssignment, it should be set.
        /// </summary>
        private StudentInfo[] ParseLoadingStudents(int assignmentIndex, bool gradecall)
        {
            return ParseLoadingStudents(dctAssignmentItems.ElementAt(assignmentIndex).Key, gradecall);
        }
        private StudentInfo[] ParseLoadingStudents(string assignmentName, bool gradecall)
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
                            string text = (td.InnerText != null ? td.InnerText.TrimEnd() : String.Empty);
                            if (td.GetAttribute("headers") == "studentname")
                            {
                                MatchCollection matches = Regex.Matches(text, @"\(([\w|\.|\s]*)\)"); // Gets student ID
// To do: add check for cases unexpected formats (coused by "off#spring")
                                studentID = matches[0].Groups[1].Value;
                                studentName = text.Substring(0, text.IndexOf(" ("));
                                gradeLink = td.GetElementsByTagName("a")[0].GetAttribute("href");
                            }
                            else if (td.GetAttribute("headers") == "submitted")
                            {
                                submitted = text;
                            }
                            else if (td.GetAttribute("headers") == "status")
                            {
                                status = text;
                            }
                            else if (td.GetAttribute("headers") == "grade")
                            {
                                grade = text;
                            }
                            else if (td.GetAttribute("headers") == "gradereleased")
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
                                Dictionary<string, StudentInfo> sid = dctAssignmentItems[assignmentName].StudentInfosDictionary;
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

            return gradecall? null: dctAssignmentItems[assignmentName].StudentInfosDictionary.Values.ToArray();
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
            return dctAssignmentItems.Count > 0 ? dctAssignmentItems.ElementAt(0).Value.StudentInfosDictionary.Keys.ToList().OrderBy(q => q).ToArray() : new string[]{};
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

        private void ProcessSubmissions(string assignmentName, string[] studentIDs, Dictionary<String, String> usersTutorComment)
        {
            if (dctAssignmentItems.Keys.Contains(assignmentName) && (studentIDs != null ? studentIDs.Length > 0 : true))
            {
                confidentLoad = true;
                AssignmentName = assignmentName;
                AssignmentStudentIDs = studentIDs;
                AssignmentUsersTutorComment = usersTutorComment;
                dctAssignmentItems[assignmentName].StudentInfosDictionary.Clear();
                webBrowserTask = WebBrowserTask.ReadStudentSubmissions;
                webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems[assignmentName].Link);
            }
        }
        /// <summary>
        /// Starts asynchronous writing of submissions
        /// </summary>
        public void WriteSubmissions(string assignmentName, Dictionary<String, String> usersTutorComment)
        {
            ProcessSubmissions(assignmentName, usersTutorComment.Keys.ToArray(), usersTutorComment);
        }
        /// <summary>
        /// Starts asynchronous reading of submissions
        /// </summary>
        public void ReadSubmissions(string assignmentName, string[] studentIDs)
        {
            ProcessSubmissions(assignmentName, studentIDs, null);
        }
        /// <summary>
        /// Starts asynchronous reading of submissions
        /// </summary>
        public void ReadSubmissions(string assignmentName)
        {
            ProcessSubmissions(assignmentName, null, null);
        }

        /// <summary>
        /// Parse students from assignment
        /// </summary>
        public void ParseStudents()
        {
            if (worksiteName != "Administration Workspace")
            {
                ParseStudents(dctAssignmentItems.Keys.ToArray()[0]);
            }
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
        public bool[] GetAssignmentItemDrafts()
        {
            return dctAssignmentItems.Values.Select(x => x.Draft).ToArray();
        }
        public string[] GetPublishedTestNames()
        {
            return dctTestAndQuizzesItems.Values.Select(x => x.Title).ToArray();
        }
        public void SetDelayOfTestDueDate(String testName, uint delay)
        {
            if (linkToTestsAndQuizzes != String.Empty)
            {
                currentTestDelay = delay;
                currentTestName = testName;
                webBrowserTask = WebBrowserTask.OpenTestAndQuizzesSettings;
                confidentLoad = true;
                webBrowser.Navigate(linkToTestsAndQuizzes);
            }
            
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
            SetTaskAndCallAgain(WebBrowserTask.GoToAdministrationWorkspaceUsersTab);
        }



        /// <summary>
        /// Renames User. Web browser should be navigated to the User tab in Administration Workspace
        /// </summary>
        public void RenameUser(string id, string firstName, string lastName)
        {
            renamingStudentLastname = lastName;
            renamingStudentName = firstName;
            renamingStudentID = id;

            SetTaskAndCallAgain(WebBrowserTask.RenameStudent);
        }

        /// <summary>
        /// Starts asynchronous assignment items parsing process
        /// </summary>
        public void ParseAssignmentItems()
        {
            if (linkToAssignments != String.Empty)
            {
                webBrowserTask = WebBrowserTask.ParseAssignments;
                webBrowser.Navigate(linkToAssignments);
            }
        }
        /// <summary>
        /// Starts asynchronous assignment items parsing process
        /// </summary>
        public void ParseTestsAndQuizzesItems()
        {
            if (linkToTestsAndQuizzes != String.Empty)
            {
                webBrowserTask = WebBrowserTask.ParseTestsAndQuizzes;
                webBrowser.Navigate(linkToTestsAndQuizzes);
            }
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
            ReadWorksites();
        }
        public void LogOut()
        {
            String logoutURL = String.Empty;
            HtmlElement htmlElementLogin = webBrowser.Document.GetElementById("loginLinks");
            if (htmlElementLogin != null)
            {
                String[] logoutLinks = htmlElementLogin.Children[1].OuterHtml.Split('"');
                foreach (string logoutLink in logoutLinks)
                {
                    if (logoutLink.Contains(webBrowser.Document.Domain))
                    {
                        logoutURL = logoutLink;
                        break;
                    }
                }
            }
            if(logoutURL == String.Empty)
            {
                if (webBrowser.Document.Url.LocalPath.LastIndexOf('/') > 0)
                {
                    logoutURL = webBrowser.Document.Domain + webBrowser.Document.Url.LocalPath.Substring(0, webBrowser.Document.Url.LocalPath.IndexOf('/', 1)) + "/logout";
                }
            }
            webBrowser.Navigate(logoutURL);
            webBrowserTask = WebBrowserTask.Busy;
        }
        public void AddAssignmentItem(  string title, string description, double grade,
                                        string openday, string openmonth, string openyear,
                                        string dueday, string duemonth, string dueyear,
                                        string closeday, string closemonth, string closeyear)
        {
            AssignmentItemTitle = title;
            AssignmentItemDecription = description;
            AssignmentItemGrade = grade;
            AssignmentItemOpenDay = openday;
            AssignmentItemOpenMonth = openmonth;
            AssignmentItemOpenYear = openyear;
            AssignmentItemDueDay = dueday;
            AssignmentItemDueMonth = duemonth;
            AssignmentItemDueYear = dueyear;
            AssignmentItemCloseDay = closeday;
            AssignmentItemCloseMonth = closemonth;
            AssignmentItemCloseYear = closeyear;
            webBrowserTask = WebBrowserTask.AddAssignmentItems;
            webBrowser.Navigate(linkToAssignments);
        }
        private void InvokeAssignmentItemPost()
        {
            webBrowserTask = WebBrowserTask.Waiting;
            HtmlElementCollection asPosts = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("input");
            foreach (HtmlElement asPost in asPosts)
            {
                if (asPost.GetAttribute("name") == "post")
                {
                    asPost.InvokeMember("CLICK");
                }
            }
            confidentLoad = true;
            webBrowserTask = WebBrowserTask.AddAssignmentItemsResultMessage;
        }
        private void GetAssignmentItemMessage()
        {
            String message = "";
            HtmlElementCollection Release = webBrowser.Document.Window.Frames[1].Document.GetElementsByTagName("div");
            foreach (HtmlElement divs in Release)
            {
                if (divs.GetAttribute("className") == "alertMessage")
                {
                    message = divs.InnerText;
                    break;
                }
                else
                {
                    message = String.Empty;
                }
            }
            webBrowserTask = WebBrowserTask.Idle;
            AssignmentNoSuccessMessage = message;
            AddNewAssignmentItemProvider(message);
        }
    }
}

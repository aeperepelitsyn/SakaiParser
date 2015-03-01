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
//
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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

    public enum WebBrowserTask
    {
        Idle,
        LogIn,
        GetMembershipLink,
        ParseWorksites,
        GoToAssignments,
        ParseAssignments,
        LoadStudents,
        ReloadStudents,
        LoadStudentAttachments
    }

    class SakaiParser285
    {
        WebBrowser webBrowser;

        WebBrowserTask webBrowserTask;

        Dictionary<string, string> dctWorksites;
        Dictionary<string, StudentInfo> dctStudentInfos;
        Dictionary<string, Assignment> dctAssignmentItems;

        string worksiteName;
        string linkToMembership;

        int indexOfProcessingAssignment;
        int indexOfProcessingStudent;

        bool confidentLoad;
        bool attachmentPresent;

        public SakaiParser285(WebBrowser webBrowser, string initialUrl, string userName, string password)
        {
            InitialUrl = initialUrl;
            UserName = userName;
            Password = password;
            confidentLoad = false;
            dctWorksites = new Dictionary<string, string>();
            dctAssignmentItems = new Dictionary<string, Assignment>();
            dctStudentInfos = new Dictionary<string, StudentInfo>();
            worksiteName = "";
            linkToMembership = "";
            this.webBrowser = webBrowser;
            webBrowser.DocumentCompleted += webBrowser_DocumentCompleted;
        }

        public string InitialUrl { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        public void Initialize(string worksiteName)
        {
            SetWorksiteName(worksiteName);
            Initialize();
            // Than login when document is completed
        }

        public void Initialize()
        {
            webBrowserTask = WebBrowserTask.LogIn;
            webBrowser.Navigate(InitialUrl);
        }

        public void SetWorksiteName(string worksiteName)
        {
            this.worksiteName = worksiteName;
        }

        void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (!confidentLoad)
            {
                if (e.Url != webBrowser.Url) return; // If we have not reached destination URL
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

                    dctWorksites.Clear();

                    HtmlElement worksiteTable = webBrowser.Document.Window.Frames[0].Document.GetElementById("currentSites");
                    HtmlElementCollection tableTDs = worksiteTable.GetElementsByTagName("td");
                    foreach(HtmlElement worksiteTD in tableTDs)
                    {
                        if(worksiteTD.GetAttribute("headers") == "worksite")
                        {
                            HtmlElement linkToWorksite = worksiteTD.GetElementsByTagName("a")[0];
                            string sLink = linkToWorksite.GetAttribute("href");

                            dctWorksites.Add(worksiteTD.InnerText, sLink);
                        }
                    }

                    // Worksites in dctWorksites

                    webBrowserTask = WebBrowserTask.GoToAssignments;
                    webBrowser.Navigate(dctWorksites[worksiteName]);

                    break;
                case WebBrowserTask.GetMembershipLink:   
           
                    HtmlElementCollection linksCollection = webBrowser.Document.GetElementsByTagName("a");
                    foreach(HtmlElement link in linksCollection)
                        if (link.GetAttribute("className") == "icon-sakai-membership")
                            linkToMembership = link.GetAttribute("href");
                    if (linkToMembership == "") throw new Exception("Unable to find Membership link");

                    webBrowserTask = WebBrowserTask.ParseWorksites;
                    webBrowser.Navigate(linkToMembership);

                    break;
                case WebBrowserTask.GoToAssignments:
                    // Now we are at HOME link, maybe

                    string linkToAssignments = "";
                    HtmlElementCollection links = webBrowser.Document.GetElementsByTagName("a");
                    foreach(HtmlElement link in links)
                        if(link.GetAttribute("className") == "icon-sakai-assignment-grades")
                            linkToAssignments = link.GetAttribute("href");

                    if (linkToAssignments == "") throw new Exception("Unable to find Assignments link");

                    webBrowserTask = WebBrowserTask.ParseAssignments;
                    webBrowser.Navigate(linkToAssignments);

                    break;
                case WebBrowserTask.ParseAssignments:

                    HtmlElementCollection assignmentsTable = webBrowser.Document.Window.Frames[1].Document.Forms["listAssignmentsForm"].Document.GetElementsByTagName("table");
                    foreach(HtmlElement table in assignmentsTable)
                    {
                        if(table.GetAttribute("className") == "listHier lines nolines")
                        {
                            HtmlElementCollection trs = table.GetElementsByTagName("tr");
                            foreach(HtmlElement tr in trs)
                            {
                                HtmlElementCollection tds = tr.GetElementsByTagName("td");

                                string title = "";
                                string status = "";
                                string open = "";
                                string url = "";
                                string due = "";
                                string innew = "";
                                string scale = "";

                                foreach(HtmlElement td in tds)
                                {
                                    if(td.GetAttribute("headers") == "title")
                                    {
                                        HtmlElement link = td.GetElementsByTagName("a")[0];
                                        title = link.InnerText;
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
                                        HtmlElement link = td.GetElementsByTagName("a")[0];
                                        string outerHtml = link.OuterHtml;
                                        MatchCollection mathes = Regex.Matches(outerHtml, "window.location\\s*=\\s*('|\")(.*?)('|\")", RegexOptions.IgnoreCase);
                                        url =  mathes[0].Groups[2].Value;

                                        innew = td.InnerText;
                                    }
                                    if (td.GetAttribute("headers") == "maxgrade")
                                    {
                                        scale = td.InnerText;
                                    }
                                }
                                if(title != "" && status != "" && open != "" && url != "" && due != "" && innew != "" && scale != "")
                                    dctAssignmentItems.Add(title, new Assignment(title, url, status, open, due, innew, scale));
                            }
                        }
                    }

                    // We have all assignments. Go to TryLoadStudents
                    indexOfProcessingAssignment = 0;
                    webBrowserTask = WebBrowserTask.LoadStudents;

                    if (dctAssignmentItems.Count >= 1) // Load into the main frame students
                    {
                        dctStudentInfos.Clear();
                        confidentLoad = true;
                        webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.Link);
                    }

                    break;
                case WebBrowserTask.ReloadStudents:

                    ParseLoadingStudents();


                    break;
                case WebBrowserTask.LoadStudents:

                    // Finding students
                    ParseLoadingStudents();

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
                        if(attachmentPresent)
                            webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.ElementAt(indexOfProcessingStudent).Value.GradeLink); // If you see it you are hero
                        else webBrowser_DocumentCompleted(webBrowser, e);
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
                        if(indexOfProcessingAssignment >= dctAssignmentItems.Count)
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
                case WebBrowserTask.Idle:
                default:
                    break;
            }
        }

        /// <summary>
        /// Method works with html and gets all required information about students
        /// Works with indexOfProcessingAssignment, it should be set
        /// </summary>
        private void ParseLoadingStudents()
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
                        bool attached = tr.InnerHtml.Contains("attachments.gif"); ;
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

                        if (studentID != "" && studentName != "")
                            dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.StudentInfosDictionary.Add(studentID,
                                new StudentInfo(studentName, studentID, submitted, status, grade, released, gradeLink, attached));
                    }

                }
            }
        }

        public Dictionary<String, StudentInfo> GetStudentsInformation(string assignmentTitle)
        {
            return dctAssignmentItems[assignmentTitle].StudentInfosDictionary;
        }

        public String[] GetStudentIDs(string assignmentTitle)
        {
            return dctAssignmentItems[assignmentTitle].StudentInfosDictionary.Keys.ToList().OrderBy(q => q).ToArray();
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

            if (dctAssignmentItems.Count >= 1) // Load into the main frame students
            {
                dctStudentInfos.Clear();
                confidentLoad = true;
                webBrowser.Document.Window.Frames[1].Navigate(dctAssignmentItems.ElementAt(indexOfProcessingAssignment).Value.Link);
            }
        }

        /// <summary>
        /// Get the names of assignments in the specified Site (Worksite).
        /// </summary>
        /// <returns>List of Assignments</returns>
        public Assignment[] GetAssignmentItems()
        {
            return dctAssignmentItems.Values.ToArray();
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
            webBrowserTask = WebBrowserTask.GetMembershipLink;
        }
    }
}

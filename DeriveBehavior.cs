using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml;

using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.Reflection;

using HPMSdk;
using Hansoft.ObjectWrapper;

using Hansoft.Jean.Behavior;
using Hansoft.ObjectWrapper.CustomColumnValues;

namespace Hansoft.Jean.Behavior.DeriveBehavior
{
    public class NoNewValueException : Exception
    { }

    public class DeriveBehavior : AbstractBehavior
    {
        const string assemblySourceTemplate = @"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using HPMSdk;
using Hansoft.ObjectWrapper;
using Hansoft.ObjectWrapper.CustomColumnValues;

namespace Hansoft.Jean.Behavior.DeriveBehavior.Expressions
{{
    public class {0}
    {{
        public static object EvaluateExpression(Task current_task)
        {{
            return {1};
        }}
    }}
}}
";

        class DerivedColumn
        {
            private bool isCustomColumn;
            private string expression;
            private MethodInfo mInfo;

            private string customColumnName;

            private EHPMProjectDefaultColumn defaultColumnType;

            internal DerivedColumn(bool isCustomColumn, string customColumnName, EHPMProjectDefaultColumn defaultColumnType, string expression)
            {
                this.isCustomColumn = isCustomColumn;
                this.customColumnName = customColumnName;
                this.defaultColumnType = defaultColumnType;
                this.expression = expression;
            }

            internal void Initialize(List<string> extensionAssemblies)
            {
                GenerateAssembly(extensionAssemblies);
            }

            internal string GetClassName()
            {
                return "ExpressionEvaluator" + GetHashCode().ToString();
            }

            internal void GenerateAssembly(List<string> extensionAssemblies)
            {
                CompilerParameters cParams = new CompilerParameters();
                CSharpCodeProvider cProvider = new CSharpCodeProvider();
                cParams.GenerateInMemory = true;
                cParams.ReferencedAssemblies.Add("System.Core.dll");
                cParams.ReferencedAssemblies.Add("ObjectWrapper.dll");
                // TODO: It is not that great to have the assembly name hard coded here since the 7.1 version of the SDK
                cParams.ReferencedAssemblies.Add("HPMSdkManaged_4_5.x86.dll");
                foreach (string extension in extensionAssemblies)
                    cParams.ReferencedAssemblies.Add(extension);

                string[] sources = new string[] { string.Format(assemblySourceTemplate, GetClassName(), expression) };
                CompilerResults results = cProvider.CompileAssemblyFromSource(cParams, sources);
                if (results.Errors.HasErrors)
                {
                    StringBuilder allErrors = new StringBuilder();
                    foreach (CompilerError err in results.Errors)
                        allErrors.Append(err.ToString());
                    throw new ArgumentException("Error in Expression parameter of Derive behavior: " + allErrors.ToString());
                }
                else
                {
                    TypeInfo tInfo = results.CompiledAssembly.DefinedTypes.First();
                    mInfo = tInfo.DeclaredMethods.First();
                }
            }

            internal void DoUpdate(Task task)
            {
                try
                {
                    Debug.Print("DeriveBehaviour - {0} - {1} - processing task \"{2}\"",
                        isCustomColumn ? customColumnName : defaultColumnType.ToString(),
                        expression,
                        task.Name);
                    var expressionValue = mInfo.Invoke(null, new object[] { task });
                    if (isCustomColumn)
                    {
                        // Ensure that we get the custom column of the right project
                        var customColumn = task.ProjectView.GetCustomColumn(customColumnName);
                        var oldValue = task.GetCustomColumnValue(customColumnName);

                        if (customColumn != null && oldValue != null && !oldValue.Equals(expressionValue))
                            task.SetCustomColumnValue(customColumn, expressionValue);
                    }
                    else
                    {
                        var oldValue = task.GetDefaultColumnValue(defaultColumnType);
                        if (oldValue != null && !oldValue.Equals(expressionValue))
                            task.SetDefaultColumnValue(defaultColumnType, expressionValue);
                    }
                }
                catch (TargetInvocationException e)
                {
                    // Hack to let column handlers bail out when they don't have any new value to set.
                    if (!(e.InnerException is NoNewValueException))
                        throw;
                }
            }
        }

        string projectName;
        EHPMReportViewType viewType;
        List<Project> projects;
        bool inverted = false;
        List<ProjectView> projectViews;
        string find;
        List<DerivedColumn> derivedColumns;
        bool changeImpact = false;
        string title;
        bool initializationOK = false;

        public DeriveBehavior(XmlElement configuration)
            : base(configuration)
        {
            projectName = GetParameter("HansoftProject");
            string invert = GetParameter("InvertedMatch");
            if (invert != null)
                inverted = invert.ToLower().Equals("yes");
            viewType = GetViewType(GetParameter("View"));
            find = GetParameter("Find");
            derivedColumns = GetDerivedColumns(configuration);
            title = "DeriveBehavior: " + configuration.InnerText;
        }

        public override string Title
        {
            get { return title; }
        }

        private void DoUpdate()
        {
            if (initializationOK)
            {
                foreach (ProjectView projectView in projectViews)
                {
                    List<Task> tasks = projectView.Find(find);
                    foreach (Task task in tasks)
                    {
                        foreach (DerivedColumn derivedColumn in derivedColumns)
                            derivedColumn.DoUpdate(task);
                    }
                }
            }
        }

        private void InitializeProjects()
        {
            projects = new List<Project>();
            projectViews = new List<ProjectView>();
            projects = HPMUtilities.FindProjects(projectName, inverted);
            foreach (Project project in projects)
            {
                ProjectView projectView;
                if (viewType == EHPMReportViewType.AgileBacklog)
                    projectView = project.ProductBacklog;
                else if (viewType == EHPMReportViewType.AllBugsInProject)
                    projectView = project.BugTracker;
                else
                    projectView = project.Schedule;

                projectViews.Add(projectView);
            }
        }

        public override void Initialize()
        {
            initializationOK = false;
            InitializeProjects();
            foreach (DerivedColumn derivedColumn in derivedColumns)
                derivedColumn.Initialize(ExtensionAssemblies);
            initializationOK = true;
            DoUpdate();
        }


        // TODO: Subject to refactoring
        private List<DerivedColumn> GetDerivedColumns(XmlElement parent)
        {
            List<DerivedColumn> columnDefaults = new List<DerivedColumn>();
            foreach (XmlNode node in parent.ChildNodes)
            {
                if (node is XmlElement)
                {
                    XmlElement el = (XmlElement)node;
                    switch (el.Name)
                    {
                        case ("CustomColumn"):
                            columnDefaults.Add(new DerivedColumn(true, el.GetAttribute("Name"), EHPMProjectDefaultColumn.NewVersionOfSDKRequired, el.GetAttribute("Expression")));
                            break;
                        case ("Risk"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.Risk, el.GetAttribute("Expression")));
                            break;
                        case ("Priority"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.BacklogPriority, el.GetAttribute("Expression")));
                            break;
                        case ("EstimatedDays"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.EstimatedIdealDays, el.GetAttribute("Expression")));
                            break;
                        case ("Category"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.BacklogCategory, el.GetAttribute("Expression")));
                            break;
                        case ("Points"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.ComplexityPoints, el.GetAttribute("Expression")));
                            break;
                        case ("Status"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.ItemStatus, el.GetAttribute("Expression")));
                            break;
                        case ("Confidence"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.Confidence, el.GetAttribute("Expression")));
                            break;
                        case ("Hyperlink"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.Hyperlink, el.GetAttribute("Expression")));
                            break;
                        case ("Name"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.ItemName, el.GetAttribute("Expression")));
                            break;
                        case ("WorkRemaining"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.WorkRemaining, el.GetAttribute("Expression")));
                            break;
                        case ("IsCompleted"):
                            columnDefaults.Add(new DerivedColumn(false, null, EHPMProjectDefaultColumn.IsCompleted, el.GetAttribute("Expression")));
                            break;
                        default:
                            throw new ArgumentException("Unknown column type specified in Derive behavior : " + el.Name);
                    }
                }
            }
            return columnDefaults;
        }

        public override void OnBeginProcessBufferedEvents(EventArgs e)
        {
            changeImpact = false;
        }

        public override void OnEndProcessBufferedEvents(EventArgs e)
        {
            if (BufferedEvents && changeImpact)
                DoUpdate();
        }

        public override void OnTaskChange(TaskChangeEventArgs e)
        {
            //            if (Task.GetTask(e.Data.m_TaskID).MainProjectID.m_ID == project.UniqueID.m_ID)
            //            {
            if (!BufferedEvents)
                DoUpdate();
            else
                changeImpact = true;
            //            }
        }

        public override void OnTaskChangeCustomColumnData(TaskChangeCustomColumnDataEventArgs e)
        {
            //            if (Task.GetTask(e.Data.m_TaskID).MainProjectID.m_ID == project.UniqueID.m_ID)
            //            {
            if (!BufferedEvents)
                DoUpdate();
            else
                changeImpact = true;
            //            }
        }

        public override void OnTaskCreate(TaskCreateEventArgs e)
        {
            //            if (e.Data.m_ProjectID.m_ID == projectView.UniqueID.m_ID)
            //            {
            if (!BufferedEvents)
                DoUpdate();
            else
                changeImpact = true;
            //            }
        }

        public override void OnTaskDelete(TaskDeleteEventArgs e)
        {
            if (!BufferedEvents)
                DoUpdate();
            else
                changeImpact = true;
        }

        public override void OnTaskMove(TaskMoveEventArgs e)
        {
            //            if (e.Data.m_ProjectID.m_ID == projectView.UniqueID.m_ID)
            //            {
            if (!BufferedEvents)
                DoUpdate();
            else
                changeImpact = true;
            //            }
        }
    }
}

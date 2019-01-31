using System.AddIn;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Text;
using System.Data;
using System.ComponentModel;
using System.Linq;
using System;
using System.ServiceModel;

////////////////////////////////////////////////////////////////////////////////
//
// File: WorkspaceAddIn.cs
//
// Comments:
//
// Notes: 
//
// Pre-Conditions: 
//
////////////////////////////////////////////////////////////////////////////////
namespace Allocation_Check
{
    public class WorkspaceAddIn : Panel, IWorkspaceComponent2
    {
        /// <summary>
        /// The current workspace record context.
        /// </summary>
        private IRecordContext _recordContext;
        public static IIncident _incidentRecord;
        private System.Windows.Forms.Label label1;
    
        public static IGlobalContext _globalContext { get; private set; }
        
        RightNowConnectService _rnConnectService;
        /// <summary>
        /// Default constructor.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        public WorkspaceAddIn(bool inDesignMode, IRecordContext RecordContext, IGlobalContext GlobalContext)
        {
            if (!inDesignMode)
            {

                _recordContext = RecordContext;
                _globalContext = GlobalContext;

                _rnConnectService = RightNowConnectService.GetService(_globalContext);

            }
        }

        #region IAddInControl Members

        /// <summary>
        /// Method called by the Add-In framework to retrieve the control.
        /// </summary>
        /// <returns>The control, typically 'this'.</returns>
        public Control GetControl()
        {
            return this;
        }

        #endregion

        #region IWorkspaceComponent2 Members

        /// <summary>
        /// Sets the ReadOnly property of this control.
        /// </summary>
        public bool ReadOnly { get; set; }

        /// <summary>
        /// Method which is called when any Workspace Rule Action is invoked.
        /// </summary>
        /// <param name="ActionName">The name of the Workspace Rule Action that was invoked.</param>
        public void RuleActionInvoked(string ActionName)
        {
            int nfPA = 0;
            int nfLA = 0;
            int nfOA = 0;

            int sPA = 0;
            int sLA = 0;
            int sOA = 0;

            int cPA = 0;
            int cLA = 0;
            int cOA = 0;

            int nfTotal = 0; 

            string val = "";
            switch (ActionName)
            {
                case "AllocationCheck":
                    _incidentRecord = (IIncident)_recordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);

                    val = getFieldFromIncidentRecord("CO", "nf_parts_allocation");
                    nfPA= Convert.ToInt32(NullCheck(val));
                    
                    val = getFieldFromIncidentRecord("CO", "cust_parts_allocation");
                    cPA = Convert.ToInt32(NullCheck(val));

                    val = getFieldFromIncidentRecord("CO", "supp_parts_allocation");
                    sPA = Convert.ToInt32(NullCheck(val));


                    nfTotal = nfPA + cPA + sPA;
                    if (nfTotal != 100)
                    {
                        MessageBox.Show("Parts Allocation does not equal 100%: " + nfTotal.ToString());
                        break; 
                    }
                    
                    val = getFieldFromIncidentRecord("CO", "nf_labor_allocation");
                    nfLA= Convert.ToInt32(NullCheck(val));

                    val = getFieldFromIncidentRecord("CO", "cust_labor_allocation");                    
                    cLA = Convert.ToInt32(NullCheck(val));

                    
                    val = getFieldFromIncidentRecord("CO", "supp_labor_allocation");
                    sLA = Convert.ToInt32(NullCheck(val));


                     nfTotal = nfLA + cLA + sLA;
                    if (nfTotal != 100)
                    {
                        MessageBox.Show("Labor Allocation does not equal 100%: " + nfTotal.ToString());
                        break; 
                    }


                    val = getFieldFromIncidentRecord("CO", "nf_other_allocation");
                    nfOA= Convert.ToInt32(NullCheck(val));


                    val = getFieldFromIncidentRecord("CO", "cust_other_allocation");
                    cOA = Convert.ToInt32(NullCheck(val));
                    
                    val = getFieldFromIncidentRecord("CO", "supp_other_allocation");
                    sOA = Convert.ToInt32(NullCheck(val));

                     nfTotal = nfOA + cOA + sOA;
                    if (nfTotal != 100)
                    {
                        MessageBox.Show("Other Allocation does not equal 100%: " + nfTotal.ToString());
                        break; 
                    }

                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Method which is called when any Workspace Rule Condition is invoked.
        /// </summary>
        /// <param name="ConditionName">The name of the Workspace Rule Condition that was invoked.</param>
        /// <returns>The result of the condition.</returns>
        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }


        public string NullCheck(string inVal)
        {
            string v="";
            if (inVal == null)
            {
                v = "0";
            }
            else
            {
                if (inVal == "")
                {
                    v = "0";
                }
                else
                {
                    v = inVal;
                }
            }

            return v;
        }

        /// <summary>
        /// Method which is called to get value of a custom field of Incident record.
        /// </summary>
        /// <param name="packageName">The name of the package.</param>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string getFieldFromIncidentRecord(string packageName, string fieldName)
        {
            string value = "";
            IList<ICustomAttribute> incCustomAttributes = _incidentRecord.CustomAttributes;

            foreach (ICustomAttribute val in incCustomAttributes)
            {
                if (val.PackageName == packageName)//if package name matches
                {
                    if (val.GenericField.Name == packageName + "$" + fieldName)//if field matches
                    {
                        if (val.GenericField.DataValue.Value != null)
                        {
                            value = val.GenericField.DataValue.Value.ToString();
                            break;
                        }
                    }
                }
            }
            return value;
        }
        /// <summary>
        /// Method which is use to set incident field 
        /// </summary>
        /// <param name="pkgName">package name of custom field</param>
        /// <param name="fieldName">field name</param>
        /// <param name="value">value of field</param>
        public static void SetIncidentField(string pkgName, string fieldName, string value)
        {
            if (pkgName == "c")
            {
                IList<ICfVal> incCustomFields = _incidentRecord.CustomField;
                int fieldID = GetCustomFieldID(fieldName);
                foreach (ICfVal val in incCustomFields)
                {
                    if (val.CfId == fieldID)
                    {
                        switch (val.DataType)
                        {
                            case (int)RightNow.AddIns.Common.DataTypeEnum.BOOLEAN_LIST:
                            case (int)RightNow.AddIns.Common.DataTypeEnum.BOOLEAN:
                                if (value == "1" || value.ToLower() == "true")
                                {
                                    val.ValInt = 1;
                                }
                                else if (value == "0" || value.ToLower() == "false")
                                {
                                    val.ValInt = 0;
                                }
                                break;
                        }

                    }
                }
            }
            else
            {
                IList<ICustomAttribute> incCustomAttributes = _incidentRecord.CustomAttributes;

                foreach (ICustomAttribute val in incCustomAttributes)
                {
                    if (val.PackageName == pkgName)
                    {
                        if (val.GenericField.Name == pkgName + "$" + fieldName)
                        {
                            switch (val.GenericField.DataType)
                            {
                                case RightNow.AddIns.Common.DataTypeEnum.BOOLEAN:
                                    if (value == "1" || value.ToLower() == "true")
                                    {
                                        val.GenericField.DataValue.Value = true;
                                    }
                                    else if (value == "0" || value.ToLower() == "false")
                                    {
                                        val.GenericField.DataValue.Value = false;
                                    }
                                    break;
                                case RightNow.AddIns.Common.DataTypeEnum.INTEGER:
                                    if (value.Trim() == "" || value.Trim() == null)
                                    {
                                        val.GenericField.DataValue.Value = null;
                                    }
                                    else
                                    {
                                        val.GenericField.DataValue.Value = Convert.ToInt32(value);
                                    }
                                    break;
                                case RightNow.AddIns.Common.DataTypeEnum.STRING:
                                    val.GenericField.DataValue.Value = value;
                                    break;
                            }
                        }
                    }
                }
            }
            return;
        }
        /// <summary>
        /// Method to get custom field id by name
        /// </summary>
        /// <param name="fieldName">Custom Field Name</param>
        public static int GetCustomFieldID(string fieldName)
        {
            IList<IOptlistItem> CustomFieldOptList = _globalContext.GetOptlist((int)RightNow.AddIns.Common.OptListID.CustomFields);//92 returns an OptList of custom fields in a hierarchy
            foreach (IOptlistItem CustomField in CustomFieldOptList)
            {
                if (CustomField.Label == fieldName)//Custom Field Name
                {
                    return (int)CustomField.ID;//Get Custom Field ID
                }
            }
            return -1;
        }


        #endregion

        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "Allocation Check";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Allocation Check";
            this.ResumeLayout(false);

        }
    }

    [AddIn("Workspace Factory AddIn", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        #region IWorkspaceComponentFactory2 Members
        static public IGlobalContext _globalContext;
        /// <summary>
        /// Method which is invoked by the AddIn framework when the control is created.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <returns>The control which implements the IWorkspaceComponent2 interface.</returns>
        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceAddIn(inDesignMode, RecordContext, _globalContext);
        }

        #endregion

        #region IFactoryBase Members

        /// <summary>
        /// The 16x16 pixel icon to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public Image Image16
        {
            get { return Properties.Resources.AddIn16; }
        }

        /// <summary>
        /// The text to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Text
        {
            get { return "Allocation Check"; }
        }

        /// <summary>
        /// The tooltip displayed when hovering over the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Tooltip
        {
            get { return "Allocation Check"; }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            return true;
        }

        #endregion
    }
}
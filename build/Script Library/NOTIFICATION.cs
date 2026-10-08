// LIBRARY NOTIFICATION
// UPDATE for 2021.11 - 12.15.13.00 - Add ex.ToString() details in Catch for SendEMail()
// UPDATE for 2022.05 - updated logic for creating object URL in notification
// UPDATE for 2024.06 - updated href variable in makeObjectUrl() function for MT SaaS v2 Portal
//Script Library: NOTIFICATION  
//AUTHOR    :   
//CREATED   :   
//REVISIONS :   WHO    DATE          CHANGE      REASON
//------------------------------------------------------------------------------------
//REV001     Nikhilesh 12052026     Updated the url to correctly display the object in Optiva upon opening.
using System;
using Formation.Shared.Defs;

public class NOTIFICATION : FcProcFuncSetEventWF
{
    private FcProcFuncSetEvent co;

    public NOTIFICATION(ref FcProcFuncSetEvent context)
    {
        co = context;
    }

    // *****************************************
    // Out of box workflow actions' Notifications 

    public long SendEMail(string sObjectCode, string sObjectSymbol, string sName, string sTemplate, int iIndicator, string sActionSetCode, string sActionCode, string sReasonCode, string sComment, string sApprovedUser, int iWipId, int iWipLineId)
    {

        try
        {
            string[] key = sObjectCode.Split('\\');
            string objectUrl;
            if (key is not null && key.Length == 2)
            {
                objectUrl = makeObjectUrl(sObjectSymbol, key[0], key[1]);
            }
            else
            {
                objectUrl = makeObjectUrl(sObjectSymbol, sObjectCode);
            }

            var gf = new GENERALFUNCTIONS(ref co);
            string sClass = co.ObjProperty("CLASS", sObjectSymbol, sObjectCode).ToString();
            string sClassName = "";
            if (co.IsBlank(sClass) == 0)
                sClassName = gf.GetEnumLabel(sClass, "CLASS", sObjectSymbol);
            if (sClass != sClassName)
                sClassName = sClassName + " (" + sClass + ")";
            string sDescription = co.ObjProperty("DESCRIPTION", sObjectSymbol, sObjectCode).ToString();
            string sWorkflowDesc = co.ObjProperty("DESCRIPTION", "ACTIONSET", sActionSetCode).ToString();
            var oStatus = co.ObjProperty("STATUSIND.STATUS", sObjectSymbol, sObjectCode);

            string sActionName = co.ObjProperty("DESCRIPTION", "ACTION", sActionCode).ToString();
            string sTaskName = co.WIPInfoGet("DESCRIPTION");    // [LineID, ActionWIPId])
                                                                // co.Messagelist("Task is ", sTaskName);
            if (sTaskName is not null && co.IsBlank(sTaskName) == 0)
                sActionName = sTaskName;

            var sReasonName = "";
            if (co.IsBlank(sReasonCode) == 0)
            {
                sReasonName = gf.GetEnumLabel(sReasonCode, "REASONCODE", "REASONCOMMENT");
                // co.Messagelist("Reason name is ", sReasonname);
            }

            string sStatusDesc = gf.GetStatusDesc(sObjectSymbol, Convert.ToInt32(oStatus));
            sStatusDesc = sStatusDesc + " (" + oStatus.ToString() + ")";

            string sStartCode = co._STARTUSER;
            string sStartName = "";
            if (co.IsBlank(sStartCode) == 0)
                sStartName = co.ObjProperty("DESCRIPTION", "USER", sStartCode).ToString();

            string sDateTime = DateTime.Now.ToString("MMM d yyyy HH:mm");


            // ---Email Notification
            co.Notify(sName, sTemplate, 1, iIndicator, sClassName, sObjectCode, sDescription, sApprovedUser, iWipId.ToString(), sActionSetCode, sWorkflowDesc, sStatusDesc, sDateTime, sActionName, objectUrl, sName, sReasonName, sComment, System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sObjectSymbol), sStartName + " (" + sStartCode + ")");
        }
        // Tokens:
        // 1 CLASS                        2 Object keycode
        // 3 Object Description           4 Task USER
        // 5 WIP ID                       6 Action Set Code 
        // 7 Action Set Description       8 Status 
        // 9 Current Date/Time            10 Action Code 
        // 11 Object url                  12 Addressee
        // 13 Reason Code                 14 Comment (global variable or any comment text sent)
        // 14 Object symbol               16 Workflow launcher

        catch (Exception Ex)
        {
            co.MessageList("Email Sending  ", sName, " Failed: ", Ex.Message, Environment.NewLine, Ex.ToString());
            co.PublishDebuggingInfo("Notification error: " + Ex.ToString());
            return 0L;
        }

        return 1L;

    }


    // *****************************************
    // Basic Example NOTIFICATION

    // To call this library function from an Action, assuming email template ACTION REQUIRED exists:
    // Dim notify As NOTIFICATION = New NOTIFICATION(Me)
    // notify.SendActionRequiredEmail("Approval is requested.", _STARTUSER, _OBJECTSYMBOL, _OBJECTKEY)

    public void SendActionRequiredEmail(string actionRequiredText, string userCode, string symbol, string objectId)
    {
        co.PublishDebuggingInfo("Entering SendActionRequiredEmail");
        string symbolDesc = co.ObjProperty("DESCRIPTION", "SYMBOL", symbol).ToString();
        string[] key = objectId.Split('\\');
        string desc = co.ObjProperty("DESCRIPTION", symbol, objectId).ToString();
        string objectUrl;
        if (key is not null && key.Length == 2)
        {
            objectUrl = makeObjectUrl(symbol, key[0], key[1]);
        }
        else
        {
            objectUrl = makeObjectUrl(symbol, objectId);
        }

        co.Notify(userCode, "ACTION REQUIRED", 1, 0, DateTime.Today.ToString(), symbolDesc, objectUrl, desc, actionRequiredText);
        co.PublishDebuggingInfo("Leaving SendActionRequiredEmail");
    }

    // *****************************************
    // Function to build the URL for the workflow target object
    private string makeObjectUrl(string symbol, string objectId, string ver = null)
    {

        string url = null;
        string displayObjectId = objectId;
        string URLParams = "type%3D" + symbol + "%26id%3D" + objectId;
        if (ver is not null)
        {
            URLParams += "%26version%3D" + ver;
            displayObjectId += @"\" + ver;
        }

        string[] ProfileLogicalId = co.GetProfileValue("ION.FROMLOGICALID");
        string logicalID = ProfileLogicalId[0];
        //REV001Starts
        if (co.IsBlank(logicalID) == 0 && (logicalID.Contains("lid://") != true))
            logicalID = "lid://" + logicalID;
        //REV001Ends
        string[] ProfileMultiTenant = co.GetProfileValue("MULTITENANT.ENABLED");

        if (ProfileMultiTenant is not null && ProfileMultiTenant[0] == "1")
        {

            string TenantURLPrefix = co._PORTALSUBDOMAIN;
            string[] ProfileTenant = co.GetProfileValue("ION.TENANT");
            string Tenant = ProfileTenant[0];
            // Old URL style for Infor OS portal v1:  string href = TenantURLPrefix + "/" + Tenant + "?type=" + symbol + "&id=" + objectId + "&version=" + ver + "&LogicalId=" + logicalID;
            // Updated URL format for Infor OS portal v2:
            string href = TenantURLPrefix + "/v2/" + Tenant + "?favoriteContext=" + URLParams + "&LogicalId=" + logicalID;

            url = "<a href='" + href + "'>" + displayObjectId + "</a>";
        }
        else
        {
            // This is the URL if Optiva runs standalone, but change "localhost" to your actual server name.  If running inside Ming.le, remove the line below and uncomment out the code below that. 
            url = "<a href='http://localhost/FsOptivaWeb/Default.aspx?type=" + symbol + "&id=" + objectId + "&version=" + ver + "'>" + displayObjectId + "</a>";

            // Start of code to use if running Optiva inside Ming.le - also remove code above
            // string MingleURLPrefix = "https://usfrvmingle19.infor.com:8443/infor/";    // This is an example
            // string href = MingleURLPrefix + "?favoriteContext=" + URLParams + "&LogicalId=" + logicalID;
            // url = "<a href='" + href + "'>" + displayObjectId + "</a>";
            // End of code to use if running Optiva inside Ming.le

        }
        return url;
    }


}
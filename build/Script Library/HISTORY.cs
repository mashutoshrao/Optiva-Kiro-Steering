//SL Code: HISTORY
//Date: 27 Mar 2026
//Author: Nikhilesh
//--------------------------------------------------------------------
//Revision Number    |   Date   |  User |   Reason
//--------------------------------------------------------------------
//REV001               27032026 Nikhilesh  Put in place an improvised wrf_history function.
Using System;
Using Formation.Shared.Defs;

Public Class HISTORY
{
    Private FcProcFuncSetEvent co;

    Public HISTORY(ref FcProcFuncSetEvent context)
    {
        co = context;
    }

    //REV001
    Public String wrf_history(String sObjectCode, String sObjectSymbol, String sSourceUser, String sFunctionCode, String sEvent, String sMsg = "")
    {
        // ---Log History
        DateTime dToday = DateTime.Now;
        String first, last, newtext, oldtext, Delim;
        first = co.ObjProperty("FIRSTNAME", "USER", sSourceUser) as string;
        last = co.ObjProperty("LASTNAME", "USER", sSourceUser) as string;
        String sComment = ".";
        If (co.IsBlank(co._COMMENT) == 0) sComment = " with comments: " + co._COMMENT + sComment;

        int iText = co.DocExist(sObjectSymbol, sObjectCode, sFunctionCode);
        If (iText == 1 || iText == 3)
        {
            Delim = "--------------";
            oldtext = "\r\n" + Delim + "\r\n" + co.ObjProperty("DOCTEXT.DOC.A", sObjectSymbol, sObjectCode, sFunctionCode, 1);
        }
        Else
        {
            oldtext = "";
        }

        // if (co.IsBlank(sMsg) == 1)
        // {
            String sAction = co.WIPInfoGet("DESCRIPTION") as string;
            If (co.IsBlank(sAction) == 1)
            {
                sAction = co.ObjProperty("DESCRIPTION", "ACTION", co._ACTIONCODE) as string;
            }
            String sActionSet = co.ObjProperty("DESCRIPTION", "ACTIONSET", co._ACTIONSETCODE) as string;

            sEvent = sEvent.ToLower();
            If (sEvent == "start")
                sMsg = " has started the task '" + sAction + "'" + " of Workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            ElseIf (sEvent == "complete")
                sMsg = " has completed the task '" + sAction + "'" + " of Workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            ElseIf (sEvent == "approve")
                sMsg = " has approved the task '" + sAction + "'" + " of Workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            ElseIf (sEvent == "reject")
                sMsg = " has rejected the task '" + sAction + "'" + " of Workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            ElseIf (sEvent == "reassign")
                sMsg = " has reassigned the task '" + sAction + "'" + " of Workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            ElseIf (sEvent == "return")
                sMsg = " has returned the task '" + sAction + "'" + " of Workflow '" + sActionSet + "'" + "{b}{font color=\"#b94e4e\"}" + sComment + "{/font}{/b}" + sMsg;
            ElseIf (sEvent == "onhold")
                sMsg = " has put the task on hold '" + sAction + "'" + " of Workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            ElseIf (sEvent == "wfcomplete")
                sMsg = " has completed the workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            ElseIf (sEvent == "wfstart")
                sMsg = " has started the workflow '" + sActionSet + "'" + "{b}" + sComment + "{/b}" + sMsg;
            Else
                sMsg = sMsg;
        // }

        newtext = dToday + " (UTC): " + first + " " + last + " " + sMsg + "\r\n" + oldtext;
        co.ObjPropertySet(newtext, 0, "DOCTEXT.DOC.A", sObjectSymbol, sObjectCode, sFunctionCode, 1);
        Return newtext;
    }
    //REV001

    Public String wf_history(String sObjectCode, String sObjectSymbol, String sSourceUser, String sFunctionCode, String sMsg)
    {

        // ---Log History
        var dToday = DateTime.Now;
        String first, last, newtext, oldtext, Delim;
        first = co.ObjProperty("FIRSTNAME", "USER", sSourceUser).ToString();
        last = co.ObjProperty("LASTNAME", "USER", sSourceUser).ToString();

        int iText = co.DocExist(sObjectSymbol, sObjectCode, sFunctionCode);
        If (iText == 1 || iText == 3)
        {
            Delim = "--------------";
            oldtext = Environment.NewLine + Delim + Environment.NewLine + co.ObjProperty("DOCTEXT.DOC.A", sObjectSymbol, sObjectCode, sFunctionCode, 1L).ToString();
        }
        Else
        {
            oldtext = "";
        }

        newtext = dToday.ToString() + ": " + first + " " + last + " " + sMsg + Environment.NewLine + oldtext;
        co.ObjPropertySet(newtext, 0L, "DOCTEXT.DOC.A", sObjectSymbol, sObjectCode, sFunctionCode, 1L);

        Return newtext;

    }



    Public Long cr_history(String sObjectCode, String sObjectSymbol, String sSourceUser, String sDocCode, String sCreateRule)
    {

        // ---Log History
        var dToday = DateTime.Now;
        String sFirstName, sLastName, sNewtext;
        sFirstName = co.ObjProperty("FIRSTNAME", "USER", sSourceUser).ToString();
        sLastName = co.ObjProperty("LASTNAME", "USER", sSourceUser).ToString();

        int iText = co.DocExist(sObjectSymbol, sObjectCode, sDocCode);
        If (iText!= -1)
        {
            sNewtext = dToday.ToString() + ": " + sFirstName + " " + sLastName + " Created the " + sObjectSymbol + " using Create Rule " + sCreateRule + Environment.NewLine;
            co.ObjPropertySet(sNewtext, 0L, "DOCTEXT.DOC.A", sObjectSymbol, sObjectCode, sDocCode, 1L);
        }
        Else
        {
            co.MessageList("Invalid DocCode ", sDocCode, " History has not been Logged");
        }

        Return 1L;

    }



    Public Long wf_call()
    {
        co.MessageList("Hello World!");
        Return 1L;
    }

}
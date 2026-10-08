// Guide to Library Functions:
// -------------------------------------
// GetStatusDesc:  Get the status label by Set type And status value
// GetObjectID:  Return xxxx_ID For two part key objects, Or keycode For 1 part key objects
// GetEnumLabel:  For a header Or custom field return the UI label if enum supported, Or existing value if Not
// GetEnumLabelTP:  For a parameter return the UI label if enum supported, Or existing value if Not
// GetEnumLabelDM:  For a custom table column return the UI label if enum supported, Or existing value if Not
// GetFieldLabel:  Retrieve the label For a header (standard Or custom) field
// GetSetDesc:  Retrieve the description For a Set code Of a specific symbol
// ClearACL:  Clear all entries from ACL tables For Item, Project, Formula, Spec
// 25-APR-2025 update:  Changed ClearACL To use Dataset scrpting instead Of Query To remove data
// MultiLangText:  Pass back a String With language-specific text In all language-pack languages, Or selected languages

// DATE		NAME		REASON/COMMENT
//============================================================
// 18DEC18	Infor			Original



Using System;
Using Formation.Shared.Defs;
Using System.Collections.Generic;
Using System.Diagnostics;
Using System.Globalization;
Using System.IO;
Using System.Linq;
Using System.Reflection;
Using System.Runtime.CompilerServices;
Using System.Security;
Using System.Text;
Using System.Threading.Tasks;
Using Microsoft.VisualBasic;
Using System.Data;
Using System.Xml;


Public Class GENERALFUNCTIONS
{
    Private FcProcFuncSetEvent co;
    String defaultlang = "EN-US";

    Public GENERALFUNCTIONS(ref FcProcFuncSetEvent context)
    {
        co = context;
    }

    // --------------------------------------------------------------------------------------------------------------------
    Public String GetStatusDesc(String symbol, int statusind, String langcode = "EN-US")
    {
        Try
        {
            Object oStatdesc;
            String sStatdesc = "";
            If (co.IsBlank(langcode) == 1)
            {
                langcode = defaultlang;
            }
            oStatdesc = co.TableLookup("GETSTATUSDESC", statusind, symbol, langcode);
            If (co.IsBlank(langcode) == 1)
            {
                langcode = defaultlang;
            }
            If (oStatdesc!= null)
                sStatdesc = oStatdesc.ToString();

            Return sStatdesc;
        }
        Catch (Exception ex)
        {
            co.MessageList("Error retrieving status description for code " + statusind + " for " + symbol + ": " + ex.Message);
            Return String.Empty;
        }
    }

    // --------------------------------------------------------------------------------------------------------------------
    Public String GetObjectID(String symbol, String objectkey)
    {
        Try
        {
            String sID = objectkey;

            If (VBStrings.InStr(sID, "\\") > 0) // vb code:  If (InStr(sID, "\") > 0)
            {
                DataSet ds = co.ObjectDataSet(symbol, objectkey);
                DataTable dt = ds.Tables["FS" + symbol];
                Object oID = dt.Rows[0][symbol + "_ID"];
                If (oID!= null)
                    sID = oID.ToString();
            }

            Return sID;
        }
        Catch (Exception ex)
        {
            Message("Error retrieving " + symbol + "_ID for " + objectkey + ": " + ex.Message);
            Return String.Empty;
        }
    }

    // --------------------------------------------------------------------------------------------------------------------
    Public String GetEnumLabel(String fldValue, String fieldname, String symbol, String langcode = "EN-US")
    {
        // Call GetFieldEnumLabel query

        String result;
        If (co.IsBlank(langcode) == 1)
        {
            langcode = defaultlang;
        }
        Object oQuery = co.TableLookup("GETFIELDENUMLABEL", fldValue, fieldname, symbol, langcode);
        If (oQuery == null)
            result = fldValue;
        Else
            result = oQuery.ToString();
        Return result;
    }

    // --------------------------------------------------------------------------------------------------------------------
    Public String GetEnumLabelTP(String paramcode, String fldValue, String langcode = "EN-US")
    {
        // Call 'GetTPEnumLabel query

        String result;
        If (co.IsBlank(langcode) == 1)
            {
                langcode = defaultlang;
            }
        Object oQuery = co.TableLookup("GETTPENUMLABEL", fldValue, paramcode, langcode);
        If (oQuery == null)
            result = fldValue;
        Else
            result = oQuery.ToString();
        Return result;
    }

    // --------------------------------------------------------------------------------------------------------------------
    Public String GetEnumLabelDM(String fldValue, String tablename, String fieldname, String langcode = "EN-US")
    {
        // Call GetDMFieldLabel query
        //tablename Is FsValidationfield.Validation_subcode, e.g. PROJECT0
        //fieldname Is FsValidationfield.Field_Name, e.g. 'FIELD8'

        String result;
        If (co.IsBlank(langcode) == 1)
            {
                langcode = defaultlang;
            }
        Object oQuery = co.TableLookup("GETDMFIELDLABEL", fldValue, fieldname, tablename, langcode);
        If (oQuery == null)
            result = fldValue;
        Else
            result = oQuery.ToString();
        Return result;
    }

    // --------------------------------------------------------------------------------------------------------------------
    Public String GetFieldLabel(String fieldname, String symbol, String langcode = "EN-US")
    {
        String result;
        If (co.IsBlank(langcode) == 1)
            {
                langcode = defaultlang;
            }
        Object oQuery = co.TableLookup("GETFIELDLABEL", symbol, langcode, fieldname);
        If (oQuery == null)
            result = fieldname;
        Else
            result = oQuery.ToString();
        Return result;
    }
    // --------------------------------------------------------------------------------------------------------------------
    Public String GetSetDesc(String symbol, String setcode, String langcode = "EN-US")
    {
        String result;
        If (co.IsBlank(langcode) == 1)
            {
                langcode = defaultlang;
            }
        Object getdesc = co.TableLookup("GETSETDESCBYCODE", symbol, setcode, langcode);
        If (getdesc == null || co.IsBlank(getdesc) == 1)
            // co.MessageList("No description");
            result = setcode;
        Else
            result = getdesc.ToString();

        Return result;
    }

    //================================================================
    // Created sub function for Messagelist to later enable multi-language Messages
    Public void Message(String msgtext)
    {
        co.MessageList(msgtext);
    }

    //================================================================
    // Created sub function to clear all ACL entries for an object.  E.g. use at the end of a workflow
    Public void ClearACL(String symbol, String key)
    {
        Try
        {
            If (!(symbol == "PROJECT" || symbol == "ITEM" || symbol == "FORMULA" || symbol == "SPECIFICATION"))
                co.MessageList("The ACL Clear function currently only supports Projects, Items, Formulas, and Specifications.");
            Else
            {
                String acltable = co.DataSetTableName(symbol, key, "ACL");
                // co.MessageList("ACL table is ", acltable);
                DataSet ds = co.ObjectDataSet(symbol, key, "HEADER;ACL");
                DataTable dt = ds.Tables[acltable];
                //co.MessageList("check datatable FS",  symbol,  "ACL:  ", dt.GetType().Name);

                If (dt == null || dt.Rows.Count == 0)
                {
                    //co.MessageList("Nothing to remove");
                    Return;
                }
                Else
                {
                    // co.MessageList("Rows before removal: ", dt.Rows.count);
                    DataRow[] currentRows = dt.Select();
                    foreach (DataRow dr in currentRows)
                        dr.Delete();
                }
            }
        }
        Catch (Exception ex)
        {
            co.MessageList("Error clearing ACL entries for ", symbol, " ", key, ": ", ex.Message);
        }
    }

    //================================================================
    // Pass symbol & keycode, And a field name with translated text, e.g. DESCRIPTION.
    // Pass a text string with tokens to be replaced by language code (%1) And translated text (%2).
    // Language code can be switched to lower case if needed.
    // If specifying languages instead of taking all installed, must be passed as "LANG1;LANG2;LANG3" etc.

    Public String MultiLangText(String objsymbol, String objkey, String textstring, String sField, String textcase = "", String langlist = "", int showall = 0)
    {
        Try
        {
            String totalstring = "";

            // Get database default language, Or English if none.
            String[] sDefaultLangArr = co.GetProfileValue("DEFAULTLANGUAGE");
            String sDefaultLang = "EN-US";
            If (textcase == "lower")
                sDefaultLang = sDefaultLang.ToLower();
            If (sDefaultLangArr!= null && sDefaultLangArr.Length > 0)
                sDefaultLang = sDefaultLangArr[0];

            // Use language list as passed, else get all installed languages
            If (co.IsBlank(langlist) == 1)
            {
                // Get installed languages:
                DataTable dtLangs = co.TableLookupEx("ENUMLANGUAGE", "langtable");
                If (dtLangs!= null && dtLangs.Rows.Count > 0)
                {
                    foreach (DataRow drLangs in dtLangs.Rows)
                    {
                        String sLangString = textstring;
                        String sLang = drLangs[0].ToString();
                        If (textcase == "lower")
                            sLang = sLang.ToLower();
                        String sDesc = GetLangText(objsymbol, objkey, sLang, sField, sDefaultLang, showall);
                        If (co.IsBlank(sDesc) == 1)
    Continue Do;
                        If (VBStrings.InStr(sLangString, "[%2]") > 0) // vb code:  If (InStr(sLangString, "[%2]") > 0)
                            sLangString = sLangString.Replace("[%2]", sDesc);
                        If (VBStrings.InStr(sLangString, "[%1]") > 0) // vb code:  If (InStr(sLangString, "[%1]") > 0)
                            sLangString = sLangString.Replace("[%1]", sLang);

                        If (totalstring == "")
                            totalstring = sLangString;
                        Else
                            totalstring = totalstring + Environment.NewLine + sLangString;
                    }
                }
            }
            ElseIf (VBStrings.InStr(langlist, ";") == 0) // vb code:  (instr(langlist, ";") = 0)     
                co.MessageList("Error: Language list must be semi-colon delimited.");
            Else
            {
                String[] arLanguages = langlist.Split(';');
                foreach (string LangInList in arLanguages)
                {
                    String LangInListUse = LangInList;
                    String sLangString2 = textstring;
                    If (textcase == "lower")
                        LangInListUse = LangInList.ToLower();
                    String sDesc2 = GetLangText(objsymbol, objkey, LangInListUse, sField, sDefaultLang, showall);
                    If (co.IsBlank(sDesc2) == 1)
    Continue Do;
                        If (VBStrings.InStr(sLangString2, "[%2]") > 0) // vb code:  If (InStr(sLangString2, "[%2]") > 0)
                        sLangString2 = sLangString2.Replace("[%2]", sDesc2);
                        If (VBStrings.InStr(sLangString2, "[%1]") > 0) // vb code:  If (InStr(sLangString2, "[%1]") > 0)
                        sLangString2 = sLangString2.Replace("[%1]", LangInList);
                    If (totalstring == "")
                        totalstring = sLangString2;
                    Else
                        totalstring = totalstring + Environment.NewLine + sLangString2;
                }
            }

            Return totalstring;
        }
        Catch (Exception ex)
        {
            co.MessageList("Error building multi-language text: ", ex.Message);
            Return String.Empty;
        }
    }

    //================================================================

    Public String GetLangText(String objsymbol, String objkey, String sLang, String sField, String sDefaultLang, int showall)
    {

        // Try to get description, else pass default
        DataSet ds = co.ObjectDataSet("", "");
        DataTable dt = ds.Tables["FSDESCRIPTION"];
        DataRow[] dr1 = dt.Select(string.Format("LANGUAGE_CODE = '{0}'", sLang));
        String sDesc = "";
        If (dr1!= null && dr1.Length > 0)
            sDesc = dr1[0][sField].ToString();
        ElseIf (showall == 1)
        {
            DataRow[] dr2 = dt.Select(string.Format("LANGUAGE_CODE = '{0}'", sDefaultLang));
            sDesc = dr2[0][sField].ToString();
        }
        Return sDesc;
    }

}
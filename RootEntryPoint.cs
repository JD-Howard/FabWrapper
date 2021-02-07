//      Fabrication ITEM .Net Wrapper 1.0
//      Written by Josh Howard 2017
//      www.houseofbim.com

using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using hand = AgileBIM.FabWrapper.TypeHandlers;
using System.Linq;

[assembly: ExtensionApplication(typeof(AgileBIM.FabWrapper.INI))]
[assembly: CommandClass(typeof(AgileBIM.FabWrapper.APIwrappers))]

namespace AgileBIM.FabWrapper
{
    //Initialization
    public class INI : IExtensionApplication
    {
        public Assembly assymblyRef;
        public void Initialize()
        {
            var acadpath = System.Diagnostics.Process.GetProcessesByName("acad").FirstOrDefault().MainModule.FileName;
            var year = new String(acadpath.ToCharArray().Where(p => char.IsDigit(p) == true).ToArray());
            try
            { assymblyRef = Assembly.LoadFrom("C:\\Program Files\\Autodesk\\Fabrication " + year.ToString() + "\\CADmep\\FabricationAPI.dll"); }
            catch (System.Exception ex)
            { Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message); }
        }
        public void Terminate() { }
    }


    //Lisp Functions
    public class APIwrappers
    {
        internal static Dictionary<string, dynamic> VarCache = new Dictionary<string, dynamic>();

        // (FabItemVarCache VarName [Value])
        // Purpose: To clear the internal global variable cache and hopefully release all stored objects
        //          hopefully preventing memory leaks and potential crashes by enabling garbage collection.
        [LispFunction("FabClearCache")]
        public bool FabClearCache(ResultBuffer args)
        {
            foreach (object o in VarCache.Values)
            {
                MethodInfo mii = null;
                foreach (MethodInfo mi in o.GetType().GetMethods())
                {
                    if (mi.Name.ToUpper() == "DISPOSE")
                    {
                        mii = mi;
                        break;
                    }                    
                }
                if (mii != null)
                {
                    try
                    {
                        mii.Invoke(o, null);
                    }
                    catch (System.Exception) { }
                }                
            }
            VarCache.Clear();
            return true;
        }


        // (FabItemVarCache VarName [Value])
        // Puprose: If a Value is provided it will set or create a variable in the GCV with the specified value.
        //          If no value is provided then it will attempt to locate and translate the value stored in the
        //          GCV to the user. IF it doesn't exist or the value is actually NULL, then it returns nil.
        [LispFunction("FabVarCache")]
        public dynamic FabVarCache(ResultBuffer args)
        {            
            ResultBuffer res = new ResultBuffer();
            if (args != null)
            {
                TypedValue[] argarray = args.AsArray();
                string name = "";
                dynamic value;
                if (argarray.Length >= 1 && argarray[0].Value != null)
                    name = hand.VariableName((string)argarray[0].Value);
                if (argarray.Length == 2)
                {                    
                    value = argarray[1];
                    if (VarCache.ContainsKey(name) == true)
                    {
                        VarCache[name] = hand.ConvertFromRB(value);
                    }
                    else
                    {
                        VarCache.Add(name, hand.ConvertFromRB(value));
                    }
                    res.Add(new TypedValue((int)LispDataType.T_atom));
                }
                else if (argarray.Length > 2)
                {
                    value = new ResultBuffer();
                    for (int i = 1; i < argarray.Length - 1; i++)
                    { value.Add(argarray[i].Value); }
                    if (VarCache.ContainsKey(name) == true)
                    {
                        VarCache[name] = value;
                    }
                    else
                    {
                        VarCache.Add(name, value);
                    }
                    res.Add(new TypedValue((int)LispDataType.T_atom));
                }
                else if (VarCache.ContainsKey(name) == true)
                {                    
                    res = new ResultBuffer();
                    res.Add(new TypedValue((int)LispDataType.ListBegin));
                    foreach (TypedValue tv in hand.ConvertToRB(VarCache[name]))
                    {
                        res.Add(tv);
                    }
                    res.Add(new TypedValue((int)LispDataType.ListEnd));
                }
            }
            if (res != null && res.AsArray().Length == 1)
            {
                return res.AsArray()[0];
            }
            else
            {
                return res;
            }            
        }

        
        // (FabGetItemProp ENAME PathToProp [VarName])
        // Purpose: This will use reflection on the provided object to navigate its properties as specified in the
        //          users PathToProp string and return that value to the user. If the VarName is specified then
        //          the program will return the value (if translatable), but also place a reference in the GCV.
        //          Further note that you can use a $VarName in place of the Ename and "THIS" as a path reference.
        [LispFunction("FabGetItemProp")]
        public dynamic FabGetItemProp(ResultBuffer args)
        {
            ResultBuffer ret = new ResultBuffer();
            object res = null;
            if (args != null)
            {
                TypedValue[] argarray = args.AsArray();
                ObjectId ent;
                string path = "";
                string varName = "";
                dynamic itm = null;
                if (argarray.Length >= 2)
                {                    
                    if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.ObjectId)
                    {
                        ent = (ObjectId)argarray[0].Value;
                        try
                        {
                            itm = Autodesk.Fabrication.Job.GetFabricationItemFromACADHandle(ent.Handle.ToString());                            
                        }
                        catch (System.Exception) { }                        
                    }
                    else if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.Text)
                    {
                        itm = (string)argarray[0].Value;
                        if (itm.StartsWith("$") == true && VarCache.ContainsKey(hand.VariableName(itm)) == true)
                        {
                            itm = VarCache[hand.VariableName(itm)];
                        }
                    }
                    if (argarray.Length >= 2 && argarray[1].TypeCode == (int)LispDataType.Text)
                    { path = (string)argarray[1].Value; }
                    if (argarray.Length >= 3 && argarray[2].TypeCode == (int)LispDataType.Text)
                    { varName = hand.VariableName((string)argarray[2].Value); }
                }
                if (itm != null && path.Length > 1)
                {
                    res = Sequencing.GenericPropertyGetter(itm, new List<string>(path.Split('.')));
                    if (varName.Length >= 1)
                    {
                        if(VarCache.ContainsKey(varName) == true)
                        {
                            VarCache[varName] = res;
                        }
                        else
                        {
                            VarCache.Add(varName, res);
                        }
                    }
                }
            }
            if (res != null)
            {   
                foreach (TypedValue x in hand.ConvertToRB(res))
                { ret.Add(x); }
            }
            else
            {
                ret.Add(new TypedValue((int)LispDataType.Nil));                
            }
            if(ret.AsArray().Length > 1)
            {
                return ret;
            }
            else
            {
                return ret.AsArray()[0];
            }            
        }

        
        // (FabSetItemProp ENAME PathToProp NewValue)
        // Purpose: This will use reflection on the provided object to navigate its properties as specified in the
        //          users PathToProp string and set its value as specified by the users NewValue. Note that there
        //          is nothing more than a try/catch to prevent improper value specifications by the user. Type
        //          conversions will be performed on the NewValue provided as needed to make lisp data types usable
        //          in .Net, but in the case of a string it will also check to see if it starts with a '$' character.
        //          If it does, then it will use the value of a GCV entry in place of the provided NewValue string.
        //          This is especially helpful for properties that require complex object stuctures that do not
        //          translate into lisp data types.
        //          Further note that you can use a $VarName in place of the Ename and "THIS" as a path reference.
        [LispFunction("FabSetItemProp")]
        public bool FabSetItemProp(ResultBuffer args)
        {
            bool ret = false;
            dynamic res = null;
            if (args != null)
            {
                TypedValue[] argarray = args.AsArray();
                ObjectId ent;
                string path = "";                
                dynamic value = null;
                dynamic itm = null;
                if (argarray.Length >= 2)
                {
                    if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.ObjectId)
                    {
                        ent = (ObjectId)argarray[0].Value;
                        try { itm = Autodesk.Fabrication.Job.GetFabricationItemFromACADHandle(ent.Handle.ToString()); }
                        catch (System.Exception) { }
                    }
                    else if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.Text)
                    {
                        itm = (string)argarray[0].Value;
                        if (itm.StartsWith("$") == true && VarCache.ContainsKey(hand.VariableName(itm)) == true)
                        {
                            itm = VarCache[hand.VariableName(itm)];
                        }
                    }
                    if (argarray.Length >= 2 && argarray[1].TypeCode == (int)LispDataType.Text)
                    { path = (string)argarray[1].Value; }
                    if (argarray.Length >= 3 )
                    { value = hand.ConvertFromRB(argarray[2]); }
                }
                if (itm != null && path.Length > 1)
                {
                    res = Sequencing.GenericPropertySetter(itm, new List<string>(path.Split('.')), value); 
                }
            }
            return ret;            
        }


        // (FabInvokeItem ENAME PathToObject MethodName ListOfArgs [VarName])
        // Purpose: This will use reflection on the provided object to navigate its properties as specified in the
        //          users PathToProp string. Once found, the property value will then use reflection to see if the
        //          provided MethodName exists on the properties returned value. If it does not exist it will just
        //          return nil to the user. If the method did exist, then the provided ListOfArgs will be broken into
        //          peieces and fed into the Method as specified. Note that there is nothing more than a try/catch
        //          to prevent improper argument value specifications by the user. Type conversions will be performed
        //          on the NewValue provided as needed to make lisp data types usable in .Net, but in the case of a string
        //          it will also check to see if it starts with a '$' character. If it does, then it will use the value
        //          of a GCV entry in place of the provided string argument. This is especially helpful for methods
        //          that require complex object stuctures that do not translate into lisp data types.
        //          Further note that you can use a $VarName in place of the Ename and "THIS" as a path reference.
        [LispFunction("FabInvokeItem")]
        public object FabInvokeItem(ResultBuffer args)
        {
            ResultBuffer ret = new ResultBuffer();
            dynamic pres = null;
            dynamic mres = null;
            Autodesk.AutoCAD.Geometry.Point3dCollection pts = new Autodesk.AutoCAD.Geometry.Point3dCollection();
            if (args != null)
            {
                TypedValue[] argarray = args.AsArray();
                ObjectId ent;
                string path = "";
                string methName = "";
                string varName = "";
                List<object> methArgs = new List<object>();
                dynamic itm = null;
                if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.ObjectId)
                {
                    ent = (ObjectId)argarray[0].Value;
                    try
                    {
                        itm = Autodesk.Fabrication.Job.GetFabricationItemFromACADHandle(ent.Handle.ToString());
                        if(itm.Connectors?.Count >= 1)
                        {
                            Transaction tr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction();
                            DBObject dbo = tr.GetObject((ObjectId)ent, OpenMode.ForRead);
                            (dbo as Entity).GetStretchPoints(pts);
                            tr.Dispose();
                        }
                    }
                    catch (System.Exception) { }                    
                }
                else if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.Text)
                {
                    itm = (string)argarray[0].Value;
                    if (itm.StartsWith("$") == true && VarCache.ContainsKey(hand.VariableName(itm)) == true)
                    {
                        itm = VarCache[hand.VariableName(itm)];
                    }
                }
                if (argarray.Length >= 2 && argarray[1].TypeCode == (int)LispDataType.Text)
                { path = (string)argarray[1].Value; }
                if (argarray.Length >= 3 && argarray[2].TypeCode == (int)LispDataType.Text)
                { methName = hand.VariableName((string)argarray[2].Value); }
                methArgs = hand.ExtractListValues(argarray);
                if (argarray[argarray.Length - 1].TypeCode == (int)LispDataType.Text)
                {
                    varName = hand.VariableName(hand.GetVarNameFromArgs(argarray[argarray.Length - 1]));
                }                
                if (itm != null && path.Length > 1 && methName.Length >= 1)
                {
                    pres = Sequencing.GenericPropertyGetter(itm, new List<string>(path.Split('.')));
                    mres = Sequencing.GenericMethodInvoke(pres, methName, methArgs);
                    if (methName == "GETCONNECTORENDPOINT")
                    {
                        mres = pts.Count >= 1 ? pts : null;
                    }
                    if (varName.Length >= 1)
                    {
                        if (VarCache.ContainsKey(varName) == true)
                        {
                            VarCache[varName] = mres;
                        }
                        else
                        {
                            VarCache.Add(varName, mres);
                        }
                    }
                }
            }            
            if (mres != null)
            {
                foreach (TypedValue x in hand.ConvertToRB(mres))
                {
                    ret.Add(x);
                }
                if (ret.AsArray().Length > 1)
                {
                    return ret;
                }
                else
                {
                    return ret.AsArray()[0];
                }
            }
            else
            { return true; }
        }


        // (FabDumpItem ENAME Integer PathToProp)
        // Purpose: This will use reflection on the provided object to navigate its properties as specified in the
        //          users PathToProp string. Once found, it will again use reflection to compile a string list of
        //          available proerties, methods or both and send that back to the user. This will help make the API
        //          slightly more self documenting from lisp and less of a constant CHM investigation. The integer
        //          dictates whether it outputs to the command line (1)properties, (2)methods, (3)both and any value
        //          greater than or less than would return just a list of properties and methods to lisp.
        //          Further note that you can use a $VarName in place of the Ename and "THIS" as a path reference.
        [LispFunction("FabDumpItem")]
        public dynamic FabDumpItem(ResultBuffer args)
        {
            Autodesk.AutoCAD.EditorInput.Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;            
            List<string> res = new List<string>();
            ObjectId ent;
            string path = "";
            int OutputType = 3;
            dynamic itm = null;
            if (args != null)
            {
                TypedValue[] argarray = args.AsArray();
                if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.ObjectId)
                {
                    ent = (ObjectId)argarray[0].Value;
                    try { itm = Autodesk.Fabrication.Job.GetFabricationItemFromACADHandle(ent.Handle.ToString()); }
                    catch (System.Exception) { }
                }
                else if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.Text)
                {
                    itm = (string)argarray[0].Value;
                    if (itm.StartsWith("$") == true && VarCache.ContainsKey(hand.VariableName(itm)) == true)
                    {
                        itm = VarCache[hand.VariableName(itm)];
                    }
                }
                if (argarray.Length >= 2 && argarray[1].TypeCode == (int)LispDataType.Text)
                { path = (string)argarray[1].Value; }
                if (argarray.Length >= 3 && argarray[2].TypeCode == (int)LispDataType.Int16)
                { OutputType = Convert.ToInt32(argarray[2].Value); }
                
                if (itm != null && path.Length > 1)
                {
                    res = Sequencing.GetReflectedTypesList(itm, new List<string>(path.Split('.')));
                    if (res.Count >= 1)
                    {
                        ed.WriteMessage("\n\n");
                        foreach (string s in res)
                        {
                            if(s.StartsWith("Property:") == true)
                            {
                                if (OutputType == 1 || OutputType == 3)
                                {
                                    ed.WriteMessage(s + "\n");
                                }
                            }
                            else if(s.StartsWith("Method:") == true)
                            {
                                if (OutputType == 2 || OutputType == 3)
                                {
                                    ed.WriteMessage(s + "\n");
                                }
                            }
                        }                        
                    }
                }
            }
            ResultBuffer rb = new ResultBuffer();
            rb.Add(new TypedValue((int)LispDataType.ListBegin));
            foreach (string s in res)
            {
                rb.Add(new TypedValue((int)LispDataType.Text, s));
            }
            rb.Add(new TypedValue((int)LispDataType.ListEnd));
            if( OutputType >= 1 && OutputType <= 3)
            {
                return null;
            }
            else
            {
                return rb;
            }
        }




        [LispFunction("FabGetDBItem")]
        public dynamic FabGetDBItem(ResultBuffer args)
        {
           
            ResultBuffer ret = new ResultBuffer();
            object res = null;
            if (args != null)
            {
                TypedValue[] argarray = args.AsArray();
                ObjectId ent;
                string path = "";
                string varName = "";
                dynamic itm = null;
                if (argarray.Length >= 2)
                {
                    if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.ObjectId)
                    {
                        ent = (ObjectId)argarray[0].Value;
                        try
                        {
                            itm = Autodesk.Fabrication.Job.GetFabricationItemFromACADHandle(ent.Handle.ToString());
                        }
                        catch (System.Exception) { }
                    }
                    else if (argarray.Length >= 1 && argarray[0].TypeCode == (int)LispDataType.Text)
                    {
                        itm = (string)argarray[0].Value;
                        if (itm.StartsWith("$") == true && VarCache.ContainsKey(hand.VariableName(itm)) == true)
                        {
                            itm = VarCache[hand.VariableName(itm)];
                        }
                    }
                    if (argarray.Length >= 2 && argarray[1].TypeCode == (int)LispDataType.Text)
                    { path = (string)argarray[1].Value; }
                    if (argarray.Length >= 3 && argarray[2].TypeCode == (int)LispDataType.Text)
                    { varName = hand.VariableName((string)argarray[2].Value); }
                }
                if (itm != null && path.Length > 1)
                {
                    res = Sequencing.GenericPropertyGetter(itm, new List<string>(path.Split('.')));
                    if (varName.Length >= 1)
                    {
                        if (VarCache.ContainsKey(varName) == true)
                        {
                            VarCache[varName] = res;
                        }
                        else
                        {
                            VarCache.Add(varName, res);
                        }
                    }
                }
            }
            if (res != null)
            {
                foreach (TypedValue x in hand.ConvertToRB(res))
                { ret.Add(x); }
            }
            else
            {
                ret.Add(new TypedValue((int)LispDataType.Nil));
            }
            if (ret.AsArray().Length > 1)
            {
                return ret;
            }
            else
            {
                return ret.AsArray()[0];
            }
        }



    }
}

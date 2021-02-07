//      Fabrication ITEM .Net Wrapper 1.0
//      Written by Josh Howard 2017
//      www.houseofbim.com

using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using root = AgileBIM.FabWrapper.APIwrappers;

namespace AgileBIM.FabWrapper
{    
    public static class TypeHandlers
    {
        static char[] trs = { ' ', '$', '[', ']', '(', ')', '\n' };

        public static Autodesk.AutoCAD.Geometry.Point3d ToAcadPoint(this Autodesk.Fabrication.Geometry.Point3D pt)
        { return new Autodesk.AutoCAD.Geometry.Point3d(pt.X, pt.Y, pt.Z); }


        public static string FormatReflectedOutput(string Name, string NameValue, string TypeName, string TypeValue)
        {
            string ret = Name;
            ret = ret.PadRight(20) + NameValue;
            ret = ret.PadRight(50) + TypeName;
            ret = ret.PadRight(70) + TypeValue;
            return ret;
        }


        public static string VariableName(string name)
        {
            string ret = name.Replace(" ", "");
            return ret.ToUpper().Trim(trs);
        }


        public static string GetVarNameFromArgs(TypedValue LastArg)
        {
            string ret = "";
            if (LastArg.TypeCode == (int)LispDataType.Text)
            {
                ret = (string)LastArg.Value;
            }
            return ret;
        }


        public static List<object> ExtractListValues(TypedValue[] args)
        {
            List<object> ret = new List<object>();
            object res = null;
            bool flag = false;
            foreach (TypedValue tv in args)
            {
                if (tv.TypeCode == (int)LispDataType.ListBegin)
                {
                    flag = true;
                }
                else if (tv.TypeCode == (int)LispDataType.ListEnd)
                {
                    flag = false;
                }
                else if (flag == true)
                {
                    res = ConvertFromRB(tv);
                    if (typeof(string) == res.GetType())
                    {
                        if (res.ToString().Trim().StartsWith("$") == true)
                        {
                            ret.Add(TryPullVar(res));
                        }
                        else { ret.Add(res); }
                    }
                    else { ret.Add(res); }                    
                }
            }
            return ret;
        }
        

        public static object TryPullVar(dynamic val)
        {
            object ret;
            if (typeof(string) == val.GetType())
            {
                if (val.Trim().StartsWith("$") == true)
                {
                    if (root.VarCache.ContainsKey(VariableName(val)) == true)
                    {
                        ret = root.VarCache[VariableName(val)];
                    }
                    else { ret = val; }
                }
                else
                { ret = val; }
            }
            else
            { ret = val; }
            return ret;
        }
                

        public static dynamic ConvertFromRB(TypedValue r)
        {
            dynamic ret;
            switch (r.TypeCode)
            {
                case (int)LispDataType.T_atom:
                    ret = true; break;
                case (int)LispDataType.Double:
                case (int)LispDataType.Angle:
                    ret = Convert.ToDouble(r.Value); break;
                case (int)LispDataType.Nil:
                    ret = false; break;
                case (int)LispDataType.Int16:
                case (int)LispDataType.Int32:                
                    ret = Convert.ToInt32(r.Value); break;
                case (int)LispDataType.ObjectId:
                    ret = (ObjectId)r.Value; break;
                case (int)LispDataType.Point2d:
                    ret = (Autodesk.AutoCAD.Geometry.Point2d)r.Value; break;
                case (int)LispDataType.Point3d:
                    ret = (Autodesk.AutoCAD.Geometry.Point3d)r.Value; break;
                case (int)LispDataType.Text:
                    ret = (string)r.Value;
                    if (ret.Trim() == "NULL") { ret = null; }
                    else if (ret.Trim().StartsWith("$") == true)
                    {
                        if (root.VarCache.ContainsKey(VariableName(ret)) == true)
                        {
                            ret = root.VarCache[VariableName(ret)];
                        }
                    }
                    break;
                default:
                    ret = null;
                    break;                    
            }
            return ret;
        }


        public static List<TypedValue> ConvertToRB(object value)
        {
            List<TypedValue> ret = new List<TypedValue>();
            Type vt = value.GetType();
            if (typeof(string) == vt)
            {
                ret.Add(new TypedValue((int)LispDataType.Text, (string)value));
            }
            else if (typeof(int) == vt)
            {
                ret.Add(new TypedValue((int)LispDataType.Int32, (int)value));
            }
            else if (typeof(double) == vt)
            {
                ret.Add(new TypedValue((int)LispDataType.Double, (double)value));
            }
            else if (typeof(bool) == vt && (bool)value == true)
            {
                ret.Add(new TypedValue((int)LispDataType.T_atom));
            }
            else if (typeof(Autodesk.Fabrication.Geometry.Point3D) == vt)
            {
                ret.Add(new TypedValue((int)LispDataType.Point3d, (value as Autodesk.Fabrication.Geometry.Point3D).ToAcadPoint()));
            }
            else if (typeof(Autodesk.AutoCAD.Geometry.Point3d) == vt)
            {
                ret.Add(new TypedValue((int)LispDataType.Point3d, value));
            }            
            else if (typeof(Autodesk.AutoCAD.Geometry.Point3dCollection) == vt)
            {
                foreach (Autodesk.AutoCAD.Geometry.Point3d pt in (Autodesk.AutoCAD.Geometry.Point3dCollection)value)
                {
                    ret.Add(new TypedValue((int)LispDataType.Point3d, pt));
                }
            }
            else if (typeof(Enum) == vt.BaseType)
            {
                ret.Add(new TypedValue((int)LispDataType.Int32, (int)value));
            }
            else
            {
                if (value != null)
                {
                    string test = null;
                    try
                    {
                        test = value.ToString();
                        ret.Add(new TypedValue((int)LispDataType.Text, (string)test));
                    }
                    catch (System.Exception)
                    { ret.Add(new TypedValue((int)LispDataType.Nil)); }
                }
                else
                {
                    ret.Add(new TypedValue((int)LispDataType.Nil));
                }
            }
            return ret;
        }

        
    }
}

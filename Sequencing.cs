//      Fabrication ITEM .Net Wrapper 1.0
//      Written by Josh Howard 2017
//      www.houseofbim.com

using System;
using System.Collections.Generic;
using System.Reflection;
using hand = AgileBIM.FabWrapper.TypeHandlers;

namespace AgileBIM.FabWrapper
{
    class Sequencing
    {
        static char[] trs = { ' ', '$', '[', ']', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
        
                
        public static object GenericPropertyGetter(dynamic co, List<string> lst)
        {            
            object ret = null;
            if (lst.Count >= 1 && lst[0].ToUpper() == "THIS")
            {
                lst.RemoveAt(0);
                ret = GenericPropertyGetter(co, lst);
            }
            else if (co != null)
            {
                System.Reflection.PropertyInfo[] props = co.GetType().GetProperties();
                if (lst.Count >= 1)
                {
                    if (lst[0].EndsWith("]") == true)
                    {
                        int si = Convert.ToInt32(lst[0].Substring(lst[0].LastIndexOf('[') + 1).Trim(']'));
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                ret = GenericPropertyGetter(pi.GetValue(co, null)[si], lst);
                                break;
                            }
                        }
                    }
                    else
                    {
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                ret = GenericPropertyGetter(pi.GetValue(co, null), lst);
                                break;
                            }
                        }
                    }

                }
                else
                {
                    ret = co;
                }
            }
            return ret;
        }


        public static bool GenericPropertySetter(dynamic co, List<string> lst, dynamic val)
        {
            bool ret = false;
            if (lst.Count > 1 && lst[0].ToUpper() == "THIS")
            {
                lst.RemoveAt(0);
                ret = GenericPropertySetter(co, lst, val);
            }
            else if (co != null)
            {
                System.Reflection.PropertyInfo[] props = co.GetType().GetProperties();
                if (lst.Count >= 1)
                {
                    if (lst[0].EndsWith("]") == true)
                    {
                        int si = Convert.ToInt32(lst[0].Substring(lst[0].LastIndexOf('[') + 1).Trim(']'));
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                if (lst.Count == 0)
                                {
                                    pi.SetValue(co[si], val);
                                    if (pi.GetValue(co, null)[si] == val)
                                    { ret = true; }
                                    else { ret = false; }
                                }
                                else
                                {
                                    ret = GenericPropertySetter(pi.GetValue(co, null)[si], lst, val);
                                }
                                break;
                            }
                        }
                    }
                    else
                    {
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                if (lst.Count == 0)
                                {
                                    pi.SetValue(co, val);
                                    if (pi.GetValue(co, null) == val)
                                    { ret = true; }
                                    else { ret = false; }
                                }
                                else
                                {
                                    ret = GenericPropertySetter(pi.GetValue(co, null), lst, val);
                                }
                                break;
                            }
                        }
                    }
                }
            }
            return ret;
        }

        
        public static object GenericMethodInvoke(object co, string MethodName, List<object> values)
        {            
            object ret = null;
            if (co != null)
            {
                MethodInfo[] meths = co.GetType().GetMethods();
                foreach (MethodInfo mi in meths)
                {
                    if (mi.Name.ToUpper() == MethodName.ToUpper())
                    {
                        try
                        { ret = mi.Invoke(co, values.ToArray()); }
                        catch (System.Exception) { }
                    }
                }
            }            
            return ret;
        }

        
        public static List<string> GetReflectedTypesList(dynamic co, List<string> lst)
        {
            List<string> ret = new List<string>();
            if (lst.Count >= 1 && lst[0].ToUpper() == "THIS")
            {
                lst.RemoveAt(0);
                ret = GetReflectedTypesList(co, lst);
            }
            else if (co != null)
            {
                System.Reflection.PropertyInfo[] props = co.GetType().GetProperties();
                System.Reflection.MethodInfo[] meths = co.GetType().GetMethods();
                if (lst.Count >= 1)
                {
                    if (lst[0].EndsWith("]") == true)
                    {
                        int si = Convert.ToInt32(lst[0].Substring(lst[0].LastIndexOf('[') + 1).Trim(']'));
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                ret = GetReflectedTypesList(pi.GetValue(co, null)[si], lst);
                                break;
                            }
                        }                        
                    }
                    else
                    {
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                ret = GetReflectedTypesList(pi.GetValue(co, null), lst);
                                break;
                            }                            
                        }
                    }
                    
                }
                else
                {
                    string item;
                    foreach (PropertyInfo pi in props)
                    {
                        item = hand.FormatReflectedOutput("Property:", pi.Name, "ValueType:", pi.PropertyType.Name);
                        if (!ret.Contains(item)){ ret.Add(item); }
                    }
                    foreach (MethodInfo mi in meths)
                    {
                        item = hand.FormatReflectedOutput("Method:", mi.Name, "ReturnType:", mi.ReturnType.Name);
                        if (mi.Name.ToUpper().StartsWith("SET_") == false && mi.Name.ToUpper().StartsWith("GET_") == false)
                        {
                            if (!ret.Contains(item)) { ret.Add(item); }
                        }
                    }
                }                
            }
            return ret;
        }




        public static List<string> GetReflectedDBList(dynamic co, List<string> lst)
        {
            //typeof(Autodesk.Fabrication.DB).GetMembers()
            List<string> ret = new List<string>();
            if (lst.Count >= 1 && lst[0].ToUpper() == "THIS")
            {
                lst.RemoveAt(0);
                ret = GetReflectedTypesList(co, lst);
            }
            else if (co != null)
            {
                System.Reflection.PropertyInfo[] props = co.GetType().GetProperties();
                System.Reflection.MethodInfo[] meths = co.GetType().GetMethods();
                if (lst.Count >= 1)
                {
                    if (lst[0].EndsWith("]") == true)
                    {
                        int si = Convert.ToInt32(lst[0].Substring(lst[0].LastIndexOf('[') + 1).Trim(']'));
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                ret = GetReflectedTypesList(pi.GetValue(co, null)[si], lst);
                                break;
                            }
                        }
                    }
                    else
                    {
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi.Name.ToUpper() == lst[0].Trim(trs).ToUpper())
                            {
                                lst.RemoveAt(0);
                                ret = GetReflectedTypesList(pi.GetValue(co, null), lst);
                                break;
                            }
                        }
                    }

                }
                else
                {
                    string item;
                    foreach (PropertyInfo pi in props)
                    {
                        item = hand.FormatReflectedOutput("Property:", pi.Name, "ValueType:", pi.PropertyType.Name);
                        if (!ret.Contains(item)) { ret.Add(item); }
                    }
                    foreach (MethodInfo mi in meths)
                    {
                        item = hand.FormatReflectedOutput("Method:", mi.Name, "ReturnType:", mi.ReturnType.Name);
                        if (mi.Name.ToUpper().StartsWith("SET_") == false && mi.Name.ToUpper().StartsWith("GET_") == false)
                        {
                            if (!ret.Contains(item)) { ret.Add(item); }
                        }
                    }
                }
            }
            return ret;
        }

    }
}

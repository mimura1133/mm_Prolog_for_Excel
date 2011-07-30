using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

using Microsoft.Win32;

using Prolog;
using Prolog.Code;

namespace PrologExcel
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class PrologExcel
    {
        public string Prolog_List(object obj)
        {
            string ret = "";
            int i,j;

            if ((obj as Excel.Range) != null)
            {
                object[,] r = (obj as Excel.Range).Value2 as object[,];

                if (r != null)
                {
                    for (i = 1; i < (r.GetLength(0)+1); i++)
                    {
                        int v = r.GetLength(1);
                        if (r.GetLength(1) >= 2)
                        {
                            if (r[i, 1] != null && r.GetLength(1)>1)
                            {
                                ret += r[i, 1] + "(";
                                for (j = 2; j < (r.GetLength(1)+1); j++)
                                {
                                    ret += r[i, j] + ",";
                                }
                                ret = ret.Substring(0,ret.Length-1) +  ").\n";
                            }
                        }
                    }
                }
            }
            return ret;
        }

        public string Prolog(string program, string query)
        {
            string ret = "";
            Query pro_query;
            Program pro_program = new Program();
            CodeSentence[] pro_pro_sent;
            CodeSentence[] pro_query_sent;

            try
            {
                pro_pro_sent = Parser.Parse(program);
            }
            catch (Exception e)
            {
                throw new Exception("PROGRAM ERROR : " + e.ToString());
            }
            try
            {
                pro_query_sent = Parser.Parse(query);
            }
            catch (Exception e)
            {
                throw new Exception("QUERY ERROR : " + e.ToString());
            }
            if (pro_pro_sent.Length == 0)
                throw new Exception("VALID PROGRAM CODE NOT AVAILABLE");
            if (pro_query_sent.Length == 0)
                throw new Exception("VALID QUERY CODE NOT AVAILABLE");

            foreach (CodeSentence s in pro_pro_sent)
            {
                pro_program.Add(s);
            }
            pro_query = new Query(pro_query_sent[0]);

            PrologMachine m = PrologMachine.Create(pro_program, pro_query);
            m.RunToSuccess();

            ret = "";
            foreach (var v in m.QueryResults.Variables)
            {
                ret += v.Text + "\n";
            }
            ret = ret.Substring(0, ret.Length - 1);

            return ret;
        }

        [ComRegisterFunctionAttribute]
        private static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(
              GetSubKeyName(type, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(
              GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll",
              RegistryValueKind.String);
        }

        [ComUnregisterFunctionAttribute]
        private static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(
                            GetSubKeyName(type, "Programmable"), false);
        }
        private static string GetSubKeyName(Type type,
          string subKeyName)
        {
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            s.Append(@"CLSID\{");
            s.Append(type.GUID.ToString().ToUpper());
            s.Append(@"}\");
            s.Append(subKeyName);
            return s.ToString();
        }

    }
}

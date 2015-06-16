using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrystalUtility
{
    class Utility
    {
        public String RemoveQuotes(String strToRet)
        {
            strToRet = strToRet.Replace("\"", "");
            return strToRet;
        }

        public Boolean UsesExec(String str)
        {
            String t_string = str.ToUpper();
            Boolean retVal = false;
            if (t_string.Substring(0, 15).Contains("EXEC") ||
                t_string.Substring(0, 15).Contains("CALL"))
            {
                retVal = true;
            }
            return retVal;
        }

        public Boolean HasProduct(String str)
        {
            String t_string = str.ToUpper();
            Boolean retVal = false;
            if (!UsesExec(str))
            {
                if(t_string.Contains("PRODUCT.DBO") ||
                    t_string.Contains("PRODUCT.."))
                {
                    retVal = true;
                }
            }
            return retVal;
        }

        public Boolean HasView(String str)
        {
            String t_string = str.ToUpper();
            Boolean retVal = false;

            if (t_string.Contains("_VW"))
            {
                retVal = true;
            }

            return retVal;
        }

        public Boolean HasProc(String str)
        {
            String t_string = str.ToUpper();
            Boolean retVal = false;

            if (t_string.Contains("PROC("))
            {
                retVal = true;
            }

            return retVal;
        }

        public Boolean HasSQL(String str)
        {
            String t_string = str.ToUpper();
            Boolean retVal = false;

            if (t_string.Contains("SELECT") &&
                t_string.Contains("FROM"))
            {
                retVal = true;
            }

            return retVal;
        }

        public void WriteToFile(String path, String fileName, List<String> listOfStrings)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + fileName);
            for (int i = 0; i < listOfStrings.Count; i++)
            {
                file.WriteLine(listOfStrings[i].ToString(), true);
            }

            file.Close();
        }
    }
}

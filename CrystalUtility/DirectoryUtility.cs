using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CrystalUtility
{
    class DirectoryUtility
    {
        public Boolean CreateDirectory(String path)
        {
            bool retVal = false;
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));
                retVal = true;
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
                retVal = false;
            }
            finally { }
            return retVal; 
        }

        public Boolean DeleteDirectory(String path)
        {
            bool retVal = false;
            try
            {
                Directory.Delete(path);
                Console.WriteLine("The directory was deleted successfully at {0}.", Directory.GetCreationTime(path));
                retVal = true;
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
                retVal = false;
            }
            finally { }
            return retVal;
        }

        /**
         * If sourceDirName is a file, destDirName must be a file as well
         * 
         */
        public Boolean MoveFileToDirectory(string sourceDirName, string destDirName)
        {
            bool retVal = false;
            try
            {
                Directory.Move(sourceDirName, destDirName);
                retVal = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return retVal;
        }

        public Boolean DoesDirectoryExist(String path)
        {
            return Directory.Exists(path);
        }
    }
}

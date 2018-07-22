using System;
using System.IO;

public static class Utilities
{
    public static void MoveFile(string sSourcePath, string txtSourceFile, string sTargetPath, string sSourceSystem, string txtEmail, string txtEmailcc)

    {
        string spath = sSourcePath + txtSourceFile;
        string tpath = sTargetPath + txtSourceFile;

        try 
        {
            if (!File.Exists(spath)) 
            {
                // This statement ensures that the file is created,
                // but the handle is not kept.
                using (FileStream fs = File.Create(spath)) {}
            }

            // Ensure that the target does not exist.
            if (File.Exists(tpath))	
                File.Delete(tpath);

            // Move the file.
            File.Move(spath, tpath);
            //Console.WriteLine("{0} was moved to {1}.", path, path2);

            // See if the original exists now.
            if (File.Exists(spath)) 
            {
                //Console.WriteLine("The original file still exists, which is unexpected.");
            } 
            else 
            {
                //Console.WriteLine("The original file no longer exists, which is expected.");
            }			

        } 
        catch (Exception e) 
        {
            //Console.WriteLine("The process failed: {0}", e.ToString());
        }
    }
}
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

class TestClass
{
    static void Usage()
    {
        Console.WriteLine("Usage");
        Console.WriteLine("    CopyExcel.exe <sourceFile> <destFolder>");
        Console.WriteLine("    ex. CopyExcel.exe C:\\work\\test.xlsx C:\\work\\resultFolder");
    }

    static int preCopyExcel(string sourceFilepath, string destDirpath, string destFilepath)
    {
        if (!System.IO.File.Exists(sourceFilepath))
        {
            Console.WriteLine(string.Format("{0} does not exist.", sourceFilepath));
            return -1;
        }

        if (System.IO.File.Exists(destFilepath))
        {
            Console.WriteLine(string.Format("{0} already exists.", destFilepath));
            return -1;
        }

        if (!System.IO.Directory.Exists(destDirpath))
        {
            try
            {
                Directory.CreateDirectory(destDirpath);
            }
            catch(Exception e)
            {
                Console.WriteLine(string.Format("Creating {0} failed. exception: {1}", destDirpath, e));
                return -1;
            }
        }
        return 0;
    }

    static int copyExcel(string sourceFilepath, string destFilepath)
    {
        try
        {
            // Open sourceFilepath readonly.
            FileStream fs = new FileStream(sourceFilepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            XLWorkbook sourceWorkbook = new XLWorkbook(fs);
            XLWorkbook destWorkbook = new XLWorkbook();

            foreach (IXLWorksheet sourceWorksheet in sourceWorkbook.Worksheets)
            {
                var destWorksheet = destWorkbook.AddWorksheet(sourceWorksheet.Name);
                int rowNum = sourceWorksheet.LastRowUsed().RowNumber();
                for (int i = 1; i <= rowNum; ++i)
                {
                    sourceWorksheet.Row(i).CopyTo(destWorksheet.Row(i));
                }
            }

            destWorkbook.SaveAs(destFilepath);
        }
        catch(Exception e)
        {
            Console.WriteLine(string.Format("Copying failed. exception:{0}", e));
            return -1;
        }
        return 0;
    }

    static int Main(string[] args)
    {
        if (args.Length != 2) 
        {
            Console.WriteLine(string.Format("The number of arguments is invalid. expected:{0}, actual:{1}", 2, args.Length));
            Usage();
            return -1;
        }

        string sourceFilepath = args[0];
        string destDirpath = args[1];
        string destFilepath = Path.Combine(destDirpath, Path.GetFileName(sourceFilepath));

        if (preCopyExcel(sourceFilepath, destDirpath, destFilepath) != 0)
        {
            return -1;
        }

        if (copyExcel(sourceFilepath, destFilepath) != 0)
        {
            return -1;
        }

        Console.WriteLine(string.Format("{0} was successfully copied.", sourceFilepath));
        return 0;
    }
}
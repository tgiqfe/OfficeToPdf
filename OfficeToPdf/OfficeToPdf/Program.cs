
using System.Diagnostics;

var arguments = args.Aggregate(new List<string>(), (list, arg) =>
{
    if (arg.Contains(";"))
    {
        list.AddRange(arg.Split(';'));
    }
    else
    {
        list.Add(arg);
    }
    return list;
});

arguments.ForEach(arg =>
{
    string extension = Path.GetExtension(arg);
    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(arg);    
    string parent = Path.GetDirectoryName(arg);
    string output = Path.Combine(parent, fileNameWithoutExtension + ".pdf");
    using(var proc = new Process())
    {
        proc.StartInfo.WorkingDirectory = @"script";
        proc.StartInfo.FileName = @".\venv\Scripts\python.exe";
        switch (extension)
        {
            case ".xls":
            case ".xlsx":
                proc.StartInfo.Arguments = $"convert_excel.py \"{arg}\" \"{output}\"";
                break;
            case ".doc":
            case ".docx":
                proc.StartInfo.Arguments = $"convert_word.py \"{arg}\" \"{output}\"";
                break;
            case ".ppt":
            case ".pptx":
                proc.StartInfo.Arguments = $"convert_powerpoint.py \"{arg}\" \"{output}\"";
                break;
        }
        proc.Start();
        proc.WaitForExit();
    }
});

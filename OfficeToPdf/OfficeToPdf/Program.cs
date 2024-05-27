
using System.Diagnostics;

if (args.Length == 0) return;

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

string currentDirectory = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);

if(!Directory.Exists(Path.Combine(currentDirectory, "script", "venv")))
{
    using (var proc = new Process())
    {
        proc.StartInfo.UseShellExecute = false;
        proc.StartInfo.CreateNoWindow = false;
        proc.StartInfo.WorkingDirectory = Path.Combine(currentDirectory, "script");

        proc.StartInfo.FileName = "cmd";
        proc.StartInfo.Arguments = "/c InitialSetup.bat";
        proc.Start();
        proc.WaitForExit();
    }
}

arguments.ForEach(arg =>
{
    string extension = Path.GetExtension(arg);
    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(arg);
    string parent = Path.GetDirectoryName(arg);
    string output = Path.Combine(parent, fileNameWithoutExtension + ".pdf");
    using (var proc = new Process())
    {
        proc.StartInfo.UseShellExecute = false;
        proc.StartInfo.CreateNoWindow = false;
        proc.StartInfo.WorkingDirectory = Path.Combine(currentDirectory, "script");
        proc.StartInfo.FileName = @"venv\Scripts\python.exe";
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

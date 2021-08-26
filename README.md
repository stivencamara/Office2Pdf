# Office2Pdf
Convert Office documents to Pdf

# Example
Convert **docx** to pdf

```C#
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.DOCX);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.docx");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testdocx.pdf");

            conveter.Convert(sourcePath, targetPath, false);
```

Convert **xls** to pdf

```C#
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.XLS);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.xls");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testxls.pdf");

            conveter.Convert(sourcePath, targetPath, false);
```

Convert **ppt** to pdf

```C#
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.PPT);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.ppt");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testppt.pdf");

            conveter.Convert(sourcePath, targetPath, false);
```

# How to setup DCOM Configuration Settings to use Office Interop Word, Excel, and Access with ASP.NET C# on a IIS Windows Server machine

**0x800A03EC Cannot access the file**
The 0x800A03EC Cannot access the file is arguably the worst error you can experience, as the given error message is completely misleading. To fix that, you have to the following:

Create the following new folders on your Windows Server + IIS machine:
C:\Windows\SysWOW64\config\systemprofile\Desktop (for 64-bit Servers only)
C:\Windows\System32\config\systemprofile\Desktop (for both 32-bit and 64-bit Servers)
Set Full control permissions for these Desktop folders for the Application Pool user (IIS AppPool\DefaultAppPool if you’re using the ApplicationPoolIdentity dynamic account).

https://www.ryadel.com/en/office-interop-dcom-config-windows-server-iis-word-excel-access-asp-net-c-sharp/

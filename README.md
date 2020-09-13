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

# Future problems with interoperability (ASP.Net IIS)

https://www.ryadel.com/en/office-interop-dcom-config-windows-server-iis-word-excel-access-asp-net-c-sharp/

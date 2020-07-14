# Office2Pdf
Convert Office documents to Pdf

# Example
The following is an example that will convert docx to pdf

```C#
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.DOCX);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.docx");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.pdf");

            conveter.Convert(sourcePath, targetPath, false);
```

The following is an example that will convert xls to pdf

```C#
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.XLS);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.xls");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testxls.pdf");

            conveter.Convert(sourcePath, targetPath, false);
```

The following is an example that will convert xls to pdf

```C#
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.PPT);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.ppt");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testppt.pdf");

            conveter.Convert(sourcePath, targetPath, false);
```

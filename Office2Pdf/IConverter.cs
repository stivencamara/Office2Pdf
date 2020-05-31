using System;

namespace Office2Pdf
{
    public interface IConverter
    {
        void Convert(string sourcePath, string targetPath, bool isPdf);
    }
}

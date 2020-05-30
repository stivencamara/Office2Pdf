using System;

namespace Convert2Office
{
    public interface IConverter
    {
        void Convert(string sourcePath, string targetPath, bool isPdf);
    }
}

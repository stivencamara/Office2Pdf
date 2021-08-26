using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Office2Pdf
{
    internal abstract class ConverterBase<T> : IConverter, IDisposable
    {
        private readonly T application;

        private protected ConverterBase(T application)
        {
            this.application = application;
        }

        public T Application => application;

        public void Convert(string sourcePath, string targetPath, bool isPdfA)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new Exception("The source path is required");

            if (string.IsNullOrEmpty(targetPath))
                throw new Exception("The target path is required");

            OnConvert(sourcePath, targetPath, isPdfA);

            Dispose();
        }

        public abstract void OnConvert(string sourcePath, string targetPath, bool isPdfA);

        protected void ReleaseObject(T obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = default(T);
            }

            catch
            {
                obj = default(T);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected ContentType GetFileExtension(string sourcePath)
        {
            var extension = Path.GetExtension(sourcePath);

            extension = extension.Replace(".", string.Empty);
            return (ContentType)Enum.Parse(typeof(ContentType), extension, true);
        }

        public void Dispose()
        {
            ReleaseObject(application);
        }
    }
}

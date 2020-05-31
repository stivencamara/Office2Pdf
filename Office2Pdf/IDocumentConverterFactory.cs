namespace Office2Pdf
{
    public interface IDocumentConverterFactory
    {
        IConverter GetConverter(ContentType conversionType);
    }
}

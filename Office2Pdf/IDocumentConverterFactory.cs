namespace Convert2Office
{
    public interface IDocumentConverterFactory
    {
        IConverter GetConverter(ContentType conversionType);
    }
}

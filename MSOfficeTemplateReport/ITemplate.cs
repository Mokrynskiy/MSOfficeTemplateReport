namespace MSOfficeTemplateReport
{
    public interface ITemplate
    {
        void AddVariable(string name, object data);
        void Generate();

        string SaveAs(string path);

        byte[] ToByteArray();
    }
}

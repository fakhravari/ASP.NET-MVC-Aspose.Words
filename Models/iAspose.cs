using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.MailMerging;
using System;
using System.Drawing;
using System.IO;
using System.Web;

namespace WebApplication2.Models
{
    public static class iAspose
    {
        public static string templateurl;
        public static string BuildPrintLetter(string[] field, object[] data, string template, string SaveDocFilePath)
        {
            templateurl = template;

            if (!Directory.Exists(HttpContext.Current.Server.MapPath(SaveDocFilePath)))
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath(SaveDocFilePath));

            Document my = new Document(HttpContext.Current.Server.MapPath(template));

            my.MailMerge.FieldMergingCallback = new HandleMergeFieldAlternatingRows();
            my.MailMerge.Execute(field, data);

            string SavedFilePathName = "PIGIS_Doc_" + DateTime.Now.Millisecond + ".docx";
            my.Save(HttpContext.Current.Server.MapPath("~/PigisFileServer/Temp/" + SavedFilePathName));

            return SavedFilePathName;
        }


        private class HandleMergeFieldAlternatingRows : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                try
                {
                    if (mBuilder == null)
                        mBuilder = new DocumentBuilder(e.Document);

                    if (e.FieldName.Equals("Shomareh"))
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;

                        Run run = new Run(mBuilder.Document);
                        run.Text = e.FieldValue.ToString();

                        Aspose.Words.Font font = run.Font;
                        font.Bidi = false;
                        if (!templateurl.Contains("English"))
                        {
                            font.Bold = true;
                            font.Name = "B Nazanin";
                            font.Size = 9;
                            font.Color = Color.Black;
                        }
                        else
                        {
                            font.Name = "Arial";
                            font.Size = 9;
                            font.Color = Color.Black;
                        }
                        mBuilder.InsertNode(run);
                    }
                    if (e.FieldName.Equals("Tarikh"))
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;

                        Run run = new Run(mBuilder.Document);
                        run.Text = e.FieldValue.ToString();

                        Aspose.Words.Font font = run.Font;

                        if (!templateurl.Contains("English"))
                        {
                            font.Bold = true;
                            font.Name = "B Nazanin";
                            font.Size = 12;
                            font.Color = Color.Black;
                        }
                        else
                        {

                            font.Name = "Arial";
                            font.Size = 11;
                            font.Color = Color.Black;
                        }
                        mBuilder.InsertNode(run);
                    }
                    if (e.FieldName.Equals("Semat"))
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;

                        Run run = new Run(mBuilder.Document);
                        run.Text = e.FieldValue.ToString();
                        Aspose.Words.Font font = run.Font;
                        font.Bold = true;
                        font.Name = "B Nazanin";
                        font.Size = 12;
                        font.Color = Color.Black;
                        mBuilder.InsertNode(run);
                    }
                    if (e.DocumentFieldName.StartsWith("Html"))
                    {
                        DocumentBuilder builder = new DocumentBuilder(e.Document);
                        builder.MoveToMergeField(e.DocumentFieldName);
                        builder.InsertHtml((string)e.FieldValue);
                        e.Text = "";
                    }
                    if (e.DocumentFieldName.StartsWith("Image"))
                    {
                        DocumentBuilder builder = new DocumentBuilder(e.Document);
                        builder.MoveToMergeField(e.DocumentFieldName);
                        Shape shape = builder.InsertImage(e.FieldValue.ToString());
                        shape.WrapType = WrapType.None;
                        shape.BehindText = true;
                        shape.Width = 135;
                        shape.Height = 110;
                    }
                }
                catch (Exception ex)
                {
                    string Field = e.FieldName;
                    string Value = e.FieldValue.ToString();
                }
            }
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // throw new NotImplementedException();
            }

            public DocumentBuilder mBuilder { set; get; }
        }
    }
}
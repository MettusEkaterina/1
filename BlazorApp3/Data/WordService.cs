using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace BlazorApp3.Data
{
    public class WordService
    {
        VigenereCipherDocx vigenereCipherDocx;
        private void IterateTextBody(WTextBody textBody)
        {
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        IterateParagraph(paragraph);
                        break;
                    case EntityType.Table:
                        IterateTable(bodyItemEntity as WTable);
                        break;
                }
            }
        }

        private void IterateTable(WTable table)
        {
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    IterateTextBody(cell);
                }
            }
        }

        private void IterateParagraph(WParagraph paragraph)
        {
            for (int i = 0; i < paragraph.ChildEntities.Count; i++)
            {
                Entity entity = paragraph.ChildEntities[i];

                if (entity.EntityType == EntityType.TextRange)
                {
                    WTextRange textRange = entity as WTextRange;
                    (entity as WTextRange).Text = vigenereCipherDocx.Handle(textRange.Text);
                }
            }
        }

        public MemoryStream CreateWord(string path, string keyWord, bool encrypting = true)
        {
            vigenereCipherDocx = new VigenereCipherDocx(keyWord, encrypting);
            using (FileStream fileStreamPath = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    using (WordDocument clonedDocument = document.Clone())
                    {
                        foreach (WSection section in clonedDocument.Sections)
                        {
                            WTextBody sectionBody = section.Body;
                            IterateTextBody(sectionBody);
                            WHeadersFooters headersFooters = section.HeadersFooters;
                            IterateTextBody(headersFooters.OddHeader);
                            IterateTextBody(headersFooters.OddFooter);
                        }

                        using (MemoryStream stream = new MemoryStream())
                        {
                            clonedDocument.Save(stream, FormatType.Docx);
                            clonedDocument.Close();
                            stream.Position = 0;
                            return stream;
                        }
                    }
                }
            }
        }

        public MemoryStream CreateWord(string resultText) 
        { 
            WordDocument document = new WordDocument();
            WSection section = document.AddSection() as WSection;
            section.PageSetup.Margins.All = 72;
            section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);
            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Arial";
            style.CharacterFormat.FontSize = 13.5f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 8;

            IWParagraph paragraph = section.AddParagraph();
            WTextRange textRange = paragraph.AppendText(resultText) as WTextRange;
            textRange.CharacterFormat.FontSize = 12f;

            using (MemoryStream stream = new MemoryStream())
            {
                document.Save(stream, FormatType.Docx);
                document.Close();
                stream.Position = 0;
                return stream;
            }
        }
    }
}

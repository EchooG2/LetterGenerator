using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace LetterGenerator
{
    public partial class MainWindow : Window
    {
        private ObservableCollection<Attachment> _attachments;

        public MainWindow()
        {
            InitializeComponent();
            _attachments = new ObservableCollection<Attachment>();
            AttachmentsGrid.ItemsSource = _attachments;
        }

        private void AddAttachment_Click(object sender, RoutedEventArgs e)
        {
            _attachments.Add(new Attachment { Title = "Новое приложение", Content = "Текст приложения..." });
        }

        private void RemoveAttachment_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var attachment = button?.Tag as Attachment;
            if (attachment != null && _attachments.Contains(attachment))
                _attachments.Remove(attachment);
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            // 收集缺失字段
            var missingFields = new System.Collections.Generic.List<string>();
            if (string.IsNullOrWhiteSpace(KomyTextBox.Text)) missingFields.Add("Кому (организация)");
            if (string.IsNullOrWhiteSpace(AddresseeTextBox.Text)) missingFields.Add("Адресат (должность и ФИО)");
            if (string.IsNullOrWhiteSpace(AddresseeNameTextBox.Text)) missingFields.Add("Имя адресата (в дательном падеже)");
            if (string.IsNullOrWhiteSpace(SubjectTextBox.Text)) missingFields.Add("Тема письма");
            if (string.IsNullOrWhiteSpace(BodyTextBox.Text)) missingFields.Add("Текст письма");
            if (string.IsNullOrWhiteSpace(ManagerTextBox.Text)) missingFields.Add("Подпись (должность и ФИО)");

            if (missingFields.Count > 0)
            {
                MessageBox.Show($"Пожалуйста, заполните следующие поля:\n{string.Join("\n", missingFields)}",
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 模板文件路径
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template.docx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show($"Файл шаблона не найден: {templatePath}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 输出文件路径（桌面）
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                                $"Letter_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

            try
            {
                File.Copy(templatePath, outputPath, true);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                    // 替换占位符
                    ReplacePlaceholder(body, "{KOMY}", KomyTextBox.Text);
                    ReplacePlaceholder(body, "{ADDRESSEE}", AddresseeTextBox.Text);
                    ReplacePlaceholder(body, "{ADDRESSEE_NAME}", AddresseeNameTextBox.Text);
                    ReplacePlaceholder(body, "{SUBJECT}", SubjectTextBox.Text);
                    ReplacePlaceholder(body, "{TEXT_BODY}", BodyTextBox.Text);
                    ReplacePlaceholder(body, "{MANAGER_NAME}", ManagerTextBox.Text);

                    // 插入附件列表（在签名之前）
                    if (_attachments.Count > 0)
                    {
                        InsertAttachmentList(body, _attachments);
                    }

                    // 生成附件内容页
                    GenerateAttachmentPages(body, _attachments);

                    doc.MainDocumentPart.Document.Save();
                }

                MessageBox.Show($"Документ успешно создан:\n{outputPath}", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
                System.Diagnostics.Process.Start("explorer.exe", $"/select, \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReplacePlaceholder(Body body, string placeholder, string newValue)
        {
            foreach (var text in body.Descendants<Text>())
            {
                if (text.Text.Contains(placeholder))
                {
                    text.Text = text.Text.Replace(placeholder, newValue);
                }
            }
        }

        private void InsertAttachmentList(Body body, ObservableCollection<Attachment> attachments)
        {
            var lastParagraph = body.Elements<Paragraph>().LastOrDefault();
            if (lastParagraph == null) return;

            string listText = "Приложение: ";
            for (int i = 0; i < attachments.Count; i++)
            {
                int pageCount = (attachments[i].Content.Length / 2000) + 1;
                listText += $"{i + 1}. {attachments[i].Title} на {pageCount} л. ";
            }

            var attachmentListPara = new Paragraph(new Run(new Text(listText)));
            body.InsertBefore(attachmentListPara, lastParagraph);
        }

        private void GenerateAttachmentPages(Body body, ObservableCollection<Attachment> attachments)
        {
            if (attachments.Count == 0) return;

            for (int i = 0; i < attachments.Count; i++)
            {
                var att = attachments[i];
                string attachmentLabel = attachments.Count == 1 ? "Приложение" : $"Приложение {i + 1}";

                var pageBreak = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                body.AppendChild(pageBreak);

                var markerPara = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Right }));
                markerPara.AppendChild(new Run(new Text(attachmentLabel)));
                body.AppendChild(markerPara);

                var titlePara = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
                titlePara.AppendChild(new Run(new Text(att.Title)));
                body.AppendChild(titlePara);

                var contentPara = new Paragraph(new Run(new Text(att.Content)));
                body.AppendChild(contentPara);
            }
        }
    }
}
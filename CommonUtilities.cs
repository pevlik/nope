using System;
using System.IO;
using System.Windows.Forms;

namespace Specification
{
    public static class CommonUtilities
    {
        public static string PromptSectionNumber(TextBox textBox)
        {
            if (string.IsNullOrEmpty(textBox.Text))
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("Номер Вашей секции:", "Ввод", "");
                if (!string.IsNullOrEmpty(input)) textBox.Text = input;
                return input;
            }
            return textBox.Text;
        }

        public static void EnsureDirectoryExists(string path)
        {
            string directory = Path.GetDirectoryName(path);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }

        public static void ShowSuccess(string message)
        {
            MessageBox.Show(message, "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ShowError(string message)
        {
            MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
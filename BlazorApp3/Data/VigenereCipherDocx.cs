using System;

namespace BlazorApp3.Data
{
    public class VigenereCipherDocx
    {
        const string alphabet = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ";
        string KeyWord
        {
            get
            {
                return key;
            }

            set
            {
                if (value == null || value.Length == 0)
                {
                    throw new Exception("Ключ не задан");
                }

                for (int i = 0; i < value.Length; i++)
                {
                    if (!alphabet.Contains(Char.ToUpper(value[i])))
                    {
                        throw new Exception("Ключ должен содержать только буквы кириллицы");
                    }
                }

                key = value.ToUpper();
            }
        }

        string key;
        int keyCurrentIndex;
        bool encrypting;

        public VigenereCipherDocx(string keyWord, bool encrypting = true)
        {
            KeyWord = keyWord;
            keyCurrentIndex = 0;
            this.encrypting = encrypting;
        }

        public string Handle(string text)
        {
            string result = "";
            var q = alphabet.Length;

            for (int i = 0; i < text.Length; i++)
            {
                bool isUpper = Char.IsUpper(text[i]);
                int letterIndex = isUpper ? alphabet.IndexOf(text[i]) : alphabet.IndexOf(Char.ToUpper(text[i]));
                var codeIndex = alphabet.IndexOf(KeyWord[keyCurrentIndex % KeyWord.Length]);

                if (letterIndex < 0)
                {
                    result += text[i].ToString();
                }
                else
                {
                    result += isUpper ? alphabet[(q + letterIndex + ((encrypting ? 1 : -1) * codeIndex)) % q].ToString() : alphabet[(q + letterIndex + ((encrypting ? 1 : -1) * codeIndex)) % q].ToString().ToLower();
                    keyCurrentIndex++;
                }
            }

            return result;
        }
    }
}

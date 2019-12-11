using System.Threading.Tasks;

namespace BlazorApp3.Data
{
    public class VigenereCipherService
    {
        public Task<string> GetResultTextAsync(string inputText, string keyWord, bool encrypting)
        {
            return Task.FromResult(new VigenereCipher().Handle(inputText, keyWord, encrypting));
        }
    }
}

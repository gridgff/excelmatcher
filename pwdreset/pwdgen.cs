using System;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace pwdgen
{
    class Program
    {
        private static Random random = new Random();
        public static string strgen(int length)
        {
            const string letter = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            StringBuilder stringBuilder = new StringBuilder(length);

            for (int i = 0; i < length; i++)
            {
                stringBuilder.Append(letter[random.Next(letter.Length)]);
            }
            return stringBuilder.ToString();
            
        }
        public static void Main(string[] args)
        {

            string randstr = strgen(2);
            StringBuilder stringBuilder2 = new StringBuilder(100);
            //Console.Write($"Временный пароль на англ.:{randstr}");
            stringBuilder2.Append($"Временный пароль на англ.:{randstr}");
            for (int i = 0; i < 6; i++)
            {
                var rand = new Random();
                stringBuilder2.Append(rand.Next(10));

            }
            stringBuilder2.ToString();

            string fp = "pwd.txt";
            //System.Console.WriteLine(stringBuilder2);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                fp = "C://Dst";
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                fp = "/home/cheetos/vs/pwdreset/pwd.txt";
            }
            char.ToUpper(stringBuilder2[27]);
            System.Console.WriteLine(stringBuilder2[27]);
            using (StreamWriter writer = new StreamWriter(fp, false))
            {
                writer.WriteLine($"{stringBuilder2}");
            }
            Process.Start(new ProcessStartInfo(fp){ UseShellExecute =true});
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneS_Baselinker
{

    class Base64Encoder
    {
        public void Encode(string inFileName, string outFileName)
        {
            System.Security.Cryptography.ICryptoTransform transform = new System.Security.Cryptography.ToBase64Transform();
            using (System.IO.FileStream inFile = System.IO.File.OpenRead(inFileName),
                                      outFile = System.IO.File.Create(outFileName))
            using (System.Security.Cryptography.CryptoStream cryptStream = new System.Security.Cryptography.CryptoStream(outFile, transform, System.Security.Cryptography.CryptoStreamMode.Write))
            {
                // Bufor 4kb
                byte[] buffer = new byte[4096];
                int bytesRead;

                while ((bytesRead = inFile.Read(buffer, 0, buffer.Length)) > 0)
                    cryptStream.Write(buffer, 0, bytesRead);

                cryptStream.FlushFinalBlock();
            }
        }

        public void Decode(string inFileName, string outFileName)
        {
            System.Security.Cryptography.ICryptoTransform transform = new System.Security.Cryptography.FromBase64Transform();
            using (System.IO.FileStream inFile = System.IO.File.OpenRead(inFileName),
                                      outFile = System.IO.File.Create(outFileName))
            using (System.Security.Cryptography.CryptoStream cryptStream = new System.Security.Cryptography.CryptoStream(inFile, transform, System.Security.Cryptography.CryptoStreamMode.Read))
            {
                byte[] buffer = new byte[4096];
                int bytesRead;

                while ((bytesRead = cryptStream.Read(buffer, 0, buffer.Length)) > 0)
                    outFile.Write(buffer, 0, bytesRead);

                outFile.Flush();
            }
        }

        //Cryptostorm powinien zużyć mniej pamięci. Tutaj wrzucane jest wszystko na raz do pamięci, co może być uciążliwe przy dużych plikach.
        //Efekt powinien być ten sam
        public void MemoryEncode(string inFileName, string outFileName)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(inFileName);
            System.IO.File.WriteAllText(outFileName, System.Convert.ToBase64String(bytes));
        }
    }

}

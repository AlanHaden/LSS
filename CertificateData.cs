using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace LSS
{
    [Serializable]
    internal class CertificateData
    {
        public static void SaveDataToFile(string filePath, byte[] imageSignatureData, string[,] formDataArray, string[] formDetails)
        {
            List<List<string>> formDataList = ConvertFormDataArrayToList(formDataArray);

            CertificateData certificateData = new CertificateData
            {
                ImageSignatureData = imageSignatureData,
                FormDataList = formDataList,
                FormDetails = formDetails
            };

            byte[] serializedData = SerializeObject(certificateData);

            File.WriteAllBytes(filePath, serializedData);

            Console.WriteLine("Certificate data has been saved to: " + filePath);
        }

        private static List<List<string>> ConvertFormDataArrayToList(string[,] formDataArray)
        {
            int rows = formDataArray.GetLength(0);
            int columns = formDataArray.GetLength(1);

            List<List<string>> formDataList = new List<List<string>>();

            for (int i = 0; i < rows; i++)
            {
                List<string> rowList = new List<string>();

                for (int j = 0; j < columns; j++)
                {
                    string value = formDataArray[i, j];
                    rowList.Add(value);
                }

                formDataList.Add(rowList);
            }

            return formDataList;
        }

        private static byte[] SerializeObject(object obj)
        {
            if (obj == null)
                return null;

            BinaryFormatter formatter = new BinaryFormatter();
            using (MemoryStream memoryStream = new MemoryStream())
            {
                formatter.Serialize(memoryStream, obj);
                return memoryStream.ToArray();
            }
        }

        public byte[] ImageSignatureData { get; set; }
        public List<List<string>> FormDataList { get; set; }
        public string[] FormDetails { get; set; }

        // Load data -----------------------------------

        public static CertificateData LoadDataFromFile(string filePath)
        {
            byte[] serializedData = File.ReadAllBytes(filePath);
            CertificateData certificateData = DeserializeObject(serializedData);
            return certificateData;
        }

        private static CertificateData DeserializeObject(byte[] serializedData)
        {
            if (serializedData == null)
                return null;

            BinaryFormatter formatter = new BinaryFormatter();
            using (MemoryStream memoryStream = new MemoryStream(serializedData))
            {
                object deserializedObject = formatter.Deserialize(memoryStream);
                if (deserializedObject is CertificateData certificateData)
                {
                    return certificateData;
                }
            }

            return null;
        }
    }
}


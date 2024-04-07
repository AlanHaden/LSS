
namespace LSS
{
    internal class Global
    {
        public static string blankTestCert = "BlankTC.pdf";
        public static string dueDate = "";
        public static string[] companyAddressValues = new string[5];// - User set company details
        public static string pdfImagePath; // - User set company logo path
        public static byte[] imageSignatureData;
        public static string[,] formDataArray = new string[10, 34];
        // 0 - source protected
        // 1 - circuit number
        // 2 - spd manufacturer
        // 3 - spd model
        // 4 - spd type
        // 5 - L1-PE M/G
        // 6 - L1-PE
        // 7 - L2-PE M/G
        // 8 - L2-PE
        // 9 - L3-PE M/G
        //10 - L3-PE
        //11 - N-PE M/G
        //12 - N-PE
        //13 - L1-N M/G
        //14 - L1-N
        //15 - L2-N M/G
        //16 - L2-N
        //17 - L3-N M/G
        //18 - L3-N
        //19 - Zs
        //20 - location
        //21 - bsen
        //22 - type
        //23 - rating
        //24 - short ciruit capacity
        //25 - reference method
        //26 - live mm2
        //27 - cpc mm2
        //28 - R2
        //29 - ir live/live
        //30 - ir live/pe
        //31 - polarity
        //32 - comments
        //33 - pass/fail


        public static string[] formDetails = new string[10];
        // 0 - reference
        // 1 - certificate number
        // 2 - site address
        // 3 - engineer
        // 4 - MFT model
        // 5 - MFT SN
        // 6 - SPD Tester model
        // 7 - SPD Tester SN
        // 8 - date
        // 9 - overall

    }
}

using System.IO;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.Content;

using static LSS.Global;
using System.Windows;
using System.IO.IsolatedStorage;
using System.Diagnostics;
using System.Drawing;
using System;
using System.Globalization;
using System.Windows.Forms;

namespace LSS
{
    public class PDF
    {

        public static void PrintToPDF()
        {

            string savePath = null;

            var dlg = new FolderPicker();
            dlg.InputPath = @"c:\";
            if (dlg.ShowDialog() == true)
            {
                savePath = dlg.ResultPath;
            }


            try
            {

                using (var doc = PdfReader.Open(blankTestCert, PdfDocumentOpenMode.Modify))
                {
                    int a = 93, b = 142, c = 100, d = 200;

                    // Print Page 1 -----------------------------------------------------
                    var page = doc.Pages[0];
                    var contents = ContentReader.ReadContent(page);

                    // Get an XGraphics object for drawing
                    XGraphics gfx = XGraphics.FromPdfPage(page);

                    // Create a font
                    XFont font = new XFont("Calibri", 8, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.WinAnsi));
                    XFont fontSym = new XFont("Wingdings", 12, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.WinAnsi));
                    XFont fontOverall = new XFont("Calibri", 10, XFontStyle.Bold, new XPdfFontOptions(PdfFontEncoding.WinAnsi));

                    // ------- BS declaration 
                    gfx.DrawString("BS7671 / BSEN 62305", font, XBrushes.Black, new XRect(350, 10, 100, 200), XStringFormats.TopRight);
                    // -------- Header

                    // Get an XGraphics object for drawing
                    using (var isolatedStorage = IsolatedStorageFile.GetUserStoreForAssembly())
                    {
                        string imageName = "UserImage.jpg";

                        // Check if the file exists
                        if (isolatedStorage.FileExists(imageName))
                        {
                            // Open the file stream
                            using (var imageStream = isolatedStorage.OpenFile(imageName, FileMode.Open))
                            {
                                // Load the image from the stream
                                XImage image = XImage.FromStream(imageStream);

                                // Draw the image on the PDF page
                                gfx.DrawImage(image, 30, 15, 135, 45);
                            }
                        }
                        else
                        {
                            // File not found in isolated storage
                            System.Windows.MessageBox.Show("Image file not found in isolated storage.");
                        }
                    }
                    if (companyAddressValues[0] != null)
                    {
                        //Address line 1
                        gfx.DrawString(companyAddressValues[0], font, XBrushes.Black, new XRect(710, 15, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[1] != null)
                    {
                        //Address line 2
                        gfx.DrawString(companyAddressValues[1], font, XBrushes.Black, new XRect(710, 24, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[2] != null)
                    {
                        //Address line 3
                        gfx.DrawString(companyAddressValues[2], font, XBrushes.Black, new XRect(710, 33, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[3] != null)
                    {
                        //Address line 4
                        gfx.DrawString(companyAddressValues[3], font, XBrushes.Black, new XRect(710, 42, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[4] != null)
                    {
                        //Telephone number
                        gfx.DrawString(companyAddressValues[4], font, XBrushes.Black, new XRect(710, 51, 100, 200), XStringFormats.TopRight);
                    }

                    // ---- Top table

                    if (formDetails[0] != null)
                    {
                        //Reference
                        gfx.DrawString(formDetails[0], font, XBrushes.Black, new XRect(a + 60, 65, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[1] != null)
                    {
                        //Cert Number top field and p1 footer
                        gfx.DrawString(formDetails[1], font, XBrushes.Black, new XRect(a + 50, 574, 100, 200), XStringFormats.TopLeft);
                        gfx.DrawString(formDetails[1], font, XBrushes.Black, new XRect(a + 465, 65, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[2] != null)
                    {
                        //Site Address
                        gfx.DrawString(formDetails[2], font, XBrushes.Black, new XRect(a + 60, 79, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[3] != null)
                    {
                        //Inspector
                        gfx.DrawString(formDetails[3], font, XBrushes.Black, new XRect(a + 60, 92, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[8] != null)
                    {
                        //Date
                        gfx.DrawString(formDetails[8], font, XBrushes.Black, new XRect(a + 395, 92, 100, 200), XStringFormats.TopLeft);
                        gfx.DrawString(dueDate, font, XBrushes.Black, new XRect(a + 613, 92, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[4] != null)
                    {
                        //MFT Tester
                        gfx.DrawString(formDetails[4], font, XBrushes.Black, new XRect(a + 115, 125, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[5] != null)
                    {
                        //MFT Tester SN
                        gfx.DrawString(formDetails[5], font, XBrushes.Black, new XRect(a + 465, 125, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[6] != null)
                    {
                        //SPD Tester
                        gfx.DrawString(formDetails[6], font, XBrushes.Black, new XRect(a + 115, 137, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[7] != null)
                    {
                        //SPD Tester SN
                        gfx.DrawString(formDetails[7], font, XBrushes.Black, new XRect(a + 465, 137, 100, 200), XStringFormats.TopLeft);
                    }

                    if (formDetails[9] != null)
                    {
                        //Overall
                        if (formDetails[9] == "PASS")
                        {
                            gfx.DrawString(formDetails[9], fontOverall, XBrushes.Black, new XRect(765, 326, 100, 200), XStringFormats.TopLeft);
                        }
                        else
                        {
                            gfx.DrawString(formDetails[9], fontOverall, XBrushes.Red, new XRect(765, 326, 100, 200), XStringFormats.TopLeft);
                        }
                    }

                    try
                    {

                        // Load the image from imgEngSig
                        using (MemoryStream memoryStream = new MemoryStream(imageSignatureData))
                        {
                            XImage imageSig = XImage.FromStream(memoryStream);

                            // Draw the image on the page using the existing gfx object
                            gfx.DrawImage(imageSig, new XRect(680, 485, 121, 38)); // Adjust the position and dimensions as needed
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle the exception
                        System.Windows.MessageBox.Show("Error caught: " + ex.Message);
                    }


                    // ---- Condition Summary

                    for (int i = 0; i < 10; i++)
                    {
                        if (formDataArray[i, 0] != null)
                        {
                            gfx.DrawString(formDataArray[i, 0], font, XBrushes.Black, new XRect(57, b + 54 + i * 12.7, 100, 200), XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 20] != null)
                        {
                            gfx.DrawString(formDataArray[i, 20], font, XBrushes.Black, new XRect(125, b + 54 + i * 12.7, 100, 200), XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 2] != null)
                        {
                            gfx.DrawString(formDataArray[i, 2] + " " + formDataArray[i, 3], font, XBrushes.Black, new XRect(303, b + 54 + i * 12.7, 100, 200), XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 32] != null)
                        {
                            gfx.DrawString(formDataArray[i, 32], font, XBrushes.Black, new XRect(523, b + 54 + i * 12.7, 100, 200), XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 33] != null)
                        {
                            if (formDataArray[i, 33] == "PASS")
                            {
                                gfx.DrawString(formDataArray[i, 33], font, XBrushes.Black, new XRect(740, b + 54 + i * 12.7, 100, 200), XStringFormats.TopLeft);
                            }
                            else
                            {
                                gfx.DrawString(formDataArray[i, 33], font, XBrushes.Red, new XRect(740, b + 54 + i * 12.7, 100, 200), XStringFormats.TopLeft);
                            }
                        }
                    }

                    //Print Page 2 ------------------------------------------------------
                    page = doc.Pages[1];
                    contents = ContentReader.ReadContent(page);

                    // Get an XGraphics object for drawing
                    gfx = XGraphics.FromPdfPage(page);

                    // ------- BS declaration 
                    gfx.DrawString("BS7671 / BSEN 62305", font, XBrushes.Black, new XRect(350, 10, 100, 200), XStringFormats.TopRight);

                    // -------- Header
                    // Get an XGraphics object for drawing
                    using (var isolatedStorage = IsolatedStorageFile.GetUserStoreForAssembly())
                    {
                        string imageName = "UserImage.jpg";

                        // Check if the file exists
                        if (isolatedStorage.FileExists(imageName))
                        {
                            // Open the file stream
                            using (var imageStream = isolatedStorage.OpenFile(imageName, FileMode.Open))
                            {
                                // Load the image from the stream
                                XImage image = XImage.FromStream(imageStream);

                                // Draw the image on the PDF page
                                gfx.DrawImage(image, 30, 10, 135, 45);
                            }
                        }
                        else
                        {
                            // File not found in isolated storage
                            System.Windows.MessageBox.Show("Image file not found in isolated storage.");
                        }
                    }
                    if (companyAddressValues[0] != null)
                    {
                        //Address line 1
                        gfx.DrawString(companyAddressValues[0], font, XBrushes.Black, new XRect(710, 10, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[1] != null)
                    {
                        //Address line 2
                        gfx.DrawString(companyAddressValues[1], font, XBrushes.Black, new XRect(710, 19, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[2] != null)
                    {
                        //Address line 3
                        gfx.DrawString(companyAddressValues[2], font, XBrushes.Black, new XRect(710, 28, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[3] != null)
                    {
                        //Address line 4
                        gfx.DrawString(companyAddressValues[3], font, XBrushes.Black, new XRect(710, 37, 100, 200), XStringFormats.TopRight);
                    }
                    if (companyAddressValues[4] != null)
                    {
                        //Telephone number
                        gfx.DrawString(companyAddressValues[4], font, XBrushes.Black, new XRect(710, 46, 100, 200), XStringFormats.TopRight);
                    }

                    //int row = 0;
                    for (int i = 0; i < 10; i++)
                    {
                        //Draw the text
                        if (formDataArray[i, 0] != null) // Source Protected
                        {
                            gfx.DrawString(formDataArray[i, 0], font, XBrushes.Black,
                              new XRect(52, b + 11 + i * 15, c, d),
                              XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 1] != null) // Circuit number
                        {
                            gfx.DrawString(formDataArray[i, 1], font, XBrushes.Black,
                          new XRect(a + 52, b + 11 + i * 15, c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 2] != null) // SPD Make & Model
                        {
                            string printedmm = formDataArray[i, 2];
                            if(formDataArray[i, 3] != null){
                                printedmm = printedmm + " " + formDataArray[i, 3];
                            }
                            gfx.DrawString(printedmm, font, XBrushes.Black,
                          new XRect(a + 95, b + 11 + i * 15, c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 4] != null) // SPD type
                        {
                            gfx.DrawString(formDataArray[i, 4], font, XBrushes.Black,
                          new XRect(a + 310, b + 11 + i * 15, c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 6] != null) // L1-PE
                        {
                            XBrush textColor = XBrushes.Black;

                            // Check if the last two characters of formDataArray[i, 5] are "OR"
                            if (formDataArray[i, 5] != null && formDataArray[i, 5].EndsWith("OR"))
                            {
                                textColor = XBrushes.Red;
                            }

                            gfx.DrawString(formDataArray[i, 6], font, textColor,
                                new XRect(440, b + 11 + i * 15, c, d),
                                XStringFormats.TopLeft);
                            // Print the first 3 characters of formDataArray[i, 5] vertically
                            if (formDataArray[i, 5] != null && formDataArray[i, 5].Length >= 3)
                            {
                                string firstThreeChars = formDataArray[i, 5].Substring(0, 3);
                                if (firstThreeChars == "gdt" || firstThreeChars == "mov")
                                {
                                    gfx.DrawString(firstThreeChars, font, XBrushes.Black,
                                    new XRect(458, b + 11 + i * 15, 10, d),
                                    XStringFormats.TopCenter);
                                }
                            }

                        }

                        if (formDataArray[i, 8] != null) // L2-PE
                        {
                            XBrush textColor = XBrushes.Black;

                            // Check if the last two characters of formDataArray[i, 5] are "OR"
                            if (formDataArray[i, 7] != null && formDataArray[i, 7].EndsWith("OR"))
                            {
                                textColor = XBrushes.Red;
                            }

                            gfx.DrawString(formDataArray[i, 8], font, textColor,
                                new XRect(489, b + 11 + i * 15, c, d),
                                XStringFormats.TopLeft);
                            // Print the first 3 characters of formDataArray[i, 5] vertically
                            if (formDataArray[i, 7] != null && formDataArray[i, 7].Length >= 3)
                            {
                                string firstThreeChars = formDataArray[i, 7].Substring(0, 3);
                                if (firstThreeChars == "gdt" || firstThreeChars == "mov")
                                {
                                    gfx.DrawString(firstThreeChars, font, XBrushes.Black,
                                    new XRect(507, b + 11 + i * 15, 10, d),
                                    XStringFormats.TopCenter);
                                }
                            }
                        }

                        if (formDataArray[i, 10] != null) // L3-PE
                        {
                            XBrush textColor = XBrushes.Black;

                            // Check if the last two characters of formDataArray[i, 5] are "OR"
                            if (formDataArray[i, 9] != null && formDataArray[i, 9].EndsWith("OR"))
                            {
                                textColor = XBrushes.Red;
                            }

                            gfx.DrawString(formDataArray[i, 10], font, textColor,
                                new XRect(538, b + 11 + i * 15, c, d),
                                XStringFormats.TopLeft);
                            // Print the first 3 characters of formDataArray[i, 5] vertically
                            if (formDataArray[i, 9] != null && formDataArray[i, 9].Length >= 3)
                            {
                                string firstThreeChars = formDataArray[i, 9].Substring(0, 3);
                                if (firstThreeChars == "gdt" || firstThreeChars == "mov")
                                {
                                    gfx.DrawString(firstThreeChars, font, XBrushes.Black,
                                    new XRect(556, b + 11 + i * 15, 10, d),
                                    XStringFormats.TopCenter);
                                }
                            }
                        }

                        if (formDataArray[i, 12] != null) // N-PE
                        {
                            XBrush textColor = XBrushes.Black;

                            // Check if the last two characters of formDataArray[i, 5] are "OR"
                            if (formDataArray[i, 11] != null && formDataArray[i, 11].EndsWith("OR"))
                            {
                                textColor = XBrushes.Red;
                            }

                            gfx.DrawString(formDataArray[i, 12], font, textColor,
                                new XRect(587, b + 11 + i * 15, c, d),
                                XStringFormats.TopLeft);
                            // Print the first 3 characters of formDataArray[i, 5] vertically
                            if (formDataArray[i, 11] != null && formDataArray[i, 11].Length >= 3)
                            {
                                string firstThreeChars = formDataArray[i, 11].Substring(0, 3);
                                if (firstThreeChars == "gdt" || firstThreeChars == "mov")
                                {
                                    gfx.DrawString(firstThreeChars, font, XBrushes.Black,
                                    new XRect(605, b + 11 + i * 15, 10, d),
                                    XStringFormats.TopCenter);
                                }
                            }
                        }

                        if (formDataArray[i, 14] != null) // L1-N
                        {
                            XBrush textColor = XBrushes.Black;

                            // Check if the last two characters of formDataArray[i, 5] are "OR"
                            if (formDataArray[i, 13] != null && formDataArray[i, 13].EndsWith("OR"))
                            {
                                textColor = XBrushes.Red;
                            }

                            gfx.DrawString(formDataArray[i, 14], font, textColor,
                                new XRect(636, b + 11 + i * 15, c, d),
                                XStringFormats.TopLeft);
                            // Print the first 3 characters of formDataArray[i, 5] vertically
                            if (formDataArray[i, 13] != null && formDataArray[i, 13].Length >= 3)
                            {
                                string firstThreeChars = formDataArray[i, 13].Substring(0, 3);
                                if (firstThreeChars == "gdt" || firstThreeChars == "mov")
                                {
                                    gfx.DrawString(firstThreeChars, font, XBrushes.Black,
                                    new XRect(654, b + 11 + i * 15, 10, d),
                                    XStringFormats.TopCenter);
                                }
                            }
                        }

                        if (formDataArray[i, 16] != null) // L2-N
                        {
                            XBrush textColor = XBrushes.Black;

                            // Check if the last two characters of formDataArray[i, 5] are "OR"
                            if (formDataArray[i, 15] != null && formDataArray[i, 15].EndsWith("OR"))
                            {
                                textColor = XBrushes.Red;
                            }

                            gfx.DrawString(formDataArray[i, 16], font, textColor,
                                new XRect(685, b + 11 + i * 15, c, d),
                                XStringFormats.TopLeft);
                            // Print the first 3 characters of formDataArray[i, 5] vertically
                            if (formDataArray[i, 15] != null && formDataArray[i, 15].Length >= 3)
                            {
                                string firstThreeChars = formDataArray[i, 15].Substring(0, 3);
                                if (firstThreeChars == "gdt" || firstThreeChars == "mov")
                                {
                                    gfx.DrawString(firstThreeChars, font, XBrushes.Black,
                                    new XRect(703, b + 11 + i * 15, 10, d),
                                    XStringFormats.TopCenter);
                                }
                            }
                        }

                        if (formDataArray[i, 18] != null) // L3-N
                        {
                            XBrush textColor = XBrushes.Black;

                            // Check if the last two characters of formDataArray[i, 5] are "OR"
                            if (formDataArray[i, 17] != null && formDataArray[i, 17].EndsWith("OR"))
                            {
                                textColor = XBrushes.Red;
                            }

                            gfx.DrawString(formDataArray[i, 18], font, textColor,
                                new XRect(734, b + 11 + i * 15, c, d),
                                XStringFormats.TopLeft);
                            // Print the first 3 characters of formDataArray[i, 5] vertically
                            if (formDataArray[i, 17] != null && formDataArray[i, 17].Length >= 3)
                            {
                                string firstThreeChars = formDataArray[i, 17].Substring(0, 3);
                                if (firstThreeChars == "gdt" || firstThreeChars == "mov")
                                {
                                    gfx.DrawString(firstThreeChars, font, XBrushes.Black,
                                    new XRect(752, b + 11 + i * 15, 10, d),
                                    XStringFormats.TopCenter);
                                }
                            }
                        }

                        if (formDataArray[i, 19] != null) //Zs
                        {
                            gfx.DrawString(formDataArray[i, 19], font, XBrushes.Black,
                          new XRect(783, b + 11 + i * 15, c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 20] != null) //Location
                        {
                            gfx.DrawString(formDataArray[i, 20], font, XBrushes.Black,
                          new XRect(52, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 21] != null) //BSEN
                        {
                            gfx.DrawString(formDataArray[i, 21], font, XBrushes.Black,
                          new XRect(318, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 22] != null) //BSEN Type
                        {
                            gfx.DrawString(formDataArray[i, 22], font, XBrushes.Black,
                          new XRect(396, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 23] != null) //BSEN Rating
                        {
                            gfx.DrawString(formDataArray[i, 23], font, XBrushes.Black,
                          new XRect(443, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 24] != null) //BSEN SSC
                        {
                            gfx.DrawString(formDataArray[i, 24], font, XBrushes.Black,
                          new XRect(a + 395, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 25] != null) //BSEN Reference Method
                        {
                            gfx.DrawString(formDataArray[i, 25], font, XBrushes.Black,
                          new XRect(533, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 26] != null) //BSEN Live mm
                        {
                            gfx.DrawString(formDataArray[i, 26], font, XBrushes.Black,
                          new XRect(575, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 27] != null) //BSEN CPC mm
                        {
                            gfx.DrawString(formDataArray[i, 27], font, XBrushes.Black,
                          new XRect(618, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 28] != null) //R2
                        {
                            gfx.DrawString(formDataArray[i, 28], font, XBrushes.Black,
                          new XRect(663, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 29] != null) //IR LL
                        {
                            gfx.DrawString(formDataArray[i, 29], font, XBrushes.Black,
                          new XRect(708, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 30] != null) //IR lPE
                        {
                            gfx.DrawString(formDataArray[i, 30], font, XBrushes.Black,
                          new XRect(751, (b + 270) + (i * 15), c, d),
                          XStringFormats.TopLeft);
                        }

                        if (formDataArray[i, 31] != null) //Polarity
                        {
                            string tc = "";
                            if (formDataArray[i, 31] == "True")
                            {
                                tc = "ü";
                                gfx.DrawString(tc, fontSym, XBrushes.Black, new XRect(793, (b + 270) + (i * 15), c, d), XStringFormats.TopLeft);
                            }
                            else if (formDataArray[i, 31] == "False")
                            {
                                tc = "û";
                                gfx.DrawString(tc, fontSym, XBrushes.Red, new XRect(793, (b + 270) + (i * 15), c, d), XStringFormats.TopLeft);
                            }
                            else
                            {

                            }


                        }

                    }

                    if (formDetails[1] != null) // certificate number in p2 footer
                    {
                        gfx.DrawString(formDetails[1], font, XBrushes.Black,
                      new XRect(a + 50, 574, 100, 200),
                      XStringFormats.TopLeft);
                    }

                    // Save the files
                    if (savePath != null)
                    {
                        string certificateFilePath = savePath + "\\" + formDetails[1] + "_Certificate.pdf";
                        string dataFilePath = savePath + "\\" + formDetails[1] + "_Data.sjh";

                        if (File.Exists(certificateFilePath) || File.Exists(dataFilePath))
                        {
                            // Files already exist, ask for confirmation
                            DialogResult overwriteResult = System.Windows.Forms.MessageBox.Show("The files already exist. Do you want to overwrite them?", "Confirmation", MessageBoxButtons.YesNo);
                            if (overwriteResult == DialogResult.No)
                            {
                                // User chose not to overwrite, exit the function or handle accordingly
                                return;
                            }
                        }

                        doc.Save(certificateFilePath);
                        System.Windows.Forms.MessageBox.Show(formDetails[1] + "_Certificate.pdf\n" + formDetails[1] + "_Data.sjh\n\nSaved to " + savePath, "Saved");
                        Process.Start(certificateFilePath);

                        CertificateData.SaveDataToFile(dataFilePath, imageSignatureData, formDataArray, formDetails);
                    }

                }
            } catch (Exception ex) 
            {
                System.Windows.MessageBox.Show("Error caught: " + ex.Message);
            }
        }

    }

}

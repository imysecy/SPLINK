using System;
using System.Text;
using System.Windows.Forms;

namespace SharePoint_Link.Utility
{
    /// <summary>
    /// <c>HtmlFromRtf</c> class
    /// Implements the functionality to convert  convert a string from rtf format to html format
    /// </summary>
    class HtmlFromRtf
    {

        /// <summary>
        /// <c>IsHTMLinRtf</c> member function 
        /// finds whether the string is in rtf format or not
        /// </summary>
        /// <param name="rtf"></param>
        /// <returns></returns>
        private bool IsHTMLinRtf(string rtf)
        {
            // We look for the words "\fromhtml" somewhere in the file.     
            // If the rtf encodes text rather than html, then instead      
            // it will only find "\fromtext".       
            return rtf.Contains(@"\fromhtml");
        }


        /// <summary>
        /// <c>Match</c> member function
        /// it matches the string within a string at the specified position
        /// </summary>
        /// <param name="str"></param>
        /// <param name="match"></param>
        /// <param name="pos"></param>
        /// <returns></returns>
        private bool Match(string str, string match, ref int pos)
        {
            if (str.Length >= match.Length + pos && str.Substring(pos, match.Length) == match) { pos += match.Length; return true; }
            else
                return false;
        }


        /// <summary> member function
        /// <c>ReadInt</c>
        /// read integer and skip following space if any   
        /// </summary>
        /// <param name="str"></param>
        /// <param name="pos"></param>
        /// <returns></returns>
        private int ReadInt(string str, ref int pos)
        {
            int i = 0;
            while (str[pos] >= '0' && str[pos] <= '9')
            {
                i = i * 10 + (str[pos] - '0');
                ++pos;
            }
            if (str[pos] == ' ')
                ++pos;
            return i;
        }



        /// <summary> 
        /// <c>SkipTo</c> member function
        /// skips to next occurence of char c, returns false if not found 
        /// </summary>
        /// <param name="str"></param>
        /// <param name="c"></param>
        /// <param name="pos"></param>
        /// <returns></returns>
        private bool SkipTo(string str, char c, ref int pos)
        {
            pos = str.IndexOf('}', pos);
            if (pos < 0)
                return false;
            ++pos;
            return true;
        }

        /// <summary>
        /// <c>GetRefinedHtmlFromRtf</c> member function
        /// convert rtf data to html data
        /// </summary>
        /// <param name="rtf"></param>
        /// <returns></returns>
        public string GetRefinedHtmlFromRtf(string rtf)
        {
            string result = string.Empty;

            try
            {
                StringBuilder s = new StringBuilder();
                int len = rtf.Length;
                int pos = rtf.IndexOf(@"{\*\htmltag");
                int ignore_tag = -1;
                while (pos < len)
                {
                    char c = rtf[pos];
                    if (c == '{') ++pos;
                    else if (c == '}') ++pos;
                    else if (c == '\r' || c == '\n')
                        ++pos;
                    else if (c == '\\')
                    {
                        ++pos;
                        if (rtf[pos] == '{')
                        {
                            s.Append('{');
                            ++pos;
                        }
                        else if (rtf[pos] == '}')
                        { s.Append('}'); ++pos; }
                        else if (rtf[pos] == '\'')
                        {
                            ++pos;
                            int hex = Convert.ToInt32(rtf.Substring(pos, 2), 16);
                            s.Append((char)hex); pos += 2;
                        }
                        else if (Match(rtf, @"*\htmltag", ref pos))
                        {
                            int tag = ReadInt(rtf, ref pos);
                            if (tag == ignore_tag)
                                if (!SkipTo(rtf, '}', ref pos))
                                    break; // abort if '}' not found    
                        }
                        else if (Match(rtf, @"*\mhtmltag", ref pos))
                        {
                            ignore_tag = ReadInt(rtf, ref pos);
                        }
                        else if (Match(rtf, "par", ref pos))
                        {
                            s.Append("\r\n");
                        }
                        else if (Match(rtf, "tab", ref pos))
                        {
                            s.Append("\t");
                        }
                        else if (Match(rtf, "li", ref pos))
                        {
                            ReadInt(rtf, ref pos);
                        }
                        else if (Match(rtf, "fi-", ref pos))
                        {
                            ReadInt(rtf, ref pos);
                        }
                        else if (Match(rtf, "pntext", ref pos))
                        {
                            SkipTo(rtf, '}', ref pos);
                        }
                        else if (Match(rtf, "htmlrtf", ref pos))
                        {
                            pos = rtf.IndexOf(@"\htmlrtf0", pos);
                            if (pos < 0)
                                break;
                            pos += 9;
                        }
                        else
                        {
                            s.Append('\\');
                            s.Append(c); ++pos;
                        }
                    }
                    else
                    {
                        s.Append(c);
                        ++pos;
                    }
                }
                if (pos < 0)
                    return null;
                //  return s.ToString();
                result = s.ToString();
                try
                {


                    int indexofalink = result.IndexOf("a:link");
                    while (indexofalink != -1)
                    {
                        if (indexofalink != -1)
                        {

                            int indclosed = result.IndexOf("}", indexofalink);
                            if (indclosed != -1)
                            {
                                result = result.Remove(indexofalink, (indclosed - indexofalink));
                            }

                        }
                        indexofalink = result.IndexOf("a:link");
                    }


                    indexofalink = result.IndexOf("a:visited");

                    while (indexofalink != -1)
                    {
                        if (indexofalink != -1)
                        {
                            int indclose = result.IndexOf("}", indexofalink);
                            if (indclose != -1)
                            {
                                result = result.Remove(indexofalink, (indclose - indexofalink));
                            }

                        }

                        indexofalink = result.IndexOf("a:visited");
                    }



                }
                catch (Exception)
                {

                    result = s.ToString();
                }


            }
            catch (Exception)
            {


            }
            if (string.IsNullOrEmpty(result))
            {
                result = getsecondalternatehtml(rtf, result);
            }
            return result;
        }


        /// <summary>
        /// <c>getsecondalternatehtml</c> member function
        /// replaces rtf tags with relevant html tags
        /// </summary>
        /// <param name="rtf"></param>
        /// <param name="strhtml"></param>
        /// <returns></returns>
        public string getsecondalternatehtml(string rtf, string strhtml)
        {

            try
            {
                if (string.IsNullOrEmpty(strhtml))
                {
                    strhtml = rtf;



                    try
                    {
                        int i = 1;
                        while (i > 0)
                        {
                            i = strhtml.IndexOf(@"{\*\htmltag84 <img");
                            if (i > -1)
                            {
                                strhtml = strhtml.Remove(i, 13);

                            }

                        }

                        strhtml = GetHtml(strhtml);

                        int start = 0;
                        while (start != -1)
                        {

                            if (start < strhtml.Length)
                            {


                                int wwindex = strhtml.ToLower().IndexOf("www.", start);

                                if (wwindex != -1)
                                {


                                    char[] characters = { '\n', ' ', '\r' };
                                    int indspace = strhtml.IndexOfAny(characters, wwindex - 1);
                                    if (indspace != -1)
                                    {
                                        string subst = strhtml.Substring(wwindex, indspace - wwindex);
                                        strhtml = strhtml.Replace(subst, "<a href=\"http://" + subst + "\">" + subst + "</a>");
                                        start = wwindex + (2 * subst.Length) + 10;
                                    }
                                    else
                                    {
                                        start = wwindex + 10;
                                    }

                                }
                                else
                                {
                                    start = -1;
                                }
                            }
                        }
                        ////////////////////////
                        start = 0;
                        while (start != -1)
                        {

                            if (start < strhtml.Length)
                            {


                                int wwindex = strhtml.ToLower().IndexOf("mailto:", start);

                                if (wwindex != -1)
                                {


                                    char[] characters = { '\n', ' ', '\r', ']' };
                                    int indspace = strhtml.IndexOfAny(characters, wwindex - 1);
                                    if (indspace != -1)
                                    {


                                        string subst = strhtml.Substring(wwindex, indspace - wwindex);
                                        strhtml = strhtml.Replace(subst, "<a href=\"" + subst + "\">" + subst + "</a>");
                                        start = indspace + 2 * subst.Length + 10;
                                    }
                                    else
                                    {
                                        start = indspace + 2 + 10;
                                    }
                                }
                                else
                                {
                                    start = -1;
                                }
                            }
                        }
                        ////////////////////////



                    }
                    catch (Exception ex)
                    {

                    }

                }

                /////////////End of ExceptionCase ////////////
            }
            catch (Exception)
            { }
            return strhtml;
        }

        /// <summary>
        /// <c>GetHtml</c> member function
        /// calls <c>converttoHtml</c> member function to convert rtf to html using richtextbox
        /// </summary>
        /// <param name="strRTF"></param>
        /// <returns></returns>
        public static string GetHtml(string strRTF)
        {
            string result = string.Empty;
            try
            {
                RichTextBox rtbox = new RichTextBox();

                rtbox.SelectedRtf = strRTF;


                result = converttoHtml(rtbox);


            }
            catch (Exception)
            {


            }
            return result;

        }

        /// <summary>
        /// <c>converttoHtml</c> member function 
        /// converts rtf from richtextbox to html
        /// </summary>
        /// <param name="rtb"></param>
        /// <returns></returns>
        public static string converttoHtml(RichTextBox rtb)
        {

            RichTextBox Box = rtb;
            string strHTML = string.Empty;
            string strColour = string.Empty;
            bool blnBold;
            bool blnItalic;

            string strFont;
            short shtSize;

            long lngOriginalStart;
            long lngOriginalLength;

            // int intCount;
            if (Box.Text.Length == 0)
            {

            }


            // Store original selections, then select first character
            lngOriginalStart = 0;
            lngOriginalLength = Box.TextLength;
            Box.Select(0, 1);
            //       ' Add HTML header
            strHTML = "<html>";
            //    ' Set up initial parameters
            strColour = Box.SelectionColor.ToKnownColor().ToString();
            blnBold = Box.SelectionFont.Bold;
            blnItalic = Box.SelectionFont.Italic;
            strFont = Box.SelectionFont.FontFamily.Name;
            shtSize = short.Parse(Box.SelectionFont.Size.ToString());

            //      ' Include first 'style' parameters in the HTML
            strHTML += "<span style=\"font-family: " + strFont + "; font-size: " + shtSize + "pt; color: " + strColour + "\">";
            //        ' Include bold tag, if required
            if (blnBold == true)
            {
                strHTML += "<b>";
            }

            //      ' Include italic tag, if required

            if (blnItalic == true)
            {
                strHTML += "<i>";
            }

            //  ' Finally, add our first character
            strHTML += Box.Text.Substring(0, 1);
            // ' Loop around all remaining characters
            for (int intCount = 2; intCount < Box.Text.Length; intCount++)
            {
                Box.Select(intCount - 1, 1);
                if (Convert.ToChar(Box.Text.Substring((intCount - 1), 1)) == Convert.ToChar(10))
                {
                    strHTML += "<br>";
                }

                //    ' Check/implement any changes in style
                if (Box.SelectionColor.ToKnownColor().ToString() != strColour || Box.SelectionFont.Size != shtSize)
                {
                    strHTML += "</span><span style=\"font-family: " + Box.SelectionFont.FontFamily.Name + "; font-size: " + Box.SelectionFont.Size + "pt; color: " + Box.SelectionColor.ToKnownColor().ToString() + "\">";

                }

                //' Check for bold changes
                if (Box.SelectionFont.Bold != blnBold)
                {
                    if (Box.SelectionFont.Bold == false)
                    {
                        strHTML += "</b>";
                    }
                    else
                    {
                        strHTML += "<b>";
                    }
                }
                if (Box.SelectionFont.Italic != blnItalic)
                {
                    if (Box.SelectionFont.Italic == false)
                    {
                        strHTML += "</i>";
                    }
                    else
                    {
                        strHTML += "<i>";
                    }
                }
                //       ' Add the actual character
                strHTML += Box.Text.Substring(intCount, 1);    //Mid(Box.Text, intCount, 1);
                //   ' Update variables with current style
                strColour = Box.SelectionColor.ToKnownColor().ToString();
                blnBold = Box.SelectionFont.Bold;
                blnItalic = Box.SelectionFont.Italic;
                strFont = Box.SelectionFont.FontFamily.Name;
                shtSize = short.Parse(Box.SelectionFont.Size.ToString());

            }







            //  Close off any open bold/italic tags
            if (blnBold == true)
            {
                strHTML += "";
            }
            if (blnItalic == true)
            {
                strHTML += "";
            }

            // Terminate outstanding HTML tags
            strHTML += "</span></html>";
            // Restore original RichTextBox selection
            Box.Select(Int32.Parse(lngOriginalStart.ToString()), Int32.Parse(lngOriginalLength.ToString()));

            return strHTML;
        }


    }
}

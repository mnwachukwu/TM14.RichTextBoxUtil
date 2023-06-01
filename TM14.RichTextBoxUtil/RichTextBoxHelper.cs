using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TM14.RichTextBoxUtil
{
    public static class RichTextBoxHelper
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);
        public const int WmVscroll = 277; // Vertical scroll
        public const int SbBottom = 7; // Scroll to bottom

        public static void AppendText(this RichTextBox richTextBox, string text, Color color, float size, bool showTimeStamp = false)
        {
            var italicFormattingActive = false;
            var boldFormattingActive = false;
            var italicBoldFormattingActive = false;
            var strikeThroughFormattingActive = false;
            var underlineFormattingActive = false;

            richTextBox.SelectionStart = richTextBox.TextLength;
            richTextBox.SelectionLength = 0;
            richTextBox.SelectionColor = color;
            richTextBox.SelectionFont = new(richTextBox.Font.FontFamily, size);

            // Just append a new line and move on
            if (text == null)
            {
                richTextBox.AppendText(Environment.NewLine);
                richTextBox.ScrollRichTextBox();
                return;
            }

            // parse the message for formatting
            var splitText = text.Trim().Split(' ');

            // Append dates
            if (showTimeStamp)
            {
                richTextBox.AppendText($"[{DateTime.Now:h:mm:ss}] ");
            }

            foreach (var s in splitText)
            {
                if (s.Length <= 0)
                {
                    continue;
                }

                // are markdown characters the only thing in this string? If so, append text according to currently active formatting (without applying new formatting)
                if (s.ContainsOnlyAsterisks() || s.ContainsOnlyTildes() || s.ContainsOnlyUnderscores())
                {
                    richTextBox.AppendText(s + " ");
                    continue;
                }

                if (s.Contains("__"))
                {
                    var ss = s.Split(new[] { "__" }, StringSplitOptions.None);
                    underlineFormattingActive = !underlineFormattingActive;

                    if (ss.Length > 2)
                    {
                        for (var i = 0; i < ss.Length; i++)
                        {
                            if (i == 1)
                            {
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            if (i == ss.Length - 1)
                            {
                                underlineFormattingActive = !underlineFormattingActive;
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            richTextBox.AppendText(ss[i]);
                        }
                    }
                    else
                    {
                        richTextBox.AppendText(ss[0]);
                        richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                        richTextBox.AppendText(ss[1]);
                    }

                    richTextBox.AppendText(" ");
                    continue;
                }

                if (s.Contains("***"))
                {
                    var ss = s.Split(new[] { "***" }, StringSplitOptions.None);
                    italicBoldFormattingActive = !italicBoldFormattingActive;

                    if (ss.Length > 2)
                    {
                        for (var i = 0; i < ss.Length; i++)
                        {
                            if (i == 1)
                            {
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            if (i == ss.Length - 1)
                            {
                                italicBoldFormattingActive = !italicBoldFormattingActive;
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            richTextBox.AppendText(ss[i]);
                        }
                    }
                    else
                    {
                        richTextBox.AppendText(ss[0]);
                        richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                        richTextBox.AppendText(ss[1]);
                    }
                    
                    richTextBox.AppendText(" ");
                    continue;
                }

                if (s.Contains("**"))
                {
                    var ss = s.Split(new[] { "**" }, StringSplitOptions.None);
                    boldFormattingActive = !boldFormattingActive;

                    if (ss.Length > 2)
                    {
                        for (var i = 0; i < ss.Length; i++)
                        {
                            if (i == 1)
                            {
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            if (i == ss.Length - 1)
                            {
                                boldFormattingActive = !boldFormattingActive;
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            richTextBox.AppendText(ss[i]);
                        }
                    }
                    else
                    {
                        richTextBox.AppendText(ss[0]);
                        richTextBox.SelectionFont = new(richTextBox.Font,
                            DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive,
                                strikeThroughFormattingActive, underlineFormattingActive));
                        richTextBox.AppendText(ss[1]);
                    }

                    richTextBox.AppendText(" ");
                    continue;
                }

                if (s.Contains("*"))
                {
                    var ss = s.Split(new[] { "*" }, StringSplitOptions.None);
                    italicFormattingActive = !italicFormattingActive;

                    if (ss.Length > 2)
                    {
                        for (var i = 0; i < ss.Length; i++)
                        {
                            if (i == 1)
                            {
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            if (i == ss.Length - 1)
                            {
                                italicFormattingActive = !italicFormattingActive;
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            richTextBox.AppendText(ss[i]);
                        }
                    }
                    else
                    {
                        richTextBox.AppendText(ss[0]);
                        richTextBox.SelectionFont = new(richTextBox.Font,
                            DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive,
                                strikeThroughFormattingActive, underlineFormattingActive));
                        richTextBox.AppendText(ss[1]);
                    }

                    richTextBox.AppendText(" ");
                    continue;
                }

                if (s.Contains("~~"))
                {
                    var ss = s.Split(new[] { "~~" }, StringSplitOptions.None);
                    strikeThroughFormattingActive = !strikeThroughFormattingActive;

                    if (ss.Length > 2)
                    {
                        for (var i = 0; i < ss.Length; i++)
                        {
                            if (i == 1)
                            {
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            if (i == ss.Length - 1)
                            {
                                strikeThroughFormattingActive = !strikeThroughFormattingActive;
                                richTextBox.SelectionFont = new(richTextBox.Font, DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive, strikeThroughFormattingActive, underlineFormattingActive));
                            }

                            richTextBox.AppendText(ss[i]);
                        }
                    }
                    else
                    {
                        richTextBox.AppendText(ss[0]);
                        richTextBox.SelectionFont = new(richTextBox.Font,
                            DetermineFontStyle(italicFormattingActive, boldFormattingActive, italicBoldFormattingActive,
                                strikeThroughFormattingActive, underlineFormattingActive));
                        richTextBox.AppendText(ss[1]);
                    }

                    richTextBox.AppendText(" ");
                    continue;
                }

                // doesn't start with a formatting character, uses previously applied formatting
                richTextBox.AppendText(s + " ");
            }

            richTextBox.AppendText(Environment.NewLine);
            richTextBox.ScrollRichTextBox();
        }

        private static FontStyle DetermineFontStyle(bool italicFormattingActive, bool boldFormattingActive, bool italicBoldFormattingActive, bool strikeThroughFormattingActive, bool underlineFormattingActive)
        {
            var style = FontStyle.Regular;

            if (italicBoldFormattingActive)
            {
                style |= FontStyle.Bold | FontStyle.Italic;
            }
            else
            {
                if (boldFormattingActive)
                {
                    style |= FontStyle.Bold;
                }

                if (italicFormattingActive)
                {
                    style |= FontStyle.Italic;
                }
            }

            if (underlineFormattingActive)
            {
                style |= FontStyle.Underline;
            }

            if (strikeThroughFormattingActive)
            {
                style |= FontStyle.Strikeout;
            }

            return style;
        }

        private static void ScrollRichTextBox(this RichTextBox richTextBox)
        {
            if (richTextBox == null || richTextBox.IsDisposed || richTextBox.Disposing)
            {
                return;
            }

            SendMessage(richTextBox.Handle, WmVscroll, (IntPtr)SbBottom, IntPtr.Zero);
        }

        public static bool ContainsOnlyAsterisks(this string text)
        {
            return text.All(t => t == '*');
        }

        public static bool ContainsOnlyTildes(this string text)
        {
            return text.All(t => t == '~');
        }

        public static bool ContainsOnlyUnderscores(this string text)
        {
            return text.All(t => t == '_');
        }
    }
}

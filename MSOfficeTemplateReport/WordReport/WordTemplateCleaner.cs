using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Text;
using System.Xml.Linq;

namespace MSOfficeTemplateReport.WordReport
{
    internal sealed class WordTemplateCleaner
    {
        private enum RunState
        {
            None,
            Starting,
            Started,
            Continuing,
        }

        private RunState state;

        private Text firstText;

        private string lastText;

        private readonly StringBuilder texts = new StringBuilder();

        public void Clean(OpenXmlElement item, Paragraph p)
        {
            if (!(item is Run r))
            {
                return;
            }

            var text = r.GetFirstChild<Text>();

            if (text == null)
            {
                return;
            }

            var str = text.Text;

            switch (state)
            {
                case RunState.None:

                    TryStart(text);

                    break;

                case RunState.Starting:

                    state = str.StartsWith("{") ? RunState.Started : RunState.None;

                    break;

                case RunState.Continuing:

                    str = lastText + str;

                    lastText = null;

                    state = RunState.Started;

                    break;
            }

            if (state != RunState.None)
            {
                texts.Append(text.Text);

                if (!ReferenceEquals(firstText, text))
                {
                    p.RemoveChild(r);
                }
            }
            
            if (state == RunState.Started)
            {
                TryEnd(str);
            }
        }

        private void TryStart(Text text)
        {
            state = text.Text.EndsWith("{{")
                ? RunState.Started
                : text.Text.EndsWith("{")
                ? RunState.Starting
                : text.Text.Contains("{{")
                ? RunState.Started
                : RunState.None;

            if (state != RunState.None)
            {
                texts.Clear();

                firstText = text;
            }
        }

        private void TryEnd(string str)
        {
            if (!str.EndsWith("{{") && str.EndsWith("{"))
            {
                lastText = str;

                state = RunState.Continuing;

                return;
            }

            var i = str.LastIndexOf("}}", StringComparison.Ordinal);

            if (i >= 0)
            {
                var j = str.LastIndexOf("{{", StringComparison.Ordinal);

                if (j <= i)
                {
                    firstText.Text = texts.ToString();
                   
                    firstText.SetAttribute(new OpenXmlAttribute("space", XNamespace.Xml.NamespaceName, "preserve"));

                    state = RunState.None;

                    firstText = null;
                }
            }
            else if (str.EndsWith("}"))
            {
                lastText = str;

                state = RunState.Continuing;
            }
        }
    }
}

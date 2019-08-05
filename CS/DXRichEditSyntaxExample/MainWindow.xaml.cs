using DevExpress.Office.Utils;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Services;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace DXRichEditSyntaxExample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            richEditControl1.ReplaceService<ISyntaxHighlightService>(new CustomSyntaxHighlightService(richEditControl1.Document));
            richEditControl1.LoadDocument("CarsXtraScheduling.sql");
            richEditControl1.Document.Sections[0].Page.Width = Units.InchesToDocumentsF(80f);
            richEditControl1.Document.DefaultCharacterProperties.FontName = "Courier New";
        }
    }
    public class CustomSyntaxHighlightService : ISyntaxHighlightService
    {
        readonly Document document;

        Regex _keywords;
        //Declare a regular expression to search text in quotes (including embedded quotes)
        Regex _quotedString = new Regex(@"'([^']|'')*'");

        //Declare a regular expression to search commented text (including multiline)
        Regex _commentedString = new Regex(@"(/\*([^*]|[\r\n]|(\*+([^*/]|[\r\n])))*\*+/)");

        public CustomSyntaxHighlightService(Document document)
        {
            this.document = document;

            //Declare keywords
            string[] keywords = { "INSERT", "SELECT", "CREATE", "TABLE", "USE", "IDENTITY", "ON", "OFF", "NOT", "NULL", "WITH", "SET", "GO", "DECLARE", "EXECUTE", "NVARCHAR", "FROM", "INTO", "VALUES", "WHERE", "AND" };
            this._keywords = new Regex(@"\b(" + string.Join("|", keywords.Select(w => Regex.Escape(w))) + @")\b");
        }

        public void Execute()
        {
            List<SyntaxHighlightToken> tSqltokens = ParseTokens();
            document.ApplySyntaxHighlight(tSqltokens);
        }
        public void ForceExecute()
        {
            Execute();
        }

        private List<SyntaxHighlightToken> ParseTokens()
        {
            List<SyntaxHighlightToken> tokens = new List<SyntaxHighlightToken>();
            DocumentRange[] ranges = null;

            // Search for quoted strings
            ranges = document.FindAll(_quotedString);
            for (int i = 0; i < ranges.Length; i++)
            {
                tokens.Add(CreateToken(ranges[i].Start.ToInt(), ranges[i].End.ToInt(), Color.Red));
            }

            //Extract all keywords
            ranges = document.FindAll(_keywords);
            for (int j = 0; j < ranges.Length; j++)
            {
                //Check whether tokens intersect
                if (!IsRangeInTokens(ranges[j], tokens))
                    tokens.Add(CreateToken(ranges[j].Start.ToInt(), ranges[j].End.ToInt(), Color.Blue));
            }

            //Find all comments
            ranges = document.FindAll(_commentedString);
            for (int j = 0; j < ranges.Length; j++)
            {
                //Check whether tokens intersect
                if (!IsRangeInTokens(ranges[j], tokens))
                    tokens.Add(CreateToken(ranges[j].Start.ToInt(), ranges[j].End.ToInt(), Color.Green));
            }

            // Sort tokens by their start position
            tokens.Sort(new SyntaxHighlightTokenComparer());

            // Fill in gaps in document coverage
            tokens = CombineWithPlainTextTokens(tokens);
            return tokens;
        }

        //Parse the remaining text into tokens:
        List<SyntaxHighlightToken> CombineWithPlainTextTokens(List<SyntaxHighlightToken> tokens)
        {
            List<SyntaxHighlightToken> result = new List<SyntaxHighlightToken>(tokens.Count * 2 + 1);
            int documentStart = this.document.Range.Start.ToInt();
            int documentEnd = this.document.Range.End.ToInt();
            if (tokens.Count == 0)
                result.Add(CreateToken(documentStart, documentEnd, Color.Black));
            else
            {
                SyntaxHighlightToken firstToken = tokens[0];
                if (documentStart < firstToken.Start)
                    result.Add(CreateToken(documentStart, firstToken.Start, Color.Black));
                result.Add(firstToken);
                for (int i = 1; i < tokens.Count; i++)
                {
                    SyntaxHighlightToken token = tokens[i];
                    SyntaxHighlightToken prevToken = tokens[i - 1];
                    if (prevToken.End != token.Start)
                        result.Add(CreateToken(prevToken.End, token.Start, Color.Black));
                    result.Add(token);
                }
                SyntaxHighlightToken lastToken = tokens[tokens.Count - 1];
                if (documentEnd > lastToken.End)
                    result.Add(CreateToken(lastToken.End, documentEnd, Color.Black));
            }
            return result;
        }

        //Check whether tokens intersect
        private bool IsRangeInTokens(DocumentRange range, List<SyntaxHighlightToken> tokens)
        {
            return tokens.Any(t => IsIntersect(range, t));
        }
        bool IsIntersect(DocumentRange range, SyntaxHighlightToken token)
        {
            int start = range.Start.ToInt();
            if (start >= token.Start && start < token.End)
                return true;
            int end = range.End.ToInt() - 1;
            if (end >= token.Start && end < token.End)
                return true;
            if (start < token.Start && end >= token.End)
                return true;
            return false;
        }

        SyntaxHighlightToken CreateToken(int start, int end, Color foreColor)
        {
            SyntaxHighlightProperties properties = new SyntaxHighlightProperties();
            properties.ForeColor = foreColor;
            return new SyntaxHighlightToken(start, end - start, properties);
        }
        public class SyntaxHighlightTokenComparer : IComparer<SyntaxHighlightToken>
        {
            public int Compare(SyntaxHighlightToken x, SyntaxHighlightToken y)
            {
                return x.Start - y.Start;
            }
        }
    }
}

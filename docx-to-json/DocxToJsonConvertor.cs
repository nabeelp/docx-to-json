using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml.Linq;

namespace docx_to_json
{
    class DocxToJsonConvertor
    {
        public string TraceID { get; set; }
        public ILogger Logger { get; set; }
        public bool IncludeCellHtml { get; set; }
        public string[] StringsToRemove { get; set; }

        public DocxToJsonConvertor(string traceID, ILogger logger, bool includeCellHtml = false)
        {
            TraceID = traceID;
            Logger = logger;
            IncludeCellHtml = includeCellHtml;
        }

        public JObject ConvertDocx(Uri blobUri, bool useManagedIdentity = true)
        {
            throw new NotImplementedException();
        }

        public JObject ConvertDocx(Stream inputDoc)
        {
            var jsonResult = new JObject();
            try
            {
                // convert Word document to HTML, as we are using HTML as the "standardised" format for all further processing
                var docxAsHtml = ConvertDocxToHtml(inputDoc);

                // clean HTML, removing unneccessary Word-specific HTML syntax elements
                if (StringsToRemove != null)
                {
                    foreach (var stringToRemove in StringsToRemove)
                    {
                        docxAsHtml = docxAsHtml.Replace(stringToRemove, "");
                    }
                }
                docxAsHtml = Regex.Replace(docxAsHtml, " {2,}", "");

                // process HTML
                ConvertHtmlToJson(docxAsHtml, ref jsonResult);

                // return JSON result
                return jsonResult;
            }
            catch (Exception ex)
            {
                Logger.LogError($"{TraceID} - {ex.Message}");
                throw new Exception($"{TraceID} - {ex.Message}", ex);
            }
        }

        private string ConvertDocxToHtml(Stream inputDoc)
        {
            // convert Stream to a memory stream
            using (var memStream = new MemoryStream())
            {
                inputDoc.CopyTo(memStream);

                // open Word document stream
                using (WordprocessingDocument doc =
                    WordprocessingDocument.Open(memStream, true))
                {
                    // remove unnecessary markup
                    SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                    {
                        AcceptRevisions = true,
                        NormalizeXml = true,
                        RemoveBookmarks = true,
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveEndAndFootNotes = true,
                        RemoveFieldCodes = true,
                        RemoveGoBackBookmark = true,
                        RemoveHyperlinks = false,
                        RemoveLastRenderedPageBreak = true,
                        RemoveMarkupForDocumentComparison = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveRsidInfo = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        RemoveWebHidden = true,
                        ReplaceTabsWithSpaces = true
                    };
                    MarkupSimplifier.SimplifyMarkup(doc, settings);

                    // export to html
                    return WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings()).ToString();
                }
            }
        }
        private void ConvertHtmlToJson(string docxAsHtml, ref JObject jsonResult)
        {
            // load the HTML document
            var htmlDoc = new HtmlDocument();
            htmlDoc.OptionOutputAsXml = true;
            htmlDoc.LoadHtml(docxAsHtml);

            // loop over each table in the html
            var tableList = htmlDoc.DocumentNode.SelectNodes("//table");
            var tableArray = new JArray();
            foreach (var table in tableList)
            {
                // prepare JSON object for this table
                var jsonTable = new JObject();

                // loop over each row in the current table
                var rowList = table.SelectNodes("./tr");
                var rowArray = new JArray();
                foreach (var curRow in rowList)
                {
                    // prepare JSON object for this row
                    var jsonTableRow = new JObject();

                    // loop over each cell in the current row in the current table
                    var cellList = curRow.SelectNodes("./td");
                    var cellArray = new JArray();
                    foreach (var curCell in cellList)
                    {
                        // prepare JSON object for this cell
                        var jsonTableRowCell = new JObject();

                        // include the full html of the cell - useful for debugging
                        if (IncludeCellHtml)
                        {
                            jsonTableRowCell.Add("html", curCell.InnerHtml);
                        }

                        // get cell content, based on each span in the output representing one line
                        var spanList = curCell.SelectNodes(".//span");
                        var lineArray = new JArray();
                        var lineText = String.Empty;
                        foreach (var curSpan in spanList)
                        {
                            // cater for possibility of multuple spans in one "line"
                            lineText += HttpUtility.HtmlDecode(curSpan.InnerHtml);
                            if (curSpan.NextSibling == null)
                            {
                                // prepare JSON object for this span / line
                                var jsonLine = new JObject();
                                jsonLine.Add("type", curSpan.ParentNode.Name);
                                jsonLine.Add("text", lineText);
                                lineArray.Add(jsonLine);

                                // reset lineText, in preparation for next line
                                lineText = String.Empty;
                            }
                        }

                        // add cell to cell Array
                        jsonTableRowCell.Add("lines", lineArray);
                        cellArray.Add(jsonTableRowCell);
                    }

                    // add table row to row array
                    jsonTableRow.Add("cells", cellArray);
                    rowArray.Add(jsonTableRow);
                }

                // add table to table array
                jsonTable.Add("rows", rowArray);
                tableArray.Add(jsonTable);
            }

            // add tables to json Result
            jsonResult.Add("tables", tableArray);
        }
        private static string GetNodeContent(HtmlNode htmlNode)
        {
            // if this node contains lists, just return the whole HTML
            if (htmlNode.InnerHtml.Contains("ListParagraph"))
            {
                return htmlNode.InnerHtml;
            }

            // if this node only has a single #text child node, return that
            if (htmlNode.ChildNodes.Count == 1 && htmlNode.ChildNodes[0].Name == "#text")
            {
                return htmlNode.InnerHtml;
            }

            if (htmlNode.HasChildNodes)
            {
                var childNodes = htmlNode.SelectNodes("./*");
                var nodeString = String.Empty;
                foreach (var childNode in childNodes)
                {
                    nodeString += GetNodeContent(childNode);
                }
                return nodeString;
            }
            else
            {
                return htmlNode.InnerHtml;
            }
        }

    }
}

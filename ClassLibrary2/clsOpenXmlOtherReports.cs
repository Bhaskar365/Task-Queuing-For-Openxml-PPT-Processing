using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

using Vt = DocumentFormat.OpenXml.VariantTypes;


using DocumentFormat.OpenXml.Drawing;
using GraphicFrame = DocumentFormat.OpenXml.Presentation.GraphicFrame;




namespace GeneratedCode
{
    public class clsPrefRankingGenCharts
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private PresentationDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = PresentationDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            //Changes the relationship ID of the parts.
            ReconfigureRelationshipID();
            //Adds new parts or new relationships.
            AddParts();
            //Changes the contents of the specified parts.
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeThumbnailPart1(((ThumbnailPart)UriPartDictionary["/docProps/thumbnail.jpeg"]));
            ChangeExtendedFilePropertiesPart1(((ExtendedFilePropertiesPart)UriPartDictionary["/docProps/app.xml"]));
            ChangeViewPropertiesPart1(((ViewPropertiesPart)UriPartDictionary["/ppt/viewProps.xml"]));
            //ChangeSlidePart1(((SlidePart)UriPartDictionary["/ppt/slides/slide1.xml"]));
        }

        /// <summary>
        /// Stores the references to all the parts in the package.
        /// They could be retrieved by their URIs later.
        /// </summary>
        private void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(document);
            while (queue.Count > 0)
            {
                foreach (var part in queue.Dequeue().Parts)
                {
                    if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                    {
                        UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                        queue.Enqueue(part.OpenXmlPart);
                    }
                }
            }
        }

        /// <summary>
        /// Changes the relationship ID of the parts in the source package to make sure these IDs are the same as those in the target package.
        /// To avoid the conflict of the relationship ID, a temporary ID is assigned first.        
        /// </summary>
        private void ReconfigureRelationshipID()
        {
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/authors.xml"], "generatedTmpID1");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/authors.xml"], "rId11");
        }

        /// <summary>
        /// Adds new parts or new relationship between parts.
        /// </summary>
        private void AddParts()
        {
            //Generate new parts.
            ExtendedPart extendedPart1 = document.PresentationPart.AddExtendedPart("http://schemas.microsoft.com/office/2015/10/relationships/revisionInfo", "application/vnd.ms-powerpoint.revisioninfo+xml", "xml", "rId10");
            GenerateExtendedPart1Content(extendedPart1);

        }

        private void GenerateExtendedPart1Content(ExtendedPart extendedPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(extendedPart1Data);
            extendedPart1.FeedData(data);
            data.Close();
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Revision = "1097";
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2024-11-12T17:22:56Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeThumbnailPart1(ThumbnailPart thumbnailPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        private void ChangeExtendedFilePropertiesPart1(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = extendedFilePropertiesPart1.Properties;

            Ap.TotalTime totalTime1 = properties1.GetFirstChild<Ap.TotalTime>();
            Ap.Words words1 = properties1.GetFirstChild<Ap.Words>();
            Ap.Paragraphs paragraphs1 = properties1.GetFirstChild<Ap.Paragraphs>();
            totalTime1.Text = "109190";

            words1.Text = "25";

            paragraphs1.Text = "17";

        }

        private void ChangeViewPropertiesPart1(ViewPropertiesPart viewPropertiesPart1)
        {
            ViewProperties viewProperties1 = viewPropertiesPart1.ViewProperties;

            SlideViewProperties slideViewProperties1 = viewProperties1.GetFirstChild<SlideViewProperties>();

            CommonSlideViewProperties commonSlideViewProperties1 = slideViewProperties1.GetFirstChild<CommonSlideViewProperties>();

            CommonViewProperties commonViewProperties1 = commonSlideViewProperties1.GetFirstChild<CommonViewProperties>();

            ScaleFactor scaleFactor1 = commonViewProperties1.GetFirstChild<ScaleFactor>();
            Origin origin1 = commonViewProperties1.GetFirstChild<Origin>();

            A.ScaleX scaleX1 = scaleFactor1.GetFirstChild<A.ScaleX>();
            A.ScaleY scaleY1 = scaleFactor1.GetFirstChild<A.ScaleY>();
            scaleX1.Numerator = 50;
            scaleY1.Numerator = 50;
            origin1.X = 56L;
            origin1.Y = 52L;
        }

        public void ChangeSlidePart1(SlidePart slidePart1,string testname, string firstRank,string secondRank,string thirdRank,string weightedScores )
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.GetFirstChild<GraphicFrame>();

            A.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<A.NonVisualGraphicFrameProperties>();
            Transform transform1 = graphicFrame1.GetFirstChild<Transform>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
            modificationId1.Val = (UInt32Value)920664006U;

            A.Extents extents1 = transform1.GetFirstChild<A.Extents>();
            extents1.Cy = 1009620L;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            A.TableRow tableRow1 = new A.TableRow() { Height = 381000L };

            A.TableCell tableCell1 = new A.TableCell();

            A.TextBody textBody1 = new A.TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", FontSize = 800, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill1.Append(rgbColorModelHex1);
            A.EffectList effectList1 = new A.EffectList();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties1.Append(solidFill1);
            runProperties1.Append(effectList1);
            runProperties1.Append(latinFont1);
            A.Text text1 = new A.Text();
            text1.Text = testname;

            run1.Append(runProperties1);
            run1.Append(text1);

            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 800, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill2.Append(rgbColorModelHex2);
            A.EffectList effectList2 = new A.EffectList();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };

            endParagraphRunProperties1.Append(solidFill2);
            endParagraphRunProperties1.Append(effectList2);
            endParagraphRunProperties1.Append(latinFont2);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            A.TableCellProperties tableCellProperties1 = new A.TableCellProperties() { LeftMargin = 3721, RightMargin = 3721, TopMargin = 3721, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties1 = new A.LeftBorderLineProperties();
            A.NoFill noFill1 = new A.NoFill();

            leftBorderLineProperties1.Append(noFill1);

            A.RightBorderLineProperties rightBorderLineProperties1 = new A.RightBorderLineProperties();
            A.NoFill noFill2 = new A.NoFill();

            rightBorderLineProperties1.Append(noFill2);

            A.TopBorderLineProperties topBorderLineProperties1 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill3.Append(rgbColorModelHex3);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd1 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties1.Append(solidFill3);
            topBorderLineProperties1.Append(presetDash1);
            topBorderLineProperties1.Append(round1);
            topBorderLineProperties1.Append(headEnd1);
            topBorderLineProperties1.Append(tailEnd1);

            A.BottomBorderLineProperties bottomBorderLineProperties1 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill4.Append(rgbColorModelHex4);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties1.Append(solidFill4);
            bottomBorderLineProperties1.Append(presetDash2);
            bottomBorderLineProperties1.Append(round2);
            bottomBorderLineProperties1.Append(headEnd2);
            bottomBorderLineProperties1.Append(tailEnd2);

            A.SolidFill solidFill5 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill5.Append(rgbColorModelHex5);

            tableCellProperties1.Append(leftBorderLineProperties1);
            tableCellProperties1.Append(rightBorderLineProperties1);
            tableCellProperties1.Append(topBorderLineProperties1);
            tableCellProperties1.Append(bottomBorderLineProperties1);
            tableCellProperties1.Append(solidFill5);

            tableCell1.Append(textBody1);
            tableCell1.Append(tableCellProperties1);

            A.TableCell tableCell2 = new A.TableCell();

            A.TextBody textBody2 = new A.TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike };
            A.EffectList effectList3 = new A.EffectList();
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties2.Append(effectList3);
            runProperties2.Append(latinFont3);
            A.Text text2 = new A.Text();
            text2.Text = firstRank;

            run2.Append(runProperties2);
            run2.Append(text2);

            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };
            A.EffectList effectList4 = new A.EffectList();
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };

            endParagraphRunProperties2.Append(effectList4);
            endParagraphRunProperties2.Append(latinFont4);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            A.TableCellProperties tableCellProperties2 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties2 = new A.LeftBorderLineProperties();
            A.NoFill noFill3 = new A.NoFill();

            leftBorderLineProperties2.Append(noFill3);

            A.RightBorderLineProperties rightBorderLineProperties2 = new A.RightBorderLineProperties();
            A.NoFill noFill4 = new A.NoFill();

            rightBorderLineProperties2.Append(noFill4);

            A.TopBorderLineProperties topBorderLineProperties2 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill6.Append(rgbColorModelHex6);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3 = new A.Round();
            A.HeadEnd headEnd3 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties2.Append(solidFill6);
            topBorderLineProperties2.Append(presetDash3);
            topBorderLineProperties2.Append(round3);
            topBorderLineProperties2.Append(headEnd3);
            topBorderLineProperties2.Append(tailEnd3);

            A.BottomBorderLineProperties bottomBorderLineProperties2 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill7.Append(rgbColorModelHex7);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round4 = new A.Round();
            A.HeadEnd headEnd4 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd4 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties2.Append(solidFill7);
            bottomBorderLineProperties2.Append(presetDash4);
            bottomBorderLineProperties2.Append(round4);
            bottomBorderLineProperties2.Append(headEnd4);
            bottomBorderLineProperties2.Append(tailEnd4);

            A.SolidFill solidFill8 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill8.Append(rgbColorModelHex8);

            tableCellProperties2.Append(leftBorderLineProperties2);
            tableCellProperties2.Append(rightBorderLineProperties2);
            tableCellProperties2.Append(topBorderLineProperties2);
            tableCellProperties2.Append(bottomBorderLineProperties2);
            tableCellProperties2.Append(solidFill8);

            tableCell2.Append(textBody2);
            tableCell2.Append(tableCellProperties2);

            A.TableCell tableCell3 = new A.TableCell();

            A.TextBody textBody3 = new A.TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run3 = new A.Run();

            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike };
            A.EffectList effectList5 = new A.EffectList();
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties3.Append(effectList5);
            runProperties3.Append(latinFont5);
            A.Text text3 = new A.Text();
            text3.Text = "14";

            run3.Append(runProperties3);
            run3.Append(text3);

            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };
            A.EffectList effectList6 = new A.EffectList();
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };

            endParagraphRunProperties3.Append(effectList6);
            endParagraphRunProperties3.Append(latinFont6);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);
            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            A.TableCellProperties tableCellProperties3 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties3 = new A.LeftBorderLineProperties();
            A.NoFill noFill5 = new A.NoFill();

            leftBorderLineProperties3.Append(noFill5);

            A.RightBorderLineProperties rightBorderLineProperties3 = new A.RightBorderLineProperties();
            A.NoFill noFill6 = new A.NoFill();

            rightBorderLineProperties3.Append(noFill6);

            A.TopBorderLineProperties topBorderLineProperties3 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill9.Append(rgbColorModelHex9);
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round5 = new A.Round();
            A.HeadEnd headEnd5 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd5 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties3.Append(solidFill9);
            topBorderLineProperties3.Append(presetDash5);
            topBorderLineProperties3.Append(round5);
            topBorderLineProperties3.Append(headEnd5);
            topBorderLineProperties3.Append(tailEnd5);

            A.BottomBorderLineProperties bottomBorderLineProperties3 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill10.Append(rgbColorModelHex10);
            A.PresetDash presetDash6 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6 = new A.Round();
            A.HeadEnd headEnd6 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties3.Append(solidFill10);
            bottomBorderLineProperties3.Append(presetDash6);
            bottomBorderLineProperties3.Append(round6);
            bottomBorderLineProperties3.Append(headEnd6);
            bottomBorderLineProperties3.Append(tailEnd6);

            A.SolidFill solidFill11 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill11.Append(rgbColorModelHex11);

            tableCellProperties3.Append(leftBorderLineProperties3);
            tableCellProperties3.Append(rightBorderLineProperties3);
            tableCellProperties3.Append(topBorderLineProperties3);
            tableCellProperties3.Append(bottomBorderLineProperties3);
            tableCellProperties3.Append(solidFill11);

            tableCell3.Append(textBody3);
            tableCell3.Append(tableCellProperties3);

            A.TableCell tableCell4 = new A.TableCell();

            A.TextBody textBody4 = new A.TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run4 = new A.Run();

            A.RunProperties runProperties4 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike };
            A.EffectList effectList7 = new A.EffectList();
            A.LatinFont latinFont7 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties4.Append(effectList7);
            runProperties4.Append(latinFont7);
            A.Text text4 = new A.Text();
            text4.Text = thirdRank;

            run4.Append(runProperties4);
            run4.Append(text4);

            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };
            A.EffectList effectList8 = new A.EffectList();
            A.LatinFont latinFont8 = new A.LatinFont() { Typeface = "+mn-lt" };

            endParagraphRunProperties4.Append(effectList8);
            endParagraphRunProperties4.Append(latinFont8);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);
            paragraph4.Append(endParagraphRunProperties4);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);

            A.TableCellProperties tableCellProperties4 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties4 = new A.LeftBorderLineProperties();
            A.NoFill noFill7 = new A.NoFill();

            leftBorderLineProperties4.Append(noFill7);

            A.RightBorderLineProperties rightBorderLineProperties4 = new A.RightBorderLineProperties();
            A.NoFill noFill8 = new A.NoFill();

            rightBorderLineProperties4.Append(noFill8);

            A.TopBorderLineProperties topBorderLineProperties4 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill12 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill12.Append(rgbColorModelHex12);
            A.PresetDash presetDash7 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round7 = new A.Round();
            A.HeadEnd headEnd7 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd7 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties4.Append(solidFill12);
            topBorderLineProperties4.Append(presetDash7);
            topBorderLineProperties4.Append(round7);
            topBorderLineProperties4.Append(headEnd7);
            topBorderLineProperties4.Append(tailEnd7);

            A.BottomBorderLineProperties bottomBorderLineProperties4 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill13.Append(rgbColorModelHex13);
            A.PresetDash presetDash8 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round8 = new A.Round();
            A.HeadEnd headEnd8 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd8 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties4.Append(solidFill13);
            bottomBorderLineProperties4.Append(presetDash8);
            bottomBorderLineProperties4.Append(round8);
            bottomBorderLineProperties4.Append(headEnd8);
            bottomBorderLineProperties4.Append(tailEnd8);

            A.SolidFill solidFill14 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill14.Append(rgbColorModelHex14);

            tableCellProperties4.Append(leftBorderLineProperties4);
            tableCellProperties4.Append(rightBorderLineProperties4);
            tableCellProperties4.Append(topBorderLineProperties4);
            tableCellProperties4.Append(bottomBorderLineProperties4);
            tableCellProperties4.Append(solidFill14);

            tableCell4.Append(textBody4);
            tableCell4.Append(tableCellProperties4);

            A.TableCell tableCell5 = new A.TableCell();

            A.TextBody textBody5 = new A.TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run5 = new A.Run();

            A.RunProperties runProperties5 = new A.RunProperties() { Language = "en-US", FontSize = 1400, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex15 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill15.Append(rgbColorModelHex15);
            A.EffectList effectList9 = new A.EffectList();
            A.LatinFont latinFont9 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties5.Append(solidFill15);
            runProperties5.Append(effectList9);
            runProperties5.Append(latinFont9);
            A.Text text5 = new A.Text();
            text5.Text = weightedScores;

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph5);

            A.TableCellProperties tableCellProperties5 = new A.TableCellProperties() { LeftMargin = 3721, RightMargin = 3721, TopMargin = 3721, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties5 = new A.LeftBorderLineProperties();
            A.NoFill noFill9 = new A.NoFill();

            leftBorderLineProperties5.Append(noFill9);

            A.RightBorderLineProperties rightBorderLineProperties5 = new A.RightBorderLineProperties();
            A.NoFill noFill10 = new A.NoFill();

            rightBorderLineProperties5.Append(noFill10);

            A.TopBorderLineProperties topBorderLineProperties5 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex16 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill16.Append(rgbColorModelHex16);
            A.PresetDash presetDash9 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round9 = new A.Round();
            A.HeadEnd headEnd9 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd9 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties5.Append(solidFill16);
            topBorderLineProperties5.Append(presetDash9);
            topBorderLineProperties5.Append(round9);
            topBorderLineProperties5.Append(headEnd9);
            topBorderLineProperties5.Append(tailEnd9);

            A.BottomBorderLineProperties bottomBorderLineProperties5 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex17 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill17.Append(rgbColorModelHex17);
            A.PresetDash presetDash10 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round10 = new A.Round();
            A.HeadEnd headEnd10 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd10 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties5.Append(solidFill17);
            bottomBorderLineProperties5.Append(presetDash10);
            bottomBorderLineProperties5.Append(round10);
            bottomBorderLineProperties5.Append(headEnd10);
            bottomBorderLineProperties5.Append(tailEnd10);

            A.SolidFill solidFill18 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex18 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill18.Append(rgbColorModelHex18);

            tableCellProperties5.Append(leftBorderLineProperties5);
            tableCellProperties5.Append(rightBorderLineProperties5);
            tableCellProperties5.Append(topBorderLineProperties5);
            tableCellProperties5.Append(bottomBorderLineProperties5);
            tableCellProperties5.Append(solidFill18);

            tableCell5.Append(textBody5);
            tableCell5.Append(tableCellProperties5);

            A.ExtensionList extensionList1 = new A.ExtensionList();

            A.Extension extension1 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

         
           
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            tableRow1.Append(tableCell5);
            tableRow1.Append(extensionList1);
            table1.Append(tableRow1);


            slidePart1.Slide.Save();


            


        }

        #region Binary Data
        private string extendedPart1Data = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0KPHAxNTEwOnJldkluZm8geG1sbnM6YT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL2RyYXdpbmdtbC8yMDA2L21haW4iIHhtbG5zOnI9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMiIHhtbG5zOnAxNTEwPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9wb3dlcnBvaW50LzIwMTUvMTAvbWFpbiI+PHAxNTEwOnJldkxzdD48cDE1MTA6Y2xpZW50IGlkPSJ7MkNGQzgyQTctQzExRC00MkVELUIzMzgtQzY1RTk1MTU0NTAwfSIgdj0iMTgiIGR0PSIyMDIzLTEwLTE3VDE3OjM5OjUwLjQ2MSIvPjwvcDE1MTA6cmV2THN0PjwvcDE1MTA6cmV2SW5mbz4=";

        private string thumbnailPart1Data = "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";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }


    public class clsLikeabilityRationalCharts
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private PresentationDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = PresentationDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            // Deletes parts only existing in the source package.
            DeleteParts();
            //Changes the relationship ID of the parts.
            ReconfigureRelationshipID();
            //Changes the contents of the specified parts.
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeThumbnailPart1(((ThumbnailPart)UriPartDictionary["/docProps/thumbnail.jpeg"]));
            ChangePresentationPart1(document.PresentationPart);
            ChangeExtendedFilePropertiesPart1(((ExtendedFilePropertiesPart)UriPartDictionary["/docProps/app.xml"]));
            //ChangeSlidePart1(((SlidePart)UriPartDictionary["/ppt/slides/slide1.xml"]));
        }

        /// <summary>
        /// Stores the references to all the parts in the package.
        /// They could be retrieved by their URIs later.
        /// </summary>
        private void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(document);
            while (queue.Count > 0)
            {
                foreach (var part in queue.Dequeue().Parts)
                {
                    if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                    {
                        UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                        queue.Enqueue(part.OpenXmlPart);
                    }
                }
            }
        }

        /// <summary>
        /// Deletes parts only existing in the source package.
        /// </summary>
        private void DeleteParts()
        {
            document.PresentationPart.DeletePart("rId8");
            document.PresentationPart.DeletePart("rId13");
            document.PresentationPart.DeletePart("rId3");
            document.PresentationPart.DeletePart("rId7");
            document.PresentationPart.DeletePart("rId12");
            document.PresentationPart.DeletePart("rId16");
            document.PresentationPart.DeletePart("rId6");
            document.PresentationPart.DeletePart("rId11");
            document.PresentationPart.DeletePart("rId5");
            document.PresentationPart.DeletePart("rId15");
            document.PresentationPart.DeletePart("rId10");
            document.PresentationPart.DeletePart("rId4");
            document.PresentationPart.DeletePart("rId9");
            document.PresentationPart.DeletePart("rId14");
        }

        /// <summary>
        /// Changes the relationship ID of the parts in the source package to make sure these IDs are the same as those in the target package.
        /// To avoid the conflict of the relationship ID, a temporary ID is assigned first.        
        /// </summary>
        private void ReconfigureRelationshipID()
        {
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/handoutMasters/handoutMaster1.xml"], "generatedTmpID1");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/viewProps.xml"], "generatedTmpID2");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/notesMasters/notesMaster1.xml"], "generatedTmpID3");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/presProps.xml"], "generatedTmpID4");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/authors.xml"], "generatedTmpID5");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/tableStyles.xml"], "generatedTmpID6");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/commentAuthors.xml"], "generatedTmpID7");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/theme/theme1.xml"], "generatedTmpID8");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/handoutMasters/handoutMaster1.xml"], "rId4");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/viewProps.xml"], "rId7");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/notesMasters/notesMaster1.xml"], "rId3");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/presProps.xml"], "rId6");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/authors.xml"], "rId10");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/tableStyles.xml"], "rId9");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/commentAuthors.xml"], "rId5");
            document.PresentationPart.ChangeIdOfPart(UriPartDictionary["/ppt/theme/theme1.xml"], "rId8");
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2024-11-13T22:02:18Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeThumbnailPart1(ThumbnailPart thumbnailPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        private void ChangePresentationPart1(PresentationPart presentationPart1)
        {
            Presentation presentation1 = presentationPart1.Presentation;

            NotesMasterIdList notesMasterIdList1 = presentation1.GetFirstChild<NotesMasterIdList>();
            HandoutMasterIdList handoutMasterIdList1 = presentation1.GetFirstChild<HandoutMasterIdList>();
            SlideIdList slideIdList1 = presentation1.GetFirstChild<SlideIdList>();

            NotesMasterId notesMasterId1 = notesMasterIdList1.GetFirstChild<NotesMasterId>();
            notesMasterId1.Id = "rId3";

            HandoutMasterId handoutMasterId1 = handoutMasterIdList1.GetFirstChild<HandoutMasterId>();
            handoutMasterId1.Id = "rId4";

            SlideId slideId1 = slideIdList1.Elements<SlideId>().ElementAt(1);
            SlideId slideId2 = slideIdList1.Elements<SlideId>().ElementAt(2);
            SlideId slideId3 = slideIdList1.Elements<SlideId>().ElementAt(3);
            SlideId slideId4 = slideIdList1.Elements<SlideId>().ElementAt(4);
            SlideId slideId5 = slideIdList1.Elements<SlideId>().ElementAt(5);
            SlideId slideId6 = slideIdList1.Elements<SlideId>().ElementAt(6);
            SlideId slideId7 = slideIdList1.Elements<SlideId>().ElementAt(7);
            SlideId slideId8 = slideIdList1.Elements<SlideId>().ElementAt(8);
            SlideId slideId9 = slideIdList1.Elements<SlideId>().ElementAt(9);
            SlideId slideId10 = slideIdList1.Elements<SlideId>().ElementAt(10);
            SlideId slideId11 = slideIdList1.Elements<SlideId>().ElementAt(11);
            SlideId slideId12 = slideIdList1.Elements<SlideId>().ElementAt(12);
            SlideId slideId13 = slideIdList1.Elements<SlideId>().ElementAt(13);
            SlideId slideId14 = slideIdList1.Elements<SlideId>().ElementAt(14);

            slideId1.Remove();
            slideId2.Remove();
            slideId3.Remove();
            slideId4.Remove();
            slideId5.Remove();
            slideId6.Remove();
            slideId7.Remove();
            slideId8.Remove();
            slideId9.Remove();
            slideId10.Remove();
            slideId11.Remove();
            slideId12.Remove();
            slideId13.Remove();
            slideId14.Remove();
        }

        private void ChangeExtendedFilePropertiesPart1(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = extendedFilePropertiesPart1.Properties;

            Ap.TotalTime totalTime1 = properties1.GetFirstChild<Ap.TotalTime>();
            Ap.Words words1 = properties1.GetFirstChild<Ap.Words>();
            Ap.Paragraphs paragraphs1 = properties1.GetFirstChild<Ap.Paragraphs>();
            Ap.Slides slides1 = properties1.GetFirstChild<Ap.Slides>();
            Ap.HeadingPairs headingPairs1 = properties1.GetFirstChild<Ap.HeadingPairs>();
            Ap.TitlesOfParts titlesOfParts1 = properties1.GetFirstChild<Ap.TitlesOfParts>();
            totalTime1.Text = "110035";

            words1.Text = "91";

            paragraphs1.Text = "27";

            slides1.Text = "1";


            Vt.VTVector vTVector1 = headingPairs1.GetFirstChild<Vt.VTVector>();

            Vt.Variant variant1 = vTVector1.Elements<Vt.Variant>().ElementAt(5);

            Vt.VTInt32 vTInt321 = variant1.GetFirstChild<Vt.VTInt32>();
            vTInt321.Text = "1";


            Vt.VTVector vTVector2 = titlesOfParts1.GetFirstChild<Vt.VTVector>();
            vTVector2.Size = (UInt32Value)10U;

            Vt.VTLPSTR vTLPSTR1 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(10);
            Vt.VTLPSTR vTLPSTR2 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(11);
            Vt.VTLPSTR vTLPSTR3 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(12);
            Vt.VTLPSTR vTLPSTR4 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(13);
            Vt.VTLPSTR vTLPSTR5 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(14);
            Vt.VTLPSTR vTLPSTR6 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(15);
            Vt.VTLPSTR vTLPSTR7 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(16);
            Vt.VTLPSTR vTLPSTR8 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(17);
            Vt.VTLPSTR vTLPSTR9 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(18);
            Vt.VTLPSTR vTLPSTR10 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(19);
            Vt.VTLPSTR vTLPSTR11 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(20);
            Vt.VTLPSTR vTLPSTR12 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(21);
            Vt.VTLPSTR vTLPSTR13 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(22);
            Vt.VTLPSTR vTLPSTR14 = vTVector2.Elements<Vt.VTLPSTR>().ElementAt(23);

            vTLPSTR1.Remove();
            vTLPSTR2.Remove();
            vTLPSTR3.Remove();
            vTLPSTR4.Remove();
            vTLPSTR5.Remove();
            vTLPSTR6.Remove();
            vTLPSTR7.Remove();
            vTLPSTR8.Remove();
            vTLPSTR9.Remove();
            vTLPSTR10.Remove();
            vTLPSTR11.Remove();
            vTLPSTR12.Remove();
            vTLPSTR13.Remove();
            vTLPSTR14.Remove();
        }

        public void ChangeSlidePartLikeabilityRational1(SlidePart slidePart1,string testName1, string testName2,string total,string sum1,string sum2,string strLikeability1,string strLikeability2,string percentage1,string percentage2)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.GetFirstChild<GraphicFrame>();
            GraphicFrame graphicFrame2 = shapeTree1.Elements<GraphicFrame>().ElementAt(1);

            A.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<A.NonVisualGraphicFrameProperties>();
            Transform transform1 = graphicFrame1.GetFirstChild<Transform>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
            modificationId1.Val = (UInt32Value)883713824U;

            A.Extents extents1 = transform1.GetFirstChild<A.Extents>();
            extents1.Cy = 1228725L;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            A.TableRow tableRow1 = table1.GetFirstChild<A.TableRow>();
            A.TableRow tableRow2 = table1.Elements<A.TableRow>().ElementAt(1);
            A.TableRow tableRow3 = table1.Elements<A.TableRow>().ElementAt(2);
            A.TableRow tableRow4 = table1.Elements<A.TableRow>().ElementAt(4);

            A.TableCell tableCell1 = tableRow1.GetFirstChild<A.TableCell>();

            A.TextBody textBody1 = tableCell1.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.Run run1 = paragraph1.GetFirstChild<A.Run>();

            A.Text text1 = run1.GetFirstChild<A.Text>();
            text1.Text = testName1;


            A.TableCell tableCell2 = tableRow2.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody2 = tableCell2.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph2 = textBody2.GetFirstChild<A.Paragraph>();

            A.Run run2 = paragraph2.GetFirstChild<A.Run>();

            A.Text text2 = run2.GetFirstChild<A.Text>();
            text2.Text = "Overall";


            A.TableCell tableCell3 = tableRow3.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody3 = tableCell3.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph3 = textBody3.GetFirstChild<A.Paragraph>();

            A.Run run3 = paragraph3.GetFirstChild<A.Run>();
            //A.Run run4 = paragraph3.Elements<A.Run>().ElementAt(1);
            //A.Run run5 = paragraph3.Elements<A.Run>().ElementAt(2);

            A.Text text3 = run3.GetFirstChild<A.Text>();
            text3.Text = "n = "+total;


            //run4.Remove();
           // run5.Remove();

            A.TableCell tableCell4 = tableRow4.GetFirstChild<A.TableCell>();
            A.TableCell tableCell5 = tableRow4.Elements<A.TableCell>().ElementAt(1);
            A.TableCell tableCell6 = tableRow4.Elements<A.TableCell>().ElementAt(2);

            A.TextBody textBody4 = tableCell4.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph4 = textBody4.GetFirstChild<A.Paragraph>();

            A.Run run6 = paragraph4.GetFirstChild<A.Run>();

            A.Text text4 = run6.GetFirstChild<A.Text>();
            text4.Text = testName2;


            A.TextBody textBody5 = tableCell5.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph5 = textBody5.GetFirstChild<A.Paragraph>();

            A.Run run7 = paragraph5.GetFirstChild<A.Run>();

            A.Text text5 = run7.GetFirstChild<A.Text>();
            text5.Text = sum1;


            A.TextBody textBody6 = tableCell6.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph6 = textBody6.GetFirstChild<A.Paragraph>();

            A.Run run8 = paragraph6.GetFirstChild<A.Run>();

            A.Text text6 = run8.GetFirstChild<A.Text>();
            text6.Text = percentage1+"%";


            A.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties2 = graphicFrame2.GetFirstChild<A.NonVisualGraphicFrameProperties>();
            Transform transform2 = graphicFrame2.GetFirstChild<Transform>();
            A.Graphic graphic2 = graphicFrame2.GetFirstChild<A.Graphic>();

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = nonVisualGraphicFrameProperties2.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList2 = applicationNonVisualDrawingProperties2.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension2 = applicationNonVisualDrawingPropertiesExtensionList2.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            P14.ModificationId modificationId2 = applicationNonVisualDrawingPropertiesExtension2.GetFirstChild<P14.ModificationId>();
            modificationId2.Val = (UInt32Value)1329561748U;

            A.Extents extents2 = transform2.GetFirstChild<A.Extents>();
            extents2.Cy = 1228725L;

            A.GraphicData graphicData2 = graphic2.GetFirstChild<A.GraphicData>();

            A.Table table2 = graphicData2.GetFirstChild<A.Table>();

            A.TableRow tableRow5 = table2.GetFirstChild<A.TableRow>();
            A.TableRow tableRow6 = table2.Elements<A.TableRow>().ElementAt(1);
            A.TableRow tableRow7 = table2.Elements<A.TableRow>().ElementAt(2);
            A.TableRow tableRow8 = table2.Elements<A.TableRow>().ElementAt(4);

            A.TableCell tableCell7 = tableRow5.GetFirstChild<A.TableCell>();

            A.TextBody textBody7 = tableCell7.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph7 = textBody7.GetFirstChild<A.Paragraph>();

            A.Run run9 = paragraph7.GetFirstChild<A.Run>();

            A.Text text7 = run9.GetFirstChild<A.Text>();
            text7.Text = testName1;


            A.TableCell tableCell8 = tableRow6.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody8 = tableCell8.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph8 = textBody8.GetFirstChild<A.Paragraph>();

            A.Run run10 = paragraph8.GetFirstChild<A.Run>();

            A.Text text8 = run10.GetFirstChild<A.Text>();
            text8.Text = "Overall";


            A.TableCell tableCell9 = tableRow7.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody9 = tableCell9.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph9 = textBody9.GetFirstChild<A.Paragraph>();

            A.Run run11 = paragraph9.GetFirstChild<A.Run>();
          //  A.Run run12 = paragraph9.Elements<A.Run>().ElementAt(1);
          //  A.Run run13 = paragraph9.Elements<A.Run>().ElementAt(2);

            A.Text text9 = run11.GetFirstChild<A.Text>();
            text9.Text = "n = "+total;


          //  run12.Remove();
           // run13.Remove();

            A.TableCell tableCell10 = tableRow8.GetFirstChild<A.TableCell>();
            A.TableCell tableCell11 = tableRow8.Elements<A.TableCell>().ElementAt(1);
            A.TableCell tableCell12 = tableRow8.Elements<A.TableCell>().ElementAt(2);

            A.TextBody textBody10 = tableCell10.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph10 = textBody10.GetFirstChild<A.Paragraph>();

            A.Run run14 = paragraph10.GetFirstChild<A.Run>();

            A.Text text10 = run14.GetFirstChild<A.Text>();
            text10.Text = testName2;


            A.TextBody textBody11 = tableCell11.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph11 = textBody11.GetFirstChild<A.Paragraph>();

            A.Run run15 = paragraph11.GetFirstChild<A.Run>();

            A.Text text11 = run15.GetFirstChild<A.Text>();
            text11.Text = sum2;


            A.TextBody textBody12 = tableCell12.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph12 = textBody12.GetFirstChild<A.Paragraph>();

            A.Run run16 = paragraph12.GetFirstChild<A.Run>();

            A.Text text12 = run16.GetFirstChild<A.Text>();
            text12.Text = percentage2+"%";

        }

        public void ChangeSlidePartLikeabilityRational(SlidePart slidePart1, string testName1, string testName2, string total, string sum1, string sum2, string strLikeability1, string strLikeability2, string percentage1, string percentage2)
        {
            Slide slide1 = slidePart1.Slide;
            //slide1.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main");
            //slide1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            //slide1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.GetFirstChild<GraphicFrame>();
            GraphicFrame graphicFrame2 = shapeTree1.Elements<GraphicFrame>().ElementAt(1);

            A.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<A.NonVisualGraphicFrameProperties>();
            Transform transform1 = graphicFrame1.GetFirstChild<Transform>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
            modificationId1.Val = (UInt32Value)883713824U;

            A.Extents extents1 = transform1.GetFirstChild<A.Extents>();
            extents1.Cy = 1228725L;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            A.TableRow tableRow1 = table1.GetFirstChild<A.TableRow>();
            A.TableRow tableRow2 = table1.Elements<A.TableRow>().ElementAt(1);
            A.TableRow tableRow3 = table1.Elements<A.TableRow>().ElementAt(2);

            A.TableCell tableCell1 = tableRow1.GetFirstChild<A.TableCell>();

            A.TextBody textBody1 = tableCell1.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.Run run1 = paragraph1.GetFirstChild<A.Run>();

            A.Text text1 = run1.GetFirstChild<A.Text>();
            text1.Text = testName1 ;


            A.TableCell tableCell2 = tableRow2.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody2 = tableCell2.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph2 = textBody2.GetFirstChild<A.Paragraph>();

            A.Run run2 = paragraph2.GetFirstChild<A.Run>();

            A.Text text2 = run2.GetFirstChild<A.Text>();
            text2.Text = "Overall";


            A.TableCell tableCell3 = tableRow3.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody3 = tableCell3.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph3 = textBody3.GetFirstChild<A.Paragraph>();

            A.Run run3 = paragraph3.GetFirstChild<A.Run>();
            //A.Run run4 = paragraph3.Elements<A.Run>().ElementAt(1);
            //A.Run run5 = paragraph3.Elements<A.Run>().ElementAt(2);

            A.Text text3 = run3.GetFirstChild<A.Text>();
            text3.Text = "n = "+total;


            //run4.Remove();
            //run5.Remove();

            A.TableRow tableRow4 = new A.TableRow() { Height = 238125L };

            A.TableCell tableCell4 = new A.TableCell();

            A.TextBody textBody4 = new A.TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run6 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill1.Append(rgbColorModelHex1);
            A.EffectList effectList1 = new A.EffectList();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties1.Append(solidFill1);
            runProperties1.Append(effectList1);
            runProperties1.Append(latinFont1);
            A.Text text4 = new A.Text();
            text4.Text = testName2;

            run6.Append(runProperties1);
            run6.Append(text4);

            paragraph4.Append(paragraphProperties1);
            paragraph4.Append(run6);

            textBody4.Append(bodyProperties1);
            textBody4.Append(listStyle1);
            textBody4.Append(paragraph4);

            A.TableCellProperties tableCellProperties1 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties1 = new A.LeftBorderLineProperties();
            A.NoFill noFill1 = new A.NoFill();

            leftBorderLineProperties1.Append(noFill1);

            A.RightBorderLineProperties rightBorderLineProperties1 = new A.RightBorderLineProperties();
            A.NoFill noFill2 = new A.NoFill();

            rightBorderLineProperties1.Append(noFill2);

            A.TopBorderLineProperties topBorderLineProperties1 = new A.TopBorderLineProperties();
            A.NoFill noFill3 = new A.NoFill();

            topBorderLineProperties1.Append(noFill3);

            A.BottomBorderLineProperties bottomBorderLineProperties1 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill2.Append(rgbColorModelHex2);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd1 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties1.Append(solidFill2);
            bottomBorderLineProperties1.Append(presetDash1);
            bottomBorderLineProperties1.Append(round1);
            bottomBorderLineProperties1.Append(headEnd1);
            bottomBorderLineProperties1.Append(tailEnd1);

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill3.Append(rgbColorModelHex3);

            tableCellProperties1.Append(leftBorderLineProperties1);
            tableCellProperties1.Append(rightBorderLineProperties1);
            tableCellProperties1.Append(topBorderLineProperties1);
            tableCellProperties1.Append(bottomBorderLineProperties1);
            tableCellProperties1.Append(solidFill3);

            tableCell4.Append(textBody4);
            tableCell4.Append(tableCellProperties1);

            A.TableCell tableCell5 = new A.TableCell();

            A.TextBody textBody5 = new A.TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run7 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill4.Append(rgbColorModelHex4);
            A.EffectList effectList2 = new A.EffectList();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties2.Append(solidFill4);
            runProperties2.Append(effectList2);
            runProperties2.Append(latinFont2);
            A.Text text5 = new A.Text();
            text5.Text = sum2;

            run7.Append(runProperties2);
            run7.Append(text5);

            paragraph5.Append(paragraphProperties2);
            paragraph5.Append(run7);

            textBody5.Append(bodyProperties2);
            textBody5.Append(listStyle2);
            textBody5.Append(paragraph5);

            A.TableCellProperties tableCellProperties2 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties2 = new A.LeftBorderLineProperties();
            A.NoFill noFill4 = new A.NoFill();

            leftBorderLineProperties2.Append(noFill4);

            A.RightBorderLineProperties rightBorderLineProperties2 = new A.RightBorderLineProperties();
            A.NoFill noFill5 = new A.NoFill();

            rightBorderLineProperties2.Append(noFill5);

            A.TopBorderLineProperties topBorderLineProperties2 = new A.TopBorderLineProperties();
            A.NoFill noFill6 = new A.NoFill();

            topBorderLineProperties2.Append(noFill6);

            A.BottomBorderLineProperties bottomBorderLineProperties2 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill5.Append(rgbColorModelHex5);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties2.Append(solidFill5);
            bottomBorderLineProperties2.Append(presetDash2);
            bottomBorderLineProperties2.Append(round2);
            bottomBorderLineProperties2.Append(headEnd2);
            bottomBorderLineProperties2.Append(tailEnd2);

            A.SolidFill solidFill6 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill6.Append(rgbColorModelHex6);

            tableCellProperties2.Append(leftBorderLineProperties2);
            tableCellProperties2.Append(rightBorderLineProperties2);
            tableCellProperties2.Append(topBorderLineProperties2);
            tableCellProperties2.Append(bottomBorderLineProperties2);
            tableCellProperties2.Append(solidFill6);

            tableCell5.Append(textBody5);
            tableCell5.Append(tableCellProperties2);

            A.TableCell tableCell6 = new A.TableCell();

            A.TextBody textBody6 = new A.TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run8 = new A.Run();

            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill7.Append(rgbColorModelHex7);
            A.EffectList effectList3 = new A.EffectList();
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties3.Append(solidFill7);
            runProperties3.Append(effectList3);
            runProperties3.Append(latinFont3);
            A.Text text6 = new A.Text();
            text6.Text = percentage2;

            run8.Append(runProperties3);
            run8.Append(text6);

            paragraph6.Append(paragraphProperties3);
            paragraph6.Append(run8);

            textBody6.Append(bodyProperties3);
            textBody6.Append(listStyle3);
            textBody6.Append(paragraph6);

            A.TableCellProperties tableCellProperties3 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties3 = new A.LeftBorderLineProperties();
            A.NoFill noFill7 = new A.NoFill();

            leftBorderLineProperties3.Append(noFill7);

            A.RightBorderLineProperties rightBorderLineProperties3 = new A.RightBorderLineProperties();
            A.NoFill noFill8 = new A.NoFill();

            rightBorderLineProperties3.Append(noFill8);

            A.TopBorderLineProperties topBorderLineProperties3 = new A.TopBorderLineProperties();
            A.NoFill noFill9 = new A.NoFill();

            topBorderLineProperties3.Append(noFill9);

            A.BottomBorderLineProperties bottomBorderLineProperties3 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill8.Append(rgbColorModelHex8);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3 = new A.Round();
            A.HeadEnd headEnd3 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties3.Append(solidFill8);
            bottomBorderLineProperties3.Append(presetDash3);
            bottomBorderLineProperties3.Append(round3);
            bottomBorderLineProperties3.Append(headEnd3);
            bottomBorderLineProperties3.Append(tailEnd3);

            A.SolidFill solidFill9 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill9.Append(rgbColorModelHex9);

            tableCellProperties3.Append(leftBorderLineProperties3);
            tableCellProperties3.Append(rightBorderLineProperties3);
            tableCellProperties3.Append(topBorderLineProperties3);
            tableCellProperties3.Append(bottomBorderLineProperties3);
            tableCellProperties3.Append(solidFill9);

            tableCell6.Append(textBody6);
            tableCell6.Append(tableCellProperties3);

            A.ExtensionList extensionList1 = new A.ExtensionList();

            A.Extension extension1 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            //OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2346549190\" />");

            //extension1.Append(openXmlUnknownElement1);

            extensionList1.Append(extension1);

            tableRow4.Append(tableCell4);
            tableRow4.Append(tableCell5);
            tableRow4.Append(tableCell6);
            tableRow4.Append(extensionList1);
            table1.Append(tableRow4);

            A.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties2 = graphicFrame2.GetFirstChild<A.NonVisualGraphicFrameProperties>();
            Transform transform2 = graphicFrame2.GetFirstChild<Transform>();
            A.Graphic graphic2 = graphicFrame2.GetFirstChild<A.Graphic>();

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = nonVisualGraphicFrameProperties2.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList2 = applicationNonVisualDrawingProperties2.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension2 = applicationNonVisualDrawingPropertiesExtensionList2.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            P14.ModificationId modificationId2 = applicationNonVisualDrawingPropertiesExtension2.GetFirstChild<P14.ModificationId>();
            modificationId2.Val = (UInt32Value)1329561748U;

            A.Extents extents2 = transform2.GetFirstChild<A.Extents>();
            extents2.Cy = 1228725L;

            A.GraphicData graphicData2 = graphic2.GetFirstChild<A.GraphicData>();

            A.Table table2 = graphicData2.GetFirstChild<A.Table>();

            A.TableRow tableRow5 = table2.GetFirstChild<A.TableRow>();
            A.TableRow tableRow6 = table2.Elements<A.TableRow>().ElementAt(1);
            A.TableRow tableRow7 = table2.Elements<A.TableRow>().ElementAt(2);

            A.TableCell tableCell7 = tableRow5.GetFirstChild<A.TableCell>();

            A.TextBody textBody7 = tableCell7.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph7 = textBody7.GetFirstChild<A.Paragraph>();

            A.Run run9 = paragraph7.GetFirstChild<A.Run>();

            A.Text text7 = run9.GetFirstChild<A.Text>();
            text7.Text = testName1;


            A.TableCell tableCell8 = tableRow6.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody8 = tableCell8.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph8 = textBody8.GetFirstChild<A.Paragraph>();

            A.Run run10 = paragraph8.GetFirstChild<A.Run>();

            A.Text text8 = run10.GetFirstChild<A.Text>();
            text8.Text = "Overall";


            A.TableCell tableCell9 = tableRow7.Elements<A.TableCell>().ElementAt(1);

            A.TextBody textBody9 = tableCell9.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph9 = textBody9.GetFirstChild<A.Paragraph>();

            A.Run run11 = paragraph9.GetFirstChild<A.Run>();
           // A.Run run12 = paragraph9.Elements<A.Run>().ElementAt(1);
          //  A.Run run13 = paragraph9.Elements<A.Run>().ElementAt(2);

            A.Text text9 = run11.GetFirstChild<A.Text>();
            text9.Text = "n = "+total;


            //run12.Remove();
           // run13.Remove();

            A.TableRow tableRow8 = new A.TableRow() { Height = 238125L };

            A.TableCell tableCell10 = new A.TableCell();

            A.TextBody textBody10 = new A.TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run14 = new A.Run();

            A.RunProperties runProperties4 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill10.Append(rgbColorModelHex10);
            A.EffectList effectList4 = new A.EffectList();
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties4.Append(solidFill10);
            runProperties4.Append(effectList4);
            runProperties4.Append(latinFont4);
            A.Text text10 = new A.Text();
            text10.Text = testName1;

            run14.Append(runProperties4);
            run14.Append(text10);

            paragraph10.Append(paragraphProperties4);
            paragraph10.Append(run14);

            textBody10.Append(bodyProperties4);
            textBody10.Append(listStyle4);
            textBody10.Append(paragraph10);

            A.TableCellProperties tableCellProperties4 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties4 = new A.LeftBorderLineProperties();
            A.NoFill noFill10 = new A.NoFill();

            leftBorderLineProperties4.Append(noFill10);

            A.RightBorderLineProperties rightBorderLineProperties4 = new A.RightBorderLineProperties();
            A.NoFill noFill11 = new A.NoFill();

            rightBorderLineProperties4.Append(noFill11);

            A.TopBorderLineProperties topBorderLineProperties4 = new A.TopBorderLineProperties();
            A.NoFill noFill12 = new A.NoFill();

            topBorderLineProperties4.Append(noFill12);

            A.BottomBorderLineProperties bottomBorderLineProperties4 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill11.Append(rgbColorModelHex11);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round4 = new A.Round();
            A.HeadEnd headEnd4 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd4 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties4.Append(solidFill11);
            bottomBorderLineProperties4.Append(presetDash4);
            bottomBorderLineProperties4.Append(round4);
            bottomBorderLineProperties4.Append(headEnd4);
            bottomBorderLineProperties4.Append(tailEnd4);

            A.SolidFill solidFill12 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill12.Append(rgbColorModelHex12);

            tableCellProperties4.Append(leftBorderLineProperties4);
            tableCellProperties4.Append(rightBorderLineProperties4);
            tableCellProperties4.Append(topBorderLineProperties4);
            tableCellProperties4.Append(bottomBorderLineProperties4);
            tableCellProperties4.Append(solidFill12);

            tableCell10.Append(textBody10);
            tableCell10.Append(tableCellProperties4);

            A.TableCell tableCell11 = new A.TableCell();

            A.TextBody textBody11 = new A.TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run15 = new A.Run();

            A.RunProperties runProperties5 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill13.Append(rgbColorModelHex13);
            A.EffectList effectList5 = new A.EffectList();
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties5.Append(solidFill13);
            runProperties5.Append(effectList5);
            runProperties5.Append(latinFont5);
            A.Text text11 = new A.Text();
            text11.Text = sum1;

            run15.Append(runProperties5);
            run15.Append(text11);

            paragraph11.Append(paragraphProperties5);
            paragraph11.Append(run15);

            textBody11.Append(bodyProperties5);
            textBody11.Append(listStyle5);
            textBody11.Append(paragraph11);

            A.TableCellProperties tableCellProperties5 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties5 = new A.LeftBorderLineProperties();
            A.NoFill noFill13 = new A.NoFill();

            leftBorderLineProperties5.Append(noFill13);

            A.RightBorderLineProperties rightBorderLineProperties5 = new A.RightBorderLineProperties();
            A.NoFill noFill14 = new A.NoFill();

            rightBorderLineProperties5.Append(noFill14);

            A.TopBorderLineProperties topBorderLineProperties5 = new A.TopBorderLineProperties();
            A.NoFill noFill15 = new A.NoFill();

            topBorderLineProperties5.Append(noFill15);

            A.BottomBorderLineProperties bottomBorderLineProperties5 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill14.Append(rgbColorModelHex14);
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round5 = new A.Round();
            A.HeadEnd headEnd5 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd5 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties5.Append(solidFill14);
            bottomBorderLineProperties5.Append(presetDash5);
            bottomBorderLineProperties5.Append(round5);
            bottomBorderLineProperties5.Append(headEnd5);
            bottomBorderLineProperties5.Append(tailEnd5);

            A.SolidFill solidFill15 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex15 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill15.Append(rgbColorModelHex15);

            tableCellProperties5.Append(leftBorderLineProperties5);
            tableCellProperties5.Append(rightBorderLineProperties5);
            tableCellProperties5.Append(topBorderLineProperties5);
            tableCellProperties5.Append(bottomBorderLineProperties5);
            tableCellProperties5.Append(solidFill15);

            tableCell11.Append(textBody11);
            tableCell11.Append(tableCellProperties5);

            A.TableCell tableCell12 = new A.TableCell();

            A.TextBody textBody12 = new A.TextBody();
            A.BodyProperties bodyProperties6 = new A.BodyProperties();
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph12 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, RightToLeft = false, FontAlignment = A.TextFontAlignmentValues.Center };

            A.Run run16 = new A.Run();

            A.RunProperties runProperties6 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex16 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill16.Append(rgbColorModelHex16);
            A.EffectList effectList6 = new A.EffectList();
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties6.Append(solidFill16);
            runProperties6.Append(effectList6);
            runProperties6.Append(latinFont6);
            A.Text text12 = new A.Text();
            text12.Text = percentage1;

            run16.Append(runProperties6);
            run16.Append(text12);

            paragraph12.Append(paragraphProperties6);
            paragraph12.Append(run16);

            textBody12.Append(bodyProperties6);
            textBody12.Append(listStyle6);
            textBody12.Append(paragraph12);

            A.TableCellProperties tableCellProperties6 = new A.TableCellProperties() { LeftMargin = 9525, RightMargin = 9525, TopMargin = 9525, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties6 = new A.LeftBorderLineProperties();
            A.NoFill noFill16 = new A.NoFill();

            leftBorderLineProperties6.Append(noFill16);

            A.RightBorderLineProperties rightBorderLineProperties6 = new A.RightBorderLineProperties();
            A.NoFill noFill17 = new A.NoFill();

            rightBorderLineProperties6.Append(noFill17);

            A.TopBorderLineProperties topBorderLineProperties6 = new A.TopBorderLineProperties();
            A.NoFill noFill18 = new A.NoFill();

            topBorderLineProperties6.Append(noFill18);

            A.BottomBorderLineProperties bottomBorderLineProperties6 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex17 = new A.RgbColorModelHex() { Val = "1F497D" };

            solidFill17.Append(rgbColorModelHex17);
            A.PresetDash presetDash6 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6 = new A.Round();
            A.HeadEnd headEnd6 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties6.Append(solidFill17);
            bottomBorderLineProperties6.Append(presetDash6);
            bottomBorderLineProperties6.Append(round6);
            bottomBorderLineProperties6.Append(headEnd6);
            bottomBorderLineProperties6.Append(tailEnd6);

            A.SolidFill solidFill18 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex18 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill18.Append(rgbColorModelHex18);

            tableCellProperties6.Append(leftBorderLineProperties6);
            tableCellProperties6.Append(rightBorderLineProperties6);
            tableCellProperties6.Append(topBorderLineProperties6);
            tableCellProperties6.Append(bottomBorderLineProperties6);
            tableCellProperties6.Append(solidFill18);

            tableCell12.Append(textBody12);
            tableCell12.Append(tableCellProperties6);

          

            tableRow8.Append(tableCell10);
            tableRow8.Append(tableCell11);
            tableRow8.Append(tableCell12);
          //  tableRow8.Append(extensionList2);
            table2.Append(tableRow8);
        }



        #region Binary Data
        private string thumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACQAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD7h+M/xok+Ed14fiTRE1ZNUeUOz3wtzEIzGCEXY3muRIcIME7eM1zLftaaMrW8Q8K+IJLqayGofZ0W3LJD9ka7JY+bgERAHGcktjqDXruveLdL8MzWyalci2E6uyu33QqlQxPsC69OmSTwCRXj8feH5NpTU42DZAIVsEjbx065dFx13MF68V2QqUVFKVO773ZjKM27qWnoeSt+1BPcX11NYeEZLrQ7SR0uJnvcXqhLeed9tvHG4L7bdgI3dG3Mobb1E1n+1JbeItastI0LwzeXN5dOIxJdXUEcaloJJYyCrtu3eWQq5Vj1x0z6lN8QvD0Nv57anH5OAxk2ttCltm8nHC7vlLdAeM54oPxC8PBSRqSsqlQzLG7BSyFwCQvBwD17jHXiq9tQ6UvxYuSp/P8AgefeLP2grrwTqn9naj4O1GaVF01pri0fzIIvtEojn3uFwpi3KQCf3hYAEYYjItf2tNM1CTTprbwvrTaZcXL2kt2Y0IjkEhTOFY/JkElzgYK4yx217BeeNND0+6a2uNUt4rhSFaJn+YEkjkduh/AZ6c0yHxzoNxbyTx6nC8MbRq7jOF3/AHM8cA889Bg56GkqtC2tLX1YOFS+kvwPKbH9qKLXNH1bUNL8K3jx2lrY3Fsl5cKr3RuLxrZgEhWV/wB2QC21WOTt2g7d1Oy/bC0NrWBr7w1rVtcs8EcsUKxyYaVN67FZllcEcD92GJBBVSpx61H8SfDMhcDWLcFSoIYkEbkDqcEcAgjn3A6kVbi8ZaRLIifamRnYIBJC68l5EA5HB3RPwem3nHFP2uH1/df+TMXJU/n/AAPJrH9rXQNTubC2tPDuu3VxeRedFHDHC5I84RdpPfPH91h94YOx4L/aDtfGGufZF0C/gsZYfMt7yKSG6DOI3kaJlhdzv2xtwu7n5Tg8V3t1460OykCT36xMVMg3RvgoGCl84xsyw+bpznOKb/wsDw9tDf2pCF4+YhgBxnnjjGRn0zziplUoNNRp/iUoVLq8vwPOfH/7Rlv4OvL2zi08XF/a3cMa6eWJuriB7M3BlEOA67WwnRvm44J4pL+1x4a3APomtRnzLaNlaOLcpmjEgyvmZyoyDx95HAztr1eLxrost3Da/blS4mYLHHIrKWJZlGMj+8pH1wOpAplt480C9miit9TiuJZV3xxxBmZ1BIJAA5AxnPYc9OaFVoJWdP8AH/gCcajekvwPn9P2yLu90vS7mx8I3E9xNHYi4twyuPOmCsyRSI5BxuCENhkYNuBK7T6Fp/7Rmn3vhm/1j/hHtWZdPW1W6htljnxLcKWhjjZWxIWBgORwBcR5x823vT480T+zrO9S6aWK84t1jhcvIS20LtxkHORg46H0NVbfxZ4U0O4vbeKW20yR5XuLgC3MIlkMjRvITtG9tyHLcnAz05rSdbDyWlK3zYlCon8f4Hk8X7Vkt/rkdtp3g+/u7O6mRLN55I7aWUGS3RgEdvmOJy4OV+UDPci9rX7Uttouror+GbyTRZbK2vI7wOyynz4UkRSrRiIEF1UqJi/JOzaCR6zJ400WF5EkvlR43kjYFW4aMneOn8IGT6AgnAINN/4TjQvs11cf2jGIbVFlnfDYjVjhSeOM8H6EHoQaXtsP/wA+vxYclT+f8DxvSf2utPvLiVrnw1qcFs2wQxR+U8vzTTopfMg+dkjjcQgFyGJG8c1LqX7XWk2X2K3HhfXF1G4ktVNpJHGzKs3luD+7dtx8pywAz8wCnGcj2G48aaPbWqXDXimOSOaSPCnMgiBMm3jkgKT744rCtPjV4OuWuEfWIrSS3uVtZEuAVKuyllyRkAEAnJPY+laRdKpeUaDa8m/8iXzxsnP8DhNX/aiitdN8K6pp3hW/1DTdaW4dzvPnQrFc21v8qxJIjFjcZG+SMZTbncwFRW/7YPhaWNZJNJ1W2ha1lu/NlMBQKiyEAlZCAWaMRqDgl3VcV6vYePtD1CztrmO7ZI7iWSCJZYXR2dJWiK7SMglkYAEAnacDg1al8WabDePbPORIucnYcZDbSB688cccGsZVsNT92dKz1+1/wOmxvToV6zvSd/RHlPgb9pq18UeJIdC1Dw9daVfXN5NbWxW4jl8xY4UlEhTIdQ28qPlK5RvmyMVQ0n9qqE+HLXV9V0KOJZo4z9hsb9ZL5S0Hnb2hdUVY3AKxMJCZWaMBcvgewW/jTSLhWP2tYysgjIkGDk5x+Bwf++T6GoovFPh/R7GzjWaHT7IrIkC+UYo0SPAbjACqoOecAAE9BSWIwsm7U/8AyY0qYTFUl+8uvVHlV5+1Pax29848O3OliHR7rVIJdYuYYVneOKKSOFQrMSZFm/3l2nK9ce6Rt5iKwBG4ZwRg1hTeJ/D+q6fdxzzRXVoIn+0QzQsylVba6spXkgkArjPzDjkU+38a6LeXVvb298k8txI0UYjVmDEZyc4xt4I3dCQQDkVnUnTkkoRs/W5lGMlfmdzcorAXx7oDruGpxY8tZjlWG1GOFY5HAPGCeu5f7y5Y3xC8PRyRo+pxL5m/a+DsO1grDdjGRuBxnO07unNc5odFRWFceONBtbuS1l1OGOeMEuhzlQJDGSeOgcFSegwfSmr4+0Blhb+04gszqkbFWAZiMgDj0IJ9Ny5xkZAN+iufl8eaJbybZbwxJ9mW8814nCeUxYBskeq4+pUdWALrvx1olheXFrc3ht5oCFcSxOo3FQ4UErgkgjAHXoOQRQBvUVQ0jXbHXopZLC4W5jicxs8fKhuuM9DwQfoR61foA5rxq2qLBCNL0C111yjb1upEVVwyEJ8x/iwTkZwUGQa5250/Vn+yXI8BaW92I5IpF8yFvKDKgwpJG5SVweBwMdvm9HooA4CS31aSOMN4E014ovLVIzLCWUFyzbR0AUZxyNxbPy811p8P6XcQqsulWm0gfu5IEOMKVA6Y4UlfoSOlaVFAFKTRdOmmaZ7C1eZmDGRoVLEgAA5x1wq/98j0praDpjxtG2nWjRsVYqYFwSowpxjqBwPar9FAGavhrSFXC6VYgegt09MenoT+dPHh/S18nGm2g8khov3CfuyCSCvHGCSeO5NX6KAM1vDWkNM0p0qxMrcM5t03H5g/Jx/eAb6gGoF8G6Et3Lc/2TaNNJsBLxBgNgwuAeFwPQCtmigCkmh6bHcm4TT7Vbgv5hlWFQ5b+9nGc+9RDw1pCyLINKshIowrfZ0yBndwcevP1rm/GXjrWfDfiGx0+w8K3WsWdxCJJNRjaTyoG85I9rBInOdrluBn5egGWHJL8bPF8zsbf4W6s8EdtHPJJLJLEVZo4S6BGg3MVeWRPlBJELMFOcUdLh5Hq39i6dutz9gtc25LQnyV/dEnJK8cZPPFI+h6bIzM+n2rMxLFmhUkktuJ6f3ufrzWH8PfGGo+MtPv7jUvDV/4YeC7aCGDUPvzxAKVmHAwCSw2nkbeeorqqAMnUPCejarn7VpttKS+8t5YBY5J5I5IyTkHg55p/wDwjOjmCSE6TYmGRdjx/Zk2su7dgjHIzz9ea06KAKH9gaYSpOm2mVDhf3C8bwQ+OP4gTn1zUH/CI6F5gk/sXT/MEglDfZY87wMBs46+9a1FUpOOzFZPczrbw3pFnHax2+lWUEdqSbdI7dFEJLbiUAHy5PPHfmpm0mxkkZ2srdnY7mYxKSTnOScdeBVuipfvasuMpQ+F2Kg0mxUgiytwVbcP3S8H16deKZJoOmyxwxvp9qyQsWiVoVIQk5JXjg59KvUUrIbqTkrNmRp/hDRNLsxa22lWscG0ptMQYlS5cgk5JG47ue9WI9B0yOa3mTTrRJrfPkyLAoaPOc7TjjOT09TV+vI/EHxi8YaVr19p9h8MNT1WCG5kt4bxLhkjmwY9r58kgKweQ5yQDHgnnIZB6Ra+FtGs7aC3h0qzjhgULEggXCAZxjj3P5mnL4a0hZEkGlWQkRdisLdMquQ2AccDIB+orya4+OnjOOGZ4vhRq0rp8ywmWZWZCbfBz9mK5AnkJXJI+zyAbjtB9O8D+IL7xT4WsdU1LR5tAvbgMZNOuCxeEh2UAkqp5AB6DrQBfuNB0y6cvPp1pM5GC0kCseSWPUepJ+pNMi8N6TDny9Lso8kE7bdByOh6dsn860aKAKDaBpbrGrabaMsa7UBgXCjngccDk/mafNoun3MkrzWFtK8pDSM8KkuQu0Ekjn5ePpxVyigCG2s4LNWW3gjgVmLsI0CgsepOO9TUUUAFcJ8TvjP4a+FOnSXWtX0UTIOUZwoGegJ9T2ABJ9K6Hxl4gXwv4avtSJUNCnybum48DPtk5+gr8bfjV8V9R+LXjW91K6uZJLFZWWzhYnCpn75H95upP4dAK9jLcveOm7u0VucWKxKw8dN2fbGrf8FKfDltfNHaabc3MGcCSO1JGP8AgUiH/wAdr0z4T/tqeDPiVdJZ+etret/yxYFJP++G+8P90tX5OVLbXM1ncRXFvK8E8TB45I2KsrA5BBHQivqamR4aUbQumeRHMKqd5ao/eG1uor23jngkWaGQbldDkEV8g/8ABTj/AJI/4Y/7Dq/+k81an7D/AMcLr4geFY7DU5d95E7W8p6ZlVQwcDtuU8/7QNSf8FEPBviDxt8K/Dtp4d0LUtfuotaWWSDS7OS5dE8iUbiqAkDJAz7ivmMJSeFzCNOp0Z61aSq4dyj1R+YVFd7/AMKA+KH/AETfxd/4Irr/AON0f8KA+KH/AETfxd/4Irr/AON1+h+2pfzL7z57kn2OCor3+6/ZG1+PSzPFZ+JDdxyzA2p8L35LxrKqxspEWAWQlsZONuDjIqzdfsf6uks3kf8ACSSREM0LP4R1BSAHIUOPL6lQDgdN3XjBw+uUP5i/Y1Ox870V75qf7IviOy0W/uba38Q317EnmW1onhTUUabBIKEmLCsflI5xjNcB/wAKA+KH/RN/F3/giuv/AI3VxxNGe0kS6U47o4Kiu9/4UB8UP+ib+Lv/AARXX/xuj/hQHxQ/6Jv4u/8ABFdf/G609tS/mX3i5J9jP8CXGk26X51bwtdeJIcxZa1neJoFJKkblUgFiy4JB5QDBDEV3Gn2egRq18nwo1eaL7OSokvpGQMokjldgU5UOhx0wyMDngCHUvgT8TPDSW8Gg+F/Gl3Fe2cE18IfD95AI58EmE5T5thP3unJxUdt8PfjjaWMtlF4Q8dLays7SQ/2PdlXZwyuxzHyWDsCepBIrlnOnO7U1r5tfqaxjKOjX4Hklfr1+xV/ybB4F/64T/8ApTLX5veKf2a/HGlyWB0bwj4u1iG4tlll3eG7yGS3kJOY2BQgkccqSOa/S/8AZF0PUvDX7OngzTdX0+60rUbeGYTWd7C0M0ZNxKQGRgCMgg8joRXh55VhUw0eR3979Gd2BjKNV3XQ+Ov+CnH/ACWDwx/2Al/9KJq+PK+6P+CiHwz8YeNvip4du/DvhTXNftYtFWKSfS9OmuUR/PlO0sikA4IOPcV8qf8ACgPih/0Tfxd/4Irr/wCN16eW1accJTTktu5y4qMnWk0jgqK9K0f9nX4j6hqUVvd+AvFljA+4G4fQbohTtJXP7voTge2a7q4/ZF1iLVLRUt/FEmnzOUlaPwpftJB+5dtxzEoYGRVTjn5we1d0sVRi7ORgqU30PnuivoWb9j3W4o4wD4hkkCs0hTwjqJXIPCrmIZ45yccnGMDNY3i/9lPxfoOlW0+k6P4j8QXjGNZra38M38e0srMWDPEMquFX1yc9KSxlCTspD9jU3seJ0V3v/CgPih/0Tfxd/wCCK6/+N0f8KA+KH/RN/F3/AIIrr/43W3tqX8y+8jkn2OCrvfCN1oEfh2FdU8DX2uP9qZm1C1u5Idyqu5oxhSMhefbAPc1PY/s8/Eu5vreGf4feLLWGSRUknbQLthGpIBYgR5OBzgela+rfBb4q6LeX+kaL4U8cXehQ3cjW8q6HeQrMOUEuzZ8pZD09Disp1aUvd519/wDkVGE1rYh8VWOk6f4O1GO3+HGp6Pet5WdUvLp5Vt9pizhSowW3gHJ/5ar04FeVV7Jb/C/4w6/JZaTrnhXx2mhyTQx3Df2DdS+VEpA3KhUZ2r0GR0xkVkeIP2cfiJpms3drp/gjxVq1lG+Ib2Pw/eIJVxkHa0QIPqD3B69amnVpw91zV/W/5jlGUtUj9j9G/wCQPY/9cI//AEEVcqrpKNHpdmjqVdYUBVhgg7RxVqvyp7n1aPPfj1pc+r/C/V7a3OJGTaD6bgVB/NhX4szQvbzPFKjRyRsVZWGCCDgg1+8GoWMOpWU9pcLvhmQo6+xFfm/+1R+x/rGl+JrzXvDVv9pS6cyTWq4XzW6l4z03HqynnOSM5xX1GSYynRlKlUdrnlY+jKolOOtj5v8AD3j6x0PT1tpfDtpqRaBYZTdPlW2yvIGCheCd+05J4UYwenL6teR6jql3dRWyWcU0rSLbx42xgnIUYA4H0FWL7wxrGl3Jt7zSr21nBwY5rd0b8iK7j4e/s/8Aivx5fQ7rCbSNMZhvvb2Ipx/sIcFz6Y49SK+xlOlRTqSdvmeFGM5+4kfRX/BO/SbsX11eBWEE+oIiH/rnExc/k4FfonL5m0eWVB77gTXj/wCzv8G7b4ZeG7RUtmtlih8q3ik++FJyzv8A7bHk17HX5rjsQsTiJVY7M+rw9N0qSgyvi7/vQ/8AfJ/xoxd/3of++T/jViiuA6Dl77wSt9rkeqvdzJOkqzeUrnyiyhQPlOcfcHIweozg1kw/CO0t4vLj1C+Rd24bbhgR9zj35jQ89COMZNd9RQBzPh3waPDNxcTW11JIZx8yTH5AfVVUAL2HA6Aelb2Lv+9D/wB8n/GrFFFxFfF3/eh/75P+NGLv+9D/AN8n/GrFFAznvDnhVvDK3At7qS487ZuNyxcjaMDHTt61o31hLqEIimMJQSJJ909VYMO/Tjn2JrQooA5S18BpZ6tbail5Obi3AVQZDtYBFTDAdflVevpXUx7to34Ld9vSnUUARS+duHlmMDvuBP8AWmYu/wC9D/3yf8asUUAV8Xf96H/vk/41yMnwxt5JL9nvLlkvJWlkjMpwu47mVTjcFJJO3OOR6LjtqKLgcJH8KreONFGoXbbQqjzJA4wCTjayleSSTxznnNb/AId8Ov4ZsTaW1w80W4MDdO0jD5VX7xOT93PPcntxW5RRcVivi7/vQ/8AfJ/xoxd/3of++T/jViigZl6vpc2s6bPZSzCKOZdrPDlWHOeDTdF0ibQ9NisoZhLHGWIebLOdzFjkjHc1rUUCMjUtFk1SaGWWRFkhHyMg6fvI5M85/iiX9apaF4PHh+9muba5kdpk2PHM5ZPvFgQvYjOMjsBnNdJRQAUUUUhhUF5Y2+oW7QXUMc8LdUkUEGp6KAOLuvhF4cupC4tpYM9VimYCtPRPAOheH5RLaWKeevSaUl3H0J6fhXQ0VXM3uwCiiipAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACio7i4jtbeSeZxHFGpd3boABkmvnT4hfFK+8SXcttayNbacpwsSnG4er+p9ug/Ws6lSFGPPN2R6OAy/EZlWVDDRu/wS7s95uPFmiWshjl1ayRxwVM65H15q/Z39tqEfmWtxFcx/3oXDD8xXx01xKxyZG/OtDRvEmoaHeJcWtzJFIv8AEpwf/r/Q158cyoSlZprzPsa/BWOp0+enOMmumq+5v9bH15RXH/Dnx5H4y04iXal/CB5ir0Yf3hXYV6h+fyjKEnGSs0FFFFBIUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQBxnxdvpLLwPd+WcGV0jJHpnJ/lj8a+aLaH7VdxRF9nmOFLntk4zX1Z450NvEXhe+soxmYrvjHqynIH44x+NfKd3avZ3DxOpUqcYYYNePmcZOEJLZN3/A/UOB61KM69Fu05JNeivf7rmwvhXd8xvoUTaCd/DKT/CR2IwcjPBFZ+qaS+l+VvmhmEm7aYW3DAOM5x3qjRXz7atoj9VjCpF3lO69Duvg/qUlj4wsVUnbK5iYeoZT/AFAr6Wr58+CvhyW+8Sx3bIfJswZXbtuIIVfryT+FfQdfY4eMoUYRlukfzvnlaliMyr1KPwt/krN/N3YUUUVueEFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFefePPhNa+KJHvLNltb1uXVh8kh9eOh969Boo6WZdOpOlJTpuzWzWjPmu8+DviG1mKrYSSjPDRujA/qK2PDvwP1a6mVr0Jp8PdnYPJ+Cjj8zXvlFYRoUYPmjBJnsVs7zLEU/ZVa8nH7vvtuZvh/wAPWXhnTY7Kxj2RryzHlnbuxPc1pUUVueIFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAHD/Fbxl4T8DaTY6h4u8Qt4cs2uBFBcLK8e6XBbblQeqqw+hb1rzK4/aL+C9zM8j/ABXnVmlaXEd9cIoBOdgAXG0dB3967b49fA+P446LodkdcuPD8+kakmpwXVtCsreYisq8McfxZ/CvIZf2F7qa6+0t8TdQE+4OZF0a0UkhVUE4HPyog5/uivUoU8HKmnWm1L+vJnJUlXUrQjdf15nTQ/tDfBaGZpP+Fr3DZcvta+nIHB+X7nTnOPYUrftGfBtmVm+Lc5x2F3Ko7+ie5/T0FckP2C51gaEfFDVRGw2kDTLf+8W6+uSeffFSf8MJ3n2iSf8A4Wnq4lkGGb+zbfnp/wDEj9fU56PY5b/z8f8AX/bplz4r+Vf18z3DwPN4d+IfhMah4c8V6prGkTTuUvob6TdvBwyhiA2AR908e1bZ8DwtHKravrDGWVJmf7awIZQQMY6A5GQODtHpXO/A/wCCth8Ffhza+EIr1tbtoJ5ZxPdQqpJdt2NoyOK7r+xtP/58bb/vyv8AhXk1FTU5Km7x6eh2xcuVcy1MdvA8LY3axrRG0rgajIvXPPBHPPXtgVNqXg+DU94bUNUt1ZmYiC+kUgnPIOcr14AIAwK0v7G0/wD58bb/AL8r/hR/Y2n/APPjbf8Aflf8Ky0K1MF/h7A7FjrOtF2cszG+bJBABXjoMDHHTJPXmpf+EEt9gjTVdajiBJCLqMuBkk4znOB0Az0AFbP9jaf/AM+Nt/35X/Cj+xtP/wCfG2/78r/hRoBi2fgWK3W6E+r6xe+fx++1CUbBkkAAMAOuPfAzUo8GRBkP9ravhURApvWIwqBQcdzkbiepbn2rV/sbT/8Anxtv+/K/4Uf2Np//AD423/flf8KNA1Me18DxWuMaxrUmPM+/fuclwwJ/DcSPQgegp0fgqGOB4v7W1ltxz5jahIWHzE8EnjrjjsAOla39jaf/AM+Nt/35X/Cj+xtP/wCfG2/78r/hRoGpjr4HhW8iuG1bWJDGwYRtet5Z+bcQygchj1z2GOKks/BlvZo6f2hqdxG0TwtHc3jyqVbH8LZGRt4PUZPrWp/Y2n/8+Nt/35X/AAo/sbT/APnxtv8Avyv+FGgGL/wgFi0MET3uqSrFhf3l9IxZecqST0OecYPHXrmS48Fx3E3mNrGsKfNEu1b1gvTG3HTb7e9a39jaf/z423/flf8ACj+xtP8A+fG2/wC/K/4UaBqZSeC4Vkt3/tXV2aHABa+c7uSfm/vcnv6AdKhm8BxXGd2t64FO/KrqDjO7ODx025OMY7egrb/sbT/+fG2/78r/AIUf2Np//Pjbf9+V/wAKNAMn/hC4tluF1jWUMAYBhfud2WLfMDkNjOBkdOK09F0ddEtGt0urq6QuXDXcpkZc9gT29u2af/Y2n/8APjbf9+V/wo/sbT/+fG2/78r/AIUaAXKKp/2Np/8Az423/flf8KP7G0//AJ8bb/vyv+FGgalyiqf9jaf/AM+Nt/35X/Cj+xtP/wCfG2/78r/hRoGpcoqn/Y2n/wDPjbf9+V/wo/sbT/8Anxtv+/K/4UaBqXKKp/2Np/8Az423/flf8Knt7WC0UiCGOEHkiNQufyo0GS0VDJaiS5hm8yRTEGGxXIRs4+8O+McelTUgP//Z";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }


    public class clsQTC
        {
            private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
            private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
            private PresentationDocument document;

            public void ChangePackage(string filePath)
            {
                using (document = PresentationDocument.Open(filePath, true))
                {
                    ChangeParts();
                }
            }

            private void ChangeParts()
            {
                //Stores the referrences to all the parts in a dictionary.
                BuildUriPartDictionary();
                //Changes the contents of the specified parts.
                ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
                ChangeThumbnailPart1(((ThumbnailPart)UriPartDictionary["/docProps/thumbnail.jpeg"]));
                ChangeExtendedFilePropertiesPart1(((ExtendedFilePropertiesPart)UriPartDictionary["/docProps/app.xml"]));
               // ChangeSlidePart1(((SlidePart)UriPartDictionary["/ppt/slides/slide1.xml"]));
            }

            /// <summary>
            /// Stores the references to all the parts in the package.
            /// They could be retrieved by their URIs later.
            /// </summary>
            private void BuildUriPartDictionary()
            {
                System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
                queue.Enqueue(document);
                while (queue.Count > 0)
                {
                    foreach (var part in queue.Dequeue().Parts)
                    {
                        if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                        {
                            UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                            queue.Enqueue(part.OpenXmlPart);
                        }
                    }
                }
            }

            private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
            {
                var package = coreFilePropertiesPart1.OpenXmlPackage;
                package.PackageProperties.Revision = "1117";
                package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2024-11-19T18:32:41Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            }

            private void ChangeThumbnailPart1(ThumbnailPart thumbnailPart1)
            {
                //System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
                //thumbnailPart1.FeedData(data);
                //data.Close();
            }

            private void ChangeExtendedFilePropertiesPart1(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
            {
                Ap.Properties properties1 = extendedFilePropertiesPart1.Properties;

                Ap.Words words1 = properties1.GetFirstChild<Ap.Words>();
                Ap.Paragraphs paragraphs1 = properties1.GetFirstChild<Ap.Paragraphs>();
                words1.Text = "33";

                paragraphs1.Text = "18";

            }

       
            public  void ChangeSlidePartFillACell(SlidePart slidePart1,string celltext, int row, int cell, string color)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.GetFirstChild<GraphicFrame>();

            A.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<A.NonVisualGraphicFrameProperties>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
            modificationId1.Val = (UInt32Value)4184461266U;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            A.TableRow tableRow1 = table1.Elements<A.TableRow>().ElementAt(row);

            A.TableCell tableCell1 = tableRow1.Elements<A.TableCell>().ElementAt(cell);

            A.TextBody textBody1 = tableCell1.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.EndParagraphRunProperties endParagraphRunProperties1 = paragraph1.GetFirstChild<A.EndParagraphRunProperties>();

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", FontSize = 1400, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FF0000" };

            solidFill1.Append(rgbColorModelHex1);
            A.EffectList effectList1 = new A.EffectList();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties1.Append(solidFill1);
            runProperties1.Append(effectList1);
            runProperties1.Append(latinFont1);
            A.Text text1 = new A.Text();
            text1.Text = celltext;

            run1.Append(runProperties1);
            run1.Append(text1);
            paragraph1.InsertBefore(run1, endParagraphRunProperties1);

           // endParagraphRunProperties1.Remove();
        }





        private System.IO.Stream GetBinaryDataStream(string base64String)
            {
                return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
            }


        }
   


    public class clsFitToTheraCategory
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private PresentationDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = PresentationDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            //Changes the contents of the specified parts.
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeThumbnailPart1(((ThumbnailPart)UriPartDictionary["/docProps/thumbnail.jpeg"]));
            ChangeExtendedFilePropertiesPart1(((ExtendedFilePropertiesPart)UriPartDictionary["/docProps/app.xml"]));
            //ChangeSlideFitToCategory(((SlidePart)UriPartDictionary["/ppt/slides/slide1.xml"]));
        }

        /// <summary>
        /// Stores the references to all the parts in the package.
        /// They could be retrieved by their URIs later.
        /// </summary>
        private void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(document);
            while (queue.Count > 0)
            {
                foreach (var part in queue.Dequeue().Parts)
                {
                    if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                    {
                        UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                        queue.Enqueue(part.OpenXmlPart);
                    }
                }
            }
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2024-11-26T15:26:41Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeThumbnailPart1(ThumbnailPart thumbnailPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        private void ChangeExtendedFilePropertiesPart1(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = extendedFilePropertiesPart1.Properties;

            Ap.Words words1 = properties1.GetFirstChild<Ap.Words>();
            Ap.Paragraphs paragraphs1 = properties1.GetFirstChild<Ap.Paragraphs>();
            words1.Text = "90";

            paragraphs1.Text = "36";

        }

        public void ChangeSlideFitToCategory(SlidePart slidePart1,string text,int row,int cell,string color,int totalRows, bool isFirstTimeLoop)
        {
           
          

            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.GetFirstChild<GraphicFrame>();

            //NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<NonVisualGraphicFrameProperties>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            //ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            //ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            //ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            //P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
            //modificationId1.Val = (UInt32Value)645911878U;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            //number of rows to be deleted

            var delRowsCount = 14 - totalRows-1;


            //delete the extra rows

            if (isFirstTimeLoop  )
            {
                for (int i = 0; i < delRowsCount; i++)
                {
                    if (table1 != null && table1.Elements<TableRow>().Any() && table1.Elements<TableRow>().Count()> totalRows)
                    {
                        // Get the last row
                        TableRow lastRow = table1.Elements<TableRow>().Last();

                        // Remove the last row
                        lastRow.Remove();
                    }

                }
            }


            A.TableRow tableRow1 = table1.GetFirstChild<A.TableRow>();
            A.TableRow tableRow2 = table1.Elements<A.TableRow>().ElementAt(row);

            A.TableCell tableCell1 = tableRow1.Elements<A.TableCell>().ElementAt(3);

            A.TextBody textBody1 = tableCell1.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.Run run1 = paragraph1.GetFirstChild<A.Run>();
            A.EndParagraphRunProperties endParagraphRunProperties1 = paragraph1.GetFirstChild<A.EndParagraphRunProperties>();

            //A.RunProperties runProperties1 = run1.GetFirstChild<A.RunProperties>();
            //runProperties1.Dirty = false;

            //endParagraphRunProperties1.Remove();

            A.TableCell tableCell2 = tableRow2.Elements<A.TableCell>().ElementAt(cell); ;

            A.TextBody textBody2 = tableCell2.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph2 = textBody2.GetFirstChild<A.Paragraph>();

            A.EndParagraphRunProperties endParagraphRunProperties2 = paragraph2.GetFirstChild<A.EndParagraphRunProperties>();

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };


            //setting the color of the cell
            //orange : "FF9933" ,baby blue: "CCFFFF",pale violet :CC99FF 

            if (tableCell2 != null && color!="")
            {
                // Set the fill color of the cell
                tableCell2.TableCellProperties = new TableCellProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = color } // Red color
                    )
                );
            }



            

           




            solidFill1.Append(schemeColor1);
            A.EffectList effectList1 = new A.EffectList();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            runProperties2.Append(solidFill1);
            runProperties2.Append(effectList1);
            runProperties2.Append(latinFont1);
            runProperties2.Append(eastAsianFont1);
            runProperties2.Append(complexScriptFont1);
            A.Text text1 = new A.Text();
            text1.Text = "Ram";


           


            run2.Append(runProperties2);
            run2.Append(text1);
            paragraph2.InsertBefore(run2, endParagraphRunProperties2);


        }



        #region Binary Data
        private string thumbnailPart1Data = "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C13OXZxtpr5eh+cnif4PeNLj4f3hj8KeJri9uNB0qAWa6TdBlljl/eLt8sj5R2BGAemOK4lvgp8QV+EcdkPAvis3za405tRpFzkRiALvK+XnqcA+xr9YPD/hh/DlvNBb38k0chDEXO6TacYJX5uM9cDjPYVin4XxtaiA6zfna7SLKJCsgZgoJ3AjP3F4Oa9+nnUkrOK3v1PnK+Dq8y9mrq39dR3wPsLrSvgr4Asr22ms7y28P6fDPb3CFJIpFtowyMp5VgQQQeQRXbUighVDEFsckDApa+anLmk5dz3YJqKTCiiioLCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiqGsf2n9nU6UbTzwwyt4G2le+CpyCOvQ5xjjORjSSeNMZSDQSd5+Vp5x8ueOdnUDPbn260AdRRXPSSeKkku8R6O0WSbdmklU47bxtPbuD37Y5csnigsx8rRyn8OJZcsMHr8vGDj1zz0oA36K5OO68arNIstho7xou9ZIp5B5nH3QCODnueMEH1A6qPf5aeZt8zA3beme+PagB1FFFABRRRQAUUE45PAoBDDIORQAUUUUAFFFFABRRRQAUUUUAFFIzBcZIGeBmloAKKKKACiiigAooooAKKKKACiiigDmPH3xN8L/C3TLfUfFes2+i2VxN9nimuA2Gk2ltowDzhSfwrhP+GwPg3/0P2m/98yf/ABFafx6+B8fxx0XQ7I65ceH59I1JNTguraFZW8xFZV4Y4/iz+FeQy/sL3U119pb4m6gJ9wcyLo1opJCqoJwOflRBz/dFepQp4OVNOtNqX9eTOSpKupWhG6/rzPQdS/a3+D13atHH8QNKWTIZfMSVlJBzgjb0rk779oD4T6pBYs/xZsrC8tnnkWSyjlVV8wg4UEcAAbfXDNggkEYg/YLnWBoR8UNVEbDaQNMt/wC8W6+uSeffFSf8MJ3n2iSf/haeriWQYZv7Nt+en/xI/X1Oen2eXWt7R/1/26YXxPNzciv/AF5nr3hG10j4oeD/AO0dC8aXmt2Mt480GpwSODG6q0ZRMnIA3fXOTnpjpb/wPcahBqMD6/qCRXZJASVgYwZQ5VTu6YBUAY4Yg5GAMX4H/BWw+Cvw5tfCEV62t20E8s4nuoVUku27G0ZHFd1/Y2n/APPjbf8Aflf8K8ioqam1Td49PQ74uXKuZanKXXgDVWs7iO28XanHMxBikkYtsAB+U885O3nrhfUklf8AhX+ptavbyeLNSkR3jcsx+YbB90HP3Tk5BznC5zhi3Vf2Np//AD423/flf8KP7G0//nxtv+/K/wCFZaFanPX3gvULyO2RfEt/biKNEcxswaUqW+Ynd/FuGcenGOMUNN+H+t2+os954s1C6tEaPy4xKytIoX5w5BwMnPK+vsMdh/Y2n/8APjbf9+V/wo/sbT/+fG2/78r/AIUaD1Mq48NXN94N/se71CSa8eELLeN1Z8hicf3c8Y9OKzbPwDcw+FLfSn1aSK4huPPS6hTJj68ICcdCeWBweQAQCOn/ALG0/wD58bb/AL8r/hR/Y2n/APPjbf8Aflf8K09o+T2d9L3+Y+Z2sc1H4D1D93v8Vap8kLR7o3IJcg4c5JBIJHGMHaAQfm3RH4d3izQSQeJdQt2RNsixsxWX97JJzlyf+WhHXOAMk9K6r+xtP/58bb/vyv8AhR/Y2n/8+Nt/35X/AArPQnU5+z8CSW13YXUut31zc2p5eSRv3ql3cq3zdDvAwOPkTjAxUKfD+6OmxWk3iTVJmjIImMzh+DGQSd3JBj75GGbjJJPTf2Np/wDz423/AH5X/Cj+xtP/AOfG2/78r/hRoBycvw/1f7NBHH4v1LzFfMkkhOXXK8cEYOAfbJ9MqXR/D/U1jZZfFmpXB2BVMhxghlJY4YcnaBkYIDNtwSCOq/sbT/8Anxtv+/K/4Uf2Np//AD423/flf8KNA1Oc1TwPqWoXxnj8UahZxjzCsMJIXLdC3zchRwAMcAH72WNC28A6+ttMZvFd6bvc7QskjbFJztLAkghc524wcYJIxjsv7G0//nxtv+/K/wCFH9jaf/z423/flf8ACjQNTmPHXgGbxZcadNb6g1m9px82W4yDuHP3uOtWNW8DS32pfbrXWbrTrgjDtCMh8KgGQTg42Z5Hc1v/ANjaf/z423/flf8ACj+xtP8A+fG2/wC/K/4VlCjThOU4rWW/nZW/I1lVnOMYSekdvzOWuvAOpSO72/ii+tXaaeTcoLHZIsaqnL/w+WOf9o42nmnz+Ab260+SGXxHfPc+e00V1lg0eQBgANjAAPTH3mxium/sbT/+fG2/78r/AIUf2Np//Pjbf9+V/wAK10MtTn7zwXfXkOowt4hvViunkZQpYGJWZjtB3ZwoYgYx75AULUuPAuvf2fcxweL703LHMUsycJ8hG3CkZBJU56/LxySa6v8AsbT/APnxtv8Avyv+FH9jaf8A8+Nt/wB+V/wo0A5aLwLrSqyS+ML6dGyfmiAIJUqOQegznBzkhc55zPdeDdWuGstvie6gWDyzJ5aNmbauMEl+AT8x4JOSCSMAdF/Y2n/8+Nt/35X/AAo/sbT/APnxtv8Avyv+FGganHWHgXxFHdSNd+LLue3UoESMlWkUKoIJ/hyQeRk89fXstHt7m00u1hvbn7ZeJGonuNu0SSY+ZgOwJzgdhxSf2Np//Pjbf9+V/wAKnt7WC0UiCGOEHkiNQufyo0AloqGS1ElzDN5kimIMNiuQjZx94d8Y49KmpDP/2Q==";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion





    }



    public class clsFitToProduct
    {

        public void ChangeSlideFitToCategory(SlidePart slidePart1, string text, int row, int cell, string color, int totalRows, bool isFirstTimeLoop)
        {



            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.GetFirstChild<GraphicFrame>();

            //NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<NonVisualGraphicFrameProperties>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            //ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            //ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            //ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            //P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
            //modificationId1.Val = (UInt32Value)645911878U;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            //number of rows to be deleted

            var delRowsCount = 14 - totalRows - 1;


            //delete the extra rows

            if (isFirstTimeLoop)
            {
                for (int i = 0; i < delRowsCount; i++)
                {
                    if (table1 != null && table1.Elements<TableRow>().Any() && table1.Elements<TableRow>().Count() > totalRows)
                    {
                        // Get the last row
                        TableRow lastRow = table1.Elements<TableRow>().Last();

                        // Remove the last row
                        lastRow.Remove();
                    }

                }
            }


            A.TableRow tableRow1 = table1.GetFirstChild<A.TableRow>();
            A.TableRow tableRow2 = table1.Elements<A.TableRow>().ElementAt(row);

            A.TableCell tableCell1 = tableRow1.Elements<A.TableCell>().ElementAt(3);

            A.TextBody textBody1 = tableCell1.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.Run run1 = paragraph1.GetFirstChild<A.Run>();
            A.EndParagraphRunProperties endParagraphRunProperties1 = paragraph1.GetFirstChild<A.EndParagraphRunProperties>();

            //A.RunProperties runProperties1 = run1.GetFirstChild<A.RunProperties>();
            //runProperties1.Dirty = false;

            //endParagraphRunProperties1.Remove();

            A.TableCell tableCell2 = tableRow2.Elements<A.TableCell>().ElementAt(cell); ;

            A.TextBody textBody2 = tableCell2.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph2 = textBody2.GetFirstChild<A.Paragraph>();

            A.EndParagraphRunProperties endParagraphRunProperties2 = paragraph2.GetFirstChild<A.EndParagraphRunProperties>();

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };


            //setting the color of the cell
            //orange : "FF9933" ,baby blue: "CCFFFF",pale violet :CC99FF 

            if (tableCell2 != null && color != "")
            {
                // Set the fill color of the cell
                tableCell2.TableCellProperties = new TableCellProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = color } // Red color
                    )
                );
            }


            solidFill1.Append(schemeColor1);
            A.EffectList effectList1 = new A.EffectList();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            runProperties2.Append(solidFill1);
            runProperties2.Append(effectList1);
            runProperties2.Append(latinFont1);
            runProperties2.Append(eastAsianFont1);
            runProperties2.Append(complexScriptFont1);
            A.Text text1 = new A.Text();
            text1.Text = text;





            run2.Append(runProperties2);
            run2.Append(text1);
            paragraph2.InsertBefore(run2, endParagraphRunProperties2);


        }


    }


    public class clsSALA
    {



        public void ChangeSlidePart1(SlidePart slidePart1)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.Elements<GraphicFrame>().ElementAt(5);

            //NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<NonVisualGraphicFrameProperties>();
            Transform transform1 = graphicFrame1.GetFirstChild<Transform>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            //NonVisualDrawingProperties nonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<NonVisualDrawingProperties>();
            //ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();
            //nonVisualDrawingProperties1.Id = (UInt32Value)4U;
            //nonVisualDrawingProperties1.Name = "Table 3";

            //A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = nonVisualDrawingProperties1.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();

            //A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = nonVisualDrawingPropertiesExtensionList1.GetFirstChild<A.NonVisualDrawingPropertiesExtension>();

            //OpenXmlUnknownElement openXmlUnknownElement1 = nonVisualDrawingPropertiesExtension1.GetFirstChild<OpenXmlUnknownElement>();

            //OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1C303C05-569C-5F63-A762-7ECB46C07B7B}\" />");
            //nonVisualDrawingPropertiesExtension1.InsertBefore(openXmlUnknownElement2, openXmlUnknownElement1);

            //openXmlUnknownElement1.Remove();

            //ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            //ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            //P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
           // modificationId1.Val = (UInt32Value)3503359290U;

            A.Offset offset1 = transform1.GetFirstChild<A.Offset>();
            A.Extents extents1 = transform1.GetFirstChild<A.Extents>();
            offset1.X = 2028825L;
            offset1.Y = 866974L;
            extents1.Cy = 605790L;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            A.TableGrid tableGrid1 = table1.GetFirstChild<A.TableGrid>();
            A.TableRow tableRow1 = table1.GetFirstChild<A.TableRow>();
            A.TableRow tableRow2 = table1.Elements<A.TableRow>().ElementAt(1);

            A.GridColumn gridColumn1 = tableGrid1.GetFirstChild<A.GridColumn>();
            gridColumn1.Width = 1268464L;

            A.ExtensionList extensionList1 = gridColumn1.GetFirstChild<A.ExtensionList>();

            A.Extension extension1 = extensionList1.GetFirstChild<A.Extension>();

            OpenXmlUnknownElement openXmlUnknownElement3 = extension1.GetFirstChild<OpenXmlUnknownElement>();

            //OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"1867896359\" />");
            //extension1.InsertBefore(openXmlUnknownElement4, openXmlUnknownElement3);

            openXmlUnknownElement3.Remove();

            A.GridColumn gridColumn2 = new A.GridColumn() { Width = 1268464L };

            A.ExtensionList extensionList2 = new A.ExtensionList();

            A.Extension extension2 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            //OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2940081160\" />");

            //extension2.Append(openXmlUnknownElement5);

            extensionList2.Append(extension2);

            gridColumn2.Append(extensionList2);
            tableGrid1.Append(gridColumn2);

            A.TableCell tableCell1 = tableRow1.GetFirstChild<A.TableCell>();
            A.ExtensionList extensionList3 = tableRow1.GetFirstChild<A.ExtensionList>();
            tableCell1.GridSpan = 2;

            A.TextBody textBody1 = tableCell1.GetFirstChild<A.TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.Run run1 = paragraph1.GetFirstChild<A.Run>();
            A.EndParagraphRunProperties endParagraphRunProperties1 = paragraph1.GetFirstChild<A.EndParagraphRunProperties>();

            A.RunProperties runProperties1 = run1.GetFirstChild<A.RunProperties>();
            runProperties1.Dirty = false;

            endParagraphRunProperties1.Remove();

            A.TableCell tableCell2 = new A.TableCell() { HorizontalMerge = true };

            A.TextBody textBody2 = new A.TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties1);
            textBody2.Append(listStyle1);
            textBody2.Append(paragraph2);
            A.TableCellProperties tableCellProperties1 = new A.TableCellProperties();

            tableCell2.Append(textBody2);
            tableCell2.Append(tableCellProperties1);
            tableRow1.InsertBefore(tableCell2, extensionList3);

            A.Extension extension3 = extensionList3.GetFirstChild<A.Extension>();

            OpenXmlUnknownElement openXmlUnknownElement6 = extension3.GetFirstChild<OpenXmlUnknownElement>();

            //OpenXmlUnknownElement openXmlUnknownElement7 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3973135540\" />");
            //extension3.InsertBefore(openXmlUnknownElement7, openXmlUnknownElement6);

           // openXmlUnknownElement6.Remove();

            A.TableCell tableCell3 = tableRow2.GetFirstChild<A.TableCell>();
            A.ExtensionList extensionList4 = tableRow2.GetFirstChild<A.ExtensionList>();
            tableCell3.GridSpan = 2;

            A.TableCell tableCell4 = new A.TableCell() { HorizontalMerge = true };

            A.TextBody textBody3 = new A.TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties2);
            textBody3.Append(listStyle2);
            textBody3.Append(paragraph3);
            A.TableCellProperties tableCellProperties2 = new A.TableCellProperties();

            tableCell4.Append(textBody3);
            tableCell4.Append(tableCellProperties2);
            tableRow2.InsertBefore(tableCell4, extensionList4);

            A.Extension extension4 = extensionList4.GetFirstChild<A.Extension>();

            OpenXmlUnknownElement openXmlUnknownElement8 = extension4.GetFirstChild<OpenXmlUnknownElement>();

            //OpenXmlUnknownElement openXmlUnknownElement9 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2425369444\" />");
            //extension4.InsertBefore(openXmlUnknownElement9, openXmlUnknownElement8);

            openXmlUnknownElement8.Remove();

            A.TableRow tableRow3 = new A.TableRow() { Height = 95250L };

            A.TableCell tableCell5 = new A.TableCell();

            A.TextBody textBody4 = new A.TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, FontAlignment = A.TextFontAlignmentValues.Bottom };

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "es-ES", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill1.Append(rgbColorModelHex1);
            A.EffectList effectList1 = new A.EffectList();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties2.Append(solidFill1);
            runProperties2.Append(effectList1);
            runProperties2.Append(latinFont1);
            A.Text text1 = new A.Text();
            text1.Text = "EMCYT";

            run2.Append(runProperties2);
            run2.Append(text1);

            paragraph4.Append(paragraphProperties1);
            paragraph4.Append(run2);

            textBody4.Append(bodyProperties3);
            textBody4.Append(listStyle3);
            textBody4.Append(paragraph4);

            A.TableCellProperties tableCellProperties3 = new A.TableCellProperties() { LeftMargin = 7620, RightMargin = 7620, TopMargin = 7620, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties1 = new A.LeftBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill2.Append(rgbColorModelHex2);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd1 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            leftBorderLineProperties1.Append(solidFill2);
            leftBorderLineProperties1.Append(presetDash1);
            leftBorderLineProperties1.Append(round1);
            leftBorderLineProperties1.Append(headEnd1);
            leftBorderLineProperties1.Append(tailEnd1);

            A.RightBorderLineProperties rightBorderLineProperties1 = new A.RightBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill3.Append(rgbColorModelHex3);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            rightBorderLineProperties1.Append(solidFill3);
            rightBorderLineProperties1.Append(presetDash2);
            rightBorderLineProperties1.Append(round2);
            rightBorderLineProperties1.Append(headEnd2);
            rightBorderLineProperties1.Append(tailEnd2);

            A.TopBorderLineProperties topBorderLineProperties1 = new A.TopBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill4.Append(rgbColorModelHex4);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3 = new A.Round();
            A.HeadEnd headEnd3 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties1.Append(solidFill4);
            topBorderLineProperties1.Append(presetDash3);
            topBorderLineProperties1.Append(round3);
            topBorderLineProperties1.Append(headEnd3);
            topBorderLineProperties1.Append(tailEnd3);

            A.BottomBorderLineProperties bottomBorderLineProperties1 = new A.BottomBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill5.Append(rgbColorModelHex5);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round4 = new A.Round();
            A.HeadEnd headEnd4 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd4 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties1.Append(solidFill5);
            bottomBorderLineProperties1.Append(presetDash4);
            bottomBorderLineProperties1.Append(round4);
            bottomBorderLineProperties1.Append(headEnd4);
            bottomBorderLineProperties1.Append(tailEnd4);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties1 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill1 = new A.NoFill();
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties1.Append(noFill1);
            topLeftToBottomRightBorderLineProperties1.Append(presetDash5);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties1 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill2 = new A.NoFill();
            A.PresetDash presetDash6 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties1.Append(noFill2);
            bottomLeftToTopRightBorderLineProperties1.Append(presetDash6);

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill6.Append(schemeColor1);

            tableCellProperties3.Append(leftBorderLineProperties1);
            tableCellProperties3.Append(rightBorderLineProperties1);
            tableCellProperties3.Append(topBorderLineProperties1);
            tableCellProperties3.Append(bottomBorderLineProperties1);
            tableCellProperties3.Append(topLeftToBottomRightBorderLineProperties1);
            tableCellProperties3.Append(bottomLeftToTopRightBorderLineProperties1);
            tableCellProperties3.Append(solidFill6);

            tableCell5.Append(textBody4);
            tableCell5.Append(tableCellProperties3);

            A.TableCell tableCell6 = new A.TableCell();

            A.TextBody textBody5 = new A.TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, FontAlignment = A.TextFontAlignmentValues.Bottom };

            A.Run run3 = new A.Run();

            A.RunProperties runProperties3 = new A.RunProperties() { Language = "es-ES", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill7.Append(rgbColorModelHex6);
            A.EffectList effectList2 = new A.EffectList();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties3.Append(solidFill7);
            runProperties3.Append(effectList2);
            runProperties3.Append(latinFont2);
            A.Text text2 = new A.Text();
            text2.Text = "EMGALITY";

            run3.Append(runProperties3);
            run3.Append(text2);

            paragraph5.Append(paragraphProperties2);
            paragraph5.Append(run3);

            textBody5.Append(bodyProperties4);
            textBody5.Append(listStyle4);
            textBody5.Append(paragraph5);

            A.TableCellProperties tableCellProperties4 = new A.TableCellProperties() { LeftMargin = 7620, RightMargin = 7620, TopMargin = 7620, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties2 = new A.LeftBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill8.Append(rgbColorModelHex7);
            A.PresetDash presetDash7 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round5 = new A.Round();
            A.HeadEnd headEnd5 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd5 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            leftBorderLineProperties2.Append(solidFill8);
            leftBorderLineProperties2.Append(presetDash7);
            leftBorderLineProperties2.Append(round5);
            leftBorderLineProperties2.Append(headEnd5);
            leftBorderLineProperties2.Append(tailEnd5);

            A.RightBorderLineProperties rightBorderLineProperties2 = new A.RightBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill9.Append(rgbColorModelHex8);
            A.PresetDash presetDash8 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6 = new A.Round();
            A.HeadEnd headEnd6 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            rightBorderLineProperties2.Append(solidFill9);
            rightBorderLineProperties2.Append(presetDash8);
            rightBorderLineProperties2.Append(round6);
            rightBorderLineProperties2.Append(headEnd6);
            rightBorderLineProperties2.Append(tailEnd6);

            A.TopBorderLineProperties topBorderLineProperties2 = new A.TopBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill10.Append(rgbColorModelHex9);
            A.PresetDash presetDash9 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round7 = new A.Round();
            A.HeadEnd headEnd7 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd7 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties2.Append(solidFill10);
            topBorderLineProperties2.Append(presetDash9);
            topBorderLineProperties2.Append(round7);
            topBorderLineProperties2.Append(headEnd7);
            topBorderLineProperties2.Append(tailEnd7);

            A.BottomBorderLineProperties bottomBorderLineProperties2 = new A.BottomBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill11.Append(rgbColorModelHex10);
            A.PresetDash presetDash10 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round8 = new A.Round();
            A.HeadEnd headEnd8 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd8 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties2.Append(solidFill11);
            bottomBorderLineProperties2.Append(presetDash10);
            bottomBorderLineProperties2.Append(round8);
            bottomBorderLineProperties2.Append(headEnd8);
            bottomBorderLineProperties2.Append(tailEnd8);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties2 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill3 = new A.NoFill();
            A.PresetDash presetDash11 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties2.Append(noFill3);
            topLeftToBottomRightBorderLineProperties2.Append(presetDash11);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties2 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill4 = new A.NoFill();
            A.PresetDash presetDash12 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties2.Append(noFill4);
            bottomLeftToTopRightBorderLineProperties2.Append(presetDash12);

            A.SolidFill solidFill12 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill12.Append(schemeColor2);

            tableCellProperties4.Append(leftBorderLineProperties2);
            tableCellProperties4.Append(rightBorderLineProperties2);
            tableCellProperties4.Append(topBorderLineProperties2);
            tableCellProperties4.Append(bottomBorderLineProperties2);
            tableCellProperties4.Append(topLeftToBottomRightBorderLineProperties2);
            tableCellProperties4.Append(bottomLeftToTopRightBorderLineProperties2);
            tableCellProperties4.Append(solidFill12);

            tableCell6.Append(textBody5);
            tableCell6.Append(tableCellProperties4);

            A.ExtensionList extensionList5 = new A.ExtensionList();

            A.Extension extension5 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };



            extensionList5.Append(extension5);

            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);            //OpenXmlUnknownElement openXmlUnknownElement10 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2428161736\" />");

            //extension5.Append(openXmlUnknownElement10);
            tableRow3.Append(extensionList5);
            table1.Append(tableRow3);
        }

        public void ChangeSlidePart2(SlidePart slidePart1)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            GraphicFrame graphicFrame1 = shapeTree1.Elements<GraphicFrame>().ElementAt(5);

            //NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = graphicFrame1.GetFirstChild<NonVisualGraphicFrameProperties>();
            Transform transform1 = graphicFrame1.GetFirstChild<Transform>();
            A.Graphic graphic1 = graphicFrame1.GetFirstChild<A.Graphic>();

            //ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.GetFirstChild<ApplicationNonVisualDrawingProperties>();

            //ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = applicationNonVisualDrawingProperties1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtensionList>();

            //ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = applicationNonVisualDrawingPropertiesExtensionList1.GetFirstChild<ApplicationNonVisualDrawingPropertiesExtension>();

            //P14.ModificationId modificationId1 = applicationNonVisualDrawingPropertiesExtension1.GetFirstChild<P14.ModificationId>();
            //modificationId1.Val = (UInt32Value)3503359290U;

            A.Extents extents1 = transform1.GetFirstChild<A.Extents>();
            extents1.Cy = 605790L;

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            A.Table table1 = graphicData1.GetFirstChild<A.Table>();

            A.TableGrid tableGrid1 = table1.GetFirstChild<A.TableGrid>();
            A.TableRow tableRow1 = table1.GetFirstChild<A.TableRow>();
            A.TableRow tableRow2 = table1.Elements<A.TableRow>().ElementAt(1);

            A.GridColumn gridColumn1 = tableGrid1.GetFirstChild<A.GridColumn>();
            gridColumn1.Width = 1268464L;

            A.GridColumn gridColumn2 = new A.GridColumn() { Width = 1268464L };

            A.ExtensionList extensionList1 = new A.ExtensionList();

            A.Extension extension1 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            //OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2940081160\" />");

            //extension1.Append(openXmlUnknownElement1);

            extensionList1.Append(extension1);

            gridColumn2.Append(extensionList1);
            tableGrid1.Append(gridColumn2);

            A.TableCell tableCell1 = tableRow1.GetFirstChild<A.TableCell>();
            A.ExtensionList extensionList2 = tableRow1.GetFirstChild<A.ExtensionList>();
            tableCell1.GridSpan = 2;

            A.TableCell tableCell2 = new A.TableCell() { HorizontalMerge = true };

            A.TextBody textBody1 = new A.TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);
            A.TableCellProperties tableCellProperties1 = new A.TableCellProperties();

            tableCell2.Append(textBody1);
            tableCell2.Append(tableCellProperties1);
            tableRow1.InsertBefore(tableCell2, extensionList2);

            A.TableCell tableCell3 = tableRow2.GetFirstChild<A.TableCell>();
            A.ExtensionList extensionList3 = tableRow2.GetFirstChild<A.ExtensionList>();
            tableCell3.GridSpan = 2;

            A.TableCell tableCell4 = new A.TableCell() { HorizontalMerge = true };

            A.TextBody textBody2 = new A.TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);
            A.TableCellProperties tableCellProperties2 = new A.TableCellProperties();

            tableCell4.Append(textBody2);
            tableCell4.Append(tableCellProperties2);
            tableRow2.InsertBefore(tableCell4, extensionList3);

            A.TableRow tableRow3 = new A.TableRow() { Height = 95250L };

            A.TableCell tableCell5 = new A.TableCell();

            A.TextBody textBody3 = new A.TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, FontAlignment = A.TextFontAlignmentValues.Bottom };

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "es-ES", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill1.Append(rgbColorModelHex1);
            A.EffectList effectList1 = new A.EffectList();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties1.Append(solidFill1);
            runProperties1.Append(effectList1);
            runProperties1.Append(latinFont1);
            A.Text text1 = new A.Text();
            text1.Text = "EMCYT";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph3.Append(paragraphProperties1);
            paragraph3.Append(run1);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            A.TableCellProperties tableCellProperties3 = new A.TableCellProperties() { LeftMargin = 7620, RightMargin = 7620, TopMargin = 7620, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties1 = new A.LeftBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill2.Append(rgbColorModelHex2);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd1 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            leftBorderLineProperties1.Append(solidFill2);
            leftBorderLineProperties1.Append(presetDash1);
            leftBorderLineProperties1.Append(round1);
            leftBorderLineProperties1.Append(headEnd1);
            leftBorderLineProperties1.Append(tailEnd1);

            A.RightBorderLineProperties rightBorderLineProperties1 = new A.RightBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill3.Append(rgbColorModelHex3);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            rightBorderLineProperties1.Append(solidFill3);
            rightBorderLineProperties1.Append(presetDash2);
            rightBorderLineProperties1.Append(round2);
            rightBorderLineProperties1.Append(headEnd2);
            rightBorderLineProperties1.Append(tailEnd2);

            A.TopBorderLineProperties topBorderLineProperties1 = new A.TopBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill4.Append(rgbColorModelHex4);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3 = new A.Round();
            A.HeadEnd headEnd3 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties1.Append(solidFill4);
            topBorderLineProperties1.Append(presetDash3);
            topBorderLineProperties1.Append(round3);
            topBorderLineProperties1.Append(headEnd3);
            topBorderLineProperties1.Append(tailEnd3);

            A.BottomBorderLineProperties bottomBorderLineProperties1 = new A.BottomBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill5.Append(rgbColorModelHex5);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round4 = new A.Round();
            A.HeadEnd headEnd4 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd4 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties1.Append(solidFill5);
            bottomBorderLineProperties1.Append(presetDash4);
            bottomBorderLineProperties1.Append(round4);
            bottomBorderLineProperties1.Append(headEnd4);
            bottomBorderLineProperties1.Append(tailEnd4);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties1 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill1 = new A.NoFill();
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties1.Append(noFill1);
            topLeftToBottomRightBorderLineProperties1.Append(presetDash5);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties1 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill2 = new A.NoFill();
            A.PresetDash presetDash6 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties1.Append(noFill2);
            bottomLeftToTopRightBorderLineProperties1.Append(presetDash6);

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill6.Append(schemeColor1);

            tableCellProperties3.Append(leftBorderLineProperties1);
            tableCellProperties3.Append(rightBorderLineProperties1);
            tableCellProperties3.Append(topBorderLineProperties1);
            tableCellProperties3.Append(bottomBorderLineProperties1);
            tableCellProperties3.Append(topLeftToBottomRightBorderLineProperties1);
            tableCellProperties3.Append(bottomLeftToTopRightBorderLineProperties1);
            tableCellProperties3.Append(solidFill6);

            tableCell5.Append(textBody3);
            tableCell5.Append(tableCellProperties3);

            A.TableCell tableCell6 = new A.TableCell();

            A.TextBody textBody4 = new A.TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center, FontAlignment = A.TextFontAlignmentValues.Bottom };

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "es-ES", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Dirty = false };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill7.Append(rgbColorModelHex6);
            A.EffectList effectList2 = new A.EffectList();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };

            runProperties2.Append(solidFill7);
            runProperties2.Append(effectList2);
            runProperties2.Append(latinFont2);
            A.Text text2 = new A.Text();
            text2.Text = "EMGALITY";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph4.Append(paragraphProperties2);
            paragraph4.Append(run2);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);

            A.TableCellProperties tableCellProperties4 = new A.TableCellProperties() { LeftMargin = 7620, RightMargin = 7620, TopMargin = 7620, BottomMargin = 0, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties2 = new A.LeftBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill8.Append(rgbColorModelHex7);
            A.PresetDash presetDash7 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round5 = new A.Round();
            A.HeadEnd headEnd5 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd5 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            leftBorderLineProperties2.Append(solidFill8);
            leftBorderLineProperties2.Append(presetDash7);
            leftBorderLineProperties2.Append(round5);
            leftBorderLineProperties2.Append(headEnd5);
            leftBorderLineProperties2.Append(tailEnd5);

            A.RightBorderLineProperties rightBorderLineProperties2 = new A.RightBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill9.Append(rgbColorModelHex8);
            A.PresetDash presetDash8 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6 = new A.Round();
            A.HeadEnd headEnd6 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            rightBorderLineProperties2.Append(solidFill9);
            rightBorderLineProperties2.Append(presetDash8);
            rightBorderLineProperties2.Append(round6);
            rightBorderLineProperties2.Append(headEnd6);
            rightBorderLineProperties2.Append(tailEnd6);

            A.TopBorderLineProperties topBorderLineProperties2 = new A.TopBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill10.Append(rgbColorModelHex9);
            A.PresetDash presetDash9 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round7 = new A.Round();
            A.HeadEnd headEnd7 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd7 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties2.Append(solidFill10);
            topBorderLineProperties2.Append(presetDash9);
            topBorderLineProperties2.Append(round7);
            topBorderLineProperties2.Append(headEnd7);
            topBorderLineProperties2.Append(tailEnd7);

            A.BottomBorderLineProperties bottomBorderLineProperties2 = new A.BottomBorderLineProperties() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "F2F2F2" };

            solidFill11.Append(rgbColorModelHex10);
            A.PresetDash presetDash10 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round8 = new A.Round();
            A.HeadEnd headEnd8 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd8 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties2.Append(solidFill11);
            bottomBorderLineProperties2.Append(presetDash10);
            bottomBorderLineProperties2.Append(round8);
            bottomBorderLineProperties2.Append(headEnd8);
            bottomBorderLineProperties2.Append(tailEnd8);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties2 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill3 = new A.NoFill();
            A.PresetDash presetDash11 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties2.Append(noFill3);
            topLeftToBottomRightBorderLineProperties2.Append(presetDash11);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties2 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill4 = new A.NoFill();
            A.PresetDash presetDash12 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties2.Append(noFill4);
            bottomLeftToTopRightBorderLineProperties2.Append(presetDash12);

            A.SolidFill solidFill12 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill12.Append(schemeColor2);

            tableCellProperties4.Append(leftBorderLineProperties2);
            tableCellProperties4.Append(rightBorderLineProperties2);
            tableCellProperties4.Append(topBorderLineProperties2);
            tableCellProperties4.Append(bottomBorderLineProperties2);
            tableCellProperties4.Append(topLeftToBottomRightBorderLineProperties2);
            tableCellProperties4.Append(bottomLeftToTopRightBorderLineProperties2);
            tableCellProperties4.Append(solidFill12);

            tableCell6.Append(textBody4);
            tableCell6.Append(tableCellProperties4);

            A.ExtensionList extensionList4 = new A.ExtensionList();

            A.Extension extension2 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            //OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2428161736\" />");

            //extension2.Append(openXmlUnknownElement2);

            extensionList4.Append(extension2);

            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);
            tableRow3.Append(extensionList4);
            table1.Append(tableRow3);
        }



    }





}





